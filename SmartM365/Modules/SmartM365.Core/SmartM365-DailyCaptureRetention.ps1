<#
.SYNOPSIS
Receipt-qualified daily capture retention. Internal, opt-in Core helper.
.VERSION
1.0.1
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Reads the established Core publication/authentication context; does not introduce operational globals.')]
param()
$script:SmartM365DailyCapture = $null

function Test-SmartM365DailyCapturePath {
    param([string]$Path)
    if (-not $script:SmartM365DailyCapture) { return $false }
    $full = [IO.Path]::GetFullPath($Path).TrimEnd([char[]]'\/')
    $root = $script:SmartM365DailyCapture.Root
    return $full -ieq $root -or $full.StartsWith($root + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)
}

function Resolve-SmartM365DailyCapturePath {
    param([string]$Root,[string]$Relative,[switch]$Metadata)
    if ([IO.Path]::IsPathRooted($Relative) -or $Relative -match '(^|[\\/])\.\.([\\/]|$)|:' -or
        $Relative -match '(?i)(^|[\\/])(WeeklyHistory|DATA-LAST|Corrections|Archive|Cache|Resume)([\\/]|$)') { throw 'Unsafe capture path.' }
    if (-not $Metadata -and ([IO.Path]::GetExtension($Relative) -ine '.csv' -or $Relative -match '(?i)receipt|manifest')) { throw 'Only owned capture CSVs are eligible.' }
    $path = [IO.Path]::GetFullPath((Join-Path $Root $Relative))
    if (-not $path.StartsWith($Root.TrimEnd([char[]]'\/') + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Capture path escapes its output root.' }
    $node = $path
    while ($node) {
        if (Test-Path -LiteralPath $node) {
            if ((Get-Item -LiteralPath $node -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked capture paths are not eligible for automatic retention.' }
        }
        $parent = [IO.Path]::GetDirectoryName($node.TrimEnd([char[]]'\/'))
        if ($parent -eq $node) { break }; $node = $parent
    }
    return $path
}

function Get-SmartM365DailyCaptureHash {
    param([string]$Path)
    $before = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    $bytes = $before.Length; $ticks = $before.LastWriteTimeUtc.Ticks
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    $after = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($after.Length -ne $bytes -or $after.LastWriteTimeUtc.Ticks -ne $ticks) { throw 'Capture changed during verification.' }
    return @{Bytes=$bytes;SHA256=$hash}
}

function Get-SmartM365DailyCaptureDay {
    param([datetimeoffset]$Utc)
    try { $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Europe/Paris') }
    catch { $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time') }
    return [TimeZoneInfo]::ConvertTime($Utc,$zone).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
}

function Enable-SmartM365DailyCaptureRetention {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot)
    $context = $script:SmartM365CmdbSourceContext
    $script:SmartM365DailyCapture = $null
    if (-not $context -or (Test-SmartM365MaxItemsMode)) { return }
    $allowed = @('SmartM365-M365UserActivity-Inventory.ps1','SmartM365-CopilotUsage-Inventory.ps1',
        'SmartM365-WorkplaceScope-Inventory.ps1','SmartM365-IntuneConfigurations-Inventory.ps1',
        'SmartM365-IntuneManagedApps-Inventory.ps1','SmartM365-Licences-Inventory.ps1',
        'SmartM365-LicensePricing-Inventory.ps1','SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1')
    if ($context.Producer -notin $allowed) { throw 'Daily capture retention is not enabled for this producer.' }
    $root = [IO.Path]::GetFullPath($OutputRoot).TrimEnd([char[]]'\/')
    if ($root -match '(?i)(^|[\\/])(DATA-LAST|WeeklyHistory|Corrections)([\\/]|$)' -or $root -ieq $context.SourceRoot) { throw 'Invalid capture output root.' }
    # Validate the root before creating any state; no legacy payload is adopted by scanning.
    $state = @{Root=$root;Producer=$context.Producer;RunId=$context.RunId;RetentionDays=7;Disabled=$false}
    try { $null = Resolve-SmartM365DailyCapturePath $root 'DailyCaptureRecords/probe.json.txt' -Metadata }
    catch {
        # An unqualified filesystem must not break collection or fall back to per-file deletion.
        $state.Disabled=$true
        WriteLog -Message ('Daily capture retention disabled for this output root; legacy cleanup also blocked: {0}' -f $_.Exception.Message) -Level WARNING
    }
    $script:SmartM365DailyCapture = $state
}

function Write-SmartM365DailyCaptureRecord {
    param([string]$Path,$Record,[string]$ExpectedHash='ABSENT')
    $actual = if (Test-Path -LiteralPath $Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash} else {'ABSENT'}
    if ($actual -ine $ExpectedHash) { throw 'Capture metadata changed concurrently.' }
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $temporary = [IO.Path]::Combine([IO.Path]::GetDirectoryName($Path), ('~sm' + [guid]::NewGuid().ToString('N').Substring(0, 12) + '.tmp'))
    try {
        [IO.File]::WriteAllText($temporary,($Record | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        $check = if (Test-Path -LiteralPath $Path) {(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash} else {'ABSENT'}
        if ($check -ine $ExpectedHash) { throw 'Capture metadata changed before promotion.' }
        if ($ExpectedHash -eq 'ABSENT') {[IO.File]::Move($temporary,$Path)}
        else {[IO.File]::Replace($temporary,$Path,[Management.Automation.Language.NullString]::Value,$true)}
    } finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function Get-SmartM365DailyCaptureRemoteItem {
    param([string]$LocalPath)
    $thumb=if($global:Thumbprint){$global:Thumbprint}else{$global:Thumb}
    if (-not (Connect-SmartM365GraphForSharePointUpload -AppId $global:AppId -TenantId $global:TenantId -Thumbprint $thumb)) { throw 'SharePoint retention authentication unavailable.' }
    $key = '{0}|{1}|{2}' -f $global:SharePointSiteHostname,$global:SharePointSitePath,$global:SharePointLibraryDisplayName
    if (-not $script:SmartM365SharePointDriveIdCache -or -not $script:SmartM365SharePointDriveIdCache.ContainsKey($key)) {
        throw 'SharePoint drive was not qualified by current-run publication.'
    }
    $drive = $script:SmartM365SharePointDriveIdCache[$key]
    $target = ConvertTo-SmartM365SharePointDataRootPath $global:SharePointTargetFolderPath
    $relative = Get-SmartM365SharePointRelativeFilePath $LocalPath
    $uri = 'https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}' -f $drive,(ConvertTo-GraphDrivePath ($target.TrimEnd('/')+'/'+$relative.TrimStart('/')))
    $item = Invoke-SmartM365GraphRestWithRetry -Method GET -Uri $uri -Operation 'Inspect daily capture'
    if (-not $item.id -or -not $item.eTag -or -not $item.PSObject.Properties['file']) { throw 'SharePoint capture is not a versioned file.' }
    return @{Uri=('https://graph.microsoft.com/v1.0/drives/{0}/items/{1}' -f $drive,$item.id);ETag=[string]$item.eTag;Bytes=[int64]$item.size}
}

function Assert-SmartM365DailyCaptureRemoteFile {
    param([string]$Path,$Expected)
    $item = Get-SmartM365DailyCaptureRemoteItem $Path
    if ($item.Bytes -ne $Expected.Bytes) { throw 'SharePoint capture size differs.' }
    $temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CaptureReadback-'+[guid]::NewGuid().ToString('N')+'.csv')
    try {
        Invoke-SmartM365GraphFileDownloadWithRetry -Uri ($item.Uri+'/content') -OutputFilePath $temporary -Operation 'Verify daily capture SHA256' | Out-Null
        if ((Get-SmartM365DailyCaptureHash $temporary).SHA256 -ine $Expected.SHA256) { throw 'SharePoint capture SHA256 differs.' }
        $after = Get-SmartM365DailyCaptureRemoteItem $Path
        if ($after.Uri -cne $item.Uri -or $after.ETag -cne $item.ETag) { throw 'SharePoint capture changed during read-back.' }
        return $item
    } finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function Invoke-SmartM365DailyCaptureRetention {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)]$Proof,[datetimeoffset]$ReferenceUtc=[datetimeoffset]::UtcNow)
    if ($WhatIfPreference) { return }
    $state = $script:SmartM365DailyCapture
    if (-not $state -or $state.Disabled -or $Proof.Status -ne 'Completed' -or $Proof.RunId -cne $state.RunId -or $Context.Producer -cne $state.Producer) { return }
    $root = $state.Root
    $recordFolder = Resolve-SmartM365DailyCapturePath $root ('DailyCaptureRecords/'+[IO.Path]::GetFileNameWithoutExtension($state.Producer)) -Metadata
    if (-not (Test-Path -LiteralPath $recordFolder)) { New-Item -ItemType Directory -Path $recordFolder -Force | Out-Null }
    $lockPath = Join-Path $recordFolder '.retention.lock'
    $lock = [IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $protectedStreams=[Collections.Generic.List[IDisposable]]::new()
    try {
        $records = [Collections.Generic.List[object]]::new()
        foreach ($meta in @(Get-ChildItem -LiteralPath $recordFolder -File -Filter '*.json.txt' -ErrorAction Stop)) {
            $null = Resolve-SmartM365DailyCapturePath $root ($meta.FullName.Substring($root.Length+1)) -Metadata
            $value = [IO.File]::ReadAllText($meta.FullName) | ConvertFrom-Json
            # Recent PowerShell versions materialize JSON timestamps as DateTime values.
            # Preserve their instant instead of reparsing their culture-dependent display text.
            if ($value.CompletedAtUtc -is [datetime] -or $value.CompletedAtUtc -is [datetimeoffset]) {
                $value.CompletedAtUtc=([datetimeoffset]$value.CompletedAtUtc).ToUniversalTime().ToString('o',[Globalization.CultureInfo]::InvariantCulture)
            }
            if ($value.Owner -cne 'SmartM365-DailyCapture' -or $value.Version -ne 1 -or $value.OutputRoot -ine $root -or $value.Status -ne 'Completed' -or $value.TimeZone -cne 'Europe/Paris' -or $value.RunId -notmatch '^[a-fA-F0-9]{32}$') { throw 'Invalid capture metadata; retention refused.' }
            if ($meta.Name -cne ($value.RunId+'.json.txt') -or @($records | Where-Object {$_.Value.RunId -eq $value.RunId}).Count) { throw 'Repeated capture run identity.' }
            foreach ($field in @('Producer','TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {if ($value.$field -cne $Context[$field]) { throw 'Capture identity differs.' }}
            $date = [datetimeoffset]::Parse($value.CompletedAtUtc,[Globalization.CultureInfo]::InvariantCulture)
            if ($value.Day -cne (Get-SmartM365DailyCaptureDay $date) -or $date -gt $ReferenceUtc) { throw 'Invalid capture completion date.' }
            if (@($value.Files).Count -eq 0 -or @($value.Files | Group-Object File | Where-Object Count -GT 1).Count) { throw 'Capture file set is empty or repeated.' }
            foreach ($file in $value.Files) {
                $null = Resolve-SmartM365DailyCapturePath $root $file.Relative
                if ($file.SHA256 -notmatch '^[a-fA-F0-9]{64}$' -or [int64]$file.Bytes -lt 0) { throw 'Invalid capture fingerprint.' }
                $leaf=[IO.Path]::GetFileName([string]$file.Relative)
                $timestamped='^'+[regex]::Escape([IO.Path]::GetFileNameWithoutExtension([string]$file.File))+'_\d{8}[-_]\d{6}\.csv$'
                $newRelative=('DailyCaptures/'+$value.RunId+'/'+$file.File) -replace '\\','/'
                if ([IO.Path]::GetFileName([string]$file.File) -cne [string]$file.File -or
                    ($leaf -notmatch $timestamped -and (($file.Relative -replace '\\','/') -cne $newRelative))) { throw 'Capture file does not match its declared export.' }
            }
            $records.Add(@{Path=$meta.FullName;Hash=(Get-FileHash -LiteralPath $meta.FullName -Algorithm SHA256).Hash;Value=$value})
        }
        if (@($records | Where-Object {$_.Value.RunId -ceq $Proof.RunId}).Count) { throw 'This run capture is already recorded; no overwrite is permitted.' }
        $completeAt = [datetimeoffset]::Parse($Proof.CompletedAtUtc,[Globalization.CultureInfo]::InvariantCulture)
        if ($completeAt -gt $ReferenceUtc) { throw 'Current capture is dated in the future.' }
        foreach ($previous in $records) {
            if ([datetimeoffset]::Parse($previous.Value.CompletedAtUtc) -gt $completeAt) { throw 'Current capture would regress the completion sequence.' }
        }
        # A later run selecting fewer exports must never displace a fuller same-day capture.
        $currentNames=@($Proof.Files | ForEach-Object {$_.File})
        foreach ($previous in $records) {
            if ($previous.Value.Day -eq (Get-SmartM365DailyCaptureDay $completeAt)) {
                foreach ($file in $previous.Value.Files) {if ($file.File -notin $currentNames) { throw 'Current capture has fewer exports than the existing daily capture; no purge.' }}
            }
        }
        $day = Get-SmartM365DailyCaptureDay $completeAt
        $files = [Collections.Generic.List[object]]::new()
        $ownedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($previous in $records) { foreach ($file in $previous.Value.Files) { [void]$ownedPaths.Add([string]$file.Relative) } }
        foreach ($source in $Proof.Files) {
            if ([IO.Path]::GetFileName([string]$source.File) -cne [string]$source.File -or $source.File -notlike '*.csv') { throw 'Invalid receipt file name.' }
            $current = Join-Path $Context.SourceRoot $source.File
            $expected = Get-SmartM365DailyCaptureHash $current
            if ($expected.SHA256 -ine [string]$source.SHA256) { throw 'Current CSV changed after receipt qualification.' }
            $capture = $null
            $pattern = '^'+[regex]::Escape([IO.Path]::GetFileNameWithoutExtension($source.File))+'_\d{8}[-_]\d{6}\.csv$'
            foreach ($published in @($global:csvGeneratedPaths | Sort-Object -Descending)) {
                if (-not (Test-SmartM365DailyCapturePath $published) -or [IO.Path]::GetFileName($published) -notmatch $pattern -or -not (Test-Path -LiteralPath $published -PathType Leaf)) { continue }
                $relative = [IO.Path]::GetFullPath($published).Substring($root.Length+1)
                $path = Resolve-SmartM365DailyCapturePath $root $relative
                if ($ownedPaths.Contains($relative)) { continue }
                $hash = Get-SmartM365DailyCaptureHash $path
                if ($hash.Bytes -eq $expected.Bytes -and $hash.SHA256 -ieq $expected.SHA256) { $capture=$path; break }
            }
            if (-not $capture) {
                $capture = Resolve-SmartM365DailyCapturePath $root ('DailyCaptures/'+$Proof.RunId+'/'+$source.File)
                $parent = Split-Path $capture -Parent
                if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
                if (Test-Path -LiteralPath $capture) { throw 'New capture payload already exists.' }
                [IO.File]::Copy($current,$capture,$false)
            }
            if ((Get-SmartM365DailyCaptureHash $capture).SHA256 -ine $expected.SHA256) { throw 'Capture copy hash mismatch.' }
            $files.Add(@{File=[string]$source.File;Relative=$capture.Substring($root.Length+1);Bytes=$expected.Bytes;SHA256=$expected.SHA256})
        }
        foreach ($required in $Proof.RequiredFiles) { if ($required -notin @($files | ForEach-Object {$_.File})) { throw 'Required file missing from complete capture.' } }
        if (-not $files.Count) { throw 'An empty capture cannot authorize retention.' }
        $remote = [bool]$global:EnableSharePointUpload
        if ($remote) {
            foreach ($file in $files) {
                $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
                $uploaded = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path -EnsureParentFolders
                if (-not $uploaded) { throw 'Capture publication did not succeed; no purge.' }
                $null = Assert-SmartM365DailyCaptureRemoteFile $path $file
            }
        }
        # The source receipt lock is still held. Fence all current sources again before committing the capture.
        foreach ($file in $files) {if ((Get-SmartM365DailyCaptureHash (Join-Path $Context.SourceRoot $file.File)).SHA256 -ine $file.SHA256) { throw 'Current source changed while closing the capture.' }}
        $record = @{Owner='SmartM365-DailyCapture';Version=1;OutputRoot=$root;RunId=$Proof.RunId;Status='Completed';TimeZone='Europe/Paris';Day=$day;CompletedAtUtc=$Proof.CompletedAtUtc;Files=$files.ToArray();RemovedFiles=@();ScopeParameters=$Proof.ScopeParameters;ScriptVersion=$Proof.ScriptVersion}
        foreach ($field in @('Producer','TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {$record[$field]=$Context[$field]}
        $recordPath = Join-Path $recordFolder ($Proof.RunId+'.json.txt')
        Write-SmartM365DailyCaptureRecord $recordPath $record
        if ($remote) {
            if (-not (Invoke-SmartM365SharePointCsvUpload -LocalFilePath $recordPath -EnsureParentFolders)) { throw 'Capture metadata publication failed; no purge.' }
            $null = Assert-SmartM365DailyCaptureRemoteFile $recordPath (Get-SmartM365DailyCaptureHash $recordPath)
        }
        $records.Add(@{Path=$recordPath;Hash=(Get-FileHash -LiteralPath $recordPath -Algorithm SHA256).Hash;Value=([pscustomobject]$record)})
        $today = Get-SmartM365DailyCaptureDay $ReferenceUtc
        $cutoff = [datetime]::ParseExact($today,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(-6).ToString('yyyy-MM-dd')
        $keep = @{}
        foreach ($entry in $records) {
            $value = $entry.Value
            if ($value.Day -lt $cutoff) { continue }
            if (-not $keep.ContainsKey($value.Day) -or [datetimeoffset]::Parse($value.CompletedAtUtc) -gt [datetimeoffset]::Parse($keep[$value.Day].Value.CompletedAtUtc)) {$keep[$value.Day]=$entry}
        }
        # Verify every retained payload before removing even the first obsolete file.
        foreach ($entry in $keep.Values) {
            if (@($entry.Value.RemovedFiles).Count) { throw 'A retained capture has retired payloads.' }
            $metadataLease=[IO.File]::Open($entry.Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
            $protectedStreams.Add($metadataLease)
            if ((Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256).Hash -ine $entry.Hash) { throw 'Retained capture metadata changed.' }
            foreach ($file in $entry.Value.Files) {
                $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
                $lease=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
                $protectedStreams.Add($lease)
                $sha=[Security.Cryptography.SHA256]::Create()
                try {$hash=[BitConverter]::ToString($sha.ComputeHash($lease)).Replace('-','')} finally {$sha.Dispose()}
                if ($lease.Length -ne $file.Bytes -or $hash -ine $file.SHA256) { throw 'Retained capture is incomplete or changed; no purge.' }
                if ($remote) { $null = Assert-SmartM365DailyCaptureRemoteFile $path $file }
            }
        }
        $deletions = [Collections.Generic.List[object]]::new()
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $records) {
            if ($keep.ContainsKey($entry.Value.Day) -and $keep[$entry.Value.Day].Value.RunId -eq $entry.Value.RunId) { continue }
            foreach ($file in $entry.Value.Files) {
                if ([string]$file.Relative -in @($entry.Value.RemovedFiles)) { continue }
                if (-not $seen.Add([string]$file.Relative)) { throw 'Capture records share a candidate path.' }
                foreach ($retained in $keep.Values) {if ([string]$file.Relative -in @($retained.Value.Files | ForEach-Object {$_.Relative})) { throw 'Candidate overlaps a retained capture.' }}
                $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
                $hash = Get-SmartM365DailyCaptureHash $path
                if ($hash.Bytes -ne $file.Bytes -or $hash.SHA256 -ine $file.SHA256) { throw 'Obsolete capture changed; no purge.' }
                $remoteItem = if ($remote) {Assert-SmartM365DailyCaptureRemoteFile $path $file} else {$null}
                $deletions.Add(@{Entry=$entry;File=$file;Path=$path;Remote=$remoteItem})
            }
        }
        $deleted = 0
        foreach ($candidate in $deletions) {
            if (-not $PSCmdlet.ShouldProcess($candidate.Path,'Remove obsolete receipt-qualified daily capture CSV')) { continue }
            # Recheck metadata, payload and remote ETag immediately before deletion.
            if ((Get-FileHash -LiteralPath $candidate.Entry.Path -Algorithm SHA256).Hash -ine $candidate.Entry.Hash) { throw 'Capture metadata changed before deletion.' }
            $path = Resolve-SmartM365DailyCapturePath $root $candidate.File.Relative
            $stream = [IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::Read -bor [IO.FileShare]::Delete))
            try {
                $sha = [Security.Cryptography.SHA256]::Create()
                try {$hash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')} finally {$sha.Dispose()}
                if ($stream.Length -ne $candidate.File.Bytes -or $hash -ine $candidate.File.SHA256) { throw 'Capture changed immediately before deletion.' }
                if ($remote) {
                    Invoke-SmartM365GraphRestWithRetry -Method DELETE -Uri $candidate.Remote.Uri -AdditionalHeaders @{'If-Match'=$candidate.Remote.ETag} -Operation 'Delete verified obsolete daily capture' | Out-Null
                }
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
            } finally {$stream.Dispose()}
            $candidate.Entry.Value.RemovedFiles=@($candidate.Entry.Value.RemovedFiles)+@([string]$candidate.File.Relative)
            Write-SmartM365DailyCaptureRecord -Path $candidate.Entry.Path -Record $candidate.Entry.Value -ExpectedHash $candidate.Entry.Hash
            $candidate.Entry.Hash=(Get-FileHash -LiteralPath $candidate.Entry.Path -Algorithm SHA256).Hash
            if ($remote) {if (-not (Invoke-SmartM365SharePointCsvUpload -LocalFilePath $candidate.Entry.Path -EnsureParentFolders)) { throw 'Retirement journal upload failed; further deletion stopped.' }}
            $deleted++
        }
        WriteLog -Message ("Daily capture retention: {0} complete day(s) retained; {1} owned CSV(s) removed; seven calendar days, Europe/Paris. Unindexed legacy files preserved." -f $keep.Count,$deleted) -Level INFO
    } finally {foreach($stream in $protectedStreams){$stream.Dispose()};$lock.Dispose()}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/oEXI2lXBn3uV
# KQa4z79V7+KP/mTDVSmJ+SXQDA0HXqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPu+w74YPUXUp6HYj6gV98GMFDT2FHobgh63iie8z2e6MA0GCSqG
# SIb3DQEBAQUABIIBgASddT6fbwD2njmTcfNaoo3V8dGnh8uEwEa3D+4xBAmiJgUZ
# tg5NJJRPGETOh14bg/XbS5zrdwnfVgTmhHi36Yyovm+z4M1vYSspcX/NUzw0v56L
# rct3wzN2mGorf0A93BjCobjEMsr9eDjjpP5CPxf2Xeu5DEKYxNMWwi2LC7d7jxwu
# KciAsZjqTo2HqBbb7ECguvDdU3i8Ne8/TxrSg3r7w9+pR0NxkM9wL2xQKeaFMWHH
# DuxQYrWg7nm7jv9cHB0I56D4VXStf8Ue9awZLtDz/vY2SlflPnpbm2VhmOrRsfYo
# o9W/864d3d7ujS7P15yTenpWqegAKGfIpC+pGNECV/0aZKkkhnuxITCw6NExFqWY
# wrhD4g74gN+Lj51Av0clJTOxm1fm+qxCfpC9w5ImpWS3lhPWBWSFOR3+7F3VZ6js
# U0PLSI8QgMRO6fNSKfMGWrs7P/Bu/u7gHZpHWUnP9BUtVQY30VsA4tFVRDU+fhkR
# 6loujrWPPrgvfpf3UaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNzE2
# MjJaMC8GCSqGSIb3DQEJBDEiBCCzxdVqUr9Tbm84sXYSQXBeJlWxJk+5pWGO3fxC
# PrUfGTANBgkqhkiG9w0BAQEFAASCAgBZnbDXfvX7jinbVy0lClQtxgS0UbmcNppU
# Q9v55VHpzGR/qB7ALHFNBLPuEtyF5+AQC8kWBNuqDRmn5/zfafykQWMZv/JVWJBB
# 27bITn/L46cWqd/GqrDxrqbISVvo4CWAJYY5zWSeWTdxQb5qwzY3G6FkjSdRSFmb
# QWyDsHG5XRQRnrrniFNwAUJqza7MlvIaGtmtjcv+/SM9P+TvHLc56KnUjot1mXti
# hySnNargv0FA7K0KpZs6HmPRj3+x9Ejd48vlHkDTuaYqjRpsgYjj45kV9p1K5ZA2
# dnmnM1DYBOiA7drv98PMF2wQNQ2KnS1YbW6a4Fs8KHUTMt6vFcYCvveNQD7dcxCQ
# qbOX9FdAxPVMhctgY3s3JbEf2ri4xFYnSXoV8EA3X9nQYeEK8jU1Mbb+9Sl551kl
# ar8iNKM4yfFqNh+74h0OZVbFuQUiQ3AxDh02WdP+5OXi/zy1IQTDGehW6XdOpkRe
# kAtQzCtCVCsz9dy1DpG/kDrH7DSCRhUO3HuR6c+vp3X8lfkPir8GJVbNCizox9Hh
# 8NLfcexv1mz+iBGrgXrfT4zdoVuna7OgYqYTnnzhCZAOaU55aMxSIyZRKfvpVnCu
# AXmw5/fxfwDMnvkQ+a7+hKavSmsn4cY1ucMsSWaS1CDC3ghdMdiUOPtYLuDyyIm4
# c7+Yd5T14g==
# SIG # End signature block
