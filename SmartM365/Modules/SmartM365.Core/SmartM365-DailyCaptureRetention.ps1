<#
.SYNOPSIS
Receipt-qualified daily capture retention. Internal, opt-in Core helper.
.VERSION
1.0.0
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
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
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
