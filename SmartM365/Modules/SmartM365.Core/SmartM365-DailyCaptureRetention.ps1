<#
.SYNOPSIS
Receipt-qualified daily capture retention. Internal, opt-in Core helper.
.VERSION
1.0.4
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

function Test-SmartM365DailyCaptureRecordedRoot {
    param([string]$Recorded,[string]$Root)
    if ($Recorded -ieq $Root) { return $true }
    # A relocated workspace keeps the tenant-relative output identity; recorded metadata is never rewritten.
    $pattern = '(?i)[\\/]Tenants[\\/](?<identity>[^\\/]+[\\/].+)$'
    $recordedIdentity = [regex]::Match($Recorded.TrimEnd([char[]]'\/'),$pattern)
    $currentIdentity = [regex]::Match($Root,$pattern)
    return $recordedIdentity.Success -and $currentIdentity.Success -and
        $recordedIdentity.Groups['identity'].Value.Replace('/','\') -ieq $currentIdentity.Groups['identity'].Value.Replace('/','\')
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

# True when this run already uploaded this local file (Publish-SmartM365Csv uploads the timestamped CSV that
# usually becomes the capture) and the upload did not fail afterwards.
function Test-SmartM365DailyCaptureUploadedInRun {
    param([string]$Path,$Expected)
    $full = [IO.Path]::GetFullPath($Path)
    if ($script:SmartM365SharePointFailedUploads -and $script:SmartM365SharePointFailedUploads.Contains($full)) { return $false }
    foreach ($record in @($global:SmartM365SharePointUploadedFiles)) {
        if ($null -eq $record -or [string]::IsNullOrWhiteSpace([string]$record.LocalFilePath)) { continue }
        if ([IO.Path]::GetFullPath([string]$record.LocalFilePath) -ieq $full -and [int64]$record.Size -eq $Expected.Bytes) { return $true }
    }
    return $false
}

# Verification memory (1.0.4): a retained file fully verified (local SHA-256 and SharePoint read-back) less than
# seven days ago is only checked for size, local write time and SharePoint ETag. Any difference, a missing or
# unreadable memory, or an older full check falls back to the full verification. Deletions always verify fully.
function Read-SmartM365DailyCaptureVerificationMemory {
    param([string]$Path)
    $memory = @{}
    try {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            $document = [IO.File]::ReadAllText($Path) | ConvertFrom-Json
            if ($document.Owner -ceq 'SmartM365-DailyCaptureVerification' -and $document.Version -eq 1 -and $document.Files) {
                foreach ($property in @($document.Files.PSObject.Properties)) { $memory[$property.Name] = $property.Value }
            }
        }
    }
    catch { WriteLog -Message ('Daily capture verification memory unreadable; retained files get a full check: {0}' -f $_.Exception.Message) -Level WARNING }
    return $memory
}

function Save-SmartM365DailyCaptureVerificationMemory {
    param([string]$Path,[hashtable]$Memory)
    $temporary = [IO.Path]::Combine([IO.Path]::GetDirectoryName($Path), ('~sm' + [guid]::NewGuid().ToString('N').Substring(0, 12) + '.tmp'))
    try {
        $files = [ordered]@{}
        foreach ($key in @($Memory.Keys | Sort-Object)) { $files[$key] = $Memory[$key] }
        [IO.File]::WriteAllText($temporary,([ordered]@{Owner='SmartM365-DailyCaptureVerification';Version=1;Files=$files} | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $Path) {[IO.File]::Replace($temporary,$Path,[Management.Automation.Language.NullString]::Value,$true)}
        else {[IO.File]::Move($temporary,$Path)}
    }
    catch { WriteLog -Message ('Daily capture verification memory not saved; the next run checks retained files fully: {0}' -f $_.Exception.Message) -Level WARNING }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue } }
}

function Get-SmartM365DailyCaptureVerificationEntry {
    param([string]$Path,$File,$RemoteItem,[datetimeoffset]$VerifiedAtUtc)
    return [ordered]@{
        Bytes=[int64]$File.Bytes;SHA256=[string]$File.SHA256
        LastWriteTimeUtcTicks=(Get-Item -LiteralPath $Path -Force -ErrorAction Stop).LastWriteTimeUtc.Ticks
        RemoteETag=if($RemoteItem){[string]$RemoteItem.ETag}else{''};VerifiedAtUtcTicks=$VerifiedAtUtc.UtcTicks
    }
}

function Test-SmartM365DailyCaptureLightCheck {
    param($Entry,$File,[IO.Stream]$Lease,[string]$Path,[bool]$Remote,[datetimeoffset]$ReferenceUtc)
    try {
        if (-not $Entry -or [string]$Entry.SHA256 -ine [string]$File.SHA256 -or [int64]$Entry.Bytes -ne [int64]$File.Bytes) { return $false }
        $age = $ReferenceUtc.UtcTicks - [int64]$Entry.VerifiedAtUtcTicks
        if ($age -lt 0 -or $age -ge [TimeSpan]::FromDays(7).Ticks) { return $false }
        if ($Lease.Length -ne [int64]$File.Bytes -or (Get-Item -LiteralPath $Path -Force -ErrorAction Stop).LastWriteTimeUtc.Ticks -ne [int64]$Entry.LastWriteTimeUtcTicks) { return $false }
        if ($Remote) {
            if ([string]::IsNullOrEmpty([string]$Entry.RemoteETag)) { return $false }
            $item = Get-SmartM365DailyCaptureRemoteItem $Path
            if ([int64]$item.Bytes -ne [int64]$File.Bytes -or [string]$item.ETag -cne [string]$Entry.RemoteETag) { return $false }
        }
        return $true
    }
    catch { return $false }
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
            if ($value.Owner -cne 'SmartM365-DailyCapture' -or $value.Version -ne 1 -or -not (Test-SmartM365DailyCaptureRecordedRoot ([string]$value.OutputRoot) $root) -or $value.Status -ne 'Completed' -or $value.TimeZone -cne 'Europe/Paris' -or $value.RunId -notmatch '^[a-fA-F0-9]{32}$') { throw 'Invalid capture metadata; retention refused.' }
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
        $verificationPath = Join-Path $recordFolder '.retention.verification.json'
        $verification = Read-SmartM365DailyCaptureVerificationMemory $verificationPath
        foreach ($file in $files) {
            $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
            $remoteItem = $null
            if ($remote) {
                # Uploaded earlier in this run: the SHA-256 read-back proves SharePoint holds these bytes, so the
                # file is not sent twice. Any read-back failure falls back to the upload below.
                if (Test-SmartM365DailyCaptureUploadedInRun $path $file) {
                    try {
                        $remoteItem = Assert-SmartM365DailyCaptureRemoteFile $path $file
                        WriteLog -Message ("Daily capture already uploaded in this run; SharePoint copy verified by SHA-256, upload skipped: {0}" -f $file.Relative) -Level INFO
                    }
                    catch { WriteLog -Message ("Daily capture read-back after this run's upload failed; uploading again: {0}" -f $_.Exception.Message) -Level WARNING }
                }
                if (-not $remoteItem) {
                    $uploaded = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path -EnsureParentFolders
                    if (-not $uploaded) { throw 'Capture publication did not succeed; no purge.' }
                    $remoteItem = Assert-SmartM365DailyCaptureRemoteFile $path $file
                }
            }
            # The capture of this run was just verified in full (copy hash and SharePoint read-back).
            $verification[[string]$file.Relative] = Get-SmartM365DailyCaptureVerificationEntry $path $file $remoteItem $ReferenceUtc
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
        $retainedTotal = @($keep.Values | ForEach-Object { @($_.Value.Files) }).Count
        $retainedIndex = 0; $lightChecks = 0; $fullChecks = 0
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $retainedMemory = @{}
        WriteLog -Message ("Daily capture retention: checking {0} retained file(s) of {1} day(s)." -f $retainedTotal,$keep.Count) -Level INFO
        foreach ($entry in @($keep.Values | Sort-Object { $_.Value.Day })) {
            if (@($entry.Value.RemovedFiles).Count) { throw 'A retained capture has retired payloads.' }
            $metadataLease=[IO.File]::Open($entry.Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
            $protectedStreams.Add($metadataLease)
            if ((Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256).Hash -ine $entry.Hash) { throw 'Retained capture metadata changed.' }
            foreach ($file in $entry.Value.Files) {
                $retainedIndex++
                $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
                $lease=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
                $protectedStreams.Add($lease)
                $memoryEntry = $verification[[string]$file.Relative]
                if (Test-SmartM365DailyCaptureLightCheck $memoryEntry $file $lease $path $remote $ReferenceUtc) {
                    $lightChecks++
                    $retainedMemory[[string]$file.Relative] = $memoryEntry
                    continue
                }
                WriteLog -Message ("Daily capture retention: full check {0}/{1} ({2}, {3:N1} MiB): {4}" -f $retainedIndex,$retainedTotal,$entry.Value.Day,([int64]$file.Bytes/1MB),$file.Relative) -Level INFO
                $sha=[Security.Cryptography.SHA256]::Create()
                try {$hash=[BitConverter]::ToString($sha.ComputeHash($lease)).Replace('-','')} finally {$sha.Dispose()}
                if ($lease.Length -ne $file.Bytes -or $hash -ine $file.SHA256) { throw 'Retained capture is incomplete or changed; no purge.' }
                $remoteItem = if ($remote) { Assert-SmartM365DailyCaptureRemoteFile $path $file } else { $null }
                $retainedMemory[[string]$file.Relative] = Get-SmartM365DailyCaptureVerificationEntry $path $file $remoteItem $ReferenceUtc
                $fullChecks++
            }
        }
        # Only retained files are remembered: removed or obsolete payloads drop out of the memory.
        Save-SmartM365DailyCaptureVerificationMemory $verificationPath $retainedMemory
        WriteLog -Message ("Daily capture retention: {0} retained file(s) checked in {1:N0} s; {2} light check(s) (size, write time, SharePoint ETag), {3} full verification(s)." -f $retainedTotal,$clock.Elapsed.TotalSeconds,$lightChecks,$fullChecks) -Level INFO
        $deletions = [Collections.Generic.List[object]]::new()
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $records) {
            if ($keep.ContainsKey($entry.Value.Day) -and $keep[$entry.Value.Day].Value.RunId -eq $entry.Value.RunId) { continue }
            foreach ($file in $entry.Value.Files) {
                if ([string]$file.Relative -in @($entry.Value.RemovedFiles)) { continue }
                if (-not $seen.Add([string]$file.Relative)) { throw 'Capture records share a candidate path.' }
                foreach ($retained in $keep.Values) {if ([string]$file.Relative -in @($retained.Value.Files | ForEach-Object {$_.Relative})) { throw 'Candidate overlaps a retained capture.' }}
                $path = Resolve-SmartM365DailyCapturePath $root $file.Relative
                WriteLog -Message ("Daily capture retention: verifying obsolete capture file {0} ({1}): {2}" -f ($deletions.Count+1),$entry.Value.Day,$file.Relative) -Level INFO
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCb5pAXVQ3H9qv6
# pFhHZE9dKKFrEYYs+Fn3Ag9b0kC4y6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKP1D6rNntRgPoFR4lNhJXrRWnZo4tUmcI0uqKWF4YuwMA0GCSqG
# SIb3DQEBAQUABIIBgKSpu9NXyPFsjqc8pbe4g7wTctQs3b3q24Jlik+brcqvI9vm
# cokqllCDCeJv+0OywEm9NJ2e7NfvzxWXfu5dRFDbQU5mPvrAb5ZJqCW1nOMRDhfp
# DhKvtH9UzWTscatbas45z2HtPZ8491Y/LAkL4zbuE/Ogj0fnzox3dhPEO43mw/aq
# EDaRODJhyNJObWHXOunhbKgXH/eRV2hbMNVqxr98dZOgl6G9PzT5FALZzJ2Bihv7
# gSvzDd/Vf3OQXEISrbDnheegVV0PJHzDX5y47rI4IPAfmGgr8C6TwC8mKuTcx9UU
# ZXemHEUV9ScluRkUorOilF583P4b970+GgcTMMVdNp1PYSLVWbsUqyHITLxpHG1U
# CFrvyEJW1smjoSJVYfd8GxjzkIrv2EbD1NdyPNq/AqK2eclFXSIa6R2Xj0UNpM/N
# BQ5bdh51CxzDB3A12b/nxZqiTxhwig1aCeW1qXw3mVQ7HpnfmoRY0JB5MPOIwR7r
# 8+Crh99ExZh0xd1eHKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxMDEy
# NTNaMC8GCSqGSIb3DQEJBDEiBCBIvJt6AKqVfRB3Uh/dn6onk1S/9b622oQiSun3
# cKb63TANBgkqhkiG9w0BAQEFAASCAgCluOHurIKpMuSsW1JyQR8D+jI+Ks5ssVAl
# 394Iz1HvMKr6QwDbExSGuM5C5SNqW8BBAGaTcarjCMKifseu5hNevjNw3LPcPG7W
# qaWvHLI+xfMqi3er4cAoi3XVCNAwfJ7pWXB+EDDPdUsIvbiSaW/Ob+AFXZrEWgEf
# s4yI5jyE80xygM+GF2wXC6GZj0QByHDPNvT+nCmMGi/DfcO4XzEzBiw+DfPhKFts
# 3lmAUdcYn5NPPf5cav9krbA8m2mvyajUMLEcGp/k22cAYuYMPrifRSC8UHst2jBf
# 8pmzeiZzzzh9SKJNnvvuW4UiOuDCSxeBu4VqoPSj1Esh5QrZn6Pmrgg51FX0ZYHZ
# HERTPCni2oeWejJ2vTDwKAloVy8Oe60E70yCX67VNNg62TAwAcRrfyC9Y6lY3lQz
# MIarHdJ57tnBtsLi4Gd2Ie+l4SvOsuMh45lIVPPQ1GlwWN0q/uRe1vw2/yoC/0G9
# iyBpd6EUvRVRMt6goqmqPq0/tmpNSwXVJbyN2u2yJKZA1p7ymCyJbyqKG6KDV4ED
# gV08TYSLoLPz+lLy1IxIeALzciii9OJt0d/vKgWucCGXtx6RHIptGXJ7z5/P1+F1
# O+FwzWq0LteXiAiDnDZpgKxUwDy+kD7ougkIsq+08im2rAG0At9WG82D9/7WRr60
# tS0kMwIQMg==
# SIG # End signature block
