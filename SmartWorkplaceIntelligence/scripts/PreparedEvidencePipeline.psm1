# Offline, single-tenant preparation and versioned local publication. No tenant APIs.
# Version: 1.2.5
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ProductRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'AdCountryEvidence.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'IntuneOperationalEvidence.psm1')
Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1')
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1')
Import-Module (Join-Path $PSScriptRoot 'PreparedClassification.psm1')
Import-Module (Join-Path $PSScriptRoot 'SharePointMigrationEvidence.psm1')
if (-not ('Workplace.Prepared.StrictCsv' -as [type])) {
    Add-Type -Path (Join-Path $PSScriptRoot 'PreparedCsvReader.cs') -CompilerOptions '-optimize+'
}
$script:PreparedTenantReferenceParses = 0

function Get-PreparedChildPath([string]$Root, [string]$Relative) {
    $base = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $path = [IO.Path]::GetFullPath((Join-Path $base $Relative))
    if (-not $path.StartsWith($base + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Path must be strictly below its declared root.'
    }
    $path
}

function Write-PreparedJson([string]$Path, $Value) {
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth 20))
    if($Path.EndsWith('.json',[StringComparison]::OrdinalIgnoreCase)){
        $Path=Resolve-SmartM365OwnedJsonPath -Path $Path -Owner 'WorkplaceEvidence-Prepare' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Prepared audit must be an object.'}}
    }
    if($Path -match '\.json(?:\.txt)?$'){Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Prepared metadata must be an object.'}} | Out-Null}
    else{[IO.File]::WriteAllBytes($Path,$bytes)}
}

function Move-PreparedFile {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination, [switch]$Overwrite,
        [int[]]$RetryDelaySeconds = @(2,4,8,15))
    # A file just written to an SMB share can be held briefly by the file server (antivirus, indexing).
    # The rename stays atomic; only a refused or busy target is retried, then the failure is final.
    for ($attempt = 1; ; $attempt++) {
        try { [IO.File]::Move($Source,$Destination,[bool]$Overwrite); return }
        catch {
            $exception = $_.Exception
            while ($exception -is [Management.Automation.MethodInvocationException] -and $exception.InnerException) { $exception = $exception.InnerException }
            $transient = ($exception -is [UnauthorizedAccessException]) -or
                ($exception -is [IO.IOException] -and $exception -isnot [IO.FileNotFoundException] -and $exception -isnot [IO.DirectoryNotFoundException])
            if (-not $transient) { throw }
            if ($attempt -gt $RetryDelaySeconds.Count) { throw "Prepared file move failed after $attempt attempt(s): $Destination. $($exception.Message)" }
            Write-Warning ("Prepared file move refused (attempt {0}): {1}. {2} Retrying in {3}s." -f $attempt,$Destination,$exception.Message,$RetryDelaySeconds[$attempt-1])
            Start-Sleep -Seconds $RetryDelaySeconds[$attempt-1]
        }
    }
}

function Remove-PreparedOwnedPath([string]$Root, [string]$Relative) {
    # Only callers with an ownership marker/receipt may use this helper.
    $target = Get-PreparedChildPath $Root $Relative
    if (-not (Test-Path -LiteralPath $target)) { return }
    $cursor = $target
    $volumeRoot = [IO.Path]::GetPathRoot($target).TrimEnd('\','/')
    while ($cursor) {
        if ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Cleanup refused linked path: $cursor" }
        if ($cursor.TrimEnd('\','/') -eq $volumeRoot) { break }
        $cursor = Split-Path $cursor -Parent
    }
    if ((Get-Item -LiteralPath $target).PSIsContainer) {
        $linked = @(Get-ChildItem -LiteralPath $target -Recurse -Force | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint })
        if ($linked.Count) { throw "Cleanup refused linked content: $target" }
    }
    Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
}

function Clear-PreparedRunPayload([string]$WorkRoot, [string]$RunId, [string]$DataRoot, [string]$TenantKey) {
    if ($RunId -notmatch '^[a-f0-9]{32}$') { throw 'Invalid owned run ID.' }
    $run = Get-PreparedChildPath $WorkRoot $RunId
    $marker = (Read-SmartM365JsonDocument (Join-Path $run 'run.json')).Document
    if ($marker.Owner -ne 'PreparedEvidencePipeline/v1' -or $marker.RunId -ne $RunId -or $marker.DataRoot -ne $DataRoot -or $marker.TenantKey -ne $TenantKey) { throw 'Run ownership mismatch; no cleanup performed.' }
    foreach($name in 'run','capture','failure'){
        $path=Join-Path $run ($name+'.json')
        if(Get-SmartM365JsonReadPath $path -Optional){
            Resolve-SmartM365OwnedJsonPath -Path $path -Owner 'WorkplaceEvidence-Prepare/run' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Invalid prepared run audit.'}} | Out-Null
        }
    }
    foreach ($name in 'source','prepared','AccountClassification.psd1','AccountClassification.json.txt','AdCountryRules.json.txt','EndpointProtectionPolicy.json.txt','PreparedEndpointContract.json.txt','IntuneSourcePolicy.json.txt') { Remove-PreparedOwnedPath $WorkRoot "$RunId/$name" }
}

function Copy-PreparedStableSource {
    param($Entry, [string]$SnapshotRoot, [int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{},
        [ValidateRange(1,3)][int]$Attempts=3, [ValidateRange(0,2000)][int]$RetryDelayMs=1000,
        [Collections.Generic.List[object]]$OptionalSourceIssues)
    $optional = $Entry.PSObject.Properties['IsOptional'] -and $Entry.IsOptional
    $to = Get-PreparedChildPath $SnapshotRoot $Entry.Relative
    New-Item -ItemType Directory -Path (Split-Path $to -Parent) -Force | Out-Null
    for ($attempt=1; $attempt -le $Attempts; $attempt++) {
        $inputStream=$null; $outputStream=$null; $hash=$null; $verifyStream=$null; $sha=$null
        try {
            $before = Get-Item -LiteralPath $Entry.SourcePath -ErrorAction Stop
            if ($before.Length -eq 0) { throw 'Empty input file.' }
            if ($Entry.Kind -eq 'Current') {
                $limit = if ($optional) { [int]$Entry.MaxAgeHours } elseif ($AgeOverrides.ContainsKey($before.Name)) { [int]$AgeOverrides[$before.Name] } else { $MaxSourceAgeHours }
                if ($before.LastWriteTimeUtc -gt [datetime]::UtcNow.AddMinutes(5)) { throw 'Input publication-age limit or future timestamp violated.' }
                # An optional source older than its limit is still captured (already recorded as stale at selection).
                if ($limit -lt 1 -or (-not $optional -and ([datetime]::UtcNow-$before.LastWriteTimeUtc).TotalHours -gt $limit)) { throw 'Input publication-age limit or future timestamp violated.' }
            }
            # Do not lock out raw collectors. Verify the bytes against a fresh source
            # read after copying; retry only this file if it changed during capture.
            $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
            $inputStream = [IO.File]::Open($Entry.SourcePath,'Open','Read',$share)
            $outputStream = [IO.File]::Open($to,'Create','Write','None')
            $hash = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
            $buffer = [byte[]]::new(1048576)
            while (($count=$inputStream.Read($buffer,0,$buffer.Length)) -gt 0) { $outputStream.Write($buffer,0,$count); $hash.AppendData($buffer,0,$count) }
            $capturedHash = [Convert]::ToHexString($hash.GetHashAndReset())
            $inputStream.Dispose(); $inputStream=$null
            $outputStream.Dispose(); $outputStream=$null
            $verifyStream = [IO.File]::Open($Entry.SourcePath,'Open','Read',$share)
            $sha = [Security.Cryptography.SHA256]::Create()
            $sourceHash = [Convert]::ToHexString($sha.ComputeHash($verifyStream))
            $verifyStream.Dispose(); $verifyStream=$null
            $after = Get-Item -LiteralPath $Entry.SourcePath -ErrorAction Stop
            $copy = Get-Item -LiteralPath $to
            if ($before.Length -ne $after.Length -or $before.LastWriteTimeUtc -ne $after.LastWriteTimeUtc -or $copy.Length -ne $after.Length -or $capturedHash -ne $sourceHash -or (Get-FileHash -LiteralPath $to).Hash -ne $capturedHash) { throw 'Source changed during capture or copy integrity check failed.' }
            [IO.File]::SetLastWriteTimeUtc($to,$after.LastWriteTimeUtc)
            return [pscustomobject]@{Relative=$Entry.Relative;SourcePath=$Entry.SourcePath;Kind=$Entry.Kind;Bytes=$copy.Length;ModifiedUtc=$after.LastWriteTimeUtc.ToString('O');SHA256=$capturedHash;CapturedUtc=[datetime]::UtcNow.ToString('O');CaptureAttempts=$attempt}
        } catch {
            if ($attempt -eq $Attempts) {
                if (-not $optional) { throw "Capture failed after $attempt attempts: $($Entry.Relative): $($_.Exception.Message)" }
                $reason = $_.Exception.Message
                # Dispose before removing only this run-owned, partial snapshot file.
                if ($inputStream) { $inputStream.Dispose(); $inputStream=$null }
                if ($outputStream) { $outputStream.Dispose(); $outputStream=$null }
                Remove-PreparedOwnedPath $SnapshotRoot $Entry.Relative
                Add-PreparedOptionalSourceIssue $OptionalSourceIssues $Entry.Relative 'Capture' $reason
                return
            }
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Capture retry {1}/{2}: {3}: {4}' -f (Get-Date),($attempt+1),$Attempts,$Entry.Relative,$_.Exception.Message)
        } finally {
            foreach ($resource in $inputStream,$outputStream,$hash,$verifyStream,$sha) { if ($null -ne $resource) { $resource.Dispose() } }
        }
        if ($RetryDelayMs) { Start-Sleep -Milliseconds $RetryDelayMs }
    }
}

function Add-PreparedOptionalSourceIssue {
    param([Collections.Generic.List[object]]$Issues, [string]$Relative, [string]$Phase, [string]$Reason)
    if ($null -ne $Issues) {
        $Issues.Add([pscustomobject]@{Relative=$Relative;Phase=$Phase;Reason=$Reason;Utc=[datetime]::UtcNow.ToString('O')})
    }
    Write-Warning "Optional monitoring unavailable: $Relative ($Phase): $Reason. Other evidence can continue."
}

function Add-PreparedStaleOptionalSource {
    # The source stays in the batch; the record tells readers its data is older than the target age.
    param([Collections.Generic.List[object]]$Issues, [string]$Relative, [int]$LimitHours, [double]$AgeHours)
    $reason = 'Older than its {0} h target (age {1:N1} h); used as the latest available data.' -f $LimitHours, $AgeHours
    if ($null -ne $Issues) {
        $Issues.Add([pscustomobject]@{Relative=$Relative;Phase='StaleUsed';Reason=$reason;Utc=[datetime]::UtcNow.ToString('O')})
    }
    Write-Warning "Optional source used although stale: $Relative. $reason"
}

function Remove-PreparedMappingWorkbooks([Alias('DataRoot')][string]$WorkRoot, [string]$MappingRoot) {
    $relative = [IO.Path]::GetRelativePath([IO.Path]::GetFullPath($WorkRoot),[IO.Path]::GetFullPath($MappingRoot)).Replace('\','/')
    if ($relative -notmatch '^mappings/[a-f0-9]{32}$') { throw 'Mapping cleanup requires the exact run-owned download folder.' }
    Remove-PreparedOwnedPath $WorkRoot $relative
}

function Remove-PreparedObsoleteBatches([string]$OutputRoot, $Pointer) {
    # Compatibility entry for old callers: publication alone no longer authorizes payload retirement.
    # Coupled retention runs after cloud pointer verification, with a qualified consumer registry.
    Write-Verbose "Batch $($Pointer.BatchId) retained under $OutputRoot; coupled retention is a post-transfer operation."
}

function Undo-PreparedMappingReplacement {
    param([string]$DataRoot,[string]$MappingRoot,$Committed,[hashtable]$Previous,[hashtable]$Hashes)
    for ($index=$Committed.Count-1; $index -ge 0; $index--) {
        $name=$Committed[$index]; $target=Get-PreparedChildPath $DataRoot $name
        if ($Previous[$name]) {
            $backup=Join-Path (Join-Path $MappingRoot 'rollback') $name
            if ((Get-FileHash -LiteralPath $backup).Hash -ne $Previous[$name]) { throw "Mapping rollback copy changed: $name" }
            Move-PreparedFile -Source $backup -Destination $target -Overwrite
            if ((Get-FileHash -LiteralPath $target).Hash -ne $Previous[$name]) { throw "Mapping rollback hash mismatch: $name" }
        } elseif (Test-Path -LiteralPath $target) {
            if ((Get-FileHash -LiteralPath $target).Hash -ne $Hashes[$name]) { throw "Mapping rollback refused unexpected target: $name" }
            Remove-PreparedOwnedPath $DataRoot $name
        }
    }
    $Committed.Clear()
}

function Receive-PreparedMappingWorkbooks {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][scriptblock]$DownloadFile,
        [switch]$MailboxMigrationExceptionsEnabled, [string]$PythonCommand='python',
        [switch]$SharePointMigrationEnabled,
        [string]$SharePointMigrationWorkbookFileName='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx',
        [string]$SharePointMigrationBindingsFileName='')
    $data = (Resolve-Path -LiteralPath $DataRoot -ErrorAction Stop).ProviderPath
    if (-not (Test-Path -LiteralPath $data -PathType Container)) { throw 'Mapping DataRoot must be an existing directory.' }
    $mappingParent=Get-Item -LiteralPath (Join-Path $data 'mappings') -Force -ErrorAction SilentlyContinue
    if ($mappingParent -and (-not $mappingParent.PSIsContainer -or ($mappingParent.Attributes -band [IO.FileAttributes]::ReparsePoint))) { throw 'Mapping download parent must be an ordinary directory.' }
    # Serialize replacement with mapping downloads and the pipeline's raw-source capture.
    $lock = [IO.File]::Open((Join-Path $data '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
    $folder = Get-PreparedChildPath $data ('mappings/'+[guid]::NewGuid().ToString('N'))
    $committed = [Collections.Generic.List[string]]::new()
    $previous = @{}
    $hashes = @{}
    $mappingReady=$false; $keepRecovery=$false
    $required = [ordered]@{
        'SmartWorkplaceIntelligence-PersonaClassification.xlsx' = @('Personas','Rules','Exclusions')
        'SmartWorkplaceIntelligence-SiteClassification.xlsx' = @('Settings','SiteTypes','Sites')
    }
    if ($MailboxMigrationExceptionsEnabled) { $required['SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx'] = @('Exceptions','Guide') }
    if ($SharePointMigrationEnabled) {
        Assert-SharePointMigrationFileName $SharePointMigrationWorkbookFileName
        Assert-SharePointMigrationFileName $SharePointMigrationBindingsFileName -AllowEmpty
        if ($required.Contains($SharePointMigrationWorkbookFileName) -or $SharePointMigrationWorkbookFileName -notlike '*.xlsx') { throw 'Invalid or colliding migration workbook file name.' }
        $required[$SharePointMigrationWorkbookFileName]=@('Overview','MigrationSP','MigrationSharedFolders','Parameters','LegacyWebs','LegacyCollections')
        if ($SharePointMigrationBindingsFileName) {
            if ($required.Contains($SharePointMigrationBindingsFileName) -or $SharePointMigrationBindingsFileName -notlike '*.json.txt') { throw 'Invalid or colliding migration registry file name.' }
            $required[$SharePointMigrationBindingsFileName]=@()
        }
    }
    try {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    foreach ($name in $required.Keys) {
        $path = Join-Path $folder $name
        $receipt = & $DownloadFile $path $name
        if (-not $receipt -or -not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Item -LiteralPath $path).Length -eq 0) {
            throw "Required SharePoint mapping download failed: $name. No cached fallback is allowed."
        }
        if ($SharePointMigrationEnabled -and $SharePointMigrationBindingsFileName -and $name -ceq $SharePointMigrationBindingsFileName) {
            $null=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
            $hashes[$name]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            continue
        }
        $archive = $null
        try {
            $archive = [IO.Compression.ZipFile]::OpenRead($path)
            $entry = $archive.GetEntry('xl/workbook.xml')
            if (-not $entry -or $entry.Length -gt 1048576) { throw 'Missing or oversized workbook metadata.' }
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $settings = [Xml.XmlReaderSettings]::new()
            $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
            $settings.XmlResolver = $null
            $xmlReader = [Xml.XmlReader]::Create([IO.StringReader]::new($xml), $settings)
            $document = [Xml.XmlDocument]::new()
            $document.XmlResolver = $null
            try { $document.Load($xmlReader) } finally { $xmlReader.Dispose() }
            $sheets = @($document.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
            foreach ($sheet in $required[$name]) { if ($sheet -notin $sheets) { throw "Required worksheet missing: $sheet" } }
        } catch { throw "Invalid SharePoint mapping workbook '$name': $($_.Exception.Message)" }
        finally { if ($archive) { $archive.Dispose() } }
        if ($name -eq 'SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx') {
            $python = Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
            & $python.Source -B (Join-Path $PSScriptRoot 'workplace_mailbox_exceptions.py') --validate-workbook $path | ForEach-Object { Write-Information $_ -InformationAction Continue }
            if ($LASTEXITCODE -ne 0) { throw 'Invalid mailbox migration exception table. Retained mappings unchanged.' }
        }
        $hashes[$name] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    if ($SharePointMigrationEnabled) {
        $migrationValidation=@('-B',(Join-Path $PSScriptRoot 'workplace_spo_migration.py'),'--validate-workbook',(Join-Path $folder $SharePointMigrationWorkbookFileName))
        if($SharePointMigrationBindingsFileName){$migrationValidation+=@('--bindings',(Join-Path $folder $SharePointMigrationBindingsFileName))}
        $python=Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
        & $python.Source @migrationValidation | ForEach-Object { Write-Information $_ -InformationAction Continue }
        if($LASTEXITCODE -ne 0){throw 'Invalid migration workbook or approval registry. Retained mappings unchanged.'}
    }
    # Prepare replacements and rollback copies before touching retained mappings.
    $promotion = Join-Path $folder 'promotion'
    $rollback = Join-Path $folder 'rollback'
    New-Item -ItemType Directory -Path $promotion,$rollback -Force | Out-Null
    foreach ($name in $required.Keys) {
        $target = Get-PreparedChildPath $data $name
        $existing = Get-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
        if ($existing -and ($existing.PSIsContainer -or ($existing.Attributes -band [IO.FileAttributes]::ReparsePoint))) { throw "Mapping replacement refused nonregular target: $target" }
        $previous[$name] = if ($existing) { (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash } else { $null }
        if ($existing) {
            $backup = Join-Path $rollback $name
            Copy-Item -LiteralPath $target -Destination $backup -ErrorAction Stop
            if ((Get-FileHash -LiteralPath $backup).Hash -ne $previous[$name]) { throw "Mapping rollback copy changed: $name" }
        }
        $replacement = Join-Path $promotion $name
        Copy-Item -LiteralPath (Join-Path $folder $name) -Destination $replacement -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $replacement).Hash -ne $hashes[$name]) { throw "Validated mapping changed before replacement: $name" }
    }
    try {
        foreach ($name in $required.Keys) {
            $target = Get-PreparedChildPath $data $name
            if ($previous[$name]) {
                if ((Get-FileHash -LiteralPath $target).Hash -ne $previous[$name]) { throw "Retained mapping changed before replacement: $name" }
            } elseif (Test-Path -LiteralPath $target) { throw "Retained mapping appeared before replacement: $name" }
            # Record the intent before the rename so Ctrl+C between statements can roll it back.
            $committed.Add($name)
            Move-PreparedFile -Source (Join-Path $promotion $name) -Destination $target -Overwrite
            if ((Get-FileHash -LiteralPath $target).Hash -ne $hashes[$name]) { throw "Retained mapping hash mismatch: $name" }
        }
    } catch {
        $replacementFailure = $_
        try {
            Undo-PreparedMappingReplacement -DataRoot $data -MappingRoot $folder -Committed $committed -Previous $previous -Hashes $hashes
        } catch {
            $recovery=[InvalidOperationException]::new("Mapping replacement and rollback failed. Recovery files retained at $folder. $($_.Exception.Message)",$replacementFailure.Exception)
            $recovery.Data['KeepMappingRecovery']=$true
            $keepRecovery=$true
            throw $recovery
        }
        throw $replacementFailure
    }
    Remove-PreparedOwnedPath $folder 'promotion'
    Remove-PreparedOwnedPath $folder 'rollback'
    # Retained root files survive cleanup. This run consumes its immutable DATA pair.
    $mappingReady=$true
    $folder
    } catch {
        $failure=$_
        if (-not $failure.Exception.Data['KeepMappingRecovery']) {
            try { Remove-PreparedMappingWorkbooks -DataRoot $data -MappingRoot $folder } catch { Write-Warning "Mapping cleanup failed: $($_.Exception.Message)" }
        }
        throw $failure
    } finally {
        try {
            if (-not $mappingReady -and -not $keepRecovery -and (Test-Path -LiteralPath $folder)) {
                try {
                    if ($committed.Count) { Undo-PreparedMappingReplacement -DataRoot $data -MappingRoot $folder -Committed $committed -Previous $previous -Hashes $hashes }
                    Remove-PreparedMappingWorkbooks -DataRoot $data -MappingRoot $folder
                } catch { Write-Warning "Mapping interruption cleanup incomplete; recovery retained at $folder. $($_.Exception.Message)" }
            }
        } finally { $lock.Dispose() }
    }
}

function Get-PreparedSourcePlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot,
        [string]$SourceContractPath = (Join-Path $script:ProductRoot 'config/prepared-source-contract.json'),
        [ValidateRange(1,8760)][int]$MaxSourceAgeHours = 168,
        [hashtable]$AgeOverrides = @{}, [string]$MappingRoot, [switch]$MetadataOnly,
        [switch]$MailboxMigrationExceptionsEnabled,
        [switch]$SharePointMigrationEnabled,
        [string]$SharePointMigrationWorkbookFileName='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx',
        [string]$SharePointMigrationBindingsFileName='',
        [Collections.Generic.List[object]]$OptionalSourceIssues)
    $root = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
    $spec = (Read-SmartM365JsonDocument $SourceContractPath).Document
    $items = [Collections.Generic.List[object]]::new()
    foreach ($name in $spec.currentFiles) { $items.Add(@{Relative="DATA-LAST/$name";Kind='Current'}) }
    foreach ($entry in $(if($spec.PSObject.Properties['governedArtifacts'] -and $spec.governedArtifacts){Get-PreparedArtifactSources -DataRoot $root})) {
        if ($entry.Relative -notin @($items | ForEach-Object { $_.Relative })) { $items.Add($entry) }
    }
    foreach ($entry in $(if($spec.PSObject.Properties['nativeEvidence'] -and $spec.nativeEvidence){Get-NativeEvidenceSourceItems -DataRoot $root})) {
        # Overlap is captured once. Native receipt age is additional to an existing
        # Intelligence publication-age rule; neither contract is weakened.
        if ($entry.Relative -notin @($items | ForEach-Object { $_.Relative })) { $items.Add($entry) }
    }
    if ($spec.PSObject.Properties['optionalSourceGroups']) {
        if ($spec.optionalSourceGroups -ne 'intune-operational-contract.json.txt') { throw 'Unknown grouped Intune source contract.' }
        foreach ($group in Get-IntuneOperationalSourceGroups -LastRoot (Join-Path $root 'DATA-LAST') -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides | Where-Object Available) {
            $items.Add(@{Relative=('DATA-LAST/'+[IO.Path]::GetFileName($group.ReceiptPath));Kind='Receipt'})
            foreach ($file in $group.Files) { $items.Add(@{Relative=('DATA-LAST/'+$file.Spec.file);Kind='Current'}) }
        }
    }
    if ($spec.PSObject.Properties['optionalCurrentFiles']) {
        $optionalNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $spec.optionalCurrentFiles) {
            if ([string]::IsNullOrWhiteSpace($entry.file) -or -not $optionalNames.Add($entry.file) -or $entry.file -in $spec.currentFiles -or [int]$entry.maxAgeHours -lt 1 -or [int]$entry.maxAgeHours -gt 8760) { throw 'Invalid or duplicate optional source contract.' }
            $items.Add(@{Relative="DATA-LAST/$($entry.file)";Kind='Current';IsOptional=$true;MaxAgeHours=[int]$entry.maxAgeHours})
        }
    }
    foreach ($name in $spec.mappingFiles) { $items.Add(@{Relative=$name;Kind='Mapping'}) }
    if ($MailboxMigrationExceptionsEnabled) {
        if (-not $spec.PSObject.Properties['mailboxMigrationExceptionsFile'] -or $spec.mailboxMigrationExceptionsFile -cne 'SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx') { throw 'Missing reviewed mailbox migration exception mapping contract.' }
        $items.Add(@{Relative=$spec.mailboxMigrationExceptionsFile;Kind='Mapping'})
    }
    if ($SharePointMigrationEnabled) {
        Assert-SharePointMigrationFileName $SharePointMigrationWorkbookFileName
        Assert-SharePointMigrationFileName $SharePointMigrationBindingsFileName -AllowEmpty
        if(-not $spec.PSObject.Properties['sharePointMigrationContract'] -or $spec.sharePointMigrationContract -cne 'sharepoint-migration-contract.json.txt'){throw 'Missing migration source group contract.'}
        if($SharePointMigrationWorkbookFileName -in $spec.mappingFiles -or $SharePointMigrationBindingsFileName -in $spec.mappingFiles){throw 'Migration inputs collide with classification workbooks.'}
        $items.Add(@{Relative=$SharePointMigrationWorkbookFileName;Kind='Mapping'})
        if($SharePointMigrationBindingsFileName){$items.Add(@{Relative=$SharePointMigrationBindingsFileName;Kind='Mapping'})}
        foreach($entry in Get-SharePointMigrationSourceItems -DataRoot $root){
            if($entry.Relative -notin @($items | ForEach-Object {$_.Relative})){$items.Add($entry)}
        }
    }
    foreach ($name in $spec.dailyFiles) { $items.Add(@{Relative=$name;Kind='History'}) }
    foreach ($entry in $spec.history) {
        $folder = Get-PreparedChildPath $root $entry.root
        $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse -Filter $entry.file -ErrorAction Stop |
            Where-Object { $_.FullName -match '[\\/]WeeklyHistory[\\/]\d{4}-W\d{2}[\\/]' })
        if (-not $files.Count) { throw "Required observed history missing: $($entry.root)/$($entry.file)" }
        foreach ($file in $files) { $items.Add(@{Relative=[IO.Path]::GetRelativePath($root,$file.FullName);Kind='History'}) }
    }
    if ($spec.PSObject.Properties['archivedHistory']) {
        foreach ($entry in $spec.archivedHistory) {
            if ([string]::IsNullOrWhiteSpace($entry.root) -or
                [IO.Path]::GetFileName([string]$entry.file) -cne [string]$entry.file -or
                [IO.Path]::GetFileName([string]$entry.manifest) -cne [string]$entry.manifest -or
                [string]::IsNullOrWhiteSpace($entry.file) -or
                [string]::IsNullOrWhiteSpace($entry.manifest)) { throw 'Invalid archived history source contract.' }
            $folder = Get-PreparedChildPath $root $entry.root
            if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
            foreach ($weekFolder in Get-ChildItem -LiteralPath $folder -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$') {
                foreach ($name in @($entry.file,$entry.manifest)) {
                    $file = Get-PreparedChildPath $weekFolder.FullName $name
                    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Archived history input missing: $file" }
                    $kind = if ($name -ceq $entry.file) { 'ArchivedHistory' } else { 'ArchiveManifest' }
                    $items.Add(@{Relative=[IO.Path]::GetRelativePath($root,$file);Kind=$kind})
                }
            }
        }
    }
    # Windows history discovery returns backslashes; contracts use slashes.
    # Normalize before de-duplication so one physical input is captured once.
    foreach ($entry in $items) { $entry.Relative = $entry.Relative.Replace('\','/') }
    foreach ($item in $items | Sort-Object Relative -Unique) {
        $sourceRoot = if ($item.Kind -eq 'Mapping' -and $MappingRoot) { $MappingRoot } else { $root }
        $path = Get-PreparedChildPath $sourceRoot $item.Relative
        $optional = $item.ContainsKey('IsOptional') -and $item.IsOptional
        try {
            $file = Get-Item -LiteralPath $path -ErrorAction Stop
            if ($file.Length -eq 0) { throw "Empty input file: $($item.Relative)" }
            # Transport freshness only; business report timestamps remain in Data Trust.
            if ($item.Kind -eq 'Current') {
                $limit = if ($optional) { [int]$item.MaxAgeHours } elseif ($AgeOverrides.ContainsKey($file.Name)) { [int]$AgeOverrides[$file.Name] } else { $MaxSourceAgeHours }
                if ($file.LastWriteTimeUtc -gt [datetime]::UtcNow.AddMinutes(5)) { throw "Future input timestamp: $($item.Relative)" }
                $age = ([datetime]::UtcNow - $file.LastWriteTimeUtc).TotalHours
                if ($limit -lt 1 -or $age -gt $limit) {
                    # Optional sources keep their latest available publication, however old: it is
                    # used and recorded as stale. Required sources keep their hard age limit.
                    if (-not $optional -or $limit -lt 1) { throw "Input exceeds configured publication-age limit: $($item.Relative)" }
                    Add-PreparedStaleOptionalSource $OptionalSourceIssues $item.Relative $limit $age
                }
            }
            [pscustomobject]@{Relative=$item.Relative;SourcePath=$file.FullName;Kind=$item.Kind;Bytes=$file.Length;ModifiedUtc=$file.LastWriteTimeUtc.ToString('O');SHA256=$(if($MetadataOnly){$null}else{(Get-FileHash -LiteralPath $path).Hash});IsOptional=[bool]$optional;MaxAgeHours=$(if($optional){$item.MaxAgeHours}else{$null})}
        } catch {
            if (-not $optional) { throw }
            Add-PreparedOptionalSourceIssue $OptionalSourceIssues $item.Relative 'Selection' $_.Exception.Message
        }
    }
}

function Test-PreparedTenant {
    param([string]$Path, [string]$TenantKey, [switch]$AllowMissingTenantKey)
    # Match the source formats consumed by the evidence generators; never retry
    # malformed records with another separator or weaken the tenant checks.
    $fileName = [IO.Path]::GetFileName($Path)
    $delimiter = if ($fileName -like '*DailyStats.csv' -or $fileName -like 'SharePoint_OnPrem_*.csv' -or $fileName -eq 'Exchange_OnPrem_Servers_Inventory.csv') { ';' } else { ',' }
    $artifactContract = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json
    $optional = @(if ($artifactContract.PSObject.Properties['optionalCurrentFiles']) { $artifactContract.optionalCurrentFiles | Where-Object { [string]$_.file -ceq $fileName } })
    if ($optional.Count -gt 1) { throw 'Duplicate optional source delimiter declaration.' }
    if ($optional.Count -eq 1 -and $optional[0].PSObject.Properties['delimiter']) {
        if ([string]$optional[0].delimiter -cnotin @(',', ';')) { throw 'Invalid optional source delimiter.' }
        $delimiter = [string]$optional[0].delimiter
    }
    # Compiled strict reader first; any other construct and every failure are re-read below.
    $fast = [Workplace.Prepared.StrictCsv]::CheckTenant($Path, [char]$delimiter, $TenantKey, [bool]$AllowMissingTenantKey)
    if ($null -ne $fast) { return [pscustomobject]@{Rows=$fast.Rows;HasTenantKey=$fast.HasTenantKey} }
    $script:PreparedTenantReferenceParses++
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.SetDelimiters([string]$delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $header = $parser.ReadFields()
        if (-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count) { throw 'Empty or duplicate CSV header.' }
        $index = [array]::IndexOf($header,'TenantKey')
        if ($index -lt 0 -and -not $AllowMissingTenantKey) { throw 'TenantKey column missing.' }
        $rows = 0L
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            $rows++
            if ($fields.Length -ne $header.Length) { throw "Malformed CSV at record $rows." }
            if ($index -ge 0 -and ([string]::IsNullOrWhiteSpace($fields[$index]) -or $fields[$index] -ne $TenantKey)) {
                throw "Empty or incompatible TenantKey at record $rows."
            }
        }
        [pscustomobject]@{Rows=$rows;HasTenantKey=($index -ge 0)}
    } finally { $parser.Dispose() }
}

function Test-PreparedSourceTenants {
    param([object[]]$Plan, [string]$TenantKey, [string]$SnapshotRoot)
    $errors = [Collections.Generic.List[string]]::new()
    $rows = 0L; $files = 0; $legacyArchived = 0
    $script:PreparedTenantReferenceParses = 0
    foreach ($entry in $Plan | Where-Object { $_.Relative -like '*.csv' }) {
        $relative = $entry.Relative.Replace('\','/')
        $path = if ($SnapshotRoot) { Get-PreparedChildPath $SnapshotRoot $relative } else { $entry.SourcePath }
        try {
            if ($entry.Kind -eq 'ArchivedHistory') {
                if ($relative -notmatch '^DATA-ALL/Intune/Applications/DiscoveredApps/SummaryHistory/(?<Week>\d{4}-W\d{2})/Intune_DiscoveredApps_Summary\.csv$') {
                    throw 'Unrecognized legacy archived summary path.'
                }
                $week = $Matches.Week
                $manifestRelative = $relative.Substring(0,$relative.LastIndexOf('/')+1) + 'summary.manifest.json.txt'
                if (@($Plan | Where-Object Relative -CEQ $manifestRelative).Count -ne 1) { throw 'Archived summary manifest was not captured.' }
                $manifestPath = if ($SnapshotRoot) { Get-PreparedChildPath $SnapshotRoot $manifestRelative } else { (@($Plan | Where-Object Relative -CEQ $manifestRelative)[0]).SourcePath }
                $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                if ($manifest.Week -cne $week -or $manifest.FileName -cne [IO.Path]::GetFileName($path) -or
                    $manifest.SHA256 -cne (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash) {
                    throw 'Archived summary provenance does not match its manifest.'
                }
                $check = Test-PreparedTenant -Path $path -TenantKey $TenantKey -AllowMissingTenantKey
                if (-not $check.HasTenantKey) { $legacyArchived++ }
            } else {
                $check = Test-PreparedTenant -Path $path -TenantKey $TenantKey
            }
            $rows += $check.Rows; $files++
        } catch { $errors.Add("${relative}: $($_.Exception.Message)") }
        if ($files -gt 0 -and $files % 25 -eq 0) { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Source tenant validation: {1} CSVs checked.' -f (Get-Date),$files) }
    }
    Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Source tenant validation: {1} CSVs checked; {2} read by the reference parser.' -f (Get-Date),$files,$script:PreparedTenantReferenceParses)
    if ($errors.Count) { throw ("Source tenant validation failed ($($errors.Count) files):`n" + ($errors -join "`n")) }
    [pscustomobject]@{CsvFiles=$files;Rows=$rows;LegacyArchivedSummaries=$legacyArchived}
}

function Publish-PreparedEvidenceBatch {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$StagingRoot,
        [Parameter(Mandatory)][string]$OutputRoot,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_.-]+$')][string]$TenantKey,
        [Parameter(Mandatory)]$Provenance,
        [string]$ContractPath = (Join-Path $script:ProductRoot 'config/prepared-evidence-contract.json'),
        [string[]]$AllowEmptyTables = @(),
        [string]$ExpectedPreviousBatchId,
        [string]$PreservePublishedFilesFrom,
        [string[]]$ReplacePublishedFiles = @(),
        [string]$ArtifactSnapshotRoot,
        [ValidateRange(0,520)][int]$HistoryWindowWeeks=0,
        [switch]$SkipRetention)
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
    $output = [IO.Path]::GetFullPath($OutputRoot)
    $staging = (Resolve-Path -LiteralPath $StagingRoot).ProviderPath
    if ($output.TrimEnd('\') -eq $staging.TrimEnd('\')) { throw 'Staging and publication roots must differ.' }
    if ((Split-Path $output -Leaf) -ne 'DATA-POWERBI') { throw 'Publication root must be a dedicated DATA-POWERBI directory.' }
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    Initialize-PreparedMetadataNames -OutputRoot $output -TenantKey $TenantKey
    # Held throughout validation and commit; a failed process releases the OS lock.
    $lock = [IO.File]::Open((Join-Path $output '.publication.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    $batch=$null; $committed=$false; $validationRoot=$null
    try {
        $currentPath = Join-Path $output 'current.json.txt'
        if (Test-Path -LiteralPath (Join-Path $output 'current.json')) { throw 'Legacy prepared metadata is preserved while JSON transport policy is Readers. Activate the approved JsonText deployment before publication.' }
        $previousPath = Get-PreparedMetadataPath $output 'current' -Optional
        $previous = if ($previousPath) { Get-Content -LiteralPath $previousPath -Raw | ConvertFrom-Json } else { $null }
        if ($ExpectedPreviousBatchId -and (($ExpectedPreviousBatchId -eq '__NONE__' -and $previous) -or ($ExpectedPreviousBatchId -ne '__NONE__' -and (-not $previous -or $previous.BatchId -ne $ExpectedPreviousBatchId)))) { throw 'Published batch changed since staging; no publication performed.' }
        if ($previous -and $previous.TenantKey -ne $TenantKey) { throw 'Output root belongs to another tenant.' }
        $contractReceipt = Read-SmartM365JsonDocument $ContractPath
        $ContractPath = $contractReceipt.Path
        $contract = $contractReceipt.Document
        $preservedFiles = @{}
        $previousPointerHash = if ($previousPath) { (Get-FileHash -LiteralPath $previousPath).Hash } else { $null }
        if ($ReplacePublishedFiles.Count -and -not $PreservePublishedFilesFrom) { throw 'Targeted replacements require predecessor preservation.' }
        if (@($ReplacePublishedFiles | Sort-Object -Unique).Count -ne $ReplacePublishedFiles.Count) { throw 'Duplicate targeted replacement file.' }
        if ($PreservePublishedFilesFrom) {
            if (-not $previous -or -not $ExpectedPreviousBatchId) { throw 'Preserving published files requires an expected current batch.' }
            $previousBatchPath = Get-PreparedChildPath $output "batches/$($previous.BatchId)"
            $preservePath = (Resolve-Path -LiteralPath $PreservePublishedFilesFrom).ProviderPath
            if ([IO.Path]::GetFullPath($preservePath).TrimEnd('\') -ne [IO.Path]::GetFullPath($previousBatchPath).TrimEnd('\')) { throw 'Only the current predecessor batch can be preserved.' }
            $preserveManifestPath = Join-Path $preservePath 'batch.json.txt'
            if ((Get-FileHash -LiteralPath $preserveManifestPath).Hash -ne $previous.ManifestSHA256) { throw 'Predecessor manifest hash mismatch.' }
            $preserveManifest = Get-Content -LiteralPath $preserveManifestPath -Raw | ConvertFrom-Json
            if ($preserveManifest.BatchId -ne $previous.BatchId -or $preserveManifest.TenantKey -ne $TenantKey) { throw 'Predecessor manifest identity mismatch.' }
            foreach ($entry in $preserveManifest.Files) {
                if ($entry.File -notin $contract.tables.file -or $preservedFiles.ContainsKey($entry.File)) { throw 'Predecessor file is absent from the contract or duplicated.' }
                $null = Get-PreparedChildPath $preservePath $entry.File
                $preservedFiles[$entry.File] = $entry
            }
            foreach ($name in $ReplacePublishedFiles) {
                if (-not $preservedFiles.ContainsKey($name)) { throw 'Targeted replacement is absent from predecessor manifest.' }
                $entry = $preservedFiles[$name]
                $source = Get-PreparedChildPath $preservePath $name
                if ((Get-Item -LiteralPath $source).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $source).Hash -ne $entry.SHA256) { throw "Predecessor file hash or size mismatch: $name" }
                $preservedFiles.Remove($name)
            }
        }
        foreach ($name in $AllowEmptyTables) { if ($name -notin $contract.tables.table) { throw "Unknown empty-table exception: $name" } }
        $batchId = [datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
        $batch = Get-PreparedChildPath $output "batches/$batchId"
        New-Item -ItemType Directory -Path $batch -Force | Out-Null
        # Build and validate beside the captured staging, not over SMB. The
        # destination bytes are checked against this validation before commit.
        $validationRelative='.publication-'+$batchId
        $validationRoot=Get-PreparedChildPath $staging $validationRelative
        New-Item -ItemType Directory -Path $validationRoot -ErrorAction Stop | Out-Null
        Write-PreparedJson (Join-Path $validationRoot '.owner.json.txt') @{BatchId=$batchId;TenantKey=$TenantKey}
        foreach ($table in $contract.tables) {
            $fileTimer=[Diagnostics.Stopwatch]::StartNew()
            $destination = Get-PreparedChildPath $validationRoot $table.file
            if ($preservedFiles.ContainsKey($table.file)) {
                $entry = $preservedFiles[$table.file]
                $source = Get-PreparedChildPath $preservePath $table.file
                if ((Get-Item -LiteralPath $source).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $source).Hash -ne $entry.SHA256) { throw "Predecessor file hash or size mismatch: $($table.file)" }
                Copy-Item -LiteralPath $source -Destination $destination
                if ((Get-FileHash -LiteralPath $destination).Hash -ne $entry.SHA256 -or (Get-FileHash -LiteralPath $source).Hash -ne $entry.SHA256) { throw "Preserved file changed during publication: $($table.file)" }
                Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Assemble preserved {1}: {2:F2}s' -f (Get-Date),$table.file,$fileTimer.Elapsed.TotalSeconds)
                continue
            }
            $source = Get-PreparedChildPath $staging $table.file
            $sourceHash = (Get-FileHash -LiteralPath $source).Hash
            # Retain the original header for a valid header-only empty export.
            $reader = [IO.StreamReader]::new($source)
            $temp = $destination + '.tenant.tmp'
            $stagingTenant=$false
            if ($table.PSObject.Properties['stagingIncludesTenantKey']) {
                if ($table.stagingIncludesTenantKey -isnot [bool]) {throw 'stagingIncludesTenantKey must be a JSON boolean.'}
                $stagingTenant=$table.stagingIncludesTenantKey
            }
            try {
                $header = $reader.ReadLine()
                if (-not $stagingTenant -and $header -match '^"?TenantKey"?,') { throw 'Staging must not contain a published TenantKey prefix.' }
                if ($stagingTenant -and $header -notmatch '^"?TenantKey"?,') {throw 'Tenant-aware staging must start with TenantKey.'}
            } finally { $reader.Dispose() }
            if ($stagingTenant) {
                # Only explicitly declared tenant-aware tables may retain this prefix.
                Copy-Item -LiteralPath $source -Destination $temp
            } else {
                # Prefix legacy tenantless staging locally; native tables retain their prefix.
                Import-Csv -LiteralPath $source | Select-Object @{n='TenantKey';e={$TenantKey}},* |
                    Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding utf8
                # Header-only files need explicit schema preservation.
                if ((Get-Item -LiteralPath $temp).Length -eq 0) {
                    [IO.File]::WriteAllText($temp, '"TenantKey",' + $header + [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
                }
            }
            if ((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash) { throw "Staging changed during publication: $($table.file)" }
            # The batch file does not exist yet: a plain rename, never a replacement of a file just written.
            Move-PreparedFile -Source $temp -Destination $destination
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Assemble {1}: {2:F2}s' -f (Get-Date),$table.file,$fileTimer.Elapsed.TotalSeconds)
        }
        $validationPath = Join-Path $batch 'validation.json.txt'
        $localValidationPath=Join-Path $validationRoot 'validation.json.txt'
        & (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $validationRoot -ContractPath $ContractPath -ReportPath $localValidationPath -TenantKey $TenantKey -LogTiming | ForEach-Object {
            if ($_ -like '[[]*') { Write-Host $_ }
        }
        $validationExit=$LASTEXITCODE
        if (Test-Path -LiteralPath $localValidationPath) { Copy-Item -LiteralPath $localValidationPath -Destination $validationPath }
        if ($validationExit -ne 0) { throw 'Batch validation failed; current.json was not changed.' }
        $validation = Get-Content -LiteralPath $validationPath -Raw | ConvertFrom-Json
        foreach ($entry in $validation.Files) {
            $table = $contract.tables | Where-Object file -EQ $entry.File
            $contractAllowsEmpty = $table.PSObject.Properties['allowEmpty'] -and $table.allowEmpty -is [bool] -and $table.allowEmpty
            if ($entry.Rows -eq 0 -and -not $contractAllowsEmpty -and $table.table -notin $AllowEmptyTables) { throw "Unexpected empty output: $($entry.File)" }
        }
        foreach ($entry in $validation.Files) {
            $copyTimer=[Diagnostics.Stopwatch]::StartNew()
            $source=Get-PreparedChildPath $validationRoot $entry.File
            $destination=Get-PreparedChildPath $batch $entry.File
            Copy-Item -LiteralPath $source -Destination $destination
            $copySeconds=$copyTimer.Elapsed.TotalSeconds
            if ((Get-FileHash -LiteralPath $destination).Hash -ne $entry.SHA256 -or (Get-FileHash -LiteralPath $source).Hash -ne $entry.SHA256) { throw "Validated file changed during publication: $($entry.File)" }
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Publish copy {1}: copy={2:F2}s; hash verification={3:F2}s' -f (Get-Date),$entry.File,$copySeconds,($copyTimer.Elapsed.TotalSeconds-$copySeconds))
        }
        $historyTables = @($contract.tables | Where-Object { $_.table -match 'History|Trend' })
        $historyKeyVersions = [ordered]@{}
        $historyKeys = @{}
        foreach ($table in $historyTables) {
            # The comparison key is declared per table: stable period and dimension columns only,
            # never measures or provisional dates that the next run legitimately replaces.
            if (-not $table.PSObject.Properties['historyKey'] -or @($table.historyKey).Count -eq 0) { throw "No historical comparison key configured: $($table.table)" }
            foreach ($column in @($table.historyKey)) { if ($column -notin @($table.columns.name)) { throw "Unknown historical key column in $($table.table): $column" } }
            $historyKeys[$table.table] = @($table.historyKey)
            $version = 1
            if ($table.PSObject.Properties['historyKeyVersion']) {
                if (-not [int]::TryParse([string]$table.historyKeyVersion,[ref]$version) -or $version -lt 1) { throw "Invalid historyKeyVersion: $($table.table)" }
            }
            $historyKeyVersions[$table.table] = $version
        }
        $historyKeyResets = [Collections.Generic.List[object]]::new()
        $historyKeyInitializations = [Collections.Generic.List[object]]::new()
        $historyWindowExpirations = [Collections.Generic.List[object]]::new()
        # Rolling window of the weekly sources (WeeklyHistoryMaxRetentionWeeks); 0 keeps every week.
        $historyCutoffWeek = $null
        if ($HistoryWindowWeeks -gt 0) {
            $cutoffDate = [datetime]::UtcNow.Date.AddDays(-7 * $HistoryWindowWeeks)
            $historyCutoffWeek = '{0:D4}-W{1:D2}' -f [Globalization.ISOWeek]::GetYear($cutoffDate), [Globalization.ISOWeek]::GetWeekOfYear($cutoffDate)
        }
        if ($previous) {
            $previousManifestPath = Get-PreparedChildPath $output "batches/$($previous.BatchId)/batch.json.txt"
            if (-not (Test-Path -LiteralPath $previousManifestPath -PathType Leaf)) { throw "Previous batch manifest is missing: $($previous.BatchId)" }
            $previousManifest = Get-Content -LiteralPath $previousManifestPath -Raw | ConvertFrom-Json
            # Loss of an observed historical key is never silently accepted.
            foreach ($table in $historyTables) {
                $previousFile = @($previousManifest.Files | Where-Object { $_.File -eq $table.file })
                if ($previousFile.Count -eq 0) {
                    $historyKeyInitializations.Add([ordered]@{Table=$table.table;PreviousBatchId=$previous.BatchId;Version=$historyKeyVersions[$table.table]})
                    Write-Information "Historical comparison initialized: $($table.table) (previous batch $($previous.BatchId))." -InformationAction Continue
                    continue
                }
                # A contract key-version change (new key grain) is the only accepted reset, and it is recorded.
                $previousVersion = 1
                if ($previousManifest.PSObject.Properties['HistoryKeyVersions'] -and $previousManifest.HistoryKeyVersions.PSObject.Properties[$table.table]) {
                    $previousVersion = [int]$previousManifest.HistoryKeyVersions.($table.table)
                }
                $version = $historyKeyVersions[$table.table]
                if ($version -lt $previousVersion) { throw "Historical key version cannot go back: $($table.table) ($previousVersion -> $version)" }
                if ($version -gt $previousVersion) {
                    $historyKeyResets.Add([ordered]@{Table=$table.table;PreviousBatchId=$previous.BatchId;PreviousVersion=$previousVersion;Version=$version})
                    Write-Warning "Historical key comparison reset by contract: $($table.table) key version $previousVersion -> $version (previous batch $($previous.BatchId))."
                    continue
                }
                $oldPath = Get-PreparedChildPath $output "batches/$($previous.BatchId)/$($table.file)"
                $newPath = Get-PreparedChildPath $batch $table.file
                $newKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $keyColumns = $historyKeys[$table.table]
                $keyOf = { param($row) ($keyColumns | ForEach-Object { $v=[string]$row.$_; $v.Length.ToString()+':'+$v }) -join '|' }
                Import-Csv -LiteralPath $newPath | ForEach-Object { $null=$newKeys.Add((& $keyOf $_)) }
                $expiredWeeks = [Collections.Generic.SortedSet[string]]::new([StringComparer]::Ordinal)
                Import-Csv -LiteralPath $oldPath | ForEach-Object {
                    if (-not $newKeys.Contains((& $keyOf $_))) {
                        $oldRow = $_
                        # Only a week at or before the rolling window cutoff may leave the history.
                        $expiredWeek = if ($historyCutoffWeek) { @($keyColumns | ForEach-Object { [string]$oldRow.$_ } | Where-Object { $_ -match '^\d{4}-W\d{2}$' -and [string]::CompareOrdinal($_, $historyCutoffWeek) -le 0 }) | Select-Object -First 1 }
                        if ($expiredWeek) { $null=$expiredWeeks.Add($expiredWeek); return }
                        throw ("Historical coverage would regress: {0} (missing {1})" -f $table.table, (($keyColumns | ForEach-Object { '{0}={1}' -f $_, $oldRow.$_ }) -join ', '))
                    }
                }
                if ($expiredWeeks.Count -gt 0) {
                    $historyWindowExpirations.Add([ordered]@{Table=$table.table;PreviousBatchId=$previous.BatchId;WindowWeeks=$HistoryWindowWeeks;CutoffWeek=$historyCutoffWeek;Weeks=@($expiredWeeks)})
                    Write-Information ("Historical weeks expired by the {0}-week window: {1} ({2})." -f $HistoryWindowWeeks, $table.table, (@($expiredWeeks) -join ', ')) -InformationAction Continue
                }
            }
        }
        $artifacts=@()
        $licenseQualification=$null
        if($ArtifactSnapshotRoot){
            $artifacts=@(Copy-PreparedArtifacts -SnapshotRoot $ArtifactSnapshotRoot -Destination $batch -TenantKey $TenantKey)
            Assert-PreparedArtifacts -BatchRoot $batch -Entries $artifacts -TenantKey $TenantKey
            # License tables carry no history key: a Not qualified report is an explicit manifest state,
            # never a historical coverage regression and never a failed batch.
            $licenseQualification=Get-PreparedLicenseReportQualification -SnapshotRoot $ArtifactSnapshotRoot -TenantKey $TenantKey -Artifacts $artifacts -RequireCurrentSources
            if($licenseQualification.State -eq 'Qualified'){Write-Information ("License report qualified by licensing receipt run {0} (snapshot {1})." -f $licenseQualification.ReceiptRunId,$licenseQualification.SnapshotId) -InformationAction Continue}
            else{Write-Warning ("License report Not qualified; license tables are published empty. Reason: {0}" -f $licenseQualification.Reason)}
        }
        $manifest = [ordered]@{SchemaVersion=1;BatchId=$batchId;TenantKey=$TenantKey;CreatedUtc=[datetime]::UtcNow.ToString('O');Provenance=$Provenance;ContractSHA256=(Get-FileHash $ContractPath).Hash;HistoryKeyVersions=$historyKeyVersions;HistoryKeyResets=@($historyKeyResets);HistoryKeyInitializations=@($historyKeyInitializations);HistoryWindowExpirations=@($historyWindowExpirations);Files=@($validation.Files | ForEach-Object { @{File=$_.File;Rows=$_.Rows;Bytes=(Get-Item (Join-Path $batch $_.File)).Length;SHA256=$_.SHA256} })}
        if($ArtifactSnapshotRoot){$manifest.Artifacts=$artifacts;$manifest.LicenseReportQualification=$licenseQualification;$manifest.ArtifactContractSHA256=(Get-FileHash -LiteralPath (Join-Path $script:ProductRoot 'config/prepared-artifact-contract.json.txt')).Hash}
        Write-PreparedJson (Join-Path $batch 'batch.json.txt') $manifest
        $pointer = @{SchemaVersion=1;BatchId=$batchId;TenantKey=$TenantKey;ManifestSHA256=(Get-FileHash (Join-Path $batch 'batch.json.txt')).Hash;PreviousBatchId=if($previous){$previous.BatchId}else{$null}}
        $pointerTemp = Join-Path $output ('current.'+$batchId+'.tmp')
        Write-PreparedJson $pointerTemp $pointer
        # Immutable copy used by a delayed/cloud transfer; never upload a newer run's pointer.
        Write-PreparedJson (Join-Path $batch 'current.json.txt') $pointer
        if ($PreservePublishedFilesFrom -and (Get-FileHash -LiteralPath $currentPath).Hash -ne $previousPointerHash) { throw 'Current pointer changed during preserved-file validation; no publication performed.' }
        Move-PreparedFile -Source $pointerTemp -Destination $currentPath -Overwrite
        $committed=$true
        if (-not $SkipRetention) { try { Remove-PreparedObsoleteBatches $output $pointer } catch { Write-Warning "Batch published; retention incomplete: $($_.Exception.Message)" } }
        [pscustomobject]@{BatchId=$batchId;BatchPath=$batch;Files=$validation.Files.Count;CurrentPath=$currentPath}
    } catch {
        if ($batch -and -not $committed) {
            $failureMessage=$_.Exception.Message
            try {
                $audit=Get-PreparedChildPath $output "failed/$batchId"
                New-Item -ItemType Directory -Path $audit -Force | Out-Null
                Write-PreparedJson (Join-Path $audit 'failure.json.txt') @{BatchId=$batchId;TenantKey=$TenantKey;Error=$failureMessage;Utc=[datetime]::UtcNow.ToString('O')}
                if (Test-Path -LiteralPath (Join-Path $batch 'validation.json.txt')) { Copy-Item -LiteralPath (Join-Path $batch 'validation.json.txt') -Destination (Join-Path $audit 'validation.json.txt') }
                $_.Exception.Data['PreparedFailureAuditBatchId'] = $batchId
                # Deletion is owned by finally, after checking the actual publication pointer.
            } catch { Write-Warning "Failed batch cleanup incomplete: $($_.Exception.Message)" }
        }
        throw
    } finally {
        try {
            if ($validationRoot -and (Test-Path -LiteralPath (Join-Path $validationRoot '.owner.json.txt'))) {
                $owner=Get-Content -LiteralPath (Join-Path $validationRoot '.owner.json.txt') -Raw | ConvertFrom-Json
                if ($owner.BatchId -ne $batchId -or $owner.TenantKey -ne $TenantKey) { throw 'Validation scratch ownership mismatch.' }
                Remove-PreparedOwnedPath $staging $validationRelative
            }
        } catch { Write-Warning "Validation scratch cleanup incomplete: $($_.Exception.Message)" }
        try {
            if ($batch -and -not $committed) {
                # The pointer rename may already have completed before Ctrl+C was received.
                $referenced=$false; $pointerKnown=$true
                try {
                    if (Get-SmartM365JsonReadPath $currentPath -Optional) { $referenced=(Read-SmartM365JsonDocument $currentPath).Document.BatchId -eq $batchId }
                } catch { $pointerKnown=$false; Write-Warning 'Publication pointer could not be read; the candidate batch is retained.' }
                if ($pointerKnown -and -not $referenced -and (Test-Path -LiteralPath $batch)) {
                    $audit=Get-PreparedChildPath $output "failed/$batchId"
                    New-Item -ItemType Directory -Path $audit -Force | Out-Null
                    if (-not (Test-Path -LiteralPath (Join-Path $audit 'failure.json.txt'))) {
                        Write-PreparedJson (Join-Path $audit 'failure.json.txt') @{BatchId=$batchId;TenantKey=$TenantKey;Error='Publication interrupted before pointer commit';Utc=[datetime]::UtcNow.ToString('O')}
                    }
                    Remove-PreparedOwnedPath $output "batches/$batchId"
                }
            }
            if (Test-Path -LiteralPath (Join-Path $output "current.$batchId.tmp")) { Remove-PreparedOwnedPath $output "current.$batchId.tmp" }
        } catch { Write-Warning "Publication cleanup incomplete: $($_.Exception.Message)" }
        finally { $lock.Dispose() }
    }
}

function Invoke-PreparedEvidencePipeline {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][string]$OutputRoot,
        [Parameter(Mandatory)][string]$WorkRoot, [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$AccountClassificationConfigPath,
        [int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{}, [switch]$AllowLegacyTenantless,
        [string[]]$AllowEmptyTables=@(), [switch]$ValidateOnly, [string]$MappingRoot,$AdCountryConfiguration,
        [string]$PythonCommand='python',
        [switch]$MailboxMigrationExceptionsEnabled, [string]$MailboxMigrationBindingsPath,
        [switch]$SharePointMigrationEnabled,
        [string]$SharePointMigrationWorkbookFileName='SmartWorkplaceIntelligence-Migration-to-SPO.xlsx',
        [string]$SharePointMigrationBindingsFileName='',
        [string]$SharePointMigrationTargetTenantUrl='',
        [ValidateRange(1,3650)][int]$SharePointMigrationReviewMaxAgeDays=7,
        [string]$SharePointMigrationReportingTimeZoneId='UTC',
        [Nullable[double]]$DeviceSystemMaxObservationAgeHours, [string]$DeviceSystemOffsetlessGraphTimeZoneId,
        [ValidateRange(1,8760)][int]$OfficeActivationWarningAgeHours=72,
        [AllowNull()][string[]]$HybridAllowedTargetDomains=@(),
        [ValidateRange(1,8)][int]$MaxParallelGenerators=1,
        [ValidateRange(0,520)][int]$HistoryWindowWeeks=0)
    # Read only the two non-secret preparation settings from the effective tenant profile.
    if ($null -eq $DeviceSystemMaxObservationAgeHours -and $null -ne $AdCountryConfiguration) {
        $policyValue=if ($AdCountryConfiguration -is [Collections.IDictionary]) {$AdCountryConfiguration['DeviceSystemMaxObservationAgeHours']} else {$property=$AdCountryConfiguration.PSObject.Properties['DeviceSystemMaxObservationAgeHours']; if($null -ne $property){$property.Value}}
        if ($null -ne $policyValue -and [string]$policyValue -notin @('','__USE_GLOBAL__','USE_GLOBAL')) {$DeviceSystemMaxObservationAgeHours=[double]$policyValue}
    }
    if (-not $DeviceSystemOffsetlessGraphTimeZoneId -and $null -ne $AdCountryConfiguration) {
        $zoneValue=if ($AdCountryConfiguration -is [Collections.IDictionary]) {$AdCountryConfiguration['DeviceSystemOffsetlessGraphTimeZoneId']} else {$property=$AdCountryConfiguration.PSObject.Properties['DeviceSystemOffsetlessGraphTimeZoneId']; if($null -ne $property){$property.Value}}
        if ([string]$zoneValue -notin @('','__USE_GLOBAL__','USE_GLOBAL')) {$DeviceSystemOffsetlessGraphTimeZoneId=[string]$zoneValue}
    }
    if ($null -ne $DeviceSystemMaxObservationAgeHours -and ($DeviceSystemMaxObservationAgeHours -le 0 -or [double]::IsNaN($DeviceSystemMaxObservationAgeHours) -or [double]::IsInfinity($DeviceSystemMaxObservationAgeHours))) {throw 'Device System observation age policy must be a finite positive number.'}
    if ($DeviceSystemOffsetlessGraphTimeZoneId) {$null=[TimeZoneInfo]::FindSystemTimeZoneById($DeviceSystemOffsetlessGraphTimeZoneId)}
    if($null -eq $AdCountryConfiguration){$AdCountryConfiguration=@{TenantKey=$TenantKey}}
    $endpointPolicy=ConvertTo-EndpointProtectionPolicy -Configuration $AdCountryConfiguration -TenantKey $TenantKey
    $adCountryRules=New-AdCountryRules -Configuration $AdCountryConfiguration -ExpectedTenantKey $TenantKey
    if ($AllowLegacyTenantless) { throw 'Global tenantless bypass is no longer supported. Repair reviewed historical sources instead.' }
    $root = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
    $work = [IO.Path]::GetFullPath($WorkRoot)
    if ($work.TrimEnd('\') -eq $root.TrimEnd('\') -or $work.StartsWith($root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'WorkRoot must be outside raw/synchronized DataRoot.' }
    if ($root.StartsWith($work.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'WorkRoot must not contain DataRoot.' }
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $lock = [IO.File]::Open((Join-Path $work '.preparation.lock'),'OpenOrCreate','ReadWrite','None')
    $sourceLock = $null; $run=$null; $runId=$null
    try {
        $sourceLock = [IO.File]::Open((Join-Path $root '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
        # Recover payloads from interrupted v0.1.6+ runs only, under both locks.
        # Unmarked legacy folders and another tenant's folders are never swept.
        foreach ($prior in Get-ChildItem -LiteralPath $work -Directory | Where-Object { $_.Name -match '^[a-f0-9]{32}$' }) {
            if (Get-SmartM365JsonReadPath (Join-Path $prior.FullName 'run.json') -Optional) { Clear-PreparedRunPayload $work $prior.Name $root $TenantKey }
        }
        $selectedUtc=[datetime]::UtcNow.ToString('O')
        $optionalIssues = [Collections.Generic.List[object]]::new()
        $migrationSourceArgs=@{SharePointMigrationEnabled=[bool]$SharePointMigrationEnabled;SharePointMigrationWorkbookFileName=$SharePointMigrationWorkbookFileName;SharePointMigrationBindingsFileName=$SharePointMigrationBindingsFileName}
        $plan = @(Get-PreparedSourcePlan -DataRoot $root -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides -MappingRoot $MappingRoot -MetadataOnly -OptionalSourceIssues $optionalIssues -MailboxMigrationExceptionsEnabled:$MailboxMigrationExceptionsEnabled @migrationSourceArgs)
        $migrationPrevious=Get-SharePointMigrationHistoryCapture -OutputRoot $OutputRoot -TenantKey $TenantKey
        if($migrationPrevious.Entry){$plan+= $migrationPrevious.Entry}
        if ($MailboxMigrationExceptionsEnabled -and $MailboxMigrationBindingsPath) {
            $bindingFile=Get-Item -LiteralPath $MailboxMigrationBindingsPath -ErrorAction Stop
            $plan+= [pscustomobject]@{Relative='MailboxMigrationBindings.json.txt';SourcePath=$bindingFile.FullName;Kind='Mapping';Bytes=$bindingFile.Length;ModifiedUtc=$bindingFile.LastWriteTimeUtc.ToString('O');SHA256=$null;IsOptional=$false;MaxAgeHours=$null}
        }
        $runId=[guid]::NewGuid().ToString('N')
        $run = Get-PreparedChildPath $work $runId
        New-Item -ItemType Directory -Path $run -Force | Out-Null
        Write-PreparedJson (Join-Path $run 'run.json') @{Owner='PreparedEvidencePipeline/v1';RunId=$runId;TenantKey=$TenantKey;DataRoot=$root;SelectedUtc=$selectedUtc}
        $snapshot = Join-Path $run 'source'
        $staging = Join-Path $run 'prepared'
        $captured=[Collections.Generic.List[object]]::new()
        foreach ($entry in $plan) {
            $copy = Copy-PreparedStableSource -Entry $entry -SnapshotRoot $snapshot -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides -OptionalSourceIssues $optionalIssues
            if($entry.Relative -ceq 'MigrationHistory/SharePointMigrationHistoryEvidence.csv' -and $copy.SHA256 -ne $entry.SHA256){throw 'Validated migration predecessor changed during capture.'}
            if ($null -ne $copy) { $captured.Add($copy) }
            if ($captured.Count % 25 -eq 0) { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Stable source copies: {1}/{2}.' -f (Get-Date),$captured.Count,$plan.Count) }
        }
        $plan=$captured.ToArray()
        Write-PreparedJson (Join-Path $run 'capture.json') @{SelectedUtc=$selectedUtc;Sources=$plan;OptionalSourceIssues=$optionalIssues.ToArray();Policy='Per-file verified capture; Intune operational groups additionally require matching complete source receipts. Later raw updates belong to the next run.'}
        $identity = Test-PreparedSourceTenants -Plan $plan -TenantKey $TenantKey -SnapshotRoot $snapshot
        $intuneSourcePolicyPath=Join-Path $run 'IntuneSourcePolicy.json.txt'
        Write-PreparedJson $intuneSourcePolicyPath @{MaxSourceAgeHours=$MaxSourceAgeHours;AgeOverrides=$AgeOverrides}
        if($SharePointMigrationEnabled){
            $migrationValidation=@('-B',(Join-Path $PSScriptRoot 'workplace_spo_migration.py'),'--validate-workbook',(Join-Path $snapshot $SharePointMigrationWorkbookFileName),'--tenant',$TenantKey,'--validate-sources','--data-root',$snapshot,'--as-of',$selectedUtc,'--source-policy',$intuneSourcePolicyPath)
            if($SharePointMigrationBindingsFileName){$migrationValidation+=@('--bindings',(Join-Path $snapshot $SharePointMigrationBindingsFileName))}
            $python=Get-Command $PythonCommand -CommandType Application -ErrorAction Stop | Select-Object -First 1
            & $python.Source @migrationValidation | ForEach-Object { Write-Information $_ -InformationAction Continue }
            if($LASTEXITCODE -ne 0){throw 'Migration inputs or receipt payloads failed preflight; no build or publication performed.'}
        }
        $sourceSpec=(Read-SmartM365JsonDocument (Join-Path $script:ProductRoot 'config/prepared-source-contract.json')).Document
        $nativeEnabled=$sourceSpec.PSObject.Properties['nativeEvidence'] -and $sourceSpec.nativeEvidence
        $nativeAuditPath=Join-Path $run 'NativeEvidenceAudit.json.txt'
        $nativeIdentityPath=Join-Path $run 'NativeIdentity.json.txt'
        if($nativeEnabled){
        $nativeIdentity=[ordered]@{}
        foreach($field in 'TenantKey','OrganizationKey','EnvironmentKey','TenantId') {
            $value=if($AdCountryConfiguration -is [Collections.IDictionary]){$AdCountryConfiguration[$field]}else{$AdCountryConfiguration.$field}
            if([string]::IsNullOrWhiteSpace([string]$value)){throw "Effective tenant identity is missing: $field"}
            $nativeIdentity[$field]=[string]$value
        }
        if($nativeIdentity.TenantKey -cne $TenantKey){throw 'Effective native identity tenant mismatch.'}
        Write-PreparedJson $nativeIdentityPath $nativeIdentity
        Invoke-NativeEvidenceStage -LastRoot (Join-Path $snapshot 'DATA-LAST') -TenantKey $TenantKey -AuditPath $nativeAuditPath -IdentityPath $nativeIdentityPath -PythonCommand $PythonCommand
        }
        if ($ValidateOnly) {
            $intuneChecks=@(Test-IntuneOperationalSnapshot -LastRoot (Join-Path $snapshot 'DATA-LAST') -TenantKey $TenantKey -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides)
            return [pscustomobject]@{SourceFiles=$plan.Count;Bytes=($plan | Measure-Object Bytes -Sum).Sum;Status='SourcesAndRowTenantValidationPassed';Publication=$false;Identity=$identity;IntuneOperations=$intuneChecks;DiagnosticPath=$run;OptionalSourceIssues=$optionalIssues.ToArray()}
        }
        if(-not $nativeEnabled -or -not $sourceSpec.governedArtifacts){throw 'Full Workplace preparation requires the unified native and artifact source contract.'}
        $rules = Join-Path $run 'AccountClassification.json.txt'
        $classificationCapture = Copy-PreparedClassificationCapture -SourcePath $AccountClassificationConfigPath -DestinationPath $rules
        # Retain only metadata after payload cleanup, including on a later failed stage.
        Write-PreparedJson (Join-Path $run 'ClassificationCapture.json.txt') $classificationCapture
        # Copy only country rules, never credentials or the complete private tenant profile.
        $countryRulesPath=Join-Path $run 'AdCountryRules.json.txt'
        Write-PreparedJson $countryRulesPath @{TenantKey=$TenantKey;AdDomainCountryMappings=$adCountryRules.Domains;AdComputerNameCountryPrefixes=$adCountryRules.Prefixes}
        $log = Join-Path $run 'preparation.log'
        $systemPolicyArgs=@('-NativeEvidenceAuditPath',$nativeAuditPath,'-PythonCommand',$PythonCommand)
        $systemPolicyArgs+=@('-SharePointMigrationWorkbookFileName',$SharePointMigrationWorkbookFileName,'-SharePointMigrationReviewMaxAgeDays',[string]$SharePointMigrationReviewMaxAgeDays,'-SharePointMigrationReportingTimeZoneId',$SharePointMigrationReportingTimeZoneId,'-SharePointMigrationMaxSourceAgeHours',[string]$MaxSourceAgeHours)
        if($SharePointMigrationEnabled){
            $systemPolicyArgs+='-SharePointMigrationEnabled'
            if($SharePointMigrationBindingsFileName){$systemPolicyArgs+=@('-SharePointMigrationBindingsFileName',$SharePointMigrationBindingsFileName)}
            if($SharePointMigrationTargetTenantUrl){$systemPolicyArgs+=@('-SharePointMigrationTargetTenantUrl',$SharePointMigrationTargetTenantUrl)}
        }
        if($migrationPrevious.Entry){$systemPolicyArgs+=@('-SharePointMigrationPreviousHistoryPath',(Join-Path $snapshot 'MigrationHistory/SharePointMigrationHistoryEvidence.csv'))}
        if ($MailboxMigrationExceptionsEnabled) {
            $systemPolicyArgs+='-MailboxMigrationExceptionsEnabled'
            if ($MailboxMigrationBindingsPath) { $systemPolicyArgs+=@('-MailboxMigrationBindingsPath',(Join-Path $snapshot 'MailboxMigrationBindings.json.txt')) }
        }
        $systemPolicyArgs+=@('-TenantKey',$TenantKey,'-IntuneSourcePolicyPath',$intuneSourcePolicyPath)
        $baseContract=Get-Content -LiteralPath (Get-SmartM365JsonReadPath (Join-Path $script:ProductRoot 'config/prepared-evidence-contract.json')) -Raw | ConvertFrom-Json
        $baseContract=Merge-NativeEvidenceContract -BaseContract $baseContract
        $migrationContract=Get-Content -LiteralPath (Join-Path $script:ProductRoot 'config/sharepoint-migration-contract.json.txt') -Raw | ConvertFrom-Json
        $baseContract.tables=@($baseContract.tables)+@($migrationContract.tables)
        if ($MailboxMigrationExceptionsEnabled) {
            $exceptionContract=Get-Content -LiteralPath (Join-Path $script:ProductRoot 'config/mailbox-migration-exceptions-contract.json.txt') -Raw | ConvertFrom-Json
            $baseContract.tables=@($baseContract.tables)+@($exceptionContract.tables)
        }
        $effectiveContractPath=Join-Path $run 'PreparedUnifiedContract.json.txt'
        $endpointContractArgs=@{ContractPath=$effectiveContractPath}; $endpointPolicyPath=$null
        if ($endpointPolicy.Enabled) {
            # Only normalized non-secret policy fields cross into the preparation child process.
            $endpointPolicyPath=Join-Path $run 'EndpointProtectionPolicy.json.txt'
            Write-PreparedJson $endpointPolicyPath $endpointPolicy
            $systemPolicyArgs+=@('-EndpointProtectionPolicyPath',$endpointPolicyPath)
            $extensionContract=Get-Content -LiteralPath (Join-Path $script:ProductRoot 'config/endpoint-protection-contract.json.txt') -Raw | ConvertFrom-Json
            $baseContract=Merge-EndpointProtectionContract $baseContract $extensionContract
        }
        Write-PreparedJson $effectiveContractPath $baseContract
        if ($null -ne $DeviceSystemMaxObservationAgeHours) {$systemPolicyArgs+=@('-DeviceSystemMaxObservationAgeHours',$DeviceSystemMaxObservationAgeHours)}
        if ($DeviceSystemOffsetlessGraphTimeZoneId) {$systemPolicyArgs+=@('-DeviceSystemOffsetlessGraphTimeZoneId',$DeviceSystemOffsetlessGraphTimeZoneId)}
        $systemPolicyArgs+=@('-OfficeActivationWarningAgeHours',$OfficeActivationWarningAgeHours)
        # Private TargetAddress domains cross into the child build as one comma-separated value.
        $hybridDomains=@($HybridAllowedTargetDomains | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Trim() })
        if ($hybridDomains.Count) { $systemPolicyArgs+=@('-HybridAllowedTargetDomains',($hybridDomains -join ',')) }
        $null = Test-PreparedClassificationCapture -Path $rules -Stage 'BeforeBuild' -ExpectedSHA256 $classificationCapture.Captured.SHA256 -ExpectedBytes $classificationCapture.Captured.Bytes
        $systemPolicyArgs+=@('-AccountClassificationExpectedSHA256',$classificationCapture.Captured.SHA256,'-AccountClassificationExpectedBytes',[string]$classificationCapture.Captured.Bytes)
        # Generators render warnings as plain text: relayed lines and the retained log never carry ANSI
        # sequences that would hide the WARNING prefix from the entry-point counters.
        $previousNoColor = $env:NO_COLOR
        $env:NO_COLOR = '1'
        try {
            & (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -DataRoot $snapshot -OutputRoot $staging -AccountClassificationConfigPath $rules -AdCountryRulesPath $countryRulesPath -MaxParallelGenerators $MaxParallelGenerators @systemPolicyArgs 2>&1 |
                ForEach-Object { foreach ($line in ([string]$_ -split '\r?\n')) { $message='[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date),$line; $message | Add-Content -LiteralPath $log; Write-Host $message } }
        } finally { if ($null -eq $previousNoColor) { Remove-Item -LiteralPath Env:\NO_COLOR -ErrorAction SilentlyContinue } else { $env:NO_COLOR = $previousNoColor } }
        if ($LASTEXITCODE -ne 0) { throw "Preparation failed. Retained diagnostics: $run" }
        $provenance = @{Mode='RebuiltFromSnapshot';CapturePolicy='PerFileVerified';SelectedUtc=$selectedUtc;Sources=$plan;OptionalSourceIssues=$optionalIssues.ToArray();TenantValidation=$identity;LegacyTenantlessAllowed=($identity.LegacyArchivedSummaries -gt 0);LegacyTenantlessScope='Manifest-verified archived Intune discovered-application summaries only';AccountRulesSHA256=(Get-FileHash $rules).Hash;Scripts=@(Get-ChildItem $PSScriptRoot -File | Where-Object Extension -In '.ps1','.psm1','.py','.cs' | ForEach-Object { @{File=$_.Name;SHA256=(Get-FileHash $_.FullName).Hash} })}
        $provenance.NativeEvidence=@{ContractSHA256=(Get-FileHash -LiteralPath (Join-Path $script:ProductRoot 'config/native-output-contract.json.txt')).Hash;Audit=(Get-Content -LiteralPath $nativeAuditPath -Raw | ConvertFrom-Json)}
        $provenance.SharePointMigration=@{ContractSHA256=(Get-FileHash -LiteralPath (Join-Path $script:ProductRoot 'config/sharepoint-migration-contract.json.txt')).Hash;Audit=(Get-Content -LiteralPath (Join-Path $staging 'SharePointMigrationAudit.json.txt') -Raw | ConvertFrom-Json)}
        if ($MailboxMigrationExceptionsEnabled) { $provenance.MailboxMigrationExceptions=@{ContractSHA256=(Get-FileHash -LiteralPath (Join-Path $script:ProductRoot 'config/mailbox-migration-exceptions-contract.json.txt')).Hash;Audit=(Get-Content -LiteralPath (Join-Path $staging 'MailboxMigrationExceptionsAudit.json.txt') -Raw | ConvertFrom-Json)} }
        $provenance.WindowsMigration = (Get-Content -LiteralPath (Join-Path $staging 'WindowsMigrationAudit.json.txt') -Raw | ConvertFrom-Json)
        $provenance.IntuneOperations=@{ContractSHA256=(Get-FileHash -LiteralPath (Join-Path $script:ProductRoot 'config/intune-operational-contract.json.txt')).Hash;SourcePolicySHA256=(Get-FileHash -LiteralPath $intuneSourcePolicyPath).Hash;Audit=(Get-Content -LiteralPath (Join-Path $staging 'IntuneOperationalAudit.json.txt') -Raw | ConvertFrom-Json)}
        $provenance.AdCountryRulesSHA256=(Get-FileHash -LiteralPath $countryRulesPath).Hash
        $provenance.OfficeActivationFreshness=@{WarningAgeHours=$OfficeActivationWarningAgeHours;Basis='DATA-LAST write time = last successful acquisition';Effect='Prepare log warning and Data Trust Aging only; never blocks publication.'}
        $provenance.DeviceSystemPolicy=@{MaxObservationAgeHours=$DeviceSystemMaxObservationAgeHours;OffsetlessGraphTimeZoneId=$DeviceSystemOffsetlessGraphTimeZoneId;HardwareFreshness='Platform timestamp is context only; not hardware collection time.'}
        $provenance.WorkforceActivity=@{DefinitionVersion='AD-M365-Console-v1';CurrentSources=@('Active Directory','Microsoft 365','Qualified local console');ConsoleIdentity='Exact unique tenant-scoped AD SID and unique AD-to-Entra link';ConsoleScope='Last observed local interactive user per device; partial coverage; RDP excluded';HistoricalSources=@('Active Directory','Microsoft 365');CurrentConsoleInjectedIntoHistory=$false}
        if ($endpointPolicy.Enabled) {$provenance.EndpointProtectionPolicy=@{PolicyVersion=$endpointPolicy.PolicyVersion;SHA256=(Get-FileHash -LiteralPath $endpointPolicyPath).Hash;Qualification='Antivirus and EDR evaluated separately; unknown is not healthy.'}}
        $migrationPublicationArgs=@{ExpectedPreviousBatchId=$migrationPrevious.BatchId}
        Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $OutputRoot -TenantKey $TenantKey -Provenance $provenance -AllowEmptyTables $AllowEmptyTables -ArtifactSnapshotRoot $snapshot -HistoryWindowWeeks $HistoryWindowWeeks @endpointContractArgs @migrationPublicationArgs
    } catch {
        if ($run) { try { Write-PreparedJson (Join-Path $run 'failure.json') @{Error=$_.Exception.Message;Utc=[datetime]::UtcNow.ToString('O')} } catch { Write-Warning 'Could not write run failure diagnostic.' } }
        throw
    } finally {
        if ($run) { try { Clear-PreparedRunPayload $work $runId $root $TenantKey } catch { Write-Warning "Temporary cleanup incomplete: $($_.Exception.Message)" } }
        if ($sourceLock) { $sourceLock.Dispose() }; $lock.Dispose()
    }
}
Export-ModuleMember -Function Get-PreparedSourcePlan, Invoke-PreparedEvidencePipeline, Publish-PreparedEvidenceBatch, Receive-PreparedMappingWorkbooks, Remove-PreparedMappingWorkbooks

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBjWqZxL99EatpA
# s8yUikDazCMJw4vaZ3RDnWEJvkcBlaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ2sxc5lhcOlgAbaMw0PsxgCerCQPRHbCChIbqhbz9VUMA0GCSqG
# SIb3DQEBAQUABIIBgJ05J8jMw129E9fEs5VGM60t+Fa93zPns7PJ5IMP7VLydshI
# PzxDQxEExmZMoQ5cMoNYmhJFeSR4HnzRtyDERRba4maRIxxK3+FuWcDIzfTTzJpP
# 0YPVMx+rl00aSlEqj8mqERN2x9ZVbLbBxMtlAQvfRGep8r7MGB/eGf2OrQ9Cws3Z
# QyKg6mWLQ++uxoym8Vop8lNFMpdOtg3ipwXiO6uQfJDTT5/TPuDSVX79+L6GEJCI
# TAE41o6j7FIyPipf12WPjUJyyeTE2CQYSFyo3FQAeVvaH6graFJ9hHXHog2jMbOc
# xSPNESs9UUZtEBDLsAgIGa3bdiQNYRzLuahOGMXE1AYBHuqwHL0fl6QL36uHfC37
# RdtH6m1XcRYewtB5C2nb4VSST4t6SXEpgjGy8wXdyKA0D4vDddZhYk8EtT20T+S0
# ehpaV8lJDTHfcFFfiJBtBttsgCI32dLEmwxv8E7W1hDRknDCWjnwyqiR2TXbEKtZ
# 6n+0S7G0X8oqz00BhaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMjAy
# NDdaMC8GCSqGSIb3DQEJBDEiBCC2uXFh9QC0qkzIudvQipp3s2sTJVjktlbJdK3R
# Q8py8zANBgkqhkiG9w0BAQEFAASCAgAR4kToSpD4AnvTE0LqaDXBjw1k2ukPQP0J
# TK2FPuZ64LFEzFgRxQKzGsWrw5o14QCWJ/zg2DtBT4MdvZrqaFTzKZ9N00O1UzxO
# 1k02KVNNV/WOJ+334Kf+RPqiL87PEdBYDLoz3KLrEi8dkIbPEccWbJfCC67Yowjy
# F5NNks5pjVO9gy6XufFobZeuC0jmYNeYfMlPM427OAfRaj7EXhnle0hCAwbcBIp+
# nnqgtzRgY6PceOfHi5eqw4ba4Or+LRPD3YYStA81JZLGqhQnb+TVPOkTshX1z75x
# tcmoRqnIdwT/VxZE3MSNPVPsA1BOXM8hPuf6xGxk0ujQ1dknf1h6uW021Yjp2+u3
# NULoYs1oHTFrOGpTQ5zl7vjqdN9ecnnE0wTi+ok2brjOLW+ztmR9JoBgeaAY1qo4
# 4nLPqJ2LwMJ0a0L2D9GDBMfOLU3rbCB9VK6/IH4HtNe5hSxZQ+JxgBaiCEpWMwaM
# 9WYAGHTA6tq3VKMR6pLAHUUYEQj71Co+Y6AAgXhIzSpunXu3VmYZI9RdWmem9RKK
# +NtGj89z6kLWJNpvPH1+LLrES1Fw2gV5EWlMjJm1QL6eDf94VcyIc2w51o4wrwi0
# 1djNnvm0o0Z8aRnO7oi34sYVCvgwg7qCUONsVAxZfI4wx5MwNnnMwNiaOtXNDo+w
# SuDpaf1pXQ==
# SIG # End signature block
