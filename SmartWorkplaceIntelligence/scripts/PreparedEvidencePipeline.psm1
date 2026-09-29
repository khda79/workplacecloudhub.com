# Offline, single-tenant preparation and versioned local publication. No tenant APIs.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ProductRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force

function Get-PreparedChildPath([string]$Root, [string]$Relative) {
    $base = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $path = [IO.Path]::GetFullPath((Join-Path $base $Relative))
    if (-not $path.StartsWith($base + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Path must be strictly below its declared root.'
    }
    $path
}

function Write-PreparedJson([string]$Path, $Value) {
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
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
    $marker = Get-Content -LiteralPath (Join-Path $run 'run.json') -Raw | ConvertFrom-Json
    if ($marker.Owner -ne 'PreparedEvidencePipeline/v1' -or $marker.RunId -ne $RunId -or $marker.DataRoot -ne $DataRoot -or $marker.TenantKey -ne $TenantKey) { throw 'Run ownership mismatch; no cleanup performed.' }
    foreach ($name in 'source','prepared','AccountClassification.psd1') { Remove-PreparedOwnedPath $WorkRoot "$RunId/$name" }
}

function Copy-PreparedStableSource {
    param($Entry, [string]$SnapshotRoot, [int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{},
        [ValidateRange(1,3)][int]$Attempts=3, [ValidateRange(0,2000)][int]$RetryDelayMs=1000)
    $to = Get-PreparedChildPath $SnapshotRoot $Entry.Relative
    New-Item -ItemType Directory -Path (Split-Path $to -Parent) -Force | Out-Null
    for ($attempt=1; $attempt -le $Attempts; $attempt++) {
        $inputStream=$null; $outputStream=$null; $hash=$null; $verifyStream=$null; $sha=$null
        try {
            $before = Get-Item -LiteralPath $Entry.SourcePath -ErrorAction Stop
            if ($before.Length -eq 0) { throw 'Empty input file.' }
            if ($Entry.Kind -eq 'Current') {
                $limit = if ($AgeOverrides.ContainsKey($before.Name)) { [int]$AgeOverrides[$before.Name] } else { $MaxSourceAgeHours }
                if ($limit -lt 1 -or ([datetime]::UtcNow-$before.LastWriteTimeUtc).TotalHours -gt $limit -or $before.LastWriteTimeUtc -gt [datetime]::UtcNow.AddMinutes(5)) { throw 'Input publication-age limit or future timestamp violated.' }
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
            if ($attempt -eq $Attempts) { throw "Capture failed after $attempt attempts: $($Entry.Relative): $($_.Exception.Message)" }
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Capture retry {1}/{2}: {3}: {4}' -f (Get-Date),($attempt+1),$Attempts,$Entry.Relative,$_.Exception.Message)
        } finally {
            foreach ($resource in $inputStream,$outputStream,$hash,$verifyStream,$sha) { if ($null -ne $resource) { $resource.Dispose() } }
        }
        if ($RetryDelayMs) { Start-Sleep -Milliseconds $RetryDelayMs }
    }
}

function Remove-PreparedMappingWorkbooks([string]$WorkRoot, [string]$MappingRoot) {
    $relative = [IO.Path]::GetRelativePath([IO.Path]::GetFullPath($WorkRoot),[IO.Path]::GetFullPath($MappingRoot)).Replace('\','/')
    if ($relative -notmatch '^mappings/[a-f0-9]{32}$') { throw 'Mapping cleanup requires the exact run-owned download folder.' }
    Remove-PreparedOwnedPath $WorkRoot $relative
}

function Remove-PreparedObsoleteBatches([string]$OutputRoot, $Pointer) {
    # Follow only the published chain, never sweep arbitrary or failed folders.
    $keep = @($Pointer.BatchId,$Pointer.PreviousBatchId)
    if (-not $Pointer.PreviousBatchId) { return }
    foreach ($id in $keep) { if ($id -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$') { throw 'Invalid protected batch ID.' } }
    $previousPath = Get-PreparedMetadataPath (Get-PreparedChildPath $OutputRoot "batches/$($Pointer.PreviousBatchId)") 'current' -Optional
    if (-not $previousPath) { return }
    $previous = Get-Content -LiteralPath $previousPath -Raw | ConvertFrom-Json
    if ($previous.BatchId -ne $Pointer.PreviousBatchId -or $previous.TenantKey -ne $Pointer.TenantKey) { throw 'Previous batch receipt mismatch.' }
    $candidate = $previous.PreviousBatchId
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    while ($candidate) {
        if ($candidate -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$' -or $candidate -in $keep -or -not $seen.Add($candidate)) { throw 'Invalid or cyclic batch retention chain.' }
        $relative = "batches/$candidate"
        $folder = Get-PreparedChildPath $OutputRoot $relative
        if (-not (Test-Path -LiteralPath $folder)) { break } # Already retired on an earlier run.
        $receiptPath = Get-PreparedMetadataPath $folder 'current'
        $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
        $manifestPath = Get-PreparedMetadataPath $folder 'batch'
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($receipt.BatchId -ne $candidate -or $manifest.BatchId -ne $candidate -or $receipt.TenantKey -ne $Pointer.TenantKey -or $manifest.TenantKey -ne $Pointer.TenantKey -or (Get-FileHash -LiteralPath $manifestPath).Hash -ne $receipt.ManifestSHA256) { throw 'Old batch ownership/hash mismatch; retained.' }
        $candidate = $receipt.PreviousBatchId
        # Preserve small audit receipts, but not duplicate historical CSV payloads.
        $audit = Get-PreparedChildPath $OutputRoot "retired/$($receipt.BatchId)"
        New-Item -ItemType Directory -Path $audit -Force | Out-Null
        foreach ($name in 'current','batch','validation') { Copy-Item -LiteralPath (Get-PreparedMetadataPath $folder $name) -Destination (Join-Path $audit ($name+'.json.txt')) -ErrorAction Stop }
        Remove-PreparedOwnedPath $OutputRoot $relative
        Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Retired prepared batch {1}; audit retained, raw history untouched.' -f (Get-Date),$receipt.BatchId)
    }
}

function Receive-PreparedMappingWorkbooks {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkRoot, [Parameter(Mandatory)][scriptblock]$DownloadFile)
    $folder = Join-Path ([IO.Path]::GetFullPath($WorkRoot)) ('mappings/'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $required = [ordered]@{
        'SmartWorkplaceIntelligence-PersonaClassification.xlsx' = @('Personas','Rules','Exclusions')
        'SmartWorkplaceIntelligence-SiteClassification.xlsx' = @('Settings','SiteTypes','Sites')
    }
    try { foreach ($name in $required.Keys) {
        $path = Join-Path $folder $name
        $receipt = & $DownloadFile $path $name
        if (-not $receipt -or -not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Item -LiteralPath $path).Length -eq 0) {
            throw "Required SharePoint mapping download failed: $name. No cached fallback is allowed."
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
    } } catch {
        try { Remove-PreparedMappingWorkbooks $WorkRoot $folder } catch { Write-Warning "Mapping cleanup failed: $($_.Exception.Message)" }
        throw
    }
    # Return only after both files passed; a partial download never reaches source planning.
    $folder
}

function Get-PreparedSourcePlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot,
        [string]$SourceContractPath = (Join-Path $script:ProductRoot 'config/prepared-source-contract.json'),
        [ValidateRange(1,8760)][int]$MaxSourceAgeHours = 168,
        [hashtable]$AgeOverrides = @{}, [string]$MappingRoot, [switch]$MetadataOnly)
    $root = (Resolve-Path -LiteralPath $DataRoot).ProviderPath
    $spec = Get-Content -LiteralPath $SourceContractPath -Raw | ConvertFrom-Json
    $items = [Collections.Generic.List[object]]::new()
    foreach ($name in $spec.currentFiles) { $items.Add(@{Relative="DATA-LAST/$name";Kind='Current'}) }
    foreach ($name in $spec.mappingFiles) { $items.Add(@{Relative=$name;Kind='Mapping'}) }
    foreach ($name in $spec.dailyFiles) { $items.Add(@{Relative=$name;Kind='History'}) }
    foreach ($entry in $spec.history) {
        $folder = Get-PreparedChildPath $root $entry.root
        $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse -Filter $entry.file -ErrorAction Stop |
            Where-Object { $_.FullName -match '[\\/]WeeklyHistory[\\/]\d{4}-W\d{2}[\\/]' })
        if (-not $files.Count) { throw "Required observed history missing: $($entry.root)/$($entry.file)" }
        foreach ($file in $files) { $items.Add(@{Relative=[IO.Path]::GetRelativePath($root,$file.FullName);Kind='History'}) }
    }
    foreach ($item in $items | Sort-Object Relative -Unique) {
        $sourceRoot = if ($item.Kind -eq 'Mapping' -and $MappingRoot) { $MappingRoot } else { $root }
        $path = Get-PreparedChildPath $sourceRoot $item.Relative
        $file = Get-Item -LiteralPath $path -ErrorAction Stop
        if ($file.Length -eq 0) { throw "Empty input file: $($item.Relative)" }
        # Transport freshness only; business report timestamps remain in Data Trust.
        if ($item.Kind -eq 'Current') {
            $limit = if ($AgeOverrides.ContainsKey($file.Name)) { [int]$AgeOverrides[$file.Name] } else { $MaxSourceAgeHours }
            if ($limit -lt 1 -or ([datetime]::UtcNow - $file.LastWriteTimeUtc).TotalHours -gt $limit) {
                throw "Input exceeds configured publication-age limit: $($item.Relative)" }
            if ($file.LastWriteTimeUtc -gt [datetime]::UtcNow.AddMinutes(5)) { throw "Future input timestamp: $($item.Relative)" }
        }
        [pscustomobject]@{Relative=$item.Relative;SourcePath=$file.FullName;Kind=$item.Kind;Bytes=$file.Length;ModifiedUtc=$file.LastWriteTimeUtc.ToString('O');SHA256=$(if($MetadataOnly){$null}else{(Get-FileHash -LiteralPath $path).Hash})}
    }
}

function Test-PreparedTenant {
    param([string]$Path, [string]$TenantKey)
    # Match the source formats consumed by the evidence generators; never retry
    # malformed records with another separator or weaken the tenant checks.
    $fileName = [IO.Path]::GetFileName($Path)
    $delimiter = if ($fileName -like '*DailyStats.csv' -or $fileName -eq 'Exchange_OnPrem_Servers_Inventory.csv') { ';' } else { ',' }
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.SetDelimiters([string]$delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $header = $parser.ReadFields()
        if (-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count) { throw 'Empty or duplicate CSV header.' }
        $index = [array]::IndexOf($header,'TenantKey')
        if ($index -lt 0) { throw 'TenantKey column missing.' }
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
    $rows = 0L; $files = 0
    foreach ($entry in $Plan | Where-Object { $_.Relative -like '*.csv' }) {
        $relative = $entry.Relative.Replace('\','/')
        $path = if ($SnapshotRoot) { Get-PreparedChildPath $SnapshotRoot $relative } else { $entry.SourcePath }
        try {
            $check = Test-PreparedTenant -Path $path -TenantKey $TenantKey
            $rows += $check.Rows; $files++
        } catch { $errors.Add("${relative}: $($_.Exception.Message)") }
        if ($files -gt 0 -and $files % 25 -eq 0) { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Source tenant validation: {1} CSVs checked.' -f (Get-Date),$files) }
    }
    if ($errors.Count) { throw ("Source tenant validation failed ($($errors.Count) files):`n" + ($errors -join "`n")) }
    [pscustomobject]@{CsvFiles=$files;Rows=$rows}
}

function Publish-PreparedEvidenceBatch {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$StagingRoot,
        [Parameter(Mandatory)][string]$OutputRoot,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_.-]+$')][string]$TenantKey,
        [Parameter(Mandatory)]$Provenance,
        [string]$ContractPath = (Join-Path $script:ProductRoot 'config/prepared-evidence-contract.json'),
        [string[]]$AllowEmptyTables = @())
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
    $output = [IO.Path]::GetFullPath($OutputRoot)
    $staging = (Resolve-Path -LiteralPath $StagingRoot).ProviderPath
    if ($output.TrimEnd('\') -eq $staging.TrimEnd('\')) { throw 'Staging and publication roots must differ.' }
    if ((Split-Path $output -Leaf) -ne 'DATA-POWERBI') { throw 'Publication root must be a dedicated DATA-POWERBI directory.' }
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    # Held throughout validation and commit; a failed process releases the OS lock.
    $lock = [IO.File]::Open((Join-Path $output '.publication.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    $batch=$null; $committed=$false
    try {
        $currentPath = Join-Path $output 'current.json.txt'
        if (Test-Path -LiteralPath (Join-Path $output 'current.json')) { throw 'Convert existing metadata with Convert-SmartM365-WorkplaceEvidence-Metadata.cmd before publishing another batch.' }
        $previousPath = Get-PreparedMetadataPath $output 'current' -Optional
        $previous = if ($previousPath) { Get-Content -LiteralPath $previousPath -Raw | ConvertFrom-Json } else { $null }
        if ($previous -and $previous.TenantKey -ne $TenantKey) { throw 'Output root belongs to another tenant.' }
        $contract = Get-Content -LiteralPath $ContractPath -Raw | ConvertFrom-Json
        foreach ($name in $AllowEmptyTables) { if ($name -notin $contract.tables.table) { throw "Unknown empty-table exception: $name" } }
        $batchId = [datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
        $batch = Get-PreparedChildPath $output "batches/$batchId"
        New-Item -ItemType Directory -Path $batch -Force | Out-Null
        foreach ($table in $contract.tables) {
            $source = Get-PreparedChildPath $staging $table.file
            $destination = Get-PreparedChildPath $batch $table.file
            Copy-Item -LiteralPath $source -Destination $destination -ErrorAction Stop
            if ((Get-FileHash -LiteralPath $source).Hash -ne (Get-FileHash -LiteralPath $destination).Hash) { throw 'Staging changed during publication.' }
            # Retain the original header for a valid header-only empty export.
            $reader = [IO.StreamReader]::new($destination)
            $temp = $destination + '.tenant.tmp'
            try {
                $header = $reader.ReadLine()
                if ($header -match '^"?TenantKey"?,') { throw 'Staging must not contain a published TenantKey prefix.' }
            } finally { $reader.Dispose() }
            # Stream records; prepared files are modest compared with source inventories.
            Import-Csv -LiteralPath $destination | Select-Object @{n='TenantKey';e={$TenantKey}},* |
                Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding utf8
            # Header-only files need explicit schema preservation.
            if ((Get-Item -LiteralPath $temp).Length -eq 0) {
                [IO.File]::WriteAllText($temp, '"TenantKey",' + $header + [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
            }
            [IO.File]::Move($temp,$destination,$true)
        }
        $validationPath = Join-Path $batch 'validation.json.txt'
        & (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $batch -ContractPath $ContractPath -ReportPath $validationPath | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Batch validation failed; current.json was not changed.' }
        $validation = Get-Content -LiteralPath $validationPath -Raw | ConvertFrom-Json
        foreach ($entry in $validation.Files) {
            $table = $contract.tables | Where-Object file -EQ $entry.File
            if ($entry.Rows -eq 0 -and $table.table -notin $AllowEmptyTables) { throw "Unexpected empty output: $($entry.File)" }
        }
        if ($previous) {
            # Loss of an observed historical key is never silently accepted.
            foreach ($table in $contract.tables | Where-Object { $_.table -match 'History|Trend' }) {
                $oldPath = Get-PreparedChildPath $output "batches/$($previous.BatchId)/$($table.file)"
                $newPath = Get-PreparedChildPath $batch $table.file
                $newKeys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $keyColumns = @($table.columns.name | Where-Object { $_ -match 'Date|Week|Metric Name|Service|Country' })
                if (-not $keyColumns.Count) { throw "No historical comparison key configured: $($table.table)" }
                $keyOf = { param($row) ($keyColumns | ForEach-Object { $v=[string]$row.$_; $v.Length.ToString()+':'+$v }) -join '|' }
                Import-Csv -LiteralPath $newPath | ForEach-Object { $null=$newKeys.Add((& $keyOf $_)) }
                Import-Csv -LiteralPath $oldPath | ForEach-Object { if (-not $newKeys.Contains((& $keyOf $_))) { throw "Historical coverage would regress: $($table.table)" } }
            }
        }
        $manifest = [ordered]@{SchemaVersion=1;BatchId=$batchId;TenantKey=$TenantKey;CreatedUtc=[datetime]::UtcNow.ToString('O');Provenance=$Provenance;ContractSHA256=(Get-FileHash $ContractPath).Hash;Files=@($validation.Files | ForEach-Object { @{File=$_.File;Rows=$_.Rows;Bytes=(Get-Item (Join-Path $batch $_.File)).Length;SHA256=$_.SHA256} })}
        Write-PreparedJson (Join-Path $batch 'batch.json.txt') $manifest
        $pointer = @{SchemaVersion=1;BatchId=$batchId;TenantKey=$TenantKey;ManifestSHA256=(Get-FileHash (Join-Path $batch 'batch.json.txt')).Hash;PreviousBatchId=if($previous){$previous.BatchId}else{$null}}
        $pointerTemp = Join-Path $output ('current.'+$batchId+'.tmp')
        Write-PreparedJson $pointerTemp $pointer
        # Immutable copy used by a delayed/cloud transfer; never upload a newer run's pointer.
        Write-PreparedJson (Join-Path $batch 'current.json.txt') $pointer
        [IO.File]::Move($pointerTemp,$currentPath,$true)
        $committed=$true
        try { Remove-PreparedObsoleteBatches $output $pointer } catch { Write-Warning "Batch published; retention incomplete: $($_.Exception.Message)" }
        [pscustomobject]@{BatchId=$batchId;BatchPath=$batch;Files=$validation.Files.Count;CurrentPath=$currentPath}
    } catch {
        if ($batch -and -not $committed) {
            $failureMessage=$_.Exception.Message
            try {
                $audit=Get-PreparedChildPath $output "failed/$batchId"
                New-Item -ItemType Directory -Path $audit -Force | Out-Null
                Write-PreparedJson (Join-Path $audit 'failure.json.txt') @{BatchId=$batchId;TenantKey=$TenantKey;Error=$failureMessage;Utc=[datetime]::UtcNow.ToString('O')}
                if (Test-Path -LiteralPath (Join-Path $batch 'validation.json.txt')) { Copy-Item -LiteralPath (Join-Path $batch 'validation.json.txt') -Destination (Join-Path $audit 'validation.json.txt') }
                Remove-PreparedOwnedPath $output "batches/$batchId"
                Remove-PreparedOwnedPath $output "current.$batchId.tmp"
            } catch { Write-Warning "Failed batch cleanup incomplete: $($_.Exception.Message)" }
        }
        throw
    } finally { $lock.Dispose() }
}

function Invoke-PreparedEvidencePipeline {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot, [Parameter(Mandatory)][string]$OutputRoot,
        [Parameter(Mandatory)][string]$WorkRoot, [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$AccountClassificationConfigPath,
        [int]$MaxSourceAgeHours=168, [hashtable]$AgeOverrides=@{}, [switch]$AllowLegacyTenantless,
        [string[]]$AllowEmptyTables=@(), [switch]$ValidateOnly, [string]$MappingRoot)
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
            if (Test-Path -LiteralPath (Join-Path $prior.FullName 'run.json')) { Clear-PreparedRunPayload $work $prior.Name $root $TenantKey }
        }
        $selectedUtc=[datetime]::UtcNow.ToString('O')
        $plan = @(Get-PreparedSourcePlan -DataRoot $root -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides -MappingRoot $MappingRoot -MetadataOnly)
        $runId=[guid]::NewGuid().ToString('N')
        $run = Get-PreparedChildPath $work $runId
        New-Item -ItemType Directory -Path $run -Force | Out-Null
        Write-PreparedJson (Join-Path $run 'run.json') @{Owner='PreparedEvidencePipeline/v1';RunId=$runId;TenantKey=$TenantKey;DataRoot=$root;SelectedUtc=$selectedUtc}
        $snapshot = Join-Path $run 'source'
        $staging = Join-Path $run 'prepared'
        $captured=[Collections.Generic.List[object]]::new()
        foreach ($entry in $plan) {
            $captured.Add((Copy-PreparedStableSource -Entry $entry -SnapshotRoot $snapshot -MaxSourceAgeHours $MaxSourceAgeHours -AgeOverrides $AgeOverrides))
            if ($captured.Count % 25 -eq 0) { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Stable source copies: {1}/{2}.' -f (Get-Date),$captured.Count,$plan.Count) }
        }
        $plan=$captured.ToArray()
        Write-PreparedJson (Join-Path $run 'capture.json') @{SelectedUtc=$selectedUtc;Sources=$plan;Policy='Per-file verified capture; later raw updates belong to the next run, not an atomic collector-wide snapshot.'}
        $identity = Test-PreparedSourceTenants -Plan $plan -TenantKey $TenantKey -SnapshotRoot $snapshot
        if ($ValidateOnly) {
            return [pscustomobject]@{SourceFiles=$plan.Count;Bytes=($plan | Measure-Object Bytes -Sum).Sum;Status='SourcesAndRowTenantValidationPassed';Publication=$false;Identity=$identity;DiagnosticPath=$run}
        }
        $rules = Join-Path $run 'AccountClassification.psd1'
        Copy-Item -LiteralPath $AccountClassificationConfigPath -Destination $rules
        $log = Join-Path $run 'preparation.log'
        & (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -DataRoot $snapshot -OutputRoot $staging -AccountClassificationConfigPath $rules 2>&1 |
            ForEach-Object { foreach ($line in ([string]$_ -split '\r?\n')) { $message='[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date),$line; $message | Add-Content -LiteralPath $log; Write-Host $message } }
        if ($LASTEXITCODE -ne 0) { throw "Preparation failed. Retained diagnostics: $run" }
        $provenance = @{Mode='RebuiltFromSnapshot';CapturePolicy='PerFileVerified';SelectedUtc=$selectedUtc;Sources=$plan;TenantValidation=$identity;LegacyTenantlessAllowed=$false;AccountRulesSHA256=(Get-FileHash $rules).Hash;Scripts=@(Get-ChildItem $PSScriptRoot -File | Where-Object Extension -In '.ps1','.psm1' | ForEach-Object { @{File=$_.Name;SHA256=(Get-FileHash $_.FullName).Hash} })}
        Publish-PreparedEvidenceBatch -StagingRoot $staging -OutputRoot $OutputRoot -TenantKey $TenantKey -Provenance $provenance -AllowEmptyTables $AllowEmptyTables
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6Q770ip1ZXeOg
# 2o1/CmJD0V9X2ta56BnVqwlO7OlczqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICtSmAr33ip6bKETTz4sTDNKExKd6ixbe6uDnKdEc0deMA0GCSqG
# SIb3DQEBAQUABIIBgCorYJpzYYaZYoj/Xo0eWmGoZ5kBvjiIZlNsVnwcuhmyAwtM
# JRbD6UGZyu4mcQ6FUVkzDtGX2Q12MossFEYJqCMovjC3wB0dr6rqht+N8if0+1QV
# wEBXrmMkm0TUUEyXG09IF3h/31X/1Mbu0/LPRB5gKnNVpPVj4bpAWXzXwMP86BF+
# Re6dJfcxDzVXMz1VSDxYk0eSXA96B/FiR7LdbJS1AuL5kz7agp/q+SXOKBWkXLDq
# 78/BUt7W6g6k9Nm/KrifRa3I2PlH7Y3awJb2c3ihIjeAJqXuEioDZZrPkaOG6oZo
# St5zXuLkhL26a9b2aoNDMe3Bc/7DyG2LuezWyUvZxFuMZRExy6nWpXvK6C5X1K5a
# v4rCIx2VcwOP6h0Awg4OQFjDzBF4axKIcteXWkfHAmoXv0kJJq2ouZMEWKFGJORb
# okxm3TyjP1xhbCYW0In2XPJrjH9ke4A4lY7s1jEbP4cMUsbsZPAeC6f0rwraYWww
# asMQJXGvlpAg8voMyKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxMzI5
# MzJaMC8GCSqGSIb3DQEJBDEiBCChSu0s/JuW0yjUT3P2tklpOKkJ6nTUtBlIWzvv
# xm9jrzANBgkqhkiG9w0BAQEFAASCAgA+MviLfPWD4XytHcnm8vQIPXXOcXNvlXeQ
# OGqM7cH9utXnIFYNvYlzVeJyIWwzd9zdqUGqZidfW7mcRKK8HTyooSCe+Y2c1hfa
# bIkUX16I07nfTUEZzk+BYWkdPudA2uAn4JaGGgI4VusM1ZI3ig9NqdKK4VwTWhrh
# 9FSESzMkpX+PFKZ5uFa/VAymUe6E9uuEA8vXz49bdvhuM5lW8mM4Gh7CEcDh6s7e
# p4vm3C78e1W9vKpjNzTneMAhZb0Qw662qxOsY8S++TV4dsjYeH6L+5pqftPq3SH3
# iMlb4EmfKAR5jg1D8AE/qrfYhcgA6w+vZPZs3ts77GM7t5Z3UMqG/YtdWuqvGGXy
# zHCwbKd/LZa5X/07dg3lBDvhoa246Pdxq7Oa85GJFzIBKXkRugpeyzcNzS8a5mQ/
# yph7ezlK0m3/RjhdWAbxiFKEwAoF/hParjiPKCQ7kcJnkx+g02GJ3nMPlT08Cj91
# ltSp/9tfbkE4B95Hmg39qwpN5bLw7inqgPaMBC+y9VqinQVLNZi6ZP8XgSvPCfFQ
# uiMRYW0JLeruzEyBpaKBb/vN6r6nuIJMpAlzt11NE/u1lu4os4ZxomhTftLMWJem
# TL5/5H4TfFOrzrEKxlEy1tJ7LZHMYS/P6TAxnTeA+X4RmLS28dVfem22pHxemuwk
# MGIQm/3rnQ==
# SIG # End signature block
