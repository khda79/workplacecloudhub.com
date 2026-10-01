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
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth 20))
    if($Path.EndsWith('.json',[StringComparison]::OrdinalIgnoreCase)){
        $Path=Resolve-SmartM365OwnedJsonPath -Path $Path -Owner 'WorkplaceEvidence-Prepare' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Prepared audit must be an object.'}}
    }
    if($Path -match '\.json(?:\.txt)?$'){Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Prepared metadata must be an object.'}} | Out-Null}
    else{[IO.File]::WriteAllBytes($Path,$bytes)}
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
    foreach ($name in 'source','prepared','AccountClassification.psd1','AccountClassification.json.txt') { Remove-PreparedOwnedPath $WorkRoot "$RunId/$name" }
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
    $spec = (Read-SmartM365JsonDocument $SourceContractPath).Document
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
    Initialize-PreparedMetadataNames -OutputRoot $output -TenantKey $TenantKey
    # Held throughout validation and commit; a failed process releases the OS lock.
    $lock = [IO.File]::Open((Join-Path $output '.publication.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    $batch=$null; $committed=$false
    try {
        $currentPath = Join-Path $output 'current.json.txt'
        if (Test-Path -LiteralPath (Join-Path $output 'current.json')) { throw 'Legacy prepared metadata is preserved while JSON transport policy is Readers. Activate the approved JsonText deployment before publication.' }
        $previousPath = Get-PreparedMetadataPath $output 'current' -Optional
        $previous = if ($previousPath) { Get-Content -LiteralPath $previousPath -Raw | ConvertFrom-Json } else { $null }
        if ($previous -and $previous.TenantKey -ne $TenantKey) { throw 'Output root belongs to another tenant.' }
        $contractReceipt = Read-SmartM365JsonDocument $ContractPath
        $ContractPath = $contractReceipt.Path
        $contract = $contractReceipt.Document
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
            if (Get-SmartM365JsonReadPath (Join-Path $prior.FullName 'run.json') -Optional) { Clear-PreparedRunPayload $work $prior.Name $root $TenantKey }
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
        $rules = Join-Path $run 'AccountClassification.json.txt'
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
