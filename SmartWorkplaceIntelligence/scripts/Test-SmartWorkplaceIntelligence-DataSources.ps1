[CmdletBinding()]
param(
    [Parameter()]
    [string]$DataRoot = $env:SMARTM365_DATA_ROOT,

    [Parameter()]
    [string]$ManifestPath,

    [Parameter()]
    [int]$FreshnessWarningHours = 72
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $PSScriptRoot '..\config\source-manifest.json'
}

if ([string]::IsNullOrWhiteSpace($DataRoot)) {
    throw 'Specify -DataRoot or set SMARTM365_DATA_ROOT. The private path is intentionally not persisted in project files.'
}

function Get-CsvHeader {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][char]$Delimiter = ','
    )

    # Import-Csv correctly handles quoted and very wide headers used by the
    # SmartM365 exports. Selecting the first record keeps this schema check
    # streaming and avoids loading large private inventories into memory.
    $firstRow = Import-Csv -LiteralPath $Path -Delimiter $Delimiter | Select-Object -First 1
    if ($null -ne $firstRow) {
        return @($firstRow.PSObject.Properties.Name)
    }

    # Header-only exports are valid for collectors that found no protected
    # objects or no findings. Parse the header without inventing a data row.
    $headerLine = Get-Content -LiteralPath $Path -TotalCount 1
    if ([string]::IsNullOrWhiteSpace($headerLine)) {
        return @()
    }
    $syntheticRow = ($headerLine -split [regex]::Escape([string]$Delimiter) | ForEach-Object { '""' }) -join $Delimiter
    $headerOnly = ConvertFrom-Csv -InputObject @($headerLine, $syntheticRow) -Delimiter $Delimiter | Select-Object -First 1
    return @($headerOnly.PSObject.Properties.Name)
}

$resolvedManifest = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifest = Get-Content -LiteralPath $resolvedManifest -Raw | ConvertFrom-Json
$currentRoot = Join-Path $DataRoot $manifest.currentRoot
$historyRoot = Join-Path $DataRoot $manifest.historyRoot
$now = Get-Date

$results = foreach ($source in $manifest.sources) {
    $currentPath = Join-Path $currentRoot $source.currentFile
    $exists = Test-Path -LiteralPath $currentPath -PathType Leaf
    $delimiter = if ($source.delimiter) { [char][string]$source.delimiter } else { ',' }
    $header = if ($exists) { Get-CsvHeader -Path $currentPath -Delimiter $delimiter } else { @() }
    $missingColumns = @($source.requiredColumns | Where-Object { $_ -notin $header })
    $file = if ($exists) { Get-Item -LiteralPath $currentPath } else { $null }
    $historyFiles = foreach ($relativeDirectory in $source.historyDirectories) {
        $candidateRoot = Join-Path $historyRoot $relativeDirectory
        if (Test-Path -LiteralPath $candidateRoot -PathType Container) {
            Get-ChildItem -LiteralPath $candidateRoot -Recurse -File -Filter $source.historyFilePattern -ErrorAction SilentlyContinue
        }
    }
    $historyFiles = @($historyFiles | Sort-Object FullName -Unique)
    $ageHours = if ($file) { [math]::Round(($now - $file.LastWriteTime).TotalHours, 1) } else { $null }
    $status = if (-not $exists) {
        'Missing'
    }
    elseif ($missingColumns.Count -gt 0) {
        'SchemaMismatch'
    }
    elseif ($ageHours -gt $FreshnessWarningHours) {
        'Stale'
    }
    else {
        'Ready'
    }

    [pscustomobject]@{
        SourceId            = $source.id
        Domain              = $source.domain
        CurrentFile         = $source.currentFile
        Status              = $status
        SizeMB              = if ($file) { [math]::Round($file.Length / 1MB, 2) } else { $null }
        LastWriteTime       = if ($file) { $file.LastWriteTime.ToString('o') } else { $null }
        AgeHours            = $ageHours
        RequiredColumnCount = @($source.requiredColumns).Count
        MissingColumns      = $missingColumns
        HistoryFileCount    = $historyFiles.Count
        LoadPriority        = $source.loadPriority
    }
}

$summary = [pscustomobject]@{
    ManifestPath         = $resolvedManifest
    DataRoot             = $DataRoot
    CheckedAt            = $now.ToString('o')
    SourceCount          = @($results).Count
    ReadyCount           = @($results | Where-Object Status -eq 'Ready').Count
    StaleCount           = @($results | Where-Object Status -eq 'Stale').Count
    MissingCount         = @($results | Where-Object Status -eq 'Missing').Count
    SchemaMismatchCount  = @($results | Where-Object Status -eq 'SchemaMismatch').Count
    Sources              = @($results)
    UnsupportedIndicators = @($manifest.unsupportedIndicators)
}

$summary | ConvertTo-Json -Depth 12

if ($summary.MissingCount -gt 0 -or $summary.SchemaMismatchCount -gt 0) {
    exit 2
}
