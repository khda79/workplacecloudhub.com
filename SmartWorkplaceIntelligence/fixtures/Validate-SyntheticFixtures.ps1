[CmdletBinding()]
param(
    [Parameter()]
    [string]$FixtureRoot = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$requiredCommonColumns = @(
    'ContractVersion',
    'SourceFamily',
    'SourceEntity',
    'TenantId',
    'OrganizationId',
    'EnvironmentId',
    'RunId',
    'CollectedAtUtc',
    'CollectionStatus',
    'IsPartialInventory',
    'SourceScriptVersion'
)

$manifestPath = Join-Path $FixtureRoot 'fixture-manifest.json'
$contractPath = Join-Path $FixtureRoot 'source-contract.json'

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Fixture manifest not found: $manifestPath"
}

if (-not (Test-Path -LiteralPath $contractPath -PathType Leaf)) {
    throw "Source contract not found: $contractPath"
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -Depth 20
$null = Get-Content -LiteralPath $contractPath -Raw | ConvertFrom-Json -Depth 20

if ($manifest.syntheticOnly -ne $true) {
    throw 'The manifest must declare syntheticOnly=true.'
}

$manifestPaths = @($manifest.files | ForEach-Object { $_.path.Replace('/', [IO.Path]::DirectorySeparatorChar) })
$diskPaths = @(
    Get-ChildItem -LiteralPath (Join-Path $FixtureRoot 'valid'), (Join-Path $FixtureRoot 'quality-cases') -Filter '*.csv' -File |
        ForEach-Object { [IO.Path]::GetRelativePath($FixtureRoot, $_.FullName) } |
        Sort-Object
)

$missingFromManifest = @($diskPaths | Where-Object { $_ -notin $manifestPaths })
$missingFromDisk = @($manifestPaths | Where-Object { $_ -notin $diskPaths })

if ($missingFromManifest.Count -gt 0 -or $missingFromDisk.Count -gt 0) {
    throw "Manifest/file mismatch. Unlisted=$($missingFromManifest -join ','); Missing=$($missingFromDisk -join ',')"
}

$sha256 = [Security.Cryptography.SHA256]::Create()
$accepted = 0
$rejected = 0
$totalRows = 0

foreach ($entry in $manifest.files) {
    $relativePath = $entry.path.Replace('/', [IO.Path]::DirectorySeparatorChar)
    $path = Join-Path $FixtureRoot $relativePath
    $lines = @(Get-Content -LiteralPath $path | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($lines.Count -eq 0) {
        throw "Empty fixture file: $relativePath"
    }

    $header = $lines[0]
    $schemaHash = [Convert]::ToHexString($sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($header)))
    $fileHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash
    $rows = @(Import-Csv -LiteralPath $path)

    if ($schemaHash -ne $entry.schemaSha256) {
        throw "Schema hash mismatch: $relativePath"
    }

    if ($fileHash -ne $entry.fileSha256) {
        throw "File hash mismatch: $relativePath"
    }

    if ($rows.Count -ne $entry.expectedRows) {
        throw "Row count mismatch: $relativePath expected=$($entry.expectedRows) actual=$($rows.Count)"
    }

    $headerCount = ($header -split ',').Count
    for ($lineIndex = 1; $lineIndex -lt $lines.Count; $lineIndex++) {
        if (($lines[$lineIndex] -split ',').Count -ne $headerCount) {
            throw "CSV field-count mismatch: $relativePath line=$($lineIndex + 1)"
        }
    }

    if ($entry.expectedOutcome -eq 'Accept') {
        $columns = @($rows[0].PSObject.Properties.Name)
        $missingColumns = @($requiredCommonColumns | Where-Object { $_ -notin $columns })
        if ($missingColumns.Count -gt 0) {
            throw "Accepted fixture is missing required columns: $relativePath ($($missingColumns -join ','))"
        }

        foreach ($row in $rows) {
            if ($row.ContractVersion -ne $manifest.contractVersion) {
                throw "Contract version mismatch: $relativePath"
            }

            if ($row.TenantId -ne $manifest.defaultContext.tenantId -or
                $row.OrganizationId -ne $manifest.defaultContext.organizationId -or
                $row.EnvironmentId -ne $manifest.defaultContext.environmentId) {
                throw "Synthetic context mismatch: $relativePath"
            }

            $timestamp = [DateTimeOffset]::MinValue
            $validTimestamp = [DateTimeOffset]::TryParse(
                $row.CollectedAtUtc,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::RoundtripKind,
                [ref]$timestamp
            )

            if (-not $validTimestamp -or $row.CollectedAtUtc -notmatch '(Z|[+-][0-9]{2}:[0-9]{2})$') {
                throw "Accepted fixture has an invalid UTC/offset timestamp: $relativePath"
            }
        }

        $accepted++
    }
    elseif ($entry.expectedOutcome -eq 'Reject') {
        if ([string]::IsNullOrWhiteSpace($entry.expectedReason)) {
            throw "Rejected fixture has no expected reason: $relativePath"
        }
        $rejected++
    }
    else {
        throw "Unsupported expected outcome: $($entry.expectedOutcome)"
    }

    $totalRows += $rows.Count
}

[pscustomobject]@{
    Status = 'PASS'
    ContractVersion = $manifest.contractVersion
    FixtureFiles = $manifest.files.Count
    AcceptedCases = $accepted
    RejectedCases = $rejected
    TotalRows = $totalRows
    SyntheticOnly = $manifest.syntheticOnly
}
