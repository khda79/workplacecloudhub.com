<#
.SYNOPSIS
    Regression test for the Discovered Apps cache validation path.
.DESCRIPTION
    Loads only the cache-related functions from the inventory script, creates an
    incomplete manifest plus a duplicate app/device relation, and verifies that
    the contaminated app is rejected while a valid app is reused.
    No Microsoft Graph connection or production CSV is used.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$inventoryScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Intune-DiscoveredApps-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    throw ($parseErrors | Out-String)
}

$requiredFunctions = @(
    'Write-DiscoveredAppsCsvRows',
    'Get-DiscoveredAppsAppDeviceCount',
    'Get-DiscoveredAppsDeviceDetailCacheManifestPath',
    'ConvertTo-DiscoveredAppsCacheStatsMap',
    'Read-DiscoveredAppsDeviceDetailCacheManifest',
    'Use-DiscoveredAppsDeviceDetailCache'
)
$definitions = @($ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $requiredFunctions -contains $node.Name
}, $true) | Sort-Object { $_.Extent.StartOffset })
if ($definitions.Count -ne $requiredFunctions.Count) {
    throw "Expected $($requiredFunctions.Count) cache function definitions, found $($definitions.Count)."
}
foreach ($definition in $definitions) {
    Invoke-Expression $definition.Extent.Text
}

$script:TestLogs = [System.Collections.Generic.List[string]]::new()
function WriteLog {
    param([string]$Message, [string]$Level)
    $script:TestLogs.Add("$Level|$Message") | Out-Null
}
function Add-SmartM365TenantKey {
    process {
        [pscustomobject]@{
            TenantKey = 'tenant-test'
            AppId     = [string]$_.AppId
            DeviceId  = [string]$_.DeviceId
        }
    }
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) {
        throw "$Label expected '$Expected', got '$Actual'."
    }
}

$testRoot = Join-Path $env:TEMP ("SmartM365-DiscoveredApps-CacheTest-" + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    $cachePath = Join-Path $testRoot 'Intune_DiscoveredApps_AppDeviceRelations.csv'
    $partialPath = Join-Path $testRoot 'partial.csv'
    @(
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-contaminated'; DeviceId = 'device-2' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-valid'; DeviceId = 'device-3' }
    ) | Export-Csv -LiteralPath $cachePath -NoTypeInformation -Encoding UTF8

    $cacheItem = Get-Item -LiteralPath $cachePath
    $manifestPath = Get-DiscoveredAppsDeviceDetailCacheManifestPath -CsvPath $cachePath
    [ordered]@{
        CacheManifestVersion = 2
        SourceCsvLength      = [int64]$cacheItem.Length
        AppCount             = 1
        TotalRows            = 2
        TotalDeviceRows      = 2
        Stats                = @(
            [pscustomobject]@{
                AppId = 'app-contaminated'; AppName = 'Contaminated'; AppVersion = '1.0'
                Publisher = 'Publisher'; Platform = 'windows'; Rows = 2; DeviceRows = 2
                MetadataOk = $true; EnrichmentOk = $true
            }
        )
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    $targetApps = @(
        [pscustomobject]@{
            id = 'app-contaminated'; displayName = 'Contaminated'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 2
        }
        [pscustomobject]@{
            id = 'app-valid'; displayName = 'Valid'; version = '1.0'
            publisher = 'Publisher'; platform = 'windows'; deviceCount = 1
        }
    )
    $processed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $cached = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $actualCounts = @{}

    $result = Use-DiscoveredAppsDeviceDetailCache `
        -CachePath $cachePath `
        -TargetApps $targetApps `
        -PartialPath $partialPath `
        -ProcessedAppIds $processed `
        -CachedAppIds $cached `
        -ActualDeviceCountsByAppId $actualCounts `
        -MaxAgeDays 7

    Assert-Equal $result.Used $true 'Cache Used'
    Assert-Equal $result.ManifestUsed $false 'ManifestUsed'
    Assert-Equal $result.Apps 1 'Reusable apps'
    Assert-Equal $result.Rows 1 'Reused rows'
    Assert-Equal $result.RejectedApps 1 'Rejected apps'
    Assert-Equal $result.RejectedRows 3 'Rejected rows'
    Assert-Equal $result.ExcessRows 1 'Excess rows'
    Assert-Equal $processed.Contains('app-valid') $true 'Valid app processed'
    Assert-Equal $processed.Contains('app-contaminated') $false 'Contaminated app excluded'
    Assert-Equal $actualCounts.ContainsKey('app-valid') $true 'Valid actual count present'
    Assert-Equal $actualCounts.ContainsKey('app-contaminated') $false 'Contaminated actual count absent'

    $partialRows = @(Import-Csv -LiteralPath $partialPath)
    Assert-Equal $partialRows.Count 1 'Partial row count'
    Assert-Equal $partialRows[0].AppId 'app-valid' 'Partial AppId'
    Assert-Equal $partialRows[0].DeviceId 'device-3' 'Partial DeviceId'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*incomplete and will not be trusted*' }).Count 1 'Incomplete manifest warning'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*excess rows detected: 1*' }).Count 1 'Excess row warning'

    Write-Host 'PASS: incomplete cache manifest was discarded; contaminated app rejected; valid app reused.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}