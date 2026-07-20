<#
.SYNOPSIS
    Offline regression test for Discovered Apps structured logging.
.DESCRIPTION
    Verifies 429 status inference, quiet retry error streams, explicit final CSV
    sample/physical row labels, and quiet retirement of an already absent
    SharePoint legacy file. No Graph connection or production CSV is used.
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
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }

$requiredFunctions = @(
    'Get-DiscoveredAppsGraphHttpErrorMessage',
    'Invoke-GraphPagedRequest',
    'Get-DiscoveredAppsCsvDataRowCount',
    'Complete-DiscoveredAppsStreamExport',
    'Remove-LegacyDiscoveredAppsDeviceDetailExport'
)
$definitions = @($ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $requiredFunctions -contains $node.Name
}, $true) | Sort-Object { $_.Extent.StartOffset })
if ($definitions.Count -ne $requiredFunctions.Count) {
    throw "Expected $($requiredFunctions.Count) function definitions, found $($definitions.Count)."
}
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psm1')).Path
$coreTokens = $null
$coreParseErrors = $null
$coreAst = [System.Management.Automation.Language.Parser]::ParseFile($coreModulePath, [ref]$coreTokens, [ref]$coreParseErrors)
if ($coreParseErrors.Count -gt 0) { throw ($coreParseErrors | Out-String) }
$coreDeleteDefinition = $coreAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Invoke-SmartM365GraphDeleteQuietly'
}, $true) | Select-Object -First 1
if (-not $coreDeleteDefinition) { throw 'Invoke-SmartM365GraphDeleteQuietly was not found in SmartM365.Core.' }
Invoke-Expression $coreDeleteDefinition.Extent.Text

$script:TestLogs = [System.Collections.Generic.List[string]]::new()
function WriteLog {
    param([string]$Message, [string]$Level)
    $script:TestLogs.Add("$Level|$Message") | Out-Null
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}

$script:GraphAttempt = 0
$script:Stat_GraphCalls = 0
$script:Stat_ThrottleRetries = 0
$script:GraphRetryMaxSeconds = 1
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param(
        [string]$Method,
        [string]$Uri,
        [string]$OutputType,
        [string]$Body,
        [string]$ContentType,
        [switch]$SkipHttpErrorCheck,
        [string]$StatusCodeVariable
    )
    $script:GraphAttempt++
    if ($script:GraphAttempt -eq 1) {
        Set-Variable -Name $StatusCodeVariable -Value 429 -Scope 1
        return [pscustomobject]@{ error = [pscustomobject]@{ message = 'TooManyRequests synthetic test response' } }
    }
    Set-Variable -Name $StatusCodeVariable -Value 200 -Scope 1
    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'device-1' }) }
}
function Get-ShortGraphErrorMessage { param($ErrorRecord) return [string]$ErrorRecord.Exception.Message }
function Get-GraphRetryDelaySeconds {
    param($ErrorRecord, [int]$Attempt, [int]$DefaultSeconds, [int]$MaximumSeconds)
    return 0
}
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) }
function Test-SmartM365MaxItemsMode { return $false }
function Remove-SmartM365SharePointFile {
    [CmdletBinding()]
    param([string]$LocalFilePath)
    return $true
}
function Get-SmartM365GraphAccessToken { param([string]$Purpose) return 'synthetic-token' }
function Invoke-RestMethod {
    [CmdletBinding()]
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [switch]$SkipHttpErrorCheck,
        [string]$StatusCodeVariable,
        [string]$ResponseHeadersVariable
    )
    Set-Variable -Name $StatusCodeVariable -Value 404 -Scope 1
    Set-Variable -Name $ResponseHeadersVariable -Value @{} -Scope 1
    return [pscustomobject]@{ error = [pscustomobject]@{ message = 'itemNotFound' } }
}

$testRoot = Join-Path $env:TEMP ("SmartM365-DiscoveredApps-LoggingTest-" + [guid]::NewGuid().ToString('N'))
try {
    $pagedResult = @(Invoke-GraphPagedRequest -InitialUri 'https://graph.microsoft.com/v1.0/test' -MaxRetries 2 -DefaultRetrySeconds 0)
    Assert-Equal $pagedResult.Count 1 'Paged result count'
    Assert-Equal $script:Stat_ThrottleRetries 1 'Throttle retry count'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*Status=429; attempt 1/2*' }).Count 1 'Structured 429 status log'

    $quietDelete = Invoke-SmartM365GraphDeleteQuietly -Uri 'https://graph.microsoft.com/v1.0/test-delete' -MaxAttempts 1
    Assert-Equal $quietDelete.Success $true 'Quiet delete success'
    Assert-Equal $quietDelete.NotFound $true 'Quiet delete not found'
    Assert-Equal $quietDelete.StatusCode 404 'Quiet delete status'

    $outputPath = Join-Path $testRoot 'current'
    $latestPath = Join-Path $testRoot 'latest'
    New-Item -ItemType Directory -Path $outputPath, $latestPath -Force | Out-Null
    $partialPath = Join-Path $testRoot 'relations.partial.csv'
    $timestampedPath = Join-Path $outputPath 'Intune_DiscoveredApps_AppDeviceRelations_20260720_200000.csv'
    @(
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-1'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-2'; DeviceId = 'device-2' }
    ) | Export-Csv -LiteralPath $partialPath -NoTypeInformation -Encoding UTF8

    $global:csvGeneratedPaths = $null
    $publishedPath = Complete-DiscoveredAppsStreamExport `
        -PartialPath $partialPath `
        -TimestampedPath $timestampedPath `
        -OutputPath $outputPath `
        -GlobalPath $latestPath `
        -BaseFileName 'Intune_DiscoveredApps_AppDeviceRelations' `
        -ExpectedDataRows 2
    Assert-Equal $publishedPath $timestampedPath 'Published path'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $latestPath 'Intune_DiscoveredApps_AppDeviceRelations.csv')) $true 'Latest relation CSV'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*SampleRows=1; CriticalFields=TenantKey, AppId, DeviceId*' }).Count 1 'SampleRows log'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*PhysicalRows=2*' }).Count 1 'PhysicalRows log'

    $global:EnableSharePointUpload = $true
    $legacyErrorOutput = @(Remove-LegacyDiscoveredAppsDeviceDetailExport -CurrentOutputPath $outputPath -LatestOutputPath $latestPath 2>&1)
    Assert-Equal @($legacyErrorOutput | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }).Count 0 'Quiet legacy SharePoint 404'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*Legacy export Intune_DiscoveredApps_DeviceDetail.csv is disabled*' }).Count 1 'Legacy retirement log'

    Write-Host 'PASS: structured 429, explicit CSV row labels, and quiet legacy 404 verified.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}