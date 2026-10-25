<#
.SYNOPSIS
Exports the full Intune and Entra inventories used by the automatic Wave engines.

.DESCRIPTION
Uses one delegated interactive Microsoft Graph connection with the union of the
required read scopes. Intune is mandatory; Entra is optional enrichment and is
reported as a warning when its export fails. Without explicit output paths, the
shared caches <Package>\Inventory\DevicesIntune.csv and DevicesEntra.csv are written.
Provenance: copied from SmartM365/Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-IntuneHybridJoinRepair-Export-AutomaticGraphInventories.ps1 1.0.0 (see Docs\PROVENANCE.json.txt).

.PARAMETER IntuneOutputPath
Intune CSV path. Defaults to the shared cache <Package>\Inventory\DevicesIntune.csv.

.PARAMETER EntraOutputPath
Entra CSV path. Defaults to the shared cache <Package>\Inventory\DevicesEntra.csv.

.VERSION
1.0.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$IntuneOutputPath,
    [string]$EntraOutputPath,
    [string]$TenantId,
    [switch]$SkipModuleInstall
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )

    foreach ($line in @(([string]$Object) -split "`r?`n")) {
        $parameters = @{ Object = ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line) }
        if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters['ForegroundColor'] = $ForegroundColor }
        if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $parameters['BackgroundColor'] = $BackgroundColor }
        if ($PSBoundParameters.ContainsKey('NoNewline')) { $parameters['NoNewline'] = $NoNewline }
        Microsoft.PowerShell.Utility\Write-Host @parameters
    }
}

function Enable-Tls12 {
    try {
        $tls12 = [Net.SecurityProtocolType]::Tls12
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tls12
    }
    catch {
        Write-Host ("WARN: Could not enable TLS 1.2: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Import-GraphAuthentication {
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if (-not $module) {
        if ($SkipModuleInstall) {
            throw 'Microsoft.Graph.Authentication is not installed.'
        }
        Enable-Tls12
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

$scriptsRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$coreModulePath = [System.IO.Path]::GetFullPath((Join-Path $scriptsRoot '..\Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'))
Import-Module $coreModulePath -MinimumVersion 1.0.0 -Force
$packageRoot = Get-SemtPackageRoot -StartPath $scriptsRoot
if ([string]::IsNullOrWhiteSpace($IntuneOutputPath)) { $IntuneOutputPath = Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind Intune }
if ([string]::IsNullOrWhiteSpace($EntraOutputPath)) { $EntraOutputPath = Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind Entra }

$intuneExporter = Join-Path $scriptsRoot 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1'
$entraExporter = Join-Path $scriptsRoot 'SmartM365-EndpointMigration-Export-EntraDevicesCsv.ps1'
foreach ($path in @($intuneExporter, $entraExporter)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Exporter not found: $path" }
}

Write-Host "Smart Endpoint Migration Toolkit - Automatic Graph inventory exporter version $ScriptVersion" -ForegroundColor Cyan
Import-GraphAuthentication
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
$connectParameters = @{
    Scopes = @('DeviceManagementManagedDevices.Read.All', 'Device.Read.All')
    NoWelcome = $true
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParameters.TenantId = $TenantId.Trim() }

$entraSucceeded = $false
$context = $null
try {
    Write-Host 'Waiting for delegated interactive Microsoft Graph authentication...' -ForegroundColor Cyan
    Connect-MgGraph @connectParameters | Out-Null
    $context = Get-MgContext -ErrorAction Stop
    if (-not $context) { throw 'Microsoft Graph context was not created.' }
    $authType = if ($context.PSObject.Properties['AuthType']) { [string]$context.AuthType } else { '' }
    if ($authType -ieq 'AppOnly') {
        throw 'App-only Microsoft Graph authentication is not supported. Use delegated interactive authentication.'
    }
    Write-Host ("Tenant      : {0}" -f $context.TenantId)
    Write-Host ("Account     : {0}" -f $context.Account)

    Write-Host 'Exporting the required full Intune managed-device inventory...' -ForegroundColor Cyan
    & $intuneExporter -OutputPath $IntuneOutputPath -NoConnect -ForceRefresh -SkipModuleInstall
    if (-not (Test-Path -LiteralPath $IntuneOutputPath -PathType Leaf)) {
        throw "Intune inventory was not created: $IntuneOutputPath"
    }

    try {
        Write-Host 'Exporting the optional full Entra device inventory...' -ForegroundColor Cyan
        & $entraExporter -OutputPath $EntraOutputPath -NoConnect -ForceRefresh -SkipModuleInstall
        if (-not (Test-Path -LiteralPath $EntraOutputPath -PathType Leaf)) {
            throw "Entra inventory was not created: $EntraOutputPath"
        }
        $entraSucceeded = $true
    }
    catch {
        Remove-Item -LiteralPath $EntraOutputPath -Force -ErrorAction SilentlyContinue
        Write-Host ("WARNING: Entra enrichment export failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        Write-Host 'The automatic Wave can continue from the required AD and Intune inventories.' -ForegroundColor Yellow
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

[pscustomobject]@{
    TenantId = if ($context) { [string]$context.TenantId } else { '' }
    AuthenticationMode = 'DelegatedInteractive'
    IntuneOutputPath = $IntuneOutputPath
    EntraOutputPath = if ($entraSucceeded) { $EntraOutputPath } else { '' }
    EntraSucceeded = $entraSucceeded
}
