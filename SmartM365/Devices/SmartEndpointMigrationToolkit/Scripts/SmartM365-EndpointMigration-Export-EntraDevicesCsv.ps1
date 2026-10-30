<#
.SYNOPSIS
Exports Microsoft Entra devices to DevicesEntra.csv for the Smart Endpoint Migration Toolkit.

.DESCRIPTION
Uses Microsoft Graph to read Entra devices from /devices.
Without -OutputPath, an unscoped export writes the shared cache <Package>\Inventory\DevicesEntra.csv
and a -ComputerListPath export writes DevicesEntra.csv next to that Computers.txt. A scoped export
never overwrites the shared cache. InventoryTenantId, InventoryAuthenticationMode, and InventoryScope
preserve the delegated full-inventory provenance used by the automatic Wave engines.
Provenance: copied from SmartM365/Devices/IntuneHybridJoinToolkit/Scripts/SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 1.0.2 (see Docs\PROVENANCE.json.txt).

.PARAMETER OutputPath
Destination CSV path. Defaults to the shared cache for unscoped exports, or DevicesEntra.csv next to -ComputerListPath.

.PARAMETER ComputerListPath
Computers.txt path. When provided, only those Entra devices are queried.

.PARAMETER PageSize
Graph page size. Defaults to 999.

.PARAMETER TenantId
Optional tenant id passed to Connect-MgGraph.

.PARAMETER NoConnect
Do not call Connect-MgGraph. Use this only when the current PowerShell session is already connected.

.PARAMETER SkipModuleInstall
Do not install Microsoft.Graph.Authentication automatically if it is missing.

.PARAMETER ForceRefresh
Regenerates the shared cache even when it is less than 2 hours old.

.VERSION
1.0.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$ComputerListPath,
    [int]$PageSize = 999,
    [string]$TenantId,
    [switch]$NoConnect,
    [switch]$SkipModuleInstall,
    [switch]$ForceRefresh
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "1.0.0"
$EntraInventoryFreshnessHours = 24

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

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$CoreModulePath = [System.IO.Path]::GetFullPath((Join-Path $BaseDir "..\Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1"))
Import-Module $CoreModulePath -MinimumVersion 1.0.0 -Force
$PackageRoot = Get-SemtPackageRoot -StartPath $BaseDir
$SharedInventoryPath = [System.IO.Path]::GetFullPath((Get-SemtSharedInventoryPath -PackageRoot $PackageRoot -Kind Entra))

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
$ComputerListPathWasProvided = -not [string]::IsNullOrWhiteSpace($ComputerListPath)
if (-not $OutputPathWasProvided) {
    if ($ComputerListPathWasProvided) {
        $OutputPath = Join-Path (Split-Path -Parent ([System.IO.Path]::GetFullPath($ComputerListPath))) "DevicesEntra.csv"
    }
    else {
        $OutputPath = $SharedInventoryPath
    }
}
$OutputIsSharedInventory = ([System.IO.Path]::GetFullPath($OutputPath) -ieq $SharedInventoryPath)
if ($ComputerListPathWasProvided -and $OutputIsSharedInventory) {
    throw "SEMT_SCOPED_INVENTORY_TARGETS_SHARED_CACHE: a -ComputerListPath export cannot overwrite the shared inventory $SharedInventoryPath."
}

if ($PageSize -lt 1) { $PageSize = 1 }
if ($PageSize -gt 999) { $PageSize = 999 }

if (-not $ForceRefresh -and -not $ComputerListPathWasProvided -and $OutputIsSharedInventory -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $sharedInventoryItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    $sharedInventoryAge = (Get-Date) - $sharedInventoryItem.LastWriteTime
    if ($sharedInventoryAge.TotalHours -le $EntraInventoryFreshnessHours) {
        Write-Host "Smart Endpoint Migration Toolkit - Export-EntraDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
        Write-Host ("Recent shared DevicesEntra.csv found: {0}" -f $OutputPath) -ForegroundColor Green
        Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $sharedInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $sharedInventoryAge.TotalHours) -ForegroundColor Green
        Write-Host "No new Graph export generated. Use -ForceRefresh to regenerate anyway." -ForegroundColor Yellow
        return
    }
}

function Enable-Tls12 {
    try {
        $tls12 = [Net.SecurityProtocolType]::Tls12
        if (([Net.ServicePointManager]::SecurityProtocol -band $tls12) -ne $tls12) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tls12
        }
    }
    catch {
        Write-Host ("WARN: Could not enable TLS 1.2: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Install-GraphModuleIfMissing {
    param([switch]$SkipInstall)

    $module = Find-SemtGraphAuthenticationModule
    if ($module) { return $module }

    if ($SkipInstall) {
        throw ('SEMT_GRAPH_MODULE_NOT_FOUND: Microsoft.Graph.Authentication is not installed. {0}' -f (Get-SemtGraphModuleDiagnostic))
    }

    Write-Host "Microsoft.Graph.Authentication module not found. Installing Microsoft.Graph.Authentication for current user..." -ForegroundColor Yellow
    Enable-Tls12

    try {
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nugetProvider) {
            Write-Host "Installing NuGet package provider for current user..." -ForegroundColor Yellow
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Host ("WARN: NuGet provider installation/check failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    $module = Find-SemtGraphAuthenticationModule
    if (-not $module) { throw ('SEMT_GRAPH_MODULE_NOT_FOUND: Microsoft.Graph.Authentication was installed for the current user but cannot be found. {0}' -f (Get-SemtGraphModuleDiagnostic)) }
    return $module
}

function Import-GraphAuthenticationModule {
    $module = Install-GraphModuleIfMissing -SkipInstall:$SkipModuleInstall
    Import-Module -Name $module.Path -ErrorAction Stop
    Write-Host ("Graph module : Microsoft.Graph.Authentication {0}" -f $module.Version) -ForegroundColor DarkCyan
}

function Get-GraphConnectionContext {
    try { return Get-MgContext -ErrorAction Stop } catch { return $null }
}

function Get-GraphProperty {
    param(
        [Parameter(Mandatory=$true)]$InputObject,
        [Parameter(Mandatory=$true)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function ConvertTo-ComputerName {
    param([string]$DeviceName)
    if ([string]::IsNullOrWhiteSpace($DeviceName)) { return "" }
    return ($DeviceName.Split(".")[0]).Trim().ToUpperInvariant()
}

function ConvertTo-ODataStringLiteral {
    param([Parameter(Mandatory=$true)][string]$Value)

    return ("'{0}'" -f ($Value -replace "'", "''"))
}

function Get-ComputerList {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Computer list not found: $Path"
    }

    @(Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") } |
        ForEach-Object {
            [PSCustomObject]@{
                RequestedComputerName = $_
                ComputerName = ConvertTo-ComputerName -DeviceName $_
            }
        } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.ComputerName) } |
        Sort-Object ComputerName -Unique)
}

function Get-CollectionCount {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 0 }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return 0 }
        return 1
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $count = 0
        foreach ($item in $Value) {
            if ($null -ne $item) { $count++ }
        }
        return $count
    }
    return 1
}

function Invoke-GraphPagedRequest {
    param([Parameter(Mandatory=$true)][string]$Uri,[switch]$Quiet)

    # One short progress line per page (the page URL with its skip token is not useful to read).
    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    $page = 0
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page++
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
        $pageItems = Get-GraphProperty -InputObject $response -Name "value"
        if ($pageItems) {
            foreach ($item in @($pageItems)) {
                if ($null -ne $item) { $items.Add($item) }
            }
        }
        if (-not $Quiet) { Write-Host ("Graph page {0}: {1} device(s) read so far." -f $page,$items.Count) -ForegroundColor DarkCyan }
        $nextUri = ""
        $nextLink = Get-GraphProperty -InputObject $response -Name "@odata.nextLink"
        if (-not [string]::IsNullOrWhiteSpace([string]$nextLink)) { $nextUri = [string]$nextLink }
    }
    return @($items.ToArray())
}

function New-EntraDeviceExportRow {
    param([Parameter(Mandatory=$true)]$Device)

    $displayName = [string](Get-GraphProperty -InputObject $Device -Name "displayName")
    $registrationDateTime = [string](Get-GraphProperty -InputObject $Device -Name "registrationDateTime")
    $trustType = [string](Get-GraphProperty -InputObject $Device -Name "trustType")
    $alternativeSecurityIdCount = Get-CollectionCount -Value (Get-GraphProperty -InputObject $Device -Name "alternativeSecurityIds")
    $registeredState = if ($trustType -ieq "ServerAd" -and $alternativeSecurityIdCount -eq 0) { "Pending" } else { "Registered" }
    $pendingReason = if ($registeredState -eq "Pending") { "ServerAd trustType with empty AlternativeSecurityIds" } else { "" }

    [PSCustomObject]@{
        ComputerName                = ConvertTo-ComputerName -DeviceName $displayName
        DisplayName                 = $displayName
        EntraInventoryPresent       = $true
        EntraRegisteredState        = $registeredState
        EntraPendingReason          = $pendingReason
        RegistrationDateTime        = $registrationDateTime
        EntraObjectId               = Get-GraphProperty -InputObject $Device -Name "id"
        DeviceId                    = Get-GraphProperty -InputObject $Device -Name "deviceId"
        TrustType                   = $trustType
        AlternativeSecurityIdCount  = $alternativeSecurityIdCount
        AccountEnabled             = Get-GraphProperty -InputObject $Device -Name "accountEnabled"
        OperatingSystem            = Get-GraphProperty -InputObject $Device -Name "operatingSystem"
        OperatingSystemVersion     = Get-GraphProperty -InputObject $Device -Name "operatingSystemVersion"
        IsManaged                  = Get-GraphProperty -InputObject $Device -Name "isManaged"
        IsCompliant                = Get-GraphProperty -InputObject $Device -Name "isCompliant"
        ManagementType             = Get-GraphProperty -InputObject $Device -Name "managementType"
        ProfileType                = Get-GraphProperty -InputObject $Device -Name "profileType"
        ApproximateLastSignInDateTime = Get-GraphProperty -InputObject $Device -Name "approximateLastSignInDateTime"
    }
}

function New-MissingEntraDeviceExportRow {
    param(
        [string]$RequestedComputerName,
        [string]$ComputerName
    )

    [PSCustomObject]@{
        ComputerName                = $ComputerName
        DisplayName                 = $RequestedComputerName
        EntraInventoryPresent       = $false
        EntraRegisteredState        = ""
        EntraPendingReason          = ""
        RegistrationDateTime        = ""
        EntraObjectId               = ""
        DeviceId                    = ""
        TrustType                   = ""
        AlternativeSecurityIdCount  = ""
        AccountEnabled             = ""
        OperatingSystem            = ""
        OperatingSystemVersion     = ""
        IsManaged                  = ""
        IsCompliant                = ""
        ManagementType             = ""
        ProfileType                = ""
        ApproximateLastSignInDateTime = ""
    }
}

Import-GraphAuthenticationModule

if (-not $NoConnect) {
    $connectParams = @{
        Scopes = @("Device.Read.All")
        NoWelcome = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParams.TenantId = $TenantId }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph @connectParams | Out-Null
}

$context = Get-GraphConnectionContext
if (-not $context) {
    throw "Not connected to Microsoft Graph. Run without -NoConnect, or connect first with Connect-MgGraph -Scopes Device.Read.All."
}
$contextAuthType = if ($context.PSObject.Properties["AuthType"]){ [string]$context.AuthType } else { "" }
if ($contextAuthType -ieq "AppOnly") {
    throw "App-only Microsoft Graph authentication is not supported. Use delegated interactive authentication."
}
$inventoryAuthenticationMode = if ($NoConnect) { "DelegatedExistingSession" } else { "DelegatedInteractive" }

Write-Host "Smart Endpoint Migration Toolkit - Export-EntraDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
Write-Host "Tenant      : $($context.TenantId)"
Write-Host "Account     : $($context.Account)"
Write-Host "Auth        : $inventoryAuthenticationMode"
Write-Host "Output      : $OutputPath"
Write-Host "Page size   : $PageSize"

$requestedComputers = @()
if ($ComputerListPathWasProvided) {
    $requestedComputers = @(Get-ComputerList -Path $ComputerListPath)
    Write-Host "Computers   : $ComputerListPath"
    Write-Host ("Requested   : {0}" -f $requestedComputers.Count)
}
else {
    Write-Host "Computers   : no -ComputerListPath; exporting all Entra devices." -ForegroundColor Yellow
}

$select = @(
    "id",
    "displayName",
    "deviceId",
    "accountEnabled",
    "operatingSystem",
    "operatingSystemVersion",
    "trustType",
    "isCompliant",
    "isManaged",
    "managementType",
    "profileType",
    "alternativeSecurityIds",
    "registrationDateTime",
    "approximateLastSignInDateTime"
) -join ","

$devices = @()
if ($requestedComputers.Count -gt 0) {
    foreach ($computer in $requestedComputers) {
        $candidateNames = @($computer.RequestedComputerName,$computer.ComputerName) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique

        foreach ($candidateName in $candidateNames) {
            $filter = "displayName eq {0}" -f (ConvertTo-ODataStringLiteral -Value ([string]$candidateName))
            $encodedFilter = [System.Uri]::EscapeDataString($filter)
            $uri = "https://graph.microsoft.com/v1.0/devices?`$select=$select&`$filter=$encodedFilter&`$top=$PageSize"
            $devices += @(Invoke-GraphPagedRequest -Uri $uri -Quiet)
        }
    }
    Write-Host ("Graph lookups: {0} computer(s) queried; {1} Entra device(s) found." -f $requestedComputers.Count,@($devices).Count) -ForegroundColor DarkCyan

    $deviceByComputer = @{}
    foreach ($device in $devices) {
        $displayName = [string](Get-GraphProperty -InputObject $device -Name "displayName")
        $computerName = ConvertTo-ComputerName -DeviceName $displayName
        if ([string]::IsNullOrWhiteSpace($computerName)) { continue }
        if (-not $deviceByComputer.ContainsKey($computerName)) { $deviceByComputer[$computerName] = $device }
    }

    $export = foreach ($computer in $requestedComputers) {
        if ($deviceByComputer.ContainsKey($computer.ComputerName)) {
            New-EntraDeviceExportRow -Device $deviceByComputer[$computer.ComputerName]
        }
        else {
            New-MissingEntraDeviceExportRow -RequestedComputerName $computer.RequestedComputerName -ComputerName $computer.ComputerName
        }
    }
}
else {
    $uri = "https://graph.microsoft.com/v1.0/devices?`$select=$select&`$top=$PageSize"
    $devices = Invoke-GraphPagedRequest -Uri $uri
    $export = foreach ($device in $devices) { New-EntraDeviceExportRow -Device $device }
}

$export = @($export)
$inventoryScope = if ($requestedComputers.Count -gt 0) { "RequestedComputers" } else { "AllEntraDevices" }
foreach ($row in $export) {
    $row | Add-Member -NotePropertyName InventoryTenantId -NotePropertyValue ([string]$context.TenantId) -Force
    $row | Add-Member -NotePropertyName InventoryAuthenticationMode -NotePropertyValue $inventoryAuthenticationMode -Force
    $row | Add-Member -NotePropertyName InventoryScope -NotePropertyValue $inventoryScope -Force
}

try {
    $outputDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path -LiteralPath $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $finalOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    $tempOutputPath = "{0}.tmp.{1}.{2}.csv" -f $finalOutputPath, $PID, ([guid]::NewGuid().ToString("N"))

    $export |
        Sort-Object ComputerName, DisplayName |
        Export-Csv -LiteralPath $tempOutputPath -NoTypeInformation -Encoding UTF8

    Copy-Item -LiteralPath $tempOutputPath -Destination $finalOutputPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $tempOutputPath -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host ""
    Write-Host ("ERROR: Cannot write DevicesEntra.csv to: {0}" -f $OutputPath) -ForegroundColor Red
    Write-Host ("Detail: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($tempOutputPath) -and (Test-Path -LiteralPath $tempOutputPath)) {
        Write-Host ("Temporary CSV kept for troubleshooting: {0}" -f $tempOutputPath) -ForegroundColor Yellow
    }
    Write-Host "Run the CMD as administrator, close the CSV if it is open in Excel, or use -OutputPath with a writable folder." -ForegroundColor Yellow
    throw
}

$uniqueComputers = @($export | Where-Object { -not [string]::IsNullOrWhiteSpace($_.ComputerName) } | Select-Object -ExpandProperty ComputerName -Unique)
$pendingCount = @($export | Where-Object { $_.EntraRegisteredState -eq "Pending" }).Count

Write-Host ""
Write-Host ("Exported Entra devices: {0}" -f @($export).Count) -ForegroundColor Green
Write-Host ("Unique computer names : {0}" -f $uniqueComputers.Count) -ForegroundColor Green
Write-Host ("Pending registrations : {0}" -f $pendingCount) -ForegroundColor Green
if ($requestedComputers.Count -gt 0) {
    $presentCount = @($export | Where-Object { $_.EntraInventoryPresent -eq $true }).Count
    Write-Host ("Requested computers   : {0}" -f $requestedComputers.Count) -ForegroundColor Green
    Write-Host ("Present in Entra     : {0}" -f $presentCount) -ForegroundColor Green
    Write-Host ("Missing from Entra    : {0}" -f ($requestedComputers.Count - $presentCount)) -ForegroundColor Green
}
Write-Host ("CSV                   : {0}" -f $OutputPath) -ForegroundColor Green

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAl0CgBegZd7MBy
# vxiwbQjqIZujeFOlx6aaCCZfYKPbEqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAudagQ01d3E8B+knJxHKQO
# FDMQo2+SuWBKKlrbAQJXbDANBgkqhkiG9w0BAQEFAASCAYAO/bKWllAHxgm7nlgN
# orAv83/iqQbsCFhBc0xNWL/aaH5Wt/IJ2DJphqayXae25h9T8eJpvb3Q1fW9YwWq
# 6Zb1nUkhPORiI0k4W18YFLLaJ+ZqLeJZBt/bmqUX7wwiVSRDfc6vQvwl8vAx2Ylx
# QmymLhuoFXZ0OTeQQNXr6ZJz9WBdLTVPwMiSTDOKQfsvb5ZLISWuK3A47NHXzkl6
# IgRWbACq9GRBh6+2RxlVjwvHZtDerNr8dY8L0CPlNN6lOTB2x3zZIm2tKOONZlJM
# +76D2g6H9j+IywTY5oKF9Xz7DPQ4CNyIZf/RwDqkD2F+yMdyxCECdTLWK5ci41Zw
# L+/dboiokvTHshqsej7pEB3hDLL9NrZ6d7Dg6PqPeTEGyTwnYLC9189rP/xN7GX/
# UuUXz7reFVElB87pqbyPm8co1BAPU3dNJWkqXQeDyUr9cqA8aYla6bA2P8XY3dZJ
# y4NsWHRMekSX/IAFh7Yl8E96Eg1TAnk5wvazn3u8LFpC3Yw=
# SIG # End signature block
