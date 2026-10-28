<#
.SYNOPSIS
Exports Intune managed devices to DevicesIntune.csv for the Smart Endpoint Migration Toolkit.

.DESCRIPTION
Uses Microsoft Graph to read Intune managed devices from:
/deviceManagement/managedDevices

Without -OutputPath, an unscoped export writes the shared cache <Package>\Inventory\DevicesIntune.csv
and a -ComputerListPath export writes DevicesIntune.csv next to that Computers.txt. A scoped export
never overwrites the shared cache. The CSV includes DeviceName and ComputerName columns plus the
columns read by the HybridJoin and Windows11Upgrade automatic Wave engines. InventoryTenantId,
InventoryAuthenticationMode, and InventoryScope record cache provenance for guarded reuse.
Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 1.2.0 (see Docs\PROVENANCE.json.txt).

.PARAMETER OutputPath
Destination CSV path. Defaults to the shared cache for unscoped exports, or DevicesIntune.csv next to -ComputerListPath.

.PARAMETER ComputerListPath
Computers.txt path. When provided, only those Intune managed devices are queried. Otherwise all Intune managed devices are exported.

.PARAMETER PageSize
Graph page size for the full managedDevices read. Defaults to 999.

.PARAMETER TenantId
Optional tenant id used to constrain the delegated interactive sign-in.

.PARAMETER NoConnect
Do not call Connect-MgGraph. Use this only when the current PowerShell session is already connected.

.PARAMETER IncludeAllProperties
Exports the full managedDevice objects returned by Graph instead of the curated default columns.

.PARAMETER SkipModuleInstall
Do not install Microsoft.Graph.Authentication automatically if it is missing.

.PARAMETER ForceRefresh
Regenerates the shared cache even when a recent delegated full inventory already exists.

.EXAMPLE
.\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1

.EXAMPLE
.\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1 -ComputerListPath .\Computers.txt -PageSize 999

.EXAMPLE
.\SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1 -TenantId "contoso.onmicrosoft.com" -OutputPath .\DevicesIntune.csv

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
    [switch]$IncludeAllProperties,
    [switch]$SkipModuleInstall,
    [switch]$ForceRefresh
)

$ErrorActionPreference = "Stop"

$ScriptVersion = "1.0.0"
$IntuneInventoryFreshnessHours = 2

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

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$CoreModulePath = [System.IO.Path]::GetFullPath((Join-Path $BaseDir "..\Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1"))
Import-Module $CoreModulePath -MinimumVersion 1.0.0 -Force
$PackageRoot = Get-SemtPackageRoot -StartPath $BaseDir
$SharedInventoryPath = [System.IO.Path]::GetFullPath((Get-SemtSharedInventoryPath -PackageRoot $PackageRoot -Kind Intune))

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
$ComputerListPathWasProvided = -not [string]::IsNullOrWhiteSpace($ComputerListPath)

if (-not $OutputPathWasProvided) {
    if ($ComputerListPathWasProvided) {
        $OutputPath = Join-Path (Split-Path -Parent ([System.IO.Path]::GetFullPath($ComputerListPath))) "DevicesIntune.csv"
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
    $sharedCheck = Test-SemtSharedInventoryUsable -Path $OutputPath -Kind Intune -MaxAgeHours $IntuneInventoryFreshnessHours -TenantId $TenantId
    if ($sharedCheck.Usable) {
        $sharedInventoryItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
        Write-Host "Smart Endpoint Migration Toolkit - Export-IntuneDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
        Write-Host ("Recent delegated shared DevicesIntune.csv found: {0}" -f $OutputPath) -ForegroundColor Green
        Write-Host ("Last write time: {0}; Age: {1:N1} hour(s); {2}" -f $sharedInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $sharedCheck.AgeHours, $sharedCheck.Provenance) -ForegroundColor Green
        Write-Host "No new Graph export generated. Use -ForceRefresh to regenerate anyway." -ForegroundColor Yellow
        return
    }
    if ($sharedCheck.Reason -eq 'ProvenanceRejected') {
        Write-Host ("Recent shared DevicesIntune.csv ignored because delegated provenance is missing or mismatched: {0}" -f $sharedCheck.Provenance) -ForegroundColor Yellow
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
    if ($module) {
        return $module
    }

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
    if (-not $module) {
        throw ('SEMT_GRAPH_MODULE_NOT_FOUND: Microsoft.Graph.Authentication was installed for the current user but cannot be found. {0}' -f (Get-SemtGraphModuleDiagnostic))
    }

    return $module
}

function Import-GraphAuthenticationModule {
    $module = Install-GraphModuleIfMissing -SkipInstall:$SkipModuleInstall

    Import-Module -Name $module.Path -ErrorAction Stop
    Write-Host ("Graph module : Microsoft.Graph.Authentication {0}" -f $module.Version) -ForegroundColor DarkCyan
}

function Get-GraphConnectionContext {
    try {
        return Get-MgContext -ErrorAction Stop
    }
    catch {
        return $null
    }
}


function ConvertTo-ComputerName {
    param([string]$DeviceName)

    if ([string]::IsNullOrWhiteSpace($DeviceName)) {
        return ""
    }

    return ($DeviceName.Split(".")[0]).Trim().ToUpperInvariant()
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

function Get-GraphProperty {
    param(
        [Parameter(Mandatory=$true)]$InputObject,
        [Parameter(Mandatory=$true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) {
        return $property.Value
    }

    return $null
}

function ConvertTo-ODataStringLiteral {
    param([Parameter(Mandatory=$true)][string]$Value)

    return ("'{0}'" -f ($Value -replace "'", "''"))
}

function New-ManagedDeviceExportRow {
    param(
        $Device,
        [string]$RequestedComputerName,
        [string]$ComputerName,
        [bool]$IntuneInventoryPresent
    )

    if ($Device) {
        $deviceName = [string](Get-GraphProperty -InputObject $Device -Name "deviceName")
        if ([string]::IsNullOrWhiteSpace($ComputerName)) {
            $ComputerName = ConvertTo-ComputerName -DeviceName $deviceName
        }

        return [PSCustomObject]@{
            RequestedComputerName  = $RequestedComputerName
            ComputerName           = $ComputerName
            IntuneInventoryPresent = $IntuneInventoryPresent
            DeviceName             = $deviceName
            ManagedDeviceName      = Get-GraphProperty -InputObject $Device -Name "managedDeviceName"
            IntuneManagedDeviceId  = Get-GraphProperty -InputObject $Device -Name "id"
            AzureADDeviceId        = Get-GraphProperty -InputObject $Device -Name "azureADDeviceId"
            AzureADRegistered      = Get-GraphProperty -InputObject $Device -Name "azureADRegistered"
            OperatingSystem        = Get-GraphProperty -InputObject $Device -Name "operatingSystem"
            OSVersion              = Get-GraphProperty -InputObject $Device -Name "osVersion"
            ManagementAgent        = Get-GraphProperty -InputObject $Device -Name "managementAgent"
            ManagementState        = Get-GraphProperty -InputObject $Device -Name "managementState"
            ComplianceState        = Get-GraphProperty -InputObject $Device -Name "complianceState"
            ManagedDeviceOwnerType = Get-GraphProperty -InputObject $Device -Name "managedDeviceOwnerType"
            EnrolledDateTime       = Get-GraphProperty -InputObject $Device -Name "enrolledDateTime"
            LastSyncDateTime       = Get-GraphProperty -InputObject $Device -Name "lastSyncDateTime"
            UserPrincipalName      = Get-GraphProperty -InputObject $Device -Name "userPrincipalName"
            EmailAddress           = Get-GraphProperty -InputObject $Device -Name "emailAddress"
            SerialNumber           = Get-GraphProperty -InputObject $Device -Name "serialNumber"
            Manufacturer           = Get-GraphProperty -InputObject $Device -Name "manufacturer"
            Model                  = Get-GraphProperty -InputObject $Device -Name "model"
            IsEncrypted            = Get-GraphProperty -InputObject $Device -Name "isEncrypted"
            IsSupervised           = Get-GraphProperty -InputObject $Device -Name "isSupervised"
        }
    }

    return [PSCustomObject]@{
        RequestedComputerName  = $RequestedComputerName
        ComputerName           = $ComputerName
        IntuneInventoryPresent = $false
        DeviceName             = ""
        ManagedDeviceName      = ""
        IntuneManagedDeviceId  = ""
        AzureADDeviceId        = ""
        AzureADRegistered      = ""
        OperatingSystem        = ""
        OSVersion              = ""
        ManagementAgent        = ""
        ManagementState        = ""
        ComplianceState        = ""
        ManagedDeviceOwnerType = ""
        EnrolledDateTime       = ""
        LastSyncDateTime       = ""
        UserPrincipalName      = ""
        EmailAddress           = ""
        SerialNumber           = ""
        Manufacturer           = ""
        Model                  = ""
        IsEncrypted            = ""
        IsSupervised           = ""
    }
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

Import-GraphAuthenticationModule

if (-not $NoConnect) {
    $connectParams = @{
        Scopes    = @('DeviceManagementManagedDevices.Read.All')
        NoWelcome = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParams.TenantId = $TenantId.Trim() }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph @connectParams | Out-Null
}

$context = Get-GraphConnectionContext
if (-not $context) {
    throw "Not connected to Microsoft Graph. Run without -NoConnect for delegated interactive authentication, or provide a delegated preconnected session with -NoConnect."
}
$contextAuthType = if ($context.PSObject.Properties['AuthType']) { [string]$context.AuthType } else { '' }
if ($contextAuthType -ieq 'AppOnly') {
    throw 'App-only Microsoft Graph authentication is not supported by this exporter. Sign in with a delegated interactive account.'
}
$inventoryAuthenticationMode = if ($NoConnect) { 'DelegatedExistingSession' } else { 'DelegatedInteractive' }

Write-Host "Smart Endpoint Migration Toolkit - Export-IntuneDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
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
    Write-Host "Computers   : no -ComputerListPath; exporting all Intune managed devices." -ForegroundColor Yellow
}

$select = @(
    "id",
    "deviceName",
    "azureADDeviceId",
    "azureADRegistered",
    "managedDeviceName",
    "operatingSystem",
    "osVersion",
    "managementAgent",
    "managementState",
    "complianceState",
    "managedDeviceOwnerType",
    "enrolledDateTime",
    "lastSyncDateTime",
    "userPrincipalName",
    "emailAddress",
    "serialNumber",
    "manufacturer",
    "model",
    "isEncrypted",
    "isSupervised"
) -join ","

$devices = @()
if ($requestedComputers.Count -gt 0) {
    foreach ($computer in $requestedComputers) {
        $candidateNames = @($computer.RequestedComputerName,$computer.ComputerName) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Select-Object -Unique

        foreach ($candidateName in $candidateNames) {
            $filter = "deviceName eq {0}" -f (ConvertTo-ODataStringLiteral -Value ([string]$candidateName))
            $encodedFilter = [System.Uri]::EscapeDataString($filter)
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$select&`$filter=$encodedFilter&`$top=$PageSize"
            $devices += @(Invoke-GraphPagedRequest -Uri $uri -Quiet)
        }
    }
    Write-Host ("Graph lookups: {0} computer(s) queried; {1} Intune device(s) found." -f $requestedComputers.Count,@($devices).Count) -ForegroundColor DarkCyan
}
else {
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$select&`$top=$PageSize"
    $devices = Invoke-GraphPagedRequest -Uri $uri
}

if ($requestedComputers.Count -gt 0) {
    $deviceByComputer = @{}
    foreach ($device in $devices) {
        $deviceName = [string](Get-GraphProperty -InputObject $device -Name "deviceName")
        $computerName = ConvertTo-ComputerName -DeviceName $deviceName
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            continue
        }

        if (-not $deviceByComputer.ContainsKey($computerName)) {
            $deviceByComputer[$computerName] = $device
        }
        else {
            $current = $deviceByComputer[$computerName]
            $currentLastSync = [datetime]::MinValue
            $candidateLastSync = [datetime]::MinValue
            [datetime]::TryParse([string](Get-GraphProperty -InputObject $current -Name "lastSyncDateTime"), [ref]$currentLastSync) | Out-Null
            [datetime]::TryParse([string](Get-GraphProperty -InputObject $device -Name "lastSyncDateTime"), [ref]$candidateLastSync) | Out-Null
            if ($candidateLastSync -gt $currentLastSync) {
                $deviceByComputer[$computerName] = $device
            }
        }
    }

    $export = foreach ($computer in $requestedComputers) {
        $found = $deviceByComputer.ContainsKey($computer.ComputerName)
        $device = if ($found) { $deviceByComputer[$computer.ComputerName] } else { $null }
        New-ManagedDeviceExportRow `
            -Device $device `
            -RequestedComputerName $computer.RequestedComputerName `
            -ComputerName $computer.ComputerName `
            -IntuneInventoryPresent $found
    }
}
elseif ($IncludeAllProperties) {
    $export = foreach ($device in $devices) {
        $deviceName = [string](Get-GraphProperty -InputObject $device -Name "deviceName")
        $row = $device | Add-Member -NotePropertyName ComputerName -NotePropertyValue (ConvertTo-ComputerName -DeviceName $deviceName) -Force -PassThru
        $row | Add-Member -NotePropertyName IntuneInventoryPresent -NotePropertyValue $true -Force
        $row
    }
}
else {
    $export = foreach ($device in $devices) {
        $deviceName = [string](Get-GraphProperty -InputObject $device -Name "deviceName")
        New-ManagedDeviceExportRow `
            -Device $device `
            -RequestedComputerName "" `
            -ComputerName (ConvertTo-ComputerName -DeviceName $deviceName) `
            -IntuneInventoryPresent $true
    }
}

$export = @($export)
$inventoryScope = if ($requestedComputers.Count -gt 0) { 'ComputerList' } else { 'AllManagedDevices' }
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
        Sort-Object ComputerName, DeviceName |
        Export-Csv -LiteralPath $tempOutputPath -NoTypeInformation -Encoding UTF8

    Copy-Item -LiteralPath $tempOutputPath -Destination $finalOutputPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $tempOutputPath -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host ""
    Write-Host ("ERROR: Cannot write DevicesIntune.csv to: {0}" -f $OutputPath) -ForegroundColor Red
    Write-Host ("Detail: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($tempOutputPath) -and (Test-Path -LiteralPath $tempOutputPath)) {
        Write-Host ("Temporary CSV kept for troubleshooting: {0}" -f $tempOutputPath) -ForegroundColor Yellow
    }
    Write-Host "Run the CMD as administrator, close the CSV if it is open in Excel, or use -OutputPath with a writable folder." -ForegroundColor Yellow
    exit 1
}

$uniqueComputers = @($export | Where-Object { -not [string]::IsNullOrWhiteSpace($_.ComputerName) } | Select-Object -ExpandProperty ComputerName -Unique)
$presentCount = @($export | Where-Object { $_.IntuneInventoryPresent -eq $true }).Count
$missingCount = @($export | Where-Object { $_.IntuneInventoryPresent -eq $false }).Count

Write-Host ""
Write-Host ("Exported Intune managed devices: {0}" -f @($export).Count) -ForegroundColor Green
Write-Host ("Unique computer names        : {0}" -f $uniqueComputers.Count) -ForegroundColor Green
if ($requestedComputers.Count -gt 0) {
    Write-Host ("Requested computers          : {0}" -f $requestedComputers.Count) -ForegroundColor Green
    Write-Host ("Present in Intune            : {0}" -f $presentCount) -ForegroundColor Green
    Write-Host ("Missing from Intune          : {0}" -f $missingCount) -ForegroundColor Green
}
Write-Host "CSV                          : $OutputPath" -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAZMRSIyt+mKVo0
# bPF69+yiR+/d7cJNzWvAoqTqC5F7FKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFRVlX+4tLWeSsvUZBNCrS7o40cSn79IIW71Xfb9r6m9MA0GCSqG
# SIb3DQEBAQUABIIBgGow94HGvPD3aE90zoBCMbhEWrmzz9pl9674Ca3vxC9T+xue
# KbksTGEKvQHsPJey7gl5IzOdzLcBxzzK4En88GN0NFbUhb4Ht13ZNf0aAx7PjVp7
# 8eC+MtAXf0ipZJ4gFSQjBR8sIWjzlhFt2oHfBp+FzP0X9hjYygJPzQNwnts8xrsR
# fcxsboNAph38QyHKJbwBYpM5hhVI23XNp9pTS9ExfrpkR3RzgD1lanuAA/hBtaUv
# hyFND2V0bJ39xiKypcBppE4a70hKZ/+ElMdtmH5RrYPsVr1HpZHNbcC89I+u+bmE
# S54OPN72xO4xMSifPPbxejATdD/8MbyCbv3i2H0o+zSi9wrJyw85UlII0ukefvOc
# 7NIZxtDCet/59WTfshoYAY4YV7OEU3Ve8lGW3KradrJvHu/XByYCgPjKvHH5zEqI
# UrllWCOe76U7+9tK1ycDItnCB4LPsUvqUBBF4oM8FVR5aDfM64wkjAAXbJzSA/R6
# pSRugADg6RdPVstYjKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNjM4
# MjhaMC8GCSqGSIb3DQEJBDEiBCAQkXQ7FHcWc5iws5IMliKuVowFuJHk2saJP7ns
# 3oD+oTANBgkqhkiG9w0BAQEFAASCAgAar4jAWkH4pEhNf6tD91mP282S9RZp8PvG
# NT+ZIytq9H9ALg6ZJ3/0Nkjc9cXkoMWZPOofHmDdODy4lw9f1EAbzhqNrxX6U0cX
# 4LDpPjIZ7uT14Sscf0KrbctxCmimLvpsrZVSb6OZGbrZYQ4+kWZH4RBlM5NHU51U
# 0a1P2XEPdmt2X9Q56fkeU0i6SitH1CUUaxPckcOJk96JVzdd+OdyDvBDSsk4KR92
# +WQQ+L2UJO3VylhirDDFwbm8sv9NqfEroZJ0BdIAKI9C0bIFpXrAlaC7LAzZDnTb
# i8DfZCOnc91JwzUihV2XCfmOiXvm/m0ZRFkQ6QbcB1WEJ72CogFQgaxTxoV3rRV+
# 7PTl1DjfTY8ONdR65jZaQpX94d3VNVmGP4pd3rUWX6HplgRssFH7Wpc2YrJGFMZv
# ZlxAajZg8YjqX59FvZNsdQfgknsWgHqkutLgy1fIhAhaWnfsK44ThZjXmTJ7z5UR
# Y36dAdoKQMzmayTTKMIa0/xR75Vh3sppWh13ItytPBqVX0LZOmAFBOwvghSYRsBE
# OHBdwVykP9Q31kIaxiN4e1ouiGRcABGwRsKniiJyLxwD6kcsmtqISorNzV6q/dup
# PhmtQBgOSybGsnSj0Oz04fIy4FFa2OrZLUQfwv8KB94ZF0h8ZnpIhDCY39PCj+r3
# /Ykbq1p8rQ==
# SIG # End signature block
