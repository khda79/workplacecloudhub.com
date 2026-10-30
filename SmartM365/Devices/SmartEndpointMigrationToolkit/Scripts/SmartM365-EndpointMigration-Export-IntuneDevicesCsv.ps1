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
$IntuneInventoryFreshnessHours = 24

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
    # Keep every distinct managed-device object. A name can legitimately be reused;
    # dropping older rows would turn an ambiguous identity into a false confirmation.
    $deviceByComputer = @{}
    $seenIds = @{}
    foreach ($device in $devices) {
        $computerName = ConvertTo-ComputerName -DeviceName ([string](Get-GraphProperty -InputObject $device -Name "deviceName"))
        $deviceId = [string](Get-GraphProperty -InputObject $device -Name "id")
        if ([string]::IsNullOrWhiteSpace($computerName) -or [string]::IsNullOrWhiteSpace($deviceId)) { throw 'Intune scoped response has an object without a name or id.' }
        if ($seenIds.ContainsKey($deviceId)) { continue }
        $seenIds[$deviceId] = $true
        if (-not $deviceByComputer.ContainsKey($computerName)) { $deviceByComputer[$computerName] = New-Object System.Collections.Generic.List[object] }
        $deviceByComputer[$computerName].Add($device)
    }
    $export = foreach ($computer in $requestedComputers) {
        if ($deviceByComputer.ContainsKey($computer.ComputerName)) {
            foreach ($device in $deviceByComputer[$computer.ComputerName]) {
                New-ManagedDeviceExportRow -Device $device -RequestedComputerName $computer.RequestedComputerName -ComputerName $computer.ComputerName -IntuneInventoryPresent $true
            }
        } else {
            New-ManagedDeviceExportRow -Device $null -RequestedComputerName $computer.RequestedComputerName -ComputerName $computer.ComputerName -IntuneInventoryPresent $false
        }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDr4MiISEitaKXY
# fRYOY3cTZMKWQ+wBqJCKLatIYxAHoqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCxWyvKMm98rcEicpH+f7x1
# an2ga6NZooWkG9Xm/WimiDANBgkqhkiG9w0BAQEFAASCAYBo3gw8+FzDnt2oO7+G
# G1O1vMqUR8iGyyauCVNJvjFB5WNK2W55DqWPaIXtEq6K6lVDLMUXw8pkOdCzxG9f
# pQzvdF1/Rh7HGI+3DYMVfOycFH/pTB11awEoSt9JU8BPfyDh6CcYUvrhf+myv70B
# QmxEirqc9g0I86cjkGF9WjdDBz6D7seFuha3wxGnVXrXNAjcgiLbkemgk4ST/yM6
# n+nZPYiNbwa8a4PVo3Pbq8le6HoEOe7gBZ/pxI8DIADYXxm3bWlFDizpoKUblcOR
# PQ32DOmex3b446zKMCl3LChtKxcu6y2WhfKWZBj6yGhXBLVcNUMECaYLCtGcOGJ9
# z0MbDTtVtdihI1Rn7WdPB0elB/3VamByakpRymigCtjAieGA+kQ8aF+vZPWPPiIr
# LKKREfR9hLyxcComwzk+BLefZ9iLaYAECTvGi7wbpad1TJuz24OhdTS3X4qVRJlj
# nI0G0vxl1Rxf4I6+03SKzTZwWGmuVMHeyZP8W2ywCiUTX4M=
# SIG # End signature block
