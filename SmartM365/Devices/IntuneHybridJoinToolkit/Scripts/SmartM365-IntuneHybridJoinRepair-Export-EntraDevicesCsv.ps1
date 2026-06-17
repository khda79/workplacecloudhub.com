<#
.SYNOPSIS
Exports Microsoft Entra devices to DevicesEntra.csv for the repair launcher.

.DESCRIPTION
Uses Microsoft Graph to read Entra devices from /devices.
When this script is stored in a Scripts folder, the default output is DevicesEntra.csv in the parent folder.

.PARAMETER OutputPath
Destination CSV path. Defaults to DevicesEntra.csv in the parent folder when running from Scripts.

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
Regenerates the CSV even when a recent DevicesEntra.csv exists in the parent folder.
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
$ScriptVersion = "1.0.1"

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$DefaultOutputDir = $BaseDir
if ((Split-Path -Leaf $BaseDir) -ieq "Scripts") {
    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) { $DefaultOutputDir = $parentDir }
}

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
if (-not $OutputPathWasProvided) {
    $OutputPath = Join-Path $DefaultOutputDir "DevicesEntra.csv"
}

$ComputerListPathWasProvided = -not [string]::IsNullOrWhiteSpace($ComputerListPath)
if ([string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $ComputerListPath = Join-Path $BaseDir "Computers.txt"
    if (-not (Test-Path -LiteralPath $ComputerListPath)) {
        $ComputerListPath = ""
    }
}

if ($PageSize -lt 1) { $PageSize = 1 }
if ($PageSize -gt 999) { $PageSize = 999 }

if (-not $OutputPathWasProvided -and -not $ForceRefresh) {
    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $parentInventoryPath = Join-Path $parentDir "DevicesEntra.csv"
        $defaultOutputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
        $parentInventoryFullPath = [System.IO.Path]::GetFullPath($parentInventoryPath)
        if ($parentInventoryFullPath -ne $defaultOutputFullPath -and (Test-Path -LiteralPath $parentInventoryPath)) {
            $parentInventoryItem = Get-Item -LiteralPath $parentInventoryPath -ErrorAction Stop
            $parentInventoryAge = (Get-Date) - $parentInventoryItem.LastWriteTime
            if ($parentInventoryAge.TotalHours -le 2) {
                Write-Host "Export-EntraDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
                Write-Host ("Recent parent DevicesEntra.csv found: {0}" -f $parentInventoryPath) -ForegroundColor Green
                Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $parentInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $parentInventoryAge.TotalHours) -ForegroundColor Green
                Write-Host "No new Graph export generated. Use -ForceRefresh to regenerate anyway." -ForegroundColor Yellow
                return
            }
        }
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

    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
    if ($module) { return $module }

    if ($SkipInstall) {
        throw "Microsoft.Graph.Authentication module is not installed. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
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
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) { throw "Microsoft.Graph.Authentication installation completed but the module was not found." }
    return $module
}

function Import-GraphAuthenticationModule {
    $module = Install-GraphModuleIfMissing -SkipInstall:$SkipModuleInstall
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
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
    param([Parameter(Mandatory=$true)][string]$Uri)

    $items = @()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        Write-Host "Reading Graph page: $nextUri" -ForegroundColor DarkCyan
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
        $pageItems = Get-GraphProperty -InputObject $response -Name "value"
        if ($pageItems) {
            foreach ($item in @($pageItems)) {
                if ($null -ne $item) { $items += $item }
            }
        }
        $nextUri = ""
        $nextLink = Get-GraphProperty -InputObject $response -Name "@odata.nextLink"
        if (-not [string]::IsNullOrWhiteSpace([string]$nextLink)) { $nextUri = [string]$nextLink }
    }
    return @($items)
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
    Connect-MgGraph @connectParams | Out-Null
}

$context = Get-GraphConnectionContext
if (-not $context) {
    throw "Not connected to Microsoft Graph. Run without -NoConnect, or connect first with Connect-MgGraph -Scopes Device.Read.All."
}

Write-Host "Export-EntraDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
Write-Host "Tenant      : $($context.TenantId)"
Write-Host "Account     : $($context.Account)"
Write-Host "Output      : $OutputPath"
Write-Host "Page size   : $PageSize"

$requestedComputers = @()
if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $requestedComputers = @(Get-ComputerList -Path $ComputerListPath)
    Write-Host "Computers   : $ComputerListPath"
    Write-Host ("Requested   : {0}" -f $requestedComputers.Count)
}
elseif (-not $ComputerListPathWasProvided) {
    Write-Host "Computers   : no Computers.txt found next to this script; exporting all Entra devices." -ForegroundColor Yellow
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
            $devices += @(Invoke-GraphPagedRequest -Uri $uri)
        }
    }

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
    exit 1
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
