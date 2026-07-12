<#
.SYNOPSIS
Exports Intune managed devices to DevicesIntune.csv for the repair launcher.

.DESCRIPTION
Uses Microsoft Graph to read Intune managed devices from:
/deviceManagement/managedDevices

When this script is stored in a Scripts folder, the default output is DevicesIntune.csv in the parent folder.
Otherwise, the default output is DevicesIntune.csv next to this script.
The CSV includes DeviceName and ComputerName columns so SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1
can use it with -IntuneInventoryCsv.

.PARAMETER OutputPath
Destination CSV path. Defaults to DevicesIntune.csv in the parent folder when running from Scripts, otherwise next to this script.

.PARAMETER ComputerListPath
Computers.txt path. Defaults to Computers.txt next to this script when present. If the default file is missing, all Intune managed devices are exported.

.PARAMETER PageSize
Graph page size for the full managedDevices read. Defaults to 999.

.PARAMETER TenantId
Optional tenant id passed to Connect-MgGraph.

.PARAMETER NoConnect
Do not call Connect-MgGraph. Use this only when the current PowerShell session is already connected.

.PARAMETER IncludeAllProperties
Exports the full managedDevice objects returned by Graph instead of the curated default columns.

.PARAMETER SkipModuleInstall
Do not install Microsoft.Graph.Authentication automatically if it is missing.

.PARAMETER ForceRefresh
Regenerates the CSV even when a recent DevicesIntune.csv exists in the parent folder.

.EXAMPLE
.\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1

.EXAMPLE
.\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 -ComputerListPath .\Computers.txt -PageSize 999

.EXAMPLE
.\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 -OutputPath .\DevicesIntune.csv -TenantId "contoso.onmicrosoft.com"

.EXAMPLE
.\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1 -IntuneInventoryCsv .\DevicesIntune.csv
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

$ScriptVersion = "1.3.7"

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$DefaultOutputDir = $BaseDir
if ((Split-Path -Leaf $BaseDir) -ieq "Scripts") {
    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $DefaultOutputDir = $parentDir
    }
}

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
if (-not $OutputPathWasProvided) {
    $OutputPath = Join-Path $DefaultOutputDir "DevicesIntune.csv"
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
        $parentInventoryPath = Join-Path $parentDir "DevicesIntune.csv"
        $defaultOutputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
        $parentInventoryFullPath = [System.IO.Path]::GetFullPath($parentInventoryPath)

        if ($parentInventoryFullPath -ne $defaultOutputFullPath -and (Test-Path -LiteralPath $parentInventoryPath)) {
            $parentInventoryItem = Get-Item -LiteralPath $parentInventoryPath -ErrorAction Stop
            $parentInventoryAge = (Get-Date) - $parentInventoryItem.LastWriteTime
            if ($parentInventoryAge.TotalHours -le 2) {
                Write-Host "Export-IntuneDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
                Write-Host ("Recent parent DevicesIntune.csv found: {0}" -f $parentInventoryPath) -ForegroundColor Green
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
    if ($module) {
        return $module
    }

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
    if (-not $module) {
        throw "Microsoft.Graph.Authentication installation completed but the module was not found."
    }

    return $module
}

function Import-GraphAuthenticationModule {
    $module = Install-GraphModuleIfMissing -SkipInstall:$SkipModuleInstall

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
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
    param([Parameter(Mandatory=$true)][string]$Uri)

    $items = @()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        Write-Host "Reading Graph page: $nextUri" -ForegroundColor DarkCyan
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop

        $pageItems = Get-GraphProperty -InputObject $response -Name "value"
        if ($pageItems) {
            foreach ($item in @($pageItems)) {
                if ($null -ne $item) {
                    $items += $item
                }
            }
        }

        $nextUri = ""
        $nextLink = Get-GraphProperty -InputObject $response -Name "@odata.nextLink"
        if (-not [string]::IsNullOrWhiteSpace([string]$nextLink)) {
            $nextUri = [string]$nextLink
        }
    }

    return @($items)
}

Import-GraphAuthenticationModule

if (-not $NoConnect) {
    $connectParams = @{
        Scopes = @("DeviceManagementManagedDevices.Read.All")
        NoWelcome = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $connectParams.TenantId = $TenantId
    }

    Connect-MgGraph @connectParams | Out-Null
}

$context = Get-GraphConnectionContext
if (-not $context) {
    throw "Not connected to Microsoft Graph. Run without -NoConnect, or connect first with Connect-MgGraph -Scopes DeviceManagementManagedDevices.Read.All."
}

Write-Host "Export-IntuneDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
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
    Write-Host "Computers   : no Computers.txt found next to this script; exporting all Intune managed devices." -ForegroundColor Yellow
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
            $devices += @(Invoke-GraphPagedRequest -Uri $uri)
        }
    }
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD324sJx0UFUPzT
# 3y6rcJv0eEhSJTqQ9JR7UwE08Ch+QKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCATOXAV3ltQPCFs+1em
# VjThHMmaBCPXEUu9DcRGup3+sTANBgkqhkiG9w0BAQEFAASCAYAWtxx8PKB2aXhb
# pKiCg4oBHwc5xZPvij5E16K6HMPbnIzd5UykmkdeLFyaaG76PrsBupVP6PCp3euP
# OyP/qDfV2Ew0jP7oH3cnmwspF8wDqAvulFBt8ZLL3zOEESWAks1VyAYM50XG12eC
# lWmpPVdD2V0RqEaRdrbVw+VT09qCygTzPxWsk+ie54xH83XyBa0rqFE8DXdeYPJL
# FUqvkFEUm/Enlbid/9jbHz86FOJ6Q7LOlGYloU1bIRKtNPMglytT/G44Gg8L17jf
# xlVQ4NKaCa6ojjb/VeAvvx0sxXyXGh1JhefNvxWOItoinrfJYw/ba2LlyHM8U5KB
# kmL30j8rehyf6zXivvOcUofeWRfRWWoQHSdDaLAk7ftlMY8tBdjDBakQZ34DOHxc
# YhrNaByuhnp2Hd6Qnpse0WLyeN1RuZ++N1Y1qwvA+J9aGX3izAsE+fX6x07V1wuh
# PsyBBVqVJSMpzjrzi+D0c6W190rz+Z86F6mTsKC3HES7D9oJMV8=
# SIG # End signature block
