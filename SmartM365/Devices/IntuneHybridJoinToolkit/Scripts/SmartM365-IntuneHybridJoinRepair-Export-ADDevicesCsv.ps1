<#
.SYNOPSIS
Exports Active Directory computer objects to DevicesAD.csv for the repair launcher.

.DESCRIPTION
Uses the ActiveDirectory PowerShell module to read computer objects.
Without -Domain, the script exports all domains in the current AD forest.
With -Domain, the script exports only that domain.
When this script is stored in a Scripts folder, the default output is DevicesAD.csv in the parent folder.

.PARAMETER OutputPath
Destination CSV path. Defaults to DevicesAD.csv in the parent folder when running from Scripts.

.PARAMETER ComputerListPath
Computers.txt path. When provided, only those AD computers are queried.

.PARAMETER Domain
Optional AD domain/controller passed to Get-ADComputer -Server. Use this for per-LOT domain selection.
When omitted, all domains in the current AD forest are exported.

.PARAMETER ForceRefresh
Regenerates the CSV even when a recent DevicesAD.csv exists in the parent folder.

.VERSION
1.0.2
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$ComputerListPath,
    [string]$Domain,
    [switch]$ForceRefresh
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "1.0.2"
$AdInventoryFreshnessHours = 12

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$DefaultOutputDir = $BaseDir
if ((Split-Path -Leaf $BaseDir) -ieq "Scripts") {
    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) { $DefaultOutputDir = $parentDir }
}

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
if (-not $OutputPathWasProvided) {
    $OutputPath = Join-Path $DefaultOutputDir "DevicesAD.csv"
}

if (-not $ForceRefresh -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $outputInventoryItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    $outputInventoryAge = (Get-Date) - $outputInventoryItem.LastWriteTime
    if ($outputInventoryAge.TotalHours -le $AdInventoryFreshnessHours) {
        Write-Host "Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
        Write-Host ("Recent DevicesAD.csv found: {0}" -f $OutputPath) -ForegroundColor Green
        Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $outputInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $outputInventoryAge.TotalHours) -ForegroundColor Green
        Write-Host ("No new AD export generated because the AD CSV is less than {0} hour(s) old. Use -ForceRefresh to regenerate anyway." -f $AdInventoryFreshnessHours) -ForegroundColor Yellow
        return
    }
}
$ComputerListPathWasProvided = -not [string]::IsNullOrWhiteSpace($ComputerListPath)
if ([string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $ComputerListPath = Join-Path $BaseDir "Computers.txt"
    if (-not (Test-Path -LiteralPath $ComputerListPath)) {
        $ComputerListPath = ""
    }
}

if (-not $OutputPathWasProvided -and -not $ForceRefresh) {
    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $parentInventoryPath = Join-Path $parentDir "DevicesAD.csv"
        $defaultOutputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
        $parentInventoryFullPath = [System.IO.Path]::GetFullPath($parentInventoryPath)
        if ($parentInventoryFullPath -ne $defaultOutputFullPath -and (Test-Path -LiteralPath $parentInventoryPath)) {
            $parentInventoryItem = Get-Item -LiteralPath $parentInventoryPath -ErrorAction Stop
            $parentInventoryAge = (Get-Date) - $parentInventoryItem.LastWriteTime
            if ($parentInventoryAge.TotalHours -le $AdInventoryFreshnessHours) {
                Write-Host "Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
                Write-Host ("Recent parent DevicesAD.csv found: {0}" -f $parentInventoryPath) -ForegroundColor Green
                Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $parentInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $parentInventoryAge.TotalHours) -ForegroundColor Green
                Write-Host ("No new AD export generated because the AD CSV is less than {0} hour(s) old. Use -ForceRefresh to regenerate anyway." -f $AdInventoryFreshnessHours) -ForegroundColor Yellow
                return
            }
        }
    }
}

function Convert-FileTimeUtc {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return "" }
    try {
        $number = [int64]$Value
        if ($number -le 0) { return "" }
        return [DateTime]::FromFileTimeUtc($number).ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return ""
    }
}

function Convert-ToComputerName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return ($Name.Split(".")[0]).Trim().ToUpperInvariant()
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
                ComputerName = Convert-ToComputerName -Name $_
            }
        } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.ComputerName) } |
        Sort-Object ComputerName -Unique)
}

function ConvertTo-AdFilterStringLiteral {
    param([Parameter(Mandatory=$true)][string]$Value)

    return ("'{0}'" -f ($Value -replace "'", "''"))
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    throw "ActiveDirectory PowerShell module is not available. Install RSAT Active Directory tools on the operator workstation. Detail: $($_.Exception.Message)"
}

$domainTargets = @()
if (-not [string]::IsNullOrWhiteSpace($Domain)) {
    try {
        $adDomain = Get-ADDomain -Server $Domain.Trim() -ErrorAction Stop
        $effectiveDomain = if ($adDomain -and -not [string]::IsNullOrWhiteSpace([string]$adDomain.DNSRoot)) { [string]$adDomain.DNSRoot } else { $Domain.Trim() }
        $domainTargets += [PSCustomObject]@{
            Server = $Domain.Trim()
            DNSRoot = $effectiveDomain
        }
    }
    catch {
        Write-Host ("WARN: Could not resolve AD domain metadata for '{0}': {1}" -f $Domain,$_.Exception.Message) -ForegroundColor Yellow
        $domainTargets += [PSCustomObject]@{
            Server = $Domain.Trim()
            DNSRoot = $Domain.Trim()
        }
    }
}
else {
    $forest = Get-ADForest -ErrorAction Stop
    foreach ($forestDomain in @($forest.Domains)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$forestDomain)) {
            $domainTargets += [PSCustomObject]@{
                Server = [string]$forestDomain
                DNSRoot = [string]$forestDomain
            }
        }
    }
}

if ($domainTargets.Count -eq 0) {
    throw "No AD domain target could be resolved for export."
}

$commonProperties = @(
    "DNSHostName",
    "Enabled",
    "OperatingSystem",
    "OperatingSystemVersion",
    "LastLogonTimestamp",
    "DistinguishedName"
)

Write-Host "Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
Write-Host ("Domains     : {0}" -f (($domainTargets | Select-Object -ExpandProperty DNSRoot) -join "; "))
Write-Host "Output      : $OutputPath"

$requestedComputers = @()
if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $requestedComputers = @(Get-ComputerList -Path $ComputerListPath)
    Write-Host "Computers   : $ComputerListPath"
    Write-Host ("Requested   : {0}" -f $requestedComputers.Count)
}
elseif (-not $ComputerListPathWasProvided) {
    Write-Host "Computers   : no Computers.txt found next to this script; exporting all AD computers." -ForegroundColor Yellow
}

function New-AdExportRow {
    param(
        [Parameter(Mandatory=$true)]$Computer,
        [Parameter(Mandatory=$true)]$DomainTarget
    )

    [PSCustomObject]@{
        ComputerName          = Convert-ToComputerName -Name $Computer.Name
        Name                  = $Computer.Name
        ADInventoryPresent    = $true
        ADDomain              = $DomainTarget.DNSRoot
        Enabled               = $Computer.Enabled
        DNSHostName           = $Computer.DNSHostName
        DistinguishedName     = $Computer.DistinguishedName
        OperatingSystem       = $Computer.OperatingSystem
        OperatingSystemVersion = $Computer.OperatingSystemVersion
        LastLogonTimestampUtc = Convert-FileTimeUtc -Value $Computer.LastLogonTimestamp
    }
}

function New-MissingAdExportRow {
    param([Parameter(Mandatory=$true)]$RequestedComputer)

    [PSCustomObject]@{
        ComputerName          = $RequestedComputer.ComputerName
        Name                  = $RequestedComputer.RequestedComputerName
        ADInventoryPresent    = $false
        ADDomain              = ""
        Enabled               = ""
        DNSHostName           = ""
        DistinguishedName     = ""
        OperatingSystem       = ""
        OperatingSystemVersion = ""
        LastLogonTimestampUtc = ""
    }
}

$export = if ($requestedComputers.Count -gt 0) {
    $foundByComputer = @{}
    foreach ($domainTarget in $domainTargets) {
        Write-Host ("Reading selected AD computers from domain: {0}" -f $domainTarget.DNSRoot) -ForegroundColor DarkCyan
        foreach ($requestedComputer in $requestedComputers) {
            $nameLiteral = ConvertTo-AdFilterStringLiteral -Value $requestedComputer.ComputerName
            $requestedLiteral = ConvertTo-AdFilterStringLiteral -Value $requestedComputer.RequestedComputerName
            $queryParams = @{
                Filter = ("Name -eq {0} -or DNSHostName -eq {1}" -f $nameLiteral,$requestedLiteral)
                Server = $domainTarget.Server
                Properties = $commonProperties
            }

            foreach ($computer in @(Get-ADComputer @queryParams)) {
                $key = Convert-ToComputerName -Name $computer.Name
                if (-not [string]::IsNullOrWhiteSpace($key) -and -not $foundByComputer.ContainsKey($key)) {
                    $foundByComputer[$key] = (New-AdExportRow -Computer $computer -DomainTarget $domainTarget)
                }
            }
        }
    }

    foreach ($requestedComputer in $requestedComputers) {
        if ($foundByComputer.ContainsKey($requestedComputer.ComputerName)) {
            $foundByComputer[$requestedComputer.ComputerName]
        }
        else {
            New-MissingAdExportRow -RequestedComputer $requestedComputer
        }
    }
}
else {
    foreach ($domainTarget in $domainTargets) {
        Write-Host ("Reading AD computers from domain: {0}" -f $domainTarget.DNSRoot) -ForegroundColor DarkCyan
        $queryParams = @{
            Filter = "*"
            Server = $domainTarget.Server
            Properties = $commonProperties
        }

        foreach ($computer in @(Get-ADComputer @queryParams)) {
            New-AdExportRow -Computer $computer -DomainTarget $domainTarget
        }
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
        Sort-Object ADDomain, ComputerName, DNSHostName |
        Export-Csv -LiteralPath $tempOutputPath -NoTypeInformation -Encoding UTF8

    Copy-Item -LiteralPath $tempOutputPath -Destination $finalOutputPath -Force -ErrorAction Stop
    Remove-Item -LiteralPath $tempOutputPath -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host ""
    Write-Host ("ERROR: Cannot write DevicesAD.csv to: {0}" -f $OutputPath) -ForegroundColor Red
    Write-Host ("Detail: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($tempOutputPath) -and (Test-Path -LiteralPath $tempOutputPath)) {
        Write-Host ("Temporary CSV kept for troubleshooting: {0}" -f $tempOutputPath) -ForegroundColor Yellow
    }
    exit 1
}

$enabledCount = @($export | Where-Object { $_.Enabled -eq $true }).Count

Write-Host ""
Write-Host ("Exported AD computers: {0}" -f @($export).Count) -ForegroundColor Green
Write-Host ("Enabled computers    : {0}" -f $enabledCount) -ForegroundColor Green
Write-Host ("Domains exported     : {0}" -f $domainTargets.Count) -ForegroundColor Green
if ($requestedComputers.Count -gt 0) {
    $presentCount = @($export | Where-Object { $_.ADInventoryPresent -eq $true }).Count
    Write-Host ("Requested computers  : {0}" -f $requestedComputers.Count) -ForegroundColor Green
    Write-Host ("Present in AD        : {0}" -f $presentCount) -ForegroundColor Green
    Write-Host ("Missing from AD      : {0}" -f ($requestedComputers.Count - $presentCount)) -ForegroundColor Green
}
Write-Host ("CSV                  : {0}" -f $OutputPath) -ForegroundColor Green

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClBLB/8+a8MFvw
# qqYNDierE/6ypNRTfCswDhaK48fhHKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAhBICiSo4/kIjQbybqZKQpUcXwKaaVcL8YQPF47NTYaDANBgkqhkiG9w0B
# AQEFAASCAYBzQxAZ8HJv7JjqmWPgaSCFza5an0iDRBuYn85BTIK8WXWUeXenJRvC
# HTuQHYm+fhYtVptHb+6dVmf9G3bxVzk1HEcWkEFqNjXjZoPGcBf/uJOg1rg6mbPn
# SqoMRFNFgnW1jxRmgGwxm516OQlpDX1+9e6xfgvxwVQsrv9ffLxkOQKi2jnlm0oo
# wmo6txSAsJFSvhBNktKAT4SPOK7h8py5iwaUl5fJOy3RZC+3EV3vpJ5KqCMW8vC0
# 5NK/oMdNls3eJlTycFCerg3CZg5XatO+euD8YTCoZB6hmZ5/HPijIe4UiOFmClgu
# oRmwnrXuoYhI7w8/mAK5uIUe9ud1OkUolZDqssBYxNDuLi29O6XTold7ilCalWzz
# G5Lt7WvLRTU4hzrFkf+oE1EAvMqbPTKCqPH7HVmhmb9crUB7WCCwt0fxd9sw7OYr
# BUBzq7ErY149sIm7mSMXfendzz3Szkcj4moT+Fk+BTscxrDxDtonaAoTOD2N6V46
# 13io8NPKu8Y=
# SIG # End signature block
