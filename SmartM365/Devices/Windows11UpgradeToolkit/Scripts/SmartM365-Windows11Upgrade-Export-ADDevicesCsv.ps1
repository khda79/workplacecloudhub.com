<#
.SYNOPSIS
Exports Active Directory computer objects to DevicesAD.csv for the Windows 11 upgrade launcher.

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
1.0.0
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
$ScriptVersion = "1.0.0"
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
        Write-Host "Export-Windows11UpgradeADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
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
                Write-Host "Export-Windows11UpgradeADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
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

Write-Host "Export-Windows11UpgradeADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALEVxPsFFoKugQ
# o/PNpwDGKOit9cAZ45x25SL6VP1M4KCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBxpV6InwM9FU60n7VB
# K1z+19sgSMS+4/u9Hp3HWNfXmjANBgkqhkiG9w0BAQEFAASCAYAE2AePDs6znAXD
# a2yAvijk68bT6IGE2eJtCRenoH/JogJKq5Z49yrYMorYufDvG+veL/E33bWAkiOI
# CW+RsNcdh7mTaZz88wMi9h5Vh39tBkfweyp77zi6c6CpTF1QicHc5m8LyOgmnWP/
# NiSOchYqXZOwtPINNPieYk/siEaDYt8nFD+YlwGsVvaNiNSgMgJQXuIompHWWJgl
# ev/GitF1tKpSAYerICUEV0vOhwOhBFHok3HTSTEbgWIQuHTu04fKdtxq2aUv5G78
# /Pi+T10QjNeGToqc7lwLMq+3/sQ8KYaraPLy/P3BQ46M1IZLcPj67RsCNh79ZrWv
# ekjyoczi5NQLXb+LjmaqxzT3jl51u7hr1hHo3rVelQEE4Lg4+4b6Kkbt6AKpG7ot
# DlAFZFf2GqX/HYCodSayPP+eQzJM2Ij7esGRXc9wOITnAawHtcONG264YrNawBJg
# LHrccXjihf7SkUOGCLul0qsUrmsA+kPPUPLTLtj2oc7C21iI/N4=
# SIG # End signature block
