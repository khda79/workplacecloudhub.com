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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClBLB/8+a8MFvw
# qqYNDierE/6ypNRTfCswDhaK48fhHKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEICEEgKJKjj+QiNBvJupkpClRxfApppVwvxhA8Xjs1NhoMA0GCSqG
# SIb3DQEBAQUABIIBgFCqsvgGIUgldBiNlN6rB6fabT+41P+JJvXEvOLKmdJ2gy5q
# C95plkl/zYSb52lTFvdcom8zKxPBxJ6ozou7fE8XHpG+/P6zNPFalg/kIUpWw4fO
# SPaK93zWnEsJnDV/T82ZHlOWoGNxtI/UwoSZmm4mPV9lhv13WyqAPgg4HGoZlmsx
# vijvQcg3x/wJuPPlFUOZDdrqrPcypm5QK47YD2NEK8ndSU5ndGXzzg1FV1XvEQJt
# /FJ+HLm/odv+LWZmjVn5+5ISt91ZWMb/f75U7Gz6MTEm8BUJ2gL0tcUWTmfEROmw
# SD7SnCDpSQlfJoRPzpqxxmgd3fZjmYt0ns4QJ4k6B63jl46Xh2fg7gBcDhzqDRt1
# /j7nl6+iSf33enzNVADztS8naHCvKRWPav97kNwcb0DZ58AYOf4zpXlyzuoAASkz
# nxy+8Ck8v7A8Py3EPqLupmz1GhWMIeYL4oH+yCvVZAHjbujcnk6uJ2S4IjTXFI4W
# aQG/4WyI0z+OpDF+v6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MzNaMC8GCSqGSIb3DQEJBDEiBCA5nS4mbcPzM9iu3a6WSPPVN2N4vr7+mYUtMnOj
# QQZ3UDANBgkqhkiG9w0BAQEFAASCAgBEwp5ZcVISV5g5wTmcj7OVHwKTr4ndanN1
# uPknJOBxV7muerWR+57PpisNJyL/JijoVx44SL2MJuNcrizSQADAl0hW2cqhLUm/
# SiFfAP4+NiPtGAXwgUjkdnHrz1JBxasDqZmfWCs9H4+dG3IjsMtC2mXuLC35YDEI
# WIvrfxCOusDuCM09ZbXHXsmZ+AVOzewEvqBrO0sb7sDYNV7/4m3w5V+wSjfUD6kk
# q0BB6+kSsZetHJwkpLF7N6P5uk5CcZk7xCDxbET0auE0kCKU70J6bopw6L+brjqb
# BXo8xfg+5pwsFYzhJjDqLK+AnIhGpwQ+cMkFruS2E+oCz/wvK3OZcRQNP890lWSS
# e9GdiIjvIJqFwdmW3UQR8qySVZfobPUa6IlnmYmJgCBFFIjsKkW3vIqZt9F1apin
# +oNR2Y4kj66rcUQhED+o4EpiDTqnSvtYCeRrfSE48z2Cg72y5jQ2woaaBAHcP7tO
# Z9qvks5ZVRO+5iehdlpApGSiZqZGVJq1WbKeNOv6kGcAdGyWBwwv5dQzyqTlSiDI
# 92XvhDmzstQU5j+ByK22nJvQLEFrbzPF36VBwigGJ0B05xJLbkBmi6p7rubuvIcL
# VXQjVNF8VwVfTLX7c8OKk5lisuCQLgxWS+XuRyFY2dGnRtMeEyjmb+X7+nKQeX74
# F5i6yK+7eQ==
# SIG # End signature block
