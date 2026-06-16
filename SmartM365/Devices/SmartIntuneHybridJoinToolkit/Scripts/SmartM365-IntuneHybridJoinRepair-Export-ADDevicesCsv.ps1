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

.PARAMETER Domain
Optional AD domain/controller passed to Get-ADComputer -Server. Use this for per-LOT domain selection.
When omitted, all domains in the current AD forest are exported.

.PARAMETER ForceRefresh
Regenerates the CSV even when a recent DevicesAD.csv exists in the parent folder.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$Domain,
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
    $OutputPath = Join-Path $DefaultOutputDir "DevicesAD.csv"
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
            if ($parentInventoryAge.TotalHours -le 2) {
                Write-Host "Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
                Write-Host ("Recent parent DevicesAD.csv found: {0}" -f $parentInventoryPath) -ForegroundColor Green
                Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $parentInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $parentInventoryAge.TotalHours) -ForegroundColor Green
                Write-Host "No new AD export generated. Use -ForceRefresh to regenerate anyway." -ForegroundColor Yellow
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

$export = foreach ($domainTarget in $domainTargets) {
    Write-Host ("Reading AD computers from domain: {0}" -f $domainTarget.DNSRoot) -ForegroundColor DarkCyan
    $queryParams = @{
        Filter = "*"
        Server = $domainTarget.Server
        Properties = $commonProperties
    }

    foreach ($computer in @(Get-ADComputer @queryParams)) {
        [PSCustomObject]@{
            ComputerName          = Convert-ToComputerName -Name $computer.Name
            Name                  = $computer.Name
            ADInventoryPresent    = $true
            ADDomain              = $domainTarget.DNSRoot
            Enabled               = $computer.Enabled
            DNSHostName           = $computer.DNSHostName
            DistinguishedName     = $computer.DistinguishedName
            OperatingSystem       = $computer.OperatingSystem
            OperatingSystemVersion = $computer.OperatingSystemVersion
            LastLogonTimestampUtc = Convert-FileTimeUtc -Value $computer.LastLogonTimestamp
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
Write-Host ("CSV                  : {0}" -f $OutputPath) -ForegroundColor Green
