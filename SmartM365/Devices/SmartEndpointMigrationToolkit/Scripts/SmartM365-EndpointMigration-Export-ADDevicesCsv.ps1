<#
.SYNOPSIS
Exports Active Directory computer objects to DevicesAD.csv for the Smart Endpoint Migration Toolkit.

.DESCRIPTION
Uses the ActiveDirectory PowerShell module to read computer objects.
Without -Domain, the script exports all domains in the current AD forest.
With -Domain, the script exports only that domain.
Without -OutputPath, an unscoped export writes the shared cache <Package>\Inventory\DevicesAD.csv
and a -ComputerListPath export writes DevicesAD.csv next to that Computers.txt. A scoped export
never overwrites the shared cache. The columns cover both the HybridJoin and the Windows11Upgrade
automatic Wave engines.
Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1 1.0.0 (see Docs\PROVENANCE.json.txt).

.PARAMETER OutputPath
Destination CSV path. Defaults to the shared cache for unscoped exports, or DevicesAD.csv next to -ComputerListPath.

.PARAMETER ComputerListPath
Computers.txt path. When provided, only those AD computers are queried.

.PARAMETER Domain
Optional AD domain/controller passed to Get-ADComputer -Server. Use this for per-Wave domain selection.
When omitted, all domains in the current AD forest are exported.

.PARAMETER ForceRefresh
Regenerates the shared cache even when it is less than 12 hours old.

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
    [string]$Domain,
    [switch]$ForceRefresh
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "1.0.0"
$AdInventoryFreshnessHours = 12

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
$SharedInventoryPath = [System.IO.Path]::GetFullPath((Get-SemtSharedInventoryPath -PackageRoot $PackageRoot -Kind AD))

$OutputPathWasProvided = -not [string]::IsNullOrWhiteSpace($OutputPath)
$ComputerListPathWasProvided = -not [string]::IsNullOrWhiteSpace($ComputerListPath)
if (-not $OutputPathWasProvided) {
    if ($ComputerListPathWasProvided) {
        $OutputPath = Join-Path (Split-Path -Parent ([System.IO.Path]::GetFullPath($ComputerListPath))) "DevicesAD.csv"
    }
    else {
        $OutputPath = $SharedInventoryPath
    }
}
$OutputIsSharedInventory = ([System.IO.Path]::GetFullPath($OutputPath) -ieq $SharedInventoryPath)
if ($ComputerListPathWasProvided -and $OutputIsSharedInventory) {
    throw "SEMT_SCOPED_INVENTORY_TARGETS_SHARED_CACHE: a -ComputerListPath export cannot overwrite the shared inventory $SharedInventoryPath."
}

if (-not $ForceRefresh -and -not $ComputerListPathWasProvided -and $OutputIsSharedInventory -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $outputInventoryItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    $outputInventoryAge = (Get-Date) - $outputInventoryItem.LastWriteTime
    if ($outputInventoryAge.TotalHours -le $AdInventoryFreshnessHours) {
        Write-Host "Smart Endpoint Migration Toolkit - Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
        Write-Host ("Recent shared DevicesAD.csv found: {0}" -f $OutputPath) -ForegroundColor Green
        Write-Host ("Last write time: {0}; Age: {1:N1} hour(s)" -f $outputInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $outputInventoryAge.TotalHours) -ForegroundColor Green
        Write-Host ("No new AD export generated because the AD CSV is less than {0} hour(s) old. Use -ForceRefresh to regenerate anyway." -f $AdInventoryFreshnessHours) -ForegroundColor Yellow
        return
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

Write-Host "Smart Endpoint Migration Toolkit - Export-ADDevicesCsv version $ScriptVersion" -ForegroundColor Cyan
Write-Host ("Domains     : {0}" -f (($domainTargets | Select-Object -ExpandProperty DNSRoot) -join "; "))
Write-Host "Output      : $OutputPath"

$requestedComputers = @()
if ($ComputerListPathWasProvided) {
    $requestedComputers = @(Get-ComputerList -Path $ComputerListPath)
    Write-Host "Computers   : $ComputerListPath"
    Write-Host ("Requested   : {0}" -f $requestedComputers.Count)
}
else {
    Write-Host "Computers   : no -ComputerListPath; exporting all AD computers." -ForegroundColor Yellow
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
