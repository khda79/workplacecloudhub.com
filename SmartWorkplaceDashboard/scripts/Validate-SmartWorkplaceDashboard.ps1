[CmdletBinding()]
param(
    [string]$DashboardRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$CheckSourceFiles
)

$ErrorActionPreference = 'Stop'

function Assert-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing file: $Path" }
    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json | Out-Null
    Write-Host "OK JSON $Path"
}

function Assert-FileContainsNoLegacyTerm {
    param([Parameter(Mandatory)][string]$Path)
    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $legacyPattern = ([string][char]101) + ([string][char]109) + ([string][char]101) + ([string][char]114) + ([string][char]105) + ([string][char]116)
    if ($content -match (('(?i)' + $legacyPattern))) { throw "Forbidden legacy term found in $Path" }
}

$pbipRoot = Join-Path $DashboardRoot 'pbip'
$reportDir = Join-Path $pbipRoot 'SmartWorkplaceDashboard.Report'
$modelDir = Join-Path $pbipRoot 'SmartWorkplaceDashboard.SemanticModel'

Assert-JsonFile -Path (Join-Path $pbipRoot 'SmartWorkplaceDashboard.pbip')
Assert-JsonFile -Path (Join-Path $reportDir 'definition.pbir')
Assert-JsonFile -Path (Join-Path $modelDir 'definition.pbism')
Assert-JsonFile -Path (Join-Path $modelDir 'model.bim')

$model = Get-Content -LiteralPath (Join-Path $modelDir 'model.bim') -Raw | ConvertFrom-Json
if (-not $model.model.tables -or $model.model.tables.Count -lt 10) { throw 'model.bim does not contain the expected initial tables.' }
if (-not $model.model.expressions -or $model.model.expressions.Count -lt 5) { throw 'model.bim does not contain the expected Power Query expressions.' }
if (-not (@($model.model.tables.name) -contains 'FactSourceFreshness')) { throw 'FactSourceFreshness table is missing.' }
if (-not (@($model.model.tables.name) -contains 'Measures')) { throw 'Measures table is missing.' }
Write-Host ("OK model tables={0} expressions={1}" -f $model.model.tables.Count, $model.model.expressions.Count)

Get-ChildItem -LiteralPath $DashboardRoot -Recurse -File | Where-Object { $_.Extension -in '.md','.pq','.json','.bim','.pbip','.pbir','.pbism','.ps1' } | ForEach-Object {
    Assert-FileContainsNoLegacyTerm -Path $_.FullName
}
Write-Host 'OK forbidden term scan'

if ($CheckSourceFiles) {
    $cmdbPath = ($model.model.expressions | Where-Object name -eq 'CMDBPowerBIPath').expression.Trim('"')
    $cmdbFullPath = if ([System.IO.Path]::IsPathRooted($cmdbPath)) { $cmdbPath } else { Join-Path $DashboardRoot $cmdbPath }
    $requiredCmdb = @('DimTenant.csv','DimDate.csv','DimUser.csv','DimDevice.csv','DimLicenseSku.csv','FactUserLicense.csv','FactDeviceCompliance.csv','FactMailbox.csv','FactDataQuality.csv')
    foreach ($fileName in $requiredCmdb) {
        $candidate = Join-Path $cmdbFullPath $fileName
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "Missing CMDB source file: $candidate" }
    }
    Write-Host "OK CMDB source files $cmdbFullPath"
}

Write-Host 'Smart Workplace Dashboard validation completed.'
