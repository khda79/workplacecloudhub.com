<#
.SYNOPSIS
Offline checks of Publish-SmartM365-DataAllBackfill.ps1 (file selection, preview default, guards).

.VERSION
1.0.0
#>
#requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '..\SmartInventory\Maintenance\Publish-SmartM365-DataAllBackfill.ps1'
$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++ }

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath), [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Script does not parse.'
$planFunction = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-DataAllBackfillPlan' }, $true)
Check ($null -ne $planFunction) 'Get-DataAllBackfillPlan not found.'
. ([scriptblock]::Create($planFunction.Extent.Text))

$text = Get-Content -LiteralPath $scriptPath -Raw
Check ($ast.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Execute') 'No -Execute switch.'
Check ($text -match '\.VERSION\s+1\.0') 'Help .VERSION missing.'
Check ($text -match "if \(\`$Execute -and -not \`$policy\.Enabled\) \{ throw") 'Execute is not refused when DATA-ALL uploads are disabled.'
Check ($text -match "Invoke-SmartM365SharePointCsvUpload -LocalFilePath \`$entry\.FullName -Enabled \`$true -EnsureParentFolders") 'Upload does not reuse the Core upload.'
Check ($text -notmatch 'Remove-SmartM365SharePoint') 'The backfill must never delete on SharePoint.'
Check (([regex]::Matches($text, 'Invoke-SmartM365SharePointCsvUpload')).Count -eq 1) 'Upload called outside the -Execute branch.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-BackfillTest-' + [guid]::NewGuid().ToString('N'))
$dataAll = Join-Path $root 'DATA-ALL'
try {
    $now = [datetime]'2026-11-17T10:00:00'
    function New-TestFile([string]$Relative, [datetime]$Written, [int]$Size = 10) {
        $path = Join-Path $dataAll $Relative
        [IO.Directory]::CreateDirectory((Split-Path $path -Parent)) | Out-Null
        [IO.File]::WriteAllBytes($path, [byte[]]::new($Size))
        (Get-Item -LiteralPath $path).LastWriteTime = $Written
    }
    New-TestFile 'Intune\Devices\Devices_20261116_010000.csv' $now.AddHours(-33)
    New-TestFile 'Intune\Devices\Devices_20261105_010000.csv' $now.AddDays(-12)
    New-TestFile 'Intune\Devices\Devices_20261020_010000.csv' $now.AddDays(-28)
    New-TestFile 'Intune\Devices\Devices_20261116-010000.xlsx' $now.AddHours(-33)
    New-TestFile 'Intune\Devices\WeeklyHistory\2026-W45\Devices.csv' $now.AddDays(-9)
    New-TestFile 'Intune\Devices\WeeklyHistory\2026-W45\manifest.json.txt' $now.AddDays(-9)
    New-TestFile 'Usage\20261110_044500\Usage.csv' $now.AddDays(-7)
    New-TestFile 'Orchestrator\Jobs.json.txt' $now.AddDays(-1)
    New-TestFile 'Intune\.staging\Devices.csv' $now.AddDays(-1)
    New-TestFile 'Intune\Devices\Devices.csv.lock' $now.AddDays(-1)
    New-TestFile 'Intune\Devices\Devices_20261117_095000.csv' $now.AddMinutes(-10)

    $plan = @(Get-DataAllBackfillPlan -DataAllRoot $dataAll -Since ([datetime]'2026-10-24') -TimestampedCutoff $now.AddDays(-7) -NewestAllowed $now.AddMinutes(-30))
    $byName = @{}; foreach ($entry in $plan) { $byName[$entry.Folder + '|' + $entry.Name] = $entry }
    function Expect([string]$Key, [string]$Action, [string]$Reason) {
        Check ($byName.ContainsKey($Key)) "Missing plan entry $Key."
        Check ($byName[$Key].Action -eq $Action -and $byName[$Key].Reason -eq $Reason) ("{0}: expected {1}/{2}, got {3}/{4}." -f $Key, $Action, $Reason, $byName[$Key].Action, $byName[$Key].Reason)
    }
    Check (-not $byName.ContainsKey('Intune\Devices|Devices_20261020_010000.csv')) 'A file written before -Since was planned.'
    Expect 'Intune\Devices|Devices_20261116_010000.csv' 'Upload' ''
    Expect 'Intune\Devices|Devices_20261116-010000.xlsx' 'Upload' ''
    Expect 'Intune\Devices|Devices_20261105_010000.csv' 'Skip' 'BeyondSharePointRetention'
    Expect 'Intune\Devices\WeeklyHistory\2026-W45|Devices.csv' 'Upload' ''
    Expect 'Intune\Devices\WeeklyHistory\2026-W45|manifest.json.txt' 'Upload' ''
    Expect 'Usage\20261110_044500|Usage.csv' 'Upload' ''
    Expect 'Orchestrator|Jobs.json.txt' 'Skip' 'OrchestratorFolder'
    Expect 'Intune\.staging|Devices.csv' 'Skip' 'HiddenFolder'
    Expect 'Intune\Devices|Devices.csv.lock' 'Skip' 'TransientFile'
    Expect 'Intune\Devices|Devices_20261117_095000.csv' 'Skip' 'TooRecent'
    $ordered = @($plan | ForEach-Object LastWriteTime)
    Check ((@($ordered | Sort-Object) -join '|') -eq ($ordered -join '|')) 'Plan is not ordered oldest first.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "DATA-ALL backfill offline: $checks checks passed."
