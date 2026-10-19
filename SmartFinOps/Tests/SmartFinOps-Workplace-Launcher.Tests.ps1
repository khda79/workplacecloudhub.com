#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of Start-SmartFinOps-Workplace-Prod.cmd: Valuation then Report, never a report on an older valuation.
.VERSION
1.0.0
.NOTES
  The launcher runs next to stub scripts that record their call and exit with the requested code; no SmartFinOps
  script, configuration or data is used.
#>
[CmdletBinding()]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-Launcher-' + [guid]::NewGuid().ToString('N'))
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

try {
  New-Item -ItemType Directory -Path $root -Force | Out-Null
  $launcher = Join-Path $root 'Start-SmartFinOps-Workplace-Prod.cmd'
  Copy-Item -LiteralPath (Join-Path $project 'Start-SmartFinOps-Workplace-Prod.cmd') -Destination $launcher
  $calls = Join-Path $root 'calls.txt'
  foreach ($name in @('Valuation', 'Report')) {
    $stub = "param([string]`$Tenant)`r`nAdd-Content -LiteralPath `$env:FINOPS_LAUNCHER_TEST_CALLS -Value ('$name|' + `$Tenant)`r`nexit [int]`$env:FINOPS_LAUNCHER_TEST_$($name.ToUpperInvariant())_EXIT`r`n"
    Set-Content -LiteralPath (Join-Path $root "SmartFinOps-Workplace-$name.ps1") -Value $stub -Encoding utf8NoBOM
  }
  $run = {
    param([int]$Valuation, [int]$Report)
    if (Test-Path -LiteralPath $calls) { Remove-Item -LiteralPath $calls -Force }
    $env:FINOPS_LAUNCHER_TEST_CALLS = $calls
    $env:FINOPS_LAUNCHER_TEST_VALUATION_EXIT = [string]$Valuation
    $env:FINOPS_LAUNCHER_TEST_REPORT_EXIT = [string]$Report
    $output = & $env:ComSpec /d /c "`"$launcher`"" 2>&1 | Out-String
    [pscustomobject]@{ ExitCode=$LASTEXITCODE; Output=$output; Calls=@(if (Test-Path -LiteralPath $calls) { Get-Content -LiteralPath $calls } else { @() }) }
  }
  $cases = @(
    @{ V=0; R=0; Exit=0; Calls='Valuation|prod,Report|prod'; Name='success then success' }
    @{ V=3; R=0; Exit=3; Calls='Valuation|prod,Report|prod'; Name='valuation warnings: report runs, exit 3' }
    @{ V=0; R=3; Exit=3; Calls='Valuation|prod,Report|prod'; Name='report warnings: exit 3' }
    @{ V=3; R=1; Exit=1; Calls='Valuation|prod,Report|prod'; Name='report failure is the most severe: exit 1' }
    @{ V=0; R=2; Exit=1; Calls='Valuation|prod,Report|prod'; Name='any other report code is a failure: exit 1' }
    @{ V=1; R=0; Exit=1; Calls='Valuation|prod'; Name='valuation failure: report not run, exit 1' }
    @{ V=2; R=0; Exit=1; Calls='Valuation|prod'; Name='any other valuation code: report not run, exit 1' }
  )
  foreach ($case in $cases) {
    $result = & $run $case.V $case.R
    Check ($result.ExitCode -eq $case.Exit -and ($result.Calls -join ',') -ceq $case.Calls) ("{0} (exit {1}, calls {2})" -f $case.Name, $result.ExitCode, ($result.Calls -join ','))
    if ($case.Calls -ceq 'Valuation|prod') {
      Check ($result.Output.Contains("Valuation failed with exit code $($case.V): the report is not run, so no report is built on an older valuation.")) "$($case.Name): explicit message"
    }
  }
  $separate = @(Get-ChildItem -LiteralPath $project -Filter 'Start-SmartFinOps-Workplace-*-Prod.cmd' | ForEach-Object Name | Sort-Object)
  Check (($separate -join ',') -ceq 'Start-SmartFinOps-Workplace-Report-Prod.cmd,Start-SmartFinOps-Workplace-Valuation-Prod.cmd') "separate launchers kept ($($separate -join ','))"
}
finally {
  foreach ($name in @('FINOPS_LAUNCHER_TEST_CALLS', 'FINOPS_LAUNCHER_TEST_VALUATION_EXIT', 'FINOPS_LAUNCHER_TEST_REPORT_EXIT')) { Remove-Item -Path "Env:$name" -ErrorAction SilentlyContinue }
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartFinOps launcher tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartFinOps launcher tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green
