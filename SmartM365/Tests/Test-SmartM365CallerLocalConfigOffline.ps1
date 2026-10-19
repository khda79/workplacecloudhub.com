#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of Get-SmartM365CallerLocalConfig: a caller without a local JSON is skipped, an existing one keeps the
  JSON transport resolution and its UNC qualification.
.VERSION
1.0.0
.NOTES
  A real caller script runs from a local path and from an unqualified UNC path (\\localhost\<drive>$). Core 1.0.91.
#>
[CmdletBinding()]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$core = Join-Path $repository 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1'
$pwsh = (Get-Process -Id $PID).Path
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CallerConfig-' + [guid]::NewGuid().ToString('N'))
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

function Invoke-Caller {
  param([string]$ScriptPath)
  $output = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $ScriptPath -Core $core 2>&1 | ForEach-Object { [string]$_ }
  $line = @($output | Where-Object { $_ -match '^(OK|ERROR)\|' } | Select-Object -Last 1)
  [pscustomobject]@{ Result=$(if ($line.Count) { $line[0] } else { 'NONE|' + ($output -join ' ') }); Output=($output -join "`n") }
}

function Get-FolderFileName { param([string]$Folder) @(Get-ChildItem -LiteralPath $Folder -File -Force | ForEach-Object Name | Sort-Object) }

try {
  $caller = @'
param([string]$Core)
Import-Module $Core -Force
try { $config = & (Get-Module SmartM365.Core) { Get-SmartM365CallerLocalConfig }; 'OK|' + ($config | ConvertTo-Json -Compress) }
catch { 'ERROR|' + $_.Exception.Message }
'@
  $toUnc = { param([string]$Path) '\\localhost\{0}${1}' -f $Path.Substring(0, 1), $Path.Substring(2) }
  foreach ($case in @('absent', 'present-txt', 'present-legacy')) {
    $folder = Join-Path $root $case
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $folder 'Caller-Script.ps1') -Value $caller -Encoding utf8NoBOM
  }
  Set-Content -LiteralPath (Join-Path $root 'present-txt\Caller-Script.local.json.txt') -Value '{ "To": "caller-override" }' -Encoding utf8NoBOM
  Set-Content -LiteralPath (Join-Path $root 'present-legacy\Caller-Script.local.json') -Value '{ "To": "caller-legacy" }' -Encoding utf8NoBOM
  Check (Test-Path -LiteralPath (& $toUnc (Join-Path $root 'absent\Caller-Script.ps1'))) 'UNC admin share reachable for the test'

  # Absent under both names: skipped, nothing read or created, also on an unqualified UNC path.
  $absentFolder = Join-Path $root 'absent'
  $local = Invoke-Caller (Join-Path $absentFolder 'Caller-Script.ps1')
  Check ($local.Result -ceq 'OK|{}') "local caller without local JSON: empty configuration ($($local.Result))"
  $unc = Invoke-Caller (& $toUnc (Join-Path $absentFolder 'Caller-Script.ps1'))
  Check ($unc.Result -ceq 'OK|{}' -and -not $unc.Output.Contains('requires qualification')) "unqualified UNC caller without local JSON: skipped, no qualification error ($($unc.Result))"
  Check ((Get-FolderFileName $absentFolder) -join ',' -ceq 'Caller-Script.ps1') "no JSON, lock or journal created next to the caller ($((Get-FolderFileName $absentFolder) -join ','))"

  # Present: unchanged resolution; on an unqualified UNC path the qualification is still required.
  $txtFolder = Join-Path $root 'present-txt'
  $txtLocal = Invoke-Caller (Join-Path $txtFolder 'Caller-Script.ps1')
  Check ($txtLocal.Result -ceq 'OK|{"To":"caller-override"}') "local caller with Caller-Script.local.json.txt is read ($($txtLocal.Result))"
  $txtUnc = Invoke-Caller (& $toUnc (Join-Path $txtFolder 'Caller-Script.ps1'))
  Check ($txtUnc.Result -ceq 'ERROR|UNC JSON activation requires qualification on the target share.') "existing local JSON on an unqualified UNC path is still refused ($($txtUnc.Result))"
  $legacyFolder = Join-Path $root 'present-legacy'
  $legacyUnc = Invoke-Caller (& $toUnc (Join-Path $legacyFolder 'Caller-Script.ps1'))
  Check ($legacyUnc.Result -ceq 'ERROR|UNC JSON activation requires qualification on the target share.') "existing legacy .local.json on an unqualified UNC path is still refused ($($legacyUnc.Result))"
  Check ((Get-FolderFileName $legacyFolder) -join ',' -ceq 'Caller-Script.local.json,Caller-Script.ps1') "refused legacy file left untouched ($((Get-FolderFileName $legacyFolder) -join ','))"
}
finally {
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartM365 caller local configuration tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartM365 caller local configuration tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green
