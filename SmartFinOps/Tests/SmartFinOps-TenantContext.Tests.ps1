#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps tenant context: SmartM365 DATA-LAST resolution and console banners.
.VERSION
1.0.0
.NOTES
  Isolated temporary runtime with a synthetic SmartM365 tenant context (no local configuration or export).
  SmartM365LatestCsvFolderPath = __USE_SMARTM365__ must never fall back silently to a default path.
#>
[CmdletBinding()]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$repository = Split-Path $project -Parent
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-Context-' + [guid]::NewGuid().ToString('N'))
$runtime = Join-Path $root 'runtime'
$engine = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1'
$report = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Report.ps1'
$context = Join-Path $runtime 'SmartFinOps\Config\SmartFinOps-TenantContext.ps1'
$smartM365Context = Join-Path $runtime 'SmartM365\Config\SmartM365-TenantContext.ps1'
$pwsh = (Get-Process -Id $PID).Path
$tenant = 'contoso'
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

foreach ($file in @('SmartFinOps-Workplace-Valuation.ps1','SmartFinOps-Workplace-Report.ps1','SmartFinOps.global.local.json.template','Config\SmartFinOps-TenantContext.ps1','Config\SmartFinOps-Contracts.local.json.template')) {
  $target = Join-Path $runtime "SmartFinOps\$file"
  New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $project $file) -Destination $target
}
New-Item -ItemType Directory -Path (Join-Path $runtime 'SmartM365\Modules'), (Split-Path $smartM365Context -Parent) -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repository 'SmartM365\Modules\SmartM365.Core') -Destination (Join-Path $runtime 'SmartM365\Modules') -Recurse
'{ "SmartM365LatestCsvFolderPath": "__USE_SMARTM365__" }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
$dataLast = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory -Path $dataLast -Force | Out-Null

function Set-SmartM365Context {
  param([string]$Body)
  if ([string]::IsNullOrEmpty($Body)) { Remove-Item -LiteralPath $smartM365Context -Force -ErrorAction SilentlyContinue; return }
  Set-Content -LiteralPath $smartM365Context -Value $Body -Encoding utf8NoBOM
}

function Invoke-Script {
  param([string]$Path, [string[]]$Arguments = @())
  $log = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $Path -Tenant $tenant @Arguments 2>&1 | Out-String
  return [pscustomobject]@{ ExitCode=$LASTEXITCODE; Log=$log }
}

function Get-ResolvedDataLast {
  # Same calls as the engine: initialize the context, then ask for the configured DATA-LAST.
  $command = ". '$context'; Initialize-SmartFinOpsTenantContext -Tenant '$tenant' -StartPath '$(Split-Path $context -Parent)' | Out-Null; try { 'OK|' + (Get-SmartFinOpsConfigValue -Name 'SmartM365LatestCsvFolderPath') } catch { 'ERROR|' + `$_.Exception.Message }"
  # The context may first report a local JSON completed from the template: keep the answer line only.
  return [string](@(& $pwsh -NoProfile -Command $command 2>&1 | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(OK|ERROR)\|' }) | Select-Object -Last 1)
}

try {
  #region 1. SmartM365 context that throws: explicit failure with message and position, no default path
  Set-SmartM365Context @'
function Get-SmartM365EffectiveGlobalConfig {
  param([string]$StartPath, [string]$TenantKey)
  throw "boom from the SmartM365 configuration of $TenantKey"
}
'@
  $resolved = Get-ResolvedDataLast
  Check ($resolved -like 'ERROR|*' -and $resolved.Contains('resolution through SmartM365 failed') -and $resolved.Contains("boom from the SmartM365 configuration of $tenant") -and $resolved -match 'SmartM365-TenantContext\.ps1:3') "SmartM365 error surfaced with message and position (got $resolved)"
  $run = Invoke-Script $engine @('-OutputRoot', (Join-Path $root 'out1'), '-ContractsPath', (Join-Path $root 'contracts1.local.json'))
  Check ($run.ExitCode -eq 1 -and $run.Log.Contains('Run failed: SmartM365LatestCsvFolderPath is __USE_SMARTM365__ but its resolution through SmartM365 failed') -and $run.Log.Contains('boom from the SmartM365 configuration')) 'engine fails with the SmartM365 resolution reason'
  Check (-not $run.Log.Contains('DATA-LAST folder not found') -and -not ($run.Log -match 'SmartM365\\Data\\Tenants')) 'no silent fallback to the default DATA-LAST path'
  Check ($run.Log.Contains(' SmartFinOps by WorkplaceCloudHub') -and $run.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $run.Log.Contains(' Status    : FAILED') -and $run.Log.Contains(' Script    : SmartFinOps-Workplace-Valuation.ps1') -and $run.Log.Contains(' CSV files : 0') -and $run.Log -match ' Errors    : [1-9]') 'engine failure: startup banner and Failed completion banner'
  $explicit = Invoke-Script $engine @('-SmartM365LatestCsvFolderPath', $dataLast, '-OutputRoot', (Join-Path $root 'out2'), '-ContractsPath', (Join-Path $root 'contracts2.local.json'), '-ValidateOnly')
  Check ($explicit.ExitCode -eq 0 -and -not $explicit.Log.Contains('resolution through SmartM365 failed') -and $explicit.Log.Contains(' Status    : ')) 'an explicit -SmartM365LatestCsvFolderPath never calls the SmartM365 context'
  #endregion

  #region 2. SmartM365 context missing
  Set-SmartM365Context $null
  $resolved = Get-ResolvedDataLast
  Check ($resolved -like 'ERROR|*' -and $resolved.Contains('SmartM365 tenant context not found')) "missing SmartM365 context is explicit (got $resolved)"
  #endregion

  #region 3. Unresolved tokens are refused
  Set-SmartM365Context @'
function Get-SmartM365EffectiveGlobalConfig {
  param([string]$StartPath, [string]$TenantKey)
  [pscustomobject]@{ WorkspaceRootPath='D:\Workplace'; LatestCsvFolderPath='{{WorkspaceRootPath}}\{{NotDefined}}\DATA-LAST' }
}
'@
  $resolved = Get-ResolvedDataLast
  Check ($resolved -like 'ERROR|*' -and $resolved.Contains('unresolved token {{NotDefined}}') -and $resolved.Contains('D:\Workplace\{{NotDefined}}\DATA-LAST')) "token left by SmartM365 refused (got $resolved)"
  Set-SmartM365Context @'
function Get-SmartM365EffectiveGlobalConfig {
  param([string]$StartPath, [string]$TenantKey)
  [pscustomobject]@{ WorkspaceRootPath='D:\Workplace'; LatestCsvFolderPath='' }
}
'@
  $resolved = Get-ResolvedDataLast
  Check ($resolved -like 'ERROR|*' -and $resolved.Contains('empty LatestCsvFolderPath')) "empty SmartM365 path refused (got $resolved)"
  '{ "SmartM365LatestCsvFolderPath": "{{SmartFinOpsRootPath}}\\{{NoSuchToken}}\\DATA-LAST" }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $resolved = Get-ResolvedDataLast
  Check ($resolved -like 'ERROR|*' -and $resolved.Contains('unresolved token {{NoSuchToken}}') -and -not $resolved.Contains('{{SmartFinOpsRootPath}}')) "token left in the SmartFinOps configuration refused (got $resolved)"
  #endregion

  #region 4. Valid resolution through SmartM365
  '{ "SmartM365LatestCsvFolderPath": "__USE_SMARTM365__" }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Set-SmartM365Context @"
function Get-SmartM365EffectiveGlobalConfig {
  param([string]`$StartPath, [string]`$TenantKey)
  [pscustomobject]@{ WorkspaceRootPath='$root'; LatestCsvFolderPath='{{WorkspaceRootPath}}\DATA-LAST' }
}
"@
  $resolved = Get-ResolvedDataLast
  Check ($resolved -ceq "OK|$dataLast") "SmartM365 tokens resolved (got $resolved)"
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Check ((Get-ResolvedDataLast) -ceq "OK|$dataLast") 'an empty SmartFinOps setting also resolves through SmartM365'
  #endregion

  #region 5. Report banners and tokenized data root
  $reportRun = Invoke-Script $report @('-InputRoot', (Join-Path $root '{{NoSuchToken}}'))
  Check ($reportRun.ExitCode -eq 1 -and $reportRun.Log.Contains('unresolved token {{NoSuchToken}}') -and $reportRun.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $reportRun.Log.Contains(' Script    : SmartFinOps-Workplace-Report.ps1')) 'report: tokenized data root refused with a Failed banner'
  #endregion
}
finally {
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartFinOps context tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartFinOps context tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green
