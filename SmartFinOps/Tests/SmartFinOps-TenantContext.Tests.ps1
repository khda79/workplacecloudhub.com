#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps tenant context: SmartM365 DATA-LAST resolution and console banners.
.VERSION
1.4.0
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
$dataAll = Join-Path $root 'DATA-ALL'
$logAll = Join-Path $root 'LOG-ALL'
New-Item -ItemType Directory -Path $dataLast, $dataAll, $logAll -Force | Out-Null
# The synthetic SmartM365 contexts below stand for a configured profile: its runtime JSON files exist.
New-Item -ItemType Directory -Path (Join-Path $runtime 'SmartM365\Config\Tenants') -Force | Out-Null
foreach ($runtimeJson in @('SmartM365\Config\SmartM365.global.local.json.txt', "SmartM365\Config\Tenants\$tenant.local.json.txt")) { Set-Content -LiteralPath (Join-Path $runtime $runtimeJson) -Value '{}' -Encoding utf8NoBOM }

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
  param([string]$Name = 'SmartM365LatestCsvFolderPath')
  # Same calls as the engine: initialize the context ($tenant is the profile key), then ask for the configured value.
  $command = ". '$context'; Initialize-SmartFinOpsTenantContext -Tenant '$tenant' -StartPath '$(Split-Path $context -Parent)' | Out-Null; try { 'OK|' + (Get-SmartFinOpsConfigValue -Name '$Name') } catch { 'ERROR|' + `$_.Exception.Message }"
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
  $run = Invoke-Script $engine @('-ContractsPath', (Join-Path $root 'contracts1.local.json'))
  Check ($run.ExitCode -eq 1 -and $run.Log.Contains('Run failed: SmartM365LatestCsvFolderPath is __USE_SMARTM365__ but its resolution through SmartM365 failed') -and $run.Log.Contains('boom from the SmartM365 configuration')) 'engine fails with the SmartM365 resolution reason'
  Check (-not $run.Log.Contains('DATA-LAST folder not found') -and -not ($run.Log -match 'SmartM365\\Data\\Tenants')) 'no silent fallback to the default DATA-LAST path'
  Check ($run.Log.Contains(' SmartFinOps by WorkplaceCloudHub') -and $run.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $run.Log.Contains(' Status    : FAILED') -and $run.Log.Contains(' Script    : SmartFinOps-Workplace-Valuation.ps1') -and $run.Log.Contains(' CSV files : 0') -and $run.Log -match ' Errors    : [1-9]') 'engine failure: startup banner and Failed completion banner'
  $explicit = Invoke-Script $engine @('-SmartM365LatestCsvFolderPath', $dataLast, '-SmartM365DataAllRootPath', $dataAll, '-SmartM365LogAllRootPath', $logAll, '-SmartM365TenantKey', 'contoso-prod', '-ContractsPath', (Join-Path $root 'contracts2.local.json'), '-ValidateOnly')
  Check ($explicit.ExitCode -eq 3 -and -not $explicit.Log.Contains('resolution through SmartM365 failed') -and $explicit.Log.Contains(' Status    : ')) 'explicit data roots and -SmartM365TenantKey never call the SmartM365 context (ValidateOnly: no upload either)'
  $tenantKeyFailure = Get-ResolvedDataLast 'SmartM365TenantKey'
  Check ($tenantKeyFailure -like 'ERROR|SmartM365TenantKey is __USE_SMARTM365__ but its resolution through SmartM365 failed*' -and $tenantKeyFailure.Contains('boom from the SmartM365 configuration')) "TenantKey resolution failure is explicit (got $tenantKeyFailure)"
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
  [pscustomobject]@{ WorkspaceRootPath='$root'; LatestCsvFolderPath='{{WorkspaceRootPath}}\DATA-LAST'; ProfileKey=`$TenantKey; TenantKey='contoso-prod' }
}
"@
  $resolved = Get-ResolvedDataLast
  Check ($resolved -ceq "OK|$dataLast") "SmartM365 tokens resolved (got $resolved)"
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Check ((Get-ResolvedDataLast) -ceq "OK|$dataLast") 'an empty SmartFinOps setting also resolves through SmartM365'
  Check ((Get-ResolvedDataLast 'SmartM365TenantKey') -ceq 'OK|contoso-prod') 'TenantKey of the profile resolved through SmartM365 (profile contoso, TenantKey contoso-prod)'
  '{ "SmartM365TenantKey": "fabrikam-test" }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Set-SmartM365Context 'function Get-SmartM365EffectiveGlobalConfig { throw "SmartM365 must not be called" }'
  Check ((Get-ResolvedDataLast 'SmartM365TenantKey') -ceq 'OK|fabrikam-test') 'a TenantKey set in the SmartFinOps configuration is used without SmartM365'
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Set-SmartM365Context 'function Get-SmartM365EffectiveGlobalConfig { [pscustomobject]@{ LatestCsvFolderPath="C:\x" } }'
  $noKey = Get-ResolvedDataLast 'SmartM365TenantKey'
  Check ($noKey -like 'ERROR|*returned no TenantKey*') "SmartM365 context without TenantKey refused (got $noKey)"
  #endregion

  #region 5. Report banners and tokenized data root
  $reportRun = Invoke-Script $report @('-SmartM365LatestCsvFolderPath', (Join-Path $root '{{NoSuchToken}}\DATA-LAST'))
  Check ($reportRun.ExitCode -eq 1 -and $reportRun.Log.Contains('unresolved token {{NoSuchToken}}') -and $reportRun.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $reportRun.Log.Contains(' Script    : SmartFinOps-Workplace-Report.ps1')) 'report: tokenized data root refused with a Failed banner'
  #endregion

  #region 6. Run from a UNC path (\\localhost\<drive>$): plain file-system paths only, never provider-qualified
  $toUnc = { param([string]$Path) '\\localhost\{0}${1}' -f $Path.Substring(0, 1), $Path.Substring(2) }
  $uncRoot = & $toUnc $root
  $uncRuntime = & $toUnc $runtime
  Check (Test-Path -LiteralPath (Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1')) "UNC admin share reachable for the test: $uncRuntime"
  # Like the real SmartM365 context, the synthetic one checks its local JSON with .NET, which reads a provider-qualified path as relative.
  Set-Content -LiteralPath (Join-Path $runtime 'SmartM365\Config\SmartM365.global.local.json.txt') -Value '{}' -Encoding utf8NoBOM
  Set-SmartM365Context @"
function Get-SmartM365EffectiveGlobalConfig {
  param([string]`$StartPath, [string]`$TenantKey)
  `$local = Join-Path `$StartPath 'Config\SmartM365.global.local.json.txt'
  if (-not [IO.File]::Exists(`$local)) { throw ('Local JSON not found: {0}' -f [IO.Path]::GetFullPath(`$local)) }
  [pscustomobject]@{ WorkspaceRootPath='$uncRoot'; LatestCsvFolderPath='{{WorkspaceRootPath}}\DATA-LAST'; DataAllRootPath='{{WorkspaceRootPath}}\DATA-ALL'; LogAllRootPath='{{WorkspaceRootPath}}\LOG-ALL'; ProfileKey=`$TenantKey; TenantKey='contoso-prod'; EnableSharePointUpload=`$false }
}
"@
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $cwd = Join-Path $root 'cwd'
  New-Item -ItemType Directory -Path $cwd -Force | Out-Null
  Push-Location -LiteralPath $cwd
  try {
    $uncEngine = & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1') -Tenant $tenant 2>&1 | Out-String
    $uncEngineExit = $LASTEXITCODE
    $uncReport = & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Report.ps1') -Tenant $tenant 2>&1 | Out-String
    $uncReportExit = $LASTEXITCODE
  }
  finally { Pop-Location }
  Check ($uncEngineExit -eq 3 -and $uncEngine.Contains(('Sources and current outputs: {0}\DATA-LAST' -f $uncRoot)) -and $uncEngine.Contains(('History: {0}\DATA-ALL\FinOps\Workplace' -f $uncRoot)) -and $uncEngine.Contains(('Log: {0}\LOG-ALL\SmartFinOps-Workplace-Valuation\' -f $uncRoot))) "engine from UNC: DATA-LAST, DATA-ALL and LOG-ALL resolved through SmartM365 (exit $uncEngineExit) $uncEngine"
  $uncReceipt = Get-Content -LiteralPath (Join-Path $dataLast 'SmartFinOps_Workplace-Valuation.current.json.txt') -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
  Check ($null -ne $uncReceipt -and $uncReceipt.ProfileKey -ceq $tenant -and $uncReceipt.TenantKey -ceq 'contoso-prod' -and $uncEngine.Contains("Profile 'contoso' - TenantKey 'contoso-prod'")) 'engine from UNC: receipt in DATA-LAST with the TenantKey of the profile'
  Check (@(Get-ChildItem -LiteralPath $dataLast -Filter 'FinOps_*.csv' -ErrorAction SilentlyContinue).Count -eq 7 -and (Test-Path -LiteralPath (Join-Path $dataAll "FinOps\Workplace\$($uncReceipt.RunId)\SmartFinOps_Workplace-Valuation.current.json.txt")) -and (Test-Path -LiteralPath (Join-Path $root ([string]$uncReceipt.LogFile).Replace('/', '\')))) 'engine from UNC: seven CSVs in DATA-LAST, run history in DATA-ALL\FinOps\Workplace, log in LOG-ALL'
  Check (-not (Test-Path -LiteralPath (Join-Path $runtime 'SmartFinOps\Data')) -and -not (Test-Path -LiteralPath (Join-Path $runtime 'SmartFinOps\Output'))) 'nothing written in the SmartFinOps code folder'
  Check ($uncReportExit -eq 3 -and (Test-Path -LiteralPath (Join-Path $dataLast 'SmartFinOps_Workplace-Report.html')) -and (Test-Path -LiteralPath (Join-Path $dataAll "FinOps\Workplace\$($uncReceipt.RunId)\SmartFinOps_Workplace-Report.html")) -and $uncReport.Contains(('Inputs and page: {0}\DATA-LAST' -f $uncRoot)) -and @(Get-ChildItem -LiteralPath (Join-Path $logAll 'SmartFinOps-Workplace-Report') -Filter '*.log').Count -eq 1) "report from UNC: page in DATA-LAST and in the run history, log in LOG-ALL (exit $uncReportExit) $uncReport"
  $written = @(Get-ChildItem -LiteralPath $dataLast, $dataAll, $logAll -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.txt', '.log', '.csv', '.html' } | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw })
  Check (-not ($uncEngine + $uncReport + ($written -join '')).Contains('FileSystem::')) 'no provider-qualified path in console, logs or outputs'
  Check (@(Get-ChildItem -LiteralPath $cwd -Recurse -Force).Count -eq 0) 'nothing written relative to the process working directory'
  # Relative data roots are taken from the PowerShell location (here on the share), not from the process directory.
  $relative = Join-Path $root 'relative'
  New-Item -ItemType Directory -Path (Join-Path $relative 'DATA-LAST'), (Join-Path $relative 'DATA-ALL'), (Join-Path $relative 'LOG-ALL') -Force | Out-Null
  Push-Location -LiteralPath $cwd
  try {
    $command = "Set-Location -LiteralPath '$(& $toUnc $relative)'; & '$(Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1')' -Tenant '$tenant' -SmartM365LatestCsvFolderPath 'DATA-LAST' -SmartM365DataAllRootPath 'DATA-ALL' -SmartM365LogAllRootPath 'LOG-ALL'; exit `$LASTEXITCODE"
    $relativeRun = & $pwsh -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1 | Out-String
    $relativeExit = $LASTEXITCODE
  }
  finally { Pop-Location }
  Check ($relativeExit -eq 3 -and (Test-Path -LiteralPath (Join-Path $relative 'DATA-LAST\SmartFinOps_Workplace-Valuation.current.json.txt')) -and @(Get-ChildItem -LiteralPath $cwd -Recurse -Force).Count -eq 0) "relative data roots resolved from the PowerShell location on the share (exit $relativeExit) $relativeRun"
  #endregion

  #region 7. Profile files: legacy TenantKey renamed to ProfileKey before template synchronization
  $tenantsFolder = Join-Path $runtime 'SmartFinOps\Config\Tenants'
  New-Item -ItemType Directory -Path $tenantsFolder -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $project 'Config\Tenants\tenant.local.json.template') -Destination $tenantsFolder
  $profileFile = Join-Path $tenantsFolder "$tenant.local.json"
  # Profile-file errors are raised by the initialization itself: capture them too.
  $getProfileResult = {
    $command = "try { . '$context'; Initialize-SmartFinOpsTenantContext -Tenant '$tenant' -StartPath '$(Split-Path $context -Parent)' | Out-Null; 'OK|' + (Get-SmartFinOpsConfigValue -Name 'ProfileKey') } catch { 'ERROR|' + `$_.Exception.Message }"
    [string](@(& $pwsh -NoProfile -Command $command 2>&1 | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(OK|ERROR)\|' }) | Select-Object -Last 1)
  }
  ('{ "TenantKey": "' + $tenant + '", "SmartM365LatestCsvFolderPath": "__USE_GLOBAL__" }') | Set-Content -LiteralPath $profileFile -Encoding utf8NoBOM
  $dataRoot = & $getProfileResult
  $migrated = Get-Content -LiteralPath $profileFile -Raw | ConvertFrom-Json
  $names = @($migrated.PSObject.Properties.Name)
  Check ($dataRoot -like 'OK|*' -and $names[0] -ceq 'ProfileKey' -and $migrated.ProfileKey -ceq $tenant -and $names -notcontains 'TenantKey' -and $names -contains 'SmartM365TenantKey') "legacy TenantKey renamed to ProfileKey in place, template key added, no 'test' profile injected ($dataRoot)"
  Check ($dataRoot -ceq "OK|$tenant") "ProfileKey of the context is the requested profile ($dataRoot)"
  ('{ "TenantKey": "other" }') | Set-Content -LiteralPath $profileFile -Encoding utf8NoBOM
  $mismatch = & $getProfileResult
  Check ($mismatch -like 'ERROR|*' -and $mismatch.Contains("contains ProfileKey 'other' but requested '$tenant'")) "legacy TenantKey of another profile refused after renaming ($mismatch)"
  $conflict = '{ "ProfileKey": "' + $tenant + '", "TenantKey": "other" }'
  $conflict | Set-Content -LiteralPath $profileFile -Encoding utf8NoBOM
  $conflictResult = & $getProfileResult
  Check ($conflictResult -like 'ERROR|*' -and $conflictResult.Contains('keep a single ProfileKey') -and (Get-Content -LiteralPath $profileFile -Raw).Trim() -ceq $conflict) "ProfileKey and a different legacy TenantKey refused, file unchanged ($conflictResult)"
  Remove-Item -LiteralPath $profileFile -Force
  $globalFile = Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json'
  $globalBefore = Get-Content -LiteralPath $globalFile -Raw
  try {
    $legacyGlobal = $globalBefore | ConvertFrom-Json -AsHashtable
    $legacyGlobal['SmartM365DataAllRootPath'] = '{{SmartFinOpsRootPath}}\Legacy\{{TenantKey}}\DATA-ALL'
    $legacyGlobal | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $globalFile -Encoding utf8NoBOM
    $legacyRoot = Get-ResolvedDataLast -Name 'SmartM365DataAllRootPath'
    Check ($legacyRoot -like ('OK|*' + [IO.Path]::DirectorySeparatorChar + 'Legacy' + [IO.Path]::DirectorySeparatorChar + $tenant + [IO.Path]::DirectorySeparatorChar + 'DATA-ALL')) "{{TenantKey}} remains a read alias of the profile key ($legacyRoot)"
  }
  finally { Set-Content -LiteralPath $globalFile -Value $globalBefore -NoNewline -Encoding utf8NoBOM }
  #endregion

  #region 8. Data layout: existing sibling DATA-LAST, DATA-ALL and LOG-ALL only, never the SmartM365 Output fallback
  $getLayout = {
    param([string]$Last, [string]$All, [string]$Logs)
    $command = "try { . '$context'; Initialize-SmartFinOpsTenantContext -Tenant '$tenant' -StartPath '$(Split-Path $context -Parent)' | Out-Null; `$l = Get-SmartFinOpsDataLayout -DataLast '$Last' -DataAll '$All' -LogAll '$Logs'; 'OK|' + `$l.HistoryRoot } catch { 'ERROR|' + `$_.Exception.Message }"
    [string](@(& $pwsh -NoProfile -Command $command 2>&1 | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(OK|ERROR)\|' }) | Select-Object -Last 1)
  }
  Check ((& $getLayout $dataLast $dataAll $logAll) -ceq ('OK|' + (Join-Path $dataAll 'FinOps\Workplace'))) 'sibling data roots: history under DATA-ALL\FinOps\Workplace'
  $otherAll = Join-Path $root 'other\DATA-ALL'; New-Item -ItemType Directory -Path $otherAll -Force | Out-Null
  $notSibling = & $getLayout $dataLast $otherAll $logAll
  Check ($notSibling -like 'ERROR|*' -and $notSibling.Contains('must be sibling folders')) "data roots of different tenants refused ($notSibling)"
  $wrongLeaf = Join-Path $root 'DATA-HISTORY'; New-Item -ItemType Directory -Path $wrongLeaf -Force | Out-Null
  $leafResult = & $getLayout $dataLast $wrongLeaf $logAll
  Check ($leafResult -like 'ERROR|*' -and $leafResult.Contains('must name a folder called DATA-ALL')) "a history root not called DATA-ALL refused ($leafResult)"
  $bare = Join-Path $root 'bare'; New-Item -ItemType Directory -Path (Join-Path $bare 'DATA-LAST'), (Join-Path $bare 'DATA-ALL') -Force | Out-Null
  $missingLogs = & $getLayout (Join-Path $bare 'DATA-LAST') (Join-Path $bare 'DATA-ALL') (Join-Path $bare 'LOG-ALL')
  Check ($missingLogs -like 'ERROR|*' -and $missingLogs.Contains('never creates the SmartM365 data roots') -and -not (Test-Path -LiteralPath (Join-Path $bare 'LOG-ALL'))) "a missing data root is refused, never created ($missingLogs)"
  $fallbackRoot = Join-Path $root 'smartm365-output'
  Set-SmartM365Context @"
function Get-SmartM365EffectiveGlobalConfig {
  param([string]`$StartPath, [string]`$TenantKey)
  [pscustomobject]@{ ScriptOutputRootPath='$fallbackRoot'; LatestCsvFolderPath='$fallbackRoot\Tenants\contoso\DATA-LAST'; TenantKey='contoso-prod' }
}
"@
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $fallback = Get-ResolvedDataLast
  Check ($fallback -like 'ERROR|*' -and $fallback.Contains('SmartM365 Output fallback')) "a SmartM365 path under its Output fallback is refused ($fallback)"
  #endregion

  #region 9. SharePoint upload through the real SmartM365 context of the profile (no synthetic context)
  # Real SmartM365 tenant context, JSON transport and SmartM365.Core; profile prod only, no test profile file.
  $real = Join-Path $root 'real'
  $realSmartM365 = Join-Path $real 'SmartM365'
  foreach ($file in @('SmartFinOps-Workplace-Valuation.ps1','SmartFinOps-Workplace-Report.ps1','SmartFinOps.global.local.json.template','Config\SmartFinOps-TenantContext.ps1','Config\SmartFinOps-Contracts.local.json.template')) {
    $target = Join-Path $real "SmartFinOps\$file"
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $project $file) -Destination $target
  }
  '{ }' | Set-Content -LiteralPath (Join-Path $real 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  foreach ($file in @('Config\SmartM365-TenantContext.ps1','Config\SmartM365-JsonTransport.policy.psd1','Config\SmartM365.global.local.json.template','Config\Tenants\tenant.local.json.template')) {
    $target = Join-Path $realSmartM365 $file
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repository "SmartM365\$file") -Destination $target
  }
  New-Item -ItemType Directory -Path (Join-Path $realSmartM365 'Modules') -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $repository 'SmartM365\Modules\SmartM365.Core') -Destination (Join-Path $realSmartM365 'Modules') -Recurse
  $workplace = Join-Path $real 'WORKPLACE'
  foreach ($folder in @('DATA-LAST','DATA-ALL','LOG-ALL')) { New-Item -ItemType Directory -Path (Join-Path $workplace "Data\Tenants\prod\$folder") -Force | Out-Null }
  '{}' | Set-Content -LiteralPath (Join-Path $realSmartM365 'Config\SmartM365.global.local.json.txt') -Encoding utf8NoBOM
  $prodProfile = [ordered]@{ ProfileKey='prod'; OrganizationKey='contoso'; EnvironmentKey='prod'; TenantKey='contoso-prod'; TenantId='11111111-2222-3333-4444-555555555555'
    AppId='66666666-7777-8888-9999-000000000001'; Thumb='ABCDEF0123456789ABCDEF0123456789ABCDEF01'; WorkspaceRootPath=$workplace
    EnableSharePointUpload=$true; SharePointSiteHostname='contoso.sharepoint.invalid'; SharePointSitePath='/sites/finops'; SharePointLibraryDisplayName='Documents'; SharePointTargetFolderPath='SMART-M365/CSV' }
  $prodProfilePath = Join-Path $realSmartM365 'Config\Tenants\prod.local.json.txt'
  $prodProfile | ConvertTo-Json | Set-Content -LiteralPath $prodProfilePath -Encoding utf8NoBOM
  $runtimeJsonFiles = { @(Get-ChildItem -LiteralPath (Join-Path $realSmartM365 'Config') -Recurse -File | Where-Object { $_.Name -match '\.json(\.txt)?$' } | ForEach-Object { $_.FullName.Substring($realSmartM365.Length) } | Sort-Object) }
  $jsonBefore = & $runtimeJsonFiles
  $realEngine = Join-Path $real 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1'
  $realReport = Join-Path $real 'SmartFinOps\SmartFinOps-Workplace-Report.ps1'
  $realLast = Join-Path $workplace 'Data\Tenants\prod\DATA-LAST'
  $upload = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $realEngine -Tenant prod 2>&1 | Out-String
  $uploadExit = $LASTEXITCODE
  $uploadReceipt = Get-Content -LiteralPath (Join-Path $realLast 'SmartFinOps_Workplace-Valuation.current.json.txt') -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
  Check ($uploadExit -eq 3 -and $null -ne $uploadReceipt -and $uploadReceipt.Status -ceq 'Completed' -and $uploadReceipt.TenantKey -ceq 'contoso-prod') "real SmartM365 context: paths and TenantKey of profile prod, local outputs published, exit 3 (got $uploadExit)"
  Check ($upload.Contains("SharePoint upload enabled (SmartM365 context of profile 'prod'): contoso.sharepoint.invalid/sites/finops, library 'Documents', data root 'SMART-M365/DATA' (DATA-LAST, DATA-ALL, LOG-ALL).")) 'upload target from the prod profile, SmartM365 data root convention (SMART-M365/CSV -> SMART-M365/DATA)'
  Check (-not $upload.Contains('Tenant profile mismatch') -and -not $upload.Contains('Graph app-only connection values are missing') -and -not $upload.Contains('Created required local configuration') -and -not $upload.Contains('Created missing global local configuration')) 'no test profile and no missing Graph identity: the SmartM365 helpers use the prod identity'
  Check ($upload.Contains('SharePoint upload failed for') -and $upload.Contains('the completed receipt is not uploaded')) 'upload failure (synthetic identity): warnings, completed receipt withheld from SharePoint'
  $report = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $realReport -Tenant prod 2>&1 | Out-String
  $reportExit = $LASTEXITCODE
  Check ($reportExit -eq 3 -and (Test-Path -LiteralPath (Join-Path $realLast 'SmartFinOps_Workplace-Report.html')) -and $report.Contains("SmartM365 context of profile 'prod'") -and -not $report.Contains('Tenant profile mismatch') -and -not $report.Contains('Graph app-only connection values are missing')) "report with the real SmartM365 context: local page, upload attempted with the prod identity, exit 3 (got $reportExit)"
  $jsonAfter = & $runtimeJsonFiles
  Check ((($jsonBefore -join '|') -ceq ($jsonAfter -join '|')) -and -not (Test-Path -LiteralPath (Join-Path $realSmartM365 'Config\Tenants\test.local.json')) -and -not (Test-Path -LiteralPath (Join-Path $realSmartM365 'Config\Tenants\test.local.json.txt'))) "no SmartM365 runtime JSON created (before: $($jsonBefore -join ', '); after: $($jsonAfter -join ', '))"
  $zeroIdentity = Get-Content -LiteralPath $prodProfilePath -Raw | ConvertFrom-Json
  $zeroIdentity.AppId = '00000000-0000-0000-0000-000000000000'
  $zeroIdentity | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $prodProfilePath -Encoding utf8NoBOM
  $incomplete = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $realEngine -Tenant prod 2>&1 | Out-String
  Check ($LASTEXITCODE -eq 3 -and $incomplete.Contains('SharePoint upload enabled but the SmartM365 configuration is incomplete (missing: AppId)')) 'placeholder SmartM365 identity: upload skipped with a warning, exit 3'
  $missingProfile = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $realEngine -Tenant prod2 2>&1 | Out-String
  $missingExit = $LASTEXITCODE
  Check ($missingExit -eq 1 -and $missingProfile.Contains('SmartM365 runtime configuration not found') -and $missingProfile.Contains('never creates SmartM365 runtime JSON') -and -not (Test-Path -LiteralPath (Join-Path $realSmartM365 'Config\Tenants\prod2.local.json.txt')) -and -not (Test-Path -LiteralPath (Join-Path $realSmartM365 'Config\Tenants\prod2.local.json'))) "a profile without SmartM365 runtime JSON is refused and nothing is created (exit $missingExit)"
  $realContext = Join-Path $real 'SmartFinOps\Config\SmartFinOps-TenantContext.ps1'
  $otherCommand = "try { . '$realContext'; Initialize-SmartFinOpsTenantContext -Tenant 'prod' -StartPath '$(Split-Path $realContext -Parent)' | Out-Null; `$global:SmartM365ProfileKey = 'test'; Get-SmartFinOpsSharePointConfiguration -TenantKey 'contoso-prod' | Out-Null; 'OK|' } catch { 'ERROR|' + `$_.Exception.Message }"
  $other = [string](@(& $pwsh -NoProfile -Command $otherCommand 2>&1 | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(OK|ERROR)\|' }) | Select-Object -Last 1)
  Check ($other -like "ERROR|*A SmartM365 context of profile 'test' is already loaded*") "a SmartM365 context of another profile is refused ($other)"
  '{ "EnableSharePointUpload": "maybe" }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $invalid = Invoke-Script $engine @('-SmartM365LatestCsvFolderPath', $dataLast, '-SmartM365DataAllRootPath', $dataAll, '-SmartM365LogAllRootPath', $logAll, '-SmartM365TenantKey', 'contoso-prod', '-ContractsPath', (Join-Path $root 'contracts-sp.local.json'))
  Check ($invalid.ExitCode -eq 1 -and $invalid.Log.Contains('EnableSharePointUpload must be true, false or __USE_SMARTM365__')) 'an invalid EnableSharePointUpload value is refused'
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  #endregion
}
finally {
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartFinOps context tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartFinOps context tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCTWIiCF6Yz0FMr
# 1huwpfMdag4yxlonJjITmUMAfJf+e6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPCVZyBb4cPusvEuQKC6KHhA0nZsx7XBqTnhzVMW/svGMA0GCSqG
# SIb3DQEBAQUABIIBgAB4OmqPKls4bTO726JcGx6OoTC0ii+A7/T0H4oiPQj5kTq9
# MhyDqAig0L3Wp0xBYkojDEtjh+JUvZiIvCObQSiRd5b2qwaixoHjPJkgA9/k9ZbR
# 0dQPhopFgoxgyBtSFOHlFHHOzDQfmVdOEBZaG791/IMOx1p0NNQETzUZmIw5Q79H
# NDmUkoDaS0LtBany6EhCQ/aJvzsVaDxaIoYXxIuKiazNg0YVq05TNKzOk+hYgmbi
# PA9Ru3U+peRXi42NJ5KXa3PMVxVqAYJwRi2w7uRdqNO/sk8VAkFpf5Ki/BRymyxc
# PWf5EIhwSLtCJ5vsamwO7NdqNJaDJHqe5WwJj+x9j3FvHtWHbwkn4dUmk8jGZO9r
# GSJLaIdIQjSzjotR6yqc05hndp7x2uzbo3Z0pqHIm4+AKplRGdPCNTdQktCkpWHa
# onfy68wO7fuFqnGGjgf2iaOmk0VuTdG7LHuq7byQzzKK+esxEivbBzmieJND7rBq
# u/i1hpMfl0KtOzEAnqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkwMDUw
# NTRaMC8GCSqGSIb3DQEJBDEiBCCvqTO00FUYR11NoobpF+Uc9JMKMt1pdSJ9aRec
# 85/q7jANBgkqhkiG9w0BAQEFAASCAgBwVcod4BWRUc5RamQPk2rXA1fFUuo0N1Va
# psQ3WFyXz1ESleDCI6ewTBOwq2VgpAhUkzEWNnN2xnriokBbOfvC3QetW+vPM01v
# LjFcGdOYv1mhMPnawNjH+0SSHr6sBL+GOWsUy4Upv6QVKvfc1Frn3CA4mzGvigGv
# 4LEWIvfCHEUp5u3o6VfDiwEbhL+wkFJdHVT538vlJGyB+4fQPXMMSOPPb66mWcPF
# /gBAouiYetBZ2JFFZsmRXcYGuqTH/4KQUN9dmO2ztrPNhulwpclFpkbRhNuGIgVW
# lmcYr8dCb05brAS3WQqx1fZWmvKWPOhf9L4Wo84feT3AEILvCAPJTMvXWJCSVnfy
# mVtbai/XT99inCmMqVmf2zOwwWoj8k5V2k3Yf6o+ft7YYtU+ZC7HukOegWM5hVge
# maWaD1BX8nLI0A67CLQE9Bx1jAYDhtXoJCnLj5zzt6L391kKr73r3gCF3i1haIp4
# TGPqBGuD+VcbHA8KuzCKIRew/aSXq2zXCnZmYJJ0/kQgTGFOGYZqnOQlssUFJQue
# rEkgJnnVo/tH8Zf2ofHMyRxZrfKt0K3a5+ssrBKGgCAVMaHm+7+/CRy4G8LXl2WC
# dd1UEbsi7m5Q+K6X5WpuIwS7f2uMm/IO1GJLK6VlL8BwHUwv0xN9RDWP37Vp9CwA
# isSaqik2uA==
# SIG # End signature block
