#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps tenant context: SmartM365 DATA-LAST resolution and console banners.
.VERSION
1.1.0
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
  $run = Invoke-Script $engine @('-OutputRoot', (Join-Path $root 'out1'), '-ContractsPath', (Join-Path $root 'contracts1.local.json'))
  Check ($run.ExitCode -eq 1 -and $run.Log.Contains('Run failed: SmartM365LatestCsvFolderPath is __USE_SMARTM365__ but its resolution through SmartM365 failed') -and $run.Log.Contains('boom from the SmartM365 configuration')) 'engine fails with the SmartM365 resolution reason'
  Check (-not $run.Log.Contains('DATA-LAST folder not found') -and -not ($run.Log -match 'SmartM365\\Data\\Tenants')) 'no silent fallback to the default DATA-LAST path'
  Check ($run.Log.Contains(' SmartFinOps by WorkplaceCloudHub') -and $run.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $run.Log.Contains(' Status    : FAILED') -and $run.Log.Contains(' Script    : SmartFinOps-Workplace-Valuation.ps1') -and $run.Log.Contains(' CSV files : 0') -and $run.Log -match ' Errors    : [1-9]') 'engine failure: startup banner and Failed completion banner'
  $explicit = Invoke-Script $engine @('-SmartM365LatestCsvFolderPath', $dataLast, '-SmartM365TenantKey', 'contoso-prod', '-OutputRoot', (Join-Path $root 'out2'), '-ContractsPath', (Join-Path $root 'contracts2.local.json'), '-ValidateOnly')
  Check ($explicit.ExitCode -eq 0 -and -not $explicit.Log.Contains('resolution through SmartM365 failed') -and $explicit.Log.Contains(' Status    : ')) 'explicit -SmartM365LatestCsvFolderPath and -SmartM365TenantKey never call the SmartM365 context'
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
  $reportRun = Invoke-Script $report @('-InputRoot', (Join-Path $root '{{NoSuchToken}}'))
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
  [pscustomobject]@{ WorkspaceRootPath='$uncRoot'; LatestCsvFolderPath='{{WorkspaceRootPath}}\DATA-LAST'; ProfileKey=`$TenantKey; TenantKey='contoso-prod' }
}
"@
  '{ }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $cwd = Join-Path $root 'cwd'
  New-Item -ItemType Directory -Path $cwd -Force | Out-Null
  $dataRoot = Join-Path $runtime 'SmartFinOps\Data\contoso'
  Push-Location -LiteralPath $cwd
  try {
    $uncEngine = & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1') -Tenant $tenant 2>&1 | Out-String
    $uncEngineExit = $LASTEXITCODE
    $uncReport = & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Report.ps1') -Tenant $tenant 2>&1 | Out-String
    $uncReportExit = $LASTEXITCODE
  }
  finally { Pop-Location }
  Check ($uncEngineExit -eq 0 -and $uncEngine.Contains(('Sources: {0}\DATA-LAST' -f $uncRoot)) -and $uncEngine.Contains(('Outputs: {0}\SmartFinOps\Data\contoso' -f $uncRuntime))) "engine from UNC: DATA-LAST resolved through SmartM365, outputs on the share (exit $uncEngineExit) $uncEngine"
  $uncReceipt = Get-Content -LiteralPath (Join-Path $dataRoot 'SmartFinOps_Workplace-Valuation.current.json.txt') -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
  Check ($null -ne $uncReceipt -and $uncReceipt.ProfileKey -ceq $tenant -and $uncReceipt.TenantKey -ceq 'contoso-prod' -and $uncEngine.Contains("Profile 'contoso' - TenantKey 'contoso-prod'")) 'engine from UNC: Data\<profile> with the TenantKey of the profile in the receipt'
  Check ((Test-Path -LiteralPath (Join-Path $dataRoot 'SmartFinOps_Workplace-Valuation.current.json.txt')) -and @(Get-ChildItem -LiteralPath $dataRoot -Filter 'FinOps_*.csv' -ErrorAction SilentlyContinue).Count -eq 7) 'engine from UNC: receipt and seven CSVs written under the UNC data root'
  Check ($uncReportExit -eq 0 -and (Test-Path -LiteralPath (Join-Path $dataRoot 'SmartFinOps_Workplace-Report.html')) -and $uncReport.Contains(('Inputs: {0}\SmartFinOps\Data\contoso' -f $uncRuntime))) "report from UNC: page written under the UNC data root (exit $uncReportExit) $uncReport"
  $written = @(Get-ChildItem -LiteralPath $dataRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.txt', '.log', '.csv', '.html' } | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw })
  Check (-not ($uncEngine + $uncReport + ($written -join '')).Contains('FileSystem::')) 'no provider-qualified path in console, logs or outputs'
  Check (@(Get-ChildItem -LiteralPath $cwd -Recurse -Force).Count -eq 0) 'nothing written relative to the process working directory'
  # A relative -OutputRoot is taken from the PowerShell location (here on the share), not from the process directory.
  $relative = Join-Path $root 'relative'
  New-Item -ItemType Directory -Path $relative -Force | Out-Null
  Push-Location -LiteralPath $cwd
  try {
    $command = "Set-Location -LiteralPath '$(& $toUnc $relative)'; & '$(Join-Path $uncRuntime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1')' -Tenant '$tenant' -OutputRoot 'relout'; exit `$LASTEXITCODE"
    $relativeRun = & $pwsh -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1 | Out-String
    $relativeExit = $LASTEXITCODE
  }
  finally { Pop-Location }
  Check ($relativeExit -eq 0 -and (Test-Path -LiteralPath (Join-Path $relative 'relout\SmartFinOps_Workplace-Valuation.current.json.txt')) -and @(Get-ChildItem -LiteralPath $cwd -Recurse -Force).Count -eq 0) "relative -OutputRoot resolved from the PowerShell location on the share (exit $relativeExit) $relativeRun"
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDzInr19zqmQ1L0
# 4DugP2U7gPAvdTFyVPuwBsxLIvOubqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO9eimfm9CoUqO8a7WhcbfkGQbHZBWHbWD0fssqWfQ/2MA0GCSqG
# SIb3DQEBAQUABIIBgCtm7V0S4xvVbEfMSZNEJBjq7fwawBDY7EaUnbMjapJZ3cWR
# rQLLI2y3FJSyFFvLVRzytzwv6u/Z8IbobYSyFJ8GUiLsLJOvE5kflz1+jZNeG9zC
# 53ydj0Puu+Qtt/lIHWd4YsLkKmfBs8Tp+1vd1AtYCuUxiqM3JJ3AnOu8Q8wGukFa
# 5xFgVPFQUuXMkdx04fk2qJZGinCBOMZoblGy4H9UPnUWZRi1GTm08LnbCZdLLeUC
# ZTXkSfGxjEOPwUG5vLg1pR17UQWmVw1+f2IRuGN64LOYkZClvCHEZBiDJF0tIWvq
# +SYQ/UMhAfpMJ41TsHUHPuZrLU8LtktCo9nCA6VDye22Gn3IbPA7G4AGNNxigOgJ
# 9uiW8yk2qvqpkwQUhGMAJIrw0QScnB+a21/rav1IkLt7JMGTQXggfMpsKNIy27Mw
# 0anSCvkLNqQZWbf3KRUXySCe7NvBYcoNe0ymguT8dxx1UzKYDOIGY3nGIWFzvwxr
# 9Nunn1dmPgDT792j1aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMDQ5
# MTlaMC8GCSqGSIb3DQEJBDEiBCDrdJt4qDvDKMVSjnx1+y/xV6brny3yp00QidXJ
# RI7KSDANBgkqhkiG9w0BAQEFAASCAgA16oK0/UqI+WSekeGLWGT2G26DqBPl0+Fe
# wEAd/YcWguzhcMXT26iwUFzPUupr+jeV2ZuNgTapJvvfjEyMJiyettCrZNTo25OX
# PDbGBsq7NAk300tJzcaEHxRegHCGu5AvphAYWVje+KV2bYhWJCupgw8tHS3X9YZ4
# 0/Si9gtPHoWJ6kEKXHH2jOWZNVnLafh6DmOQ+Gl3iHyOiV8zh3ifhk2Whdlk/4Ij
# WxNmD5Zbl1fJnTFt+i4OTSoQZzEa0yxVKl7uxTk/9grDfldItwTPPlIDJr6Mo30X
# 93On03qEa5shY5ECtLabwS1OM7vEKaxFO2Ay5MMAoGmSSfkrG/rgvZJ/cbXyJPwL
# vZ1mDW/Vg4X7EuWaVL8taz/j0lksTqcH8TjprJWxaaOpFDQr05F9WH4SLuVmMTy3
# NIK1CBtxMtYgGNqyftfgpNL1FsCH9ZDTQ16nxHkiZfMItHxElgizrMaSC6mUPjOZ
# ULWOGxb3UqoNoysDYi7zjw8bzi6SVCoHFBxHnXL0NuPzRpnJvnKP1wSosu6woCjI
# VZK2APOR9sDUWKDZXdGMLPCGNfLRVqnGtZuMv6PlJJO5w7oKuqq2X1+l6cFx2pOD
# 344GOY5i/a1EDr1ZNevZTy7k2jwdn3KFPkBEQLY1NG3KGst44xyni0hZCPA7wfNT
# M2OzJNJhfw==
# SIG # End signature block
