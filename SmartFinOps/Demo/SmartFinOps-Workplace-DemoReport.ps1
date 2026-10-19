#Requires -Version 7.0
<#
.SYNOPSIS
  Builds the SmartFinOps Workplace demo report from synthetic data with the real valuation engine and report.
.DESCRIPTION
  1. Creates an isolated temporary runtime: the V2 engine, report, tenant context and templates, plus the tracked
     SmartM365.Core module files (no local configuration, export, profile or tenant is copied or read).
  2. Writes a fictional SmartInventory DATA-LAST (SmartFinOps-Workplace-DemoFixture.ps1, contoso.invalid) with real
     SmartM365.Core source receipts, in two stages: a baseline plateau, then the current day after a reduction.
  3. Runs SmartFinOps-Workplace-Valuation.ps1 on each stage, then SmartFinOps-Workplace-Report.ps1, unchanged, with
     SharePoint upload and the daily email disabled. Exit codes 0 and 3 are accepted (the demo includes provisional
     amounts); 1 or a "Report not available" page fails the run.
  4. Adds a visible DEMO banner, scans the page for leaks (e-mail addresses, GUIDs, IP addresses, file-system paths,
     host or user names of this computer, private terms, foreign domains or URLs) and writes the page and a manifest
     to -OutputPath.
.VERSION
1.0.0
.NOTES
  Read-only and offline: no API call, no SharePoint upload, no mail. -OutputPath must lie outside the SmartFinOps
  code folder. The page and its manifest are generated files: never commit them.
  Exit code: 0 page written and leak scan passed, 1 failure.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='SmartM365.Core receipt identity globals are saved and restored in finally.')]
param(
  [string]$OutputPath,
  [string]$SourceCommit = '',
  [string]$ForbiddenTermsPath = '',
  [switch]$KeepWorkspace
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptVersion = '1.0.0'
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
$script:DemoTenantKey = 'contoso-demo'
$script:DemoProfileKey = 'demo'
$script:ReportFileName = 'SmartFinOps_Workplace-Report.demo.html'
$script:ManifestFileName = 'SmartFinOps_Workplace-Report.demo.manifest.json'

function Write-DemoLog {
  param([Parameter(Mandatory)][string]$Message, [ValidateSet('INFO','WARNING','ERROR')][string]$Level = 'INFO')
  Write-Host ('{0} [{1}] {2}' -f [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant), $Level, $Message)
}

function Get-DemoScriptVersion {
  param([Parameter(Mandatory)][string]$Path)
  $match = [regex]::Match([IO.File]::ReadAllText($Path), "(?m)^\`$script:ScriptVersion\s*=\s*'([^']+)'")
  if (-not $match.Success) { throw "Script version not found in $([IO.Path]::GetFileName($Path))." }
  return $match.Groups[1].Value
}

function Copy-DemoCoreModule {
  # Only the tracked module files: local configuration, transport locks and logs next to the module stay out.
  param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
  foreach ($file in Get-ChildItem -LiteralPath $Source -File -Recurse) {
    if ($file.Name -like '*.local.json*' -or $file.Extension -in @('.lock','.log')) { continue }
    $target = Join-Path $Destination ([IO.Path]::GetRelativePath($Source, $file.FullName))
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $target
  }
}

function Invoke-DemoScript {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string[]]$Arguments)
  $pwsh = (Get-Process -Id $PID).Path
  $output = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $Path @Arguments 2>&1 | Out-String
  return [pscustomobject]@{ ExitCode=$LASTEXITCODE; Output=$output }
}

function Get-DemoPrivateTermList {
  param([string]$Path, [string]$RepositoryRoot)
  $candidate = $Path
  if (-not $candidate) {
    $gitDir = & git -C $RepositoryRoot rev-parse --git-common-dir 2>$null
    if ($LASTEXITCODE -eq 0 -and $gitDir) {
      $gitDir = [string]$gitDir
      if (-not [IO.Path]::IsPathRooted($gitDir)) { $gitDir = Join-Path $RepositoryRoot $gitDir }
      $candidate = Join-Path $gitDir 'forbidden-terms.txt'
    }
  }
  if (-not $candidate -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) { return @() }
  return @([IO.File]::ReadAllLines($candidate) | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
}

function Test-DemoLeak {
  <#
    Returns the leak findings of a page as "rule: detail" strings. Allow-list approach: the only domain is
    contoso.invalid, the only URLs are the SVG namespace and the WorkplaceCloudHub website.
    Private terms are reported by their line number only, never by value.
  #>
  param([Parameter(Mandatory)][string]$Html, [string[]]$PrivateTerms = @(), [string[]]$LocalValues = @())
  $findings = [System.Collections.Generic.List[string]]::new()
  $rules = [ordered]@{
    'e-mail address or UPN' = '[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+'
    'GUID' = '\b[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\b'
    'IPv4 address' = '(?<![\d.])(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}(?![\d.])'
    'IPv6 address' = '(?i)(?<![0-9a-f:])(?:[0-9a-f]{1,4}:){3,7}[0-9a-f]{1,4}(?![0-9a-f:])|(?i)\b[0-9a-f]{1,4}::[0-9a-f]{1,4}\b'
    'Windows path' = '(?i)(?<![a-z])[a-z]:\\'
    'UNC path' = '\\\\[A-Za-z0-9_.$-]+\\'
    'profile folder' = '(?i)\\Users\\|AppData|/home/|/Users/'
    'script element' = '(?i)<script'
  }
  foreach ($rule in $rules.GetEnumerator()) {
    $hits = [regex]::Matches($Html, $rule.Value)
    if ($hits.Count -gt 0) { $findings.Add(('{0}: {1} occurrence(s), first "{2}"' -f $rule.Key, $hits.Count, $hits[0].Value)) }
  }
  foreach ($match in [regex]::Matches($Html, '(?i)\b(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+(?:com|net|org|fr|de|es|it|eu|io|uk|ch|be|local|lan|corp|internal|intra|ad|onmicrosoft|sharepoint|invalid|example|test)\b')) {
    $domain = $match.Value.ToLowerInvariant()
    if ($domain -eq 'contoso.invalid' -or $domain.EndsWith('.contoso.invalid') -or $domain -in @('www.w3.org','workplacecloudhub.com')) { continue }
    $findings.Add("foreign domain: $domain"); break
  }
  foreach ($match in [regex]::Matches($Html, '(?i)https?://[^\s"''<>)]+')) {
    $url = $match.Value
    if ($url -ceq 'http://www.w3.org/2000/svg' -or $url -like 'https://workplacecloudhub.com*') { continue }
    $findings.Add("external URL: $url"); break
  }
  foreach ($key in [regex]::Matches($Html, '(?i)\b[a-z0-9]+-(?:prod|test|dev|qa|demo|validation)\b') | ForEach-Object { $_.Value.ToLowerInvariant() } | Sort-Object -Unique) {
    if ($key -ne $script:DemoTenantKey) { $findings.Add("tenant key other than ${script:DemoTenantKey}: $key") }
  }
  foreach ($value in $LocalValues | Where-Object { $_ -and $_.Length -ge 3 } | Sort-Object -Unique) {
    if ($Html.IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $findings.Add('local value of this computer (user, host, domain or workspace path)') }
  }
  for ($index = 0; $index -lt $PrivateTerms.Count; $index++) {
    if ($Html.IndexOf($PrivateTerms[$index], [StringComparison]::OrdinalIgnoreCase) -ge 0) { $findings.Add("private term (line $($index + 1) of the private term list)") }
  }
  return @($findings)
}

function Add-DemoBanner {
  # Visible DEMO label at the top of the page and in its title; the rest of the generated page is unchanged.
  param([Parameter(Mandatory)][string]$Html, [Parameter(Mandatory)][string]$Text)
  $css = '.demo-banner{margin:0;padding:10px 16px;background:#fff4d6;color:#5c4100;border-bottom:2px solid #e0a100;font:600 14px/1.45 "Segoe UI",Arial,sans-serif}.demo-banner strong{margin-right:8px;padding:1px 8px;border-radius:4px;background:#e0a100;color:#1c1300;letter-spacing:.04em}'
  $styleEnd = $Html.IndexOf('</style>', [StringComparison]::OrdinalIgnoreCase)
  if ($styleEnd -ge 0) { $Html = $Html.Insert($styleEnd, $css) }
  else {
    $headEnd = $Html.IndexOf('</head>', [StringComparison]::OrdinalIgnoreCase)
    if ($headEnd -lt 0) { throw 'The report page has no </head> element.' }
    $Html = $Html.Insert($headEnd, "<style>$css</style>")
  }
  $body = [regex]::Match($Html, '(?i)<body[^>]*>')
  if (-not $body.Success) { throw 'The report page has no <body> element.' }
  $Html = $Html.Insert($body.Index + $body.Length, "`n<p class=""demo-banner"" role=""note""><strong>DEMO</strong>$([Net.WebUtility]::HtmlEncode($Text))</p>")
  return [regex]::Replace($Html, '(?i)<title>', '<title>DEMO - ', 1)
}

if ($MyInvocation.InvocationName -eq '.') { return }

$project = Split-Path $PSScriptRoot -Parent
$repository = Split-Path $project -Parent
$workspace = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-Demo-' + [guid]::NewGuid().ToString('N'))
$savedGlobals = @{}
$exitCode = 1
try {
  Write-DemoLog "Starting SmartFinOps Workplace demo report v$($script:ScriptVersion)"
  if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw '-OutputPath is required.' }
  $outputFull = [IO.Path]::GetFullPath($OutputPath)
  $projectFull = [IO.Path]::GetFullPath($project).TrimEnd('\') + '\'
  if (($outputFull.TrimEnd('\') + '\').StartsWith($projectFull, [StringComparison]::OrdinalIgnoreCase)) { throw '-OutputPath must lie outside the SmartFinOps code folder.' }

  # 1. Isolated runtime.
  $runtime = Join-Path $workspace 'runtime'
  foreach ($file in @('SmartFinOps-Workplace-Valuation.ps1','SmartFinOps-Workplace-Report.ps1','SmartFinOps.global.local.json.template','Config\SmartFinOps-TenantContext.ps1','Config\SmartFinOps-Contracts.local.json.template')) {
    $target = Join-Path $runtime "SmartFinOps\$file"
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $project $file) -Destination $target
  }
  Copy-DemoCoreModule -Source (Join-Path $repository 'SmartM365\Modules\SmartM365.Core') -Destination (Join-Path $runtime 'SmartM365\Modules\SmartM365.Core')
  # Fixed TenantKey, no SharePoint mirror, no daily email: the SmartM365 context is never resolved.
  ([ordered]@{ SmartM365TenantKey=$script:DemoTenantKey; EnableSharePointUpload=$false; Report=[ordered]@{ ShowbackMinimumGroupUsers=5; EnableDailyEmail=$false } } | ConvertTo-Json -Depth 3) |
    Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $engine = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1'
  $report = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Report.ps1'
  $versions = [ordered]@{ Valuation=(Get-DemoScriptVersion $engine); Report=(Get-DemoScriptVersion $report) }
  Write-DemoLog "Runtime ready: Valuation $($versions.Valuation), Report $($versions.Report)."

  $producers = @{
    Licences = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1'
    Pricing = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-LicensePricing-Inventory.ps1'
    AD = Join-Path $repository 'SmartM365\SmartInventory\ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1'
    Users = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1'
  }
  foreach ($path in $producers.Values) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "SmartInventory producer script not found: $([IO.Path]::GetFileName($path))." } }

  # 2. Synthetic sources with real receipts (identity globals of SmartM365.Core, restored afterwards).
  foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $savedGlobals[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
  }
  Import-Module (Join-Path $runtime 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.90' -Force | Out-Null
  $global:SmartM365TenantKey = $script:DemoTenantKey; $global:SmartM365OrganizationKey = 'contoso'; $global:SmartM365EnvironmentKey = 'demo'; $global:SmartM365TenantId = 'contoso-demo-tenant'
  . (Join-Path $PSScriptRoot 'SmartFinOps-Workplace-DemoFixture.ps1')
  $contracts = Join-Path $workspace 'contracts\SmartFinOps-Contracts.local.json'
  New-SmartFinOpsDemoContract -Path $contracts
  $tenantRoot = Join-Path $workspace 'tenant'
  $roots = [ordered]@{ Last=(Join-Path $tenantRoot 'DATA-LAST'); All=(Join-Path $tenantRoot 'DATA-ALL'); Logs=(Join-Path $tenantRoot 'LOG-ALL') }
  New-Item -ItemType Directory -Path $roots.Values -Force | Out-Null
  $rootArguments = @('-Tenant',$script:DemoProfileKey,'-SmartM365LatestCsvFolderPath',$roots.Last,'-SmartM365DataAllRootPath',$roots.All,'-SmartM365LogAllRootPath',$roots.Logs)

  # 3. Valuation of the baseline plateau, then of the current day (realization against that baseline), then the report.
  $valuationExitCodes = [System.Collections.Generic.List[int]]::new()
  $offset = 1
  foreach ($stage in @('Baseline','Current')) {
    $fixture = New-SmartFinOpsDemoFixture -Folder (Join-Path $workspace "sources-$($stage.ToLowerInvariant())") -TenantKey $script:DemoTenantKey -Stage $stage -ProducerScripts $producers
    Get-ChildItem -LiteralPath $fixture.Folder -File | Copy-Item -Destination $roots.Last -Force
    $asOf = [datetime]::UtcNow.AddMinutes($offset).ToString('o', $script:Invariant); $offset++
    $run = Invoke-DemoScript -Path $engine -Arguments ($rootArguments + @('-ContractsPath',$contracts,'-AsOfUtc',$asOf))
    $valuationExitCodes.Add($run.ExitCode)
    if ($run.ExitCode -notin @(0,3)) { throw "Valuation ($stage) ended with exit code $($run.ExitCode).`n$($run.Output)" }
    $receipt = Get-Content -LiteralPath (Join-Path $roots.Last 'SmartFinOps_Workplace-Valuation.current.json.txt') -Raw | ConvertFrom-Json
    if ($receipt.Status -ne 'Completed' -or -not $receipt.LicencesMeasured -or $receipt.SnapshotId -ne $fixture.SnapshotId) { throw "Valuation ($stage) receipt is not Completed for the demo snapshot." }
    $failedControls = @(Import-Csv -LiteralPath (Join-Path $roots.Last 'FinOps_Controls.csv') -Encoding utf8 | Where-Object Status -ne 'Pass')
    if ($failedControls.Count -gt 0) { throw "Valuation ($stage): $($failedControls.Count) equivalence control(s) did not pass." }
    Write-DemoLog "Valuation $stage completed (exit code $($run.ExitCode), run $($receipt.RunId))."
  }
  $reportRun = Invoke-DemoScript -Path $report -Arguments $rootArguments
  if ($reportRun.ExitCode -notin @(0,3)) { throw "Report ended with exit code $($reportRun.ExitCode).`n$($reportRun.Output)" }
  $html = [IO.File]::ReadAllText((Join-Path $roots.Last 'SmartFinOps_Workplace-Report.html'))
  if ($html -match '(?i)<h1>\s*Report not available') { throw 'The report page is "Report not available".' }
  Write-DemoLog "Report completed (exit code $($reportRun.ExitCode))."

  # 4. DEMO banner, leak scan, outputs.
  $bannerText = 'Fictional sample data for the organization contoso.invalid. No tenant, directory, mailbox or license was read. Built from synthetic SmartInventory exports by SmartFinOps Workplace Valuation {0} and Report {1}.' -f $versions.Valuation, $versions.Report
  $html = Add-DemoBanner -Html $html -Text $bannerText
  $privateTerms = Get-DemoPrivateTermList -Path $ForbiddenTermsPath -RepositoryRoot $repository
  $localValues = @($env:USERNAME, $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERDNSDOMAIN, $workspace, [IO.Path]::GetTempPath().TrimEnd('\'), $repository)
  $findings = @(Test-DemoLeak -Html $html -PrivateTerms $privateTerms -LocalValues $localValues)
  if ($findings.Count -gt 0) { throw "Leak scan failed:`n - $($findings -join "`n - ")" }
  Write-DemoLog "Leak scan passed ($($privateTerms.Count) private term(s) checked)."

  New-Item -ItemType Directory -Path $outputFull -Force | Out-Null
  $pagePath = Join-Path $outputFull $script:ReportFileName
  [IO.File]::WriteAllText($pagePath, $html, [Text.UTF8Encoding]::new($false))
  $manifest = [ordered]@{
    Owner='SmartFinOps-DemoReport'; GeneratorVersion=$script:ScriptVersion; FixtureVersion=$script:DemoFixtureVersion
    ValuationVersion=$versions.Valuation; ReportVersion=$versions.Report; SourceCommit=$SourceCommit
    GeneratedAtUtc=[datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', $script:Invariant); TenantKey=$script:DemoTenantKey
    Scenario='Baseline plateau then current day after a reduction; partial AD scope (domain west not collected)'
    ValuationExitCodes=@($valuationExitCodes); ReportExitCode=$reportRun.ExitCode
    ReportFile=$script:ReportFileName; ReportSHA256=(Get-FileHash -LiteralPath $pagePath -Algorithm SHA256).Hash
    LeakScan='Passed'; PrivateTermsChecked=$privateTerms.Count
  }
  ($manifest | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath (Join-Path $outputFull $script:ManifestFileName) -Encoding utf8NoBOM
  Write-DemoLog "Demo report written: $pagePath"
  $exitCode = 0
}
catch {
  Write-DemoLog -Level ERROR -Message ("{0} (line {1})" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber)
}
finally {
  foreach ($name in $savedGlobals.Keys) {
    if ($savedGlobals[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $savedGlobals[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if ($KeepWorkspace) { Write-DemoLog "Workspace kept: $workspace" }
  elseif (Test-Path -LiteralPath $workspace) { Remove-Item -LiteralPath $workspace -Recurse -Force -ErrorAction SilentlyContinue }
}
exit $exitCode

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD7kVe4zP6cDVql
# oenE7G8lFku7ZLPSrkUg+0Zj+Kj8rKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIgtg4KhH7Vgg7GNVzAh14dBS9XTTamLfhN+pIlPVTNiMA0GCSqG
# SIb3DQEBAQUABIIBgAL1atMTsy1MePcQtWrQRN9BwTGkN2GW1OkqMJqed43gF+p6
# PYX2sdZms7irO+w7A39YIvCGRVdvaYFUvOXKB5zCdQIUMC0u/o022J4mrYS1ahmJ
# qy65aeW+UQY6TyaEVkl5kQGKxM+F0dxiDhi0SBHE/We6amOZFhHYLXM2OupMaS7R
# /zkadeoRts563xkzGH9XENg9cROIFLXJKgoOXrevlrIi3xim6NUAxQ2YVrEISHGU
# g/v5ooXuCzmpCR2BkiQrnVmtuOZaHiLRqKjUosRkOI3MubWIjtF6Bz/1dg07Wdud
# 2VvBFR/qP6GuGstxE0njBj5NTCPU+FumR2z66DyaMir/MzmAvx2nVookWSW2C1H9
# NF97/iINh4HCtucIaOEyD9pL7pwlyCgMH5IlpQ+K/GbD1KPJ8l/JnqGSNv8JW7YV
# ZmPEONu6Ozf6fsulWteAAoF+wyaM65oyq7auXj/HcvLCf5IT8X/uuzNde6om6MuL
# yXfMrL4gg5Z9vlnpj6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxMjUz
# MDdaMC8GCSqGSIb3DQEJBDEiBCAcRqw9LDZen119iQOoZyE0OpPCZJq68JpHcUcn
# 3N1GrzANBgkqhkiG9w0BAQEFAASCAgCk1xSYV5Pt56jZdBCNHrtqgD4qOLchAd1g
# cg7d5/pxLt/ZM4Ix5xHAB1oI09nMBZrCh6noMR7Xd88rgLODlmtkM6UF/Vsg1+FM
# Xsm7sLGlx6BOVyMkeYHvHoKwxbGB+/NoDJ0Yva85cHy7CQjPoZiyHenvPD4tYkZB
# Iovqp6Bqj6OJmNogliCa0zPG5bKeyhpIECnlzI7JMxVYBP+8GB5LHrMsM2GjVaCz
# DJ+rPQaJQBRge0dQ7CKtPrYCviP0id4mgov1cPLWiRY5Mu8lplJy8lJ0R3hnxp42
# pt/6JNy0Zp/raIVPnqTcB9QxelmyFomH9OZTPgTxGEgF0qDr+7DGMS1GdvMgKXe2
# Uoui788kq1PPJNsf0hNGKV3cOT4OJJZU/uWJ7mTqpAM4IDQi3MuAu+kY/cBp6u1u
# IKb2DEE6ZS9udc6fcmHlRc24w4/lJM7ffJi7QXoO9CE8wu+1ti999rjpLvxDAMyE
# SNbg3tEgWps/5CNG20FgJ6VoONzu4GTWMgNE0MKvHql+KzI2bg5DZHegAImK7Dw+
# IQ+HyBTWJv43ySHhMIsrn8XrJZzR45hf1kXtfb+SEIKVJ/V7PzgQpVPdqZjTKO/A
# ocWJcjOyTOG26Ga7Ao2HBBwJ88/raaAiBwVly5QvR2n2bkdqxUBf+NFqM1rIH3N2
# inC9xKKA2w==
# SIG # End signature block
