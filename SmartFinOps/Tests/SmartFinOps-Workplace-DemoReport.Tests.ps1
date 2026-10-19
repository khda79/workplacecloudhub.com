#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps Workplace demo report generator.
.VERSION
1.0.0
.NOTES
  Runs Demo\SmartFinOps-Workplace-DemoReport.ps1 (real engine and report on synthetic contoso data) into a temporary
  folder, then checks the page, its manifest, the DEMO banner, the statuses it must show and the leak scanner.
  No tenant, export, SharePoint upload or mail is involved.
#>
[CmdletBinding()]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$generator = Join-Path $project 'Demo\SmartFinOps-Workplace-DemoReport.ps1'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-DemoTest-' + [guid]::NewGuid().ToString('N'))
$pwsh = (Get-Process -Id $PID).Path
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

try {
  #region 1. Generation with the real engine and report
  $out = Join-Path $root 'out'
  $workspacesBefore = @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'SmartFinOps-Demo-*' | ForEach-Object Name)
  $log = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $generator -OutputPath $out -SourceCommit 'test-commit' 2>&1 | Out-String
  $exit = $LASTEXITCODE
  Check ($exit -eq 0) "generator exit code 0 (got $exit) $log"
  $page = Join-Path $out 'SmartFinOps_Workplace-Report.demo.html'
  $manifestPath = Join-Path $out 'SmartFinOps_Workplace-Report.demo.manifest.json'
  Check ((Test-Path -LiteralPath $page) -and (Test-Path -LiteralPath $manifestPath)) 'page and manifest written to -OutputPath'
  Check (@(Get-ChildItem -LiteralPath $out -Force).Count -eq 2) 'nothing else written to -OutputPath'
  Check (@($log -split "`r?`n" | Where-Object { $_ -and $_ -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[' }).Count -eq 0) 'every generator log line is timestamped'
  Check (@(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'SmartFinOps-Demo-*' | Where-Object Name -notin $workspacesBefore).Count -eq 0) 'temporary workspace removed'
  #endregion

  if (Test-Path -LiteralPath $page) {
    $html = [IO.File]::ReadAllText($page)
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

    #region 2. Manifest
    Check ($manifest.Owner -eq 'SmartFinOps-DemoReport' -and $manifest.TenantKey -eq 'contoso-demo' -and $manifest.LeakScan -eq 'Passed' -and $manifest.SourceCommit -eq 'test-commit') 'manifest identity, tenant key, leak scan and source commit'
    Check ($manifest.ValuationVersion -match '^\d+\.\d+\.\d+$' -and $manifest.ReportVersion -match '^\d+\.\d+\.\d+$' -and $manifest.GeneratorVersion -eq '1.0.0' -and $manifest.FixtureVersion -eq '1.0.0') 'manifest versions'
    Check (@($manifest.ValuationExitCodes).Count -eq 2 -and @($manifest.ValuationExitCodes | Where-Object { $_ -notin @(0,3) }).Count -eq 0 -and $manifest.ReportExitCode -in @(0,3)) 'two valuation runs and the report ended with 0 or 3'
    Check ($manifest.ReportSHA256 -eq (Get-FileHash -LiteralPath $page -Algorithm SHA256).Hash) 'manifest hash matches the page'
    Check (-not ((Get-Content -LiteralPath $manifestPath -Raw) -match '(?i)[a-z]:\\|\\\\|temp')) 'manifest holds no local path'
    #endregion

    #region 3. Page content: DEMO label and the statuses the demo must show
    Check ([regex]::Matches($html, '<p class="demo-banner" role="note"><strong>DEMO</strong>').Count -eq 1 -and $html.Contains('<title>DEMO - ')) 'one DEMO banner and a DEMO title'
    Check ($html.Contains('contoso.invalid') -and $html.Contains("Valuation $($manifest.ValuationVersion) and Report $($manifest.ReportVersion)")) 'banner names the fictional organization and the versions'
    Check ($html -notmatch '(?i)Report not available') 'figures published (report available)'
    foreach ($status in @('Provisional','Not valued','Review','Blocked','Capped','Not measured','Realized','Reused')) { Check ($html.Contains($status)) "page shows status $status" }
    Check ($html.Contains('AD domains not collected: west')) 'partial AD scope shown'
    Check ($html -match 'enabled accounts without M365 activity in 90 days' -and $html -match 'E3 to F3 \(not in the R total\)') 'enabled inactive accounts and conditional E3 to F3 shown separately'
    Check ($html -notmatch '(?i)<script|https?://(?!www\.w3\.org/2000/svg)') 'no script and no external resource'
    #endregion
  }

  #region 4. Leak scanner
  . $generator
  $clean = '<html><head><title>x</title><style>p{}</style></head><body><p>north.contoso.invalid 2.4.1 12:30:45 contoso-demo https://workplacecloudhub.com</p><svg xmlns="http://www.w3.org/2000/svg"></svg></body></html>'
  Check (@(Test-DemoLeak -Html $clean).Count -eq 0) 'clean page passes the leak scan'
  $cases = [ordered]@{
    'e-mail address or UPN' = 'contact user1@fabrikam.invalid'
    'GUID' = '11111111-2222-3333-4444-555555555555'
    'IPv4 address' = 'host 10.20.30.40'
    'IPv6 address' = 'fe80:0:0:0:200:f8ff:fe21:67cf'
    'Windows path' = 'D:\Exports\file.csv'
    'UNC path' = '\\server01\share'
    'profile folder' = 'C\Users\someone'
    'script element' = '<script>x</script>'
    'foreign domain' = 'mail.fabrikam.com'
    'external URL' = 'see https://cdn.fabrikam.invalid/x.js'
    'tenant key other than' = 'tenant fabrikam-prod'
  }
  foreach ($case in $cases.GetEnumerator()) {
    $findings = @(Test-DemoLeak -Html ($clean.Replace('</body>', "<p>$($case.Value)</p></body>")))
    Check (@($findings | Where-Object { $_.StartsWith($case.Key) }).Count -ge 1) "leak scan detects: $($case.Key) (got $($findings -join ' | '))"
  }
  $private = @(Test-DemoLeak -Html $clean.Replace('</body>', '<p>Acme Internal</p></body>') -PrivateTerms @('first term','acme internal'))
  Check ($private.Count -eq 1 -and $private[0] -eq 'private term (line 2 of the private term list)' -and $private[0] -notmatch '(?i)acme') 'private term reported by line number, never by value'
  $local = @(Test-DemoLeak -Html $clean.Replace('</body>', '<p>WORKSTATION42</p></body>') -LocalValues @('workstation42'))
  Check ($local.Count -eq 1 -and $local[0] -notmatch '(?i)workstation') 'local computer value detected without echoing it'
  $bannered = Add-DemoBanner -Html $clean -Text 'Fictional & synthetic'
  Check ($bannered.Contains('<body>') -and $bannered.Contains('<strong>DEMO</strong>Fictional &amp; synthetic') -and $bannered.Contains('<title>DEMO - x</title>') -and $bannered.Contains('.demo-banner{')) 'banner inserted after <body>, encoded, CSS in the existing style element, title prefixed'
  $refused = $false
  $insideLog = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $generator -OutputPath (Join-Path $project 'Demo\out') 2>&1 | Out-String
  $refused = $LASTEXITCODE -eq 1 -and $insideLog.Contains('outside the SmartFinOps code folder') -and -not (Test-Path -LiteralPath (Join-Path $project 'Demo\out'))
  Check $refused 'output inside the SmartFinOps code folder refused'
  #endregion
}
finally {
  if ($KeepEvidence) { Write-Host "Evidence kept: $root" }
  elseif (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($script:failures.Count -gt 0) {
  Write-Host ("SmartFinOps demo report tests: {0} failure(s) in {1} checks." -f $script:failures.Count, $script:checks) -ForegroundColor Red
  exit 1
}
Write-Host ("SmartFinOps demo report tests: PASS ({0} checks)." -f $script:checks) -ForegroundColor Green
exit 0
