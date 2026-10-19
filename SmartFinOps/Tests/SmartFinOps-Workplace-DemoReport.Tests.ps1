#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps Workplace demo report generator.
.VERSION
1.3.0
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
    Check ($manifest.ValuationVersion -match '^\d+\.\d+\.\d+$' -and $manifest.ReportVersion -match '^\d+\.\d+\.\d+$' -and $manifest.GeneratorVersion -eq '1.0.3' -and $manifest.FixtureVersion -eq '1.1.0') 'manifest versions'
    Check (@($manifest.ValuationExitCodes).Count -eq 2 -and @($manifest.ValuationExitCodes | Where-Object { $_ -notin @(0,3) }).Count -eq 0 -and $manifest.ReportExitCode -in @(0,3)) 'two valuation runs and the report ended with 0 or 3'
    Check ($manifest.ReportSHA256 -eq (Get-FileHash -LiteralPath $page -Algorithm SHA256).Hash) 'manifest hash matches the page'
    Check (-not ((Get-Content -LiteralPath $manifestPath -Raw) -match '(?i)[a-z]:\\|\\\\|temp')) 'manifest holds no local path'
    #endregion

    #region 3. Page content: DEMO label and the statuses the demo must show
    Check ([regex]::Matches($html, '<p class="demo-banner" role="note"><strong>DEMO</strong>').Count -eq 1 -and $html.Contains('<title>DEMO - ')) 'one DEMO banner and a DEMO title'
    Check ($html.Contains('contoso.invalid') -and $html.Contains("Valuation $($manifest.ValuationVersion) and Report $($manifest.ReportVersion)")) 'banner names the fictional organization and the versions'
    Check ($html -notmatch '(?i)Report not available') 'figures published (report available)'
    foreach ($status in @('Final','Not valued','Review','Blocked','Capped','Not measured','Realized','Reused')) { Check ($html.Contains($status)) "page shows status $status" }
    # Complete source receipts (fixture 1.1.0): Final figures, no provisional source, no missing AD domain, every run exit 0.
    Check ($html.Contains('Final figures') -and -not $html.Contains('Provisional figures') -and $html.Contains('0 provisional')) 'figures are Final, no provisional source'
    Check (-not $html.Contains('AD domains not collected')) 'every AD domain collected'
    Check (@($manifest.ValuationExitCodes | Where-Object { $_ -ne 0 }).Count -eq 0 -and $manifest.ReportExitCode -eq 0) "valuation and report end with exit code 0 (got $(@($manifest.ValuationExitCodes) -join ',') / $($manifest.ReportExitCode))"
    Check ($html -match 'enabled accounts without M365 activity in 90 days' -and $html -match 'E3 to F3 \(not in the R total\)') 'enabled inactive accounts and conditional E3 to F3 shown separately'
    Check ($html -notmatch '(?i)<script|https?://(?!www\.w3\.org/2000/svg)') 'no script and no external resource'
    Check ([regex]::Matches($html, '(?i)<img\b').Count -eq 1 -and $html.Contains('<img alt="WorkplaceCloudHub" ') -and -not $html.Contains('class="client-logo"')) 'one image only, the WorkplaceCloudHub logo: never a client logo'
    Check ($log.Contains('Report run without client logo (-NoClientLogo confirmed).')) 'the generator forces the report without SmartM365 branding'
    #endregion
  }

  #region 4. Outside a Git repository and without -ForbiddenTermsPath: no private term list, generation still succeeds (1.0.0 failed here)
  $isolated = Join-Path $root 'no-git'
  $repository = Split-Path $project -Parent
  $needed = @(
    'SmartFinOps\Demo\SmartFinOps-Workplace-DemoReport.ps1', 'SmartFinOps\Demo\SmartFinOps-Workplace-DemoFixture.ps1'
    'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1', 'SmartFinOps\SmartFinOps-Workplace-Report.ps1', 'SmartFinOps\SmartFinOps.global.local.json.template'
    'SmartFinOps\Config\SmartFinOps-TenantContext.ps1', 'SmartFinOps\Config\SmartFinOps-Contracts.local.json.template'
    'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1', 'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-LicensePricing-Inventory.ps1'
    'SmartM365\SmartInventory\ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1', 'SmartM365\SmartInventory\M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1'
  )
  foreach ($relative in $needed) {
    $target = Join-Path $isolated $relative
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repository $relative) -Destination $target
  }
  # Tracked module files only, as the generator copies them (no local configuration, lock or log).
  $coreSource = Join-Path $repository 'SmartM365\Modules\SmartM365.Core'
  foreach ($file in Get-ChildItem -LiteralPath $coreSource -File -Recurse | Where-Object { $_.Name -notlike '*.local.json*' -and $_.Extension -notin @('.lock','.log') }) {
    $target = Join-Path (Join-Path $isolated 'SmartM365\Modules\SmartM365.Core') ([IO.Path]::GetRelativePath($coreSource, $file.FullName))
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $target
  }
  & git -C $isolated rev-parse --git-dir 2>$null | Out-Null
  Check ($LASTEXITCODE -ne 0) 'isolated copy is outside any Git repository'
  $noTermsOut = Join-Path $root 'out-no-terms'
  $noTermsLog = & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $isolated 'SmartFinOps\Demo\SmartFinOps-Workplace-DemoReport.ps1') -OutputPath $noTermsOut 2>&1 | Out-String
  $noTermsExit = $LASTEXITCODE
  Check ($noTermsExit -eq 0 -and $noTermsLog.Contains('Leak scan passed (0 private term(s) checked)')) "generation without a private term list succeeds (got $noTermsExit) $noTermsLog"
  $noTermsManifest = Join-Path $noTermsOut 'SmartFinOps_Workplace-Report.demo.manifest.json'
  Check ((Test-Path -LiteralPath $noTermsManifest) -and (Get-Content -LiteralPath $noTermsManifest -Raw | ConvertFrom-Json).PrivateTermsChecked -eq 0) 'manifest records 0 private terms checked'
  #endregion

  #region 5. Leak scanner
  . $generator
  $clean = '<html><head><title>x</title><style>p{}</style></head><body><p>north.contoso.invalid 2.4.1 12:30:45 contoso-demo https://workplacecloudhub.com</p><svg xmlns="http://www.w3.org/2000/svg"></svg></body></html>'
  Check (@(Test-DemoLeak -Html $clean).Count -eq 0) 'clean page passes the leak scan'
  Check (@(Test-DemoLeak -Html $clean.Replace('</body>', '<div class="logo"><img alt="WorkplaceCloudHub" width="120" height="97" src="data:image/png;base64,AAAA"></div></body>')).Count -eq 0) 'the WorkplaceCloudHub logo alone passes the leak scan'
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
    'image other than the WorkplaceCloudHub logo' = '<img alt="Client" src="data:image/png;base64,AAAA">'
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0JYJlXNveItuZ
# IlRdPyJr+ab9TS6lTm4cKHi7qF9bNKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINLyvBISmWTp1uulF5xCcEHP5QvFXeKknEr2zmOKZ/0XMA0GCSqG
# SIb3DQEBAQUABIIBgKqn9Kuzz44gQma5T7fIxGeYIY0nd73Flngl5OJiHNl9x9yj
# Y8U6q8Np5r7a2ftbgwovJxZX7tlsov7YU/BKEjlmjiEJK3BbvSnee+zirWcL2azo
# TJXNbgQ9ZrPRxAFSDwXU1KdleLwyeyTkgviHSPoUPpJnczplrSDt6IPidAfRdhMm
# IIZB3502L+kiHPOGgIMKtBXPxuKM2FMXxIRbc9FTiG6wmllbqfhGMLTBm4omftLt
# CxoZm25WhN7Rb/G/I9s6S56ESBpt6KPG166t08VyqQdtGBMmo4lUjtX6w7XGrdJO
# hCqD/F7Db5ksk6KnHnnD2DgY8zknQnh8cZnjRUgjsO3iT1dmXpLT3d8wB7mL4CM4
# ahLjGNdd4yB0O5Nl9vtGYtB+Hz9rZg458I3Ec/8m3XwS7XJcHZFt/1dmeZsRArfQ
# 0ch2kHThrRO1Ghnh+IzCG/4f0/AGUv016lGYY2MakcJHt3cHq01eVES9slqtTN0m
# J33/COZer0F4xijuPKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxOTUz
# MzZaMC8GCSqGSIb3DQEJBDEiBCCKOci13a5qzoEONCP/dkP3WTHktwmG6arWCBgD
# t1YfqzANBgkqhkiG9w0BAQEFAASCAgBqKJASMWIbJ7Dx75kPW/7b4qOuHE/as1ol
# UQgB3zOlv+BPQNhtmGhFLwn7MS8DMTpVRPvN6R2sILFILT0GrPHOpDCip3r4qo4g
# nMEIj9lZHW11Qryw4c2/g1EC7InwOusIcbJCcD381os+Dg3ae1CqYTNT4BirXqm1
# ApFKGecvtcqF+26oEsO/yraNens5IYi6vpEZpyzUaktg0GAcFVSqCj6dt7hvg2ky
# S9o3m2WB2q5lPAwqbN5Z+q5esNB8PO9y2IFOHzx5vfzkw3e9fQHe/t+EYen0bE7d
# 2IpzUJTXapvzBFkzaYbdQwYoG2rGXb5IQPOtM74MPsn0UEy2DHJnWhrm/2GxYfSK
# xb5uYSFvLeJ5X6VXjBjLbCMI4lS0ZJcE58pepeI36i7XP5CVZdSc+Q9asq/NTRrp
# iwD1tA16Ve8HqsPe1ZbgZBp6qtmCuQb53DwScCDl2aObJGW2UTRN+ZNvH8anvgjK
# 0h6LVw6j0WyLcMIl+j7/++0nkJJu+4HsbKiCeiZJZi11y/sv0brKZ28T4wrLvVKL
# oqI5dsIs5SO2OHfUxR1c5U9IErhfJfhbBYGTWjQvM7qQBeC+leTHr01891afXVXO
# 1kbZ5u7SZ+88rNFO5K5f8TwH4z5imhn0lE0mnn/B6wsQv6JdpQZIxzXGDvqlDXBV
# qX11N8AwhA==
# SIG # End signature block
