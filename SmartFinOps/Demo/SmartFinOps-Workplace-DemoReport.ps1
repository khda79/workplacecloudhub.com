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
