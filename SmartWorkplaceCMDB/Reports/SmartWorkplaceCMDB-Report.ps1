<#
.SYNOPSIS
Creates a local SmartWorkplaceCMDB HTML overview report.

.VERSION
0.1.1
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'Default',
    [string]$LatestOutputRootPath,
    [string]$OutputPath,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.1'
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'

Import-Module $modulePath -Force

$paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant $Tenant -LatestOutputRootPath $LatestOutputRootPath

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $paths.LatestOutputRootPath -ChildPath 'SmartWorkplaceCMDB-Overview.html'
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status        = 'Valid'
        ScriptVersion = $ScriptVersion
        Tenant        = $paths.TenantKey
        OutputPath    = $OutputPath
    } | Format-List
    return
}

Initialize-SmartWorkplaceCMDBTenantFolder -Paths $paths

$cmdbPath = $paths.CmdbLatestPath
$powerBiPath = $paths.PowerBILatestPath
$cmdbTables = if (Test-Path -LiteralPath $cmdbPath) { Get-ChildItem -LiteralPath $cmdbPath -Filter '*.csv' -File } else { @() }
$powerBiTables = if (Test-Path -LiteralPath $powerBiPath) { Get-ChildItem -LiteralPath $powerBiPath -Filter '*.csv' -File } else { @() }

$generatedAt = Get-Date
$html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>SmartWorkplaceCMDB Overview</title>
  <style>
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #F5F8FB; color: #1F2937; }
    main { max-width: 1180px; margin: 0 auto; padding: 28px; }
    header, section { background: #FFFFFF; border: 1px solid #DDE7F0; border-radius: 8px; padding: 22px; margin-bottom: 18px; }
    h1, h2 { margin: 0 0 10px 0; font-weight: 650; }
    .muted { color: #5F6B7A; }
    .grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
    .metric { border: 1px solid #DDE7F0; border-radius: 8px; padding: 14px; background: #F8FBFE; }
    .value { font-size: 28px; font-weight: 650; color: #0078D4; }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>SmartWorkplaceCMDB Overview</h1>
      <div class="muted">Tenant: $($paths.TenantKey) | Generated: $($generatedAt.ToString('yyyy-MM-dd HH:mm:ss')) | Script: $ScriptVersion</div>
    </header>
    <section class="grid">
      <div class="metric"><div class="value">$($cmdbTables.Count)</div><div>CMDB tables</div></div>
      <div class="metric"><div class="value">$($powerBiTables.Count)</div><div>Power BI tables</div></div>
      <div class="metric"><div class="value">0</div><div>Collector findings</div></div>
    </section>
    <section>
      <h2>Executive Overview Status</h2>
      <p class="muted">The initial scaffold is ready. Collector data will populate coverage, compliance, licensing, mailbox, freshness, and data quality indicators in the next phase.</p>
    </section>
  </main>
</body>
</html>
"@

$folder = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $folder)) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8 -Force
Write-Information ("SmartWorkplaceCMDB overview report created: {0}" -f $OutputPath) -InformationAction Continue


