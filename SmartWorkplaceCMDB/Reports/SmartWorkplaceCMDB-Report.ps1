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



# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOxuuRQaRtjNUF
# ZmCQa5F1kJ73L4fK6x3DipBBPAKEb6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCD6XtV2q/GDQPEbK3b8+EXg6yq65oZMQtTAZ6qIirVOnjANBgkqhkiG9w0B
# AQEFAASCAYCwcBeOLPyFW6rkBgHsdRX1n0zbp+h5gstjbIIdkpK7dJIaZRBGO3i7
# GB3G7wqbrMsPQYMYRes6TVQUO/7sd0qjdpU8wNjupllTCnqw5UVPPgNR8TS57/Yx
# hXNeHRW9nuVkBKj7tElxZToBm1F5rel7TRGRpGk8qdLwKDJkwoeyaL/UjtSBGZeK
# qc7AKc3tW8yOqVziXGD7UQ80T++ZxR7q0TDyX9FwFb+Kvy1SxYxQWe0cSqBBaaQK
# uhKOGX0DE356zCJkyv2RDUDDr3Xtef/xSIS78V4wzlfsik3K1y2XokhZZ8S2ijJL
# II8rzoRXZqWUMQT3fmd/yCE3jgD/ONFD0Nw2auNBsToD/f7+qXmMJzVZN77l8Nsz
# OEP/Yo6GiD35f8y1PRcQEDIMcJceBTiN42cyCF87evWw81sLkIa0Z0o0gcH8NWOD
# 6z8lOAxCxagnVScOUqvWW3OLvU+wxq7aVaKmJdh4gy+lS/UumZO4wtfIiWj0I4fx
# S80DVT8kUas=
# SIG # End signature block
