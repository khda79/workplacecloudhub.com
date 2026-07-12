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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOxuuRQaRtjNUF
# ZmCQa5F1kJ73L4fK6x3DipBBPAKEb6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD6XtV2q/GDQPEbK3b8
# +EXg6yq65oZMQtTAZ6qIirVOnjANBgkqhkiG9w0BAQEFAASCAYCIpivWKk1xlX49
# Us4Nj3f4v78ReD0EDm+QSuZVfFS5tU308ogMbb2qg60KCWZqALn8Zkcdn7frYq14
# 8oTm4ApxOkn+5Kjw7SQU1Z0PjZ53YbCAkItcgjsfoBBVR0of1Tkre0oRCNp7qLQj
# 49GaPWI0/XU+rRRsaUveYd3vFm9UuoZHaOFDwCJt93rDOSRgGk2zU8uPIdg0lxqm
# sOKGoM4Y1DC9uD04ERL0G7ml964qL3xCvvc35f98dDNSBNkpK9TvSzuC1dlFGlBQ
# dQWqmE+agQ8+YwNwO50nC0Qr8SIhMDzhnheu9cG1Lpz2wQVTLonvZzJr9ZAyDB4f
# gi3ohetnA0HBUmc9ZxB5ZGSucs63uf5Thjykstw6DaghdwRUMqLbR/fNn5OccUa8
# n0h+F2vZ2T806bkCzCH2cTP9hDUssNstpzROQ7b1lwFZ3UMO0lyT/HV7o0ofeEs1
# HDt+UYSlCmV5qUuvOP8dgHKwuIJ/Ck12xQnbEVdj/iMpJ7GCFsM=
# SIG # End signature block
