<#
.SYNOPSIS
Creates a local SmartWorkplaceCMDB HTML overview report.

.DESCRIPTION
Loads the autonomous SmartWorkplaceCMDB configuration and reports CSV contract
readiness, table counts, row counts, and data-quality finding counts.

.VERSION
0.2.2-beta.2
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')]
    [string]$Tenant = 'default',
    [string]$OrganizationKey,
    [string]$EnvironmentKey,
    [string]$TenantKey,
    [string]$TenantId,
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [string]$GlobalConfigPath,
    [string]$TenantConfigPath,
    [string]$OutputPath,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.2.2-beta.2'
$ErrorActionPreference = 'Stop'

function Get-CsvDataRowCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]$Paths,
        [switch]$TenantNeutral
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return 0
    }

    $rowCount = 0
    Import-Csv -LiteralPath $Path | ForEach-Object {
        if (-not $TenantNeutral) {
            foreach ($name in @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')) {
                if ([string]$_.$name -ne [string]$Paths.$name) {
                    throw "CSV tenant identity mismatch in '$Path' for '$name'."
                }
            }
        }
        $rowCount++
    }
    return $rowCount
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'

Import-Module $modulePath -Force

$boundParameterCopy = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $boundParameterCopy[$key] = $PSBoundParameters[$key]
}
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $boundParameterCopy -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$contract = Get-SmartWorkplaceCMDBTableContract -Path $context.ContractPath
$contractResults = @(Test-SmartWorkplaceCMDBCsvContract -LatestOutputRootPath $paths.LatestOutputRootPath -ContractPath $context.ContractPath)
$incompatibleTables = @($contractResults | Where-Object Status -eq 'Incompatible')
if ($incompatibleTables.Count -gt 0) {
    $details = @($incompatibleTables | ForEach-Object {
        '{0}: missing=[{1}] unexpected=[{2}]' -f $_.Name, $_.MissingColumns, $_.UnexpectedColumns
    })
    throw ("Cannot report on incompatible SmartWorkplaceCMDB CSV outputs. {0}" -f ($details -join '; '))
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $paths.LatestOutputRootPath -ChildPath 'SmartWorkplaceCMDB-Overview.html'
}

$rowCounts = @{}
foreach ($table in @($contractResults | Where-Object Status -eq 'Valid')) {
    $rowCounts[$table.Name] = Get-CsvDataRowCount -Path $table.Path -Paths $paths -TenantNeutral:($table.Name -eq 'DimDate.csv')
}

if ($ValidateOnly) {
    [pscustomobject]@{
        Status               = 'Valid'
        ScriptVersion        = $ScriptVersion
        ContractVersion      = [string]$contract.contractVersion
        ValidContractFiles   = @($contractResults | Where-Object Status -eq 'Valid').Count
        MissingContractFiles = @($contractResults | Where-Object Status -eq 'Missing').Count
        ProfileKey           = $paths.ProfileKey
        OrganizationKey      = $paths.OrganizationKey
        EnvironmentKey       = $paths.EnvironmentKey
        TenantKey            = $paths.TenantKey
        TenantId             = $paths.TenantId
        LatestOutputRootPath = $paths.LatestOutputRootPath
        OutputPath           = $OutputPath
        GlobalConfigPath     = $context.GlobalConfigPath
        TenantConfigPath     = $context.TenantConfigPath
    } | Format-List
    return
}

$validTables = @($contractResults | Where-Object Status -eq 'Valid')
$missingTables = @($contractResults | Where-Object Status -eq 'Missing')
$cmdbTables = @($validTables | Where-Object Area -eq 'CMDB')
$powerBiTables = @($validTables | Where-Object Area -eq 'PowerBI')
$cmdbRows = 0
foreach ($table in $cmdbTables) {
    if ($table.Name -ne 'CMDB_BuildManifest.csv') {
        $cmdbRows += $rowCounts[$table.Name]
    }
}
$powerBiRows = 0
foreach ($table in $powerBiTables) {
    $powerBiRows += $rowCounts[$table.Name]
}
$dataQualityPath = Join-Path -Path $paths.CmdbLatestPath -ChildPath 'CMDB_DataQuality.csv'
$dataQualityFindings = if ($rowCounts.ContainsKey('CMDB_DataQuality.csv')) { $rowCounts['CMDB_DataQuality.csv'] } else { 0 }

$generatedAt = Get-Date
$sourceHealth = @(Get-SmartWorkplaceCMDBSourceHealth -Paths $paths)
$sourceHtml = (@($sourceHealth | ForEach-Object {
    $source = $_
    $cells = @('SourceName','Status','Coverage','CompletedUtc','RowCount') | ForEach-Object {
        '<td>{0}</td>' -f [System.Net.WebUtility]::HtmlEncode([string]$source.$_)
    }
    '<tr>{0}</tr>' -f ($cells -join '')
}) -join "`n")
$profileText = [System.Net.WebUtility]::HtmlEncode([string]$paths.ProfileKey)
$tenantText = [System.Net.WebUtility]::HtmlEncode([string]$paths.TenantKey)
$missingText = if ($missingTables.Count -eq 0) {
    'None'
}
else {
    [System.Net.WebUtility]::HtmlEncode((($missingTables.Name | Sort-Object) -join ', '))
}

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
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th, td { text-align: left; border-bottom: 1px solid #DDE7F0; padding: 8px; }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>SmartWorkplaceCMDB Overview — BETA</h1>
      <div class="muted">Profile: $profileText | Tenant: $tenantText | Generated: $($generatedAt.ToString('yyyy-MM-dd HH:mm:ss')) | Script: $ScriptVersion | Contract: $($contract.contractVersion)</div>
    </header>
    <section class="grid">
      <div class="metric"><div class="value">$cmdbRows</div><div>CMDB rows</div></div>
      <div class="metric"><div class="value">$powerBiRows</div><div>Power BI rows</div></div>
      <div class="metric"><div class="value">$dataQualityFindings</div><div>Data quality findings</div></div>
    </section>
    <section>
      <h2>Contract Readiness</h2>
      <p>Valid files: $($validTables.Count) / $(@($contract.tables).Count)</p>
      <p class="muted">Missing files: $missingText</p>
    </section>
    <section>
      <h2>Source evidence</h2>
      <p>Complete means a successful unbounded live snapshot within the recorded collector scope. Fixture, bounded, scoped and unknown sources do not prove full live coverage. Unknown includes legacy files without collection evidence. A valid CSV contract alone does not prove collection success.</p>
      <div style="overflow-x:auto"><table><thead><tr><th>Source</th><th>Status</th><th>Coverage</th><th>Completed (UTC)</th><th>Rows</th></tr></thead><tbody>$sourceHtml</tbody></table></div>
    </section>
  </main>
</body>
</html>
"@

$folder = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $folder)) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

$tempPath = '{0}.tmp.{1}' -f $OutputPath, ([guid]::NewGuid().ToString('N'))
try {
    Set-Content -LiteralPath $tempPath -Value $html -Encoding UTF8 -Force
    Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
}
finally {
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
}

Write-Information ("SmartWorkplaceCMDB overview report created: {0}" -f $OutputPath) -InformationAction Continue

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+bCVLB24tk7+j
# 2ZM85G3cNnh21d/cNTo44Ut8DZcHXqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIChY9HRkHZwuhEnbFDJJT8xNeUfjssVMj+1f4RQYICOeMA0GCSqG
# SIb3DQEBAQUABIIBgJnwJ34kFBAAXcQBMHwh+mPVk0Dn/gOoyLz9+4Hgc0QH/rGp
# Mv6A26UrzM9M+nUwdn7vkJ3cOimaWm+AeQzAtJsMV/IyapJypMwXd+3eeDwge86c
# haKHxdQ57ZHWp9La9BspK36Uuyl1A2BIr/ZqfEnnbaxGdbBpguBu7HA6XXAtFdoI
# uYeFD9oHM4hgSc39a5obYO4/bmsTGprrF8ggJetTJryMlio5yOxEqIxlCwTwIlsc
# jVbyEg6Tf/vt4ldS6b6WP8ATCzXeTi/HdkJbb9Jcx2NjC7UTeDTJGooOAFc6zvry
# RTqk2qT3sEgkB21KThVVTGg6c/cjHjAeYF0STkE7hvcQiU9snRmhgChhTfYpZgHR
# 7BKpKGHM3K9vtYE8V147e877Agme6+vaKa9Li/o6CWaEUxzVuHq3JgNFq19K5uQW
# aXJzw0gXZs3yp1pu203CMMwHX16rptny77O0ttoKIKz0idH8sd4SrdrygVr5u0Qt
# bv3eoU5cf0Zp0KZwh6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkxNDIy
# MDlaMC8GCSqGSIb3DQEJBDEiBCBYGdZnAyYUtBbbkpJv+eCD/XQzrGs69iUIEwJD
# knYtWjANBgkqhkiG9w0BAQEFAASCAgA4keTnhWITXCmyrBRy72DKjuxFDay6lRDS
# uWj/mm9JQKveR/J0THk9sFOWdalRe2mSiKP4NvvnFuxGTJP4LFyX3Mbh1NtCWe+C
# sPxjY0iFCbipExK6rZxVYoWTF182JX0gZH4S5KUZK7zusBkTIIOCHy+UYVWyWuwF
# YGnyztcS1RWqAYhm9TYHVZ9ATG4O/FKuUxX4eQhiuqH9hAQf+2ISswaSpDFMzwEx
# T0EKe1QLZjztWTSthf5LVFxIUZA72Ssd/a4B5ywSP4JD69VXm031W26lRYw+yGUz
# JGHdHknvOytzJzLVZGl1GIJBvT5MLwvzkhAn9/9ykXGGgbEuEHL18FosYg6u/Zce
# LtN0bzev5zPCeW584gZcDkv4G457kLyQd31s1BgVFIRRyEc3Kigbo0fGVqkXbXW7
# hXytHXg9IPM/gvT1ctm325fv74N0C8uutCYnidoEtKtpmxIf5J+J2fz3aV0Wol/O
# ZWiRzWJ4xJ+KRnNWKe/f4OACTla/pxhWdOq2k3xoLyb1HkZ7xZRCB0grZO4MrLdN
# A88rU0jeBjqUpUfP85cdSh1aWB6CyiFyrrJKcIUs/vXFH5WRT7LKmJiDHK/E3GEX
# vo5ZQ4lmsAnHz6nHDKdyuuDpT0gybzyzzreN37KBjIaDYjhr6tduvlwtkwHDSU9l
# VAizXo43ww==
# SIG # End signature block
