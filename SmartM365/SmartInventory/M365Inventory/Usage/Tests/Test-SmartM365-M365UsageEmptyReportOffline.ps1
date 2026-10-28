<#
.SYNOPSIS
    Offline checks for empty Microsoft 365 usage reports.
.DESCRIPTION
    Runs the collector's actual retry and fallback functions on synthetic CSVs
    with a mocked download. No Graph, tenant collection, or publication is
    performed.
.VERSION
1.1
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$collector = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-M365UserActivity-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($name in @('Get-M365UsageReportDefinition','Assert-M365UsageReportRefreshDate','Get-M365UsageRawFileSummary','Invoke-M365UsageReportAcquisition','Resolve-M365UsageEmptyReportFallback')) {
    $definition = @($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true))
    if ($definition.Count -ne 1) { throw "Expected one $name function." }
    Invoke-Expression $definition[0].Extent.Text
}

$script:warnings = [System.Collections.Generic.List[string]]::new()
$script:sleeps = [System.Collections.Generic.List[int]]::new()
$script:responses = [System.Collections.Generic.Queue[string]]::new()
$script:downloads = 0
function WriteLog { param([string]$Message,[string]$Level) if ($Level -eq 'WARNING') { $script:warnings.Add($Message) | Out-Null } }
function Start-Sleep { param([int]$Seconds) $script:sleeps.Add($Seconds) | Out-Null }
function Invoke-M365UsageReportDownload {
    param($ReportDefinition,[string]$RawPath)
    $script:downloads++
    Set-Content -LiteralPath $RawPath -Value $script:responses.Dequeue() -Encoding UTF8
}
function Assert-Case { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Assert-Rejected { param([scriptblock]$Action,[string]$Pattern,[string]$Message)
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $_.Exception.Message -match $Pattern }
    Assert-Case $caught $Message
}
function Reset-Case { param([string[]]$Responses)
    $script:warnings.Clear(); $script:sleeps.Clear(); $script:responses.Clear(); $script:downloads = 0
    foreach ($response in $Responses) { $script:responses.Enqueue($response) }
}

$definitions = Get-M365UsageReportDefinition
$activation = $definitions['Office365ActivationUserDetail']
$activeUsers = $definitions['Office365ActiveUserDetail']
$refreshText = (Get-Date).Date.AddDays(-2).ToString('yyyy-MM-dd')
$header = 'Report Refresh Date,User Principal Name,Product Type'
$data = "$header`r`n$refreshText,user1@contoso.example,MICROSOFT 365 APPS FOR ENTERPRISE"
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-UsageEmpty-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $rawPath = Join-Path $tempRoot 'raw.csv'
    $latest = Join-Path $tempRoot 'DATA-LAST'
    New-Item -ItemType Directory -Path $latest -Force | Out-Null

    Assert-Case (@($definitions.Values | Where-Object { $_.PSObject.Properties['Required'] -and $_.Required }).Name -join ',' -eq 'Office365ActiveUserDetail') 'Only the active user detail report may be required.'

    Set-Content -LiteralPath $rawPath -Value $header -Encoding UTF8
    Assert-Rejected { Assert-M365UsageReportRefreshDate -RawPath $rawPath -ReportName 'X' -WarningAgeDays 3 } 'Graph report is empty: X' 'Strict refresh-date check accepted an empty report.'

    # Empty, then populated: the retry recovers within the same run.
    Reset-Case @($header,$data)
    $date = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($null -ne $date -and $date.ToString('yyyy-MM-dd') -eq $refreshText) 'Populated retry did not return its refresh date.'
    Assert-Case ($script:downloads -eq 2 -and ($script:sleeps -join ',') -eq '120') "Unexpected retry sequence: downloads=$script:downloads sleeps=$($script:sleeps -join ',')."
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'attempt 1/3; Bytes=\d+; Columns=3; Header=Report Refresh Date,User Principal Name,Product Type\. Retrying in 120 seconds') "Retry warning missing or without the raw summary: $($script:warnings -join ' | ')"

    # A populated first answer is neither retried nor warned.
    Reset-Case @($data)
    $null = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($script:downloads -eq 1 -and $script:sleeps.Count -eq 0 -and $script:warnings.Count -eq 0) 'Populated report was retried.'

    # Persistently empty: every attempt is used, no sleep after the last one.
    Reset-Case @($header,$header,$header)
    $date = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($null -eq $date -and $script:downloads -eq 3 -and $script:sleeps.Count -eq 2) 'Persistently empty report was not retried exactly three times.'
    Assert-Case ($script:warnings.Count -eq 3 -and $script:warnings[2] -match 'attempt 3/3; Bytes=\d+; Columns=3; Header=' -and $script:warnings[2] -notmatch 'Retrying') 'Last empty attempt was not logged with its raw summary.'

    # Raw summary: missing file, quoted header with a comma inside a column, long header truncated.
    Remove-Item -LiteralPath $rawPath -Force
    Assert-Case ((Get-M365UsageRawFileSummary -RawPath $rawPath) -eq 'Bytes=0; Columns=0; Header=') 'Missing raw file summary is wrong.'
    Set-Content -LiteralPath $rawPath -Value '"Report Refresh Date","Name, Display","Product Type"' -Encoding UTF8
    Assert-Case ((Get-M365UsageRawFileSummary -RawPath $rawPath) -match 'Columns=3;') 'Quoted header columns were miscounted.'
    Set-Content -LiteralPath $rawPath -Value ('A' * 400) -Encoding UTF8
    Assert-Case ((Get-M365UsageRawFileSummary -RawPath $rawPath) -match "Header=A{300}\.\.\.$") 'Long header was not truncated.'

    # Non-required report: previous CSV kept byte for byte, its refresh date and age reported.
    $previous = Join-Path $latest 'M365_Apps_Activations.csv'
    'RunId,ReportName,ReportPeriodRequested,Report Refresh Date,User Principal Name','old,Office365ActivationUserDetail,D180,2026-10-18,user1@contoso.example' | Set-Content -LiteralPath $previous -Encoding UTF8
    $written = (Get-Date).ToUniversalTime().AddHours(-100)
    (Get-Item -LiteralPath $previous).LastWriteTimeUtc = $written
    $before = (Get-FileHash -LiteralPath $previous -Algorithm SHA256).Hash
    $script:warnings.Clear()
    $result = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3 -AlertAgeHours 72
    Assert-Case ($result.Entry -eq 'Office365ActivationUserDetail=2026-10-18') "Unexpected carried-over entry: $($result.Entry)"
    Assert-Case ($result.FileAgeHours -eq 100 -and $result.Alert) "Kept CSV older than the threshold was not flagged: age=$($result.FileAgeHours) alert=$($result.Alert)"
    Assert-Case ($result.Detail -eq 'Office365ActivationUserDetail: Report Refresh Date=2026-10-18, M365_Apps_Activations.csv written 100 h ago') "Unexpected carried-over detail: $($result.Detail)"
    Assert-Case ((Get-FileHash -LiteralPath $previous -Algorithm SHA256).Hash -eq $before -and (Get-Item -LiteralPath $previous).LastWriteTimeUtc -eq $written) 'Previous CSV was modified.'
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'kept unchanged \(Report Refresh Date=2026-10-18; file age=100 h\).*not part of this run''s receipt') 'Carried-over warning missing.'

    # Below the threshold: kept and warned, but no alert. Exactly at the threshold: alert.
    $result = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3 -AlertAgeHours 101
    Assert-Case (-not $result.Alert) 'Kept CSV below the threshold was flagged.'
    $result = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3 -AlertAgeHours 100
    Assert-Case ($result.Alert) 'Kept CSV at the threshold was not flagged.'

    Remove-Item -LiteralPath $previous -Force
    $script:warnings.Clear()
    $result = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3
    Assert-Case ($result.Entry -eq 'Office365ActivationUserDetail=none' -and $null -eq $result.FileAgeHours -and $result.Alert -and $script:warnings[0] -match 'No previous M365_Apps_Activations.csv exists') 'Missing previous CSV was not reported and flagged.'

    # Required report: an empty answer stays fatal.
    Assert-Rejected { Resolve-M365UsageEmptyReportFallback -ReportDefinition $activeUsers -LatestFolderPath $latest -Attempts 3 } 'Graph report is empty: Office365ActiveUserDetail after 3 attempt\(s\)' 'Required empty report did not fail the run.'

    # The run loop uses the retry and the fallback, and records the fallback in the receipt.
    $runText = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and $_.Finally })[0].Body.Extent.Text
    Assert-Case ($runText -match 'Invoke-M365UsageReportAcquisition' -and $runText -match 'Resolve-M365UsageEmptyReportFallback' -and $runText -notmatch 'Assert-M365UsageReportRefreshDate') 'Run loop does not use the retry and fallback.'
    Assert-Case ($runText -match 'not receipted' -and $runText -match 'EmptyKeptPrevious') 'Carried-over reports are missing from the receipt qualifications or summary.'
    Assert-Case ($runText -match "'CarriedOverReportAlertAgeHours' -DefaultValue 72" -and $runText -match '-AlertAgeHours \$carriedOverAlertAgeHours') 'Alert threshold is not read from configuration or passed to the fallback.'
    Assert-Case ($runText -match "-Level WARNING -Channel Alerts -Title 'SmartM365 M365 usage report kept from a previous run'") 'Stale carried-over CSV is not sent to the Alerts channel.'
    $template = Get-Content -LiteralPath (Join-Path (Split-Path $collector -Parent) 'SmartM365-M365UserActivity-Inventory.local.json.template') -Raw | ConvertFrom-Json
    Assert-Case ($template.CarriedOverReportAlertAgeHours -eq 72) 'Template does not declare CarriedOverReportAlertAgeHours=72.'
    'PASS: empty reports retry then recover and log their raw summary; persistent empties keep the previous CSV with its age and alert past the threshold; the required report stays fatal.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $base = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($base,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-UsageEmpty-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCwahvcQ+aHlOxb
# 0BMDSOtY1bATlEzHbZiiuNbd+9HU5aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPOjj9nod6r5PL0M52X674plfX6+xYTWLPVaDoGIrnKQMA0GCSqG
# SIb3DQEBAQUABIIBgDldeyl844/dM77syOX7whiu/l/j5+kfk3wxxynG26uy2C6A
# Q2HlPlcTcI6zIa77RXMhzeH6YJORH7MwEPFo13yWP4GCY7NXzwje8DnlHaC7TOH8
# UaUWnrkzpZ/FmW1TFkZmFwFYAnwyuNLBKXUTlaiX8YwTwGYKkDm1l/Ch7MvqyNLB
# knZqUQX5EH2z2vP7snEk6jZW0G7TDiMZoWyXIu41B+h1/bXSts//rGzvNLV+4CE+
# s1+ruu/Vcep2e9YiOTDqevvHj3EV5hdMcWAt7XNE9aBcBLPXeOr2NuK2ReesOz8O
# D8oxiNYk+gsp2FW9YfdC6qvLh7UQXWMbOnh6OKvgL0I2Eo/USMyw7iBKPXsnkrcl
# mN5H3nDO7K4NeWUw6qHVON8O+lsYCrqxF50hhfNeAM6VESzqAUKbgIdICDz1ZcG/
# gQ3CfH3VwvjDrn6gQP8C1ZyVwbc/2SMQlVe5tZeX/XQo+B7x3VMIXEDMDP+TaIsc
# 4hsluqSDYmssgIU/UqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIwOTEy
# MjVaMC8GCSqGSIb3DQEJBDEiBCAkCuvUGhPErrwMVwQMJq7SJQSFaPdKfAslH0iY
# d6TLKTANBgkqhkiG9w0BAQEFAASCAgBGOyR6EKKdauVdhwjP/po3O3TRRSgI9upH
# /LO8M17TWQP9YyT4O4ap9sYWlWHiC4euZCTstsGkGP6GUtooQDSNFORwXp40Yr85
# yVbvxeKxPHC1p6ERGUVf/eU1PvZovjCkJTErBQK/hjxoj0uLVwRN+7D7Elx0+LsV
# 4w2TID3JmzwGYR+YvEDWLCfeOUuehpQlV0bsrQI0ojUK743+g/Ifal9mDEjLMY47
# fCK5APN30Jt9sVrh/5VNnbrYOvhC0HTYyMSZPk7qeWPs6rZLF6YXNR5w2GFP092F
# VPkMWXZQ4WsKPkBqemR2cE6RjDbMVUpAbKDaKcm1fg5jOhgKXNGcmYOkcciXPgx3
# 5MqYmWLU9/p64km4VAFxZTQcTmIzeeg8RyiPG8eNpu+dovL3L9OHSlYxjdkV+SGO
# 7nXkVeU6agZOuJJO9rDv8//Qdz20S4LCiIjfYP5SQlD2Ry3wFMFau0GmrxqQNORg
# 4XhsfCOP6pKdz36AI/JOI+sXIMtBeoqb/eT6O1a6G/C9vLsK0t55KCCX3wuUDO7c
# qiRDOyTaXtwFB9b3L4H9DL2vG/L93rF6ZysOKzsviHqWzfq1ADAKluUQd/8vjWZY
# dzt8Df2QMw2Ow+Rdaa/wFEnRfSKX+CJ3J372aSXdtyWkGB3pUOhoOCau0WaKF67q
# 730WGoF9og==
# SIG # End signature block
