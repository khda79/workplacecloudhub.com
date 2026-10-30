#Requires -Version 7.0
<#
.SYNOPSIS
Checks the present / functional capability columns of Servers and versions, offline.
.DESCRIPTION
Present capabilities come from the orchestrator probes (Orchestrator-Capabilities.json, used by
the election). Functional capabilities come from the last server acceptance test run as the
orchestrator service account (Orchestrator-Acceptance.json); a run under another account
(Orchestrator-Acceptance.OtherAccount.json) is reported but never validates. Display only.
Synthetic files in a temporary folder; no tenant, collector or publication.
.VERSION
1.0.1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$script:Checks = 0
function Assert-Display {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}
function Write-TestJson([string]$Path, $Value) { [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false)) }
function New-Run([string]$Account, [bool]$ServiceAccount, [object[]]$Results) {
    [ordered]@{ SchemaVersion = 1; ServerName = 'X'; GeneratedAtUtc = '2026-10-30T10:15:00.0000000Z'; AcceptanceVersion = '1.0.12'; RunAccount = $Account; ServiceAccount = $ServiceAccount; Results = $Results }
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AcceptanceDisplay-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    foreach ($server in @('SRV-A', 'SRV-B', 'SRV-C', 'SRV-D', 'SRV-E')) {
        $folder = Join-Path $root $server
        $null = New-Item -ItemType Directory -Path $folder -Force
        Write-TestJson (Join-Path $folder 'Orchestrator-Heartbeat.json') ([ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.80' })
        Write-TestJson (Join-Path $folder 'Orchestrator-Capabilities.json') ([ordered]@{ ReadyCapabilities = @('Graph', 'Python', 'SharedRuntime', 'TeamsPowerShell') })
    }
    $results = @(
        [ordered]@{ Name = 'Graph'; Status = 'PASS'; Detail = '100 users' }
        [ordered]@{ Name = 'SharedRuntime'; Status = 'PASS'; Detail = 'ok' }
        [ordered]@{ Name = 'Python'; Status = 'PASS'; Detail = 'Python 3.12.1' }
        [ordered]@{ Name = 'TeamsPowerShell'; Status = 'WARN'; Detail = 'module copy not readable' }
        [ordered]@{ Name = 'EXO'; Status = 'FAIL'; Detail = 'Forbidden' }
        [ordered]@{ Name = 'ExchangeOnPrem'; Status = 'SKIPPED'; Detail = 'snap-in not registered' }
        [ordered]@{ Name = 'Network'; Status = 'FAIL'; Detail = 'not a capability' }
    )
    # SRV-A: service account run; a later admin run must not hide it.
    Write-TestJson (Join-Path $root 'SRV-A/Orchestrator-Acceptance.json') (New-Run 'CONTOSO\svc-orch' $true $results)
    Write-TestJson (Join-Path $root 'SRV-A/Orchestrator-Acceptance.OtherAccount.json') (New-Run 'CONTOSO\admin' $false @([ordered]@{ Name = 'Graph'; Status = 'FAIL'; Detail = 'x' }))
    # SRV-B: only an admin run.
    Write-TestJson (Join-Path $root 'SRV-B/Orchestrator-Acceptance.OtherAccount.json') (New-Run 'CONTOSO\admin' $false $results)
    # SRV-C: never tested. SRV-D: unreadable result. SRV-E: older document without optional fields.
    [IO.File]::WriteAllText((Join-Path $root 'SRV-D/Orchestrator-Acceptance.json'), '{ not json', [Text.UTF8Encoding]::new($false))
    Write-TestJson (Join-Path $root 'SRV-E/Orchestrator-Acceptance.json') ([ordered]@{ Results = @([ordered]@{ Name = 'AD'; Status = 'PASS' }) })

    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C', 'SRV-D', 'SRV-E'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $rows = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)
    $row = { param($Name) @($rows | Where-Object Server -eq $Name)[0] }
    $localAt = ([datetimeoffset]::Parse('2026-10-30T10:15:00Z')).LocalDateTime.ToString('yyyy-MM-dd HH:mm')

    $a = & $row 'SRV-A'
    Assert-Display ($a.Capabilities -eq 'Graph, Python, SharedRuntime, TeamsPowerShell') "Present capabilities still come from the probes: $($a.Capabilities)"
    Assert-Display ($a.FunctionalCapabilities -eq 'Graph, Python, SharedRuntime') "Functional = PASS of the service account run: $($a.FunctionalCapabilities)"
    Assert-Display ($a.NotFunctionalCapabilities -eq 'EXO (FAIL), TeamsPowerShell (WARN)') "Not functional = WARN and FAIL capabilities only: $($a.NotFunctionalCapabilities)"
    Assert-Display ($a.Acceptance -eq "$localAt, service account CONTOSO\svc-orch, acceptance 1.0.12") "Acceptance time and account: $($a.Acceptance)"

    $b = & $row 'SRV-B'
    Assert-Display (-not $b.FunctionalCapabilities -and -not $b.NotFunctionalCapabilities -and $b.Acceptance -eq "Not validating: $localAt, run as CONTOSO\admin, not the service account, acceptance 1.0.12") "An admin run never validates: $($b.Acceptance)"
    $c = & $row 'SRV-C'
    Assert-Display (-not $c.FunctionalCapabilities -and $c.Acceptance -eq 'Never run as the service account') "Never tested: $($c.Acceptance)"
    $d = & $row 'SRV-D'
    Assert-Display (-not $d.FunctionalCapabilities -and $d.Acceptance -like 'Acceptance result unreadable: *' -and $d.Capabilities -like 'Graph*') "Unreadable result does not break the row: $($d.Acceptance)"
    $e = & $row 'SRV-E'
    Assert-Display ($e.FunctionalCapabilities -eq 'AD' -and $e.Acceptance -eq 'service account') "Older document without optional fields: $($e.Acceptance)"

    # GUI columns and the acceptance summary are display only: the election never reads them.
    $guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
    foreach ($column in @('Header="Present (probe)" Binding="{Binding Capabilities}"', 'Header="Functional (acceptance)" Binding="{Binding FunctionalCapabilities}"', 'Header="Not functional" Binding="{Binding NotFunctionalCapabilities}"', 'Header="Acceptance" Binding="{Binding Acceptance}"')) {
        Assert-Display ($guiText.Contains($column)) "GUI column: $column"
    }
    Assert-Display (-not $guiText.Contains('nothing kept')) 'GUI no longer says that the acceptance test keeps nothing.'
    Assert-Display (-not ([IO.File]::ReadAllText((Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1'))).Contains('Orchestrator-Acceptance')) 'The election module does not read the acceptance result.'
    # The resident reads only GeneratedAtUtc, for the daily acceptance catch-up (orchestrator 1.5.81).
    $orchestratorText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'))
    $readers = @([regex]::Matches($orchestratorText, '(?ms)^function ([\w-]+)\s*\{(?:(?!^function ).)*?Orchestrator-Acceptance\.json') | ForEach-Object { $_.Groups[1].Value })
    Assert-Display (($readers -join ',') -eq 'Get-OrchestratorAcceptanceLastResultUtc') "Only the daily catch-up reads the result in the resident: $($readers -join ',')"
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

"[{0}] PASS: present / functional capability display, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAy8yh9x4/XrjpZ
# S83eewlTuVJKI1suPhLiUfnf6D9S5qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFOJboJmOjuUEj4rA2RJ141ZL99XRJdOoWMLZbisbeigMA0GCSqG
# SIb3DQEBAQUABIIBgA9+pwxlINdqf87ztdS3fDg2uNURvYamEz7EQ24FPi9mdymn
# 06Z8bM9umi1Jrwx+Y+488dfBOXoUCNtBVCpIuRoUdclCWYuIpqjJmgrMNzIkWf/H
# LGJIA4o11n1eX3qfk2YqbEiDCJPosJSNh1uGxrJSjeiunXMVuhjYFIs9598qS9p6
# 1D0YF1CmyKj4Hr4PfmtPxVTXbEHa+d0nZ1kzx3jJkHakwpXEe5JNPh1DTo5MHM5c
# LVaA5CwWCfSHRqLPGs7fI1EJrk6X6alCxwMP95YOjWngTWN2TRReIPZ195lUfG9O
# KJ/4CFN1FNiR3zDXR3cvHKAf2Nb9E1NiYHI3HHJ7fbGofpoYiQPvPGQSNekd2Mvz
# +0urhk5W+j0r1TX12hOUnVMd9viFrfMqcNXxFZLTrejBb9jmLdDWwy89vm14EalZ
# xeMn8aFdvxcgxqf8dZ7HkJsnQvqWgBGRzkBQ2aBlQywiiZNo0jPz9u0+69+wwj66
# nBZBpCSiCzzt+n/iRaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMTEx
# MjBaMC8GCSqGSIb3DQEJBDEiBCCrpffFXlw1T9drEbOcxMZVolOs+PfD+4DpXLDS
# VR6wlzANBgkqhkiG9w0BAQEFAASCAgB0L8zuGHyaGHLvzr+6W0WF3tE5OKr6x5NZ
# 0YJ8Vl2Uvi3Gnx4lEPQcuPgw6DqKDmI7hkQUq0pLl3LS+pX3Dpem0WKuLFT7seTn
# SpGlaAUNe9EyrFio6iragtshnvB3e4gNQ0WbFMtc4N3Ig5JFVv5044Rmc0tdDn+X
# Guxh/3mmoS0g+b6dWZv8gtmnQQb+ORwc9c0oW6/LLvoPoqUj9V+CtWvDxtfGnkj9
# FNGS1nO/+ryWEfW7mhqQDVVfHxzoQtGJL484PUW2jrV9e/1CiUDBrYkO0pcAb7rv
# 5E0PhgK+N53cB3qaoFjf5U3aSF65H2hhCEcs8PGwtNgJ764mQ34nMo13tJAx1CVG
# JfwQ9BmfNioe+h+YrDoJIotK8QTN7mVEfskfvG7pxlogJV5ViJMRfCSDRrCehjn7
# AxdRgTp8K1y6l87zQkkwF1cgjBYyF69rejENVCTU1REw1qXjRuWVLHeehLi/6SJ7
# NWHTHowmYfEv2qZo9oQUgqgv2ehT5jfk7erypVfZrw5ZMqeEdFhelwN4evKDOvFN
# V5c0S+rm4jzkpdFy2e7bJ1TIA1KDN5YHkuKZrTXUFr9JJpiPwGuYewhRLRUHAt7H
# t2jyU5kF7U+choHrn9akkROHPSc2+3od6ilH0o1cux0zuGzOwzG/iH5Lk19srN5B
# SY/m9Nky/Q==
# SIG # End signature block
