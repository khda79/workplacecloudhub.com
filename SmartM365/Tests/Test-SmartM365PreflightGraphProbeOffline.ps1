#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the preflight Graph probe and of the orchestrator cloud capability probes: transient
Graph errors (5xx, 429 with Retry-After) are retried, a persistent one is reported as a transient
service error distinct from a permission failure, and 401/403 fail at once without retry.
.VERSION
1.0.0
.NOTES
Graph calls are replaced by synthetic HTTP responses. No tenant, Graph, certificate or network access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Recorders accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic module test state only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='The probe attempt setting is saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.97' -Force -PassThru
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$probeUri = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=1'
$savedAttempts = Get-Variable -Name GraphProbeMaxAttempts -Scope Global -ErrorAction Ignore
$global:GraphProbeMaxAttempts = $null

& $core {
    $script:Calls = New-Object 'Collections.Generic.List[object]'
    $script:Sleeps = New-Object 'Collections.Generic.List[int]'
    $script:Logs = New-Object 'Collections.Generic.List[string]'
    $script:Answers = New-Object 'Collections.Generic.Queue[object]'
    # Each answer is a status code (0 = success) or an exception to throw.
    function script:New-SyntheticHttpFailure {
        param([int]$Status, [string]$RetryAfter)
        $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
        if ($RetryAfter) { $response.Headers.Add('Retry-After', $RetryAfter) }
        [Microsoft.PowerShell.Commands.HttpResponseException]::new(("Response status code does not indicate success: {0} ({1})." -f $Status, $response.ReasonPhrase), $response)
    }
    function script:Invoke-SyntheticAnswer {
        param([string]$Kind, [string]$Uri, $TimeoutSec)
        $script:Calls.Add([pscustomobject]@{ Kind = $Kind; Uri = $Uri; TimeoutSec = $TimeoutSec })
        $answer = $script:Answers.Dequeue()
        if ($answer -is [Exception]) { throw $answer }
        [pscustomobject]@{ value = @() }
    }
    function script:Invoke-RestMethod { [CmdletBinding()] param($Method, $Uri, $Headers, $TimeoutSec) Invoke-SyntheticAnswer -Kind 'Rest' -Uri $Uri -TimeoutSec $TimeoutSec }
    function script:Invoke-MgGraphRequest { [CmdletBinding()] param($Method, $Uri, $OutputType) Invoke-SyntheticAnswer -Kind 'Mg' -Uri $Uri }
    function script:Get-MgContext { [CmdletBinding()] param() [pscustomobject]@{ ClientId = ''; Scopes = @() } }
    function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
    function script:WriteLog { param($Message, $Level) $script:Logs.Add(("{0} {1}" -f $Level, $Message)) }
}
function Invoke-Preflight([object[]]$Answers, [switch]$UseGraphSdk) {
    & $core {
        param($Answers, $UseGraphSdk, $Uri)
        $script:Calls.Clear(); $script:Sleeps.Clear(); $script:Logs.Clear(); $script:Answers.Clear()
        foreach ($answer in $Answers) {
            $script:Answers.Enqueue($(if ($answer -is [int] -and $answer -gt 0) { New-SyntheticHttpFailure -Status $answer } else { $answer }))
        }
        $token = if ($UseGraphSdk) { '' } else { 'synthetic.token.value' }
        $outcome = [pscustomobject]@{ Passed = $false; Error = '' }
        try { $outcome.Passed = [bool](Invoke-SmartM365Preflight -ScriptName 'Synthetic' -GraphProbeUris @($Uri) -GraphAccessToken $token) }
        catch { $outcome.Error = $_.Exception.Message }
        [pscustomobject]@{ Passed = $outcome.Passed; Error = $outcome.Error; Calls = $script:Calls.Count; Sleeps = ($script:Sleeps -join ','); Logs = $script:Logs.ToArray(); Timeouts = @($script:Calls | ForEach-Object TimeoutSec) }
    } $Answers ([bool]$UseGraphSdk) $probeUri
}

try {
    # 1. 500 then success: the preflight passes after one progressive wait.
    $r = Invoke-Preflight @(500, 0)
    Check ($r.Passed -and $r.Calls -eq 2 -and $r.Sleeps -eq '15') "A transient 500 failed the preflight: $($r.Error)"
    Check (@($r.Logs | Where-Object { $_ -like 'WARNING Preflight Graph probe transient failure*Status=500; attempt 1/4*' }).Count -eq 1 -and @($r.Logs | Where-Object { $_ -like 'INFO Preflight Graph permission OK:*(after 2 attempts)' }).Count -eq 1) 'The retried probe is not logged.'
    Check (@($r.Timeouts | Where-Object { $_ -ne 180 }).Count -eq 0) 'The probe lost its request time limit.'

    # 2. 500 on every attempt: a persistent transient error, not a permission failure.
    $r = Invoke-Preflight @(500, 500, 500, 500)
    Check (-not $r.Passed -and $r.Calls -eq 4 -and $r.Sleeps -eq '15,30,45') 'A persistent 500 was not retried with progressive waits.'
    Check ($r.Error -match [regex]::Escape("Persistent transient Microsoft Graph error for '$probeUri' after 4 attempt(s) (Status=500); not a permission failure") -and $r.Error -notmatch 'Graph permission probe failed') "A persistent 500 is reported as a permission failure: $($r.Error)"

    # 3. 403: the permission failure is immediate and keeps its message.
    $r = Invoke-Preflight @(403)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Sleeps -eq '' -and $r.Error -match [regex]::Escape("Graph permission probe failed for '$probeUri'")) "A 403 was retried or reworded: $($r.Error)"
    $r = Invoke-Preflight @(401)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A 401 was retried.'

    # 4. 429 with Retry-After: the wait follows the service.
    $throttled = & $core { New-SyntheticHttpFailure -Status 429 -RetryAfter '7' }
    $r = Invoke-Preflight @($throttled, 0)
    Check ($r.Passed -and $r.Calls -eq 2 -and $r.Sleeps -eq '7') "A 429 did not honour Retry-After: $($r.Sleeps)"

    # 5. Graph SDK path: the status exists only in the message, as in the 17/10 failure.
    $sdk500 = [Exception]::new('Response status code does not indicate success: InternalServerError (Internal Server Error).')
    $r = Invoke-Preflight @($sdk500, 0) -UseGraphSdk
    Check ($r.Passed -and $r.Calls -eq 2) "A Graph SDK 500 failed the preflight: $($r.Error)"
    $sdk403 = [Exception]::new('Forbidden (Forbidden). Request timed out while authorizing.')
    $r = Invoke-Preflight @($sdk403) -UseGraphSdk
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A Graph SDK 403 was retried because its text mentions a timeout.'

    # 6. A 404 (wrong endpoint) is not transient: unchanged permission-probe failure.
    $r = Invoke-Preflight @(404)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A 404 was retried.'

    # 7. Orchestrator cloud capability probes: transient errors retried, authorization errors not.
    $distributed = [IO.File]::ReadAllText((Join-Path $root 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1'))
    $retryText = [regex]::Match($distributed, "(?s)\`$cloudProbeRetry = @'\r?\n(?<Body>.*?)\r?\n'@").Groups['Body'].Value
    Check ($retryText -match 'function Invoke-ProbeWithRetry') 'The capability probe retry helper is missing.'
    Check ($distributed -match "\`$graphScript = \`$cloudProbeRetry \+" -and $distributed -match "\`$exoScript = \`$cloudProbeRetry \+" -and $distributed -match 'Invoke-ProbeWithRetry \{ Connect-MgGraph' -and $distributed -match '(?s)Invoke-ProbeWithRetry \{\s+Connect-ExchangeOnline') 'Graph or EXO capability probes do not retry transient errors.'
    $probe = New-Module -ScriptBlock ([scriptblock]::Create($retryText + "`nfunction Start-Sleep { param([int]`$Seconds) `$script:Slept += `$Seconds }`n`$script:Slept = 0"))
    try {
        $state = & $probe { $script:n = 0; $value = Invoke-ProbeWithRetry { $script:n++; if ($script:n -lt 3) { throw 'Response status code does not indicate success: ServiceUnavailable (Service Unavailable).' }; 'ok' }; [pscustomobject]@{ Value = $value; Calls = $script:n; Slept = $script:Slept } }
        Check ($state.Value -eq 'ok' -and $state.Calls -eq 3 -and $state.Slept -eq 15) 'A transient capability probe error was not retried.'
        $state = & $probe { $script:n = 0; $message = ''; try { Invoke-ProbeWithRetry { $script:n++; throw 'An error occurred while sending the request.' } } catch { $message = $_.Exception.Message }; [pscustomobject]@{ Calls = $script:n; Message = $message } }
        Check ($state.Calls -eq 3 -and $state.Message -like 'Persistent transient error after 3 attempt(s): *') 'A persistent capability probe error is not reported as such.'
        $state = & $probe { $script:n = 0; try { Invoke-ProbeWithRetry { $script:n++; throw 'AADSTS700027: Client assertion failed; request timed out.' } } catch { $null = $_ }; $script:n }
        Check ($state -eq 1) 'An authorization error of a capability probe was retried.'
    }
    finally { Remove-Module $probe -Force }
}
finally {
    if ($savedAttempts) { $global:GraphProbeMaxAttempts = $savedAttempts.Value } else { Remove-Variable -Name GraphProbeMaxAttempts -Scope Global -ErrorAction Ignore }
    Remove-Module $core -Force -ErrorAction SilentlyContinue
}
"PASS: $script:checks preflight Graph probe checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCAL5dLXsmYTGI+
# B6BIX8kBiJFVptUO7/C61XGmeU4lcqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICsanSfdm5EUMNb23ZH7zKcyIq5oFZTf0PcsxpsFBZjFMA0GCSqG
# SIb3DQEBAQUABIIBgFndo334Xs/e3dYvI2pejAhVHttOqRFqY2pJpfsK15EMKjr9
# AASj8Zw8+EBsDIauOarPVDrlSu5rwiSPFdpwKPmZA68v9dJwRwnpke+xEc1bgcCW
# TiV6Te3NxkAWxzg5VOlwKSZZNLMI5lwoYcQo9mnr1vJJQANKFuXw4u2YBJfpYp36
# lasYEhSZwzUb4VVXYKEwwa5R9w1RkVnYu7dM9Dd7oZ8mtSqUdqqIra+0lvIijyXQ
# 5bUurmJx2U3CpAPp4O8Pp5ky/6lZ7wpf/2nU3kKukzerbcaSYq9T63QLpcrSMby3
# VbwYlEvvcXPjDDu9n3s/YTrmdps83RVnN4KmPjGfPbEJjsr2YC6SKFb14iHSLcdE
# p8PWnQLK/+WgD5Bg63CCJYGm+yPVqAe53DNawcNLgr13DOak34+yVhOs3wbvtxhR
# rNw15L1XtWnq7U9hU5S8HYvW6ZKxaipj+2cXB0OczKFRxlAhjqUFjk3Nt9ydtnYG
# P8otgSpqXf7dPpXfGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEwODEy
# MDZaMC8GCSqGSIb3DQEJBDEiBCBfPs1Z5+M5oQEr1XIieo9PbHnzGH93wjqKcVDE
# XxXxRDANBgkqhkiG9w0BAQEFAASCAgABYfvqeDJp73ko07RfX01H4WVFtZLk0ckP
# FbNrkfhG26strPyYmDCykexFQjf5h5rwz+S6XzpLY+iAtk6QOO7gITgHrAIpY4jL
# 3s1cBQtjzFUzqnnKqHIw1F1u/AN7L2sk8lur6fM8hYnAt+dHT/6lGBHOuPyO2fgr
# iL/9Qlp2HHlPdqnnNpQ6wbEyDsD0sJ97DY6ExudaUx2liFI4Ox6SfLL1U+YBOTPD
# QJxIrpT1z/hz75wlPH7rCdGszhjY6Pn0JgIsyu5/328AMFxhQoTO4GjOe9OJYnoX
# lgwvNVbUpa9pYugHGJreSeyy4u7fR8NLi24pC25F8vD5jlm50/3VUwMCOgrgQZ0j
# 4s+pxWD48myTSr1NE6bjfvepanTYLRw4IqB+tTlRXlZYTQ+mveenvfjiGyZUaP5k
# YKE3i+NTmlTfeko1DgYgbTjl8bSPnM720KQj7SozdzaJiJ4x4nZhYGUxZ3W26tbL
# BS2+SxNpR99KXV41F/fJoTpV0PGIW1RBugvIJ3N0RolgwAo4e/MlYkpTMAPARuyG
# OYkSoPkYmBpGYYvWNzdM6aCyKMC3pes6DADYm383RfM1H5dKzv0hTV8udNvufVTg
# wPaGYd0EUujm7ooZ6dPzpGA6dm7nc27v/Vh5y6qKkpPZjXVTtGmZyQRMv0PqBX3D
# ZhW4PW4nQQ==
# SIG # End signature block
