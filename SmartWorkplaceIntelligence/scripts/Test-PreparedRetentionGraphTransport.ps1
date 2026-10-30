<#
.SYNOPSIS
    Real-transport retention check: the installed Invoke-MgGraphRequest against a local HTTP listener.
.DESCRIPTION
    Uses a synthetic access token and a localhost listener. No tenant, Graph endpoint or network
    access outside this machine. Requires PowerShell 7 and Microsoft.Graph.Authentication; reports
    NOT RUN with exit code 3 when the SDK is absent instead of passing.
    -GraphAuthenticationVersion selects an installed SDK version (for example 2.30.0, the server
    version). When the SDK session type is reachable (2.30.0), the 2026-10-22 incident is replayed:
    after an absolute-URI request the session is switched to an app-only certificate context with
    an absent thumbprint, so a client rebuild fails exactly as on the server, without any network.
.VERSION
    1.1.0
#>
[CmdletBinding()]
param([string]$GraphAuthenticationVersion)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'PowerShell 7 is required.'}
$sdk=@(Get-Module -ListAvailable Microsoft.Graph.Authentication | Where-Object {-not $GraphAuthenticationVersion -or $_.Version -eq [version]$GraphAuthenticationVersion})
if(-not $sdk.Count){Write-Host "NOT RUN: Microsoft.Graph.Authentication $GraphAuthenticationVersion is not installed.";exit 3}
if($GraphAuthenticationVersion){Import-Module Microsoft.Graph.Authentication -RequiredVersion $GraphAuthenticationVersion -ErrorAction Stop}
else{Import-Module Microsoft.Graph.Authentication -ErrorAction Stop}
$modulePath=Join-Path $PSScriptRoot 'PreparedSharePointRetention.psm1'
Import-Module $modulePath -Force
$module=Get-Module PreparedSharePointRetention
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}

$port=Get-Random -Minimum 20000 -Maximum 45000
$prefix="http://localhost:$port/"
$listener=[Net.HttpListener]::new();$listener.Prefixes.Add($prefix);$listener.Start()
$responses=[Collections.Concurrent.ConcurrentQueue[int]]::new()
$seen=[Collections.Concurrent.ConcurrentQueue[object]]::new()
$runspace=[runspacefactory]::CreateRunspace();$runspace.Open()
$server=[powershell]::Create();$server.Runspace=$runspace
$null=$server.AddScript({param($Listener,$Responses,$Seen)
    while($Listener.IsListening){
        try {$context=$Listener.GetContext()} catch {break}
        $status=0;if(-not $Responses.TryDequeue([ref]$status)){$status=500}
        $Seen.Enqueue([pscustomobject]@{Method=$context.Request.HttpMethod;Path=$context.Request.RawUrl;IfMatch=$context.Request.Headers['If-Match']})
        $context.Response.StatusCode=$status
        $body=switch($status){
            200 {'{"id":"item"}'}
            # Body observed on 2026-10-29 when a Purview retention policy covered the site.
            403 {'{"error":{"code":"accessDenied","message":"Request was cancelled by event received. If attempting to delete a non-empty folder, it''s possible that it''s on hold"}}'}
            default {$null}
        }
        if($body){
            $bytes=[Text.Encoding]::UTF8.GetBytes($body);$context.Response.ContentType='application/json'
            $context.Response.OutputStream.Write($bytes,0,$bytes.Length)
        }
        $context.Response.Close()
    }
}).AddArgument($listener).AddArgument($responses).AddArgument($seen)
$handle=$server.BeginInvoke()
function Get-NextRequest {$item=$null;if($seen.TryDequeue([ref]$item)){$item}else{$null}}
function Invoke-Delete([string]$Uri,[string]$ETag){& $module {param($u,$e) Invoke-RetentionGraphDelete $u $e} $Uri $ETag}
try {
    # Synthetic unsigned token: Connect-MgGraph -AccessToken performs no network call.
    $claims=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"tid":"00000000-0000-0000-0000-000000000001","exp":4102444800}')).TrimEnd('=').Replace('+','-').Replace('/','_')
    $token=ConvertTo-SecureString ('eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.'+$claims+'.') -AsPlainText -Force
    Connect-MgGraph -AccessToken $token | Out-Null
    # Test setup only: one listener response per request, so SDK retries must not consume queued statuses.
    Set-MgRequestContext -MaxRetry 0 | Out-Null
    $version=(Get-Module Microsoft.Graph.Authentication).Version
    & $module {Assert-RetentionGraphDeleteSupport}
    Check $true 'Installed Invoke-MgGraphRequest exposes the retention DELETE parameters'
    $etag='"{6A5B1C2D-0000-0000-0000-000000000001},3"'
    $uri=$prefix+'v1.0/drives/b%21synthetic/items/01SYNTHETICITEM'
    # Same order as the incident run: an absolute-URI GET first, then the DELETE, without touching the request context.
    $responses.Enqueue(200)
    $null=Invoke-MgGraphRequest -Method GET -Uri ($prefix+'v1.0/drives/b%21synthetic/items/01SYNTHETICITEM') -OutputType PSObject -ErrorAction Stop
    $null=Get-NextRequest
    $environment='unavailable'
    try {$environment=[string][Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.Environment.AzureADEndpoint} catch {Write-Verbose 'GraphSession is not reachable from this load context.'}
    foreach($case in @(
        @{Status=204;Outcome='Recycled';Reached=$false},
        @{Status=404;Outcome='NotFound';Reached=$false},
        @{Status=412;Outcome='PreconditionFailed';Reached=$false},
        @{Status=429;Outcome='HttpError';Reached=$false},
        @{Status=500;Outcome='HttpError';Reached=$true})){
        $responses.Enqueue($case.Status)
        $result=Invoke-Delete $uri $etag
        $request=Get-NextRequest
        Check ($null -ne $request -and $request.Method -ceq 'DELETE' -and $request.Path -ceq '/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM' -and $request.IfMatch -ceq $etag) "HTTP $($case.Status): the real cmdlet did not send the exact conditional DELETE"
        Check ($result.Outcome -ceq $case.Outcome -and $result.HttpStatus -eq $case.Status -and $result.MayHaveReachedService -eq $case.Reached) "HTTP $($case.Status): -StatusCodeVariable outcome was $($result.Outcome)/$($result.HttpStatus) instead of $($case.Outcome)/$($case.Status)"
    }
    # A refusal keeps the Graph error code and message in the journaled error, with the retention-hold hint.
    $responses.Enqueue(403)
    $result=Invoke-Delete $uri $etag
    $null=Get-NextRequest
    Check ($result.Outcome -ceq 'HttpError' -and $result.HttpStatus -eq 403 -and -not $result.MayHaveReachedService) "HTTP 403: outcome was $($result.Outcome)/$($result.HttpStatus)"
    Check ($result.Error -match 'HTTP 403' -and $result.Error -match 'Graph error: accessDenied: Request was cancelled' -and $result.Error -match 'Purview retention policy or hold') "HTTP 403: Graph error detail missing from '$($result.Error)'"
    # Without a body the message stays the plain status text.
    $responses.Enqueue(429)
    $result=Invoke-Delete $uri $etag
    $null=Get-NextRequest
    Check ($result.Error -ceq 'Graph DELETE returned HTTP 429; only 204 confirms removal.' -and $result.GraphError -ceq '') "HTTP 429 without body: unexpected error text '$($result.Error)'"
    # Regression: a URI the SDK cannot use fails before sending, is reported explicitly and is never 'in flight'.
    # A relative string is unsafe here: the SDK resolves it against the host of the previous absolute request.
    $result=Invoke-Delete 'https://graph.microsoft.com:99999/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM' $etag
    Check ($result.Outcome -ceq 'RequestFailed' -and -not $result.MayHaveReachedService -and $result.Error -match 'UriFormatException|Invalid URI|URI') 'Invalid URI was not reported as a pre-send failure'
    Check ($null -eq (Get-NextRequest)) 'Invalid URI request reached the listener'
    # A refused connection stays uncertain: the module cannot prove the request was not accepted.
    $closed=Get-Random -Minimum 45001 -Maximum 60000
    $result=Invoke-Delete ("http://localhost:$closed/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM") $etag
    Check ($result.Outcome -ceq 'RequestFailed' -and $result.MayHaveReachedService) 'Connection failure was not kept uncertain'
    # Incident replay, last because it leaves the SDK session unusable.
    $sessionType='Microsoft.Graph.PowerShell.Authentication.GraphSession' -as [type]
    $replay='not run (SDK session type not reachable in this version)'
    if($sessionType){
        $session=$sessionType::Instance
        Check ([string]::IsNullOrEmpty([string]$session.Environment.AzureADEndpoint)) 'Expected the SDK session environment to lose its authority after an absolute-URI request'
        $context=$session.AuthContext
        $context.AuthType=[Microsoft.Graph.PowerShell.Authentication.AuthenticationType]::AppOnly
        $context.TokenCredentialType=[Microsoft.Graph.PowerShell.Authentication.TokenCredentialType]::ClientCertificate
        $context.ClientId='00000000-0000-0000-0000-0000000000aa';$context.CertificateThumbprint=('0'*40)
        # Fixed path: the cached client is kept, so the DELETE is sent.
        $responses.Enqueue(204)
        $result=Invoke-Delete $uri $etag
        Check ($result.Outcome -ceq 'Recycled' -and $null -ne (Get-NextRequest)) 'Fixed DELETE path did not reach the listener after the session lost its authority'
        # Former path: a request-context change discards the client; the rebuild must fail before sending.
        Set-MgRequestContext -MaxRetry 0 | Out-Null
        $result=Invoke-Delete $uri $etag
        Check ($result.Outcome -ceq 'RequestFailed' -and -not $result.MayHaveReachedService -and $result.Error -match 'UriFormatException: Invalid URI') "Incident replay did not reproduce a pre-send Invalid URI: $($result.Error)"
        Check ($null -eq (Get-NextRequest)) 'Incident replay sent a request'
        $replay='reproduced: Set-MgRequestContext + client rebuild raises Invalid URI before sending; the fixed path sends the DELETE'
    }
    Write-Host "PASS: $script:checks real-transport checks with Microsoft.Graph.Authentication $version (synthetic token, localhost listener). Session AzureADEndpoint after an absolute-URI request: '$environment'."
    Write-Host "Incident replay: $replay."
} finally {
    try {Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null} catch {Write-Verbose 'No Graph context to disconnect.'}
    $listener.Stop();$listener.Close()
    try {$null=$server.EndInvoke($handle)} catch {Write-Verbose 'Listener stopped.'}
    $server.Dispose();$runspace.Dispose()
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBs4wk2cNFDWhCZ
# MVhL96eywWjfPYaiDcu8AdlAuQ2946CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBgNO7hrVTmLbXxh1FFzuDBs+uAk3K9BdrdP+ihsdE/2MA0GCSqG
# SIb3DQEBAQUABIIBgCRUywpIPb87UbTuqlFLAdAeRach7JWNiPuOEEnufA6+1vN3
# 6lB/ZtYeb9wYvl6D+NsmSyeD6cOD1HYaH5w5sGKrnJqVWW9C6UAlE0F2kFNu21aI
# Q1KN8ezAGeV37EljdwW1+qSKhuEw1Un4yJCN1P2dENI1mpGFlhmWPv7MUM5AHnda
# ZkfNUXojdXc9GaW3PROWTTDk+S6E/B+FaDdy5Kd3V51c++POYovD3Jsbc58j9+Gj
# YJU/B9Vlqe7XH+uOS/W1IJ9HAUSBcb+QxvelAEoEvtVU5r1cOCwCXxip3lXHwPC2
# Sxj2YdP7wNj23t7FJHI0lg6ywbgBULWZDhJt6bSLSydpw86ku7rXmPjBJ2ooXh7F
# 40XZi1Nvba9OQJop1Xm1rAct1Mwu2ux9XpuGywaE/Zjt3dM087JAYlRXpGGEmikV
# rs/SBeCKqMnXXJv3Zt/dT8tfySOoz4CxudSba5lidRrjw8zzsXwl2wyU5iP7aIUV
# PUKwwmsnmswOrjqSeaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwMDUw
# NTZaMC8GCSqGSIb3DQEJBDEiBCDQgit8vxArTSn6QPwLzMin4qzx4rTDXVa+H3fS
# OTUS4DANBgkqhkiG9w0BAQEFAASCAgAzB9gEEDecgtf//BJuUq4kufDbsgumqng9
# tEi7KQwR1aP9tMFvR6BxQcqOKhLBGTNEhIf3vKPrZDcq5MIOlbtwV+8g3WMGZCLJ
# Nqt32//KCNcLxgTbsn4U2UFhzbWANIE7hrjkwqE90z1ju/atAXBw4AVu66z7q5eD
# oThWiVxElx7RdyF8l66asP3EprgW5zgSnS9CmMvKgeirv10ODBKdYoa0bSV2uXoo
# k2YM0pMX060UCAwAa4oMKEsdNK3Y5YMLQvwHW5vaPifOFDiZO8qh5Mar3qOZ0k11
# d6RPf1Wrf+biHCI5P07hGlBmJWgsimfbxM9AAiOfzWj1G+K2JUhJf8mXNcMQ/0Wq
# sq6NNVeFUIqlD24rKVdXZrmgCVKeHGqX5XBzV4o+p+P4GihuOuo5gXGtR/p1rr9a
# 947lG/c9lhiPk+z2TxdD2x7BpOh92m+SHk8aSHgsEuJV/rXgzC7C/ef39q9F5xcv
# Qt8O7wFuXfmnn2P8Nn5CRFjoBknO3xbFXCZ6CPyovms1VoXhPAuC6VgJUyiszcT8
# pjwEASD1ucd62mUz7+8rWSaJBzJLKxrsHtnvS165DgCik/f3lTgYl5YlLkBUqRl+
# A1PR1N9IBtf6AZddbPDL/c44VJ8sm2WIQNkVPkLFvHOKnv/f4yEYVvfPjA/SZD7C
# YrX4Y5Vxuw==
# SIG # End signature block
