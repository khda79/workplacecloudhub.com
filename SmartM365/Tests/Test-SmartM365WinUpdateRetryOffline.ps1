#Requires -Version 7.0
<#
.SYNOPSIS
Exercise Windows Update polling and ZIP retries with the real prefixed Core module offline.
.VERSION
1.0.0
.NOTES
Only extracted collector functions run. HTTP, waits and random samples are stand-ins;
no tenant, SDK connection, collector, synchronized data or real network access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='HTTP, waits and randomness are controlled inside the test modules.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.103' -Prefix Core -Force -PassThru -Global -WarningAction SilentlyContinue
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function Get-SyntheticFailure([int]$Status, [string]$RetryAfter) {
    $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
    if ($RetryAfter) { $null = $response.Headers.TryAddWithoutValidation('Retry-After', $RetryAfter) }
    [Microsoft.PowerShell.Commands.HttpResponseException]::new("HTTP $Status synthetic failure", $response)
}
function Get-SyntheticRecord($Exception) { [Management.Automation.ErrorRecord]::new($Exception, 'Synthetic', 'NotSpecified', $null) }
$sourcePath = Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$null, [ref]$null)
$names = @('Get-GraphHttpStatusCode','Get-GraphRetryDelaySeconds','Test-GraphTransientError','Invoke-GraphRestMethodWithRetry','Get-IntuneExportJob','Download-ExportZip')
$definitions = foreach ($name in $names) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing collector function: $name" }
    $node.Extent.Text
}
$collector = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-WinUpdateRetry-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
try {
    Check ($null -eq (Get-Command Get-SmartM365RetryDelay -ErrorAction SilentlyContinue)) 'An unprefixed Core import hides collector prefix errors.'
    Check ((Get-Content $sourcePath -Raw) -match "-MinimumVersion '1.0.103' -Prefix Core") 'Collector does not require the shared retry helper version.'
    # Guard the actual touched request path; these functions must remain REST-only.
    foreach ($text in @($definitions) + @(& $core { ${function:Invoke-SmartM365ExportFileDownload}.ToString(); ${function:Get-SmartM365RetryDelay}.ToString() })) {
        Check ($text -notmatch 'Set-MgRequestContext|Invoke-MgGraphRequest|EvidenceCollector') 'The tested request path now depends on the Graph SDK request context.'
    }
    & $collector {
        $script:Answers = [Collections.Generic.Queue[object]]::new()
        $script:Sleeps = [Collections.Generic.List[int]]::new()
        function script:Write-Log { param($Message, $Level, $Stage) }
        function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
        function script:Invoke-RestMethod {
            [CmdletBinding()] param($Method, $Uri, $Headers, $Body, $ContentType)
            $answer = $script:Answers.Dequeue()
            if ($answer -is [Exception]) { throw $answer }
            $answer
        }
    }
    $errorRecord = Get-SyntheticRecord (Get-SyntheticFailure 503 '2')
    foreach ($attempt in 1..8) {
        $delay = & $collector { param($e,$a) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt $a } $errorRecord $attempt
        $floor = [math]::Min(300, 15 * [math]::Pow(2, $attempt - 1))
        Check ($delay -ge $floor -and $delay -le 300) "Polling floor or cap violated at attempt $attempt : $delay"
    }
    $delay = & $collector { param($e) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt 1 } (Get-SyntheticRecord (Get-SyntheticFailure 503 '90'))
    Check ($delay -eq 90) 'Polling shortened Retry-After above its local floor.'
    $samples = @(1..40 | ForEach-Object { & $collector { param($e) Get-GraphRetryDelaySeconds -ErrorRecord $e -Attempt 1 } (Get-SyntheticRecord (Get-SyntheticFailure 503)) })
    Check (($samples | Measure-Object -Minimum).Minimum -ge 15 -and ($samples | Measure-Object -Maximum).Maximum -le 30 -and @($samples | Sort-Object -Unique).Count -gt 1) 'Polling lost bounded jitter.'
    $result = & $collector {
        param($failure)
        $script:Answers.Enqueue($failure); $script:Answers.Enqueue([pscustomobject]@{status='completed'})
        $value = Get-IntuneExportJob -JobId 'synthetic' -Headers @{} -MaxAttempts 5
        [pscustomobject]@{ Status=$value.status; Sleeps=$script:Sleeps.ToArray() }
    } (Get-SyntheticFailure 504 '40')
    Check ($result.Status -eq 'completed' -and $result.Sleeps.Count -eq 1 -and $result.Sleeps[0] -eq 40) 'Real status wrapper did not retry 504 with Retry-After through prefixed Core.'
    foreach ($status in @(401,403)) {
        $retry = & $collector { param($e) Test-GraphTransientError -ErrorRecord $e } (Get-SyntheticRecord (Get-SyntheticFailure $status))
        Check (-not $retry) "Permanent HTTP $status is retried."
    }
    # ZIP wrapper calls the real Core downloader; mock only its HTTP boundary.
    & $core {
        $script:Answers = [Collections.Generic.Queue[object]]::new()
        $script:Sleeps = [Collections.Generic.List[int]]::new()
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:Calls = 0; $script:Authorized = $false
        $script:Sample = 0.0
        function script:Get-Random { param($Minimum,$Maximum) $script:Sample }
        function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
        function script:WriteLog { param($Message,$Level) $script:Logs.Add($Message) }
        function script:Invoke-WebRequest {
            [CmdletBinding()] param($Method,$Uri,$OutFile,$TimeoutSec,$Headers)
            $script:Calls++; if ($Headers) { $script:Authorized = $true }
            [IO.File]::WriteAllText($OutFile, 'partial')
            $answer = $script:Answers.Dequeue()
            if ($answer -is [Exception]) { throw $answer }
            [IO.File]::WriteAllText($OutFile, 'zip')
        }
    }
    $target = Join-Path $temporary 'report.zip'
    & $core {
        param($withHeader,$withoutHeader,$capped)
        $script:Answers.Enqueue($withHeader); $script:Answers.Enqueue($withoutHeader); $script:Answers.Enqueue($capped); $script:Answers.Enqueue('ok')
    } (Get-SyntheticFailure 503 '37') (Get-SyntheticFailure 503) (Get-SyntheticFailure 503 '900')
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip?sig=synthetic-secret' -DestinationZipPath $out } $target
    $state = & $core { [pscustomobject]@{Sleeps=$script:Sleeps.ToArray(); Calls=$script:Calls; Authorized=$script:Authorized; Logs=$script:Logs -join "`n"} }
    Check (($state.Sleeps -join ',') -eq '37,10,120') "ZIP ignored Retry-After, jitter lower bound or 120s cap: $($state.Sleeps -join ',')"
    Check ($state.Calls -eq 4 -and -not $state.Authorized -and [IO.File]::ReadAllText($target) -eq 'zip') 'ZIP wrapper failed to publish the recovered download without bearer.'
    Check ($state.Logs -match 'Retry-After' -and $state.Logs -match 'Backoff' -and $state.Logs -notmatch 'synthetic-secret') 'ZIP retry source missing or SAS query leaked.'
    & $core { $script:Sleeps.Clear(); $script:Answers.Enqueue([Exception]::new('HTTP 503')); $script:Answers.Enqueue('ok') }
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target
    Check ((& $core { $script:Sleeps[0] }) -eq 5) 'ZIP lower jitter bound did not reach 5s.'
    & $core { $script:Sleeps.Clear(); $script:Sample=1.0; $script:Answers.Enqueue([Exception]::new('HTTP 503')); $script:Answers.Enqueue('ok') }
    & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target
    Check ((& $core { $script:Sleeps[0] }) -eq 10) 'ZIP upper jitter bound did not reach 10s (lower bound is 5s).'
    & $core { $script:Sleeps.Clear(); $script:Calls=0; foreach ($i in 1..6) { $script:Answers.Enqueue([Exception]::new('HTTP 503')) } }
    $terminal = $null
    try { & $collector { param($out) Download-ExportZip -DownloadUrl 'https://storage.example.invalid/report.zip' -DestinationZipPath $out } $target } catch { $terminal = $_ }
    $state = & $core { [pscustomobject]@{Calls=$script:Calls; Sleeps=$script:Sleeps.ToArray()} }
    Check ($terminal -and $terminal.Exception.Message -match 'failed after 6 attempt' -and $state.Calls -eq 6 -and $state.Sleeps.Count -eq 5) 'ZIP retry attempts are not bounded or sleep after terminal failure.'
    Check (($state.Sleeps | Measure-Object -Maximum).Maximum -le 120 -and [IO.File]::ReadAllText($target) -eq 'zip') 'ZIP retry exceeded its cap or overwrote the previous successful file on failure.'
    Check (@(Get-ChildItem -LiteralPath $temporary -Filter '*.partial').Count -eq 0) 'Failed ZIP left partial files.'
    Write-Output "PASS: $script:checks Windows Update retry checks (offline)."
}
finally {
    & $core { Remove-Item Function:\Invoke-WebRequest,Function:\Start-Sleep,Function:\WriteLog,Function:\Get-Random -ErrorAction SilentlyContinue }
    Remove-Module $collector -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiH/8gBlnQnRlI
# WyS1kYhJI/W1ahHqnhYt9TOiDOrMXKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHQNxxOQ4ZUeWfwFXgcy+0wgcLuHTo0BkxFnXh3cegrFMA0GCSqG
# SIb3DQEBAQUABIIBgBAB8zisbLOrEJ3sdajo/WQfYgR/hrtbKFg0706QjLF6wKTe
# Lg++be0L0m9IJANPdQ4rLd/4tRlA9TlRlFuorXabKxxHWHp9pBx7oUdc3Lf9Qxi8
# bCzwi/lsY/ahS1lFG5amF8QPNJo3Eu33cfb8o3rmVhtBGv9h8uCUi0VspC51MsxO
# nQlhT9BaPYE+0oRRYpk3TJhV+vQ8L+pdAlIUtgNpgfjHJFTKq5ag1GbVWaBoo7on
# em8+FKcXM32Qv05Nf+4c/mbFgY6qCpn3AKQsA7M0LfhhSH+OvN2phRaIsh7FjvYt
# I1wLCbRUXXxFpUuxgOiSuTFDUTjuuI0VjNo+mp3h2PQcetnaCJ6hSvAdYu9bB2+N
# h0OewUtAa6K6SOWKwcIxLRg7c99UZ+3fLqNYkU670tSoDB7K7WU0ft2PL/5JJ1O2
# FNuSzsDzAACUyAcmLOZ/3rig/axLnGqjFmleMGAs6andbU9GppHsfkrbhZZE1Hpw
# rYPAZozoxh1v1jV7/qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxMDU0
# MjhaMC8GCSqGSIb3DQEJBDEiBCD0xiQDWtmpT1F/b5JJF3siAfPWA3PcQbg52cwQ
# SeOgwzANBgkqhkiG9w0BAQEFAASCAgB+GQn7SyH15TxYHAutzZuXvdQWBDDXvZ7D
# kkUoc+C8YCo05jb/4p/FPF1/ZTFMUuTOejKWzsu5MrL65z8USRu68TAP6X974Gnd
# aVlYl5ySlyxX2cX4GHE3C+JigVEKn2grXI++VVzFy1tT3NByFJJFx7sbqnDZmS1l
# Idi98VLCxatHPuffa8gtgb8GzUhjnX/qDzZaxcGF3t8Ubsic6uYIc32jxOItAfJY
# erUBw6QP8G9ye82wpVKhoopmNvJeI0mM7oFHWjpWfwWnrPFd7HrFsSg1SWb0NeRx
# OldPMofOVsvvvoBw6JimK2lkyGbXufzifc2t53FSzyqMcqKoLbZx/LXqulTk2rp/
# yH9tVJ/FTeihWulGUG3lsaXy7jkCmOqWtAUZ9Z4pewWohFb5IQ9jvZipch5M1HEv
# JBm6/r8mri3SXgSnLK/BIsStvzSkmL4A1zTHx65Jcc3XxihwMP1od+Q44e6Zl5Xo
# ZqNQCnDeyX41WE+vNd64fAbzapjOsrCHn95193IkyFtKDRf29P+12dXcnbXoRFtG
# ieGtlJiizVFRvd5/nT4eb2W0BRUeSzWaE/WuWlALehpuvPVEWIO98dCNETSpRU0Y
# zEXY0K3/RYhpXTH5ICwg6acpx03bJlk5q4JJHXHa6HMFyPqK+fLcrUNezhZi5bTT
# nA7bpfSk+w==
# SIG # End signature block
