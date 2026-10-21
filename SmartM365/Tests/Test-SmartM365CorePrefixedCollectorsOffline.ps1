#Requires -Version 7.0
<#
.SYNOPSIS
Offline check of the collectors that import SmartM365.Core with -Prefix Core (Endpoint Analytics,
Autopatch): their Graph retry and export download paths run against the Core module imported exactly
as in production, without any Core function loaded by hand, and resolve every Core call.
.VERSION
1.0.0
.NOTES
The export download reads a loopback HTTP URL through the real Core helper. Graph calls are replaced by
synthetic errors. No tenant, Graph, storage or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Retry waits are recorded instead of slept.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
# Same import as the collectors: prefixed only. No unprefixed Core command may resolve in this process.
Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Prefix Core -Force -WarningAction SilentlyContinue
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
Check ($null -eq (Get-Command Get-SmartM365GraphErrorStatusCode -ErrorAction SilentlyContinue) -and $null -ne (Get-Command Get-CoreSmartM365GraphErrorStatusCode -ErrorAction SilentlyContinue)) 'The test process does not reproduce the prefixed Core import.'

function Import-CollectorFunction {
    param([Parameter(Mandatory)][string]$RelativePath, [Parameter(Mandatory)][string[]]$Names, [string]$Prelude)
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root $RelativePath), [ref]$null, [ref]$null)
    $text = foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) { throw "$name is missing in $RelativePath." }
        $node.Extent.Text
    }
    New-Module -ScriptBlock ([scriptblock]::Create($Prelude + "`n" + ($text -join "`n")))
}
$standIns = @'
$script:Answers = [Collections.Generic.Queue[object]]::new()
$script:Calls = 0
function Start-Sleep { param([int]$Seconds) $null = $Seconds }
function Write-Log { param($Message, $Level) $null = $Message, $Level }
function Write-EALog { param($Message, $Level) $null = $Message, $Level }
function Next-Answer { $script:Calls++; $answer = $script:Answers.Dequeue(); if ($answer -is [Exception]) { throw $answer }; $answer }
function Set-Answers([object[]]$Items) { $script:Answers.Clear(); $script:Calls = 0; foreach ($item in $Items) { $script:Answers.Enqueue($item) } }
'@
$timeout = [System.Threading.Tasks.TaskCanceledException]::new('The request was canceled due to the configured HttpClient.Timeout of 100 seconds elapsing.')
$sdk500 = [Exception]::new('Response status code does not indicate success: InternalServerError (Internal Server Error).')

# A real export package served from a local file:// URL.
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-CorePrefixed-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
try {
    $csv = Join-Path $temporary 'report.csv'
    [IO.File]::WriteAllText($csv, "DeviceId,Name`r`nd1,Device 1`r`n")
    $zip = Join-Path $temporary 'report.zip'
    Compress-Archive -Path $csv -DestinationPath $zip
    # Minimal HTTP server on the loopback address: one response per download, then it stops.
    $zipBytes = [IO.File]::ReadAllBytes($zip)
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $server = [powershell]::Create()
    [void]$server.AddScript({
        param($Listener, [byte[]]$Body, [int]$Count)
        for ($i = 0; $i -lt $Count; $i++) {
            $client = $Listener.AcceptTcpClient()
            try {
                $stream = $client.GetStream()
                $buffer = [byte[]]::new(8192)
                $null = $stream.Read($buffer, 0, $buffer.Length)
                $header = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: application/zip`r`nContent-Length: $($Body.Length)`r`nConnection: close`r`n`r`n")
                $stream.Write($header, 0, $header.Length); $stream.Write($Body, 0, $Body.Length); $stream.Flush()
            } finally { $client.Close() }
        }
    }).AddArgument($listener).AddArgument($zipBytes).AddArgument(2)
    $null = $server.BeginInvoke()
    $zipUrl = 'http://127.0.0.1:{0}/exports/report.zip?sv=synthetic&sig=synthetic' -f ([Net.IPEndPoint]$listener.LocalEndpoint).Port

    # ---- Endpoint Analytics: status, retry and download paths call Core with the Core prefix ----
    $ea = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1' -Names @('Get-EAStatusCode', 'Get-EARetryAfterSeconds', 'Invoke-EAGraphRequest', 'Invoke-EAWebDownload') -Prelude ($standIns + "`n`$MaxRetryCount = 3")
    try {
        $r = & $ea { param($a) Set-Answers $a; $v = Invoke-EAGraphRequest GET 'https://graph.microsoft.com/beta/x' -NoSleep -Invoker { Next-Answer }; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($sdk500, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Endpoint Analytics did not classify a Graph SDK 500 through the prefixed Core.'
        $r = & $ea { param($a) Set-Answers $a; $v = Invoke-EAGraphRequest GET 'https://graph.microsoft.com/beta/x' -NoSleep -Invoker { Next-Answer }; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($timeout, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Endpoint Analytics did not retry a timeout through the prefixed Core.'
        $target = Join-Path $temporary 'ea.zip'
        & $ea { param($u, $o) Invoke-EAWebDownload $u $o } $zipUrl $target
        Check ((Get-FileHash -LiteralPath $target).Hash -eq (Get-FileHash -LiteralPath $zip).Hash) 'Endpoint Analytics did not download the export through the prefixed Core.'
    } finally { Remove-Module $ea -Force }

    # ---- Autopatch: status, retry and the export import path ----
    $ap = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1' -Names @('Get-AutopatchGraphStatusCode', 'Get-AutopatchGraphRetryDelaySeconds', 'Invoke-AutopatchGraphRequest', 'Import-ExportedCsv') -Prelude ($standIns + @'

$script:GraphTransientRetryCount = 0
function Invoke-MgGraphRequest { [CmdletBinding()] param($Method, $Uri, $Body, $ContentType, $OutputType) Next-Answer }
function Ensure-Folder { param($Path) New-Item -ItemType Directory -Path $Path -Force | Out-Null }
function Start-ExportJob { param($ReportName, $Select, $Filter) [pscustomobject]@{ id = [guid]::NewGuid().ToString('N') } }
function Wait-ExportJob { param($JobId) [pscustomobject]@{ status = 'completed'; url = $script:ZipUrl } }
'@)
    try {
        $r = & $ap { param($a) Set-Answers $a; $v = Invoke-AutopatchGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/x'; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($sdk500, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Autopatch did not classify a Graph SDK 500 through the prefixed Core.'
        $r = & $ap { param($a) Set-Answers $a; $v = Invoke-AutopatchGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/x'; [pscustomobject]@{ Value = $v; Calls = $script:Calls } } @($timeout, 'ok')
        Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Autopatch did not retry a timeout through the prefixed Core.'
        $rows = & $ap { param($u) $script:ZipUrl = $u; @(Import-ExportedCsv -ReportName 'SyntheticAutopatch') } $zipUrl
        Check (@($rows).Count -eq 1 -and $rows[0].DeviceId -eq 'd1') 'Autopatch did not import the export through the prefixed Core.'
    } finally { Remove-Module $ap -Force }
}
finally {
    if (Get-Variable -Name listener -ErrorAction Ignore) { $listener.Stop() }
    if (Get-Variable -Name server -ErrorAction Ignore) { $server.Dispose() }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks prefixed Core collector checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAr/S8SM33hHsN6
# Zhij0cPmhSaXZNDwPGT6QFOeykVJCaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH/ziIs/8T9L6zZZ+Nn4zFKg2ja2qd1QnUFkIY9geUwdMA0GCSqG
# SIb3DQEBAQUABIIBgC5CM2unR9mQdqRjNjuewOsGJ+LNEhF1imxSfxuTgQx34tlI
# C5myIr82UmURjFfe7is34cpsPtQQZRHczndgSPrg5ooM97FW2Ojq7vrGu/XXCXRD
# ehIghMsxEOXEKc0kV0pgkvj35ycIHrlcBjYd0gRcDdHxtPPogWpdB1me3+sw80kQ
# xqT3L0VU7dDXUVbabIMllSXLrl6gogDLMFWahg9Naksc5z1Y4xo5WxsvbLgxZfwG
# njlPj92mDg7iA+XsO1hi9cx0wTNGkKwZPusqZsSEIHPM2OZB8SxEagxtuycWF2xU
# 038jZF/cr9u0yxtCLPNbTOTNEdkedg/7UO8sDI5rWXQONNHSFCUJOTM/X66gjm8I
# qb9/vQ4NK2/4vgh9nH1o4CZDo4df+cC0SknKvqplhz0cLzlQND4fsQ0H8D9jcoKE
# WzJkshEaNyWWzBKmUhwJHtoKqexkWMURZeU/7wh/GFL6FsuDNRTRmbWsjueqibb0
# Ob+f1/aOTljlvQQ4sqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExMzQ0
# NDZaMC8GCSqGSIb3DQEJBDEiBCAUSjQeyq/LJDjrzFHmoB5WDkcTpYX6+Nprffcl
# B/OLlzANBgkqhkiG9w0BAQEFAASCAgCcXbo1juSpOvxa197QG7zeEfW0oBkaFHh9
# ePjhOHH9KtQZIbio3XpaiIYU/kQ1MC+ygMsDXrkLrHVA3N55KGbaetrfEeFMDFYf
# eMNod55aA1Jn/r+VzF3PyFZ+onKVlvn+gh+xlLbN17KW9gBSMEdg72BMh+ylBsEg
# 9Fo3DV4h+/AqfIIDfJIlMlaWt1XrW1hHYlzzcLZLTkEVeKi3xDsDx9do4iWUSIcG
# FtIlsIcvrLUOm2LesYkXpAXTMGCYYASPniR0p64zKiVGjOPB6+3zkh1cjm34xO7P
# EzSFh/TiMGx80fMdz5wsA5IuuQ88oHoBGh+sAUvglAlbOZXq1buhZfGsNGE7ElIA
# tFbb1daq9Yvs2dj1Y6/iyKAFSJ3//lATSpg7zH+lSC2V0W4CScWdFeoS3XWounjH
# dL4AvBrDCnv0jQ6AxdwK2FH+Y0HcKsRwCaUvDLlpMslIgwpsJSbs7mBTLjQh33R7
# IbjOMJy/ADTJuOwEB92sG45nRF2M1/Hr+I8B8/dFTAi0usT9u5WZdCdIPACC2YJM
# NjQuHV0RGymOJL9ZNV25VeQ147Quvyloj4U3uTUvEEP7x2hilbeGExVPcPphfmpu
# ijjkRopdgN39/6W14H+y1YY8xriQMCUx9WvP+RetG00x7InevYr8sj75s3U4tXH8
# 6jPwXUwVXg==
# SIG # End signature block
