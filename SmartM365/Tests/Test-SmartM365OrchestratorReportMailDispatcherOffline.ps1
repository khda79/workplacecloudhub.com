<#
.SYNOPSIS
Offline tests of the report mail queue dispatcher (orchestrator 1.5.93): one server sends the
shared report mail queue, the first Active live server by name; Inactive servers never send it.
.VERSION
1.0.0
.NOTES
Runs the real orchestrator functions extracted from the script with stubbed live servers, logs and
queue worker. No mail, tenant initialization, collector or shared folder is used.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}
foreach ($name in @('Test-OrchestratorLocalServerActive', 'Get-OrchestratorReportMailDispatcher', 'Invoke-OrchestratorReportMailQueue')) {
    . ([scriptblock]::Create((Get-FunctionText $name)))
}

$script:Logs = [Collections.Generic.List[object]]::new()
function Write-OrchestratorLog([string]$Message, [string]$Level = 'INFO') { $script:Logs.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
$script:QueueCalls = 0
function Invoke-SmartM365DeferredMail { $script:QueueCalls++; [pscustomobject]@{ Sent = 1; Pending = 0; Retry = 0; Busy = 0; Invalid = 0; Uncertain = 0 } }
$script:Live = @()
$script:LiveFailure = ''
function Get-LiveOrchestratorServerCapabilities {
    if ($script:LiveFailure) { throw $script:LiveFailure }
    foreach ($server in $script:Live) { [pscustomobject]@{ ServerName = $server } }
}
function Reset-Node([string]$Computer) {
    $env:COMPUTERNAME = $Computer
    $script:Logs.Clear(); $script:QueueCalls = 0
    foreach ($variable in @('ReportMailDispatcher', 'ReportMailDispatcherFallback', 'ReportMailDispatcherError')) { Remove-Variable -Name $variable -Scope Script -ErrorAction SilentlyContinue }
}
function New-Settings([hashtable]$States, [bool]$Explicit = $true, [bool]$Distributed = $true) {
    [pscustomobject]@{ DistributedSchedulingEnabled = $Distributed; ServerStates = $States; ServerStatesExplicit = $Explicit }
}

$originalComputer = $env:COMPUTERNAME
try {
    # Synthetic layout like production: SRV-A1, SRV-A2 and SRV-B1 Inactive, SRV-C1 and SRV-C2 Active.
    $states = @{ 'SRV-A1' = 'Inactive'; 'SRV-A2' = 'Inactive'; 'SRV-B1' = 'Inactive'; 'SRV-C1' = 'Active'; 'SRV-C2' = 'Active' }
    $script:Settings = New-Settings $states
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C1', 'SRV-C2')

    # 1. Exactly one server sends: the first Active live server by name.
    $senders = foreach ($server in $script:Live) {
        Reset-Node $server
        Invoke-OrchestratorReportMailQueue
        Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C1') "Every server must agree on SRV-C1 ($server)."
        if ($script:QueueCalls) { $server }
    }
    Check ((@($senders) -join ',') -eq 'SRV-C1') "Only SRV-C1 must send the queue (senders: $(@($senders) -join ','))."

    # 2. Inactive servers never send, even when first by name; a non-dispatcher logs why once.
    Reset-Node 'SRV-A1'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'An Inactive server must not send the queue.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1; this server does not send*').Count -eq 1) 'A non-dispatcher must log the dispatcher once.'

    # 3. The dispatcher sends at every tick and logs its role once, then the queue counts.
    Reset-Node 'SRV-C1'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2) 'The dispatcher must send the queue at every tick.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: this server*').Count -eq 1) 'The dispatcher role must be logged once.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue: {"Sent":1*').Count -eq 2) 'Queue counts must still be logged.'

    # 4. Failover: a stale or deactivated dispatcher is replaced at the next tick, and given back.
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'SRV-C2 must wait while SRV-C1 is live.'
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C2')
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 1 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: this server*').Count -eq 1) 'SRV-C2 must take over when SRV-C1 is no longer live.'
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C1', 'SRV-C2')
    $states['SRV-C1'] = 'Inactive'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2) 'SRV-C2 must keep the queue while SRV-C1 is Inactive.'
    $states['SRV-C1'] = 'Active'
    Invoke-OrchestratorReportMailQueue
    # Second SRV-C1 line of this node: the first one was logged before the failover.
    Check ($script:QueueCalls -eq 2 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1*').Count -eq 2) 'The queue must return to SRV-C1 once it is Active and live again.'

    # 5. Names are compared without case; a server absent from explicit states is Inactive.
    $script:Live = @('srv-c2', 'SRV-Z9')
    Reset-Node 'SRV-C2'
    Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C2') 'Live names must be compared without case, and an unlisted server must be Inactive with explicit states.'
    $script:Settings = New-Settings @{} $false
    Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C2') 'Without explicit states every live server is Active.'
    $script:Settings = New-Settings $states

    # 6. No Active live server: nobody sends and the wait is a warning.
    $script:Live = @('SRV-A1', 'SRV-B1')
    Reset-Node 'SRV-B1'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like '*dispatcher: none*' }).Count -eq 1) 'Without an Active live server nobody sends and a WARN is logged.'

    # 7. Unreadable peer documents: an Active server keeps sending, an Inactive one does not; one WARN each.
    $script:Live = @('SRV-C1', 'SRV-C2')
    $script:LiveFailure = 'Shared folder unreachable'
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2 -and @($script:Logs | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like '*could not be determined (Shared folder unreachable); this Active server sends it*' }).Count -eq 1) 'An Active server must keep sending with one WARN when peers are unreadable.'
    Reset-Node 'SRV-A1'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object Message -like '*this Inactive server does not send it*').Count -eq 1) 'An Inactive server must not send when peers are unreadable.'
    $script:LiveFailure = ''
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1*').Count -eq 1) 'Readable peers must restore the elected dispatcher.'

    # 8. Without distributed scheduling an Active server sends and an Inactive one does not.
    $script:Settings = New-Settings $states $true $false
    Reset-Node 'SRV-C2'; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 1) 'A standalone Active server must send the queue.'
    Reset-Node 'SRV-B1'; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'A standalone Inactive server must not send the queue.'

    # 9. The scheduler loop goes through the dispatcher only.
    $loopText = [IO.File]::ReadAllText($orchestratorPath)
    $direct = [regex]::Matches($loopText, '(?m)^(?!\s*#).*\bInvoke-SmartM365DeferredMail\b')
    Check ($direct.Count -eq 1 -and (Get-FunctionText 'Invoke-OrchestratorReportMailQueue').Contains('Invoke-SmartM365DeferredMail')) 'Only Invoke-OrchestratorReportMailQueue may call the queue worker.'
    Check ($loopText -match 'try \{\s*Invoke-OrchestratorReportMailQueue\s*\} catch \{') 'The scheduler loop must call the dispatcher inside its try block.'

    Write-Host ("PASS: report mail queue dispatcher, {0} offline checks." -f $script:passed)
}
finally {
    $env:COMPUTERNAME = $originalComputer
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAZiD4zN6vATDRA
# jk1q4kDGbclnzwKREG37caDrnvh1gqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIS62fftViEbjHFuuyHsQoE6Pe1K9Zkg5dGA2y3Vp9P0MA0GCSqG
# SIb3DQEBAQUABIIBgDLzJKuc8foiLcEsz1StzxuY2Z/1Ldot1cublUep2WQGN7Ty
# W2E3d4i+IOyfDqd6WbmcdgXfm74vnOCc24zeUxuHIeOd1qCLrL+1zZ+csiSRwG8X
# ud9vYtGRBRO1cCr9mj4htRbrUgON3OvBWOfSnp9+Fb4oThCyPfI15XIs210s4mlz
# 9rJQUUdrFO2w6tDLgOeLlFfWutR1tkPY9fWQ8KnYBQzKB1Y0Q7xXRXAVB+TZtfO1
# kqw73GaSwJC4D0xl5DWkMJXL7vAJoEoUpV5kY2P4Dv50KRsxHENzZ/ZD53ahwbNF
# zdzVI7IKN85I6tn3dFwLbUhXKjIm8vHSOYADk6u5JxU2OPIi4Vq52QVjSsAHrQd6
# iTEQPTbf7Y9IouSgg1xlGPyrAX488jlELQKmgfpyDFraQ8Tq384ODk8FEnbN7Ud6
# uw7jOkQTHjYDShcu06sLYAudEqSSRMv3lxYt07Dg9Anglf8VfnvXJEPIzr87t5rh
# 9vhW/LQOvsb3PwYszKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTUyMDUy
# MzRaMC8GCSqGSIb3DQEJBDEiBCBxTuzbblM/cdsszpynElHNwrEx8cnEdgG/uZLC
# XfaieTANBgkqhkiG9w0BAQEFAASCAgAQDZGWnseS/vFJ9BDeTH9IGAxZ4AlumL9O
# 41QlToT8rM4sfL5vqghYFh1Aj4NLtmw9oVeG58JBDQsRHkJcRESs8mab90fzQWt4
# Lo9YbXsV6N8CZw6cmgBR/5fKwNneMXSd1IzHuaYV4cq74ycxyv0z7lLSmf005o2n
# 2c4Ok0bZOdgWps8I6K1FB7880Bt3+/b425Uc252j/bqo5MoQacblx+zj4+PYIQSU
# 85foNrMGM9b6quxZBq7H6N3DvIqF56vU/CYkZHHo49RLjon8kvsoXz3RmZ0InGJO
# YWb9ndxamHbhmj/DGoYvcON8ZRIJKTIB2aZaE7BoKfFiVYoVtJ3KCQE+ZChNSl7W
# PFFPUQakcUx0e9xiiUTikHikYXGDmeQsvFV8gRXGHrRcVst2YeIt3kgRsG8JK1xI
# lZZQxUBs34r6MTDarVoQicbAflNzUd4Q5QhvVedCuYvFvcU2uFrnBMzj5WPjPsye
# hr9X5jh3uh8RBoYCq5ZP77WwF3VBmr4T7xs3n4S2jSBTXm7NJRhUZjXJhPdvRsD4
# zeZbC1FYIroMo+yqn78p5Z6v54RvGKJT3XAqwoY44a9l0c+pjjAX2epdB4n/h4tI
# JD2jfWvuz0OTbJjkH8t7bp+R+PHT/CF6yoDEicyp6w3sVtlmw+EJC/dfmx/ODiUo
# LZs1Yc/7ZQ==
# SIG # End signature block
