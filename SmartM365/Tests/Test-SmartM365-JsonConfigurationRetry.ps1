#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic configuration lock, retry and diagnostic regression tests.
.VERSION
1.0.0
.NOTES
Only newly created temporary files are used. No configuration initialization,
collector, tenant, real share, authentication or cloud operation is executed.
#>
[CmdletBinding()]
param([string]$FixtureRoot = (Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonRetry-' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../Config/SmartM365-TenantContext.ps1')
$module = Get-Module SmartM365.JsonTransport
$originalPolicy = & $module { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root = [IO.Path]::GetFullPath($FixtureRoot)
if (Test-Path -LiteralPath $root) { throw 'Use a new fixture directory; existing data is never accepted.' }
New-Item -ItemType Directory -Path $root | Out-Null
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Fixture([string]$Name) {
    $path = Join-Path $root ($Name + '.local.json.txt')
    [IO.File]::WriteAllText($path, '{"Setting":"private","Nested":{"Keep":"custom"},"Extra":42}')
    [IO.File]::WriteAllText(($path + '.template'), '{"Setting":"default","Nested":{"Keep":"default","Add":7},"Added":true}')
    return $path
}
function Failure([scriptblock]$Action) {
    try { & $Action | Out-Null } catch { return $_.Exception.Message }
    throw 'Expected synthetic failure did not occur.'
}
if (-not ('SmartM365.Tests.JsonRetryLease' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Threading;
namespace SmartM365.Tests {
    public sealed class JsonRetryLease : IDisposable {
        private FileStream stream;
        private Thread worker;
        public Exception Error;
        public static JsonRetryLease Hold(string path, int milliseconds, string update, bool exclusive) {
            var lease = new JsonRetryLease();
            lease.stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                exclusive ? FileShare.None : FileShare.ReadWrite);
            if (milliseconds > 0) {
                lease.worker = new Thread(delegate() {
                    try {
                        Thread.Sleep(milliseconds);
                        if (!String.IsNullOrEmpty(update)) File.WriteAllText(path, update);
                    } catch (Exception e) { lease.Error = e; }
                    finally { lease.stream.Dispose(); }
                });
                lease.worker.IsBackground = true;
                lease.worker.Start();
            }
            return lease;
        }
        public void Dispose() {
            if (worker != null) worker.Join();
            stream.Dispose();
        }
    }
}
'@
}
try {
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{ Mode='Readers'; QualifiedUncRoots=@(); QualifiedLocalDataRoots=@() } } }
    $path = Fixture 'nominal'
    $config = Read-SmartM365JsonConfig $path -Required
    Check ($config.Setting -eq 'private' -and $config.Nested.Keep -eq 'custom' -and $config.Extra -eq 42) 'Private values changed.'
    Check ($config.Added -and $config.Nested.Add -eq 7) 'Missing keys were not added.'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $null = Read-SmartM365JsonConfig $path -Required
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Already enriched configuration was rewritten.'
    Check ((Get-Command Read-SmartM365JsonConfig).Parameters.ContainsKey('ConfigurationWriteTimeoutSeconds')) 'Configuration wait budget is not configurable.'

    $path = Fixture 'released'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 700, $null, $false)
    try {
        $records = @(Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 6>&1)
        $config = $records | Where-Object { $_ -is [Collections.IDictionary] }
        Check ($config.Added -and $config.Setting -eq 'private') 'Released lock did not permit enrichment.'
        $text = ($records | Where-Object { $_ -is [Management.Automation.InformationRecord] } | Out-String)
        Check ($text -match 'JSON destination locked' -and $text -match 'retry in' -and $text -match 'timeout=3s') 'Retry progress was not logged.'
        Check ($text -match 'JSON replacement succeeded after') 'Recovery was not logged.'
    } finally { $lease.Dispose() }
    Check ($null -eq $lease.Error) 'Synthetic release worker failed.'

    # Prove that configuration enrichment survives a lock beyond the former 10s budget.
    $path = Fixture 'default-budget'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 11000, $null, $false)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try { $config = Read-SmartM365JsonConfig $path -Required 6>$null }
    finally { $lease.Dispose() }
    Check ($config.Added -and $null -eq $lease.Error) 'Default configuration budget did not survive an 11s lock.'
    Check ($watch.Elapsed.TotalSeconds -ge 10 -and $watch.Elapsed.TotalSeconds -lt 25) 'Default retry recovery exceeded its expected test window.'

    $path = Fixture 'timeout'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 0, $null, $false)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 1 6>$null } }
    finally { $lease.Dispose() }
    $watch.Stop()
    Check ($message -match 'Atomic replacement timed out' -and $message -like "*$path*" -and $message -match 'during template update') 'Timeout diagnosis lost its phase or path.'
    Check ($message -notmatch 'not valid JSON|Invalid JSON syntax') 'A lock was reported as invalid JSON.'
    Check ($watch.Elapsed.TotalSeconds -ge 0.9 -and $watch.Elapsed.TotalSeconds -lt 5) 'Replacement timeout was not bounded.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Timed-out write changed the configuration.'

    $path = Fixture 'zero-budget'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 0, $null, $false)
    $watch.Restart()
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 0 6>$null } }
    finally { $lease.Dispose() }
    Check ($message -match 'timed out') 'Zero retry budget did not fail explicitly.'
    Check ($watch.Elapsed.TotalSeconds -lt 1) 'Zero retry budget waited.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'Zero-budget failure changed configuration.'

    $path = Fixture 'concurrent'
    $concurrent = '{"Setting":"changed-by-other-writer","Extra":99}'
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 600, $concurrent, $false)
    try { $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 6>$null } }
    finally { $lease.Dispose() }
    Check ($null -eq $lease.Error) 'Synthetic concurrent writer failed.'
    Check ($message -match 'refusing a lost update') 'A change during the wait was not detected.'
    Check ([IO.File]::ReadAllText($path) -ceq $concurrent) 'A concurrent update was overwritten.'

    $path = Fixture 'exclusive'
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $lease = [SmartM365.Tests.JsonRetryLease]::Hold($path, 500, $null, $true)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes('{"Setting":"private","Added":true}')
        $null = Write-SmartM365JsonBytesAtomically -Path $path -Bytes $bytes -ExpectedSHA256 $hash -ReplacementTimeoutSeconds 3 -Validate { param($d) if ($d.Setting -ne 'private') { throw 'Wrong fixture.' } } 6>$null
    } finally { $lease.Dispose() }
    Check ((Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).Added) 'A transient exclusive lock during hash verification was not retried.'

    $path = Fixture 'invalid'
    [IO.File]::WriteAllText($path, '{')
    $watch.Restart()
    $message = Failure { Read-SmartM365JsonConfig $path -Required -ConfigurationWriteTimeoutSeconds 3 }
    Check ($message -match 'Invalid JSON syntax' -and $message -match 'during read') 'Invalid JSON lost its syntax diagnosis.'
    Check ($watch.Elapsed.TotalSeconds -lt 1 -and [IO.File]::ReadAllText($path) -eq '{') 'Invalid JSON was retried or overwritten.'

    $path = Fixture 'invalid-template'
    [IO.File]::WriteAllText(($path + '.template'), '{')
    $hash = (Get-FileHash -LiteralPath $path).Hash
    $message = Failure { Read-SmartM365JsonConfig $path -Required }
    Check ($message -match 'template update' -and $message -notmatch 'not valid JSON|Invalid JSON syntax') 'A bad template was attributed to runtime JSON syntax.'
    Check ((Get-FileHash -LiteralPath $path).Hash -eq $hash) 'A bad template changed runtime configuration.'

    $path = Join-Path $root 'directory.json.txt'
    New-Item -ItemType Directory -Path $path | Out-Null
    $watch.Restart()
    $message = Failure { Write-SmartM365JsonBytesAtomically -Path $path -Bytes ([Text.Encoding]::UTF8.GetBytes('{}')) -ReplacementTimeoutSeconds 3 -Validate { param($d) if ($null -eq $d) { throw 'Missing fixture.' } } 6>$null }
    Check ($message -match 'Atomic replacement failed' -and $message -notmatch 'timed out') 'A permanent destination error was treated as a transient lock.'
    Check ($watch.Elapsed.TotalSeconds -lt 1 -and (Test-Path -LiteralPath $path -PathType Container)) 'Permanent destination failure was retried or changed its target.'

    foreach ($code in @(32, 33, 5, 19)) {
        $native = [IO.IOException]::new('Synthetic native error.', (-2147024896 + $code))
        $wrapped = [InvalidOperationException]::new('Synthetic wrapper.', $native)
        $actual = & $module { param($e) Get-SmartM365JsonLockErrorCode $e } $wrapped
        $expectedCode = if ($code -in @(32, 33)) { $code } else { 0 }
        Check ($actual -eq $expectedCode) 'Native error classification retried a permanent error or missed a lock.'
    }
    [pscustomobject]@{ Passed=$script:passed; FixtureRoot=$root; Evidence='Synthetic local files only; no preparation, private configuration, real share or cloud access.' }
} finally {
    & $module { param($body) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $body } $originalPolicy
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBLh4yRz3jf4MQt
# tFPNkMxqVrRZb/wliGckhpneCrvMEKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKRnD3vL75F/l9397h5kuE1kTF3e+fJqrbyImKX2HBtmMA0GCSqG
# SIb3DQEBAQUABIIBgEx+mXsCxKAqE/L0qVfQVY6g0DFGMVZzxfSYH2X+Y4gs+mqN
# q1zVqJXqyU2E3vSqquCUJm6Gc3sPJKcfYACufa7hqsAFjCl9QzUxEj6gih2TdGsQ
# P6cbSguvjjTLZT4LHszuniVAWd7G1oxUuDcJHi2iyRGG5iEotq+p3iHCmUG+v1+T
# m511mRITljXw/EFasiC8ZwD/X6S9k93FZiH5CnARbHauypvUAr/m55BfssffLRXc
# 6x5AAJeEsw4Eqh2loYSd/y5EUeOxOo4QusqWDPii+cNiv1oJWbmo7ok0GyW+JGBM
# 5C7lynpL0fBCUygyQI7s5+wPw0GRt5oX54Ko5aN5pX8Y/e64TADlklbLrJXHsFMq
# ehywATlN0X5RM+RIDALs4R1gPWTXXKSAnIyId4hKpmZuLpEAKMCbCa1WXVxi3LUj
# 9pWUxQWaIQQVndoPfH+F32YU/PJnIvzWOORpt/GyMVrRt8PSuBUiak5+jscXIEou
# E67Lt9exMDG7F63nkqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkyMTM2
# NThaMC8GCSqGSIb3DQEJBDEiBCBMkHN5IbnwTQrGNEyGztQ+q+q8h2NmCMyRom68
# p30cWTANBgkqhkiG9w0BAQEFAASCAgAQ/+FC/CDPUI+6MDZCeruUYxtaSZRLYWMq
# Ol0aL0Wh/7GLxqwJjP4RtM4nBRIQABEyO5/NA7ooGvPXoxrAm351asbG9jPUdENn
# HqQq1Gh7yzh0t9gfcuckDT6GDO1fwAAMjEzQYpl5cgAuN1xnbBIW1XsnecYmnPHO
# djsoth/eE7lw6fJFkzDZMP6UkHuieuNF5mVjhCQmsKHegxNB243My5pJoaos1qM6
# CBn8XknnA/89Poc07G8YPO3uloMqeS/aQr/h0CWCaQf75iQvWqqvkogSixSgnnXZ
# +cKiSIJ6G+/kK9fLPmut41EfoVZYQ7qnAUmEabtGAdiPcms8KEUgbJ6VpJQTEhAx
# svqokKR55cZW1rZCRPY9LbO67WAyYswOBq55ZvkINgcvqGvPhuc6bVLuKldndkLH
# yQTEPlOmvr6vdacDgT/CpYBiyRG7/rj5xOtQxMXC/AzAgVumFBzgJYQQqdB4cCO8
# 3S3OnjVWlWv0cyQUaB5V891annZhZ+5hzimz8ezR3enId/3QdPQFTPvehV9WHt96
# u8ibI9DFk2XZvm3885ONYMF6zEAMjTYP4SF1YQoXtcEi+/BN8St1jB9k+JgPq6eG
# RRZPKD64BLVniKIfv8hhHQIfvrL00q00PdJAfuTxVqU+65PzkX3lOwkn6+Jl3pqw
# dK3Tkf4H+g==
# SIG # End signature block
