<#
.SYNOPSIS
Synthetic, offline tests for the progress lines of the EXO mailboxes parallel permission collection (-PermissionsOnly).
.VERSION
1.0.0
.NOTES
Synthetic counters and a local ForEach-Object -Parallel sample only. No tenant, Exchange, API, mail or synchronized data access.
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' })
    }
    catch {
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=$_.Exception.Message })
    }
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw "Source parse failed: $Path"}
    foreach($name in $Names){
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if($null -eq $node){throw "Function not found: $name"}
        $node.Extent.Text
    }
}

$producerPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1'
$functionText = @(Get-FunctionText -Path $producerPath -Names @('New-MailboxPermissionProgressTracker', 'Get-MailboxPermissionProgressLine')) -join "`n"
$module = New-Module -ScriptBlock ([scriptblock]::Create($functionText))

try {
    Test-OfflineCase 'A line is logged at each 1,000-mailbox step and on the last mailbox only' {
        $lines = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 2500
            for ($i = 1; $i -le 2500; $i++) {
                $tracker.Processed = $i
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds ($i * 0.05) -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        Assert-Offline ($lines.Count -eq 3) ("Expected 3 count-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like 'Mailbox permissions progress: 1000/2500 mailboxes (40.0%);*') "Unexpected first line: $($lines[0])"
        Assert-Offline ($lines[1] -like '*: 2000/2500 mailboxes (80.0%);*') "Unexpected second line: $($lines[1])"
        Assert-Offline ($lines[2] -like '*: 2500/2500 mailboxes (100.0%);*ETA done.') "Unexpected last line: $($lines[2])"
    }

    Test-OfflineCase 'A slow collection logs at least every 5 minutes while mailboxes complete' {
        $lines = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 50
            for ($i = 1; $i -le 20; $i++) {
                $tracker.Processed = $i
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds ($i * 70.0) -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        # 70 s per mailbox: lines after 5, 10, 15 and 20 mailboxes (350, 700, 1050, 1400 s).
        Assert-Offline ($lines.Count -eq 4) ("Expected 4 time-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like '*: 5/50 mailboxes (10.0%);*') "Unexpected first time-based line: $($lines[0])"
    }

    Test-OfflineCase 'No progress means no line, so the stalled-job alert still fires' {
        $observed = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 100
            $tracker.Processed = 10
            $first = Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 400
            $stalled = Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 7200
            [pscustomobject]@{ First=$first; Stalled=$stalled }
        }
        Assert-Offline (-not [string]::IsNullOrEmpty($observed.First)) 'The first time-based line was not produced.'
        Assert-Offline ($null -eq $observed.Stalled) 'A line was produced without new completed mailboxes.'
    }

    Test-OfflineCase 'Rate, ETA, counters and worker detail are exact and culture-invariant ASCII' {
        $previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            $line = & $module {
                $tracker = New-MailboxPermissionProgressTracker -Total 4000 -WorkerTotals @(2000, 2000)
                $tracker.Processed = 1000
                $tracker.Errors = 3
                $tracker.ConnectionFailures = 1
                $tracker.WorkerProcessed[0] = 600
                $tracker.WorkerProcessed[1] = 400
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 500 -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
        }
        $expected = 'Mailbox permissions progress: 1000/4000 mailboxes (25.0%); errors=3; connectionFailures=1; rate=2.00 mailbox/s; elapsed=8.3 min; ETA ~25 min (06:25); workers 600/2000 400/2000.'
        Assert-Offline ($line -ceq $expected) "Unexpected line: $line"
        Assert-Offline (-not ($line.ToCharArray() | Where-Object { [int]$_ -gt 127 })) 'Progress line contains non-ASCII characters.'
    }

    Test-OfflineCase 'Workers are omitted when not given; a single mailbox completes' {
        $line = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 1
            $tracker.Processed = 1
            Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 0
        }
        Assert-Offline ($line -ceq 'Mailbox permissions progress: 1/1 mailboxes (100.0%); errors=0; connectionFailures=0; rate=0.00 mailbox/s; elapsed=0.0 min; ETA done.') "Unexpected line: $line"
    }

    Test-OfflineCase 'An empty collection produces no line' {
        $line = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 0
            Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 900
        }
        Assert-Offline ($null -eq $line) "A line was produced for an empty collection: $line"
    }

    Test-OfflineCase 'The parent consumes worker results as a stream and keeps the publication guards' {
        $source = Get-Content -LiteralPath $producerPath -Raw
        Assert-Offline ($source -notmatch '\$parallelPermissionOutput\s*=\s*@\(\s*\$permissionWorkItems') 'Parallel output is still buffered before any progress line.'
        Assert-Offline ($source -match '(?s)\$permissionWorkItems \| ForEach-Object -ThrottleLimit \$workerCount -Parallel \{.*?\} \| ForEach-Object \{.*?\[void\]\$parallelPermissionOutput\.Add\(\$permissionOutput\).*?Get-MailboxPermissionProgressLine -Tracker \$permissionProgress') 'Parallel results are not piped into the progress consumer.'
        Assert-Offline ($source -match '(?s)Get-MailboxPermissionProgressLine -Tracker \$permissionProgress.*?refusing to publish blank permission data.*?row-count mismatch: dispatched=') 'Connection-failure or row-parity guards are missing after the progress consumer.'
    }

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Test-OfflineCase 'ForEach-Object -Parallel results reach the parent before all workers finish' {
            $clock = [System.Diagnostics.Stopwatch]::StartNew()
            $received = New-Object 'System.Collections.Generic.List[double]'
            @(1, 2) | ForEach-Object -ThrottleLimit 2 -Parallel {
                if ($_ -eq 1) { 'fast' } else { Start-Sleep -Seconds 4; 'slow' }
            } | ForEach-Object { [void]$received.Add($clock.Elapsed.TotalSeconds) }
            Assert-Offline ($received.Count -eq 2) 'Parallel results were lost.'
            Assert-Offline ($received[0] -lt 3.0) ("First result arrived only after {0:F1} s: output is not streamed." -f $received[0])
        }
    }
}
finally {
    Remove-Module $module -Force
}

$results.ToArray() | Format-Table -AutoSize
if ($ResultPath) { $results.ToArray() | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 }
$failed = @($results | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    throw ("{0} mailbox permission progress offline test(s) failed." -f $failed.Count)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBW9uJJi/NrvKC+
# Mv71tLAQ8LTCtlAtcxyODzIzvuQ/e6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDfUOL+08l3kFCQ6KOVjxuDvu5LoRYHT7h54eQcDhWvSMA0GCSqG
# SIb3DQEBAQUABIIBgEweL8RMsXb4xfkUcD97BEMjhBtpgkPZhkxLPweus+fclBJ2
# 5J60elVYDRSjl8bp3N8jr4QSP14Cx+W6/CqqTIuuQOcZp64I7QsDUBkny9cq9X0f
# VkYx7MTkItx2mf7sCIf06Jyawun4TzdLeRRrU8CBIplwXX01vDaRe9u0EtR3Zi3e
# 7I7OUnudN4nAeupRzOEQu/nZBa5y60cn+ZdHVX9x9qy91fj2zDj3xX9LP/GWc9Vr
# LVmzQK/TrtDTdwCee+iiWWpic2ON0q8nDXPQ3rfdEAqG+sAgZbCrDbNwN+MBWmQf
# AkEYdmhYyt8E4431TttjLFqsQmSfM2qX28oZvg9tMK9VqP/RC74GYEMPjHnQn1a+
# p1Hr4LSDe35viYiSR0NWWC1Zgg5aIvhpy0bjfCXpolzq1CrCTbKUf9e9CqAnejOK
# xmcpDPrFblrq2Vjg6ZtzJ6VqKqHZHwlurFffBmIno2cBB3ewU9Kxq75WnN0FRHpX
# Dr1YecAM76rEyRvYmqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwNzU3
# MzBaMC8GCSqGSIb3DQEJBDEiBCBzs/nGpG7O4QhtqnsaatP9QMPKMtaotCS2ruwB
# kv9lDjANBgkqhkiG9w0BAQEFAASCAgAKegkzOEbXoVBnCCuLLOFQkOjBXAra8Jeh
# nlBzqKAVRAfHnr8quLtmz1wUwQHhNYR8yf6tuqwqG1Yk1gzysPAKczFUW3p95+/5
# sIq6l7aLI9XSU9GWbUXSRUk7o/MjMq3nz+kpT/J34YdbaI5M+vP37AUBGujFVdD+
# mQ0g5So532aa/JOSbp97YXXDou6swJUNoklAwhI/ug0VJyJghht7IW6w6BGkOapN
# lemKQxZ8sQStc6NCX8OY6h2AP7GX9zYdr9onTkh8bNgOo3JCW2jyTgk3fGig/nsI
# +WBieFRj5VV+/jcsAxoprtAwQZXVJp/8MlYKwghoqGEKcz6gtDPPcG7MBFQ/yGqQ
# kTCjjcZC4zFSEBQRwfIs1DiWFXH3BdU1GglxsxXcG1xKgiLTAvM00SnPi9gIuT55
# MUyaDWHv3rcggOIW1QdL0FeV7xz9JympxT62e6j8u4QQsE/Bm0f2B9EUkogZGdM6
# 90GPDWRm93YroVjuZesnGVKf5PRELR+B7TYHGHquGAlgYkR+vBHqap13xsdwF3i7
# jNEAG9xWIKNtJK3p1o60wnfJ2AhXsfzBI2nigPnfnktwSnklEeTzRjvhKr4pYC29
# OweSVoK+ZCEp6cQU4AtQHg+kqIheyijtp45AjTCMJQ4NsSWA+H9vbt2/p+NQ1kNo
# UyJFQ7vvWw==
# SIG # End signature block
