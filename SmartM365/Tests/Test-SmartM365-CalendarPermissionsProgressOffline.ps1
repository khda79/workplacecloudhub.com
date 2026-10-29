<#
.SYNOPSIS
Synthetic, offline tests for the calendar-permission collector progress lines (Exchange Online and Exchange on-premises).
.VERSION
1.0.0
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

$exoPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1'
$onPremPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
$functionNames = @('New-CalendarProgressTracker', 'Get-CalendarProgressLine')
$exoText = @(Get-FunctionText -Path $exoPath -Names $functionNames) -join "`n"
$onPremText = @(Get-FunctionText -Path $onPremPath -Names $functionNames) -join "`n"
$module = New-Module -ScriptBlock ([scriptblock]::Create($exoText))

try {
    Test-OfflineCase 'Both collectors carry the same progress helpers' {
        Assert-Offline ($exoText -ceq $onPremText) 'EXO and on-premises progress helpers diverged.'
    }

    Test-OfflineCase 'A line is logged at each 1,000-mailbox step and on the last mailbox only' {
        $lines = & $module {
            $tracker = New-CalendarProgressTracker -Total 2500
            for ($i = 1; $i -le 2500; $i++) {
                $tracker.Processed = $i
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds ($i * 0.05) -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        Assert-Offline ($lines.Count -eq 3) ("Expected 3 count-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like 'Calendar permissions progress: 1000/2500 mailboxes (40.0%);*') "Unexpected first line: $($lines[0])"
        Assert-Offline ($lines[1] -like '*: 2000/2500 mailboxes (80.0%);*') "Unexpected second line: $($lines[1])"
        Assert-Offline ($lines[2] -like '*: 2500/2500 mailboxes (100.0%);*ETA done.') "Unexpected last line: $($lines[2])"
    }

    Test-OfflineCase 'A slow collection logs at least every 5 minutes while mailboxes complete' {
        $lines = & $module {
            $tracker = New-CalendarProgressTracker -Total 50
            for ($i = 1; $i -le 20; $i++) {
                $tracker.Processed = $i
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds ($i * 70.0) -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        # 70 s per mailbox: lines after 5, 10, 15 and 20 mailboxes (350, 700, 1050, 1400 s).
        Assert-Offline ($lines.Count -eq 4) ("Expected 4 time-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like '*: 5/50 mailboxes (10.0%);*') "Unexpected first time-based line: $($lines[0])"
    }

    Test-OfflineCase 'No progress means no line, so the stalled-job alert still fires' {
        $observed = & $module {
            $tracker = New-CalendarProgressTracker -Total 100
            $tracker.Processed = 10
            $first = Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 400
            $stalled = Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 7200
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
                $tracker = New-CalendarProgressTracker -Total 4000 -WorkerTotals @(2000, 2000) -TrackWarnings
                $tracker.Processed = 1000
                $tracker.Errors = 3
                $tracker.Warnings = 2
                $tracker.Rows = 12345
                $tracker.WorkerProcessed[0] = 600
                $tracker.WorkerProcessed[1] = 400
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 500 -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
        }
        $expected = 'Calendar permissions progress: 1000/4000 mailboxes (25.0%); errors=3; warnings=2; rows=12345; rate=2.00 mailbox/s; elapsed=8.3 min; ETA ~25 min (20:25); workers 600/2000 400/2000.'
        Assert-Offline ($line -ceq $expected) "Unexpected line: $line"
        Assert-Offline (-not ($line.ToCharArray() | Where-Object { [int]$_ -gt 127 })) 'Progress line contains non-ASCII characters.'
    }

    Test-OfflineCase 'Warnings and workers are omitted when not tracked; single mailbox completes' {
        $line = & $module {
            $tracker = New-CalendarProgressTracker -Total 1
            $tracker.Processed = 1
            Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 0
        }
        Assert-Offline ($line -ceq 'Calendar permissions progress: 1/1 mailboxes (100.0%); errors=0; rows=0; rate=0.00 mailbox/s; elapsed=0.0 min; ETA done.') "Unexpected line: $line"
    }

    Test-OfflineCase 'EXO parent consumes worker results as a stream and logs progress' {
        $source = Get-Content -LiteralPath $exoPath -Raw
        Assert-Offline ($source -notmatch '\$parallelCalendarOutput\s*=\s*@\(\s*\$calendarWorkItems') 'Parallel output is still buffered before any progress line.'
        Assert-Offline ($source -match '(?s)ForEach-Object -ThrottleLimit \$workerCount -Parallel \{.*?\} \| ForEach-Object \{.*?Get-CalendarProgressLine -Tracker \$progressTracker') 'Parallel results are not piped into the progress consumer.'
        $onPremSource = Get-Content -LiteralPath $onPremPath -Raw
        Assert-Offline ($onPremSource -match '(?s)foreach \(\$mbx in \$mailboxes\) \{.*Get-CalendarProgressLine -Tracker \$progressTracker.*?\}\s*Write-Progress -Id 0 -Activity \$overallActivity -Completed') 'On-premises loop does not log progress per mailbox iteration.'
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
    throw ("{0} calendar-permission progress offline test(s) failed." -f $failed.Count)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBbU7xNCaD46So4
# /tidtLp5stjnsIpgIphunZf1fs9dlaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPZ2Gbl3//pPVf+Hyt2iyDg5CyWG99hRkh9Tph/JdPXxMA0GCSqG
# SIb3DQEBAQUABIIBgJ3RRePpAlVsO3jM/ra9S1IBIER2s7DvNHgXhDb+nME7uYYW
# ZO2kYRF1jAunVqV15+ggaBoqkMCfjQQMVvENEyRYtlsmrAGSpNXuGmUWvEuSZQ/a
# leRvRSanfSbjWMYYm/qJkQiEXOEmgrNFtud60FDUNzxeuvOrA8WBRu3onq5zZfH4
# wgaRdBlcBId2jxzjhrW1HHPZ8VOEfeBLIuKRnnVK2qT16u4qpKTrDYDoIYMkXad5
# EJ4skAOZ1MbNE7hWqYR5QC0IxFxY4gUs8THsjGfJUk9gSt/D47qrGPqTD6w4lB6G
# LzYFmDLbcVBH/IqRyMK5BLTon8NO4VTdofT8ATTbLdOCTWlyXY8KtpJudZfjXdTe
# xdhlndq2VGid+cU5+m84wuv63mnRNq1f5D632Ig5o41etCDn4Sb9BoK/pqR6xK7n
# wYwwdpCpteZiL0eo7dbl43u7C3Pd6RhOb+6qrYc9TIg2Mx2SrXZb/UCl7eP1Fq1T
# 4irtZbvACcvKiYzmCaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMzU2
# MDhaMC8GCSqGSIb3DQEJBDEiBCCjkA/9SJ88LOtUt9PEnnj1S6vQbxp4dqYdtcDH
# pknSHTANBgkqhkiG9w0BAQEFAASCAgBa0PdC2HZvy86zDy0vfXUL2Cx5tyGsa9lO
# o+lVjHdjgeizD/J/d/2ThJc2On8aTyUvYdos3e48XVOakq+R9FGva+3uAtqvYQ1+
# IbWvCy2FeWsvzijfF0yR5V1mit0UcI3XbFZ/KlDLfzGx0ZrgF9YDEKYzqzcaRmzP
# okYdrFN7jDoWa8UXhp1Y0gqYDvtclUnS2qkybI0Zo2Gx3YK5UCkcPM0XrijUE+pU
# +skyGf90hnsN1SvoBOs/XbUqf2+6J+iLChwSEGIfnnZzrutW+dqp+2vCiWRW0a4l
# BjjXiHyj2+fWdhtmmXG+y1MQoU4hzpHKDoNtMNMDMpqMi2aUlIxlFQWDciq8BtkP
# xuD9cGnCBub91W0WdCU2PKToMS4ZfkOLCkzpfdev6c+E1iZAB0QNmf173r9ybLuU
# LhTAwuMwPrSeQ3/lZyLndBoO29BQvN58m2OK1znsat8JpNj1/b+GPcRf4ZumW1n+
# wUdHOjL7kU7YUYjRWhtHpj9Ljw8/NhWVphLwKf+e9bvMQlaOoLIV7FqeZZJfPwA9
# 6BlYDD+IfM520LstjtWm8tt3/nptzZMMSaWdfhxjtLkvNo9kBMWbLBVxtaoCvVCi
# E93sHInlg6PNhEjn6Ef7wFxXSxKT0lhdeiRPmQlEAMgyycJhwtuqV00dOVH2B4D5
# 2K2rcOpuTw==
# SIG # End signature block
