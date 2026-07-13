<#
    Name: SmartM365-Upgrade-Diagnostics-Detection.ps1
    Version: 1.0
    Description: Consolidated SetupDiag and Windows upgrade blocker detection.
#>

[CmdletBinding()]
param(
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MaxAnalysisAgeDays = 7
)

$ErrorActionPreference = "Stop"
$SetupDiagResultPath = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml"
$SetupDiagToolPaths = @(
    (Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Tools\SetupDiag\SetupDiag.exe"),
    (Join-Path $env:ProgramData "SetupDiag\SetupDiag.exe")
)
$EnterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"

$BlockingCodes = @(
    "0xC1900208",
    "0xC1900101",
    "0x80070070",
    "0x80240020",
    "0x80242016",
    "0x800F0922",
    "0xC1900200",
    "0x8007042B",
    "0x8007001F"
)

$RollbackPaths = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther\NewOS\Rollback"),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

$SetupLogRoots = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther"),
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Panther'),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Source,
        [AllowNull()]$Events,
        [Parameter(Mandatory = $true)][string[]]$Codes,
        [Parameter(Mandatory = $true)][int]$MaxIssueLength
    )

    foreach ($eventRecord in @($Events)) {
        if ($null -eq $eventRecord -or [string]::IsNullOrWhiteSpace($eventRecord.Message)) {
            continue
        }

        foreach ($code in $Codes) {
            if ($eventRecord.Message -match [regex]::Escape($code)) {
                $message = $eventRecord.Message -replace "`r|`n", " "
                if ($message.Length -gt $MaxIssueLength) {
                    $message = $message.Substring(0, $MaxIssueLength) + "..."
                }
                Add-Issue -Issues $Issues -Message "$Source Id=$($eventRecord.Id) Code=$code Message=$message"
            }
        }
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $hasSetupDiagResult = Test-Path -LiteralPath $SetupDiagResultPath -PathType Leaf
    $hasSetupDiagTool = $false
    foreach ($setupDiagToolPath in $SetupDiagToolPaths) {
        if (Test-Path -LiteralPath $setupDiagToolPath -PathType Leaf) {
            $hasSetupDiagTool = $true
            break
        }
    }
    $hasSetupLogs = $false
    $hasRollbackLogs = $false

    foreach ($setupLogRoot in $SetupLogRoots) {
        if (Test-Path -LiteralPath $setupLogRoot) {
            $hasSetupLogs = $true
            break
        }
    }

    foreach ($rollbackPath in $RollbackPaths) {
        if (Test-Path -LiteralPath $rollbackPath) {
            $hasRollbackLogs = $true
            Add-Issue -Issues $issues -Message "RollbackDetected=$rollbackPath"
        }
    }

    if ($hasSetupLogs -and -not $hasSetupDiagResult) {
        if ($hasSetupDiagTool) {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedWithoutSetupDiagAnalysis"
        }
        else {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedSetupDiagToolMissing"
        }
    }

    if ($hasRollbackLogs -and -not $hasSetupDiagResult) {
        Add-Issue -Issues $issues -Message "RollbackLogsDetectedWithoutSetupDiagAnalysis"
    }

    if ($hasSetupDiagResult) {
        $xmlAgeDays = (New-TimeSpan -Start (Get-Item -LiteralPath $SetupDiagResultPath).LastWriteTimeUtc -End (Get-Date).ToUniversalTime()).TotalDays
        if ($xmlAgeDays -gt $MaxAnalysisAgeDays) {
            Add-Issue -Issues $issues -Message ("SetupDiagResultOutdated AgeDays={0:N1}" -f $xmlAgeDays)
        }

        $xmlContent = Get-Content -LiteralPath $SetupDiagResultPath -Raw -ErrorAction SilentlyContinue
        foreach ($code in $BlockingCodes) {
            if ($xmlContent -match [regex]::Escape($code)) {
                Add-Issue -Issues $issues -Message "SetupDiagDetected=$code"
            }
        }

        if ($xmlContent -match "Matching Profile found:\s*(.+)") {
            Add-Issue -Issues $issues -Message "SetupDiagProfile=$($matches[1].Trim())"
        }
    }
    elseif (Test-Path -Path $EnterpriseRegistryPath) {
        $setupDiagState = Get-ItemProperty -Path $EnterpriseRegistryPath -Name LastRunUtc -ErrorAction SilentlyContinue
        if ($setupDiagState -and -not [string]::IsNullOrWhiteSpace($setupDiagState.LastRunUtc)) {
            $analysisAgeDays = (New-TimeSpan -Start ([datetime]::Parse($setupDiagState.LastRunUtc)) -End (Get-Date).ToUniversalTime()).TotalDays
            if ($analysisAgeDays -gt $MaxAnalysisAgeDays) {
                Add-Issue -Issues $issues -Message ("SetupDiagRegistryStateOutdated AgeDays={0:N1}" -f $analysisAgeDays)
            }
        }
    }

    $wuEvents = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 500 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "WUEvent" -Events $wuEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    $setupEvents = Get-WinEvent -LogName "Setup" -MaxEvents 300 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "SetupEvent" -Events $setupEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    if ($issues.Count -gt 0) {
        Write-Output "Status=UpgradeDiagnosticsRequired"
        Write-Output "IssueCount=$($issues.Count)"
        $issues | Select-Object -First $MaxIssuesToDisplay | ForEach-Object { Write-Output $_ }
        exit 1
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCcPlC2r/4jfpas
# ut/F4wZqk6alsUmfhMOxk6KUJnNU16CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIO+Qn1M7U+8j16++lTdryrqPaoF2AHy2KP8jKBApZ1iYMA0GCSqG
# SIb3DQEBAQUABIIBgE1V9yOMEqe6AUeR8/mxfPKAarchKITBMFtPWD0bEIT5hlwA
# odpry9zKvfMHDxbFVgtOSdnMtvYQB/WoLYaKE93yHf59EG7l9X/0jJtmIEUZfbDo
# cX60ibuo2wE5TlbAQqZPER8AU8lIE+3wMdF2kZ0wO9doR1eMkN3q3BqcGfsagbhM
# GEveY/Lp5zpVAQKqL4v8CujCn4sWPD1+1KQ5bW8XQqOKVPiuh3H8nWfbKp3Rwz1w
# QMM0q++xCjBmyLHfS5hGeZQdXBm34nr2gv/xxSOaeoZvlymOE5QdbaSnL6ffxDhP
# ScBCbnbAKMio0X+wZtN+SQdY0HMoLKwbVbiw32oC2arF1jQ9V3XG3Q67p7S8Y+Cg
# 8SN49mAR0zsX6EnPBDu36KgtqbKLQ/medXyf5/KsXsZJ0lJCmLf/Bb1iyBzXpVZR
# y3u2YT9QxeGqNWsG59QvigMcF+T0GiYOhfuzEAigOaXqAaLO6Am/BaFvTLpOGt2d
# dkAtpD6or5/nXcos5aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NTFaMC8GCSqGSIb3DQEJBDEiBCDRVReDfSgZlUknc+P7HXAOKq/oWK//+8F+NRkL
# gtbo+zANBgkqhkiG9w0BAQEFAASCAgChH+aHMOx3LCQkYQBBlE+RLcTslBP4azjY
# lvZiHOnQApveDQ5Qrr1iJcS0aPw+dMb+p3HGb+aAWTa5Xk4nnF4O0tkmTMi/lRqa
# EHhdLb/JU/zq215PEgTrE8pKydd0i84VFhFdmXBMpcfY3Fr37HTvFvC5Hgeluflf
# TNE6cnAzs62frHJF2IBGQ5IZYFDU2P4D5MYeCtDXAVyYBXtVxLURXhYRzmlWU1h3
# bth+FYHOjR4o5swCgPHhBJWavHkHDnWUe+gZ1tghsfsrRcrmnxF+BfsgNeC7abq9
# Uj2zmxwf3aW0O3S4KOMhx6VoOXdu5nNHtpD8JCn0N342Wl/p1QeHVeINWN/iENIP
# SBxW6Ni4sOs3h5ON8u0YRSY3t5/fLi+sWwNOAc/fm1xRZxrSAWYLGQJZXzHysFoA
# FJGNlbcvOvhg36A2pEXEpEe5ghsyGi+YzEUTqkLFw03zPPRSf0w+7mq7QGIWH8wS
# g5bAh/5retZo09bfMcClc9YVRukahpmnHn5uCELF2vLqKqO7OFDw91qPys9BbEPv
# 4Ej1KDZpOW7u78uQnL4vOHapIcDNnoOGNt0ZMwzvsoil+M5OS9wyJetLETSSuXNA
# HKxmU/BGCw6MxNJq4YoR53EjgY4NZuSij1pOl0zsc3tUrPf4+zoNMdsfB8wxJEBT
# J7t4IAu/+Q==
# SIG # End signature block
