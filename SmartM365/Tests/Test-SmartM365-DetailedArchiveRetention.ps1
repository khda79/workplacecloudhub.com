<#
.SYNOPSIS
Offline tests for the seven-day detailed archive retention policy.
.VERSION
1.0.0
#>
#requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$referenceTime = [datetime]'2026-09-23T12:00:00'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Import-FunctionFromScript {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "PowerShell parse errors in $Path"
    $definition = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    }, $true))[0]
    Assert-True ($null -ne $definition) "Function $Name not found in $Path"
    $bodyText = $definition.Body.Extent.Text
    return [scriptblock]::Create($bodyText.Substring(1, $bodyText.Length - 2))
}

function Log { param([string]$Message) }
function Warn { param([string]$Message) }
function Write-Log { param([string]$Message, [string]$Level, [string]$Stage) }

$cases = @(
    [pscustomobject]@{
        Name = 'Exchange HybridIdentity Issues'
        Path = Join-Path $smartM365Root 'SmartInventory\ExchangeInventory\Migration\SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1'
        Prefix = 'Exchange_HybridIdentity_Issues'
        Format = 'yyyyMMdd-HHmmss'
        Version = '1.17'
        ArchiveWriteMarker = 'CopyCsv $main $archive'
        RetentionCallMarker = "Remove-DetailedArchiveFilesOlderThan -Folder (Join-Path `$OutputFolder 'Archive')"
    },
    [pscustomobject]@{
        Name = 'Windows11 Readiness Issues'
        Path = Join-Path $smartM365Root 'SmartInventory\M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
        Prefix = 'Intune_Windows11_Readiness_Issues'
        Format = 'yyyyMMdd-HHmmss'
        Version = '1.23'
        ArchiveWriteMarker = 'CopyCsv $main $archive'
        RetentionCallMarker = "Remove-DetailedArchiveFilesOlderThan -Folder (Join-Path `$OutputFolder 'Archive')"
    },
    [pscustomobject]@{
        Name = 'Windows Update Status'
        Path = Join-Path $smartM365Root 'SmartInventory\M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1'
        Prefix = 'Intune_WindowsUpdate_Status'
        Format = 'yyyyMMdd_HHmmss'
        Version = '1.41'
        ArchiveWriteMarker = 'Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal'
        RetentionCallMarker = 'Remove-DetailedArchiveFilesOlderThan -Folder $ArchiveFolder'
    }
)

foreach ($case in $cases) {
    $source = Get-Content -LiteralPath $case.Path -Raw
    Assert-True ($source -match ('(?m)^\s*\.VERSION\s*\r?\n' + [regex]::Escape($case.Version) + '\s*$')) "$($case.Name) metadata version"
    Assert-True ($source -match ('\$ScriptVersion\s*=\s*["'']' + [regex]::Escape($case.Version) + '["'']')) "$($case.Name) runtime version"
    Assert-True ($source -match '\[int\]\$DetailedArchiveRetentionDays\s*=\s*7') "$($case.Name) default retention is seven days"
    $localConfigTemplatePath = $case.Path -replace '\.ps1$', '.local.json.template'
    $weeklyHistoryEvidence = $source
    if (Test-Path -LiteralPath $localConfigTemplatePath) { $weeklyHistoryEvidence += Get-Content -LiteralPath $localConfigTemplatePath -Raw }
    Assert-True ($weeklyHistoryEvidence.Contains('WeeklyHistory')) "$($case.Name) weekly history remains configured"

    $writeIndex = $source.IndexOf($case.ArchiveWriteMarker, [StringComparison]::Ordinal)
    $retentionIndex = $source.IndexOf($case.RetentionCallMarker, [StringComparison]::Ordinal)
    Assert-True ($writeIndex -ge 0) "$($case.Name) archive write marker"
    Assert-True ($retentionIndex -gt $writeIndex) "$($case.Name) retention runs only after archive publication"

    $retentionFunction = Import-FunctionFromScript -Path $case.Path -Name 'Remove-DetailedArchiveFilesOlderThan'
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ArchiveRetention-' + [guid]::NewGuid().ToString('N'))
    $archiveFolder = Join-Path $tempRoot 'Archive'
    $weeklyFolder = Join-Path $tempRoot 'WeeklyHistory\2026-W39'
    New-Item -ItemType Directory -Path $archiveFolder,$weeklyFolder -Force | Out-Null

    try {
        $oldStamp = $referenceTime.AddDays(-7).AddSeconds(-1).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $cutoffStamp = $referenceTime.AddDays(-7).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $freshStamp = $referenceTime.AddHours(-1).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $oldPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$oldStamp)
        $cutoffPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$cutoffStamp)
        $freshPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$freshStamp)
        $unrelatedPath = Join-Path $archiveFolder ("Unrelated_{0}.csv" -f $oldStamp)
        $malformedPath = Join-Path $archiveFolder ("{0}_20260901.csv" -f $case.Prefix)
        $weeklyPath = Join-Path $weeklyFolder ("{0}_{1}.csv" -f $case.Prefix,$oldStamp)
        foreach ($path in @($oldPath,$cutoffPath,$freshPath,$unrelatedPath,$malformedPath,$weeklyPath)) {
            Set-Content -LiteralPath $path -Value 'test' -Encoding utf8
        }

        & $retentionFunction -Folder $archiveFolder -BaseNameWithoutExt $case.Prefix -RetentionDays 7 -ReferenceTime $referenceTime

        Assert-True (-not (Test-Path -LiteralPath $oldPath)) "$($case.Name) removes a matching snapshot older than seven days"
        Assert-True (Test-Path -LiteralPath $cutoffPath) "$($case.Name) retains a snapshot exactly at the cutoff"
        Assert-True (Test-Path -LiteralPath $freshPath) "$($case.Name) retains a fresh snapshot"
        Assert-True (Test-Path -LiteralPath $unrelatedPath) "$($case.Name) retains another collector's snapshot"
        Assert-True (Test-Path -LiteralPath $malformedPath) "$($case.Name) retains a malformed timestamp instead of guessing"
        Assert-True (Test-Path -LiteralPath $weeklyPath) "$($case.Name) does not recurse into WeeklyHistory"
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }
}

$statusSource = Get-Content -LiteralPath $cases[2].Path -Raw
Assert-True (-not $statusSource.Contains('$global:RetentionMaxCSV')) 'Windows Update Status no longer applies count-based CSV archive retention'
Assert-True ($statusSource.Contains('Prune-Files -Folder $LogsPath')) 'Windows Update Status keeps the independent log-count retention'

Write-Host "Detailed archive retention tests passed for $($cases.Count) collectors." -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCChBl9mk3eQVk6P
# 4//eOp76jALqvKMijMD4GN1hq0qnS6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGc4Q7M0JcEPCZMM3P7gAKB9oWHtTc8P3HpE8Zcto/bPMA0GCSqG
# SIb3DQEBAQUABIIBgFom8z+eZWYBirkoPyBJyo+FYurQRFN/4nnb5MmyBz7GYqpA
# UaPYn+dOvOF4r3mjdbhbNfZmmUWgN/y1WSqXSiTt6ntNYJDPV1AxVZcEGbaL0wvz
# s4RKKzW/7IxBqfqvLJsaebN3JYCAz9UkXB4MgTjdGPeJDlElmTItEFlg8WJlCITn
# VS5rLlqRCM6FJiMi9Yb74Qe8vxXabthG/Gjeq5df0Ue8kaFgxmkCL6If9JTAaV/i
# I+Fv3MHGuqduWaR6PYQ/JaiUC1rzea0m38dw4PRs5aK1sh/zrMFHWKfAOz6JvWkG
# ltFci1hjKUQ7ZWEH8GKo9blS3CgWUEnJ1BPKeQMp6y6BYeoKEBqexv00a2r+P58N
# UVxlWIJo1Jt4+HiH9AY735Vq96jCNZNU6/Fpr9zHpDcTWYHG+2bUO7YhqbUJJ3zM
# q8y9zwYXNBtKvLsUQoLN0npbyxPj55+g8daq4FsGz5FFmmsovaA5blURL3y8s1Ki
# U9QdkTbBjVREDneZ5aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjMyMTM1
# NDRaMC8GCSqGSIb3DQEJBDEiBCDxFkePfB+bF5qN+q72esQ32UCFzTpO+5sD5hWo
# cjXQqzANBgkqhkiG9w0BAQEFAASCAgArcSHzfk300Fwh2GpiY1iyx+U0Vt4py0tQ
# MCPu0hNgj2yCj18oA0KCjxW0aEQBdAD9NtJvuuzJZpE6QU38IFrv6SoWqEbSiHnS
# /A3NoO+mXgyTDCdGHsUlFZCIQycRhlWqhem1iton3wk5rxfOiY5+BkP7+F4Fa10N
# N1NzQqUQpECv04UItITdtfSDhdCCdCkBtn4kW2LBerdzdRlNqpoJ1oHCRldkIvaa
# WQHWTDWf/OKYM3hxqAkZlmZiC6zEoQ5QuUJ8pefbHjpzXvu1Y1d/0FjXA7hUitgS
# WDe6kRXNjyh0/2Ntd+IbFBIdB/Z0vhLGicr9cPiO1FFi9bXUOPUfcWae/AQKWnFq
# VZkpEIyqgAOYJVmrOgORVHuSEVQKWQQE9gb1TlWXaumtnrsI23Pl6Wyi5UuSzQCo
# mE8BjE8u4F/1wURbq1NRRIEACERGN1uglKP4TovOTxdeE53x/cvqOVaI1hz12Lzf
# 8ty+k69+9LuzNWwwy0oUYgKaU9tTIqsLayQNHyE6TQQi5RqVZBtk/zpm0cScFXwy
# Fu8ZV4Mghll3CmXTuWJq98lH/oizrxG0RoQJmXzWigA7YoTG0+Wnjl9chRDLQl+a
# LiF3SJimooVsXVQepfRXVgeL4rjNbJ3NZxnBwxrp7xAljZ7NJ+t12larU1Vi0hmF
# DUTenEIBjg==
# SIG # End signature block
