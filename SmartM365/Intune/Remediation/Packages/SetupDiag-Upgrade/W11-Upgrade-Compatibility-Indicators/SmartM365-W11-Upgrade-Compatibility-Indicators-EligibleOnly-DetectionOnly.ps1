<#
    Name: SmartM365-W11-Upgrade-Compatibility-Indicators-EligibleOnly-DetectionOnly.ps1
    Version: 1.3
    Description: Detects Windows 11 upgrade compatibility blockers, ignoring devices that are only hardware-ineligible.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 20)]
    [int]$MaxEntriesToDisplay = 5,

    [ValidateRange(20, 300)]
    [int]$MaxValueLength = 100
)

$ErrorActionPreference = "Stop"

$EffectiveMaxEntriesToDisplay = $MaxEntriesToDisplay
$IndicatorRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators"
$NonBlockingPlaceholderValues = @("None", "N/A", "NotApplicable", "Not Applicable")
$IgnoredHardwareIneligibleReasons = @(
    "Cpu",
    "CpuFms",
    "CpuVendor",
    "CpuFamily",
    "CpuModel",
    "CpuSpeed",
    "CpuCores",
    "Tpm",
    "TpmVersion",
    "SecureBoot",
    "UefiSecureBoot",
    "Uefi",
    "Ram",
    "Memory",
    "Storage",
    "Disk",
    "SystemDriveSize"
)
$IgnoredNonWindows11TargetVersions = @(
    "1507",
    "1511",
    "1607",
    "1703",
    "1709",
    "1803",
    "1809",
    "1903",
    "1909",
    "19H1",
    "19H2",
    "2004",
    "20H1",
    "20H2",
    "21H1",
    "TH1",
    "TH2",
    "RS1",
    "RS2",
    "RS3",
    "RS4",
    "RS5"
)

function ConvertTo-SafeIndicatorValue {
    param(
        [AllowNull()]$Value,

        [Parameter(Mandatory = $true)]
        [int]$MaximumLength
    )

    if ($null -eq $Value) {
        return ""
    }

    $safeValues = @($Value) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and
            $_ -notin $NonBlockingPlaceholderValues
        }

    $safeValue = $safeValues -join ", "
    $safeValue = $safeValue -replace "[^\x20-\x7E]", "?"
    if ($safeValue.Length -gt $MaximumLength) {
        return $safeValue.Substring(0, $MaximumLength) + "..."
    }

    return $safeValue
}

function ConvertTo-CompactValueList {
    param(
        [AllowNull()]$Entries,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $true)]
        [int]$MaximumEntries,

        [Parameter(Mandatory = $true)]
        [int]$MaximumLength
    )

    $values = @($Entries) |
        Where-Object { $null -ne $_ } |
        ForEach-Object {
            $property = $_.PSObject.Properties[$PropertyName]
            if ($null -ne $property) {
                [string]$property.Value
            }
        } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique |
        Select-Object -First $MaximumEntries

    $value = $values -join "|"
    if ($value.Length -gt $MaximumLength) {
        return $value.Substring(0, $MaximumLength) + "..."
    }

    return $value
}

function Write-DetectionSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [int]$IndicatorCount,

        [int]$BlockingCount = 0,

        [int]$IgnoredHardwareIneligibleCount = 0,

        [int]$IgnoredNonWindows11TargetCount = 0,

        [int]$ReadErrorCount = 0,

        [AllowNull()]$ActionableEntries,

        [AllowNull()][string[]]$FailedTargets
    )

    $summaryParts = @(
        "Status=$Status",
        "IndicatorCount=$IndicatorCount",
        "BlockingCount=$BlockingCount",
        "IgnoredHardwareIneligibleCount=$IgnoredHardwareIneligibleCount",
        "IgnoredNonWindows11TargetCount=$IgnoredNonWindows11TargetCount"
    )

    if ($ReadErrorCount -gt 0) {
        $summaryParts += "ReadErrorCount=$ReadErrorCount"
    }

    $targetVersions = ConvertTo-CompactValueList -Entries $ActionableEntries -PropertyName "TargetVersion" -MaximumEntries $EffectiveMaxEntriesToDisplay -MaximumLength $MaxValueLength
    $gatedBlockIds = ConvertTo-CompactValueList -Entries $ActionableEntries -PropertyName "GatedBlockId" -MaximumEntries $EffectiveMaxEntriesToDisplay -MaximumLength $MaxValueLength
    $redReasons = ConvertTo-CompactValueList -Entries $ActionableEntries -PropertyName "RedReason" -MaximumEntries $EffectiveMaxEntriesToDisplay -MaximumLength $MaxValueLength
    $sysReqIssues = ConvertTo-CompactValueList -Entries $ActionableEntries -PropertyName "SysReqIssue" -MaximumEntries $EffectiveMaxEntriesToDisplay -MaximumLength $MaxValueLength

    if (-not [string]::IsNullOrWhiteSpace($targetVersions)) { $summaryParts += "Targets=$targetVersions" }
    if (-not [string]::IsNullOrWhiteSpace($gatedBlockIds)) { $summaryParts += "GatedBlockIds=$gatedBlockIds" }
    if (-not [string]::IsNullOrWhiteSpace($redReasons)) { $summaryParts += "RedReasons=$redReasons" }
    if (-not [string]::IsNullOrWhiteSpace($sysReqIssues)) { $summaryParts += "SysReqIssues=$sysReqIssues" }

    if ($ReadErrorCount -gt 0 -and $null -ne $FailedTargets -and $FailedTargets.Count -gt 0) {
        $failedTargetValue = (@($FailedTargets) | Select-Object -First $EffectiveMaxEntriesToDisplay) -join "|"
        if ($failedTargetValue.Length -gt $MaxValueLength) {
            $failedTargetValue = $failedTargetValue.Substring(0, $MaxValueLength) + "..."
        }

        $summaryParts += "FailedTargets=$failedTargetValue"
    }

    Write-Output ($summaryParts -join "; ")
}

function Get-IndicatorValueToken {
    param(
        [AllowNull()][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @($Value -split "[,; ]+") |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

function Test-BlockingIndicator {
    param(
        [Parameter(Mandatory = $true)]$Indicator
    )

    return (
        $Indicator.UpEx -match "(^|,\s*)(Red|Blocked|Hold)(\s*,|$)" -or
        -not [string]::IsNullOrWhiteSpace($Indicator.GatedBlockId) -or
        -not [string]::IsNullOrWhiteSpace($Indicator.RedReason) -or
        -not [string]::IsNullOrWhiteSpace($Indicator.SysReqIssue)
    )
}

function Test-IgnoredTargetVersion {
    param(
        [AllowNull()][string]$TargetVersion
    )

    if ([string]::IsNullOrWhiteSpace($TargetVersion)) {
        return $false
    }

    return ($TargetVersion.Trim() -in $IgnoredNonWindows11TargetVersions)
}

function Test-ActionableBlockingIndicator {
    param(
        [Parameter(Mandatory = $true)]$Indicator
    )

    if (-not (Test-BlockingIndicator -Indicator $Indicator)) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($Indicator.GatedBlockId)) {
        return $true
    }

    $reasonTokens = @()
    $reasonTokens += Get-IndicatorValueToken -Value $Indicator.RedReason
    $reasonTokens += Get-IndicatorValueToken -Value $Indicator.SysReqIssue

    if ($reasonTokens.Count -eq 0) {
        return ($Indicator.UpEx -match "(^|,\s*)(Red|Blocked|Hold)(\s*,|$)")
    }

    foreach ($reasonToken in $reasonTokens) {
        if ($reasonToken -notin $IgnoredHardwareIneligibleReasons) {
            return $true
        }
    }

    return $false
}

function Get-IndicatorReasonText {
    param(
        [Parameter(Mandatory = $true)]$Indicator
    )

    $reasonParts = @()
    if (-not [string]::IsNullOrWhiteSpace($Indicator.UpEx)) { $reasonParts += "UpEx=$($Indicator.UpEx)" }
    if (-not [string]::IsNullOrWhiteSpace($Indicator.GatedBlockId)) { $reasonParts += "GatedBlockId=$($Indicator.GatedBlockId)" }
    if (-not [string]::IsNullOrWhiteSpace($Indicator.RedReason)) { $reasonParts += "RedReason=$($Indicator.RedReason)" }
    if (-not [string]::IsNullOrWhiteSpace($Indicator.SysReqIssue)) { $reasonParts += "SysReqIssue=$($Indicator.SysReqIssue)" }

    return ($reasonParts -join "; ")
}

try {
    if (-not (Test-Path -LiteralPath $IndicatorRegistryPath)) {
        Write-Output "Status=NoIndicatorsPath; Path=$IndicatorRegistryPath"
        exit 0
    }

    $indicatorEntries = New-Object System.Collections.Generic.List[object]
    $readErrors = New-Object System.Collections.Generic.List[string]
    $ignoredNonWindows11Targets = New-Object System.Collections.Generic.List[string]
    $subKeys = Get-ChildItem -LiteralPath $IndicatorRegistryPath -ErrorAction Stop

    foreach ($subKey in $subKeys) {
        try {
            if (Test-IgnoredTargetVersion -TargetVersion $subKey.PSChildName) {
                $ignoredNonWindows11Targets.Add($subKey.PSChildName)
                continue
            }

            $properties = Get-ItemProperty -LiteralPath $subKey.PSPath -ErrorAction Stop
            $entry = [PSCustomObject]@{
                TargetVersion = $subKey.PSChildName
                UpEx          = ConvertTo-SafeIndicatorValue -Value $properties.UpEx -MaximumLength $MaxValueLength
                GatedBlockId  = ConvertTo-SafeIndicatorValue -Value $properties.GatedBlockId -MaximumLength $MaxValueLength
                RedReason     = ConvertTo-SafeIndicatorValue -Value $properties.RedReason -MaximumLength $MaxValueLength
                SysReqIssue   = ConvertTo-SafeIndicatorValue -Value $properties.SysReqIssue -MaximumLength $MaxValueLength
            }

            $indicatorEntries.Add($entry)
        }
        catch {
            $readErrors.Add($subKey.PSChildName)
        }
    }

    if ($indicatorEntries.Count -eq 0) {
        if ($readErrors.Count -gt 0) {
            Write-DetectionSummary -Status "ReadError" -IndicatorCount 0 -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count -ReadErrorCount $readErrors.Count -FailedTargets @($readErrors)
            exit 1
        }

        if ($ignoredNonWindows11Targets.Count -gt 0) {
            Write-DetectionSummary -Status "NoRelevantWindows11Indicators" -IndicatorCount 0 -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count
            exit 0
        }

        Write-DetectionSummary -Status "NoIndicators" -IndicatorCount 0
        exit 0
    }

    $blockingEntries = @($indicatorEntries | Where-Object { Test-BlockingIndicator -Indicator $_ })
    $actionableBlockingEntries = @($blockingEntries | Where-Object { Test-ActionableBlockingIndicator -Indicator $_ })
    $ignoredHardwareEntries = @($blockingEntries | Where-Object { -not (Test-ActionableBlockingIndicator -Indicator $_) })

    if ($actionableBlockingEntries.Count -gt 0) {
        Write-DetectionSummary `
            -Status "ActionableBlockingConditionDetected" `
            -IndicatorCount $indicatorEntries.Count `
            -BlockingCount $actionableBlockingEntries.Count `
            -IgnoredHardwareIneligibleCount $ignoredHardwareEntries.Count `
            -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count `
            -ReadErrorCount $readErrors.Count `
            -ActionableEntries $actionableBlockingEntries `
            -FailedTargets @($readErrors)
        exit 1
    }

    if ($readErrors.Count -gt 0) {
        Write-DetectionSummary `
            -Status "PartialReadNoActionableBlockingCondition" `
            -IndicatorCount $indicatorEntries.Count `
            -IgnoredHardwareIneligibleCount $ignoredHardwareEntries.Count `
            -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count `
            -ReadErrorCount $readErrors.Count `
            -FailedTargets @($readErrors)
        exit 1
    }

    if ($ignoredHardwareEntries.Count -gt 0) {
        Write-DetectionSummary `
            -Status "OnlyHardwareIneligibleIndicatorsIgnored" `
            -IndicatorCount $indicatorEntries.Count `
            -IgnoredHardwareIneligibleCount $ignoredHardwareEntries.Count `
            -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count
        exit 0
    }

    Write-DetectionSummary -Status "NoBlockingConditionDetected" -IndicatorCount $indicatorEntries.Count -IgnoredNonWindows11TargetCount $ignoredNonWindows11Targets.Count
    exit 0
}
catch {
    Write-Output "Status=Error; Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCfIHgCh7ZF/V7Z
# 3xL/mYsKeBq98xaIb/35gKeTRfe+tKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBUar52ql8Krbn0//wWC3BeUDQJYl4pIzKNp5Iuv090+MA0GCSqG
# SIb3DQEBAQUABIIBgC0MaK18NxRK55qT2XzalEw0hQ4J3xeUVU2mD3ExReL1bE5p
# 2xbDiPbCfY1oxbWTmVmAa91as6AC0jxfjNyve/1lzxzZX629LVxULHqnVDzzayB6
# QdazHLTtfoDZlw1gr9qS2mxwnKQsKxd0NydcAtwEWdPtpJAs7x+dnmXGTpvpuJTy
# JdZ/Wv96C7ooX6CEBIz2eBZLPcffcGyhNFJiai/ob0zqYez68aNSd3XXu6UfTmhk
# CHQDqTwb8R2MrV61XDUo97SE7C8vNXJIdMIiCfTz7E2twFo9mf3ifbm0FZhCL9pt
# Gx44iRUTGgLivhgaxnr9Rz5uyMZ8FJA9ZiIrtG7L4xxW4qQ1CaBT7Pr2SjjZ5a4P
# yUjjDdPnP7S7ZCM01Vh898n/M7jh62DytYnfYLS8YqBDHploME0DCMED0B2j4uyH
# evzgFi5/fgBWX0kTeGUMzNwB1QNZqwWglzc+QcZ4JwI4QF2GX6869smDRoG4PkZ1
# OpJyFk8hr2Stl1JaLqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NTJaMC8GCSqGSIb3DQEJBDEiBCALjloQov6KsmKnW7kzvnFUtTO3Wt7CylscRosg
# eXpB2TANBgkqhkiG9w0BAQEFAASCAgAxNFx3VKjNoAuLzNYJeMnfggKHfubQAQ98
# RM+oBiNBl9LaHyWT4oaE6KQDzhFmeiXjPgcbicRkEgxhG73Zs1iReZ/xL4Soya9A
# pfDVXz11beeuYo881n+7Q353g38FnheJJJc8MZGSHNc8eobkxdFOaoUJ5OBxXm6j
# 2688AxxAar9rRtu77j1vSiOxbolx+LHb+B99l5+L993NmX+qxL1aEikkC++FFVnc
# IOJBFShmR28FmXrOZVVEkyg+QcUD/asfxYlUmH4BcOZnSNRKMffWOC49fTi7/MP/
# uB171KIonzpVeODgBXhTVBFeTL9+iRC87ZijRI7medWU8CAnb5uQ6F34gFNWB/83
# FQWaliuQ1516MKiuKUJgwGq8qm0EYsN3nIcRZQrWLSiSGjEZrS7RTzRjjcw5/M1T
# 4IJZjzSCmmWrYa/Qjtn8rFmv5blB7HqMx9R/KzzawQogld2j/rkq2bA8Pq0tV5ty
# c2SWBhKNm24a2OLO0LwaIaBVjAa7BjLZWp/fqDcZ6nmVba2TmaS4gjpunYs6gkNZ
# kw5mbXQTNDcouS1fuuC0rzO0tocCNHxNXhtpjvpyDqOJZ/j1p4jcyrYGcYDos9Ox
# DBrtmDoUMRuWzG3r5cQBeb02yotgdTEUBbcVtM/89fNDdjYdFnTd/rmFoZtI2Hrt
# VQG5zLMXmA==
# SIG # End signature block
