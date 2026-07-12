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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIKKktbi7mbEp4
# QwiKVE65aO+rv79jv8PSFZKbScPN9aCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDT8EKz7nOZNhJ53Bc/HGhvvEhAXVv4EzStIMQcFWeGujANBgkqhkiG9w0B
# AQEFAASCAYC16zbRCSgieeC5mtzRYbHXw6CyiKoH8cgFJg+qu2FegkbADBgAr070
# l3uNZM0OhRDV/FP8HnXgN5JNhXvan2nl3l9cw6C90QYZBfZd8m09SCmYAbjCr9Dr
# KGNtiKuQ/r7Sx253t/xbYCr5cFnMwJI5Dx2SDL3uTvFTgPF/gw0boXTKjNdkmM+U
# mvqDg2ePm9FXHa2YzjseV1+Sr9AplfAwodrwQnGueHYZxE8c1yGhTwmQFTNIMHc1
# V02L9U85EyAt/yObhH8aQMjUcz30GteIjxmWTyvX+I7lEynutlUg1gTWAUxZOXEk
# TBOomskRYlCniA3OsnFUIg3OcE/S68ML2GPSuKTpCjK2e4oRnfp38mnYFugVZ/+C
# /CGSkYEd9BAM+OvqSRTfLtCTZyoPknF+NjdRuIDLzZXCQbl5xcNBnMPT9+ITgMhs
# ynajXr65sT9642PeUjw+9X4P0hx15W07mrok9v1VFTbsfD6TFn8/sYRsQuB1DAet
# 1miorKOZGA8=
# SIG # End signature block
