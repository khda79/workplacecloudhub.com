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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIKKktbi7mbEp4
# QwiKVE65aO+rv79jv8PSFZKbScPN9aCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDT8EKz7nOZNhJ53Bc/
# HGhvvEhAXVv4EzStIMQcFWeGujANBgkqhkiG9w0BAQEFAASCAYCag/DjL3T+H3i5
# KkhR02ux6w2Cew2vTyD4Pm9z6qqLzGUzuWtriemd5hJKFw1ZdANRH7PJ5cPD39mR
# tygY1Wg9yn9xixPO/yv3Nvf2rQsqtWPEluO5DFVRXgTRjfRZ0XZdghFx1AAOivSk
# 3lt1+poa3thaNJK6cws2YMgOfaCRblh3C0ZjHI2nJ9WifzY4/sHx8Xj/mSv/0YxE
# hS+yctKTKTZyFvbVsAit4bV5O65ST7xi0i5q/oyKBXxa0eBLMWebGemsdyltyOZ3
# +x8a4r66Ck9flsJvwGphjtGBoGELgl2xwd36T4iJ9rmCz2w6lsU8E330MjlU06Lo
# uLNb72CxcG6L+zi3C0qdDq/ZtOGMxWVUt4HxMvjrHKlzneR5N35e0fVNoek6GT6T
# 9oLZxlXTD5Dp93bzQBlq+5BCrQq8eE02PaTAh84VvyzS6f8J19MNsAOtlPdYKyUy
# Yiuk8yxWg5hvV80hdK4YeMufgfoAxCvcAY7F/rZgRr3SsHzOUtg=
# SIG # End signature block
