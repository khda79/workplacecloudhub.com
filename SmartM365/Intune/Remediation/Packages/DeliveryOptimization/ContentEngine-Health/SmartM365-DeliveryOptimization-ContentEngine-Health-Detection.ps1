<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Detection.ps1
    Version: 1.1
    Description: Detects actionable Delivery Optimization, Dynamic Download, BITS, and content download issues without flagging common informational event noise.

    Detection goals:
    - Keep Intune output compact and single-line
    - Does not treat a missing Delivery Optimization folder as an error
    - Does not treat empty event logs as an error
    - Ignores BITS peer helper event 310 with 0x80070032, which is common non-actionable noise
    - Detects Delivery Optimization and Dynamic Download errors
    - Detects BITS download errors related to content transfer
    - Detects Windows Update errors only when clearly related to Delivery Optimization or Dynamic Download

    Exit codes:
    0 = Healthy
    1 = Issues detected
#>

[CmdletBinding()]
param(
    [int]$MaxEvents = 200,
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MinimumCacheSizeBytes = 1024
)

$ErrorActionPreference = "Stop"

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 180
    )

    $compactText = ($Text -replace "\s+", " ").Trim()

    if ($compactText.Length -gt $MaxLength) {
        return ($compactText.Substring(0, $MaxLength) + "...")
    }

    return $compactText
}

function Write-IntuneResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [hashtable]$Data = @{}
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Status=$Status")

    foreach ($key in ($Data.Keys | Sort-Object)) {
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 320
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Get-EventMessage {
    param(
        [AllowNull()]
        [object]$EventRecord
    )

    if ($null -eq $EventRecord -or [string]::IsNullOrWhiteSpace($EventRecord.Message)) {
        return ""
    }

    return (($EventRecord.Message -replace "`r|`n", " ").Trim())
}

function Get-RecentEventsSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogName,

        [Parameter(Mandatory = $true)]
        [datetime]$CutoffDate,

        [Parameter(Mandatory = $true)]
        [int]$MaxEvents
    )

    try {
        $events = @(
            Get-WinEvent -LogName $LogName -MaxEvents $MaxEvents -ErrorAction SilentlyContinue |
            Where-Object { $_.TimeCreated -ge $CutoffDate }
        )

        if ($null -eq $events) {
            return @()
        }

        return $events
    }
    catch {
        return @()
    }
}

function Test-BitsPeerHelperNoise {
    param(
        [Parameter(Mandatory = $true)]
        [object]$EventRecord
    )

    $message = Get-EventMessage -EventRecord $EventRecord

    if ($EventRecord.Id -ne 310) {
        return $false
    }

    if ($message -match "0x80070032") {
        return $true
    }

    if ($message -match "peer helper|modules auxiliaires|Peerhilfs") {
        return $true
    }

    return $false
}

function Test-DoRelatedMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    return (
        $Message -match "Delivery Optimization" -or
        $Message -match "DeliveryOptimization" -or
        $Message -match "DynamicDownload" -or
        $Message -match "0x80D0[0-9A-Fa-f]{4}"
    )
}

function Test-ActionableMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    return (
        $Message -match "0x80D0[0-9A-Fa-f]{4}" -or
        $Message -match "\berror\b" -or
        $Message -match "\bfail(?:ed|ure)?\b" -or
        $Message -match "\btimeout\b" -or
        $Message -match "\bproxy\b" -or
        $Message -match "download.*fail" -or
        $Message -match "fail.*download"
    )
}

function Test-ActionableContentEvent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [object]$EventRecord
    )

    $message = Get-EventMessage -EventRecord $EventRecord

    if ([string]::IsNullOrWhiteSpace($message)) {
        return $false
    }

    switch ($Source) {
        "DO" {
            return (Test-ActionableMessage -Message $message)
        }
        "WU" {
            return ((Test-DoRelatedMessage -Message $message) -and (Test-ActionableMessage -Message $message))
        }
        "BITS" {
            if (Test-BitsPeerHelperNoise -EventRecord $EventRecord) {
                return $false
            }

            return (
                $message -match "0x80D0[0-9A-Fa-f]{4}" -or
                $message -match "0x802000[0-9A-Fa-f]{2}" -or
                $message -match "\bBG_E_" -or
                ((Test-DoRelatedMessage -Message $message) -and (Test-ActionableMessage -Message $message)) -or
                ($message -match "download.*fail" -or $message -match "fail.*download")
            )
        }
    }

    return $false
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Events,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$IssueList,

        [Parameter(Mandatory = $true)]
        [int]$Limit,

        [Parameter(Mandatory = $true)]
        [int]$MaxIssueLength
    )

    if ($null -eq $Events -or $Events.Count -eq 0) {
        return
    }

    foreach ($eventRecord in $Events) {
        if ($null -eq $eventRecord) {
            continue
        }

        if ($IssueList.Count -ge $Limit) {
            return
        }

        if (-not (Test-ActionableContentEvent -Source $Source -EventRecord $eventRecord)) {
            continue
        }

        $message = Format-CompactText -Text (Get-EventMessage -EventRecord $eventRecord) -MaxLength $MaxIssueLength
        $IssueList.Add("${Source}:Time=$($eventRecord.TimeCreated);Id=$($eventRecord.Id);Message=$message")
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $ignoredNoiseCount = 0
    $checkedEventCount = 0

    $deliveryOptimizationFolder = "C:\ProgramData\Microsoft\Windows\DeliveryOptimization"
    $deliveryOptimizationFolderStatus = "NotCreated"
    $deliveryOptimizationSizeMB = 0

    if (Test-Path -Path $deliveryOptimizationFolder) {
        $deliveryOptimizationFiles = @(Get-ChildItem -Path $deliveryOptimizationFolder -Recurse -File -ErrorAction SilentlyContinue)
        $deliveryOptimizationSizeBytes = 0

        if ($deliveryOptimizationFiles.Count -gt 0) {
            $deliveryOptimizationSizeBytes = ($deliveryOptimizationFiles | Measure-Object -Property Length -Sum).Sum

            if ($null -eq $deliveryOptimizationSizeBytes) {
                $deliveryOptimizationSizeBytes = 0
            }
        }

        $deliveryOptimizationSizeMB = [math]::Round(($deliveryOptimizationSizeBytes / 1MB), 2)
        $deliveryOptimizationFolderStatus = "Present"

        if ($deliveryOptimizationSizeBytes -lt $MinimumCacheSizeBytes) {
            $deliveryOptimizationFolderStatus = "EmptyOrSmall"
        }
    }

    $windowsUpdateEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $deliveryOptimizationEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-DeliveryOptimization/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $bitsEvents = Get-RecentEventsSafe `
        -LogName "Microsoft-Windows-Bits-Client/Operational" `
        -CutoffDate $cutoffDate `
        -MaxEvents $MaxEvents

    $checkedEventCount = $windowsUpdateEvents.Count + $deliveryOptimizationEvents.Count + $bitsEvents.Count
    $ignoredNoiseCount = @($bitsEvents | Where-Object { Test-BitsPeerHelperNoise -EventRecord $_ }).Count

    Add-EventIssue `
        -Source "WU" `
        -Events $windowsUpdateEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    Add-EventIssue `
        -Source "DO" `
        -Events $deliveryOptimizationEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    Add-EventIssue `
        -Source "BITS" `
        -Events $bitsEvents `
        -IssueList $issues `
        -Limit $MaxIssuesToDisplay `
        -MaxIssueLength $MaxIssueLength

    $uniqueIssues = @($issues | Select-Object -Unique)

    if ($uniqueIssues.Count -gt 0) {
        $sourceSummary = ($uniqueIssues | ForEach-Object {
            if ($_ -match "^([^:]+):") { $matches[1] } else { "Unknown" }
        } | Group-Object | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ","

        Write-IntuneResult -Status "DeliveryOptimizationContentEngineIssuesDetected" -Data @{
            IssueCount = $uniqueIssues.Count
            Sources = $sourceSummary
            CheckedEvents = $checkedEventCount
            IgnoredNoise = $ignoredNoiseCount
            DOFolderStatus = $deliveryOptimizationFolderStatus
            DOFolderSizeMB = $deliveryOptimizationSizeMB
            Examples = (($uniqueIssues | Select-Object -First $MaxIssuesToDisplay) -join " | ")
        }

        exit 1
    }

    Write-IntuneResult -Status "DeliveryOptimizationContentEngineHealthy" -Data @{
        CheckedEvents = $checkedEventCount
        IgnoredNoise = $ignoredNoiseCount
        DOFolderStatus = $deliveryOptimizationFolderStatus
        DOFolderSizeMB = $deliveryOptimizationSizeMB
    }

    exit 0
}
catch {
    Write-IntuneResult -Status "TechnicalError" -Data @{
        Message = $_.Exception.Message
    }

    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9/tVdUWhlp/M4
# ovZv1GummKtyGwiaL45FXn9u2cyNAqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBMNLPCJ0D6I5KJEzMvdFxlZ4y5FSp9i1r6wxZMRe1K/zANBgkqhkiG9w0B
# AQEFAASCAYB2qEvKwDr0zS7ALYfiPtVMFXMhmOOKpiiSo9YGjR/s3POO9VKyB1tS
# PKGZgd9ayT4VVj+gMZubtzJnuU20FhBYW6VvbVr+pyxLqv3KcB6R17lc3gmwbBEh
# xF4S76z4SO9tHpuh6I3H4fLmCz4E+xDVkYzuXc2HpGstc/DjPZFHpDs8VcHwjdVc
# mg8AFgZ2IzmYBVtD7byooHDpV8w0xnLBq8iwbmrwROpmybMWr7mton5bABnuF/x8
# hiNrzVQuozbirGXz0fNztPbLR3x63v+XirXolvRM8sswztgTEZk1oJHn1RsiFe1f
# C+y98M/ociXTwQU+lwJ7ET6Ft8vJIrq0iRCpHvMS6ql2KbtfoArssqgeOYf5FrJo
# 1JlyYk+M58AwsiboXMfLFR9p1U2Jgxarq/N51USTq8gvWHxWBjOVFNUmaSMQiH8M
# KM0TZ5w256hx+YRPK0pjYdEp0ajPvyy0jTTSiRQVSJSKNu9KH8xRV7GtU6k9JAi9
# Qthn9rBerEQ=
# SIG # End signature block
