<#
.SYNOPSIS
    Version: 1.1
    Generates WindowsUpdate.log and detects recent Windows Update errors.

.VERSION
    1.1

.DESCRIPTION
    Generates WindowsUpdate.log in C:\ProgramData\SmartM365\IntuneRemediation\Temp.
    Extracts actionable Windows Update related error lines.
    Ignores benign trace lines such as "error 0", flags, product type values, and successful progress callbacks.
    Emits a compact single-line output for Intune readability.
#>

[CmdletBinding()]
param(
    [int]$MaxErrorsToDisplay = 3,
    [int]$MinimumLogSizeBytes = 1024,
    [int]$MaxLineLength = 140
)

$ErrorActionPreference = "Stop"

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 140
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
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Test-BenignWindowsUpdateLogLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Line
    )

    $benignPatterns = @(
        "\bNo error\b",
        "\berror\s*[:=]?\s*0\b",
        "\berror\s*[:=]?\s*0x0+\b",
        "\bSucceeded\b",
        "\bsuccessfully\b",
        "\bCompleted\b",
        "\bRefresh complete\b",
        "\* START \*",
        "\bFlags:\s*0X[0-9A-Fa-f]+\b",
        "\bOS Product Type\s*=\s*0x[0-9A-Fa-f]+\b",
        "\bauth token of type\s*0x[0-9A-Fa-f]+\b",
        "\bcode Call (progress|complete) and error 0\b"
    )

    foreach ($pattern in $benignPatterns) {
        if ($Line -match $pattern) {
            return $true
        }
    }

    return $false
}

function Test-ActionableWindowsUpdateLogLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Line
    )

    if (Test-BenignWindowsUpdateLogLine -Line $Line) {
        return $false
    }

    if ($Line -match "\*FAILED\*\s*\[[0-9A-Fa-f]{8}\]") {
        return $true
    }

    if ($Line -match "\b(fatal|failed|failure)\b") {
        return $true
    }

    if ($Line -match "\berror\b" -and $Line -match "\b(0x[0-9A-Fa-f]{8}|[78][0-9A-Fa-f]{7}|C[0-9A-Fa-f]{7}|80D[0-9A-Fa-f]{5})\b") {
        return $true
    }

    return $false
}

try {
    $localFolder = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Temp\WindowsUpdateLog"
    $computerName = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $outputFile = "WindowsUpdate-$date-$computerName.log"
    $localPath = Join-Path -Path $localFolder -ChildPath $outputFile

    if (-not (Test-Path -Path $localFolder)) {
        New-Item -Path $localFolder -ItemType Directory -Force | Out-Null
    }

    try {
        Get-WindowsUpdateLog -LogPath $localPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-IntuneResult -Status "WindowsUpdateLogUnavailable" -Data @{
            Reason = $_.Exception.Message
        }
        exit 0
    }

    if (-not (Test-Path -Path $localPath -PathType Leaf)) {
        Write-IntuneResult -Status "WindowsUpdateLogNotGenerated"
        exit 0
    }

    $logFile = Get-Item -Path $localPath -ErrorAction Stop

    if ($logFile.Length -lt $MinimumLogSizeBytes) {
        Write-IntuneResult -Status "WindowsUpdateLogTooSmall" -Data @{
            LogSizeBytes = $logFile.Length
        }
        exit 0
    }

    $candidatePatterns = @("\berror\b", "\bfailed\b", "\bfailure\b", "\bfatal\b", "\*FAILED\*")
    $rawMatches = @(Select-String -Path $localPath -Pattern $candidatePatterns -CaseSensitive:$false -ErrorAction Stop)
    $filteredErrors = New-Object System.Collections.Generic.List[string]

    foreach ($match in $rawMatches) {
        if (Test-ActionableWindowsUpdateLogLine -Line $match.Line) {
            $filteredErrors.Add((Format-CompactText -Text $match.Line -MaxLength $MaxLineLength))
        }
    }

    if ($filteredErrors.Count -gt 0) {
        $sampleErrors = @($filteredErrors | Select-Object -First $MaxErrorsToDisplay)

        Write-IntuneResult -Status "WindowsUpdateErrorsDetected" -Data @{
            ErrorCount = $filteredErrors.Count
            LogSizeBytes = $logFile.Length
            Samples = ($sampleErrors -join " | ")
        }

        exit 1
    }

    Write-IntuneResult -Status "NoActionableWindowsUpdateErrorsFound" -Data @{
        CandidateLineCount = $rawMatches.Count
        LogSizeBytes = $logFile.Length
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = $_.Exception.Message
    }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/wZ31uMwgPmOC
# qBqOgy65PHAp5GLkez00ECdlE29WqqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDKjTK1glZ+4WGWrACiB0esA21kUF1/2s0xKmbZbpEZOjANBgkqhkiG9w0B
# AQEFAASCAYC59p22GAUVPe2tRNCeaF+ce51VQl6kln/vENpJ3vAqJUscgeFv0g5i
# wiUtR19CS2hyXdviDyqHaU8QVW3wvPKOPTZDlNk9l2zlA3ZMfwvxfbzfHHwcMPtm
# ApCnLqKxn6lnXx3WQ9XCtAatix0dq9rWikkaitfPni+xrN2/PtKPWUBVrvwOz7bs
# A3TBnmzqZhqiVEK9nL/VRTog6IWslPtWcaccZi1uoQzKJmxsi4rr88vilbMe5jeX
# MQFjvNrY12JF/rUk06jN6A9ZiZfotZtEWu+Af0mjHKiOJlXSvUgcPDSpI2H48bJs
# vYvyTvW8xPZM5y1adWWAAJ95b91X4ODj+nOPvDSgMO3u1oRkH+MrBDhJE2AXyAwA
# UzBz/XKCpsHxFZ6rqAZGoGCe1FI7Xa3gPsf6uHjaHQvfVPxrOX5Q9AEMuVEdSTnn
# 1e9M07nX83Zovp8teyjH/Dj3BRE1JLZ0NpGE6yO00ChZyPDd6FLA9LV3mM7mKNGk
# QnBPyL+GZLk=
# SIG # End signature block
