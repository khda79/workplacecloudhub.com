<#
.SYNOPSIS
    Version: 1.0
    Generates a fresh WindowsUpdate.log diagnostic artifact after actionable Windows Update log errors are detected.

.VERSION
    1.0

.DESCRIPTION
    This remediation is intentionally diagnostic-only.
    It does not reset Windows Update or change device state, because WindowsUpdate.log errors can represent many different causes.
    Repair actions are covered by the dedicated Windows Update remediation packages.
    The generated log is stored under C:\ProgramData\SmartM365\IntuneRemediation\Temp\WindowsUpdateLog.
#>

[CmdletBinding()]
param(
    [int]$MinimumLogSizeBytes = 1024,
    [int]$MaxReasonLength = 180
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
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

try {
    $localFolder = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Temp\WindowsUpdateLog"
    $computerName = $env:COMPUTERNAME
    $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $outputFile = "WindowsUpdate-Remediation-$date-$computerName.log"
    $localPath = Join-Path -Path $localFolder -ChildPath $outputFile

    if (-not (Test-Path -Path $localFolder)) {
        New-Item -Path $localFolder -ItemType Directory -Force | Out-Null
    }

    try {
        Get-WindowsUpdateLog -LogPath $localPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-IntuneResult -Status "WindowsUpdateLogUnavailable" -Data @{
            Reason = (Format-CompactText -Text $_.Exception.Message -MaxLength $MaxReasonLength)
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

    Write-IntuneResult -Status "WindowsUpdateLogGenerated" -Data @{
        LogPath = $localPath
        LogSizeBytes = $logFile.Length
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = (Format-CompactText -Text $_.Exception.Message -MaxLength $MaxReasonLength)
    }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB8Otkdmt/hmH4T
# KaBoMi0zkgGYYgyNdjT5CrkTYPT1sKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCA2yxtvSRiBQDSyUt6gnFO+MwzbuFmYdPUYCHyrlN0IKjANBgkqhkiG9w0B
# AQEFAASCAYAhoC/2Zc9vpUJL0O5uxn3vbU9/M5c+hYixeOqjmhpMKcCv4PAXsVN3
# ha+qWzqMw3r7Yvz35+Z8p0X+hM4onhVYenP9tgOp2jlVwqTcNRvECDr3dTHOzBsQ
# 7cywBE0GzaZOQvlvvntMLfoMBt0GtQ/t+r2WgGwmFrqnrO85fHW9x78nIM01gn7N
# jaX1EEswt1t2wsEN63dxYUpOC2LWu6s8XGudyj2lrYzpAwDRTvAH5uAKJICvFOXu
# 6rRW1yg6NFP89EBfF7vXvpbidgw3czawrF8rjRXs8fxzHqe62p/xNJGuRiVyiikC
# K3yswKhGD3WTV64RWHSOQVhy6y67gsLHA3+gOZM+W3mQ8cAeLpZ7ZmsNs51kaeEP
# PBYmXxt1XbsWQ/j7kshoSj01GyU2IpwLjMNDOiS1X5E0ahbFI1plEwKvPm1H4TBx
# AJRKDAQDJS6r5RAO3mYk04E+zIrWyN7CA0qwFqLbq2zfWBDmRiT+OOWv4+89QSDP
# 6S4Gs5x4YI8=
# SIG # End signature block
