<#
.SYNOPSIS
    Version: 1.0
    Repairs Intune Management Extension health for IntuneManagementExtension-Health.
.DESCRIPTION
    Ensures the IME service is configured to start automatically, recreates the expected log directory, restarts the IME service, starts EnterpriseMgmt tasks, and triggers a Windows Update scan to refresh policy-driven activity.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'IntuneManagementExtension-Health'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"

function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }
function Invoke-UsoClient { param([string]$Action) $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'; if (Test-Path -LiteralPath $uso) { Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue; Write-SmartM365Log "UsoClient $Action triggered." } }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    New-Item -Path 'C:\ProgramData\SmartM365\IntuneRemediation\Logs' -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'

    $service = Get-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
    Set-Service -Name 'IntuneManagementExtension' -StartupType Automatic -ErrorAction SilentlyContinue
    if ($service.Status -eq 'Running') {
        Restart-Service -Name 'IntuneManagementExtension' -Force -ErrorAction Stop
        Write-SmartM365Log 'Intune Management Extension service restarted.'
    }
    else {
        Start-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
        Write-SmartM365Log 'Intune Management Extension service started.'
    }

    Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Disabled' -and ($_.TaskName -like '*PushLaunch*' -or $_.TaskName -like '*Schedule*') } |
        ForEach-Object {
            try { Start-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop; Write-SmartM365Log "Started scheduled task: $($_.TaskPath)$($_.TaskName)" } catch { Write-SmartM365Log "Could not start scheduled task $($_.TaskName): $($_.Exception.Message)" }
        }

    Invoke-UsoClient -Action 'RefreshSettings'
    Invoke-UsoClient -Action 'StartScan'

    Write-SmartM365Log 'Remediation completed.'
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC9Znw9MBANPaDl
# zYHhFKa1BX0/VZeoCYvVzPUKw8ZCjaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAmpktjTFLRa+xaeU9QU6tzv96R88TCY3HTvfm7VgadbzANBgkqhkiG9w0B
# AQEFAASCAYAf3ONNfWmTe1VEsBYCATWtBSChnkr6caeSK1mFxB7L5PMzoHo4WJ6r
# pv7RxLTzpoJT1RQt8SLj3dMBx/DK9DJf8mVUnEPU4MAGe20rs8MrqYSzwGrXKlKQ
# H5k1tWvh2aeFnTTKGdUGTV9vs3RjqOR8kw5KL2xjWYzZCZay/tubCMnspLMa/dj0
# 1vDpSujBx/j5DI2hs3wE9dgKlUKAXH912LRFxnxQwJZecWHROgVAseWh+g+C+Dgo
# GlKWCqB++bGR6QaXhQi8Wx69fost5H7VyK/1KNT9A2WiGagkyVLisLCziFQYwDB5
# vpzOZ7YVm90acLZFdYTLr2HPdHkNL8MwPi3yn2Nbj7cemkfpMWNVlbIPAO8527o+
# 3HQd6PfQu0nZIHCcwzFRia3RrIdxUAhMdQUBLBA3Eqb6KQu4Hr9/2sNUrVGPiF/z
# rq/mbk6Cw9ECwq2uJSiUNMdGtQ4Pq/S9Saz0G3dcM7K7vdYIk37v5ok5ZI+m3jhh
# ATTfASJMu/w=
# SIG # End signature block
