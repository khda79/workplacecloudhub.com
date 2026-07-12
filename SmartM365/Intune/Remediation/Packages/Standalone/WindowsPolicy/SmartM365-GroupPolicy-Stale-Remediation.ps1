# Name: SmartM365-GroupPolicy-Stale-Remediation.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$Scenario = "GroupPolicy-Stale"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $process = Start-Process -FilePath (Join-Path $env:SystemRoot "System32\gpupdate.exe") -ArgumentList "/force" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
    Write-SmartM365Log "GpupdateExitCode=$($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        Write-SmartM365Log "Status=GpupdateFailed"
        exit 1
    }

    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop
    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateHours = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalHours

    if ($lastGPUpdateHours -le 24) {
        Write-SmartM365Log ("Status=Completed LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
        exit 0
    }

    Write-SmartM365Log ("Status=CompletedButStillStale LastGPUpdateHours={0:N1}" -f $lastGPUpdateHours)
    exit 1
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "Status=Error Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBExWjLUcV8oUTp
# 4o1sMpbjNHQmlnsfeM0JOkqF41B6EKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBzV+FFmPnDiucByqxjFFNsdulGTE3w9cHFuGMb2ol8LDANBgkqhkiG9w0B
# AQEFAASCAYAZ7IymGQljB6AZK5ix31LW6qiISJMSHJGV7P3eJ45ZldWW5YEWvzoa
# ip9KS0jjZkbdouOLgvZ5+elNonK3y0ZAGeRnRs80WlXJntzMBmfYRH85ZRwwTIUO
# Vq7nYnGaNi8HgFzKSeJrIm7eYWIF/5IRN/8172bmzFCieHLeX/pPvcMyeS8IPjSW
# PMUfftcpODAk8KX6+jfHtdXoWz/lNhhNkDYj+aZMV7hyWntM3oZOnNbI1wJey4y5
# 5edTe+Rkfa2ZqlZi1mu/uLeeo3UpxMaUtvKmjpp51xwjt5TDhUI2QuqGgFmhqxgK
# 5obPGZ2G3wGAkTT8vg+Vwa6mtFVD7XVJagSigFlRn8x2GZ8JM31Cl9wlueqV+VaD
# QbAvSjdmbWoBH1+SvME02Pm+08HYr4SzH044uncYfIc7eMiGZpudm714dEHALbmt
# 71C546pCU/+JIPF6cl2cCHuGHhf5GmqKQA+LAvqhI1xSdsYgh0t5IE8ZSahS68bq
# hQPDangR23s=
# SIG # End signature block
