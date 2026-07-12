# Name: SmartM365-IntuneManagementExtension-Health-Detection.ps1
# Version: 1.0
# Description: Verifies whether Intune Management Extension (IME) is installed and healthy

$ErrorActionPreference = "Stop"

try {
    # Verify service
    $serviceName = "IntuneManagementExtension"
    $service = Get-Service -Name $serviceName -ErrorAction Stop

    if ($service.Status -ne "Running") {
        Write-Output "Intune Management Extension service is stopped"
        exit 1
    }

    # Verify process
    $process = Get-Process -Name IntuneManagementExtension -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        Write-Output "IME process is not running"
        exit 1
    }

    # Verify installation path
    $installPath = "C:\Program Files (x86)\Microsoft Intune Management Extension"
    if (-not (Test-Path $installPath)) {
        Write-Output "IME installation folder is missing"
        exit 1
    }

    # Verify logs folder
    $logPath = "C:\ProgramData\SmartM365\IntuneRemediation\Logs"
    if (-not (Test-Path $logPath)) {
        Write-Output "IME logs folder is missing"
        exit 1
    }

    # Verify main log file
    $agentLog = Join-Path $logPath "AgentExecutor.log"

    if (-not (Test-Path $agentLog)) {
        Write-Output "AgentExecutor.log is missing"
        exit 1
    }

    # Verify recent activity
    $lastWrite = (Get-Item $agentLog).LastWriteTime
    $delta = (Get-Date) - $lastWrite

    if ($delta.TotalHours -gt 24) {
        Write-Output "IME appears inactive (no recent activity detected)"
        exit 1
    }

    # Verify execution activity in logs
    $hasExecution = Select-String -Path $agentLog -Pattern "Executing" -Quiet

    if (-not $hasExecution) {
        Write-Output "No IME execution activity detected in logs"
        exit 1
    }

    Write-Output "Intune Management Extension is healthy"
    exit 0
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0WLChJnnIb9on
# aUrPJYVOWibOfFLxuBG5gjvOWbnVTqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCUch/TZK4GRLp3Zj1KMj5ySfR1ViYkqEvXLeLfbm3G7TANBgkqhkiG9w0B
# AQEFAASCAYC8dkdxRGriyFilSzDw7vdUPkKjmH7BKtAvsvOYdXF1USomZt1fOF+F
# 7YHV0kSGoI3pCBr9nUkrSrkbxMIb28WNbk2zyuLN7rTkmVztHfiP82+KeAONY5Db
# /TttamczKAz1pJvOMIn159tlSDpzmzDGFWQ2ObWGPLOQla7cQ4XtwUysXrRdfBYl
# YidloXL2SimGeaLgxrJ/dhbL4Gi4+hNixpQJyQzvdUwa96yz//ZuoOgu3JAYYYY4
# Mk2hvp3z8MvW0Wt/okWUvxqA4/GEoAX6NB88HOAvBqR4C4sXjI7u2wTf7NOX8Yhf
# GtzJs8o7QkP5zSyKATNIEqSaf+czGqsixRHuFyoH4YpErescpg0feI791eqsBdxk
# orpTqqtIyk7GKwBFNktOGdjzhSHrjRh2fhKsn2Cu+guJmYmmzj024S35Oo8E9zUn
# w72D44pDMxvHuT9+wsm+6reRSVo2H8SHvXrZfLVL7q9R+Km8l6kSZ1haYLBkD4jC
# LlMmpBNiEXM=
# SIG # End signature block
