# Name: SmartM365-RepairWindowsUpdateandDISMissues-Remediation.ps1
# Version: 1.0

# Remediation Script for Windows Update and DISM
# Purpose: Repair Windows Update and DISM issues blocking Windows 11 upgrade

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Repair-DISM'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'RepairWindowsUpdateandDISMissues.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$hadError = $false

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [Repair-DISM] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

# Step 1: Restart Windows Update service
Write-SmartM365Log "Restarting Windows Update service..."
Try {
    Restart-Service -Name wuauserv -Force -ErrorAction Stop
    Write-SmartM365Log "Windows Update service restarted successfully."
} Catch {
    Write-SmartM365Log "Failed to restart Windows Update service: $($_.Exception.Message)"
    $hadError = $true
}

# Step 2: Run DISM to repair system image
Write-SmartM365Log "Running DISM /Online /Cleanup-Image /RestoreHealth..."
Try {
    $dismResult = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "DISM completed with exit code: $($dismResult.ExitCode)"
    if ($dismResult.ExitCode -notin @(0, 3010)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "DISM failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 3: Run SFC to repair system files
Write-SmartM365Log "Running SFC /scannow..."
Try {
    $sfcResult = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "SFC completed with exit code: $($sfcResult.ExitCode)"
    if ($sfcResult.ExitCode -notin @(0, 1)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "SFC failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 4: Clear Windows Update cache
Write-SmartM365Log "Clearing Windows Update cache..."
Try {
    net stop wuauserv
    Remove-Item -Path "$env:windir\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    net start wuauserv
    Write-SmartM365Log "Windows Update cache cleared."
} Catch {
    Write-SmartM365Log "Failed to clear Windows Update cache: $($_.Exception.Message)"
    $hadError = $true
}

# Step 5: Log completion
if ($hadError) {
    Write-SmartM365Log "Status=CompletedWithErrors"
    exit 1
}

Write-SmartM365Log "Status=Completed"
exit 0

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOjr3yD9hpnwA/
# ilYSlvg7z1RYXt7EEltSbeo65fYodaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCD1jQ9qTiqld5BxhvM0AScsAKLMeiT8cl1UqwxAs97FQzANBgkqhkiG9w0B
# AQEFAASCAYCuLPtWK/b9wK89aoKqJXJvJ2k4CXWDripVTf4zcT+upUCn0nbxi4wG
# grVYtE0XPXQuI/gA91yBIdyC9YZCFleJmBaIVnBcTyiTRzSazZSkht9aDYLLCgCS
# ABmD9vzfHXqtG7yzzxMxaNHUSncH2aahF/6UYgLYZ3gzDVkmAlMpyDavduHnKQq2
# dMSXrIlyiVFIdQoaP15XOIxN3QwL8jjJhc8YxkCoyTKuPTxplCUhfTcR2IBw5iyr
# BGCU8RgsbtwmVbPLFdMCyAFI7iWSjBce3Lu/y46lF1GcP0aeDyXx1PAIiJH0cMdX
# E/ff0qOHkZFwFO4WHYGFwix3Z+/XtgpTFfTquPd63q0V/WZH+r/4OxDt+fCerBse
# w4CqR+AKL/HpKtvLuOnQ9epAvj+YQzPBenhG+GEnEfFTsA++6kjEtp+RBeT2IVyG
# MT+EoEp+6Zdv5p5lUazhiwQ7KxnP0FfAUvqn4L3+uf7zDtoV0jr6JCPKARxBuE7y
# MGoB6aoUsTw=
# SIG # End signature block
