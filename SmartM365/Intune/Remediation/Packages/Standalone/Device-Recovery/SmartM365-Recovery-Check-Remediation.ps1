# Name: SmartM365-Recovery-Check-Remediation.ps1
# Version: 1.0
$taskPath = "\Microsoft\Windows\Workplace Join"
$taskName = "Recovery-Check"

try {
    # Active la tâche
    Enable-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop

    # Exécute la tâche
    Start-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
}
catch {
    Write-Output "Erreur lors de la remédiation : $($_.Exception.Message)"
    exit 1
}

exit 0

## Test
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Detection.ps1
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Remediation.ps1

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7tLlpV35V+FG0
# 9kt8vGuR7x+jSrRMTuW+9B/eSSqqmaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCsezkwKTgkjCLXshC52djiioNbKkB1LEYPiQpxoV/DUTANBgkqhkiG9w0B
# AQEFAASCAYBUTpyFfodsEghUVVEZSw/hoiKHdPT9GYecm2fG0nHGBHIeq8XEooRF
# nNQZ6o4wb5RF4K6uvUL2K2MeUbafgf3sInyovBBgallz/9oHDL4pqXNjMo4rgzo/
# jmmMnB9qPQGrysaig4PcrRAPF8DgrPLQSzBwfMn3jHxKR/Xe99EfMfkwgn34x+B4
# IKewDj4PJtGsyy6hYmUqNCnBFr2qsHWZJoSwI/1bDaT69NdlcPe8h7VmbUGjF8VW
# d9bpWjMPSll/dF27LHpNgljgl6HZGHDbV4yuuLYjCGxZmVbx9HMmuYORAVsQV3EP
# NPT1ahx3ntGsOj3DwEQkJYE/fS+ZeMw+pwggZaJryaA/rZ/b9i2+XDuTvfwymDN4
# h2QIJ0hweUbUewYKEj7Re/Ulx6m517ZGWKBpRWBAn0KHVFSTrkIOoa85gDXaYIke
# qGn4W3h46TmKgJdjZ1nZuQmDtMwZXU1IfqZwFcF9mjtfrRFJ9wajYjCGZxe4dcGE
# gC0Pl8Ihv7E=
# SIG # End signature block
