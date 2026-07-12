# Name: SmartM365-Force-AllowOSUpgrade-Detection.ps1
# Version: 1.0
# Detection script - Check if AllowOSUpgrade is set correctly
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade"
$regName = "AllowOSUpgrade"
$expectedValue = 1

if (Test-Path $regPath) {
    try {
        $actualValue = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop | Select-Object -ExpandProperty $regName
        if ($actualValue -eq $expectedValue) {
            Write-Output "Compliant"
            exit 0
        } else {
            Write-Output "Non-compliant: Value is $actualValue"
            exit 1
        }
    } catch {
        Write-Output "Non-compliant: Value not found"
        exit 1
    }
} else {
    Write-Output "Non-compliant: Registry path not found"
    exit 1
}
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIbb5Hix77y1vf
# XJli2KWkzR3cDhieZDFW/K54U9dkqqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDXU4IyTbBtgU9H+USifzsFzFybiPqZqmGV7gEdRfzCgzANBgkqhkiG9w0B
# AQEFAASCAYCV18C6YzPv+YGxo+NJDGhHDZX3tZ4qzEQomNYI0XDUFZOdz7aBDVcZ
# SDSVNI7bFAgmo0oqPTtR0cviHYlF2ye0jMYghTcCYmJCUbiwgUs1hIHCB1WFFeEk
# 97/nyXrCnsEtN2v7hAd2sLkLYEqPb7/cDdGqheJszCX/CVISq57dLrs0YWNulR3f
# a6odfQk1yOqlQuH1f/guePkyDUvWIedQebH3q83c29IPg2tKk8L0Iw3eYuqgq9em
# cSHEI65PPmocTvfS1Lgm22miuNZs0Ewr9qTZIir7WcX8CLfkGzjPHOXfeFyDt03P
# gv8DPACGaTVJP/Y5DcyezC4ybEnv/yUomvBhgSKz5IkmODoO26J8oDBeEBJbyp6w
# 4qEkZmF+0L5vKyRVq/CLB1dtx4QwPLRd4DHMKWnSqQpMO/o6RaDxGQpEFS0lLRln
# LAAwg2zjDfccFW3NoM0ww7qHrhEUUPQWQpTRICpjcDatb15csLBruGuxBooo/+m3
# +Xn4pe+pGc8=
# SIG # End signature block
