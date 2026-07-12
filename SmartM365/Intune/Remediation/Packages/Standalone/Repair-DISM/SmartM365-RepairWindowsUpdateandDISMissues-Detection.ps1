# Name: SmartM365-RepairWindowsUpdateandDISMissues-Detection.ps1
# Version: 1.0
# Detection Script for Windows Update and DISM issues

$RemediationNeeded = $false

# Check Windows Update service
$WUService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($WUService.Status -ne 'Running') {
    Write-Output "Windows Update service is not running."
    $RemediationNeeded = $true
}

# Check DISM health
$DismLog = "$env:windir\Logs\DISM\dism.log"
$DismError = $false
if (Test-Path $DismLog) {
    $DismError = Select-String -Path $DismLog -Pattern "error|failed" -SimpleMatch
    if ($DismError) {
        Write-Output "DISM log contains errors."
        $RemediationNeeded = $true
    }
}

# Check CBS log for update errors
$CBSLog = "$env:windir\Logs\CBS\cbs.log"
$CBSError = $false
if (Test-Path $CBSLog) {
    $CBSError = Select-String -Path $CBSLog -Pattern "error|failed" -SimpleMatch
    if ($CBSError) {
        Write-Output "CBS log contains errors."
        $RemediationNeeded = $true
    }
}

# Check Windows Update cache
$WUCache = "$env:windir\SoftwareDistribution\Download"
if (Test-Path $WUCache) {
    $CacheFiles = Get-ChildItem -Path $WUCache -Recurse -ErrorAction SilentlyContinue
    if ($CacheFiles.Count -gt 1000) {
        Write-Output "Windows Update cache contains excessive files."
        $RemediationNeeded = $true
    }
}

# Output for Intune Remediation
if ($RemediationNeeded) {
    Write-Output "Remediation required."
    exit 1
} else {
    Write-Output "No remediation required."
    exit 0
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHjyKtzCKIFy1s
# GbKe7HmrriN2dq6XAVkME3litivK06CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBMKnfs0B3+z/Z+C4c1QgNnlhsJ1zUmn9ZP5tGHARCJXDANBgkqhkiG9w0B
# AQEFAASCAYC7KFDX8n0x5QtHlXo0pU32rDiLvai98lEqe8ZQci6m/v8IQbsJkuXO
# SA4//HgIoADKD0g6zMeSXwpIVkNy3LCn4b3lNg75fdazppXdVARbXyC/Yg62pUnr
# N9SlYJ+QqFOZQs91VtiltGY8dbpjcuTGLuHCofzXX5DEKbjlE3lmaWa5/lvIn2AC
# MOiHe3KQNANrHoVRTc50ReINacP8imac8rs/Q9461wIEuRnBUalZtJahF4Vq37Oc
# PZbFK4kZEMiH/TPw1b1H2szVDAvcIpK+spnyS0YfrVEaD0O2LkrEiwuvEWfm2UaJ
# FCK7IXSsW65Pf3Mgq7SsF+kCSqG34XAazOC3fnIVIpsd+p10TuMzOFBDz9eK11V6
# ZUz8O95nqvCfvhyl08PJig4rs0FLI+x/MRlcA7viA5OaqxMRvisNgEEwp2O4XAE5
# 6To/T/NiNS+OUPn4e4LGnqUZxOs7w+jqF5WmKJMaGKp0W0uJoKW8I13fKBZfy5em
# PrRu9UDuF84=
# SIG # End signature block
