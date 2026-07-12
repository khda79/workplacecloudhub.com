<#
  Name: SmartM365-Force-W11-Upgrade-Now-Remediation.ps1

    Version: 1.0
  Purpose: Force an immediate Windows Update scan/download/install (for assigned updates such as Windows 11 feature upgrade).
  Use case: Manual "Run Script" from Intune per device.
  Run as: SYSTEM (administrator)
#>

$ErrorActionPreference = 'Stop'

function Invoke-AssignedUpdatesInstall {
    $session  = New-Object -ComObject 'Microsoft.Update.Session'
    $searcher = $session.CreateUpdateSearcher()
    $criteria = "IsInstalled=0 and IsHidden=0 and IsAssigned=1"
    $result   = $searcher.Search($criteria)

    if ($result.Updates.Count -eq 0) {
        Write-Output "No applicable assigned updates found via COM API."
        return "NoUpdates"
    }

    $toInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    foreach ($u in $result.Updates) { [void]$toInstall.Add($u) }

    Write-Output "Downloading updates..."
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $toInstall
    [void]$downloader.Download()

    Write-Output "Installing updates..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $toInstall
    $installResult = $installer.Install()

    if ($installResult.RebootRequired) {
        Write-Output "Reboot required; scheduling reboot in 60 seconds."
        shutdown.exe /r /t 60 /c "Rebooting to complete Windows update"
        return "RebootScheduled"
    }
    else {
        Write-Output "Installation completed. No reboot required."
        return "Installed"
    }
}

try {
    $result = Invoke-AssignedUpdatesInstall
    if ($result -eq "NoUpdates") {
        Write-Output "Fallback to USOClient triggers..."
        $usoClient = Join-Path $env:SystemRoot "System32\UsoClient.exe"
        if (-not (Test-Path -LiteralPath $usoClient -PathType Leaf)) {
            throw "UsoClient.exe was not found."
        }
        Start-Process $usoClient -ArgumentList "StartScan" -WindowStyle Hidden -Wait
        Start-Process $usoClient -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
        Start-Process $usoClient -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
        exit 0
    }
    Write-Output "Status=$result"
    exit 0
}
catch {
    Write-Output ("Status=Error Message=" + $_.Exception.Message)
    try {
        $usoClient = Join-Path $env:SystemRoot "System32\UsoClient.exe"
        if (Test-Path -LiteralPath $usoClient -PathType Leaf) {
            Write-Output "Attempting USOClient fallback after COM API failure."
            Start-Process $usoClient -ArgumentList "StartScan" -WindowStyle Hidden -Wait
            Start-Process $usoClient -ArgumentList "StartDownload" -WindowStyle Hidden -Wait
            Start-Process $usoClient -ArgumentList "StartInstall" -WindowStyle Hidden -Wait
            Write-Output "Status=FallbackTriggered"
            exit 0
        }
    }
    catch {
        Write-Output ("Fallback failed: " + $_.Exception.Message)
    }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCct6j0GIeKwfLN
# 1GjwoCd6+9oiSmeTW8lvEVANBWw6MaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAa0r8yCd/eW/mTH22F3BFWJjjDO4IOK6iLhoDarua8DTANBgkqhkiG9w0B
# AQEFAASCAYCdPkgKYbYLb/9agdxs4Oh1+FVCxWbELGz/iqgouzuYVIWVfe1eUemO
# MLPI1QYuqtCoEFB5LifT6m7VMGSIzoG4daLYkkYRfYHg6dN+VH9Bq7D+ZAXMmTJk
# gOOLPLX+4WRvMG+utxdCd2s6LrJXe8rZNJ7tMEZhbT9cjQtl7e39UBJGvpy/idC1
# OwIGSu1SLUfRdRGP2h9EyGu0cdyaRMiUaLIhdYj1GxHd0POTPJyuFlVJYSbi8/Br
# BTrmyv39AAIsawtOUVHAhR8y8ALjJO4bGxPDO2FygtXnPR0v0yGRkJSJMlXQuRr+
# JY6Xm/8Yd+kl2lLf4QAoQ3bsl6BSY+WiqOSOhJP9cpDIPXjY5RfPTC72D3oyRlmW
# +Sktk4a2JcbJbIwwUq2bZye6luZNhUgqXQOV1pBMiaI08RjplMsBxSM/bMantwKg
# yGXaYTqUDGgMy+O0njp9PaF3mRNyZbd3g+LhHU83iOvYOZ6ogz2/2rec20j/NNgJ
# hvdly212hHk=
# SIG # End signature block
