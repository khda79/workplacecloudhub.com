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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCct6j0GIeKwfLN
# 1GjwoCd6+9oiSmeTW8lvEVANBWw6MaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAa0r8yCd/eW/mTH22F
# 3BFWJjjDO4IOK6iLhoDarua8DTANBgkqhkiG9w0BAQEFAASCAYARvw7hqozNA1gG
# 9+H42+a9uPT5TKisBYdAlF2bmfi6fQZSxoOrq4oMDA6vdVtm7zDBRWfg03gIpCjM
# WmYgPpnV4f8M3+8dS8C7Ta64Gh7EtGx9u177q9CGDS+MeEbwK3D0DS1DgAwBBIcu
# 5BhCNprs5vfkfoT+0H2Siavb9ywyUm0503b47yVSs/mvJzbZk+WEVm2zL2GY3MDo
# PptPhdCvsL0WiuBAgYTEQpLNL0+2/qNLERblOAelU/SIHvRz+o3kdbmHmYWTuvCi
# 1rRLrjOSN5oLXgiPun0Z2egs1Nb5ZM9k2mZ6U93f4QMtZldGTC1rFlSw7kvVZkfK
# QAMXzzZQW+O2z/IQfUcPaJptUoqV2z4yg23y3/aaj+7xFeDdMSwTlBgfkaP26ho4
# Il5Jmc0batJJenAA3gjFQwzkIbsjzUFe3U9+VBXZci1KBvr2kz9Co6uTJ/4vQFuE
# Yfu5jQ/ElG4XvfuGIWc4ZHiu32LRGowWj4lZA2GQnumuIu2KhyQ=
# SIG # End signature block
