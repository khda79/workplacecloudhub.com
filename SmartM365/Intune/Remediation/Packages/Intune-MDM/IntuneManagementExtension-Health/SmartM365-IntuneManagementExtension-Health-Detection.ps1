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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0WLChJnnIb9on
# aUrPJYVOWibOfFLxuBG5gjvOWbnVTqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCUch/TZK4GRLp3Zj1K
# Mj5ySfR1ViYkqEvXLeLfbm3G7TANBgkqhkiG9w0BAQEFAASCAYBvjkTbnJUPVR2r
# w3eii7Bp8KdLygIjyvmg6W97D0TOFFjVkIPas1CBd64mp/6QUFd0skrEM3WLcGXd
# ouCgS5bcckSkyMQ7V6lSIwoUPN8nwl/dx6h8exj2oi8+G20VfFwacSdrL5cm0eou
# Z/35V5wpmkm3DxHGLJM5TO20ZPBf6Eb2PJyKLnQzgyOxKj33ZED5xluLEV3vV7sj
# FV43FlbXlCrJMRgPIgmh8N/6zuU+B0GyCHolbp4NW4bLLS0cxO3HThoZu5rpQ/9O
# vKjP7yn4KF3di4XFA4WAYCG9IbRLUi/Gj1OKM6xjE75/8RljgGLT9zxbOo3juTUG
# QVD41fl+v3umeVcR7dc61hZ/H/S7nBaoVm8bnS21iOWXvp/q6A4tfm78jRFKIzCp
# GSs7plxanm9b3gxgHhIvmeST0P4ByEXrZUvkDbqM0+z6hdEBI473QUWwgr5ksdoT
# 2ZcloxbjOShztXj7LzdXB4EMyhqRxr/SmVZeM0SORC/FtbGS4BM=
# SIG # End signature block
