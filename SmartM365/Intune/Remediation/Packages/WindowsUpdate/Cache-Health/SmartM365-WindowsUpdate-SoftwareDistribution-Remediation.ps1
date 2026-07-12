# Name: SmartM365-WindowsUpdate-SoftwareDistribution-Remediation.ps1
# Version: 1.0
# Description: Safely resets the Windows Update download cache and triggers a new scan without forcing a reboot

$ErrorActionPreference = "Stop"

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Cache-Health'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Remediate-WindowsUpdate-SoftwareDistribution.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStopped=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyStopped=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStopFailed=$Name Message=$($_.Exception.Message)"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStarted=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyRunning=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStartFailed=$Name Message=$($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Windows Update remediation started ====="

    $softwareDistributionPath = "C:\Windows\SoftwareDistribution"
    $downloadPath = Join-Path -Path $softwareDistributionPath -ChildPath "Download"

    # Stop Windows Update related services
    foreach ($serviceName in @("wuauserv", "bits", "usosvc")) {
        Invoke-ServiceStopSafe -Name $serviceName
    }

    Start-Sleep -Seconds 3

    # Clean only Download content instead of deleting the entire SoftwareDistribution folder
    if (Test-Path -Path $downloadPath) {
        Write-SmartM365Log "Cleanup=Start Path=$downloadPath"

        try {
            Get-ChildItem -Path $downloadPath -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-SmartM365Log "CleanupItemSkipped=$($_.FullName)"
                }
            }

            Write-SmartM365Log "Cleanup=Completed Path=$downloadPath"
        }
        catch {
            Write-SmartM365Log "Cleanup=Partial Path=$downloadPath Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "Cleanup=DownloadFolderNotFound"
    }

    # Restart services
    foreach ($serviceName in @("bits", "usosvc", "wuauserv")) {
        Invoke-ServiceStartSafe -Name $serviceName
    }

    Start-Sleep -Seconds 5

    # Trigger Windows Update scan
    $usoClientPath = Join-Path -Path $env:windir -ChildPath "System32\UsoClient.exe"

    if (Test-Path -Path $usoClientPath) {
        try {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-SmartM365Log "WindowsUpdateScan=Triggered"
        }
        catch {
            Write-SmartM365Log "WindowsUpdateScan=Failed Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "UsoClientNotFound"
    }

    Write-SmartM365Log "===== Windows Update remediation completed ====="
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGJTfsr85nkJaQ
# o7QQLKJkHSQmVghYemetGUvaIuMIBqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBfEmrWKvgvQZCM45rs
# eb7V+qS3xgrDEilIo5ESy838vDANBgkqhkiG9w0BAQEFAASCAYAh3PsUYlQ+yaRr
# hcaz7r87yrrYJL4Sp1RGnMJb8YPoeHWCUHXGIKHyBlv+2tQILVicl1TycW2cD7E4
# FdgYrZf11Kt6STK+5VL5pxUrnxsRHzyWKx4kT0eHy3zU/AI2M5w5tyJ5mci1Jolg
# MkR0JizPKlpJ05+xgjjr8WvxkUAmLpn6J3MhMWf5nftVox9DUOuRAIXF8yGgagGx
# O0aDqXwu05BM++NMKVS/zjdFdAYJ4F1h8qUwtrTXrWCW+i8nzpKt/v7V2vaOkUGY
# zALAHsc/nhVrSnuZiq4dvAjqq1TMTtqmkcaYy7caIYdgexN0BPkKXTbe+lVsqDOB
# 99nMh3eupZNpXnhVXdIp3gl4fPEYyNYga3W+miV+WDUtfCjx5Sk1t4SD87PLstnN
# qL1mrM9sd7aZ2WOB40GyK+oF0onOOFL21APn8TqRZv9VTMPRViGZW0HO0j13BLJy
# xoNQ/HfdCVaix025e8sbdngti4818CDhu8GWVS/ljQLVc2JGSlw=
# SIG # End signature block
