# Name: SmartM365-WU-FullCycle-Remediation.ps1
# Version: 1.0
# SmartM365-WU-FullCycle-Remediation.ps1
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$ForceReboot,        # If specified, reboot automatically when required
    [int]$RebootTimeoutSeconds = 300  # Grace period before reboot
)

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Utility-Scripts'
$logFile = Join-Path -Path $LogRoot -ChildPath 'Remediate-WU-FullCycle.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
Start-Transcript -Path $logFile -Append | Out-Null

function Update-TimestampedTranscript {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $updatedLines = [System.IO.File]::ReadAllLines($Path) | ForEach-Object {
        if ($_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b') {
            $_
        }
        elseif ($_ -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
            "{0} {1}" -f $Matches[1], $Matches[2]
        }
        else {
            "{0} {1}" -f $timestamp, $_
        }
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, $utf8NoBom)
}

function Test-PendingReboot {
    try {
        $rebootWU = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        $rebootCBS = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        return ($rebootWU -or $rebootCBS)
    } catch { return $false }
}

try {
    Write-Output "Starting remediation: Full Windows Update cycle (scan, download, install)."

    # 1) Ensure core services are running
    foreach ($svcName in 'wuauserv','bits','cryptsvc') {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            if ($svc.Status -ne 'Running') {
                Write-Output "$svcName is $($svc.Status). Starting..."
                Start-Service -Name $svcName
            }
        } catch {
            Write-Warning "Service $svcName not available or failed to start: $_"
        }
    }

    $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
    if (-not (Test-Path $uso)) {
        throw "UsoClient.exe not found at $uso"
    }

    # 2) Refresh and run full cycle
    Write-Output "Refreshing WU settings..."
    & $uso RefreshSettings

    Write-Output "Running ScanInstallWait (scan + download + install + wait)..."
    & $uso ScanInstallWait

    # 3) Wait a bit for logs/installation bookkeeping
    Write-Output "Sleeping 120s to allow events to register..."
    Start-Sleep -Seconds 120

    # 4) Reporting: show last WU event
    $last = Get-WinEvent -LogName 'Microsoft-Windows-WindowsUpdateClient/Operational' -MaxEvents 1 -ErrorAction SilentlyContinue
    if ($last) {
        Write-Output ("Last WU event: {0:yyyy-MM-dd HH:mm:ss} (ID {1})" -f $last.TimeCreated, $last.Id)
    } else {
        Write-Output "No WU events found after ScanInstallWait."
    }

    # 5) Reboot if required (optional)
    if (Test-PendingReboot) {
        if ($ForceReboot.IsPresent) {
            Write-Output "Reboot is required. Restarting device in $RebootTimeoutSeconds seconds..."
            shutdown.exe /r /t $RebootTimeoutSeconds /c "Updates installed by Intune remediation; device will restart." /d p:2:4
        } else {
            Write-Output "Reboot required but ForceReboot not specified. Device will not auto-restart."
        }
    } else {
        Write-Output "No reboot required."
    }

    Write-Output "Remediation completed."
    exit 0
}
catch {
    Write-Error "Remediation failed: $_"
    exit 1
}
finally {
    Stop-Transcript | Out-Null
    Update-TimestampedTranscript -Path $logFile
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7Ys7AKOTiMtxW
# hGeWPSHy9zR6XM39H1N/p/lrYT+bB6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCvuUUxPaYx5P62Z/FC0IZPMXY35PlbV8byGQOC/TQ3kTANBgkqhkiG9w0B
# AQEFAASCAYBbntwj+oy5BjKFS7wtvhyM79SYeBfmiayXoc6ghlQPME2OLMTtagSF
# tG/PSK2r/0slWw+OD24IFUjdMMCzqLOQRCLTYLondG3FAv5N21NXu2RNexmYBbBU
# jCJ6HJoUxsD/Nk6s45rwhHFdJ+SaFmHW0PR4pufZBtx9dRUEUyweXZnqozs+3add
# ILPojbO4ERdRyFpPf/fbO8H7Y1qD0ZrYQze3KQAu8CVrrwnVv6zQl9BDYueA/oh+
# 7hYwE600jS3oPdBuvu42tyT6yy3bu18YYp/FW135X49lYWLr+Fhk8On3gpca6NaN
# T+MM7HUGLznQmN7lsGBgFzM931p092vk2rdR97psbAnipHRNZLfoSZIfDizta2PO
# TZbcygvOWUM9wOIt2fLbG2Su6lt/K8Ml3cgLXPKpNbh9J0AomYNnNJa9R30Wp//y
# 0LbYFZ4oFScQpLP43gfw9wjs3UU6pTMo20PK0a7qa883za3yM8cTz/ffBNg0Z2Jj
# 5NCiyXjmXak=
# SIG # End signature block
