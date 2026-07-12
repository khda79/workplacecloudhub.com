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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7Ys7AKOTiMtxW
# hGeWPSHy9zR6XM39H1N/p/lrYT+bB6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCvuUUxPaYx5P62Z/FC
# 0IZPMXY35PlbV8byGQOC/TQ3kTANBgkqhkiG9w0BAQEFAASCAYBTIuJe8e9nHxEY
# ux/WnZUlqJdIchOu32w5UWJy5a+XzB1j1HLoDwZ9iFsoUsKgfETQbQX1dNoCmXRX
# h8bGeZd3tTEdARk7rGrNF4DI8BMzFEuh1cJgA5sgr7OmlsVE35l6WaKSlk4lJICs
# 4tvSWJ1gdTmZdScmI6pFsV6C7qS1Y+QuPg2ZPj3/BKR2ofegTpA2A1F25j5hMXLb
# dP0Id9UzRbvXAbGAeURvRHaevUfPYh8hhh8uV+ArNY1izEgCLw8YMnm60S2ldJoV
# j1W+g3PyGwZjGL4kj0Z0TFYqT5MWDgL7GzRTG6jjWrw/GZzv+uye1p3cnh+DQytB
# hiiROvYaEnGMujyEf/pu98HRLFFUsvQixoLS9O7bbsecv6z6ycQVUyWoAkqPvYwi
# d6pw1+KXaB7OiISU90yXx1tKRCS5XaENUgWcE8r2dvnwI5Qzct6xjN4ylN9z95qN
# 8CPvySNvSX1sMsQZr8WhAJ/bNt5LhzHza9tFX+Pex/ggPiMjyZc=
# SIG # End signature block
