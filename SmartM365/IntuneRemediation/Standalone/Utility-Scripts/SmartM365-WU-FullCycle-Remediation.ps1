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
}