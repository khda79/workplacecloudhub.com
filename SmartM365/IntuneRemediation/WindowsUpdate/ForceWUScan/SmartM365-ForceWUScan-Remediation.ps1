# Name: SmartM365-ForceWUScan-Remediation.ps1
# Version: 1.0
[CmdletBinding(SupportsShouldProcess=$true)]
param()

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\ForceWUScan'
$logFile = Join-Path -Path $LogRoot -ChildPath 'Remediate-ForceWUScan.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
Start-Transcript -Path $logFile -Append | Out-Null

try {
    Write-Output "Starting remediation: Force Windows Update scan."

    # 1) Ensure Windows Update service is running
    $svc = Get-Service -Name wuauserv -ErrorAction Stop
    if ($svc.Status -ne 'Running') {
        Write-Output "wuauserv is $($svc.Status). Starting..."
        Start-Service -Name wuauserv
    }

    $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'

    # 2) Refresh settings (optional)
    Write-Output "Refreshing Windows Update settings..."
    & $uso RefreshSettings
    # 3) Trigger the scan (silent)
    Write-Output "Triggering Windows Update scan (USOClient StartInteractiveScan)..."
    & $uso StartInteractiveScan
    # --- OPTIONAL: force full workflow (uncomment if needed — may bypass ring cadence) ---
    # & $uso StartDownload
    # & $uso StartInstall
    # Single-step alternative:
    # & $uso ScanInstallWait
    # ---------------------------------------------------------------------

    # 4) Give it time to register in the event log
    Write-Output "Sleeping 120s to let scan events register..."
    Start-Sleep -Seconds 120
    # 5) Echo the last WU event for reporting
    $last = Get-WinEvent -LogName 'Microsoft-Windows-WindowsUpdateClient/Operational' -MaxEvents 1 -ErrorAction SilentlyContinue
    if ($last) {
        Write-Output ("Last WU event: {0:yyyy-MM-dd HH:mm:ss} (ID {1})" -f $last.TimeCreated, $last.Id)
    } else {
        Write-Output "No events found after scan trigger."
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