# Name: SmartM365-RepairWindowsUpdateandDISMissues-Remediation.ps1
# Version: 1.0

# Remediation Script for Windows Update and DISM
# Purpose: Repair Windows Update and DISM issues blocking Windows 11 upgrade

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Repair-DISM'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'RepairWindowsUpdateandDISMissues.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$hadError = $false

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [Repair-DISM] {1}" -f (Get-Date -Format "s"), $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

# Step 1: Restart Windows Update service
Write-SmartM365Log "Restarting Windows Update service..."
Try {
    Restart-Service -Name wuauserv -Force -ErrorAction Stop
    Write-SmartM365Log "Windows Update service restarted successfully."
} Catch {
    Write-SmartM365Log "Failed to restart Windows Update service: $($_.Exception.Message)"
    $hadError = $true
}

# Step 2: Run DISM to repair system image
Write-SmartM365Log "Running DISM /Online /Cleanup-Image /RestoreHealth..."
Try {
    $dismResult = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "DISM completed with exit code: $($dismResult.ExitCode)"
    if ($dismResult.ExitCode -notin @(0, 3010)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "DISM failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 3: Run SFC to repair system files
Write-SmartM365Log "Running SFC /scannow..."
Try {
    $sfcResult = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-SmartM365Log "SFC completed with exit code: $($sfcResult.ExitCode)"
    if ($sfcResult.ExitCode -notin @(0, 1)) {
        $hadError = $true
    }
} Catch {
    Write-SmartM365Log "SFC failed: $($_.Exception.Message)"
    $hadError = $true
}

# Step 4: Clear Windows Update cache
Write-SmartM365Log "Clearing Windows Update cache..."
Try {
    net stop wuauserv
    Remove-Item -Path "$env:windir\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    net start wuauserv
    Write-SmartM365Log "Windows Update cache cleared."
} Catch {
    Write-SmartM365Log "Failed to clear Windows Update cache: $($_.Exception.Message)"
    $hadError = $true
}

# Step 5: Log completion
if ($hadError) {
    Write-SmartM365Log "Status=CompletedWithErrors"
    exit 1
}

Write-SmartM365Log "Status=Completed"
exit 0
