# Name: SmartM365-RepairWindowsUpdateandDISMissues-Remediation.ps1
# Version: 1.0

# Remediation Script for Windows Update and DISM
# Purpose: Repair Windows Update and DISM issues blocking Windows 11 upgrade

# Enable transcript for logging
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Repair-DISM'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'RepairWindowsUpdateandDISMissues.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
Start-Transcript -Path $LogPath -Append

# Step 1: Restart Windows Update service
Write-Output "Restarting Windows Update service..."
Try {
    Restart-Service -Name wuauserv -Force -ErrorAction Stop
    Write-Output "Windows Update service restarted successfully."
} Catch {
    Write-Output "Failed to restart Windows Update service: $_"
}

# Step 2: Run DISM to repair system image
Write-Output "Running DISM /Online /Cleanup-Image /RestoreHealth..."
Try {
    $dismResult = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-Output "DISM completed with exit code: $($dismResult.ExitCode)"
} Catch {
    Write-Output "DISM failed: $_"
}

# Step 3: Run SFC to repair system files
Write-Output "Running SFC /scannow..."
Try {
    $sfcResult = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-Output "SFC completed with exit code: $($sfcResult.ExitCode)"
} Catch {
    Write-Output "SFC failed: $_"
}

# Step 4: Clear Windows Update cache
Write-Output "Clearing Windows Update cache..."
Try {
    net stop wuauserv
    Remove-Item -Path "$env:windir\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    net start wuauserv
    Write-Output "Windows Update cache cleared."
} Catch {
    Write-Output "Failed to clear Windows Update cache: $_"
}

# Step 5: Log completion
Write-Output "Remediation completed. Please retry Windows 11 upgrade."

# End transcript
Stop-Transcript
