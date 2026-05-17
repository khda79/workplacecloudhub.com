<#
.SYNOPSIS
    Version: 1.0
    Enables or repairs BitLocker protection for BitLocker-OSDrive-Protection.
.DESCRIPTION
    Validates TPM readiness, adds a recovery password protector when needed, enables BitLocker on the operating system volume with used-space-only encryption, resumes protection when suspended, and attempts to back up the recovery key to Entra ID.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'
$Scenario = 'BitLocker-OSDrive-Protection'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log 'Remediation started.'

    $tpm = Get-Tpm -ErrorAction Stop
    if (-not $tpm.TpmPresent -or -not $tpm.TpmReady) { throw 'TPM is not present or not ready. BitLocker cannot be enabled safely by this remediation.' }

    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop
    if ($volume.ProtectionStatus -eq 'On' -and $volume.EncryptionPercentage -eq 100) {
        Write-Log 'BitLocker is already fully enabled and protected.'
        exit 0
    }

    $recoveryProtector = $volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
    if (-not $recoveryProtector) {
        $protector = Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop
        $recoveryProtector = $protector.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
        Write-Log 'Recovery password protector added.'
    }

    if ($volume.VolumeStatus -eq 'FullyDecrypted') {
        Enable-BitLocker -MountPoint $MountPoint -TpmProtector -UsedSpaceOnly -SkipHardwareTest -ErrorAction Stop
        Write-Log 'BitLocker enablement started with TPM protector and used-space-only encryption.'
    }
    elseif ($volume.ProtectionStatus -ne 'On') {
        Resume-BitLocker -MountPoint $MountPoint -ErrorAction SilentlyContinue
        Write-Log 'BitLocker protection resume requested.'
    }

    if ($recoveryProtector) {
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recoveryProtector.KeyProtectorId -ErrorAction Stop
            Write-Log 'Recovery key backup to Entra ID requested.'
        }
        catch {
            Write-Log "Recovery key backup to Entra ID failed or is not supported on this device: $($_.Exception.Message)"
        }
    }

    Write-Log 'Remediation completed.'
    exit 0
}
catch {
    Write-Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}

