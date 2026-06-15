<#
.SYNOPSIS
    Version: 1.0
    Detects BitLocker protection health on the operating system drive.
.DESCRIPTION
    Validates that the operating system drive exists, BitLocker protection is enabled,
    encryption is complete, and a recovery password protector is present.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'

try {
    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop

    if ($null -eq $volume) {
        Write-Output "Status=BitLockerVolumeNotFound"
        exit 1
    }

    $issues = New-Object System.Collections.Generic.List[string]

    if ($volume.ProtectionStatus -ne 'On') {
        $issues.Add("ProtectionStatus=$($volume.ProtectionStatus)")
    }

    if ($volume.EncryptionPercentage -lt 100) {
        $issues.Add("EncryptionPercentage=$($volume.EncryptionPercentage)")
    }

    $recoveryProtector = $volume.KeyProtector |
        Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } |
        Select-Object -First 1

    if (-not $recoveryProtector) {
        $issues.Add('RecoveryPasswordProtector=Missing')
    }

    if ($issues.Count -gt 0) {
        Write-Output "Status=BitLockerNonCompliant"
        Write-Output ("Issues=" + ($issues -join '; '))
        exit 1
    }

    Write-Output "Status=BitLockerProtected"
    Write-Output "MountPoint=$MountPoint"
    Write-Output "EncryptionPercentage=$($volume.EncryptionPercentage)"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
