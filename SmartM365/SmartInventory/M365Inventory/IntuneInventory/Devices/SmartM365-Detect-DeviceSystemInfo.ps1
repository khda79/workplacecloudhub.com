<#
.SYNOPSIS
    Detects device system information for the SmartM365 Intune inventory.

.DESCRIPTION
    This script is intended to be deployed as an Intune Platform Script on Windows devices.
    It collects Secure Boot status, BIOS information, firmware type, and the operating system
    last boot time. Results are written to stdout in a pipe-delimited format for retrieval
    through Microsoft Graph deviceManagementScripts/deviceRunStates.

    Output format:
    SecureBoot:<value>|BIOSVersion:<value>|BIOSDate:<value>|FirmwareType:<value>|LastBootUpTime:<value>

.NOTES
    Version: 1.2.0
    Author: https://github.com/khda79/workplacecloudhub.com
    Deploy via: Intune > Devices > Scripts > Platform scripts (Windows)
    Run as: System
    Run in 64-bit PowerShell: Yes
#>

$outputParts = @()

# Secure Boot
try {
    $secureBootEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    $outputParts += if ($secureBootEnabled) { 'SecureBoot:Enabled' } else { 'SecureBoot:Disabled' }
}
catch [System.PlatformNotSupportedException] {
    $outputParts += 'SecureBoot:NotSupported'
}
catch {
    $outputParts += 'SecureBoot:Error'
}

# SMBIOS BIOS version and release date
try {
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $biosVersion = if ($bios.SMBIOSBIOSVersion) { $bios.SMBIOSBIOSVersion.Trim() } else { 'Unknown' }
    $biosDate = if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
    $outputParts += "BIOSVersion:$biosVersion"
    $outputParts += "BIOSDate:$biosDate"
}
catch {
    $outputParts += 'BIOSVersion:Error'
    $outputParts += 'BIOSDate:Error'
}

# Firmware type
try {
    $computerInfo = Get-ComputerInfo -Property BiosFirmwareType -ErrorAction Stop
    $firmwareType = if ($computerInfo.BiosFirmwareType) { $computerInfo.BiosFirmwareType.ToString() } else { 'Unknown' }
    $outputParts += "FirmwareType:$firmwareType"
}
catch {
    $outputParts += 'FirmwareType:Error'
}

# Operating system last boot time, expressed in the device local time
try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $lastBootUpTime = if ($operatingSystem.LastBootUpTime) {
        ([datetime]$operatingSystem.LastBootUpTime).ToString('yyyy-MM-dd HH:mm:ss')
    }
    else {
        'Unknown'
    }
    $outputParts += "LastBootUpTime:$lastBootUpTime"
}
catch {
    $outputParts += 'LastBootUpTime:Error'
}

Write-Output ($outputParts -join '|')
exit 0
