<#
.SYNOPSIS
    Detects whether the system drive is below the managed free-space threshold.

.VERSION
    1.6
#>
# Name: SmartM365-Disk-Space-Cleanup-Detection.ps1
# Version: 1.6
# Description: Detects whether the system drive is below the managed free-space threshold.

$ErrorActionPreference = "Stop"

$ScriptName = "SmartM365-Disk-Space-Cleanup-Detection"
$Version = "1.6"
$MinimumFreeSpaceGB = 50
$Windows10Only = $true

function ConvertTo-SingleLineValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s{2,}', ' '
    return $text.Trim()
}

function Write-IntuneResult {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Values)

    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Values.Keys) {
        $parts.Add(("{0}={1}" -f $key, (ConvertTo-SingleLineValue -Value $Values[$key])))
    }

    Write-Output ($parts -join " ")
}

function Get-SystemDriveFreeSpaceGB {
    try {
        $systemDrive = $env:SystemDrive

        if ([string]::IsNullOrWhiteSpace($systemDrive)) {
            return $null
        }

        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop

        if ($null -eq $drive -or $drive.DriveType -ne 3) {
            return $null
        }

        return [math]::Round(($drive.FreeSpace / 1GB), 2)
    }
    catch {
        return $null
    }
}

try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-IntuneResult -Values ([ordered]@{ Status = "NotApplicable"; Reason = "NotWindows10"; Script = $ScriptName; Version = $Version })
        exit 0
    }

    $freeSpaceGB = Get-SystemDriveFreeSpaceGB

    if ($null -eq $freeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; Reason = "FreeSpaceUnknown"; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 1
    }

    if ($freeSpaceGB -ge $MinimumFreeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "Ready"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 0
    }

    Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
    exit 1
}
catch {
    Write-IntuneResult -Values ([ordered]@{ Status = "Error"; Script = $ScriptName; Version = $Version; Message = $_.Exception.Message })
    exit 1
}
