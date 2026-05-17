# Name: SmartM365-Upgrade-Storage-Readiness-Detection.ps1
# Version: 1.0
# Description: Detects whether the system drive has enough free space for Windows upgrade readiness and identifies cleanup candidates.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-Upgrade-Storage-Readiness"
$Version = "1.0"
$MinimumFreeSpaceGB = 30
$Windows10Only = $true

$FoldersToCheck = @(
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "Temp")
        ThresholdMB = 512
        Label = "WindowsTemp"
    },
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download")
        ThresholdMB = 2048
        Label = "WindowsUpdateDownloadCache"
    },
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache")
        ThresholdMB = 2048
        Label = "DeliveryOptimizationCache"
    },
    @{
        Path = (Join-Path -Path $env:SystemDrive -ChildPath "Windows.old")
        ThresholdMB = 1024
        Label = "WindowsOld"
    }
)

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

function Get-FolderSizeBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $sum = (
            Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum
        ).Sum

        if ($null -eq $sum) {
            $sum = 0
        }

        return [int64]$sum
    }
    catch {
        return $null
    }
}

function Convert-BytesToMB {
    param(
        [Parameter(Mandatory = $true)]
        [Int64]$Bytes
    )

    return [math]::Round(($Bytes / 1MB), 2)
}

try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-Output ("{0} v{1}: OS is not Windows 10. Status=NotApplicable" -f $ScriptName, $Version)
        exit 0
    }

    $freeSpaceGB = Get-SystemDriveFreeSpaceGB
    $cleanupCandidates = New-Object System.Collections.Generic.List[string]

    if ($null -eq $freeSpaceGB) {
        Write-Output ("{0} v{1}: remediation required. FreeSpaceGB=Unknown" -f $ScriptName, $Version)
        exit 1
    }

    foreach ($folder in $FoldersToCheck) {
        $sizeBytes = Get-FolderSizeBytes -Path $folder.Path

        if ($null -eq $sizeBytes) {
            continue
        }

        $sizeMB = Convert-BytesToMB -Bytes $sizeBytes

        if ($sizeMB -gt [double]$folder.ThresholdMB) {
            $cleanupCandidates.Add("{0} SizeMB={1} ThresholdMB={2}" -f $folder.Label, $sizeMB, $folder.ThresholdMB)
        }
    }

    if ($freeSpaceGB -lt $MinimumFreeSpaceGB) {
        if ($cleanupCandidates.Count -gt 0) {
            Write-Output ("{0} v{1}: remediation required. FreeSpaceGB={2}; RequiredFreeSpaceGB={3}; CleanupCandidates={4}" -f $ScriptName, $Version, $freeSpaceGB, $MinimumFreeSpaceGB, ($cleanupCandidates -join " | "))
        }
        else {
            Write-Output ("{0} v{1}: remediation required. FreeSpaceGB={2}; RequiredFreeSpaceGB={3}; CleanupCandidates=None" -f $ScriptName, $Version, $freeSpaceGB, $MinimumFreeSpaceGB)
        }

        exit 1
    }

    if ($cleanupCandidates.Count -gt 0) {
        Write-Output ("{0} v{1}: cleanup candidates detected but free space is compliant. FreeSpaceGB={2}; Candidates={3}" -f $ScriptName, $Version, $freeSpaceGB, ($cleanupCandidates -join " | "))
        exit 0
    }

    Write-Output ("{0} v{1}: storage is ready. FreeSpaceGB={2}; RequiredFreeSpaceGB={3}" -f $ScriptName, $Version, $freeSpaceGB, $MinimumFreeSpaceGB)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
