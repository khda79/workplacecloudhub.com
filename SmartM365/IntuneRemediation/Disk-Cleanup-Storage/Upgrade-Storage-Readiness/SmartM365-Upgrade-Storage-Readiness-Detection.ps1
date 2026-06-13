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

function Get-FolderByteSize {
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

function ConvertTo-SingleLineValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    $text = $text -replace '[\r\n]+', ' '
    return $text.Trim()
}

function Get-CleanupCandidateSummary {
    $cleanupCandidates = New-Object System.Collections.Generic.List[string]

    foreach ($folder in $FoldersToCheck) {
        $sizeBytes = Get-FolderByteSize -Path $folder.Path

        if ($null -eq $sizeBytes) {
            continue
        }

        $sizeMB = Convert-BytesToMB -Bytes $sizeBytes

        if ($sizeMB -gt [double]$folder.ThresholdMB) {
            $cleanupCandidates.Add("{0}:{1}MB" -f $folder.Label, $sizeMB)
        }
    }

    if ($cleanupCandidates.Count -eq 0) {
        return "None"
    }

    return ($cleanupCandidates -join ",")
}

try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-Output ("Status=NotApplicable Reason=NotWindows10 Script={0} Version={1}" -f $ScriptName, $Version)
        exit 0
    }

    $freeSpaceGB = Get-SystemDriveFreeSpaceGB

    if ($null -eq $freeSpaceGB) {
        Write-Output ("Status=RemediationRequired Reason=FreeSpaceUnknown Script={0} Version={1}" -f $ScriptName, $Version)
        exit 1
    }

    if ($freeSpaceGB -ge $MinimumFreeSpaceGB) {
        Write-Output ("Status=Ready FreeSpaceGB={0} RequiredFreeSpaceGB={1}" -f $freeSpaceGB, $MinimumFreeSpaceGB)
        exit 0
    }

    $cleanupCandidateSummary = Get-CleanupCandidateSummary
    Write-Output ("Status=RemediationRequired FreeSpaceGB={0} RequiredFreeSpaceGB={1} CleanupCandidates={2}" -f $freeSpaceGB, $MinimumFreeSpaceGB, $cleanupCandidateSummary)
    exit 1
}
catch {
    $message = ConvertTo-SingleLineValue $_.Exception.Message
    Write-Output ("Status=Error Script={0} Version={1} Message={2}" -f $ScriptName, $Version, $message)
    exit 1
}
