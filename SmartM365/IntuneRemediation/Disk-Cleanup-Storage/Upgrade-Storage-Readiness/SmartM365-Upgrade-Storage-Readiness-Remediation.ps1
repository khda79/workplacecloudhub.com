# Name: SmartM365-Upgrade-Storage-Readiness-Remediation.ps1
# Version: 1.0
# Description: Frees disk space to improve Windows upgrade readiness without forcing a reboot.

$ErrorActionPreference = "Stop"

$Scenario = "Upgrade-Storage-Readiness"
$MinimumTargetFreeSpaceGB = 30
$Windows10Only = $true
$SystemDrive = $env:SystemDrive
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$IncludeInstallerPatchCache = $false

$FoldersToPurge = @(
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "Temp")
        Label = "WindowsTemp"
    },
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download")
        Label = "WindowsUpdateDownloadCache"
        RequiresUpdateServicesStop = $true
    },
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath "ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache")
        Label = "DeliveryOptimizationCache"
        RequiresUpdateServicesStop = $true
    },
    @{
        Path = (Join-Path -Path $SystemDrive -ChildPath "Windows.old")
        Label = "WindowsOld"
    },
    @{
        Path = (Join-Path -Path $env:WINDIR -ChildPath 'Installer\$PatchCache$')
        Label = "InstallerPatchCache"
        Sensitive = $true
    }
)

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

function Get-SystemDriveFreeSpaceGB {
    try {
        if ([string]::IsNullOrWhiteSpace($SystemDrive)) {
            return $null
        }

        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$SystemDrive'" -ErrorAction Stop

        if ($null -eq $drive -or $drive.DriveType -ne 3) {
            return $null
        }

        return [math]::Round(($drive.FreeSpace / 1GB), 2)
    }
    catch {
        return $null
    }
}

function Invoke-UpdateServiceStop {
    foreach ($serviceName in @("wuauserv", "bits", "dosvc")) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

            if ($null -ne $service -and $service.Status -ne "Stopped") {
                Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                Write-SmartM365Log "ServiceStopRequested=$serviceName"
            }
        }
        catch {
            Write-SmartM365Log "ServiceStopFailed=$serviceName Message=$($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2
}

function Invoke-UpdateServiceStart {
    foreach ($serviceName in @("bits", "dosvc", "wuauserv")) {
        try {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

            if ($null -ne $service) {
                Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                Write-SmartM365Log "ServiceStartRequested=$serviceName"
            }
        }
        catch {
            Write-SmartM365Log "ServiceStartFailed=$serviceName Message=$($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2
}

function Clear-FolderContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-SmartM365Log "CleanupSkipped=EmptyPath"
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-SmartM365Log "CleanupNotFound=$Path"
        return
    }

    try {
        Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-SmartM365Log "CleanupItemSkipped=$($_.FullName) Message=$($_.Exception.Message)"
                }
            }

        Write-SmartM365Log "CleanupCompleted=$Path"
    }
    catch {
        Write-SmartM365Log "CleanupPartial=$Path Message=$($_.Exception.Message)"
    }
}

function Clear-UserTemp {
    try {
        Get-ChildItem -LiteralPath "$SystemDrive\Users" -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin @("Default", "Default User", "Public", "All Users") } |
            ForEach-Object {
                $userTemp = Join-Path -Path $_.FullName -ChildPath "AppData\Local\Temp"

                if (Test-Path -LiteralPath $userTemp) {
                    Clear-FolderContent -Path $userTemp
                }
            }
    }
    catch {
        Write-SmartM365Log "UserTempCleanupPartial Message=$($_.Exception.Message)"
    }
}

function Clear-RecycleBinSafe {
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
        Write-SmartM365Log "RecycleBinCleanupCompleted"
    }
    catch {
        Write-SmartM365Log "RecycleBinCleanupSkipped Message=$($_.Exception.Message)"
    }
}

function Invoke-DismComponentCleanup {
    try {
        $process = Start-Process -FilePath "dism.exe" -ArgumentList "/Online", "/Cleanup-Image", "/StartComponentCleanup" -Wait -PassThru -WindowStyle Hidden
        Write-SmartM365Log "DismComponentCleanupExitCode=$($process.ExitCode)"
    }
    catch {
        Write-SmartM365Log "DismComponentCleanupFailed Message=$($_.Exception.Message)"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null

    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-Output "OS=NotWindows10 Status=NotApplicable"
        exit 0
    }

    Write-SmartM365Log "RemediationStarted"
    Write-SmartM365Log "SystemDrive=$SystemDrive"
    Write-SmartM365Log "MinimumTargetFreeSpaceGB=$MinimumTargetFreeSpaceGB"

    $beforeFreeGB = Get-SystemDriveFreeSpaceGB
    Write-SmartM365Log "FreeSpaceGBBefore=$beforeFreeGB"

    $requiresServiceStop = $false

    foreach ($folder in $FoldersToPurge) {
        if ($folder.ContainsKey("RequiresUpdateServicesStop") -and $folder.RequiresUpdateServicesStop -eq $true) {
            $requiresServiceStop = $true
        }
    }

    if ($requiresServiceStop) {
        Invoke-UpdateServiceStop
    }

    foreach ($folder in $FoldersToPurge) {
        if ($folder.ContainsKey("Sensitive") -and $folder.Sensitive -eq $true -and -not $IncludeInstallerPatchCache) {
            Write-SmartM365Log "CleanupSkipped=$($folder.Label) Reason=SensitiveTargetDisabled"
            continue
        }

        Clear-FolderContent -Path $folder.Path
    }

    if ($requiresServiceStop) {
        Invoke-UpdateServiceStart
    }

    Clear-UserTemp
    Clear-RecycleBinSafe
    Invoke-DismComponentCleanup

    $afterFreeGB = Get-SystemDriveFreeSpaceGB
    Write-SmartM365Log "FreeSpaceGBAfter=$afterFreeGB"

    if ($null -eq $afterFreeGB) {
        Write-SmartM365Log "Status=CompletedButFreeSpaceUnknown"
        exit 1
    }

    if ($afterFreeGB -lt $MinimumTargetFreeSpaceGB) {
        Write-SmartM365Log "Status=CompletedButStillBelowThreshold"
        exit 1
    }

    Write-SmartM365Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
