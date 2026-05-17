<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Remediation.ps1
    Version: 1.0
    Description: Remediates common Delivery Optimization, Dynamic Download, BITS, and Windows Update content engine issues.

    Logs:
    - C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-DeliveryOptimization-ContentEngine-Health\
#>

[CmdletBinding()]
param(
    [bool]$ResetBitsJobs = $true,
    [bool]$TriggerWindowsUpdateScan = $true
)

$ErrorActionPreference = "Stop"

$RemediationName = "Remediate-DeliveryOptimization-ContentEngine-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$RemediationName"

if (-not (Test-Path -Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"
$ErrorFound = $false

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Set-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Stop-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-Log "ServiceNotFound=${Name}"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-Log "ServiceStopRequested=${Name}"
        }
        else {
            Write-Log "ServiceAlreadyStopped=${Name}"
        }
    }
    catch {
        Set-RemediationError "Failed to stop service ${Name}: $($_.Exception.Message)"
    }
}

function Start-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-Log "ServiceNotFound=${Name}"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Write-Log "ServiceStartupTypeChanged=${Name} StartupType=Manual"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-Log "ServiceStartRequested=${Name}"
        }
        else {
            Write-Log "ServiceAlreadyRunning=${Name}"
        }
    }
    catch {
        Set-RemediationError "Failed to start service ${Name}: $($_.Exception.Message)"
    }
}

function Clear-FolderContentSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log "Cleanup=NotFound Path=${Path}"
            return
        }

        Write-Log "Cleanup=Start Path=${Path}"

        Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log "Cleanup=ItemSkipped Path=$($_.FullName) Message=$($_.Exception.Message)"
            }
        }

        Write-Log "Cleanup=Completed Path=${Path}"
    }
    catch {
        Set-RemediationError "Failed to clean folder ${Path}: $($_.Exception.Message)"
    }
}

function Reset-BitsTransfersSafe {
    try {
        if (-not $ResetBitsJobs) {
            Write-Log "BITSReset=Skipped"
            return
        }

        Write-Log "BITSReset=Start"

        $bitsAdminPath = Join-Path -Path $env:WINDIR -ChildPath "System32\bitsadmin.exe"

        if (Test-Path -LiteralPath $bitsAdminPath -PathType Leaf) {
            $process = Start-Process -FilePath $bitsAdminPath -ArgumentList "/reset", "/allusers" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-Log "BITSReset=Completed ExitCode=$($process.ExitCode)"
        }
        else {
            Write-Log "BITSReset=BitsadminNotFound"
        }
    }
    catch {
        Set-RemediationError "Failed to reset BITS transfers: $($_.Exception.Message)"
    }
}

function Trigger-WindowsUpdateScanSafe {
    try {
        if (-not $TriggerWindowsUpdateScan) {
            Write-Log "WindowsUpdateScan=Skipped"
            return
        }

        $usoClientPath = Join-Path -Path $env:WINDIR -ChildPath "System32\UsoClient.exe"

        if (Test-Path -LiteralPath $usoClientPath -PathType Leaf) {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-Log "WindowsUpdateScan=Triggered"
        }
        else {
            Write-Log "WindowsUpdateScan=UsoClientNotFound"
        }
    }
    catch {
        Set-RemediationError "Failed to trigger Windows Update scan: $($_.Exception.Message)"
    }
}

try {
    Write-Log "===== Delivery Optimization remediation started ====="

    $deliveryOptimizationCachePaths = @(
        "C:\ProgramData\Microsoft\Windows\DeliveryOptimization\Cache",
        "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
    )

    $windowsUpdateDownloadCache = Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download"

    foreach ($serviceName in @("DoSvc", "BITS", "wuauserv", "UsoSvc")) {
        Stop-ServiceSafe -Name $serviceName
    }

    Start-Sleep -Seconds 3

    foreach ($cachePath in $deliveryOptimizationCachePaths) {
        Clear-FolderContentSafe -Path $cachePath
    }

    Clear-FolderContentSafe -Path $windowsUpdateDownloadCache

    Reset-BitsTransfersSafe

    foreach ($serviceName in @("BITS", "DoSvc", "wuauserv", "UsoSvc")) {
        Start-ServiceSafe -Name $serviceName
    }

    Start-Sleep -Seconds 5

    Trigger-WindowsUpdateScanSafe

    Write-Log "===== Delivery Optimization remediation finished ====="

    if ($ErrorFound) {
        Write-Log "Status=CompletedWithErrors"
        exit 1
    }

    Write-Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"

    try {
        Set-RemediationError $_.Exception.Message
    }
    catch { }

    exit 1
}

