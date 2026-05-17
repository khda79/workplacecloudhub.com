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

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Add-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=${Name}"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStopRequested=${Name}"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyStopped=${Name}"
        }
    }
    catch {
        Add-RemediationError "Failed to stop service ${Name}: $($_.Exception.Message)"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=${Name}"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartupTypeChanged=${Name} StartupType=Manual"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartRequested=${Name}"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyRunning=${Name}"
        }
    }
    catch {
        Add-RemediationError "Failed to start service ${Name}: $($_.Exception.Message)"
    }
}

function Clear-FolderContentSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-SmartM365Log "Cleanup=NotFound Path=${Path}"
            return
        }

        Write-SmartM365Log "Cleanup=Start Path=${Path}"

        Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-SmartM365Log "Cleanup=ItemSkipped Path=$($_.FullName) Message=$($_.Exception.Message)"
            }
        }

        Write-SmartM365Log "Cleanup=Completed Path=${Path}"
    }
    catch {
        Add-RemediationError "Failed to clean folder ${Path}: $($_.Exception.Message)"
    }
}

function Invoke-BitsTransferResetSafe {
    param([bool]$ResetBitsJobs)

    try {
        if (-not $ResetBitsJobs) {
            Write-SmartM365Log "BITSReset=Skipped"
            return
        }

        Write-SmartM365Log "BITSReset=Start"

        $bitsAdminPath = Join-Path -Path $env:WINDIR -ChildPath "System32\bitsadmin.exe"

        if (Test-Path -LiteralPath $bitsAdminPath -PathType Leaf) {
            $process = Start-Process -FilePath $bitsAdminPath -ArgumentList "/reset", "/allusers" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-SmartM365Log "BITSReset=Completed ExitCode=$($process.ExitCode)"
        }
        else {
            Write-SmartM365Log "BITSReset=BitsadminNotFound"
        }
    }
    catch {
        Add-RemediationError "Failed to reset BITS transfers: $($_.Exception.Message)"
    }
}

function Invoke-WindowsUpdateScanSafe {
    param([bool]$TriggerWindowsUpdateScan)

    try {
        if (-not $TriggerWindowsUpdateScan) {
            Write-SmartM365Log "WindowsUpdateScan=Skipped"
            return
        }

        $usoClientPath = Join-Path -Path $env:WINDIR -ChildPath "System32\UsoClient.exe"

        if (Test-Path -LiteralPath $usoClientPath -PathType Leaf) {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-SmartM365Log "WindowsUpdateScan=Triggered"
        }
        else {
            Write-SmartM365Log "WindowsUpdateScan=UsoClientNotFound"
        }
    }
    catch {
        Add-RemediationError "Failed to trigger Windows Update scan: $($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Delivery Optimization remediation started ====="

    $deliveryOptimizationCachePaths = @(
        "C:\ProgramData\Microsoft\Windows\DeliveryOptimization\Cache",
        "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
    )

    $windowsUpdateDownloadCache = Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download"

    foreach ($serviceName in @("DoSvc", "BITS", "wuauserv", "UsoSvc")) {
        Invoke-ServiceStopSafe -Name $serviceName
    }

    Start-Sleep -Seconds 3

    foreach ($cachePath in $deliveryOptimizationCachePaths) {
        Clear-FolderContentSafe -Path $cachePath
    }

    Clear-FolderContentSafe -Path $windowsUpdateDownloadCache

    Invoke-BitsTransferResetSafe -ResetBitsJobs $ResetBitsJobs

    foreach ($serviceName in @("BITS", "DoSvc", "wuauserv", "UsoSvc")) {
        Invoke-ServiceStartSafe -Name $serviceName
    }

    Start-Sleep -Seconds 5

    Invoke-WindowsUpdateScanSafe -TriggerWindowsUpdateScan $TriggerWindowsUpdateScan

    Write-SmartM365Log "===== Delivery Optimization remediation finished ====="

    if ($ErrorFound) {
        Write-SmartM365Log "Status=CompletedWithErrors"
        exit 1
    }

    Write-SmartM365Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"

    try {
        Add-RemediationError $_.Exception.Message
    }
    catch {
        Write-Output "Status=ErrorDuringErrorHandling"
        Write-Output "Message=$($_.Exception.Message)"
    }

    exit 1
}

