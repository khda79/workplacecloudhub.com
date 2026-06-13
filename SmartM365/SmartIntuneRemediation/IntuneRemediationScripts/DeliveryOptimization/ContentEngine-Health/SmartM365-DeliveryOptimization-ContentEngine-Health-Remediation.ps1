<#
    Name: SmartM365-DeliveryOptimization-ContentEngine-Health-Remediation.ps1
    Version: 1.1
    Description: Safely remediates Delivery Optimization and BITS content engine state with compact Intune output and detailed local logging.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
    - No forced reboot

    Logs:
    - C:\ProgramData\SmartM365\IntuneRemediation\Logs\DeliveryOptimization-ContentEngine-Health\Remediate-DeliveryOptimization-ContentEngine-Health.log
#>

[CmdletBinding()]
param(
    [bool]$ResetBitsJobs = $true,
    [bool]$TriggerWindowsUpdateScan = $true,
    [bool]$CleanWindowsUpdateDownloadCache = $true
)

$ErrorActionPreference = "Stop"

$ScenarioName = "DeliveryOptimization-ContentEngine-Health"
$RemediationName = "Remediate-DeliveryOptimization-ContentEngine-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$ScenarioName"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"

if (-not (Test-Path -LiteralPath $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$ErrorFound = $false
$Actions = New-Object System.Collections.Generic.List[string]
$Skipped = New-Object System.Collections.Generic.List[string]
$RemediationErrors = New-Object System.Collections.Generic.List[string]

function Format-CompactText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$MaxLength = 180
    )

    $compactText = ($Text -replace "\s+", " ").Trim()

    if ($compactText.Length -gt $MaxLength) {
        return ($compactText.Substring(0, $MaxLength) + "...")
    }

    return $compactText
}

function Write-IntuneResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [hashtable]$Data = @{}
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Status=$Status")

    foreach ($key in ($Data.Keys | Sort-Object)) {
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 260
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Add-Action {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log $Message
    $script:Actions.Add($Message)
}

function Add-SkippedAction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log $Message
    $script:Skipped.Add($Message)
}

function Add-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
    $script:RemediationErrors.Add($Message)
}

function Invoke-ServiceStopIfRunning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-SkippedAction "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Add-Action "ServiceStopRequested=$Name"
        }
        else {
            Add-SkippedAction "ServiceAlreadyStopped=$Name"
        }
    }
    catch {
        Add-RemediationError "Failed to stop service ${Name}: $($_.Exception.Message)"
    }
}

function Invoke-ServiceStartIfAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-SkippedAction "ServiceNotFound=$Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Add-Action "ServiceStartupTypeChanged=$Name StartupType=Manual"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Add-Action "ServiceStartRequested=$Name"
        }
        else {
            Add-SkippedAction "ServiceAlreadyRunning=$Name"
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
            Add-SkippedAction "CleanupNotFound=$Path"
            return
        }

        $removedCount = 0
        $skippedCount = 0
        Write-SmartM365Log "CleanupStart=$Path"

        $items = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)

        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                $removedCount++
            }
            catch {
                $skippedCount++
                Write-SmartM365Log "CleanupItemSkipped=$($item.FullName) Message=$($_.Exception.Message)"
            }
        }

        Add-Action "CleanupCompleted=$Path Removed=$removedCount Skipped=$skippedCount"
    }
    catch {
        Add-RemediationError "Failed to clean folder ${Path}: $($_.Exception.Message)"
    }
}

function Invoke-BitsTransferResetSafe {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$ResetBitsJobs
    )

    try {
        if (-not $ResetBitsJobs) {
            Add-SkippedAction "BITSReset=Skipped"
            return
        }

        $bitsAdminPath = Join-Path -Path $env:WINDIR -ChildPath "System32\bitsadmin.exe"

        if (-not (Test-Path -LiteralPath $bitsAdminPath -PathType Leaf)) {
            Add-SkippedAction "BITSReset=BitsadminNotFound"
            return
        }

        $process = Start-Process -FilePath $bitsAdminPath -ArgumentList "/reset", "/allusers" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
        Add-Action "BITSReset=Completed ExitCode=$($process.ExitCode)"
    }
    catch {
        Add-RemediationError "Failed to reset BITS transfers: $($_.Exception.Message)"
    }
}

function Invoke-WindowsUpdateScanSafe {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$TriggerWindowsUpdateScan
    )

    try {
        if (-not $TriggerWindowsUpdateScan) {
            Add-SkippedAction "WindowsUpdateScan=Skipped"
            return
        }

        $usoClientPath = Join-Path -Path $env:WINDIR -ChildPath "System32\UsoClient.exe"

        if (-not (Test-Path -LiteralPath $usoClientPath -PathType Leaf)) {
            Add-SkippedAction "WindowsUpdateScan=UsoClientNotFound"
            return
        }

        Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
        Add-Action "WindowsUpdateScan=Triggered"
    }
    catch {
        Add-RemediationError "Failed to trigger Windows Update scan: $($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Delivery Optimization remediation started ====="
    Write-SmartM365Log "ResetBitsJobs=$ResetBitsJobs TriggerWindowsUpdateScan=$TriggerWindowsUpdateScan CleanWindowsUpdateDownloadCache=$CleanWindowsUpdateDownloadCache"

    $deliveryOptimizationCachePaths = @(
        "C:\ProgramData\Microsoft\Windows\DeliveryOptimization\Cache",
        "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
    )

    foreach ($serviceName in @("DoSvc", "BITS")) {
        Invoke-ServiceStopIfRunning -Name $serviceName
    }

    Start-Sleep -Seconds 2

    foreach ($cachePath in $deliveryOptimizationCachePaths) {
        Clear-FolderContentSafe -Path $cachePath
    }

    if ($CleanWindowsUpdateDownloadCache) {
        foreach ($serviceName in @("wuauserv", "UsoSvc")) {
            Invoke-ServiceStopIfRunning -Name $serviceName
        }

        $windowsUpdateDownloadCache = Join-Path -Path $env:WINDIR -ChildPath "SoftwareDistribution\Download"
        Clear-FolderContentSafe -Path $windowsUpdateDownloadCache
    }
    else {
        Add-SkippedAction "WindowsUpdateDownloadCacheCleanup=Skipped"
    }

    Invoke-BitsTransferResetSafe -ResetBitsJobs $ResetBitsJobs

    foreach ($serviceName in @("BITS", "DoSvc", "wuauserv", "UsoSvc")) {
        Invoke-ServiceStartIfAvailable -Name $serviceName
    }

    Start-Sleep -Seconds 3

    Invoke-WindowsUpdateScanSafe -TriggerWindowsUpdateScan $TriggerWindowsUpdateScan

    Write-SmartM365Log "===== Delivery Optimization remediation finished ====="

    $sampleErrors = ""
    if ($RemediationErrors.Count -gt 0) {
        $sampleErrors = (($RemediationErrors | Select-Object -First 3) -join " | ")
    }

    $sampleActions = (($Actions | Select-Object -First 5) -join " | ")

    if ($ErrorFound) {
        Write-IntuneResult -Status "CompletedWithErrors" -Data @{
            ActionCount = $Actions.Count
            SkippedCount = $Skipped.Count
            ErrorCount = $RemediationErrors.Count
            Errors = $sampleErrors
            Log = $LogPath
        }

        exit 1
    }

    Write-IntuneResult -Status "Completed" -Data @{
        ActionCount = $Actions.Count
        SkippedCount = $Skipped.Count
        Actions = $sampleActions
        Log = $LogPath
    }

    exit 0
}
catch {
    try {
        Add-RemediationError $_.Exception.Message
    }
    catch {
        $null = $_
    }

    Write-IntuneResult -Status "TechnicalError" -Data @{
        Message = $_.Exception.Message
        Log = $LogPath
    }

    exit 1
}
