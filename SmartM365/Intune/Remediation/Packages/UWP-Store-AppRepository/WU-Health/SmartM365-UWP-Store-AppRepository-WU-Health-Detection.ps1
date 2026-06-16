<# 
    Name: SmartM365-UWP-Store-AppRepository-WU-Health-Detection.ps1
    Version: 1.1
    Description: Detects issues that may prevent Windows upgrade or UWP/AppX operations because Microsoft Store, AppRepository, WindowsApps, or Windows Update cache components are corrupted or unavailable.

    Intended use:
    - Microsoft Intune detection script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
#>

$ErrorActionPreference = "Stop"

$issues = New-Object System.Collections.Generic.List[string]

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
        $value = Format-CompactText -Text ([string]$Data[$key]) -MaxLength 240
        $parts.Add(("{0}={1}" -f $key, $value))
    }

    Write-Output ($parts -join "; ")
}

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $script:issues.Add($Message)
}

function Test-ServiceHealth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [bool]$RequireRunning = $false
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($null -eq $service) {
        Add-Issue "Required service is missing: $Name"
        return
    }

    $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

    if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
        Add-Issue "Required service is disabled: $Name"
        return
    }

    if ($RequireRunning -and $service.Status -ne "Running") {
        Add-Issue "Required service is not running: $Name"
        return
    }

    if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
        Add-Issue "Required service is in an unhealthy state: $Name ($($service.Status))"
    }
}

try {
    # ---------------------------------------------------------
    # 1. WindowsApps folder health
    # ---------------------------------------------------------
    $windowsAppsPath = "C:\Program Files\WindowsApps"

    if (-not (Test-Path -Path $windowsAppsPath)) {
        Add-Issue "WindowsApps folder is missing"
    }
    else {
        try {
            $windowsAppsItems = Get-ChildItem -Path $windowsAppsPath -Directory -ErrorAction Stop

            if ($windowsAppsItems.Count -lt 50) {
                Add-Issue "WindowsApps contains too few package folders; possible corruption"
            }

            $knownPackages = @(
                "Microsoft.WindowsStore*",
                "Microsoft.DesktopAppInstaller*",
                "Microsoft.VCLibs*",
                "Microsoft.NET.Native.Runtime*",
                "Microsoft.NET.Native.Framework*"
            )

            foreach ($knownPackage in $knownPackages) {
                $match = $windowsAppsItems | Where-Object { $_.Name -like $knownPackage } | Select-Object -First 1
                if ($null -eq $match) {
                    Add-Issue "Expected WindowsApps package folder was not found: $knownPackage"
                }
            }
        }
        catch {
            Add-Issue "WindowsApps folder is unreadable or permissions appear broken"
        }
    }

    # ---------------------------------------------------------
    # 2. AppRepository health
    # ---------------------------------------------------------
    $appRepositoryPath = "C:\ProgramData\Microsoft\Windows\AppRepository"

    if (-not (Test-Path -Path $appRepositoryPath)) {
        Add-Issue "AppRepository folder is missing"
    }
    else {
        try {
            $appRepositoryItems = Get-ChildItem -Path $appRepositoryPath -ErrorAction Stop

            if ($appRepositoryItems.Count -lt 20) {
                Add-Issue "AppRepository contains too few files; possible corruption"
            }

            $requiredAppRepositoryFiles = @(
                "StateRepository-Machine.srd",
                "StateRepository-Machine.srd-shm",
                "StateRepository-Machine.srd-wal"
            )

            foreach ($requiredFile in $requiredAppRepositoryFiles) {
                $requiredFilePath = Join-Path -Path $appRepositoryPath -ChildPath $requiredFile
                if (-not (Test-Path -Path $requiredFilePath)) {
                    Add-Issue "Required AppRepository file is missing: $requiredFile"
                }
            }
        }
        catch {
            Add-Issue "AppRepository folder is unreadable or corrupted"
        }
    }

    # ---------------------------------------------------------
    # 3. Critical UWP/AppX services
    # ---------------------------------------------------------
    # These services do not always need to be running permanently.
    # The detection focuses on missing or disabled services, and unhealthy transient states.
    $criticalServices = @(
        "AppXSvc",
        "ClipSVC",
        "StateRepository"
    )

    foreach ($serviceName in $criticalServices) {
        Test-ServiceHealth -Name $serviceName -RequireRunning $false
    }

    # WpnService may be disabled by hardening baselines and is not always required for Windows upgrade.
    # It is checked as informational only when present.
    $wpnService = Get-Service -Name "WpnService" -ErrorAction SilentlyContinue
    if ($null -ne $wpnService -and $wpnService.Status -in @("StopPending", "PausePending", "Paused")) {
        Add-Issue "WpnService is in an unhealthy state: $($wpnService.Status)"
    }

    # ---------------------------------------------------------
    # 4. Microsoft Store package registration
    # ---------------------------------------------------------
    $storePackage = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue

    if ($null -eq $storePackage) {
        Add-Issue "Microsoft Store package is missing or not registered"
    }
    else {
        $storePackageLocations = $storePackage | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.InstallLocation) -and
            (Test-Path -Path $_.InstallLocation)
        }

        if ($null -eq $storePackageLocations) {
            Add-Issue "Microsoft Store package is registered but its install location is missing"
        }
    }

    # ---------------------------------------------------------
    # 5. Essential UWP frameworks
    # ---------------------------------------------------------
    $frameworks = @(
        "Microsoft.NET.Native.Runtime",
        "Microsoft.NET.Native.Framework",
        "Microsoft.VCLibs",
        "Microsoft.UI.Xaml"
    )

    $allAppxPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

    foreach ($framework in $frameworks) {
        $package = $allAppxPackages | Where-Object { $_.Name -like "$framework*" } | Select-Object -First 1

        if ($null -eq $package) {
            Add-Issue "Required UWP framework is missing: $framework"
        }
    }

    # ---------------------------------------------------------
    # 6. Windows Update SoftwareDistribution cache sanity
    # ---------------------------------------------------------
    $softwareDistributionDownloadPath = "C:\Windows\SoftwareDistribution\Download"
    $softwareDistributionIsEmpty = $false

    if (-not (Test-Path -Path $softwareDistributionDownloadPath)) {
        Add-Issue "Windows Update download cache folder is missing"
    }
    else {
        $downloadFiles = Get-ChildItem -Path $softwareDistributionDownloadPath -Recurse -File -ErrorAction SilentlyContinue
        $downloadDirectories = Get-ChildItem -Path $softwareDistributionDownloadPath -Directory -ErrorAction SilentlyContinue

        # An empty Download folder is not automatically corruption.
        # It becomes suspicious only if recent Windows Update/AppX errors are also present.
        $softwareDistributionIsEmpty = (($null -eq $downloadFiles -or $downloadFiles.Count -eq 0) -and ($null -eq $downloadDirectories -or $downloadDirectories.Count -eq 0))
    }

    # ---------------------------------------------------------
    # 7. WinRM health
    # ---------------------------------------------------------
    # WinRM is not required for Windows upgrade itself, but can be required by enterprise remediation tooling.
    $winRmService = Get-Service -Name "WinRM" -ErrorAction SilentlyContinue

    if ($null -eq $winRmService) {
        Add-Issue "WinRM service is missing"
    }
    else {
        $winRmServiceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue

        if ($null -ne $winRmServiceCim -and $winRmServiceCim.StartMode -eq "Disabled") {
            Add-Issue "WinRM service is disabled"
        }
    }

    # ---------------------------------------------------------
    # 8. Event Log checks for AppX / Store / UWP / Windows Update
    # ---------------------------------------------------------
    $appxErrors = Get-WinEvent -LogName "Microsoft-Windows-AppXDeploymentServer/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LevelDisplayName -eq "Error" -and
            $_.TimeCreated -ge (Get-Date).AddDays(-7)
        }

    if ($null -ne $appxErrors -and $appxErrors.Count -gt 0) {
        Add-Issue "Recent AppX deployment errors were detected in the event log"
    }

    $windowsUpdateErrors = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.LevelDisplayName -eq "Error" -and
            $_.TimeCreated -ge (Get-Date).AddDays(-7)
        }

    if ($null -ne $windowsUpdateErrors -and $windowsUpdateErrors.Count -gt 0) {
        Add-Issue "Recent Windows Update errors were detected in the event log"
    }

    if ($softwareDistributionIsEmpty -and (($null -ne $windowsUpdateErrors -and $windowsUpdateErrors.Count -gt 0) -or ($null -ne $appxErrors -and $appxErrors.Count -gt 0))) {
        Add-Issue "SoftwareDistribution download cache is empty while recent Windows Update or AppX errors exist"
    }

    # ---------------------------------------------------------
    # Final result
    # ---------------------------------------------------------
    if ($issues.Count -gt 0) {
        $sampleIssues = @($issues | Select-Object -First 5)
        Write-IntuneResult -Status "IssuesDetected" -Data @{
            IssueCount = $issues.Count
            Samples = ($sampleIssues -join " | ")
        }
        exit 1
    }

    Write-IntuneResult -Status "Healthy" -Data @{
        Checks = "UWP,Store,AppRepository,WindowsApps,WinRM,WindowsUpdate"
    }
    exit 0
}
catch {
    Write-IntuneResult -Status "Error" -Data @{
        Message = $_.Exception.Message
    }
    exit 1
}
