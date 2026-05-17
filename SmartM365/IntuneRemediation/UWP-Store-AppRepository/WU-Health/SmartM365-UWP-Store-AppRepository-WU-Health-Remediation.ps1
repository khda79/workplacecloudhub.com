<# 
    Name: SmartM365-UWP-Store-AppRepository-WU-Health-Remediation.ps1
    Version: 1.0
    Description: Safely remediates common Windows Update, Microsoft Store, UWP/AppX, AppRepository, and WinRM health issues that may block Windows upgrade scenarios.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell when possible
    - No forced reboot

    Notes:
    - This script avoids destructive AppRepository repair by default.
    - This script does not reset WindowsApps ACLs recursively because that can break Microsoft Store/UWP servicing.
#>

$ErrorActionPreference = "Stop"

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\WU-Health'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Remediate-UWP-Store-AppRepository-WU-Health.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$ErrorFound = $false

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $line
}

function Set-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Start-ServiceIfAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$StartupType = "Manual"
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Set-RemediationError "Service is missing: $Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction SilentlyContinue
            Write-Log "Service startup type updated: $Name -> $StartupType"
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-Log "Service start requested: $Name"
        }
    }
    catch {
        Set-RemediationError "Failed to repair service ${Name}: $($_.Exception.Message)"
    }
}

Write-Log "===== Remediation started ====="

# ---------------------------------------------------------
# 1. WinRM
# ---------------------------------------------------------
Write-Log "Checking WinRM service"
try {
    $winRmService = Get-Service -Name "WinRM" -ErrorAction SilentlyContinue

    if ($null -eq $winRmService) {
        Set-RemediationError "WinRM service is missing"
    }
    else {
        $winRmCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue

        if ($null -ne $winRmCim -and $winRmCim.StartMode -eq "Disabled") {
            Set-Service -Name "WinRM" -StartupType Manual -ErrorAction SilentlyContinue
            Write-Log "WinRM startup type changed from Disabled to Manual"
        }

        Start-Service -Name "WinRM" -ErrorAction SilentlyContinue
        Write-Log "WinRM start requested"
    }
}
catch {
    Set-RemediationError "WinRM remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 2. Windows Update cache remediation
# ---------------------------------------------------------
Write-Log "Cleaning Windows Update download cache"
try {
    $servicesToStop = @("wuauserv", "bits", "dosvc")

    foreach ($serviceName in $servicesToStop) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -ne $service -and $service.Status -ne "Stopped") {
            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
            Write-Log "Service stop requested: $serviceName"
        }
    }

    $downloadCachePath = "C:\Windows\SoftwareDistribution\Download"

    if (Test-Path -Path $downloadCachePath) {
        Get-ChildItem -Path $downloadCachePath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Windows Update download cache cleaned"
    }
    else {
        Write-Log "Windows Update download cache folder not found; nothing to clean"
    }

    foreach ($serviceName in @("bits", "dosvc", "wuauserv")) {
        Start-ServiceIfAvailable -Name $serviceName -StartupType Manual
    }
}
catch {
    Set-RemediationError "Windows Update cache remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 3. DISM and SFC
# ---------------------------------------------------------
Write-Log "Running DISM RestoreHealth"
try {
    $dismProcess = Start-Process -FilePath "dism.exe" -ArgumentList "/Online", "/Cleanup-Image", "/RestoreHealth" -Wait -NoNewWindow -PassThru
    Write-Log "DISM completed with exit code: $($dismProcess.ExitCode)"

    if ($dismProcess.ExitCode -notin @(0, 3010)) {
        Set-RemediationError "DISM returned unexpected exit code: $($dismProcess.ExitCode)"
    }
}
catch {
    Set-RemediationError "DISM execution failed: $($_.Exception.Message)"
}

Write-Log "Running SFC scan"
try {
    $sfcProcess = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -PassThru
    Write-Log "SFC completed with exit code: $($sfcProcess.ExitCode)"

    if ($sfcProcess.ExitCode -notin @(0, 1)) {
        Set-RemediationError "SFC returned unexpected exit code: $($sfcProcess.ExitCode)"
    }
}
catch {
    Set-RemediationError "SFC execution failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 4. UWP/AppX critical services
# ---------------------------------------------------------
Write-Log "Repairing UWP/AppX service configuration"
foreach ($serviceName in @("AppXSvc", "ClipSVC", "StateRepository")) {
    Start-ServiceIfAvailable -Name $serviceName -StartupType Manual
}

# WpnService is not forced to Automatic because it may be intentionally managed by security baselines.
Start-ServiceIfAvailable -Name "WpnService" -StartupType Manual

# ---------------------------------------------------------
# 5. AppRepository sanity check
# ---------------------------------------------------------
Write-Log "Checking AppRepository"
try {
    $appRepositoryPath = "C:\ProgramData\Microsoft\Windows\AppRepository"
    $stateRepositoryFile = Join-Path -Path $appRepositoryPath -ChildPath "StateRepository-Machine.srd"

    if (-not (Test-Path -Path $appRepositoryPath)) {
        Set-RemediationError "AppRepository folder is missing"
    }
    elseif (-not (Test-Path -Path $stateRepositoryFile)) {
        Set-RemediationError "StateRepository-Machine.srd is missing"
    }
    else {
        Write-Log "AppRepository core database file is present"

        # Do not run esentutl /p automatically.
        # Hard repair may cause data loss and should be performed only after backup and explicit approval.
    }
}
catch {
    Set-RemediationError "AppRepository check failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 6. WindowsApps ACL sanity check
# ---------------------------------------------------------
Write-Log "Checking WindowsApps accessibility"
try {
    $windowsAppsPath = "C:\Program Files\WindowsApps"

    if (-not (Test-Path -Path $windowsAppsPath)) {
        Set-RemediationError "WindowsApps folder is missing"
    }
    else {
        Get-ChildItem -Path $windowsAppsPath -Directory -ErrorAction Stop | Select-Object -First 1 | Out-Null
        Write-Log "WindowsApps folder is readable"

        # Do not run icacls /reset /t on WindowsApps automatically.
        # Recursive ACL reset can break Microsoft Store/UWP servicing.
    }
}
catch {
    Set-RemediationError "WindowsApps folder is unreadable or ACLs may be broken: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 7. Re-register UWP packages with valid manifests
# ---------------------------------------------------------
Write-Log "Re-registering UWP packages with valid manifests"
try {
    $packages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

    foreach ($package in $packages) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($package.InstallLocation) -and (Test-Path -Path $package.InstallLocation)) {
                $manifestPath = Join-Path -Path $package.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $manifestPath) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $manifestPath -ErrorAction SilentlyContinue
                    Write-Log "Package re-registration requested: $($package.Name)"
                }
            }
        }
        catch {
            Set-RemediationError "Failed to re-register package $($package.Name): $($_.Exception.Message)"
        }
    }
}
catch {
    Set-RemediationError "Global UWP package re-registration failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 8. Microsoft Store remediation
# ---------------------------------------------------------
Write-Log "Repairing Microsoft Store registration"
try {
    $storePackages = Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" -ErrorAction SilentlyContinue

    if ($null -eq $storePackages) {
        Set-RemediationError "Microsoft Store package is missing; offline package or Microsoft Store source is required for reinstall"
    }
    else {
        foreach ($storePackage in $storePackages) {
            if (-not [string]::IsNullOrWhiteSpace($storePackage.InstallLocation)) {
                $storeManifest = Join-Path -Path $storePackage.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $storeManifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $storeManifest -ErrorAction SilentlyContinue
                    Write-Log "Microsoft Store re-registration requested"
                }
                else {
                    Set-RemediationError "Microsoft Store manifest is missing"
                }
            }
            else {
                Set-RemediationError "Microsoft Store package has no InstallLocation"
            }
        }
    }
}
catch {
    Set-RemediationError "Microsoft Store remediation failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# 9. UWP frameworks remediation
# ---------------------------------------------------------
Write-Log "Re-registering UWP frameworks"
$frameworkPatterns = @(
    "Microsoft.NET.Native.Runtime*",
    "Microsoft.NET.Native.Framework*",
    "Microsoft.VCLibs*",
    "Microsoft.UI.Xaml*"
)

foreach ($frameworkPattern in $frameworkPatterns) {
    try {
        $frameworkPackages = Get-AppxPackage -AllUsers -Name $frameworkPattern -ErrorAction SilentlyContinue

        if ($null -eq $frameworkPackages) {
            Set-RemediationError "Required UWP framework package was not found: $frameworkPattern"
            continue
        }

        foreach ($frameworkPackage in $frameworkPackages) {
            if (-not [string]::IsNullOrWhiteSpace($frameworkPackage.InstallLocation)) {
                $frameworkManifest = Join-Path -Path $frameworkPackage.InstallLocation -ChildPath "AppxManifest.xml"

                if (Test-Path -Path $frameworkManifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $frameworkManifest -ErrorAction SilentlyContinue
                    Write-Log "Framework re-registration requested: $($frameworkPackage.Name)"
                }
                else {
                    Set-RemediationError "Framework manifest is missing: $($frameworkPackage.Name)"
                }
            }
            else {
                Set-RemediationError "Framework package has no InstallLocation: $($frameworkPackage.Name)"
            }
        }
    }
    catch {
        Set-RemediationError "Framework remediation failed for ${frameworkPattern}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------
# 10. Final result
# ---------------------------------------------------------
Write-Log "===== Remediation finished; no reboot was forced ====="

if ($ErrorFound) {
    Write-Log "Script exit code: 1"
    exit 1
}

Write-Log "Script exit code: 0"
exit 0
