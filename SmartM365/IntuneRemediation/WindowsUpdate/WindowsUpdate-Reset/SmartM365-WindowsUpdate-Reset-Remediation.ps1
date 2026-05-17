# Name: SmartM365-WindowsUpdate-Reset-Remediation.ps1
# Version: 1.0
# Description: Resets Windows Update components and removes legacy WSUS policies that block WUfB or Autopatch.

$ErrorActionPreference = "Stop"

$ScriptName = "Remediate-WindowsUpdate-Reset"
$Scenario = "WindowsUpdate-Reset"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$ScriptName.log"
$WuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$WuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

function Stop-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service -and $service.Status -ne "Stopped") {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-Log "ServiceStopRequested=$Name"
    }
}

function Start-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service) {
        try {
            $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

            if ($serviceCim -and $serviceCim.StartMode -eq "Disabled") {
                Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
                Write-Log "ServiceStartupTypeChanged=$Name StartupType=Manual"
            }
        }
        catch {
            Write-Log "ServiceStartupTypeCheckFailed=$Name Message=$($_.Exception.Message)"
        }

        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-Log "ServiceStartRequested=$Name"
    }
}

function Remove-RegistryValueSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

    if ($item -and $item.PSObject.Properties.Name -contains $Name) {
        Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue
        Write-Log "RegistryValueRemoved=$Path\$Name"
    }
}

function Rename-FolderSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "FolderNotFound=$Path"
        return
    }

    $suffix = Get-Date -Format "yyyyMMddHHmmss"
    $destination = "{0}.SmartM365.old.{1}" -f $Path, $suffix

    try {
        Rename-Item -LiteralPath $Path -NewName (Split-Path -Leaf $destination) -Force -ErrorAction Stop
        Write-Log "FolderRenamed=$Path Destination=$destination"
    }
    catch {
        Write-Log "FolderRenameFailed=$Path Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $usoClient = Join-Path -Path $env:windir -ChildPath "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $usoClient) {
        Start-Process -FilePath $usoClient -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-Log "UsoClient=$Action Status=Triggered"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log "RemediationStarted"

    foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate", "DoNotConnectToWindowsUpdateInternetLocations", "DisableWindowsUpdateAccess", "SetDisableUXWUAccess")) {
        Remove-RegistryValueSafe -Path $WuPolicyPath -Name $name
    }

    foreach ($name in @("UseWUServer", "NoAutoUpdate", "AUOptions")) {
        Remove-RegistryValueSafe -Path $WuAuPolicyPath -Name $name
    }

    foreach ($service in @("bits", "wuauserv", "dosvc", "cryptsvc")) {
        Stop-ServiceSafe -Name $service
    }

    Rename-FolderSafe -Path (Join-Path -Path $env:windir -ChildPath "SoftwareDistribution")
    Rename-FolderSafe -Path (Join-Path -Path $env:windir -ChildPath "System32\catroot2")

    foreach ($service in @("cryptsvc", "dosvc", "wuauserv", "bits")) {
        Start-ServiceSafe -Name $service
    }

    Invoke-UsoClientSafe -Action "RefreshSettings"
    Invoke-UsoClientSafe -Action "StartScan"

    Write-Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}
