# Name: SmartM365-WindowsUpdate-Policy-Blockers-Remediation.ps1
# Version: 1.0
# Description: Removes WSUS, Windows Update, and WUfB policy values that block cloud-managed update flows.

$ErrorActionPreference = "Stop"

$Scenario = "Policy-Blockers"
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$WuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$WuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
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

function Restart-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service) {
        Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue

        if ($service.Status -eq "Running") {
            Restart-Service -Name $Name -Force -ErrorAction SilentlyContinue
        }
        else {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
        }

        Write-Log "ServiceRefreshed=$Name"
    }
}

function Start-UsoClient {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $uso) {
        Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
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

    foreach ($service in @("wuauserv", "bits", "dosvc")) {
        Restart-ServiceSafe -Name $service
    }

    Start-UsoClient -Action "RefreshSettings"
    Start-UsoClient -Action "StartScan"

    Write-Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}
