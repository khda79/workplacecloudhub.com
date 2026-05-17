# Name: SmartM365-WindowsUpdate-Service-Refresh-Remediation.ps1
# Version: 1.0
# Description: Refreshes Windows Update services and triggers a fresh update scan.

$ErrorActionPreference = "Stop"

$Scenario = "Service-Refresh"
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
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
    else {
        Write-Log "ServiceNotFound=$Name"
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

    foreach ($service in @("bits", "wuauserv", "dosvc", "cryptsvc")) {
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
