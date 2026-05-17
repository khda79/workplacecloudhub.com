# Name: SmartM365-Upgrade-Staging-Health-Remediation.ps1
# Version: 1.0
# Description: Removes stale Windows upgrade staging folders only when no recent setup activity is detected.

$ErrorActionPreference = "Stop"

$Scenario = "Upgrade-Staging-Health"
$RecentSetupActivityHours = 6
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$UpgradePaths = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')
$SetupIndicators = @(
    "C:\Windows\Panther\setupact.log",
    "C:\Windows\Panther\setuperr.log",
    'C:\$WINDOWS.~BT\Sources\Panther\setupact.log',
    'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
)

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Stop-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-Log "ServiceStopRequested=$Name"
    }
}

function Start-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-Log "ServiceStartRequested=$Name"
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

    $recentSetupActivity = $false

    foreach ($indicator in $SetupIndicators) {
        if (Test-Path -LiteralPath $indicator) {
            $lastWrite = (Get-Item -LiteralPath $indicator -ErrorAction SilentlyContinue).LastWriteTime

            if ($lastWrite -and ((Get-Date) - $lastWrite).TotalHours -le $RecentSetupActivityHours) {
                $recentSetupActivity = $true
            }
        }
    }

    if ($recentSetupActivity) {
        Write-Log "RecentSetupActivityDetected=True CleanupSkipped=True"
        exit 0
    }

    foreach ($service in @("bits", "wuauserv", "dosvc")) {
        Stop-ServiceSafe -Name $service
    }

    foreach ($path in $UpgradePaths) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "UpgradeFolderRemoved=$path"
        }
        else {
            Write-Log "UpgradeFolderNotFound=$path"
        }
    }

    foreach ($service in @("dosvc", "wuauserv", "bits")) {
        Start-ServiceSafe -Name $service
    }

    Start-UsoClient -Action "RefreshSettings"
    Start-UsoClient -Action "StartScan"
    Start-UsoClient -Action "StartDownload"

    Write-Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}
