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

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStopRequested=$Name"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStartRequested=$Name"
    }
}

function Invoke-UsoClient {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $uso) {
        Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

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
        Write-SmartM365Log "RecentSetupActivityDetected=True CleanupSkipped=True"
        exit 0
    }

    foreach ($service in @("bits", "wuauserv", "dosvc")) {
        Invoke-ServiceStopSafe -Name $service
    }

    foreach ($path in $UpgradePaths) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "UpgradeFolderRemoved=$path"
        }
        else {
            Write-SmartM365Log "UpgradeFolderNotFound=$path"
        }
    }

    foreach ($service in @("dosvc", "wuauserv", "bits")) {
        Invoke-ServiceStartSafe -Name $service
    }

    Invoke-UsoClient -Action "RefreshSettings"
    Invoke-UsoClient -Action "StartScan"
    Invoke-UsoClient -Action "StartDownload"

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}
