# Name: SmartM365-WindowsUpdate-SoftwareDistribution-Remediation.ps1
# Version: 1.0
# Description: Safely resets the Windows Update download cache and triggers a new scan without forcing a reboot

$ErrorActionPreference = "Stop"

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Cache-Health'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Remediate-WindowsUpdate-SoftwareDistribution.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStopped=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyStopped=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStopFailed=$Name Message=$($_.Exception.Message)"
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
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStarted=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyRunning=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStartFailed=$Name Message=$($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Windows Update remediation started ====="

    $softwareDistributionPath = "C:\Windows\SoftwareDistribution"
    $downloadPath = Join-Path -Path $softwareDistributionPath -ChildPath "Download"

    # Stop Windows Update related services
    foreach ($serviceName in @("wuauserv", "bits", "usosvc")) {
        Invoke-ServiceStopSafe -Name $serviceName
    }

    Start-Sleep -Seconds 3

    # Clean only Download content instead of deleting the entire SoftwareDistribution folder
    if (Test-Path -Path $downloadPath) {
        Write-SmartM365Log "Cleanup=Start Path=$downloadPath"

        try {
            Get-ChildItem -Path $downloadPath -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-SmartM365Log "CleanupItemSkipped=$($_.FullName)"
                }
            }

            Write-SmartM365Log "Cleanup=Completed Path=$downloadPath"
        }
        catch {
            Write-SmartM365Log "Cleanup=Partial Path=$downloadPath Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "Cleanup=DownloadFolderNotFound"
    }

    # Restart services
    foreach ($serviceName in @("bits", "usosvc", "wuauserv")) {
        Invoke-ServiceStartSafe -Name $serviceName
    }

    Start-Sleep -Seconds 5

    # Trigger Windows Update scan
    $usoClientPath = Join-Path -Path $env:windir -ChildPath "System32\UsoClient.exe"

    if (Test-Path -Path $usoClientPath) {
        try {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-SmartM365Log "WindowsUpdateScan=Triggered"
        }
        catch {
            Write-SmartM365Log "WindowsUpdateScan=Failed Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "UsoClientNotFound"
    }

    Write-SmartM365Log "===== Windows Update remediation completed ====="
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
