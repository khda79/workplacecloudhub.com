# Name: SmartM365-WindowsUpdate-Service-Health-Remediation.ps1
# Version: 1.0
# Description: Remediates Windows Update related services health issues

$ErrorActionPreference = "Stop"

$RemediationName = "Remediate-WindowsUpdate-Service-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$RemediationName"

if (-not (Test-Path -Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"
$ErrorFound = $false

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Set-RemediationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Repair-ServiceHealth {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$StartupType
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Set-RemediationError "Service is missing: ${Name}"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
            Write-Log "Service startup type changed: ${Name} -> ${StartupType}"
        }

        if ($service.Status -in @("StopPending", "PausePending")) {
            Write-Log "Service is in pending state, waiting before retry: ${Name} Status=$($service.Status)"
            Start-Sleep -Seconds 10
            $service.Refresh()
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-Log "Service start requested: ${Name}"
        }
        else {
            Write-Log "Service already running: ${Name}"
        }
    }
    catch {
        Set-RemediationError "Failed to repair service ${Name}: $($_.Exception.Message)"
    }
}

function Trigger-WindowsUpdateScan {
    try {
        $usoClientPath = Join-Path -Path $env:WINDIR -ChildPath "System32\UsoClient.exe"

        if (Test-Path -LiteralPath $usoClientPath -PathType Leaf) {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-Log "Windows Update scan triggered"
        }
        else {
            Write-Log "UsoClient.exe not found; scan trigger skipped"
        }
    }
    catch {
        Set-RemediationError "Failed to trigger Windows Update scan: $($_.Exception.Message)"
    }
}

try {
    Write-Log "===== Windows Update service remediation started ====="

    $serviceRemediationPlan = @(
        @{
            Name = "BITS"
            StartupType = "Manual"
        },
        @{
            Name = "DoSvc"
            StartupType = "Manual"
        },
        @{
            Name = "wuauserv"
            StartupType = "Manual"
        },
        @{
            Name = "UsoSvc"
            StartupType = "Automatic"
        }
    )

    foreach ($serviceItem in $serviceRemediationPlan) {
        Repair-ServiceHealth -Name $serviceItem.Name -StartupType $serviceItem.StartupType
    }

    Start-Sleep -Seconds 5

    Trigger-WindowsUpdateScan

    Write-Log "===== Windows Update service remediation finished ====="

    if ($ErrorFound) {
        Write-Log "Status=CompletedWithErrors"
        exit 1
    }

    Write-Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"

    try {
        Set-RemediationError $_.Exception.Message
    }
    catch { }

    exit 1
}

