<#
    Name: SmartM365-WindowsUpdate-Service-And-Scan-Health-Remediation.ps1
    Version: 1.0
    Description: Consolidated remediation for Windows Update service health, settings refresh, and scan trigger.
#>

[CmdletBinding()]
param(
    [int]$PostScanWaitSeconds = 30
)

$ErrorActionPreference = "Stop"
$Scenario = "Service-And-Scan-Health"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"
$ErrorFound = $false

function Write-SmartM365Log {
    param([Parameter(Mandatory = $true)][string]$Message)

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Add-RemediationError {
    param([Parameter(Mandatory = $true)][string]$Message)

    Write-SmartM365Log "ERROR: $Message"
    $script:ErrorFound = $true
}

function Repair-WindowsUpdateService {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$StartupType
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Add-RemediationError "ServiceMissing=$Name"
            return
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType $StartupType -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartupTypeChanged=$Name StartupType=$StartupType"
        }

        if ($service.Status -in @("StopPending", "PausePending")) {
            Write-SmartM365Log "ServicePending=$Name Status=$($service.Status)"
            Start-Sleep -Seconds 10
            $service.Refresh()
        }

        if ($service.Status -eq "Running") {
            Restart-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceRestartRequested=$Name"
        }
        else {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStartRequested=$Name"
        }
    }
    catch {
        Add-RemediationError "ServiceRepairFailed=$Name Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientAction {
    param([Parameter(Mandatory = $true)][string]$Action)

    try {
        $usoClient = Join-Path -Path $env:SystemRoot -ChildPath "System32\UsoClient.exe"

        if (-not (Test-Path -LiteralPath $usoClient -PathType Leaf)) {
            Add-RemediationError "UsoClientMissing"
            return
        }

        Start-Process -FilePath $usoClient -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
    catch {
        Add-RemediationError "UsoClientFailed=$Action Message=$($_.Exception.Message)"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $servicePlan = @(
        @{ Name = "bits"; StartupType = "Manual" },
        @{ Name = "wuauserv"; StartupType = "Manual" },
        @{ Name = "dosvc"; StartupType = "Manual" },
        @{ Name = "cryptsvc"; StartupType = "Manual" },
        @{ Name = "UsoSvc"; StartupType = "Automatic" }
    )

    foreach ($serviceItem in $servicePlan) {
        Repair-WindowsUpdateService -Name $serviceItem.Name -StartupType $serviceItem.StartupType
    }

    Invoke-UsoClientAction -Action "RefreshSettings"
    Invoke-UsoClientAction -Action "StartScan"

    if ($PostScanWaitSeconds -gt 0) {
        Start-Sleep -Seconds $PostScanWaitSeconds
    }

    $lastEvent = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 1 -ErrorAction SilentlyContinue
    if ($lastEvent) {
        Write-SmartM365Log ("LastWUEvent={0:s} Id={1}" -f $lastEvent.TimeCreated, $lastEvent.Id)
    }

    if ($ErrorFound) {
        Write-SmartM365Log "Status=CompletedWithErrors"
        exit 1
    }

    Write-SmartM365Log "Status=Completed"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Add-RemediationError $_.Exception.Message } catch { Write-Output "ErrorDuringErrorHandling=$($_.Exception.Message)" }
    exit 1
}
