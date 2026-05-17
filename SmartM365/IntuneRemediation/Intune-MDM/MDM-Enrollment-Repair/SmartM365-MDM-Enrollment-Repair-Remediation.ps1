# Name: SmartM365-MDM-Enrollment-Repair-Remediation.ps1
# Version: 1.0
# Description: Repairs local device registration and MDM enrollment signals without deleting the current identity.

$ErrorActionPreference = "Stop"

$Scenario = "MDM-Enrollment-Repair"
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

function Invoke-ProcessSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$Arguments,

        [int]$TimeoutSeconds = 120
    )

    if (-not (Test-Path -LiteralPath $FilePath)) {
        Write-Log "ProcessSkipped FilePath=$FilePath Reason=NotFound"
        return
    }

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden -ErrorAction Stop
        $completed = $process.WaitForExit($TimeoutSeconds * 1000)

        if (-not $completed) {
            Write-Log "ProcessTimeout FilePath=$FilePath Arguments=$Arguments"
            return
        }

        Write-Log "ProcessCompleted FilePath=$FilePath Arguments=$Arguments ExitCode=$($process.ExitCode)"
    }
    catch {
        Write-Log "ProcessFailed FilePath=$FilePath Arguments=$Arguments Message=$($_.Exception.Message)"
    }
}

function Start-TaskSafe {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Task
    )

    try {
        if ($Task.State -eq "Disabled") {
            Enable-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction SilentlyContinue | Out-Null
            Write-Log "TaskEnabled=$($Task.TaskPath)$($Task.TaskName)"
        }

        Start-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction Stop
        Write-Log "TaskStarted=$($Task.TaskPath)$($Task.TaskName)"
    }
    catch {
        Write-Log "TaskStartFailed=$($Task.TaskPath)$($Task.TaskName) Message=$($_.Exception.Message)"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log "RemediationStarted"

    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\gpupdate.exe") -Arguments "/target:computer /force" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\DeviceEnroller.exe") -Arguments "/c /AutoEnrollMDM" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\dsregcmd.exe") -Arguments "/join" -TimeoutSeconds 180
    Invoke-ProcessSafe -FilePath (Join-Path $env:SystemRoot "System32\dsregcmd.exe") -Arguments "/refreshprt" -TimeoutSeconds 60

    Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object {
            $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" -and
            ($_.TaskName -like "*PushLaunch*" -or $_.TaskName -like "*Schedule*")
        } |
        ForEach-Object { Start-TaskSafe -Task $_ }

    $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue

    if ($imeService) {
        Restart-Service -Name "IntuneManagementExtension" -Force -ErrorAction SilentlyContinue
        Write-Log "ServiceRestartRequested=IntuneManagementExtension"
    }
    else {
        Write-Log "ServiceNotFound=IntuneManagementExtension"
    }

    Write-Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}
