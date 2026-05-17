<#
    Name: SmartM365-WUfB-Identity-Binding-Remediation.ps1
    Version: 1.0
    Description: Repairs Windows Update for Business identity binding by resetting computed PolicyState and refreshing MDM/WU policy state safely.

    Intended use:
    - Microsoft Intune remediation script
    - Windows 10 / Windows 11
    - Run as 64-bit PowerShell
    - No PRT refresh
    - No forced reboot

    Logs:
    - C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-WUfB-Identity-Binding\

    Exit codes:
    0 = Remediation completed
    1 = Remediation completed with errors
#>

[CmdletBinding()]
param(
    [switch]$ExportPolicyStateReg
)

$ErrorActionPreference = "Stop"

# Relaunch in 64-bit PowerShell if Intune starts the script in 32-bit PowerShell
if ($env:PROCESSOR_ARCHITEW6432) {
    & "$env:WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe" `
        -ExecutionPolicy Bypass `
        -NoProfile `
        -File "$PSCommandPath" `
        -ExportPolicyStateReg:$ExportPolicyStateReg

    exit $LASTEXITCODE
}

$RemediationName = "Remediate-WUfB-Identity-Binding"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$RemediationName"

if (-not (Test-Path -Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

$LogPath = Join-Path -Path $LogRoot -ChildPath "$RemediationName.log"

$PolicyStatePath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
$MdmUpdatePolicyPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
$UsoClientPath = Join-Path -Path $env:SystemRoot -ChildPath "System32\UsoClient.exe"

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

function Restart-ServiceSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-Log "ServiceNotFound=${Name}"
            return
        }

        if ($service.Status -ne "Stopped") {
            Write-Log "ServiceStopRequested=${Name}"
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        }
        else {
            Write-Log "ServiceAlreadyStopped=${Name}"
        }

        Start-Sleep -Seconds 2

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue

        if ($null -ne $serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
            Write-Log "ServiceStartupTypeChanged=${Name} StartupType=Manual"
        }

        Write-Log "ServiceStartRequested=${Name}"
        Start-Service -Name $Name -ErrorAction SilentlyContinue
    }
    catch {
        Set-RemediationError "Failed to restart service ${Name}: $($_.Exception.Message)"
    }
}

function Trigger-EnterpriseMgmtPush {
    try {
        $enterpriseMgmtTasks = @(
            Get-ScheduledTask -TaskPath "\Microsoft\Windows\EnterpriseMgmt\" -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskName -eq "PushLaunch" }
        )

        if ($null -eq $enterpriseMgmtTasks -or $enterpriseMgmtTasks.Count -eq 0) {
            Write-Log "EnterpriseMgmtPushLaunch=NotFound"
            return
        }

        foreach ($task in $enterpriseMgmtTasks) {
            try {
                Start-ScheduledTask -InputObject $task -ErrorAction Stop
                Write-Log "EnterpriseMgmtPushLaunch=Triggered TaskPath=$($task.TaskPath) TaskName=$($task.TaskName)"
            }
            catch {
                Write-Log "EnterpriseMgmtPushLaunch=Failed TaskName=$($task.TaskName) Message=$($_.Exception.Message)"
            }
        }
    }
    catch {
        Set-RemediationError "Failed to enumerate EnterpriseMgmt tasks: $($_.Exception.Message)"
    }
}

function Get-RegistryPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

    if ($null -eq $item) {
        return $null
    }

    if ($item.PSObject.Properties.Name -contains $Name) {
        return $item.$Name
    }

    return $null
}

function Get-IsWUfBConfigured {
    $value = Get-RegistryPropertyValue -Path $PolicyStatePath -Name "IsWUfBConfigured"

    if ($null -eq $value) {
        return $null
    }

    try {
        return [int]$value
    }
    catch {
        return $null
    }
}

function Export-PolicyState {
    try {
        if (-not (Test-Path -Path $PolicyStatePath)) {
            Write-Log "PolicyStateExport=Skipped Reason=PolicyStateNotFound"
            return
        }

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $exportPath = Join-Path -Path $LogRoot -ChildPath "WUfB-PolicyState-$timestamp.reg"
        $regExe = Join-Path -Path $env:SystemRoot -ChildPath "System32\reg.exe"

        & $regExe export "HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState" "$exportPath" /y | Out-Null

        if (Test-Path -LiteralPath $exportPath -PathType Leaf) {
            Write-Log "PolicyStateExport=Completed Path=$exportPath"
        }
        else {
            Write-Log "PolicyStateExport=Failed Path=$exportPath"
        }
    }
    catch {
        Write-Log "PolicyStateExport=Failed Message=$($_.Exception.Message)"
    }
}

function Invoke-UsoClientSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    try {
        if (-not (Test-Path -LiteralPath $UsoClientPath -PathType Leaf)) {
            Write-Log "UsoClient=${Action} Status=UsoClientNotFound"
            return
        }

        Start-Process -FilePath $UsoClientPath -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-Log "UsoClient=${Action} Status=Triggered"
    }
    catch {
        Write-Log "UsoClient=${Action} Status=Failed Message=$($_.Exception.Message)"
    }
}

try {
    Write-Log "===== WUfB identity binding remediation started ====="

    $mdmPresentBefore = Test-Path -Path $MdmUpdatePolicyPath
    $wuConfiguredBefore = Get-IsWUfBConfigured

    Write-Log "Before_MDMWUfBPolicyPresent=$mdmPresentBefore"
    Write-Log "Before_IsWUfBConfigured=$wuConfiguredBefore"

    # Stop/restart Windows Update services around PolicyState reset
    foreach ($serviceName in @("UsoSvc", "wuauserv")) {
        Restart-ServiceSafe -Name $serviceName
    }

    if ($ExportPolicyStateReg) {
        Export-PolicyState
    }

    if (Test-Path -Path $PolicyStatePath) {
        try {
            Write-Log "PolicyStateReset=Start Path=$PolicyStatePath"
            Remove-Item -Path $PolicyStatePath -Recurse -Force -ErrorAction Stop
            Write-Log "PolicyStateReset=Completed"
        }
        catch {
            Set-RemediationError "Failed to remove PolicyState: $($_.Exception.Message)"
        }
    }
    else {
        Write-Log "PolicyStateReset=Skipped Reason=PolicyStateNotFound"
    }

    # Restart related services again after reset
    foreach ($serviceName in @("wuauserv", "UsoSvc")) {
        Restart-ServiceSafe -Name $serviceName
    }

    # Trigger MDM and Windows Update policy refresh
    Trigger-EnterpriseMgmtPush

    Start-Sleep -Seconds 20

    Invoke-UsoClientSafe -Action "RefreshSettings"
    Invoke-UsoClientSafe -Action "StartScan"

    Start-Sleep -Seconds 15

    $mdmPresentAfter = Test-Path -Path $MdmUpdatePolicyPath
    $wuConfiguredAfter = Get-IsWUfBConfigured

    Write-Log "After_MDMWUfBPolicyPresent=$mdmPresentAfter"
    Write-Log "After_IsWUfBConfigured=$wuConfiguredAfter"

    if ($mdmPresentAfter -and $wuConfiguredAfter -eq 0) {
        Write-Log "Status=CompletedButDriftStillPresent"
        Write-Log "Result=MDM WUfB policy is present but IsWUfBConfigured is still 0"
        exit 0
    }

    if ($mdmPresentAfter -and $wuConfiguredAfter -eq 1) {
        Write-Log "Status=CompletedHealthy"
        Write-Log "Result=WUfB identity binding is healthy after remediation"
        exit 0
    }

    if (-not $mdmPresentAfter) {
        Write-Log "Status=CompletedNotApplicable"
        Write-Log "Result=No MDM WUfB policy detected after remediation"
        exit 0
    }

    Write-Log "Status=CompletedInconclusive"
    Write-Log "Result=Sanity check inconclusive"
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

