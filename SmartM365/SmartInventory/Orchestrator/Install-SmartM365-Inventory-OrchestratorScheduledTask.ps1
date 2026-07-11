#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
Installs or removes the SmartM365 Inventory Orchestrator scheduled task.

.DESCRIPTION
Registers an unattended scheduled task that runs the SmartM365 Inventory Orchestrator
directly with PowerShell 7 under a dedicated service account. The service-account
password is requested through Get-Credential and is never accepted as a command-line
parameter.

The task starts five minutes after server startup. A daily trigger at midnight repeats
every 30 minutes for one day so the orchestrator is restarted after its normal lifetime
recycle. Concurrent instances are ignored, missed starts run as soon as possible, and
Task Scheduler retries a failed start three times at one-minute intervals.

.PARAMETER Tenant
Tenant profile passed to the orchestrator. Valid values are prod and test.

.PARAMETER ServiceAccount
Dedicated service account used to run the task. SYSTEM and LocalSystem are refused.
The parameter is required when installing and ignored when uninstalling.

.PARAMETER TaskName
Scheduled task name. Defaults to SmartM365 Inventory Orchestrator - <Tenant>.

.PARAMETER StartNow
Starts the task immediately after successful registration.

.PARAMETER Uninstall
Removes the scheduled task instead of installing it. No credential is requested.

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -ServiceAccount 'CONTOSO\svc-smartm365' -StartNow

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -Uninstall

.VERSION
1.0.0

.REQUIREMENTS
    Windows PowerShell 5.1 or PowerShell 7 running as administrator.
    PowerShell 7 installed on the local computer.
    A dedicated service account with access to the repository, tenant configuration,
    certificates, data roots, and log roots required by the orchestrated jobs.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateSet('prod', 'test')]
    [string]$Tenant = 'prod',
    [string]$ServiceAccount = '',
    [string]$TaskName = '',
    [switch]$StartNow,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$taskPath = '\'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$orchestratorPath = Join-Path $scriptDirectory 'SmartM365-Inventory-Orchestrator.ps1'
$jobsManifestTemplatePath = Join-Path $scriptDirectory 'Orchestrator-Jobs.json.template'
$orchestratorConfigTemplatePath = Join-Path $scriptDirectory 'SmartM365-Inventory-Orchestrator.local.json.template'
$smartM365Root = Split-Path -Parent (Split-Path -Parent $scriptDirectory)
$tenantContextPath = Join-Path $smartM365Root 'Config\SmartM365-TenantContext.ps1'

function Write-InstallerMessage {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Output ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-AccountSid {
    param([Parameter(Mandatory = $true)][string]$AccountName)
    try {
        $account = [Security.Principal.NTAccount]::new($AccountName)
        return $account.Translate([Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        throw "Service account '$AccountName' could not be resolved. Verify the account name and domain connectivity. $($_.Exception.Message)"
    }
}

function Get-PowerShell7Path {
    $candidates = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $candidates.Add((Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))
    }
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) {
        $candidates.Add((Join-Path $programFilesX86 'PowerShell\7\pwsh.exe'))
    }
    $pathCommand = Get-Command pwsh.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $pathCommand -and -not [string]::IsNullOrWhiteSpace($pathCommand.Source)) {
        $candidates.Add($pathCommand.Source)
    }
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            continue
        }
        $resolvedCandidate = (Resolve-Path -LiteralPath $candidate).Path
        $majorVersionText = & $resolvedCandidate -NoProfile -NonInteractive -Command '$PSVersionTable.PSVersion.Major' 2>$null
        if ($LASTEXITCODE -eq 0 -and [int]$majorVersionText -ge 7) {
            return $resolvedCandidate
        }
    }
    throw 'PowerShell 7 (pwsh.exe) was not found. Install PowerShell 7 before registering the orchestrator task.'
}

function ConvertTo-QuotedArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

if (-not (Test-IsAdministrator)) {
    throw 'Administrator rights are required to install or remove the scheduled task.'
}
if ([string]::IsNullOrWhiteSpace($TaskName)) {
    $TaskName = "SmartM365 Inventory Orchestrator - $Tenant"
}
if ($TaskName.IndexOf('\') -ge 0) {
    throw 'TaskName must not contain a backslash. The task is registered in the Task Scheduler root folder.'
}

if ($Uninstall) {
    if ($StartNow) {
        throw 'StartNow cannot be combined with Uninstall.'
    }
    $existingTask = Get-ScheduledTask -TaskPath $taskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -eq $existingTask) {
        Write-InstallerMessage "Scheduled task does not exist: $TaskName"
        return
    }
    if ($PSCmdlet.ShouldProcess("$taskPath$TaskName", 'Unregister scheduled task')) {
        Unregister-ScheduledTask -TaskPath $taskPath -TaskName $TaskName -Confirm:$false
        Write-InstallerMessage "Scheduled task removed: $TaskName"
    }
    return
}

if ([string]::IsNullOrWhiteSpace($ServiceAccount)) {
    throw 'ServiceAccount is required when installing the scheduled task.'
}
$serviceAccountSid = Resolve-AccountSid -AccountName $ServiceAccount
$normalizedServiceAccount = $ServiceAccount.Trim().ToLowerInvariant()
$forbiddenServiceAccounts = @('system', 'localsystem', 'nt authority\system', 's-1-5-18')
if ($normalizedServiceAccount -in $forbiddenServiceAccounts) {
    throw 'SYSTEM and LocalSystem are refused. Use a dedicated service account because the repository scripts may be modifiable outside the Administrators group.'
}
if ($serviceAccountSid -eq 'S-1-5-18') {
    throw 'SYSTEM and LocalSystem are refused. Use a dedicated service account because the repository scripts may be modifiable outside the Administrators group.'
}

$requiredFiles = @(
    $orchestratorPath,
    $jobsManifestTemplatePath,
    $orchestratorConfigTemplatePath,
    $tenantContextPath
)
foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Required file not found: $requiredFile"
    }
}

$powerShell7Path = Get-PowerShell7Path
$actionArguments = @(
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    (ConvertTo-QuotedArgument -Value $orchestratorPath),
    '-Tenant',
    $Tenant,
    '-Connect'
) -join ' '
$actionParameters = @{
    Execute = $powerShell7Path
    Argument = $actionArguments
    WorkingDirectory = $scriptDirectory
}
$action = New-ScheduledTaskAction @actionParameters

$startupTrigger = New-ScheduledTaskTrigger -AtStartup
$startupTrigger.Delay = 'PT5M'
$dailyTrigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::Today)
$repetitionProperties = @{
    Interval = 'PT30M'
    Duration = 'P1D'
    StopAtDurationEnd = $false
}
$dailyTrigger.Repetition = New-CimInstance -Namespace 'Root\Microsoft\Windows\TaskScheduler' -ClassName 'MSFT_TaskRepetitionPattern' -ClientOnly -Property $repetitionProperties

$settingsParameters = @{
    MultipleInstances = 'IgnoreNew'
    StartWhenAvailable = $true
    RestartCount = 3
    RestartInterval = New-TimeSpan -Minutes 1
    ExecutionTimeLimit = [timespan]::Zero
}
$settings = New-ScheduledTaskSettingsSet @settingsParameters

$description = "Runs SmartM365 Inventory Orchestrator for tenant '$Tenant' under a dedicated service account."
$credential = $null
$passwordPointer = [IntPtr]::Zero
$plainTextPassword = $null

try {
    if (-not $PSCmdlet.ShouldProcess("$taskPath$TaskName", "Register scheduled task as $ServiceAccount")) {
        return
    }
    $credentialParameters = @{
        UserName = $ServiceAccount
        Message = "Enter the password for the SmartM365 Orchestrator service account ($ServiceAccount)."
    }
    $credential = Get-Credential @credentialParameters
    if ($null -eq $credential) {
        throw 'Credential entry was cancelled.'
    }
    $credentialSid = Resolve-AccountSid -AccountName $credential.UserName
    if ($credentialSid -ne $serviceAccountSid) {
        throw "The entered credential belongs to '$($credential.UserName)', not '$ServiceAccount'."
    }
    $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($credential.Password)
    $plainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)

    $registrationParameters = @{
        TaskPath = $taskPath
        TaskName = $TaskName
        Action = $action
        Trigger = @($startupTrigger, $dailyTrigger)
        Settings = $settings
        Description = $description
        User = $ServiceAccount
        Password = $plainTextPassword
        RunLevel = 'Highest'
        Force = $true
    }
    Register-ScheduledTask @registrationParameters | Out-Null

    Write-InstallerMessage "Scheduled task registered: $TaskName"
    Write-InstallerMessage "Run account: $ServiceAccount"
    Write-InstallerMessage "Action: $powerShell7Path $actionArguments"
    if ($StartNow) {
        Start-ScheduledTask -TaskPath $taskPath -TaskName $TaskName
        Write-InstallerMessage "Scheduled task started: $TaskName"
    }
}
finally {
    $plainTextPassword = $null
    $credential = $null
    if ($passwordPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
    }
}
