#Requires -Version 5.1

<#
.SYNOPSIS
Installs or removes the SmartM365 Inventory Orchestrator scheduled task.

.DESCRIPTION
Registers an unattended scheduled task that runs the SmartM365 Inventory Orchestrator
directly with PowerShell 7 under a dedicated service account. The service-account
password is requested through Get-Credential and is never accepted as a command-line
parameter. When launched without parameters, the script requests elevation and guides the
user through tenant, action, service-account, and immediate-start choices.

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

.PARAMETER Interactive
Forces the guided interactive workflow. This is used by the CMD launcher.

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -ServiceAccount 'CONTOSO\svc-smartm365' -StartNow

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -Uninstall

.VERSION
1.1.2

.REQUIREMENTS
    Windows PowerShell 5.1 or PowerShell 7. The script requests UAC elevation when needed.
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
    [switch]$Uninstall,
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
$taskFolderName = 'WCH'
$taskPath = '\WCH\'
$legacyTaskPath = '\'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$orchestratorPath = Join-Path $scriptDirectory 'SmartM365-Inventory-Orchestrator.ps1'
$jobsManifestTemplatePath = Join-Path $scriptDirectory 'Orchestrator-Jobs.json.template'
$orchestratorConfigTemplatePath = Join-Path $scriptDirectory 'SmartM365-Inventory-Orchestrator.local.json.template'
$smartM365Root = Split-Path -Parent (Split-Path -Parent $scriptDirectory)
$tenantContextPath = Join-Path $smartM365Root 'Config\SmartM365-TenantContext.ps1'

function Write-InstallerMessage {
    param([Parameter(Mandatory = $true)][string]$Message)
    Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
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

function Read-InstallerChoice {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][string[]]$AllowedValues
    )

    while ($true) {
        $value = Read-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        if ($value -in $AllowedValues) {
            return $value
        }
        Write-InstallerMessage ('Invalid choice. Allowed values: {0}' -f ($AllowedValues -join ', '))
    }
}

function Read-InstallerYesNo {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    while ($true) {
        $value = Read-Host ('[{0}] {1} [Y/N]' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Prompt)
        switch ($value.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default { Write-InstallerMessage 'Invalid choice. Enter Y or N.' }
        }
    }
}

function Wait-InteractiveClose {
    if ($script:InteractiveMode) {
        [void](Read-Host ('[{0}] Press Enter to close' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    }
}


function Initialize-WchScheduledTaskFolder {
    param([Parameter(Mandatory = $true)][string]$FolderName)

    $taskService = $null
    $rootFolder = $null
    $targetFolder = $null
    try {
        $taskService = New-Object -ComObject 'Schedule.Service'
        $taskService.Connect()
        $normalizedFolderPath = '\{0}' -f $FolderName.Trim('\')
        try {
            $targetFolder = $taskService.GetFolder($normalizedFolderPath)
        }
        catch {
            $rootFolder = $taskService.GetFolder('\')
            $targetFolder = $rootFolder.CreateFolder($FolderName.Trim('\'))
            Write-InstallerMessage "Task Scheduler folder created: $normalizedFolderPath"
        }
    }
    finally {
        foreach ($comObject in @($targetFolder, $rootFolder, $taskService)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }
    }
}
$script:InteractiveMode = $Interactive -or $PSBoundParameters.Count -eq 0

trap {
    Write-InstallerMessage ('ERROR: {0}' -f $_.Exception.Message)
    Wait-InteractiveClose
    exit 1
}

if (-not (Test-IsAdministrator)) {
    $windowsPowerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $elevationPowerShellPath = $null
    try {
        $elevationPowerShellPath = Get-PowerShell7Path
    }
    catch {
        if (-not (Test-Path -LiteralPath $windowsPowerShellPath -PathType Leaf)) {
            throw "No trusted PowerShell host was found for UAC elevation."
        }
        $elevationPowerShellPath = $windowsPowerShellPath
        Write-InstallerMessage 'PowerShell 7 was not found in Program Files; using Windows PowerShell for UAC elevation.'
    }

    foreach ($unsafeValue in @($ServiceAccount, $TaskName)) {
        if (-not [string]::IsNullOrWhiteSpace($unsafeValue) -and $unsafeValue -match '[\r\n"]') {
            throw 'ServiceAccount and TaskName must not contain quotes or line breaks.'
        }
    }

    $elevationArguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (ConvertTo-QuotedArgument -Value $MyInvocation.MyCommand.Path)
    )

    if ($script:InteractiveMode) {
        $elevationArguments += '-Interactive'
    }
    else {
        $elevationArguments += @('-Tenant', $Tenant)
        if (-not [string]::IsNullOrWhiteSpace($ServiceAccount)) {
            $elevationArguments += @('-ServiceAccount', (ConvertTo-QuotedArgument -Value $ServiceAccount))
        }
        if (-not [string]::IsNullOrWhiteSpace($TaskName)) {
            $elevationArguments += @('-TaskName', (ConvertTo-QuotedArgument -Value $TaskName))
        }
        if ($StartNow) { $elevationArguments += '-StartNow' }
        if ($Uninstall) { $elevationArguments += '-Uninstall' }
        if ($WhatIfPreference) { $elevationArguments += '-WhatIf' }
    }

    $elevationWorkingDirectory = $scriptDirectory
    if ($scriptDirectory.StartsWith('\\', [StringComparison]::Ordinal)) {
        $elevationWorkingDirectory = $env:SystemRoot
    }

    Write-InstallerMessage 'Administrator rights are required. Requesting elevation through UAC.'
    $elevationProcess = Start-Process -FilePath $elevationPowerShellPath `
        -ArgumentList ($elevationArguments -join ' ') `
        -WorkingDirectory $elevationWorkingDirectory `
        -Verb RunAs `
        -Wait `
        -PassThru
    exit $elevationProcess.ExitCode
}

if ($script:InteractiveMode) {
    Write-InstallerMessage 'SmartM365 Inventory Orchestrator scheduled-task installer'
    Write-InstallerMessage '1 - Install or update the scheduled task'
    Write-InstallerMessage '2 - Uninstall the scheduled task'
    $actionChoice = Read-InstallerChoice -Prompt 'Select an action [1-2]' -AllowedValues @('1', '2')
    $Uninstall = $actionChoice -eq '2'

    Write-InstallerMessage '1 - prod'
    Write-InstallerMessage '2 - test'
    $tenantChoice = Read-InstallerChoice -Prompt 'Select the tenant profile [1-2]' -AllowedValues @('1', '2')
    $Tenant = if ($tenantChoice -eq '1') { 'prod' } else { 'test' }

    if ($Uninstall) {
        if (-not (Read-InstallerYesNo -Prompt "Remove SmartM365 Inventory Orchestrator - $Tenant?")) {
            Write-InstallerMessage 'Uninstall cancelled.'
            Wait-InteractiveClose
            return
        }
        $StartNow = $false
        $ServiceAccount = ''
    }
    else {
        while ($true) {
            $candidateAccount = Read-Host ('[{0}] Dedicated service account (DOMAIN\user or user@domain)' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            if ([string]::IsNullOrWhiteSpace($candidateAccount)) {
                Write-InstallerMessage 'The service account is required.'
                continue
            }
            if ($candidateAccount -notmatch '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._$+-]+|[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+)$') {
                Write-InstallerMessage 'Invalid account format. Use DOMAIN\user or user@domain without spaces.'
                continue
            }
            $ServiceAccount = $candidateAccount
            break
        }
        $StartNow = Read-InstallerYesNo -Prompt 'Start the task immediately after installation?'
    }
}

if ([string]::IsNullOrWhiteSpace($TaskName)) {
    $TaskName = "SmartM365 Inventory Orchestrator - $Tenant"
}
if ($TaskName.IndexOfAny([char[]]'\/:*?"<>|[]') -ge 0) {
    throw 'TaskName contains a character that is not allowed in a scheduled-task name.'
}

if ($Uninstall) {
    if ($StartNow) {
        throw 'StartNow cannot be combined with Uninstall.'
    }

    $taskWasFound = $false
    foreach ($candidateTaskPath in @($taskPath, $legacyTaskPath)) {
        $existingTask = Get-ScheduledTask -TaskPath $candidateTaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($null -eq $existingTask) {
            continue
        }
        $taskWasFound = $true
        if ($PSCmdlet.ShouldProcess("$candidateTaskPath$TaskName", 'Unregister scheduled task')) {
            Unregister-ScheduledTask -TaskPath $candidateTaskPath -TaskName $TaskName -Confirm:$false
            Write-InstallerMessage "Scheduled task removed: $candidateTaskPath$TaskName"
        }
    }

    if (-not $taskWasFound) {
        Write-InstallerMessage "Scheduled task does not exist in $taskPath or the legacy root folder: $TaskName"
    }
    Wait-InteractiveClose
    return
}

if ([string]::IsNullOrWhiteSpace($ServiceAccount)) {
    throw 'ServiceAccount is required when installing the scheduled task.'
}
$normalizedServiceAccount = $ServiceAccount.Trim().ToLowerInvariant()
$forbiddenServiceAccounts = @('system', 'localsystem', 'nt authority\system', 's-1-5-18')
if ($normalizedServiceAccount -in $forbiddenServiceAccounts) {
    throw 'SYSTEM and LocalSystem are refused. Use a dedicated service account because the repository scripts may be modifiable outside the Administrators group.'
}
$serviceAccountSid = Resolve-AccountSid -AccountName $ServiceAccount
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
$taskWorkingDirectory = $scriptDirectory
if ($scriptDirectory.StartsWith('\\', [StringComparison]::Ordinal)) {
    $taskWorkingDirectory = $env:SystemRoot
}

$actionParameters = @{
    Execute = $powerShell7Path
    Argument = $actionArguments
    WorkingDirectory = $taskWorkingDirectory
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
        Wait-InteractiveClose
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

    Initialize-WchScheduledTaskFolder -FolderName $taskFolderName

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

    Write-InstallerMessage "Scheduled task registered: $taskPath$TaskName"
    $legacyTask = Get-ScheduledTask -TaskPath $legacyTaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $legacyTask -and $PSCmdlet.ShouldProcess("$legacyTaskPath$TaskName", 'Remove legacy root scheduled task')) {
        Unregister-ScheduledTask -TaskPath $legacyTaskPath -TaskName $TaskName -Confirm:$false
        Write-InstallerMessage "Legacy root scheduled task removed: $legacyTaskPath$TaskName"
    }

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

Wait-InteractiveClose
