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
every five minutes for one day so the orchestrator is restarted after its normal lifetime
recycle. Concurrent instances are ignored, missed starts run as soon as possible, and
Task Scheduler retries a failed start three times at one-minute intervals.

The installation also creates the shortcut "SmartM365 Orchestrator - <Tenant>" to the
orchestrator GUI in the Start menu folder "WorkplaceCloudHub" and on the all-users
(public) desktop. It starts PowerShell 7 hidden (no console window); uninstalling removes
both. Its WorkplaceCloudHub icon is copied to %ProgramData%\WorkplaceCloudHub\SmartM365
because Windows does not display an icon read from the UNC code share.

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
1.2.2

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

# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
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
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) {
    throw "SmartM365 tenant context not found: $tenantContextPath"
}
. $tenantContextPath
$script:InstallerStartedAt = Get-Date
$global:SmartM365ScriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$global:SmartM365ExecutionStartTime = $script:InstallerStartedAt
$installerLogFolder = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath 'SmartM365\Logs\Orchestrator'
if (-not (Test-Path -LiteralPath $installerLogFolder)) { New-Item -ItemType Directory -Path $installerLogFolder -Force | Out-Null }
$global:LogTextFile = Join-Path -Path $installerLogFolder -ChildPath ('{0}_{1}.log' -f $global:SmartM365ScriptName, $script:InstallerStartedAt.ToString('yyyyMMdd_HHmmss'))
$script:InstallerCompletionStatus = 'Auto'

Write-SmartM365StartupBanner

function Write-InstallerMessage {
    param([Parameter(Mandatory = $true)][string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Microsoft.PowerShell.Utility\Write-Host $line
    Add-Content -LiteralPath $global:LogTextFile -Value $line -Encoding UTF8
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

function Set-OrchestratorGuiShortcut {
    # Shortcut to the orchestrator GUI without a console window: hidden PowerShell 7, started minimized.
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$PowerShell7Path,
        [Parameter(Mandatory = $true)][string]$GuiPath,
        [Parameter(Mandatory = $true)][string]$TenantName,
        [string]$IconPath = '',
        [string]$WorkingDirectory = '',
        [string]$LocalIconDirectory = ''
    )
    New-Item -ItemType Directory -Path (Split-Path -Path $ShortcutPath -Parent) -Force | Out-Null
    # The Start menu does not render an icon read from a UNC code share: use a local copy when possible.
    if (-not [string]::IsNullOrWhiteSpace($IconPath) -and -not [string]::IsNullOrWhiteSpace($LocalIconDirectory) -and (Test-Path -LiteralPath $IconPath -PathType Leaf)) {
        try {
            New-Item -ItemType Directory -Path $LocalIconDirectory -Force | Out-Null
            $localIconPath = Join-Path -Path $LocalIconDirectory -ChildPath (Split-Path -Path $IconPath -Leaf)
            Copy-Item -LiteralPath $IconPath -Destination $localIconPath -Force
            $IconPath = $localIconPath
        }
        catch {
            Write-Warning "Local copy of the shortcut icon failed, the shared icon is used: $($_.Exception.Message)"
        }
    }
    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $PowerShell7Path
        $shortcut.Arguments = '-NoProfile -NoLogo -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File {0} -Tenant {1}' -f (ConvertTo-QuotedArgument -Value $GuiPath), $TenantName
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $shortcut.WorkingDirectory = $WorkingDirectory }
        if (-not [string]::IsNullOrWhiteSpace($IconPath) -and (Test-Path -LiteralPath $IconPath -PathType Leaf)) { $shortcut.IconLocation = '{0},0' -f $IconPath }
        $shortcut.WindowStyle = 7
        $shortcut.Description = "SmartM365 Orchestrator console for tenant '$TenantName'."
        $shortcut.Save()
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
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
function Complete-InstallerExecution {
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = $script:InstallerCompletionStatus
    )
    Write-SmartM365CompletionBanner -Status $Status -ScriptName $global:SmartM365ScriptName -StartedAt $script:InstallerStartedAt -LogPath $global:LogTextFile
}

$script:InteractiveMode = $Interactive -or $PSBoundParameters.Count -eq 0

trap {
    Write-InstallerMessage ('ERROR: {0}' -f $_.Exception.Message)
    $script:InstallerCompletionStatus = 'Failed'
    Wait-InteractiveClose
    Complete-InstallerExecution -Status 'Failed'
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
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
    $script:InstallerCompletionStatus = if ($elevationProcess.ExitCode -eq 0) { 'Auto' } else { 'Failed' }
    Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=($elevationProcess.ExitCode); exit $smartM365RunGuard.ExitCode
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
            Complete-InstallerExecution -Status 'CompletedWithWarnings'
            $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
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
$guiShortcutName = 'SmartM365 Orchestrator - {0}.lnk' -f $Tenant
$guiShortcutPaths = @(
    (Join-Path -Path $env:ProgramData -ChildPath ("Microsoft\Windows\Start Menu\Programs\WorkplaceCloudHub\{0}" -f $guiShortcutName)),
    (Join-Path -Path ([Environment]::GetFolderPath('CommonDesktopDirectory')) -ChildPath $guiShortcutName)
)

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
    foreach ($guiShortcutPath in $guiShortcutPaths) {
        if ((Test-Path -LiteralPath $guiShortcutPath -PathType Leaf) -and $PSCmdlet.ShouldProcess($guiShortcutPath, 'Remove GUI shortcut')) {
            Remove-Item -LiteralPath $guiShortcutPath -Force
            Write-InstallerMessage "GUI shortcut removed: $guiShortcutPath"
        }
    }
    Wait-InteractiveClose
    Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
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
    Interval = 'PT5M'
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
        Complete-InstallerExecution -Status 'CompletedWithWarnings'
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
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
    # A missing shortcut does not invalidate the task registration: warn and continue.
    foreach ($guiShortcutPath in $guiShortcutPaths) {
        try {
            Set-OrchestratorGuiShortcut -ShortcutPath $guiShortcutPath -PowerShell7Path $powerShell7Path -GuiPath (Join-Path -Path $scriptDirectory -ChildPath 'SmartM365-Inventory-Orchestrator-GUI.ps1') -TenantName $Tenant -IconPath (Join-Path -Path $scriptDirectory -ChildPath 'WorkplaceCloudHub.ico') -WorkingDirectory $taskWorkingDirectory -LocalIconDirectory (Join-Path -Path $env:ProgramData -ChildPath 'WorkplaceCloudHub\SmartM365')
            Write-InstallerMessage "GUI shortcut created (no console window): $guiShortcutPath"
        }
        catch {
            Write-InstallerMessage "WARNING: GUI shortcut not created: $guiShortcutPath - $($_.Exception.Message)"
        }
    }
    if ($StartNow) {
        Start-ScheduledTask -TaskPath $taskPath -TaskName $TaskName
        Write-InstallerMessage "Scheduled task started: $TaskName"
    }

    $smartM365RunGuard.ReachedTerminal=$true
}
finally {
    $smartM365RunGuard.InCleanup=$true

    $plainTextPassword = $null
    $credential = $null
    if ($passwordPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
    }
}

Wait-InteractiveClose

Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAnSEQu3gTo/Gob
# wHwTeZGd3TqybiNbWOugkEHIcTIAnKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPYouWBDX4YCOIUMPVawOaTIwm+6t7oVUwniDpwXuA2xMA0GCSqG
# SIb3DQEBAQUABIIBgDQnQPz5Kx2cXB9fcqOqoiV6+WPmpJU7KnJ6FRvixg/BYMXi
# EJPzYwmM4PPU+snXfo8QYFwwIPB6XOwnHt3SyubF7vLC4eNqqc3cndjq8xPrdN6c
# +B4aAVTSvUhDrdjvJPJia7PzQo5Vp7+rTtlP9SvOIUNeT5JU6JiBh2FYd17WLiDo
# 7DK+HzKxfH1YTcJXzAgwddUG0/LyoJDv1qGE+Vm9+Bzn9zGKnvVCg7E34ojvfNDu
# A5sv6GdEV3OxzbVJxInM4Pd9Vl0kT9z355I5Jo7kiYmu7BYaUeLvMqMqcTHWYsXv
# Uwxjd+S6ZZcI35AQ2qt1tjV+OzwNMLxJ7mOrhZ5n+t6svOExk1RmMGb4IY6cWrBj
# 9a+EytlsBsVPEKH1d9A9EdW9EECV996SMp2Brh0d1L20CWNVCJt31aKUS1PTiqN5
# 9zWzcRltn5mvNHo2rqkgeflm7rEVNOgujHkcW68mk5J6QttjaK4vyTzv/pQwOWnn
# gCiQqxEhovy1SeLEZaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMDU5
# NDBaMC8GCSqGSIb3DQEJBDEiBCDxLa1IbQP0LznPtj3AAitzIKBTCD5tQk0pjXji
# Zx5KLjANBgkqhkiG9w0BAQEFAASCAgCDFA+EkpnQxAuJ52iydjyhLhV1wHFSlxez
# XszGLcyBikcgmrbmWmlqodge3bvYSAsZOG8B42tUKnBgPWtV+QtypMIKIUj2Xoqo
# AmQ/n4k7LnUDitDK9Son4zRb1tQpQ/oAix9HOChUqfaVi9W9+Ir7L/xSTIOe8yAg
# tpiZUYVcWV18XsNXdH+k+ZtY8a1RzD1+GyS2mp3GRQWBY9PEscoZ0w1Y6A9d7dEy
# UWDyywIWai98UgBa/e1aSAxP8ejnqd2dUmp4nWoN91dBGMc7nL7dabXwJEgOAVet
# opI53e7Rj2zpqgoZtAKt4JqjhUHSDWQbcxYmqhvI1z6t+idz6CX7RemJLmQb/wzc
# 47/ClyjRZCGzKAUCaijPclouLSwPakcg9TUD8OZwRTAU/s2HiVmwLTJlMMIM+8eB
# /JaJITS6q8Y2SwrHKrIRXbLRBYEkAQJ3S9ZXl34OOwWgXcjhu3EtkKmpal36NcB4
# zi5ZnLkpJwtTnSewPLYAPvy4jPY3I8ZqWSpxiX6FaqS3651Nd5uRAfoncJ70NPcu
# ViMbnsfu4xrZA0R6ngUD4XuMq7mU2KoxR00B8g4sYu5Z4K7uzx2fWU6MwUmxTdmm
# HwjNZvTgtKE2DuTh0ynBSLp4opU7C39+qHBHRUNnihY84geC9Vx2alZt0djzBAJe
# iM7JeE3qUw==
# SIG # End signature block
