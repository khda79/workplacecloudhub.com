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
Before registering a task, the installer imports or verifies the profile's application
certificate in the service account's personal store using the exported PFX and its password.
The installer grants batch logon when needed and verifies it; its Windows profile is loaded for the import.

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
Tenant profile key passed to the orchestrator: 1 to 64 letters, digits, underscores
or hyphens, starting with a letter or digit.

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

.PARAMETER UninstallInteractive
Guides removal of the local task, including a clean orchestrator stop.

.PARAMETER CreateOnlyDisabled
Creates a disabled task without replacing an existing or legacy task, starting
the task or changing shortcuts. The password is requested interactively.

.PARAMETER ExpectedServiceSid
Requires the account to resolve to the reviewed SID before registration.

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -ServiceAccount 'CONTOSO\svc-smartm365' -StartNow

.EXAMPLE
./Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 -Tenant prod -Uninstall

.VERSION
1.3.3

.REQUIREMENTS
    Windows PowerShell 5.1 or PowerShell 7. The script requests UAC elevation when needed.
    PowerShell 7 installed on the local computer.
    A dedicated service account with access to the repository, tenant configuration,
    certificates, data roots, and log roots required by the orchestrated jobs.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,63}\z')]
    [string]$Tenant = 'prod',
    [string]$ServiceAccount = '',
    [string]$TaskName = '',
    [switch]$StartNow,
    [switch]$Uninstall,
    [switch]$Interactive,
    [switch]$UninstallInteractive,
    [switch]$CreateOnlyDisabled,
    [string]$ExpectedServiceSid = ''
)
if ($UninstallInteractive -and ($Interactive -or $Uninstall -or $StartNow -or $CreateOnlyDisabled -or $ServiceAccount -or $TaskName -or $PSBoundParameters.ContainsKey('Tenant'))) {throw 'UninstallInteractive cannot be combined with task, tenant or installation options.'}
if ($CreateOnlyDisabled -and ($StartNow -or $Uninstall -or $Interactive)) {throw 'CreateOnlyDisabled cannot be combined with StartNow, Uninstall or Interactive.'}
if ($CreateOnlyDisabled -and ($ExpectedServiceSid -notmatch '^S-1-(?:\d+-)+\d+$' -or $ExpectedServiceSid -eq 'S-1-5-18')) {throw 'A reviewed service account SID is required.'}
if ($CreateOnlyDisabled -and $ServiceAccount.Trim().EndsWith('$')) {throw 'Managed service accounts are not supported in this mode.'}

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

function Get-InstallerHostLaunchMode {
    param([Parameter(Mandatory)][bool]$IsAdministrator,[Parameter(Mandatory)][int]$PowerShellMajor)
    if (-not $IsAdministrator) {return 'Elevate'}
    if ($PowerShellMajor -lt 7) {return 'PowerShell7'}
    return 'Current'
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

function Get-InstallerAvailableProfiles {
    param([Parameter(Mandatory = $true)][string]$Root)
    $folder = Join-Path $Root 'Config\Tenants'
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw 'The tenant configuration folder is unavailable.' }
    $groups = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $folder -File -ErrorAction Stop)) {
        if ($file.Name -match '(?i)\.mail\.local\.json(?:\.txt)?$') { continue }
        if ($file.Name -notmatch '(?i)^([a-z0-9][a-z0-9_-]{0,63})\.local\.json(?:\.txt)?$') { continue }
        $key = $Matches[1]
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [Collections.Generic.List[object]]::new() }
        $groups[$key].Add($file)
    }
    foreach ($key in @($groups.Keys | Sort-Object)) {
        $files = $groups[$key]
        if ($files.Count -ne 1) { Write-InstallerMessage "Profile '$key' has conflicting files and is not available."; continue }
        try {
            $doc = Get-Content -LiteralPath $files[0].FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($doc -isnot [pscustomobject] -or [string]$doc.ProfileKey -ine $key) { throw 'ProfileKey does not match the file name.' }
            $key
        } catch { Write-InstallerMessage "Profile '$key' is invalid or unreadable and is not available." }
    }
}

function Read-InstallerProfileDocument {
    param([Parameter(Mandatory = $true)][string]$Root,[Parameter(Mandatory = $true)][string]$ProfileKey)
    if ($ProfileKey -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,63}\z') {throw 'Invalid tenant profile key.'}
    $path=Join-Path $Root ("Config\Tenants\{0}.local.json" -f $ProfileKey)
    $validate={param($document) if ($document -isnot [pscustomobject]) {throw 'Tenant configuration must be a JSON object.'}}
    $document=(Read-SmartM365JsonDocument -Path $path -Validate $validate -ErrorAction Stop).Document
    if ([string]$document.ProfileKey -ine $ProfileKey) {throw 'Tenant profile does not match the selected profile key.'}
    return $document
}

function Read-InstallerProfile {
    param([Parameter(Mandatory = $true)][string]$Root)
    $profiles = @(Get-InstallerAvailableProfiles -Root $Root)
    if ($profiles.Count -eq 0) { throw 'No valid tenant profiles were found. Create a profile before installing or removing a task.' }
    Write-InstallerMessage 'Available tenant profiles:'
    for ($i = 0; $i -lt $profiles.Count; $i++) { Write-InstallerMessage ('{0} - {1}' -f ($i + 1), $profiles[$i]) }
    while ($true) {
        $choice = Read-Host ('[{0}] Select a profile number [1-{1}]' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $profiles.Count)
        $number = 0
        if ([int]::TryParse($choice, [ref]$number) -and $number -ge 1 -and $number -le $profiles.Count) {
            $selected = $profiles[$number - 1]
            if (@(Get-InstallerAvailableProfiles -Root $Root | Where-Object { $_ -ieq $selected }).Count -ne 1) { throw 'The selected profile is no longer available. Restart the installer.' }
            return $selected
        }
        Write-InstallerMessage 'Invalid choice. Enter a number from the profile list.'
    }
}

function Read-InstallerServiceAccount {
    param([Parameter(Mandatory = $true)][string]$Root,[Parameter(Mandatory = $true)][string]$ProfileKey)
    $profile=Read-InstallerProfileDocument -Root $Root -ProfileKey $ProfileKey
    $configured=if ($null -ne $profile.PSObject.Properties['ServiceAccount']) {[string]$profile.ServiceAccount} else {''}
    if ($configured -and ($configured -notmatch '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._+-]+|[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+)$' -or $configured -in @('SYSTEM','LocalSystem','NT AUTHORITY\SYSTEM'))) {throw 'The configured service account is invalid. Correct the profile before continuing.'}
    while ($true) {
        $prompt=if ($configured) {'Dedicated service account (DOMAIN\user or user@domain) [{0}]' -f $configured} else {'Dedicated service account (DOMAIN\user or user@domain)'}
        $candidate=Read-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$prompt)
        if ([string]::IsNullOrWhiteSpace($candidate)) {$candidate=$configured}
        if ($candidate -match '^(?:[A-Za-z0-9._-]+\\[A-Za-z0-9._+-]+|[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+)$' -and $candidate -notin @('SYSTEM','LocalSystem','NT AUTHORITY\SYSTEM')) {return $candidate}
        Write-InstallerMessage 'Enter a standard service account in DOMAIN\user or user@domain format.'
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

$script:InteractiveMode = $Interactive -or $UninstallInteractive -or $PSBoundParameters.Count -eq 0
$interactiveTaskEnabled = $true

trap {
    Write-InstallerMessage ('ERROR: {0}' -f $_.Exception.Message)
    $script:InstallerCompletionStatus = 'Failed'
    Wait-InteractiveClose
    Complete-InstallerExecution -Status 'Failed'
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

$hostLaunchMode=Get-InstallerHostLaunchMode -IsAdministrator (Test-IsAdministrator) -PowerShellMajor $PSVersionTable.PSVersion.Major
if ($hostLaunchMode -ne 'Current') {
    $elevationPowerShellPath = Get-PowerShell7Path

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

    if ($UninstallInteractive) {
        $elevationArguments += '-UninstallInteractive'
        if ($WhatIfPreference) { $elevationArguments += '-WhatIf' }
    }
    elseif ($script:InteractiveMode) {
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
        if ($CreateOnlyDisabled) { $elevationArguments += @('-CreateOnlyDisabled','-ExpectedServiceSid',$ExpectedServiceSid) }
        if ($WhatIfPreference) { $elevationArguments += '-WhatIf' }
    }

    $elevationWorkingDirectory = $scriptDirectory
    if ($scriptDirectory.StartsWith('\\', [StringComparison]::Ordinal)) {
        $elevationWorkingDirectory = $env:SystemRoot
    }

    $launchOptions=@{FilePath=$elevationPowerShellPath;ArgumentList=($elevationArguments -join ' ');WorkingDirectory=$elevationWorkingDirectory;Wait=$true;PassThru=$true}
    if ($hostLaunchMode -eq 'Elevate') {
        Write-InstallerMessage 'Administrator rights are required. Requesting elevation through UAC.'
        $launchOptions.Verb='RunAs'
    } else {
        Write-InstallerMessage 'Restarting the installer with PowerShell 7.'
    }
    $elevationProcess = Start-Process @launchOptions
    $script:InstallerCompletionStatus = if ($elevationProcess.ExitCode -eq 0) { 'Auto' } else { 'Failed' }
    Complete-InstallerExecution
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=($elevationProcess.ExitCode); exit $smartM365RunGuard.ExitCode
}

if ($script:InteractiveMode) {
    Write-InstallerMessage 'SmartM365 Inventory Orchestrator scheduled-task installer'
    if ($UninstallInteractive) { $Uninstall = $true }
    else {
        Write-InstallerMessage '1 - Install or update the scheduled task'
        Write-InstallerMessage '2 - Uninstall the scheduled task'
        $actionChoice = Read-InstallerChoice -Prompt 'Select an action [1-2]' -AllowedValues @('1', '2')
        $Uninstall = $actionChoice -eq '2'
    }

    $Tenant = Read-InstallerProfile -Root $smartM365Root

    if ($Uninstall) {
        $prompt = if ($UninstallInteractive) { "Disable the task, stop the local orchestrator and remove SmartM365 Inventory Orchestrator - $Tenant on $env:COMPUTERNAME?" } else { "Remove SmartM365 Inventory Orchestrator - $Tenant?" }
        if (-not (Read-InstallerYesNo -Prompt $prompt)) {
            Write-InstallerMessage 'Uninstall cancelled.'
            Wait-InteractiveClose
            Complete-InstallerExecution -Status 'CompletedWithWarnings'
            $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
        }
        $StartNow = $false
        $ServiceAccount = ''
    }
    else {
        $ServiceAccount=Read-InstallerServiceAccount -Root $smartM365Root -ProfileKey $Tenant
        $interactiveTaskEnabled = Read-InstallerYesNo -Prompt 'Enable the scheduled task after installation?'
        $StartNow = $false
    }
}

if (@(Get-InstallerAvailableProfiles -Root $smartM365Root | Where-Object { $_ -ieq $Tenant }).Count -ne 1) {
    throw 'The selected tenant profile is unavailable or invalid. Create or repair the profile before continuing.'
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

    if ($UninstallInteractive) {
        if (-not $PSCmdlet.ShouldProcess("$env:COMPUTERNAME / $Tenant", 'Disable task and stop local orchestrator')) {
            Write-InstallerMessage 'Uninstall was not performed.'
            Wait-InteractiveClose
            Complete-InstallerExecution -Status 'CompletedWithWarnings'
            $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
        }
        foreach ($candidateTaskPath in @($taskPath, $legacyTaskPath)) {
            $existingTask = Get-ScheduledTask -TaskPath $candidateTaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
            if ($null -ne $existingTask) {
                Disable-ScheduledTask -TaskPath $candidateTaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
                Write-InstallerMessage "Scheduled task disabled: $candidateTaskPath$TaskName"
            }
        }
        Write-InstallerMessage "Stopping the local orchestrator for tenant $Tenant."
        & (Get-PowerShell7Path) -NoProfile -ExecutionPolicy Bypass -File $orchestratorPath -Tenant $Tenant -Stop -StopTimeoutSeconds 180
        if ($LASTEXITCODE -ne 0) { throw "Orchestrator stop failed (exit code $LASTEXITCODE). Task removal cancelled; any disabled task remains disabled." }
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
if ($CreateOnlyDisabled -and $serviceAccountSid -ne $ExpectedServiceSid) {throw 'The service account no longer matches the reviewed SID.'}
if ($CreateOnlyDisabled) {
    $knownTasks=@(Get-ScheduledTask -ErrorAction Stop)
    foreach ($candidateTaskPath in @($taskPath,$legacyTaskPath)) {
        if (@($knownTasks | Where-Object {$_.TaskName -ieq $TaskName -and $_.TaskPath -ieq $candidateTaskPath}).Count) {throw 'An existing standard or legacy task requires review; no replacement was requested.'}
    }
}

$requiredFiles = @(
    $orchestratorPath,
    $jobsManifestTemplatePath,
    $orchestratorConfigTemplatePath,
    (Join-Path $scriptDirectory 'SmartM365-OrchestratorServerRegistration.ps1'),
    (Join-Path $scriptDirectory 'SmartM365.Orchestrator.Management.psm1'),
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
$settingsParameters.Disable = $true
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
    if ($PSVersionTable.PSVersion.Major -lt 7) {throw 'PowerShell 7 is required to import the application certificate into the service account profile.'}
    $serviceCertificateHelper=Join-Path $scriptDirectory 'SmartM365-ServiceCertificateImport.ps1'
    if (-not (Test-Path -LiteralPath $serviceCertificateHelper -PathType Leaf)) {throw 'The service certificate import helper is missing.'}
    . $serviceCertificateHelper
    $profileDocument=Read-InstallerProfileDocument -Root $smartM365Root -ProfileKey $Tenant
    $certificateThumb=[string]$profileDocument.Thumbprint
    if ($certificateThumb -notmatch '^[0-9A-Fa-f]{40}$' -or $certificateThumb -match '^0+$' -or [string]$profileDocument.Thumb -ine $certificateThumb) {throw 'The selected profile does not contain a matching application certificate thumbprint.'}
    $pfxFiles=@(Get-SiServiceCertificatePfx -Root $smartM365Root -Thumbprint $certificateThumb)
    if ($pfxFiles.Count -eq 1) {$pfxFile=$pfxFiles[0]} else {
        Write-InstallerMessage 'Multiple PFX exports match the application certificate:'
        for ($i=0;$i -lt $pfxFiles.Count;$i++) {Write-InstallerMessage ('{0} - {1}' -f ($i+1),$pfxFiles[$i].Name)}
        $allowed=@(1..$pfxFiles.Count | ForEach-Object {[string]$_})
        $selected=Read-InstallerChoice -Prompt ('Select a PFX number [1-{0}]' -f $pfxFiles.Count) -AllowedValues $allowed
        $pfxFile=$pfxFiles[[int]$selected-1]
    }
    $pfxPassword=Read-Host 'Enter the application PFX password' -AsSecureString
    if ($null -eq $pfxPassword -or $pfxPassword.Length -eq 0) {throw 'The PFX password is required.'}
    try {
        $batchHelper=Join-Path $scriptDirectory 'SmartM365-BatchLogonRight.ps1'
        if (-not (Test-Path -LiteralPath $batchHelper -PathType Leaf)) {throw 'The batch logon rights helper is missing.'}
        . $batchHelper
        if ((Resolve-AccountSid -AccountName $ServiceAccount) -ne $serviceAccountSid) {throw 'The service account SID changed; no local right was granted.'}
        Write-InstallerMessage 'Checking batch logon for the service account...'
        $batchResult=Ensure-SiBatchLogonRight -Credential $credential -ExpectedSid $serviceAccountSid
        if ($batchResult.Added) {Write-InstallerMessage 'Batch logon right added and verified for the service account.'} else {Write-InstallerMessage 'Batch logon is already effective for the service account.'}
        Write-InstallerMessage 'Checking local interactive logon for the service account...'
        $interactiveResult=Ensure-SiInteractiveLogonRight -Credential $credential -ExpectedSid $serviceAccountSid
        if ($interactiveResult.Added) {Write-InstallerMessage 'Local interactive logon right added and verified for the service account.'} else {Write-InstallerMessage 'Local interactive logon is already effective for the service account.'}
        Write-InstallerMessage 'Importing the application certificate into the service account profile...'
        $certificateResult=Invoke-SiServiceCertificateImport -PfxPath $pfxFile.FullName -PfxPassword $pfxPassword -Credential $credential -ExpectedSid $serviceAccountSid -Thumbprint $certificateThumb -PowerShell $powerShell7Path
        Write-InstallerMessage ('Application certificate verified in service account profile: '+$certificateResult.Thumbprint)
    } finally {$pfxPassword.Dispose();$pfxPassword=$null}
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
        Force = (-not $CreateOnlyDisabled)
    }
    Register-ScheduledTask @registrationParameters | Out-Null

    $observedTask=Get-ScheduledTask -TaskPath $taskPath -TaskName $TaskName -ErrorAction Stop
    $observedActions=@($observedTask.Actions);$observedTriggers=@($observedTask.Triggers)
    $observedSid=Resolve-AccountSid -AccountName ([string]$observedTask.Principal.UserId)
    if ($observedTask.Settings.Enabled -ne $false -or $observedSid -ne $serviceAccountSid -or $observedActions.Count -ne 1 -or [string]$observedActions[0].Execute -ine $powerShell7Path -or ([string]$observedActions[0].Arguments).Trim() -ine $actionArguments -or $observedTriggers.Count -ne 2) {
        throw 'Task may have been created, but disabled state, identity or action could not be confirmed. Inspect Task Scheduler before retrying.'
    }

    Write-InstallerMessage "Scheduled task registered: $taskPath$TaskName"
    Write-InstallerMessage 'Scheduled task is disabled while the shared server registration is checked.'
    try {
        $effectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $scriptDirectory
        Import-Module (Join-Path $smartM365Root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.87' -ErrorAction Stop
        Import-Module (Join-Path $scriptDirectory 'SmartM365.Orchestrator.Management.psm1') -Force -ErrorAction Stop
        . (Join-Path $scriptDirectory 'SmartM365-OrchestratorServerRegistration.ps1')
        function Get-ConfigValue { param($Config,[string]$Name,$DefaultValue) Get-SiOrchestratorConfigValue -Config $Config -EffectiveConfig $effectiveConfig -Name $Name -DefaultValue $DefaultValue }
        function Resolve-ConfigTokens { param($Value) Resolve-SiOrchestratorConfigTokens -Value $Value -EffectiveConfig $effectiveConfig }
        $localConfigPath = Resolve-SmartM365JsonConfigurationPath -Path (Join-Path $scriptDirectory 'SmartM365-Inventory-Orchestrator.local.json')
        if (-not (Test-Path -LiteralPath $localConfigPath -PathType Leaf)) { $localConfigPath = $orchestratorConfigTemplatePath }
        $localConfig = Get-Content -LiteralPath $localConfigPath -Raw | ConvertFrom-Json -Depth 100
        $dataFolder = Resolve-ConfigTokens -Value (Get-ConfigValue -Config $localConfig -Name 'OrchestratorDataFolderPath' -DefaultValue (Join-Path $scriptDirectory 'Output'))
        if ((Split-Path -Path $dataFolder -Leaf) -eq $env:COMPUTERNAME) { $dataFolder = Split-Path -Path $dataFolder -Parent }
        $jobsPath = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $scriptDirectory 'Orchestrator-Jobs.json')
        $selectedJobs = Get-SmartM365JsonReadPath $jobsPath -Optional
        if ($selectedJobs) { $jobsPath = $selectedJobs } else { $jobsPath = $jobsManifestTemplatePath }
        $bootstrapCluster = [pscustomobject][ordered]@{
            SchemaVersion=1
            ExpectedOrchestratorServers=@(Get-ConfigValue -Config $localConfig -Name 'ExpectedOrchestratorServers' -DefaultValue @())
            ElectionWeightsByServer=Get-ConfigValue -Config $localConfig -Name 'ElectionWeightsByServer' -DefaultValue ([pscustomobject]@{})
            ServerJobPolicies=Get-ConfigValue -Config $localConfig -Name 'ServerJobPolicies' -DefaultValue ([pscustomobject]@{})
            PeerMonitoringEnabled=[bool](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringEnabled' -DefaultValue $true)
            PeerJobMonitoringEnabled=[bool](Get-ConfigValue -Config $localConfig -Name 'PeerJobMonitoringEnabled' -DefaultValue $true)
            PeerMonitoringCheckIntervalSeconds=[int](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringCheckIntervalSeconds' -DefaultValue 60)
            PeerHeartbeatStaleMinutes=[int](Get-ConfigValue -Config $localConfig -Name 'PeerHeartbeatStaleMinutes' -DefaultValue 5)
            PeerMonitoringConfirmationChecks=[int](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringConfirmationChecks' -DefaultValue 2)
            PeerJobStartGraceMinutes=[int](Get-ConfigValue -Config $localConfig -Name 'PeerJobStartGraceMinutes' -DefaultValue 15)
            PeerAlertReminderMinutes=[int](Get-ConfigValue -Config $localConfig -Name 'PeerAlertReminderMinutes' -DefaultValue 240)
            PeerAlertMailRetryMinutes=[int](Get-ConfigValue -Config $localConfig -Name 'PeerAlertMailRetryMinutes' -DefaultValue 15)
            PeerRecoveryEmailEnabled=[bool](Get-ConfigValue -Config $localConfig -Name 'PeerRecoveryEmailEnabled' -DefaultValue $true)
        }
        $serverRegistration = Register-SmartM365OrchestratorServer -SharedDataFolderPath ([IO.Path]::GetFullPath($dataFolder)) -BootstrapJobsPath $jobsPath -BootstrapClusterDocument $bootstrapCluster -Server $env:COMPUTERNAME
        Write-InstallerMessage ("Shared server registration: {0}; state={1}; added={2}. New servers remain inactive until reviewed." -f $serverRegistration.Server,$serverRegistration.State,$serverRegistration.Added)
    }
    catch { throw "The task was registered, but server registration failed: $($_.Exception.Message). Inspect the task and shared configuration before retrying." }
    if (-not $CreateOnlyDisabled -and (-not $script:InteractiveMode -or $interactiveTaskEnabled)) {
        Enable-ScheduledTask -TaskPath $taskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
        Write-InstallerMessage 'Scheduled task is enabled; its triggers can start it automatically.'
    } else { Write-InstallerMessage 'Scheduled task remains disabled; it will not run until enabled.' }
    $legacyTask = Get-ScheduledTask -TaskPath $legacyTaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $CreateOnlyDisabled -and $null -ne $legacyTask -and $PSCmdlet.ShouldProcess("$legacyTaskPath$TaskName", 'Remove legacy root scheduled task')) {
        Unregister-ScheduledTask -TaskPath $legacyTaskPath -TaskName $TaskName -Confirm:$false
        Write-InstallerMessage "Legacy root scheduled task removed: $legacyTaskPath$TaskName"
    }

    Write-InstallerMessage "Run account: $ServiceAccount"
    Write-InstallerMessage "Action: $powerShell7Path $actionArguments"
    # A missing shortcut does not invalidate the task registration: warn and continue.
    if (-not $CreateOnlyDisabled) {
    foreach ($guiShortcutPath in $guiShortcutPaths) {
        try {
            Set-OrchestratorGuiShortcut -ShortcutPath $guiShortcutPath -PowerShell7Path $powerShell7Path -GuiPath (Join-Path -Path $scriptDirectory -ChildPath 'SmartM365-Inventory-Orchestrator-GUI.ps1') -TenantName $Tenant -IconPath (Join-Path -Path $scriptDirectory -ChildPath 'WorkplaceCloudHub.ico') -WorkingDirectory $taskWorkingDirectory -LocalIconDirectory (Join-Path -Path $env:ProgramData -ChildPath 'WorkplaceCloudHub\SmartM365')
            Write-InstallerMessage "GUI shortcut created (no console window): $guiShortcutPath"
        }
        catch {
            Write-InstallerMessage "WARNING: GUI shortcut not created: $guiShortcutPath - $($_.Exception.Message)"
        }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBLWrN4KX2cBIQl
# Uw97yPaNnF8cudzaVEwXhp3z4f73qKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHXSlcDCycJlH4p+kMA5TnrMQt97hMtHFhz3HVD0snIIMA0GCSqG
# SIb3DQEBAQUABIIBgEnVUKcUmln9hFm01b3UIlbdKP6enUy96r8L+MQHqQUGc+N6
# wTrjXgNzuQYUlWIOYgTe44Ndhx08C+WMfvytX2fdMV0g3LA/EspecEpdfO/OAgwN
# qeb6i6Wxfj8F+kszTY01MifnBDgiELcUPxEvwH0cb9/TRRod4ZHKVE5ktGtn3DQ0
# GMXZvOrEXDzHG45+zoVLPdYEze1mbMpoCU51seD7qcS5EGWm5EecLC8fMofc5Kef
# n5WRh9pQdCpti7FS7gcFoFU0OeDBpBiGa8H5XyBDuiyAdspCnhigFWt4Se6TEnYA
# R8Lm5sqX44joT/fnlzRWDTSAR43OC6hUye3/l+ZKE7oWLRK4x9pAWUK9e6ketYgJ
# BqAjyxSJ+XhIzUzMX0He9YkRlK5ybBH+U1h0XnoAu3pb7Oc+r3C8yM1aUgQpoeAi
# m+bWhwuIZAVjxeft8nHQwYdg/HrYeKMB18P8v/71Cf3eiH8BGNesqxxVqlQLokjf
# H+nEQoEUP123Fqu42qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTAwOTAw
# MDlaMC8GCSqGSIb3DQEJBDEiBCDt3b5mE8aUqr5F5/b5fOAtGHiWMvS2D5zqTVg0
# iuftqTANBgkqhkiG9w0BAQEFAASCAgAKKOJhBMhPGPiq55nZL/GZvArtgBYJbjFZ
# aAaAKAVRV23kydd4mjOuSwQCVL/q9nWaGQNmmDUrlrY6KpHpU6BleZZFmM55CZU4
# IhTC5fR2PxDCHv0kB4dEwnvdgiBxoxyX6KP7qR8dkxvyf4P+3T2nPFZA7iB66xF7
# qh5K77L5wWOCp9/GAFCCBxgv45XD/gU40Y34FacoDjDw0LDbgbw9eGo8jDL+hHOm
# hUvgAIXkTZ7Wjh8iDgI/AVlfJF2sXyOJobnhfqIaDWTGfc7hE5aTZxEaDFjfDD35
# p+II8hZQRzCpqLKU6E36B7TjCibAU/v9AnFT35NYX0jLoHSbiQ6uYnzPzcnvBhD9
# o1bTXXuhhEdqbJSnsvqI86+fUIRfHBX5jFOCIdXIxt4rdFLuFRTVJwI8Cw9cZrYm
# YzaBt05fyLfkShu1EjM6rRxEXq+jWcB3BaPY9gsZAZ0kq014mVDAnK4ypD9naE+I
# 9fcQeg3Zs0GMtqaQwAQXh+qmT28qysoM8PhanIFLTbFMPc8QGTv1jVUC/7pBy9Fh
# 7kvZOPI6tP2TD/Ph94tZBVEDiImeKo+1B5WjgofSh2XIrytQB1GUcxn9yir5TH9G
# s2iNyZ4Mm9bOtp4lZtL0b5Pel6wv7zwmVbAngrg1pSaAYDev8ilg5wRLrFtMblwg
# OIdRDwkyxw==
# SIG # End signature block
