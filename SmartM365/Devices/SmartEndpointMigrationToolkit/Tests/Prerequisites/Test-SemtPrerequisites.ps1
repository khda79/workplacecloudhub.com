<#
.SYNOPSIS
    Offline tests of the Prerequisites campaign and of the scheduled-task consumers.
.DESCRIPTION
    The Prerequisites endpoint functions run against an in-memory task store: creation of the
    HybridJoin and Windows11Upgrade tasks (enabled state at creation), repeat runs, audit, a task
    with another definition (updated, enabled state kept), registration failure and unreadable
    Task Scheduler. Task state parsing is checked with a localized SYSTEM account name.
    The obsolete user-context deviceenroller task is reported in audit and removed otherwise.
    Hybrid Join consumer functions: missing user-context tasks are reported; deviceenroller runs
    from SYSTEM with its real exit code; a UserLicense fault after the trigger is found; the
    next-logon task is only enabled (when up to date) or disabled.
    Automatic prerequisites: the launcher expectations match the endpoint definitions; outdated,
    missing, obsolete and unreadable tasks are classified; the worker runs the Prerequisites
    endpoint only when needed and never in audit (remote check and run stubbed).
    No scheduled task, registry or endpoint run on the computer; the only network use is one DNS
    lookup of a .invalid name.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
    Runs on Windows PowerShell 5.1 and PowerShell 7.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$packageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$prereqEndpoint = Join-Path $packageRoot 'Campaigns\Prerequisites\Scripts\SmartM365-Invoke-EndpointMigration-Prerequisites.ps1'
$hjEndpoint = Join-Path $packageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Name)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    return (@($Name | ForEach-Object {
        $functionName = $_
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        if (-not $definition) { throw "Function $functionName not found in $Path" }
        $definition.Extent.Text
    }) -join "`n`n")
}

function Get-RegionText {
    param([string]$Path, [string]$Region)
    $match = [regex]::Match([IO.File]::ReadAllText($Path), ('(?s)#region {0}.*?#endregion {0}' -f [regex]::Escape($Region)))
    if (-not $match.Success) { throw "Region '$Region' not found in $Path" }
    return $match.Value
}

try {
    . ([scriptblock]::Create((Get-RegionText -Path $prereqEndpoint -Region 'SEMT prerequisite tasks v1')))
    . ([scriptblock]::Create((Get-FunctionText -Path $prereqEndpoint -Name @('Get-PrerequisiteTaskDefinitions','Test-PrerequisiteTaskMatch','Invoke-PrerequisiteTask','Remove-PrerequisiteObsoleteTask'))))

    #region Task state parsing (Get-ScheduledTask stubbed)
    $localizedSystem = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')).Translate([System.Security.Principal.NTAccount]).Value
    function Get-ScheduledTask {
        param([string]$TaskPath, [string]$TaskName, $ErrorAction, $ErrorVariable)
        return [pscustomobject]@{
            TaskName = $TaskName; TaskPath = $TaskPath; State = 'Disabled'
            Settings = [pscustomobject]@{ Enabled = $false }
            Actions = @([pscustomobject]@{ Execute = '%windir%\System32\WindowsPowerShell\v1.0\powershell.exe'; Arguments = ' -NoProfile -File "C:\x.ps1" ' }, [pscustomobject]@{ ClassId = 'com-handler' })
            Triggers = @([pscustomobject]@{ Enabled = $true })
            Principal = [pscustomobject]@{ UserId = $script:StubUserId; GroupId = '' }
        }
    }
    $script:StubUserId = $localizedSystem
    $parsed = Get-SemtScheduledTaskState -TaskPath '\' -TaskName 'SEMT-Stub'
    Test-Step ("task state: localized SYSTEM ({0}) is read as S-1-5-18; environment variables expanded; non-exec action kept empty" -f $localizedSystem) {
        $parsed.Exists -eq $true -and -not $parsed.Enabled -and $parsed.UserSid -eq 'S-1-5-18' -and $parsed.TriggerCount -eq 1 -and @($parsed.Actions).Count -eq 2 -and @($parsed.Actions)[0] -eq ('{0}|-NoProfile -File "C:\x.ps1"' -f (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe')) -and @($parsed.Actions)[1] -eq '|'
    }
    Remove-Item -Path Function:\Get-ScheduledTask
    #endregion

    #region Prerequisites endpoint against an in-memory task store
    $script:ComputerName = 'SEMT-QUALIF-PC'
    $script:Store = @{}
    $script:Registered = New-Object System.Collections.Generic.List[string]
    $script:FailRegister = @()
    $script:Unreadable = @()
    $script:Unregistered = New-Object System.Collections.Generic.List[string]
    function Write-PrerequisitesLog { param([string]$Message, [string]$Level = 'INFO') }
    function Get-SemtScheduledTaskState {
        param([string]$TaskPath, [string]$TaskName)
        $key = $TaskPath + $TaskName
        if ($script:Unreadable -contains $key) { return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $null; Enabled = $false; Actions = @(); TriggerCount = 0; UserSid = ''; GroupSid = ''; Detail = 'Task Scheduler read failed: synthetic' } }
        if (-not $script:Store.ContainsKey($key)) { return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $false; Enabled = $false; Actions = @(); TriggerCount = 0; UserSid = ''; GroupSid = ''; Detail = 'Task not found.' } }
        return $script:Store[$key]
    }
    function Register-PrerequisiteTask {
        param($Definition, [bool]$Enabled)
        $key = $Definition.TaskPath + $Definition.TaskName
        if ($script:FailRegister -contains $key) { throw 'Access is denied. (synthetic EDR block)' }
        $script:Registered.Add($key)
        $script:Store[$key] = [pscustomobject]@{
            TaskPath = $Definition.TaskPath; TaskName = $Definition.TaskName; Exists = $true; Enabled = $Enabled
            Actions = @($Definition.Actions | ForEach-Object { '{0}|{1}' -f [Environment]::ExpandEnvironmentVariables([string]$_.Execute),[string]$_.Arguments })
            TriggerCount = $(if ($Definition.Trigger -eq 'None') { 0 } else { 1 })
            UserSid = $(if ($Definition.Principal -eq 'System') { 'S-1-5-18' } else { '' })
            GroupSid = $(if ($Definition.Principal -eq 'Interactive') { 'S-1-5-4' } else { '' })
            Detail = ('Enabled={0}' -f $Enabled)
        }
    }
    function Unregister-PrerequisiteTask { param([string]$TaskPath, [string]$TaskName) $script:Unregistered.Add($TaskPath + $TaskName); [void]$script:Store.Remove($TaskPath + $TaskName) }
    function Invoke-Scenario {
        param([switch]$Audit, [switch]$HybridJoin, [switch]$Windows11)
        $script:AuditOnly = [bool]$Audit; $AuditOnly = [bool]$Audit
        $script:HybridJoinTasks = [bool]$HybridJoin; $HybridJoinTasks = [bool]$HybridJoin
        $script:Windows11UpgradeTasks = [bool]$Windows11; $Windows11UpgradeTasks = [bool]$Windows11
        $script:Counters = [ordered]@{ TasksExpected = 0; TasksPresent = 0; TasksCreated = 0; TasksUpdated = 0; TasksMissing = 0; TasksFailed = 0 }
        $script:TaskRows = New-Object System.Collections.Generic.List[object]
        $script:Registered.Clear()
        $script:Unregistered.Clear()
        foreach ($definition in @(Get-PrerequisiteTaskDefinitions)) { Invoke-PrerequisiteTask -Definition $definition }
        if ($HybridJoinTasks) { Remove-PrerequisiteObsoleteTask -Campaign 'HybridJoin' -TaskPath '\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\' -TaskName 'SEMT-HJ-UserMdmAutoEnroll' }
        return [pscustomobject]@{ Counters = $script:Counters; Rows = @($script:TaskRows.ToArray()); Registered = @($script:Registered.ToArray()) }
    }
    $hjFolder = '\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\'
    $userTasks = @('SEMT-HJ-UserDsregStatus','SEMT-HJ-UserRefreshPrt' | ForEach-Object { $hjFolder + $_ })
    $obsoleteAutoEnrollTask = $hjFolder + 'SEMT-HJ-UserMdmAutoEnroll'
    $logonTask = $hjFolder + 'SEMT-HJ-RunUserAutoEnrollAtLogon'
    $hjRetry = '\SmartM365-SEMT-HybridJoin-RetryAfterReboot'
    $w11Retry = '\SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot'
    $deviceEnrollerAction = '{0}|/c /AutoEnrollMDM' -f (Join-Path $env:WINDIR 'System32\deviceenroller.exe')

    $audit = Invoke-Scenario -Audit -HybridJoin -Windows11
    Test-Step 'audit, no task: five tasks reported missing, nothing registered' { $audit.Counters.TasksExpected -eq 5 -and $audit.Counters.TasksMissing -eq 5 -and $audit.Registered.Count -eq 0 -and @($audit.Rows | Where-Object { $_.Result -ne 'Missing' }).Count -eq 0 }

    $first = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'repair, no task: the five tasks are created' { $first.Counters.TasksCreated -eq 5 -and $first.Counters.TasksFailed -eq 0 -and $first.Registered.Count -eq 5 }
    Test-Step 'on-demand user-context tasks are created enabled (no trigger), group INTERACTIVE; no user-context deviceenroller task' { @($userTasks | Where-Object { $script:Store[$_].Enabled -and $script:Store[$_].TriggerCount -eq 0 -and $script:Store[$_].GroupSid -eq 'S-1-5-4' }).Count -eq 2 -and -not $script:Store.ContainsKey($obsoleteAutoEnrollTask) }
    Test-Step 'next-logon and both retry-after-reboot tasks are created disabled, as SYSTEM' { @(@($logonTask,$hjRetry,$w11Retry) | Where-Object { -not $script:Store[$_].Enabled -and $script:Store[$_].UserSid -eq 'S-1-5-18' -and $script:Store[$_].TriggerCount -eq 1 }).Count -eq 3 }
    Test-Step 'retry-after-reboot tasks run the fixed runner command of the shared region' {
        $ok = $true
        foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
            $spec = Get-SemtRetryAfterRebootTaskSpec -Campaign $campaign
            if (-not (Test-SemtRetryAfterRebootTaskState -State $script:Store[$spec.TaskPath + $spec.TaskName] -Spec $spec)) { $ok = $false }
        }
        $ok
    }
    Test-Step 'dsreg status keeps its cmd.exe redirection; refreshprt starts the executable directly; the next-logon task runs deviceenroller from SYSTEM' {
        @($script:Store[$userTasks[0]].Actions)[0] -eq 'cmd.exe|/d /c %windir%\System32\dsregcmd.exe /status > %windir%\Temp\SEMT-HJ-UserDsregStatus_%USERNAME%.txt 2>&1' -and @($script:Store[$userTasks[1]].Actions)[0] -eq ('{0}|/refreshprt' -f (Join-Path $env:WINDIR 'System32\dsregcmd.exe')) -and @($script:Store[$logonTask].Actions)[0] -eq ('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$userTasks[0]) -and @($script:Store[$logonTask].Actions)[1] -eq $deviceEnrollerAction
    }
    $script:Store[$logonTask].Enabled = $true
    $script:Store[$logonTask].Actions = @(@($script:Store[$logonTask].Actions)[0], ('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$obsoleteAutoEnrollTask))
    $script:Store[$obsoleteAutoEnrollTask] = [pscustomobject]@{ TaskPath = $hjFolder; TaskName = 'SEMT-HJ-UserMdmAutoEnroll'; Exists = $true; Enabled = $true; Actions = @($deviceEnrollerAction); TriggerCount = 0; UserSid = ''; GroupSid = 'S-1-5-4'; Detail = 'Enabled=True' }
    $auditLegacy = Invoke-Scenario -Audit -HybridJoin -Windows11
    Test-Step 'audit: the obsolete user-context deviceenroller task is reported, not removed' { $script:Store.ContainsKey($obsoleteAutoEnrollTask) -and $script:Unregistered.Count -eq 0 -and @($auditLegacy.Rows | Where-Object { $_.TaskName -eq $obsoleteAutoEnrollTask -and $_.Result -eq 'Obsolete' }).Count -eq 1 -and $auditLegacy.Counters.TasksExpected -eq 5 }
    $legacy = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'an older next-logon task is updated to deviceenroller from SYSTEM and keeps its enabled state' { $legacy.Counters.TasksUpdated -eq 1 -and $legacy.Registered -contains $logonTask -and @($script:Store[$logonTask].Actions)[1] -eq $deviceEnrollerAction -and $script:Store[$logonTask].Enabled }
    Test-Step 'the obsolete user-context deviceenroller task is removed, outside the expected count' { -not $script:Store.ContainsKey($obsoleteAutoEnrollTask) -and $script:Unregistered -contains $obsoleteAutoEnrollTask -and @($legacy.Rows | Where-Object { $_.TaskName -eq $obsoleteAutoEnrollTask -and $_.Result -eq 'Removed' }).Count -eq 1 -and $legacy.Counters.TasksExpected -eq 5 -and $legacy.Counters.TasksFailed -eq 0 }

    $second = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'repeat run: every task Present, nothing registered again' { $second.Counters.TasksPresent -eq 5 -and $second.Registered.Count -eq 0 }

    $script:Store[$w11Retry] = [pscustomobject]@{ TaskPath = '\'; TaskName = 'SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot'; Exists = $true; Enabled = $true; Actions = @('powershell.exe|-NoProfile -ExecutionPolicy RemoteSigned -File "C:\Program Files\old\RetryAfterRebootRunner.ps1"'); TriggerCount = 1; UserSid = 'S-1-5-18'; GroupSid = ''; Detail = 'Enabled=True' }
    $third = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'a task with another definition is updated and keeps its enabled state' { $third.Counters.TasksUpdated -eq 1 -and $third.Registered -contains $w11Retry -and $script:Store[$w11Retry].Enabled -and @($third.Rows | Where-Object { $_.TaskName -eq $w11Retry -and $_.Before -eq 'Different' -and $_.Result -eq 'Updated' }).Count -eq 1 }
    Test-Step 'an enabled task with the expected definition is left unchanged (a campaign may be waiting on it)' { $third.Registered -notcontains $logonTask -and $script:Store[$logonTask].Enabled }

    $script:Store = @{}
    $script:FailRegister = @($hjRetry)
    $script:Unreadable = @($w11Retry)
    $fourth = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'a registration failure and an unreadable Task Scheduler are counted as failures; other tasks are still created' { $fourth.Counters.TasksFailed -eq 2 -and $fourth.Counters.TasksCreated -eq 3 -and $fourth.Registered -notcontains $w11Retry -and @($fourth.Rows | Where-Object { $_.TaskName -eq $hjRetry -and $_.Result -eq 'Failed' -and $_.Detail -match 'Access is denied' }).Count -eq 1 }
    $script:FailRegister = @()
    $script:Unreadable = @()

    $script:Store = @{}
    $w11Only = Invoke-Scenario -Windows11
    Test-Step 'Windows11Upgrade only: one task, the retry-after-reboot task' { $w11Only.Counters.TasksExpected -eq 1 -and $w11Only.Registered.Count -eq 1 -and $w11Only.Registered[0] -eq $w11Retry }
    Test-Step 'nothing selected: no definition' { @(Invoke-Scenario).Counters.TasksExpected -eq 0 }

    $endpointText = [IO.File]::ReadAllText($prereqEndpoint)
    Test-Step 'Prerequisites endpoint: the target guard is taken before any task operation, audit never registers' {
        $endpointText.IndexOf('Enter-SemtTargetOperationGuard -Campaign $script:SemtCampaign') -lt $endpointText.IndexOf('foreach ($definition in @(Get-PrerequisiteTaskDefinitions))') -and $endpointText -match "elseif \(\`$AuditOnly -or \`$beforeText -eq 'Unreadable'\)"
    }
    #endregion

    #region Hybrid Join consumers (missing tasks, SYSTEM deviceenroller, license fault)
    foreach ($name in @('Get-SemtScheduledTaskState','Register-PrerequisiteTask','Unregister-PrerequisiteTask')) { Remove-Item -Path ('Function:\' + $name) -ErrorAction SilentlyContinue }
    . ([scriptblock]::Create((Get-FunctionText -Path $hjEndpoint -Name @('Get-HybridJoinTaskState','Test-UserContextHelperTasks','Start-IntuneAutoEnrollment','Get-RecentMdmUserLicenseFailure','Enable-UserAutoEnrollAtLogonTask','Disable-UserAutoEnrollAtLogonTask'))))
    $UserStatusTaskName = $userTasks[0]; $UserRefreshPrtTaskName = $userTasks[1]; $UserNextLogonAutoEnrollTaskName = $logonTask
    $OutputDir = $env:TEMP; $ComputerName = 'SEMT-QUALIF-PC'; $RunId = 'test'; $IntuneAutoEnrollTimeoutSeconds = 300
    $script:HjStore = @{}
    $script:HjActions = @{}
    $script:EnabledCalls = New-Object System.Collections.Generic.List[string]
    $script:Processes = New-Object System.Collections.Generic.List[string]
    $script:ProcessExitCode = 0
    function Write-RunLog { param([string]$Message) }
    function Get-SemtScheduledTaskState {
        param([string]$TaskPath, [string]$TaskName)
        $key = $TaskPath + $TaskName
        if ($script:HjStore.ContainsKey($key)) { return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $true; Enabled = [bool]$script:HjStore[$key]; Actions = @($script:HjActions[$key]); Detail = 'present' } }
        return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $false; Enabled = $false; Actions = @(); Detail = 'Task not found.' }
    }
    function Set-SemtScheduledTaskEnabled { param([string]$TaskPath, [string]$TaskName, [bool]$Enabled) $script:EnabledCalls.Add(('{0}{1}={2}' -f $TaskPath,$TaskName,$Enabled)); $script:HjStore[$TaskPath + $TaskName] = $Enabled }
    function Invoke-ProcessWithTimeout { param($FilePath, $Arguments, $TimeoutSeconds, $OutputFile) $script:Processes.Add(('{0} {1}' -f $FilePath,($Arguments -join ' '))); return [pscustomobject]@{ ExitCode = $script:ProcessExitCode; TimedOut = $false; Output = ''; OutputFile = $OutputFile; StopDetail = '' } }

    $missingHelpers = Test-UserContextHelperTasks
    Test-Step 'Hybrid Join: missing user-context tasks are reported, never created' { -not $missingHelpers.Success -and -not $missingHelpers.StatusTaskPresent -and $script:EnabledCalls.Count -eq 0 }
    $logonMissing = Enable-UserAutoEnrollAtLogonTask
    Test-Step 'Hybrid Join: next-logon task not enabled when the Prerequisites tasks are missing' { -not $logonMissing.Success -and $logonMissing.Detail -match 'SEMT_PREREQUISITE_TASK_MISSING' -and $script:EnabledCalls.Count -eq 0 }

    $script:ProcessExitCode = -2147024891
    $denied = Start-IntuneAutoEnrollment -Mode 'SystemUserCredential'
    Test-Step 'Hybrid Join: deviceenroller runs from SYSTEM (no scheduled task); a non-zero exit code is a failure, shown in hexadecimal' { $script:Processes.Count -eq 1 -and $script:Processes[0] -like '*\deviceenroller.exe /c /AutoEnrollMDM' -and -not $denied.Success -and $denied.ToolFound -and $denied.Mode -eq 'SystemUserCredential' -and $denied.Detail -match '0x80070005' }
    $script:ProcessExitCode = 0
    $ok = Start-IntuneAutoEnrollment -Mode 'SystemDeviceCredential'
    Test-Step 'Hybrid Join: deviceenroller exit code 0 is a success' { $ok.Success -and $ok.Mode -eq 'SystemDeviceCredential' -and [string]$ok.ExitCode -eq '0' }

    $script:LicenseEvents = @()
    function Get-WinEvent {
        param($FilterHashtable, $ErrorAction)
        $events = @($script:LicenseEvents | Where-Object { $_.Log -eq $FilterHashtable.LogName -and $_.TimeCreated -ge $FilterHashtable.StartTime -and $FilterHashtable.Id -contains $_.Id })
        if ($events.Count -eq 0) { $exception = New-Object System.Exception 'No events were found'; throw (New-Object System.Management.Automation.ErrorRecord $exception,'NoMatchingEventsFound,Microsoft.PowerShell.Commands.GetWinEventCommand','ObjectNotFound',$null) }
        return $events
    }
    $enrollmentLog = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Enrollment'
    $trigger = Get-Date '2026-11-05 10:15:53'
    $script:LicenseEvents = @([pscustomobject]@{ Log = $enrollmentLog; Id = 52; RecordId = 26558; TimeCreated = (Get-Date '2026-11-05 10:21:07'); Message = 'MDM Enroll: Server Returned Fault/Code/Subcode/Value=(UserLicense) Fault/Reason/Text=(Failed to issue token: UserValidation).' })
    $license = Get-RecentMdmUserLicenseFailure -Since $trigger
    Test-Step 'Hybrid Join: a UserLicense fault logged after the trigger is reported with its record' { $license -and $license.Id -eq 52 -and $license.RecordId -eq 26558 -and $license.LogName -eq $enrollmentLog }
    Test-Step 'Hybrid Join: a UserLicense fault older than the trigger is ignored; no event is not an error' { $null -eq (Get-RecentMdmUserLicenseFailure -Since (Get-Date '2026-11-05 10:30')) }
    $script:LicenseEvents = @([pscustomobject]@{ Log = $enrollmentLog; Id = 52; RecordId = 1; TimeCreated = (Get-Date '2026-11-05 10:21:07'); Message = 'MDM Enroll: Server Returned Fault/Code/Subcode/Value=(DeviceCapReached)' })
    Test-Step 'Hybrid Join: another server fault is not a license fault' { $null -eq (Get-RecentMdmUserLicenseFailure -Since $trigger) }
    Remove-Item -Path Function:\Get-WinEvent

    foreach ($task in $userTasks) { $script:HjStore[$task] = $true }
    $script:HjStore[$logonTask] = $false
    $script:HjActions[$logonTask] = @(('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$userTasks[0]), ('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$obsoleteAutoEnrollTask))
    $logonOutdated = Enable-UserAutoEnrollAtLogonTask
    Test-Step 'Hybrid Join: an outdated next-logon task (user-context deviceenroller) is not enabled' { -not $logonOutdated.Success -and $logonOutdated.Detail -match 'SEMT_PREREQUISITE_TASK_OUTDATED' -and $script:EnabledCalls.Count -eq 0 }
    $script:HjActions[$logonTask] = @(('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$userTasks[0]), $deviceEnrollerAction)
    $logonEnabled = Enable-UserAutoEnrollAtLogonTask
    Test-Step 'Hybrid Join: next-logon task is enabled (definition untouched)' { $logonEnabled.Success -and $script:EnabledCalls.Count -eq 1 -and $script:EnabledCalls[0] -eq ($logonTask + '=True') }
    Disable-UserAutoEnrollAtLogonTask -Reason 'SUCCESS'
    Disable-UserAutoEnrollAtLogonTask -Reason 'SUCCESS'
    Test-Step 'Hybrid Join: next-logon task is disabled once, after enrollment' { $script:EnabledCalls.Count -eq 2 -and $script:EnabledCalls[1] -eq ($logonTask + '=False') }

    $hjText = [IO.File]::ReadAllText($hjEndpoint)
    Test-Step 'Hybrid Join: no user-context deviceenroller task and no native task fallback remain' { $hjText -notmatch 'SEMT-HJ-UserMdmAutoEnroll|Start-UserContextIntuneAutoEnrollment|Get-NativeMdmAutoEnrollmentTask' -and $hjText -match '"INTUNE_USER_LICENSE_REJECTED" \{ return "ASSIGN_USER_INTUNE_LICENSE" \}' }
    #endregion
    #region Automatic prerequisites (launcher check and worker decision)
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force -ErrorAction Stop
    function Get-DefinitionsFor {
        param([string]$Campaign)
        $HybridJoinTasks = ($Campaign -eq 'HybridJoin'); $Windows11UpgradeTasks = ($Campaign -eq 'Windows11Upgrade')
        return @(Get-PrerequisiteTaskDefinitions | Where-Object { $_.Campaign -eq $Campaign })
    }
    function ConvertTo-ReadFromDefinition {
        param($Definition)
        return [pscustomobject]@{
            Exists = $true; Error = ''
            Actions = @($Definition.Actions | ForEach-Object { '{0}|{1}' -f [Environment]::ExpandEnvironmentVariables([string]$_.Execute),[string]$_.Arguments })
            TriggerCount = $(if ($Definition.Trigger -eq 'None') { 0 } else { 1 })
        }
    }
    foreach ($campaignName in @('HybridJoin','Windows11Upgrade')) {
        $definitions = @(Get-DefinitionsFor -Campaign $campaignName)
        $required = @(Get-SemtPrerequisiteTaskExpectations -Campaign $campaignName | Where-Object { $_.Kind -eq 'Required' })
        Test-Step ('Automatic prerequisites: {0} expectations list the same tasks as the Prerequisites endpoint' -f $campaignName) {
            $left = @($definitions | ForEach-Object { $_.TaskPath + $_.TaskName } | Sort-Object)
            $right = @($required | ForEach-Object { $_.TaskPath + $_.TaskName } | Sort-Object)
            $left.Count -gt 0 -and (($left -join ';') -eq ($right -join ';'))
        }
        Test-Step ('Automatic prerequisites: {0} tasks created by the Prerequisites endpoint are seen as present' -f $campaignName) {
            $states = @($required | ForEach-Object {
                $expectation = $_
                $definition = @($definitions | Where-Object { $_.TaskPath -eq $expectation.TaskPath -and $_.TaskName -eq $expectation.TaskName })[0]
                Test-SemtPrerequisiteTaskReadState -Expectation $expectation -Read (ConvertTo-ReadFromDefinition -Definition $definition)
            })
            @($states | Where-Object { $_ -ne 'Present' }).Count -eq 0
        }
    }
    $prereqText = [IO.File]::ReadAllText($prereqEndpoint)
    $obsolete = @(Get-SemtPrerequisiteTaskExpectations -Campaign 'HybridJoin' | Where-Object { $_.Kind -eq 'Obsolete' })
    Test-Step 'Automatic prerequisites: the obsolete task is the one the Prerequisites endpoint removes' {
        $obsolete.Count -eq 1 -and $prereqText.Contains(("Remove-PrerequisiteObsoleteTask -Campaign 'HybridJoin' -TaskPath `$script:SemtHybridJoinTaskPath -TaskName '{0}'" -f $obsolete[0].TaskName))
    }

    $hjExpectations = @(Get-SemtPrerequisiteTaskExpectations -Campaign 'HybridJoin')
    $refreshPrt = @($hjExpectations | Where-Object { $_.TaskName -eq 'SEMT-HJ-UserRefreshPrt' })[0]
    $logon = @($hjExpectations | Where-Object { $_.TaskName -eq 'SEMT-HJ-RunUserAutoEnrollAtLogon' })[0]
    $retry = @($hjExpectations | Where-Object { $_.TaskName -eq 'SmartM365-SEMT-HybridJoin-RetryAfterReboot' })[0]
    Test-Step 'Automatic prerequisites: older RefreshPrt task started through cmd.exe is outdated' {
        (Test-SemtPrerequisiteTaskReadState -Expectation $refreshPrt -Read ([pscustomobject]@{ Exists = $true; Error = ''; TriggerCount = 0; Actions = @('C:\WINDOWS\System32\cmd.exe|/d /c %windir%\System32\dsregcmd.exe /refreshprt') })) -eq 'Outdated'
    }
    Test-Step 'Automatic prerequisites: logon task without the SYSTEM deviceenroller action is outdated' {
        (Test-SemtPrerequisiteTaskReadState -Expectation $logon -Read ([pscustomobject]@{ Exists = $true; Error = ''; TriggerCount = 1; Actions = @('C:\WINDOWS\System32\schtasks.exe|/Run /TN "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserDsregStatus"','C:\WINDOWS\System32\schtasks.exe|/Run /TN "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserMdmAutoEnroll"') })) -eq 'Outdated'
    }
    Test-Step 'Automatic prerequisites: retry task without trigger is outdated; missing task is missing; read error is unknown' {
        $retryRead = [pscustomobject]@{ Exists = $true; Error = ''; TriggerCount = 0; Actions = @('C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe|-NoProfile -ExecutionPolicy RemoteSigned -File "C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\RetryAfterReboot\RetryAfterRebootRunner.ps1"') }
        $outdated = Test-SemtPrerequisiteTaskReadState -Expectation $retry -Read $retryRead
        $retryRead.TriggerCount = 1
        $present = Test-SemtPrerequisiteTaskReadState -Expectation $retry -Read $retryRead
        $missing = Test-SemtPrerequisiteTaskReadState -Expectation $retry -Read ([pscustomobject]@{ Exists = $false; Error = ''; TriggerCount = 0; Actions = @() })
        $unknown = Test-SemtPrerequisiteTaskReadState -Expectation $retry -Read ([pscustomobject]@{ Exists = $false; Error = 'Access is denied.'; TriggerCount = 0; Actions = @() })
        $outdated -eq 'Outdated' -and $present -eq 'Present' -and $missing -eq 'Missing' -and $unknown -eq 'Unknown' -and (Test-SemtPrerequisiteTaskReadState -Expectation $retry -Read $null) -eq 'Unknown'
    }
    Test-Step 'Automatic prerequisites: obsolete task present or absent' {
        (Test-SemtPrerequisiteTaskReadState -Expectation $obsolete[0] -Read ([pscustomobject]@{ Exists = $true; Error = ''; TriggerCount = 0; Actions = @('x|y') })) -eq 'Obsolete' -and (Test-SemtPrerequisiteTaskReadState -Expectation $obsolete[0] -Read ([pscustomobject]@{ Exists = $false; Error = ''; TriggerCount = 0; Actions = @() })) -eq 'Absent'
    }
    Test-Step 'Automatic prerequisites: decision is NeedsRun when proven missing, CheckFailed when only unreadable, Ready otherwise' {
        $row = { param($name,$state) [pscustomobject]@{ TaskName = $name; State = $state } }
        (Get-SemtPrerequisiteTaskSummary -Tasks @((& $row 'A' 'Present'),(& $row 'B' 'Unknown'),(& $row 'C' 'Missing'))).Decision -eq 'NeedsRun' -and
        (Get-SemtPrerequisiteTaskSummary -Tasks @((& $row 'A' 'Present'),(& $row 'B' 'Obsolete'))).Decision -eq 'NeedsRun' -and
        (Get-SemtPrerequisiteTaskSummary -Tasks @((& $row 'A' 'Present'),(& $row 'B' 'Unknown'))).Decision -eq 'CheckFailed' -and
        (Get-SemtPrerequisiteTaskSummary -Tasks @((& $row 'A' 'Present'),(& $row 'B' 'Absent'))).Decision -eq 'Ready'
    }
    Test-Step 'Automatic prerequisites: an unreachable target gives CheckFailed (nothing run)' {
        $unreachable = Get-SemtRemotePrerequisiteTaskState -ConnectionTarget 'semt-qualif-unreachable.invalid' -Campaign 'Windows11Upgrade' -TimeoutSeconds 60
        $unreachable.Decision -eq 'CheckFailed' -and @($unreachable.Tasks | Where-Object { $_.State -ne 'Unknown' }).Count -eq 0
    }

    # Worker decision with the remote check and the Prerequisites run stubbed.
    $workerPath = Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
    . ([scriptblock]::Create((Get-FunctionText -Path $workerPath -Name @('Invoke-WorkerAutoPrerequisites'))))
    $script:RemoteDecision = 'Ready'
    $script:PrerequisitesRuns = 0
    function Get-SemtRemotePrerequisiteTaskState { param($ConnectionTarget,$Campaign,$Mode) [pscustomobject]@{ Decision = $script:RemoteDecision; Detail = 'synthetic'; Tasks = @() } }
    function Invoke-WorkerPrerequisitesEndpoint { param($Auto) $script:PrerequisitesRuns++; [pscustomobject]@{ State = 'Created'; Detail = 'synthetic run.' } }
    function Write-SemtLogLine { param($Path,$Message,$Level) }
    function Invoke-WorkerDecision {
        param([string]$Decision,[string[]]$Arguments,[bool]$Enabled = $true,[string]$CampaignName = 'HybridJoin')
        $script:RemoteDecision = $Decision
        $script:PrerequisitesRuns = 0
        $contract = [pscustomobject]@{ Prerequisites = [pscustomobject]@{ Enabled = $Enabled } }
        $campaign = $CampaignName; $target = 'pc'; $remoteMode = 'ScheduledTask'; $logPath = ''
        $result = [ordered]@{}
        Invoke-WorkerAutoPrerequisites -EndpointArguments $Arguments
        return [pscustomobject]@{ State = [string]$result['PrerequisitesAutoState']; Runs = $script:PrerequisitesRuns; Has = $result.Contains('PrerequisitesAutoState') }
    }
    Test-Step 'Worker: tasks ready => the Prerequisites endpoint is not run' { $r = Invoke-WorkerDecision -Decision 'Ready' -Arguments @('-DataRoot','x'); $r.State -eq 'Ready' -and $r.Runs -eq 0 }
    Test-Step 'Worker: tasks missing => the Prerequisites endpoint runs first' { $r = Invoke-WorkerDecision -Decision 'NeedsRun' -Arguments @('-DataRoot','x'); $r.State -eq 'Created' -and $r.Runs -eq 1 }
    Test-Step 'Worker: tasks missing in audit only => reported, nothing run' { $r = Invoke-WorkerDecision -Decision 'NeedsRun' -Arguments @('-AuditOnly','-DataRoot','x'); $r.State -eq 'MissingAuditOnly' -and $r.Runs -eq 0 }
    Test-Step 'Worker: unreadable tasks => CheckFailed, nothing run' { $r = Invoke-WorkerDecision -Decision 'CheckFailed' -Arguments @(); $r.State -eq 'CheckFailed' -and $r.Runs -eq 0 }
    Test-Step 'Worker: option off or Prerequisites campaign => no check' { (-not (Invoke-WorkerDecision -Decision 'NeedsRun' -Arguments @() -Enabled $false).Has) -and (-not (Invoke-WorkerDecision -Decision 'NeedsRun' -Arguments @() -CampaignName 'Prerequisites').Has) }

    $workerText = [IO.File]::ReadAllText($workerPath)
    Test-Step 'Worker: automatic prerequisites run after the dry-run branch and before the campaign payload copy' {
        $dry = $workerText.IndexOf("if ([bool]`$contract.Flags.DryRun) {")
        $auto = $workerText.IndexOf('try { Invoke-WorkerAutoPrerequisites')
        $copy = $workerText.IndexOf('$copied = Copy-SemtRemotePayload')
        $dry -gt 0 -and $auto -gt $dry -and $copy -gt $auto
    }
    Test-Step 'Worker: unconfirmed automatic prerequisites => the Campaign endpoint is not copied nor started' {
        $flag = $workerText.IndexOf("`$prerequisitesUnconfirmed = ([string]`$result['PrerequisitesAutoState'] -eq 'Unconfirmed')")
        $loop = $workerText.IndexOf('for ($attempt = 1; $attempt -le 3 -and -not $prerequisitesUnconfirmed -and -not $prerequisitesNotStarted; $attempt++) {')
        $status = $workerText.IndexOf("Set-WorkerStatus -Status 'REMOTE_PREREQUISITES_UNCONFIRMED'")
        $next = $workerText.IndexOf("-NextAction 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH'", [math]::Max(0,$status))
        $copiedBranch = $workerText.IndexOf('elseif ($null -eq $copied) {')
        $auto = $workerText.IndexOf('try { Invoke-WorkerAutoPrerequisites')
        $flag -gt $auto -and $loop -gt $flag -and $status -gt $loop -and $next -gt $status -and $copiedBranch -gt $status
    }
    . ([scriptblock]::Create((Get-FunctionText -Path (Join-Path $packageRoot 'Campaigns\HybridJoin\SmartEndpointMigration.Campaign.HybridJoin.psm1') -Name @('Get-HybridJoinCentralLogBucket','Get-HybridJoinCooldownHours'))))
    Test-Step 'Worker: the unconfirmed status is a remote failure (not target busy) with the HybridJoin default cooldown' {
        (Get-HybridJoinCentralLogBucket -Result ([pscustomobject]@{ Status = 'REMOTE_PREREQUISITES_UNCONFIRMED' })) -eq 'RemoteCollectionFailed' -and
        (Get-HybridJoinCooldownHours -Status 'REMOTE_PREREQUISITES_UNCONFIRMED') -eq (Get-HybridJoinCooldownHours -Status 'REMOTE_EXECUTION_UNCONFIRMED')
    }
    Test-Step 'Worker: a Prerequisites task that did not start => the Campaign endpoint is not copied nor started' {
        $return = $workerText.IndexOf("State = 'NotStarted'")
        $flag = $workerText.IndexOf("`$prerequisitesNotStarted = ([string]`$result['PrerequisitesAutoState'] -eq 'NotStarted')")
        $loop = $workerText.IndexOf('-and -not $prerequisitesNotStarted; $attempt++) {')
        $branch = $workerText.IndexOf('elseif ($prerequisitesNotStarted) {')
        $status = $workerText.IndexOf("Set-WorkerStatus -Status 'REMOTE_TASK_NOT_STARTED'", [math]::Max(0,$branch))
        $copiedBranch = $workerText.IndexOf('elseif ($null -eq $copied) {')
        $return -gt 0 -and $flag -gt $return -and $loop -gt $flag -and $branch -gt $loop -and $status -gt $branch -and $copiedBranch -gt $status
    }
    Test-Step 'Worker: a task that did not start is retried with the short HybridJoin backoff' {
        (Get-HybridJoinCentralLogBucket -Result ([pscustomobject]@{ Status = 'REMOTE_TASK_NOT_STARTED' })) -eq 'RemoteCollectionFailed' -and
        (Get-HybridJoinCooldownHours -Status 'REMOTE_TASK_NOT_STARTED' -BackoffCount 1) -eq 0.25 -and
        (Get-HybridJoinCooldownHours -Status 'REMOTE_TASK_NOT_STARTED' -BackoffCount 9) -eq 6
    }
    $launcherText = [IO.File]::ReadAllText((Join-Path $packageRoot 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'))
    $catalogText = [IO.File]::ReadAllText((Join-Path $packageRoot 'Config\SmartEndpointMigrationToolkit.options.json.txt'))
    Test-Step 'Launcher: AutoPrerequisites option (default on) adds the Prerequisites endpoint to the integrity check and the worker contract' {
        $catalogText -match '"Name": "AutoPrerequisites", "Scope": "Common", "Type": "Boolean", "Default": true' -and
        $launcherText.Contains('if ($useAutoPrerequisites) { $componentPaths += $autoPrerequisitesRelativePath }') -and
        $launcherText.Contains('-Prerequisites $script:AutoPrerequisites') -and
        $launcherText.Contains("`$prerequisitesArguments = @(`$(if (`$Campaign -eq 'HybridJoin') { '-HybridJoinTasks' } else { '-Windows11UpgradeTasks' }))")
    }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAxTPmIMUE/kN6J
# QFddu229ZsLLNxHdEidqYvviRRuK7aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPwWuKLfDp6tefykqPNGtw9bM6Mg3JPYg+veMIjqgLObMA0GCSqG
# SIb3DQEBAQUABIIBgDF/55gQQIQoVnB0qSW17pSUBAoJkXo4nv5zWWXyYe7HHchX
# S31kVXMTYbpR2HH2ou2+ZrPxYaSz7s1QrgAXV8FWs9M37O5scabtgHmgb4DmFJY7
# 4JOe83Qi3YO21CdEFmqUrR35/KjPTedIx7YDUfXsZ8/IjydOyXlZQjGsWrORhfZS
# 5UhYgr9XKxeDKcN+5wkxDYc4RVxC7QH9j2AmqdSUyTZ42VJBx8CN2wrRBCgkgcZ1
# UlTN0mVKbF36uybgvMEaXbJbKCCj9ApFOXPXdynPxeu7W1ZtUxJwRh6ajdVxiU07
# xBV1zvu2/Xbb0lpZ10j8knVPbeeFdr/R8ago7IBD2TLhMi8kKT5Mhyn7Dg/MpTSt
# pseQjSoys67StotZnz8kjgAtnNE+5eTn6dc20eGKojkAlD8GtNZ81jQeGoEtOvXR
# 7lfqJa3mGqKhIGMpUi6mjp7UPexciWcLlDA4rY/M7k0BQ4H4tgJGTzzTDy3Mwd/V
# kipqoOwJLLuPNtkbj6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYyMTI5
# MTdaMC8GCSqGSIb3DQEJBDEiBCAAljy7xF4NbwGvoi7syQIvQTeWXUi8FYs9fqYv
# X8OxCDANBgkqhkiG9w0BAQEFAASCAgAIZVsotgE8BqkENBKAPPM8UvSNWed1fkEj
# K/hyjU2ai7X2nxHEQHqdcctIfJAudj3j6HgSGGkFbPUwuIiQrJC28ZeoJCQyZiTJ
# jTvCXpZ9U6HDoBe757OeC0blXNAbTLVFaq4rqy3BptDjrxEAkfp89uZnhKvfih/O
# gjgo4qePsqM6gAsrxe0Px1/3RUWIEvvvrDrCo/qwU6GkbX7eQuM8VvlQcRATeI2C
# 3LswQRLzbzgi0DtmNI5Niq1AguW1AKYmoQNcnUVZm2i+WvnqW2MCI9wdalowCQ2J
# Cy/tvBQdqSowUfWieQYqm0/NDuD33au4utsMZIDFjBLrrJ8/wEZONi/vXHO5glI+
# 1DSviak+RL56WGwwjJseyljWdCoK6ve3XueQFDl92lXTbKnoTWC/EKI2gGozEEzh
# WiC2OliN/kcjc6gkTns2V7j9IHWZa2fwwv9ypXHQ0lMs/nIxF6AYnPRyxTs6O9Ha
# 3RXmCqeY6VNrgWAloGour7rwxJ2iN6IUMYyuUkJZDvF271ZpusOdf7on0QP7aIoq
# +NAndG35uUTbw2PF0tX7v7JlgTpgSIjdMUShh1kTRfHWby22kzBZA2vFHd7atXbe
# 8NRZ+eDLmxbAnlPW4W6i3EGUCrv07saElaCS4bNNfTQlwEBUSr/ycb4+eyTFdo/m
# GU4ZNIHk8w==
# SIG # End signature block
