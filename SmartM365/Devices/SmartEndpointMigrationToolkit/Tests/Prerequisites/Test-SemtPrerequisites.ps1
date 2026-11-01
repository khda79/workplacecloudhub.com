<#
.SYNOPSIS
    Offline tests of the Prerequisites campaign and of the scheduled-task consumers.
.DESCRIPTION
    The Prerequisites endpoint functions run against an in-memory task store: creation of the
    HybridJoin and Windows11Upgrade tasks (enabled state at creation), repeat runs, audit, a task
    with another definition (updated, enabled state kept), registration failure and unreadable
    Task Scheduler. Task state parsing is checked with a localized SYSTEM account name.
    Hybrid Join consumer functions: missing user-context tasks fall back to the native
    EnterpriseMgmt task; the next-logon task is only enabled or disabled.
    No scheduled task, registry, network or endpoint run on the computer.
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
    . ([scriptblock]::Create((Get-FunctionText -Path $prereqEndpoint -Name @('Get-PrerequisiteTaskDefinitions','Test-PrerequisiteTaskMatch','Invoke-PrerequisiteTask'))))

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
    function Invoke-Scenario {
        param([switch]$Audit, [switch]$HybridJoin, [switch]$Windows11)
        $script:AuditOnly = [bool]$Audit; $AuditOnly = [bool]$Audit
        $script:HybridJoinTasks = [bool]$HybridJoin; $HybridJoinTasks = [bool]$HybridJoin
        $script:Windows11UpgradeTasks = [bool]$Windows11; $Windows11UpgradeTasks = [bool]$Windows11
        $script:Counters = [ordered]@{ TasksExpected = 0; TasksPresent = 0; TasksCreated = 0; TasksUpdated = 0; TasksMissing = 0; TasksFailed = 0 }
        $script:TaskRows = New-Object System.Collections.Generic.List[object]
        $script:Registered.Clear()
        foreach ($definition in @(Get-PrerequisiteTaskDefinitions)) { Invoke-PrerequisiteTask -Definition $definition }
        return [pscustomobject]@{ Counters = $script:Counters; Rows = @($script:TaskRows.ToArray()); Registered = @($script:Registered.ToArray()) }
    }
    $hjFolder = '\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\'
    $userTasks = @('SEMT-HJ-UserDsregStatus','SEMT-HJ-UserRefreshPrt','SEMT-HJ-UserMdmAutoEnroll' | ForEach-Object { $hjFolder + $_ })
    $logonTask = $hjFolder + 'SEMT-HJ-RunUserAutoEnrollAtLogon'
    $hjRetry = '\SmartM365-SEMT-HybridJoin-RetryAfterReboot'
    $w11Retry = '\SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot'

    $audit = Invoke-Scenario -Audit -HybridJoin -Windows11
    Test-Step 'audit, no task: six tasks reported missing, nothing registered' { $audit.Counters.TasksExpected -eq 6 -and $audit.Counters.TasksMissing -eq 6 -and $audit.Registered.Count -eq 0 -and @($audit.Rows | Where-Object { $_.Result -ne 'Missing' }).Count -eq 0 }

    $first = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'repair, no task: the six tasks are created' { $first.Counters.TasksCreated -eq 6 -and $first.Counters.TasksFailed -eq 0 -and $first.Registered.Count -eq 6 }
    Test-Step 'on-demand user-context tasks are created enabled (no trigger), group INTERACTIVE' { @($userTasks | Where-Object { $script:Store[$_].Enabled -and $script:Store[$_].TriggerCount -eq 0 -and $script:Store[$_].GroupSid -eq 'S-1-5-4' }).Count -eq 3 }
    Test-Step 'next-logon and both retry-after-reboot tasks are created disabled, as SYSTEM' { @(@($logonTask,$hjRetry,$w11Retry) | Where-Object { -not $script:Store[$_].Enabled -and $script:Store[$_].UserSid -eq 'S-1-5-18' -and $script:Store[$_].TriggerCount -eq 1 }).Count -eq 3 }
    Test-Step 'retry-after-reboot tasks run the fixed runner command of the shared region' {
        $ok = $true
        foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
            $spec = Get-SemtRetryAfterRebootTaskSpec -Campaign $campaign
            if (-not (Test-SemtRetryAfterRebootTaskState -State $script:Store[$spec.TaskPath + $spec.TaskName] -Spec $spec)) { $ok = $false }
        }
        $ok
    }
    Test-Step 'dsreg status keeps its cmd.exe redirection; refreshprt and AutoEnrollMDM start the executable directly; next-logon command unchanged' {
        @($script:Store[$userTasks[0]].Actions)[0] -eq 'cmd.exe|/d /c %windir%\System32\dsregcmd.exe /status > %windir%\Temp\SEMT-HJ-UserDsregStatus_%USERNAME%.txt 2>&1' -and @($script:Store[$userTasks[1]].Actions)[0] -eq ('{0}|/refreshprt' -f (Join-Path $env:WINDIR 'System32\dsregcmd.exe')) -and @($script:Store[$userTasks[2]].Actions)[0] -eq ('{0}|/c /AutoEnrollMDM' -f (Join-Path $env:WINDIR 'System32\deviceenroller.exe')) -and @($script:Store[$logonTask].Actions)[1] -eq ('{0}|/Run /TN "{1}"' -f (Join-Path $env:WINDIR 'System32\schtasks.exe'),$userTasks[2])
    }
    $script:Store[$userTasks[2]].Actions = @('cmd.exe|/d /c %windir%\System32\deviceenroller.exe /c /AutoEnrollMDM > %windir%\Temp\SEMT-HJ-UserMdmAutoEnroll_%USERNAME%.txt 2>&1')
    $legacy = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'an older cmd.exe auto-enrollment task is updated to the direct command' { $legacy.Counters.TasksUpdated -eq 1 -and $legacy.Registered -contains $userTasks[2] -and @($script:Store[$userTasks[2]].Actions)[0] -like '*deviceenroller.exe|/c /AutoEnrollMDM' }

    $second = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'repeat run: every task Present, nothing registered again' { $second.Counters.TasksPresent -eq 6 -and $second.Registered.Count -eq 0 }

    $script:Store[$logonTask].Enabled = $true
    $script:Store[$w11Retry] = [pscustomobject]@{ TaskPath = '\'; TaskName = 'SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot'; Exists = $true; Enabled = $true; Actions = @('powershell.exe|-NoProfile -ExecutionPolicy RemoteSigned -File "C:\Program Files\old\RetryAfterRebootRunner.ps1"'); TriggerCount = 1; UserSid = 'S-1-5-18'; GroupSid = ''; Detail = 'Enabled=True' }
    $third = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'a task with another definition is updated and keeps its enabled state' { $third.Counters.TasksUpdated -eq 1 -and $third.Registered -contains $w11Retry -and $script:Store[$w11Retry].Enabled -and @($third.Rows | Where-Object { $_.TaskName -eq $w11Retry -and $_.Before -eq 'Different' -and $_.Result -eq 'Updated' }).Count -eq 1 }
    Test-Step 'an enabled task with the expected definition is left unchanged (a campaign may be waiting on it)' { $third.Registered -notcontains $logonTask -and $script:Store[$logonTask].Enabled }

    $script:Store = @{}
    $script:FailRegister = @($hjRetry)
    $script:Unreadable = @($w11Retry)
    $fourth = Invoke-Scenario -HybridJoin -Windows11
    Test-Step 'a registration failure and an unreadable Task Scheduler are counted as failures; other tasks are still created' { $fourth.Counters.TasksFailed -eq 2 -and $fourth.Counters.TasksCreated -eq 4 -and $fourth.Registered -notcontains $w11Retry -and @($fourth.Rows | Where-Object { $_.TaskName -eq $hjRetry -and $_.Result -eq 'Failed' -and $_.Detail -match 'Access is denied' }).Count -eq 1 }
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

    #region Hybrid Join consumers (missing tasks)
    foreach ($name in @('Get-SemtScheduledTaskState','Register-PrerequisiteTask')) { Remove-Item -Path ('Function:\' + $name) -ErrorAction SilentlyContinue }
    . ([scriptblock]::Create((Get-FunctionText -Path $hjEndpoint -Name @('Get-HybridJoinTaskState','Test-UserContextHelperTasks','Start-UserContextIntuneAutoEnrollment','Enable-UserAutoEnrollAtLogonTask','Disable-UserAutoEnrollAtLogonTask'))))
    $UserStatusTaskName = $userTasks[0]; $UserRefreshPrtTaskName = $userTasks[1]; $UserIntuneAutoEnrollTaskName = $userTasks[2]; $UserNextLogonAutoEnrollTaskName = $logonTask
    $script:HjStore = @{}
    $script:Runs = New-Object System.Collections.Generic.List[string]
    $script:EnabledCalls = New-Object System.Collections.Generic.List[string]
    function Write-RunLog { param([string]$Message) }
    function Get-SemtScheduledTaskState {
        param([string]$TaskPath, [string]$TaskName)
        $key = $TaskPath + $TaskName
        if ($script:HjStore.ContainsKey($key)) { return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $true; Enabled = [bool]$script:HjStore[$key]; Detail = 'present' } }
        return [pscustomobject]@{ TaskPath = $TaskPath; TaskName = $TaskName; Exists = $false; Enabled = $false; Detail = 'Task not found.' }
    }
    function Set-SemtScheduledTaskEnabled { param([string]$TaskPath, [string]$TaskName, [bool]$Enabled) $script:EnabledCalls.Add(('{0}{1}={2}' -f $TaskPath,$TaskName,$Enabled)); $script:HjStore[$TaskPath + $TaskName] = $Enabled }
    function Invoke-OnDemandScheduledTask { param([string]$TaskName) $script:Runs.Add($TaskName); return [pscustomobject]@{ Success = $true; ExitCode = 0; Detail = ''; ResolvedTaskName = $TaskName } }
    function Get-NativeMdmAutoEnrollmentTask { return [pscustomobject]@{ TaskName = 'Schedule created by enrollment client'; TaskPath = '\Microsoft\Windows\EnterpriseMgmt\x\'; FullName = '\Microsoft\Windows\EnterpriseMgmt\x\Schedule created by enrollment client' } }
    function Get-LatestUserTaskOutputFile { param($Filter, $Since, $ExpectedFileName) return $null }

    $missingHelpers = Test-UserContextHelperTasks
    Test-Step 'Hybrid Join: missing user-context tasks are reported, never created' { -not $missingHelpers.Success -and -not $missingHelpers.StatusTaskPresent -and $script:EnabledCalls.Count -eq 0 }
    $native = Start-UserContextIntuneAutoEnrollment -PreferredTaskName $UserIntuneAutoEnrollTaskName -WaitSeconds 0 -ExpectedUserName 'user1'
    Test-Step 'Hybrid Join: missing user auto-enrollment task falls back to the native EnterpriseMgmt task without running the missing task' { $native.Mode -eq 'NativeEnterpriseMgmtTask' -and $native.Success -and $script:Runs -notcontains $UserIntuneAutoEnrollTaskName -and $script:Runs.Count -eq 1 }
    $logonMissing = Enable-UserAutoEnrollAtLogonTask
    Test-Step 'Hybrid Join: next-logon task not enabled when the Prerequisites tasks are missing' { -not $logonMissing.Success -and $logonMissing.Detail -match 'SEMT_PREREQUISITE_TASK_MISSING' -and $script:EnabledCalls.Count -eq 0 }

    foreach ($task in $userTasks) { $script:HjStore[$task] = $true }
    $script:HjStore[$logonTask] = $false
    $script:Runs.Clear()
    $userTask = Start-UserContextIntuneAutoEnrollment -PreferredTaskName $UserIntuneAutoEnrollTaskName -WaitSeconds 0 -ExpectedUserName 'user1'
    Test-Step 'Hybrid Join: present user auto-enrollment task is started; no output file is read' { $userTask.Mode -eq 'UserTask' -and $script:Runs -contains $UserIntuneAutoEnrollTaskName -and $userTask.OutputFile -eq '' }
    $logonEnabled = Enable-UserAutoEnrollAtLogonTask
    Test-Step 'Hybrid Join: next-logon task is enabled (definition untouched)' { $logonEnabled.Success -and $script:EnabledCalls.Count -eq 1 -and $script:EnabledCalls[0] -eq ($logonTask + '=True') }
    Disable-UserAutoEnrollAtLogonTask -Reason 'SUCCESS'
    Disable-UserAutoEnrollAtLogonTask -Reason 'SUCCESS'
    Test-Step 'Hybrid Join: next-logon task is disabled once, after enrollment' { $script:EnabledCalls.Count -eq 2 -and $script:EnabledCalls[1] -eq ($logonTask + '=False') }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
