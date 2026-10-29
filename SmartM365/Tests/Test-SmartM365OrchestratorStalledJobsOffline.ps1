<#
.SYNOPSIS
Offline tests of the stalled-job alert: a running job whose log stays silent beyond MaxSilentMinutes
is reported once per run (log line, email, GUI and Dashboard) and is never stopped automatically.
.VERSION
1.0.0
.NOTES
Synthetic job log and running state in a temporary folder; one short child pwsh of this test
measures the process tree CPU. No tenant initialization, collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}

# Wiring: checked for every supervised job after the timeout handling; never a stop or a kill.
$update = Get-FunctionText 'Update-RunningJobs'
Check ($update -match 'Update-RunningJobActivity -JobName \$name -RunInfo \$info -Now \$Now') 'Update-RunningJobs no longer checks the job activity.'
$activity = Get-FunctionText 'Update-RunningJobActivity'
Check ($activity -notmatch 'Stop-ProcessTree|taskkill|Stop-Process|Complete-JobRun|TimeoutRequested') 'The stalled-job check must never stop a job.'
Check ((Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'Stalled = ' -and (Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'CpuSecondsPerMinute = ') 'The heartbeat no longer publishes the stalled-job signals.'
Check ((Get-FunctionText 'Restore-RunningJobs') -match 'StallAlertSent = ') 'The once-per-run alert flag is not restored after an orchestrator restart.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Stalled-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$child = $null
try {
    $functions = @('Get-OrchestratorProcessTreeCpuTime', 'Get-OrchestratorJobLogActivity', 'Update-RunningJobActivity')
    $stubs = @'
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add(('{0} {1}' -f $Level, $Message)) }
function Get-JobState { param($JobName) $script:JobState }
function Save-OrchestratorState { $script:Saves++ }
function Send-JobStalledEmail { param($JobName, $RunInfo) $script:Mails.Add($JobName); $true }
'@
    $module = New-Module -ScriptBlock ([scriptblock]::Create((@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs))
    & $module {
        $script:Logs = [Collections.Generic.List[string]]::new(); $script:Mails = [Collections.Generic.List[string]]::new(); $script:Saves = 0
        $script:JobState = @{ Running = @{} }
        $script:JobActivityCheckSeconds = 60
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ 'Slow-Job' = [pscustomobject]@{ MaxSilentMinutes = 60 }; 'Quiet-Job' = [pscustomobject]@{ MaxSilentMinutes = 0 } } }
    }

    # Process tree CPU: the launcher and its child pwsh are both counted.
    $child = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList '-NoProfile', '-Command', '$end=(Get-Date).AddSeconds(2); while((Get-Date) -lt $end){}; Start-Sleep -Seconds 60' -PassThru -WindowStyle Hidden
    Start-Sleep -Seconds 3
    $treeCpu = & $module { param($p) Get-OrchestratorProcessTreeCpuTime -RootPid $p } $PID
    $selfCpu = (Get-Process -Id $PID).TotalProcessorTime.TotalSeconds
    Check ($treeCpu -ge $selfCpu + 1) "The process tree CPU must include the child process ($treeCpu s, this process $selfCpu s)."

    # Log silence: the size read through a shared handle drives the last activity.
    $logPath = Join-Path $root 'Job-Slow-Job_SRV-A_20261029_043000.log'
    [IO.File]::WriteAllText($logPath, "start`n")
    $start = (Get-Date).AddMinutes(-10)
    [IO.File]::SetLastWriteTimeUtc($logPath, $start.ToUniversalTime().AddMinutes(2))
    $runInfo = @{ StartTime = $start; LogPath = $logPath; Process = $child }
    $t0 = Get-Date
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t0 $runInfo
    Check (-not $runInfo.Stalled -and [math]::Abs($runInfo.SilentMinutes - 8) -lt 0.2 -and $runInfo.MaxSilentMinutes -eq 60 -and $null -ne $runInfo.CpuSeconds) "First check: silence measured from the last log write ($($runInfo.SilentMinutes) min)."
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t0.AddSeconds(30) $runInfo
    Check ([math]::Abs($runInfo.SilentMinutes - 8) -lt 0.2) 'The check must run at most once a minute.'

    # 61 minutes later without any write: stalled, one WARN line and one email, never a stop.
    $t1 = $t0.AddMinutes(53)
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t1 $runInfo
    $logs = @(& $module { $script:Logs })
    Check ($runInfo.Stalled -and $runInfo.SilentMinutes -ge 60 -and @($logs | Where-Object { $_ -like 'WARN Job Slow-Job: stalled*Not stopped*' }).Count -eq 1) "A silent job must be reported as stalled ($($logs -join ' | '))."
    Check ((& $module { $script:Mails.Count }) -eq 1 -and $runInfo.StallAlertSent -and (& $module { $script:JobState.Running.StallAlertSent }) -and (& $module { $script:Saves }) -eq 1) 'The alert must be emailed once and persisted with the running state.'
    Check ($null -ne $runInfo.CpuSecondsPerMinute -and $runInfo.CpuSecondsPerMinute -ge 0) 'The CPU progression per minute must be measured.'
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t1.AddMinutes(5) $runInfo
    Check ((& $module { $script:Mails.Count }) -eq 1 -and @(& $module { $script:Logs } | Where-Object { $_ -like 'WARN*stalled*' }).Count -eq 1) 'A still-stalled job must not be alerted again.'

    # The log moves again: the state clears with an INFO line; a second silence in the same run sends no new email.
    [IO.File]::AppendAllText($logPath, "progress`n")
    $t2 = $t1.AddMinutes(10)
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t2 $runInfo
    Check (-not $runInfo.Stalled -and $runInfo.SilentMinutes -eq 0 -and @(& $module { $script:Logs } | Where-Object { $_ -like 'INFO Job Slow-Job: log activity resumed*' }).Count -eq 1) 'Log activity must clear the stalled state.'
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } $t2.AddMinutes(61) $runInfo
    Check ($runInfo.Stalled -and (& $module { $script:Mails.Count }) -eq 1) 'The email is sent once per run, even after a second silence.'

    # MaxSilentMinutes = 0 disables the alert; a restored run that already alerted sends no email.
    $quiet = @{ StartTime = (Get-Date).AddHours(-5); LogPath = $logPath; Process = $null }
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Quiet-Job' -RunInfo $r -Now $n } (Get-Date).AddHours(3) $quiet
    Check (-not $quiet.Stalled -and $quiet.MaxSilentMinutes -eq 0) 'MaxSilentMinutes = 0 must disable the alert.'
    $restored = @{ StartTime = (Get-Date).AddHours(-3); LogPath = (Join-Path $root 'missing.log'); Process = $null; StallAlertSent = $true }
    & $module { param($n, $r) Update-RunningJobActivity -JobName 'Slow-Job' -RunInfo $r -Now $n } (Get-Date) $restored
    Check ($restored.Stalled -and (& $module { $script:Mails.Count }) -eq 1) 'A restored run without log is measured from its start and keeps its once-per-run flag.'

    # Manifest: default 60, negative refused by the orchestrator and by the GUI publication check.
    $parse = Get-FunctionText 'ConvertTo-NormalizedJob'
    Check ($parse -match '\$maxSilentMinutes = 60' -and $parse -match 'MaxSilentMinutes cannot be negative') 'The manifest must default MaxSilentMinutes to 60 and refuse a negative value.'
    $management = [IO.File]::ReadAllText((Join-Path $base 'SmartM365.Orchestrator.Management.psm1'))
    Check ($management -match "'MaxSilentMinutes'") 'The GUI publication check must refuse a negative MaxSilentMinutes.'

    # GUI side: heartbeat signals become Running rows and a Critical Dashboard item.
    New-Item -ItemType Directory -Path (Join-Path $root 'Shared/SRV-A'), (Join-Path $root 'Shared/SRV-B') -Force | Out-Null
    $now = Get-Date
    $heartbeat = [ordered]@{ Timestamp = $now.ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.72'; JobStopProtocol = 1
        RunningJobs = @(
            [ordered]@{ Name = 'Slow-Job'; Pid = 4242; StartTime = $now.AddMinutes(-110).ToString('o'); ScheduledOccurrence = $now.AddMinutes(-110).ToString('o'); LastLogWriteUtc = $now.ToUniversalTime().AddMinutes(-75).ToString('o'); SilentMinutes = 75; MaxSilentMinutes = 60; Stalled = $true; CpuSeconds = 58.2; CpuSecondsPerMinute = 0.3 }
            [ordered]@{ Name = 'Busy-Job'; Pid = 4243; StartTime = $now.AddMinutes(-5).ToString('o'); ScheduledOccurrence = $now.AddMinutes(-5).ToString('o'); SilentMinutes = 0.5; MaxSilentMinutes = 60; Stalled = $false; CpuSeconds = 40; CpuSecondsPerMinute = 9.5 }
        ) }
    [IO.File]::WriteAllText((Join-Path $root 'Shared/SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    # SRV-B runs an older orchestrator without the signals.
    $old = [ordered]@{ Timestamp = $now.ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.71'; RunningJobs = @([ordered]@{ Name = 'Old-Job'; Pid = 5000; StartTime = $now.AddMinutes(-3).ToString('o') }) }
    [IO.File]::WriteAllText((Join-Path $root 'Shared/SRV-B/Orchestrator-Heartbeat.json'), ($old | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Global
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); PeerHeartbeatStaleMinutes = 5 }
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath (Join-Path $root 'Shared') -ClusterDocument $cluster -Now $now
    $slow = @($operations.Running | Where-Object Job -eq 'Slow-Job')[0]
    $busy = @($operations.Running | Where-Object Job -eq 'Busy-Job')[0]
    $oldRow = @($operations.Running | Where-Object Job -eq 'Old-Job')[0]
    Check ($slow.Activity -eq 'Stalled' -and $slow.SilentMinutes -eq 75 -and $slow.CpuSeconds -eq 58.2 -and $slow.CpuSecondsPerMinute -eq 0.3) 'A stalled run must reach the Running rows with its silence and CPU.'
    Check ($busy.Activity -eq 'Active' -and $oldRow.Activity -eq '' -and $null -eq $oldRow.SilentMinutes) 'An active run and an older orchestrator must be shown as such.'
    $insights = [IO.File]::ReadAllText((Join-Path $base 'SmartM365.Orchestrator.Insights.psm1'))
    Check ($insights -match "Add-DashboardAction 'Critical' 'Stalled job'") 'The Dashboard must list a stalled job as Critical.'

    $guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
    foreach ($binding in @('Activity', 'SilentMinutes', 'CpuSeconds', 'CpuSecondsPerMinute')) { Check ($guiText -match ('Binding="\{Binding ' + $binding + '\}"')) "Running jobs column missing: $binding" }
    Check ($guiText -match '<DataTrigger Binding="\{Binding Activity\}" Value="Stalled">') 'A stalled run must be highlighted in Running jobs.'
}
finally {
    if ($null -ne $child) { try { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } catch { } }
    Remove-Module -Name SmartM365.Orchestrator.Insights -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: stalled-job alert, $($script:passed) offline checks."
