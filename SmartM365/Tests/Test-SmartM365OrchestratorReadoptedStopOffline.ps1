<#
.SYNOPSIS
Offline test of an operator stop on a job re-adopted after an orchestrator restart: the same
process is stopped, a process whose start time differs is refused.
.VERSION
1.0.0
.NOTES
Runs the real Restore-RunningJobs, Update-RunningJobs and Test-ProcessMatchesRecord against short
child pwsh processes of this test. Shared claims, leases, state saving and job completion are
stubbed. No tenant initialization, collectors, scheduling or publication.
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

$functions = @('Restore-RunningJobs', 'Update-RunningJobs', 'Test-ProcessMatchesRecord', 'ConvertFrom-StateTime', 'ConvertTo-StateTime', 'Get-RunningJobTimeoutWindow', 'Stop-ProcessTree')
$stubs = @'
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add(('{0} {1}' -f $Level, $Message)) }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now) $script:Logs.Add('WARN ' + $Message) }
function Sync-RunningJobConcurrencyLease { param($JobName, $RunInfo, $Now) }
function Test-OrchestratorStatePersistenceReady { $true }
function Update-RunningJobActivity { param($JobName, $RunInfo, $Now) }
function Get-SmartM365JsonReadPath { param($Path, [switch]$Optional) $null }
function Enter-SmartM365OrchestratorConcurrencyLease { [pscustomobject]@{ Acquired = $false; Reason = 'stub' } }
function Get-CompletedJobResult { param($JobName, $Record, $ExpectedStart) $null }
function Complete-JobRun { param($JobName, $RunInfo, $StatusHint, $ExitCode, $EndTime, $ErrorText) $script:Completed.Add(('{0}={1}' -f $JobName, $StatusHint)); $script:RunningJobs.Remove($JobName) }
'@
$module = New-Module -ScriptBlock ([scriptblock]::Create((@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs))
$pwshPath = (Get-Process -Id $PID).Path
$children = [Collections.Generic.List[object]]::new()
function Start-Child { $p = Start-Process -FilePath $pwshPath -ArgumentList '-NoProfile', '-Command', 'Start-Sleep -Seconds 120' -PassThru -WindowStyle Hidden; $children.Add($p); $p }
function Get-Record($Process, [datetime]$StartTime) {
    # As written by a previous orchestrator instance (state file, round-trip time format).
    @{ Pid = $Process.Id; StartTime = $StartTime.ToString('o'); ScheduledOccurrence = $StartTime.ToString('o'); LogPath = ''; Attempt = 1; TimeoutMinutes = 1380; ProcessName = 'pwsh' }
}
try {
    $same = Start-Child
    $reused = Start-Child
    Start-Sleep -Seconds 2
    $sameStart = (Get-Process -Id $same.Id).StartTime
    & $module {
        param($SameRecord, $ReusedRecord)
        $script:Logs = [Collections.Generic.List[string]]::new(); $script:Completed = [Collections.Generic.List[string]]::new()
        $script:RunningJobs = @{}
        $script:State = @{ Jobs = @{ 'Stuck-Job' = @{ Running = $SameRecord }; 'Reused-Job' = @{ Running = $ReusedRecord } } }
        $script:Manifest = [pscustomobject]@{ OrderedJobs = @(); JobsByName = @{} }
        $script:Settings = [pscustomobject]@{ ElectionClaimGraceMinutes = 30; ConcurrencyLeasesPath = ''; SharedDataFolderPath = ''; PeerHeartbeatStaleMinutes = 5 }
        $script:ElectionPlan = $null
        Restore-RunningJobs
    } (Get-Record $same $sameStart) (Get-Record $reused (Get-Date).AddHours(-3))

    # Restart: the same process is re-adopted; a PID whose start time differs is not.
    $running = & $module { $script:RunningJobs }
    Check ($running.ContainsKey('Stuck-Job') -and $running['Stuck-Job'].ProcessName -eq 'pwsh') 'The re-adopted job must keep its expected process name.'
    Check (-not $running.ContainsKey('Reused-Job') -and (& $module { $script:Completed }) -contains 'Reused-Job=Interrupted') 'A PID with another start time must not be re-adopted.'

    # Operator stop after the restart: the identity check passes and the process tree is stopped.
    & $module { $info = $script:RunningJobs['Stuck-Job']; $info['OperatorStopRequested'] = $true; $info['OperatorStopBy'] = 'admin'; $info['OperatorStopReason'] = 'stalled'; Update-RunningJobs -Now (Get-Date) }
    $logs = @(& $module { $script:Logs })
    Check (@($logs | Where-Object { $_ -like '*stop deferred*' }).Count -eq 0) "The stop of a re-adopted job must not be deferred: $($logs -join ' | ')"
    Check ((& $module { $script:Completed }) -contains 'Stuck-Job=Cancelled' -and $same.HasExited) 'The re-adopted job must be stopped and recorded as cancelled.'

    # A process that no longer matches the selected run is still refused.
    $other = Start-Child
    Start-Sleep -Seconds 2
    & $module {
        param($Process)
        $script:Logs.Clear()
        $script:RunningJobs['Other-Job'] = @{ Process = (Get-Process -Id $Process.Id); StartTime = (Get-Date).AddHours(-2); ProcessName = 'pwsh'; OperatorStopRequested = $true; OperatorStopBy = 'admin'; OperatorStopReason = 'test'; TimeoutMinutes = 1380 }
        Update-RunningJobs -Now (Get-Date)
    } $other
    $logs = @(& $module { $script:Logs })
    Check (@($logs | Where-Object { $_ -like 'ERROR Job Other-Job: stop deferred because PID*no longer matches its recorded process*' }).Count -eq 1 -and -not $other.HasExited) 'A start-time mismatch must keep refusing the stop.'

    # A job ended outside the orchestrator (taskkill) is completed and leaves Running jobs.
    Stop-Process -Id $other.Id -Force
    $other.WaitForExit(10000) | Out-Null
    & $module { Update-RunningJobs -Now (Get-Date) }
    Check ((& $module { $script:Completed }) -contains 'Other-Job=Exited' -and -not (& $module { $script:RunningJobs }).ContainsKey('Other-Job')) 'An externally ended job must be completed and removed from the running jobs.'
}
finally {
    foreach ($child in $children) { try { if (-not $child.HasExited) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } } catch { } }
}
"PASS: operator stop after re-adoption, $($script:passed) offline checks."
