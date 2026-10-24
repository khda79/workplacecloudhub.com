<#
.SYNOPSIS
Offline tests for the local ConcurrencyKey guard of the orchestrator: a scheduled occurrence that waits
for a key held on this server is logged once at INFO, forced requests keep their warning, and the guard
still runs before the dependency evaluation without changing which jobs are blocked.
.VERSION
1.0
.NOTES
Synthetic jobs only. No tenant initialization, collectors, scheduling, authentication or publication.
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
function Get-FunctionText([string]$Name) { $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true).Extent.Text }

# The guard keeps its place in the launch phase: after the occurrence is known, before dependencies.
$launchPhase = Get-FunctionText 'Invoke-LaunchPhase'
$guardIndex = $launchPhase.IndexOf('Test-OrchestratorLocalConcurrencyBlocked -Job $job -Occurrence $occurrence')
Check ($guardIndex -gt $launchPhase.IndexOf('if ($null -eq $occurrence) { continue }')) 'The local concurrency guard runs before the occurrence is known.'
Check ($guardIndex -lt $launchPhase.IndexOf('Get-OrchestratorExternalDependencyStatus') -and $guardIndex -lt $launchPhase.IndexOf('Get-OrchestratorSharedDependencyStatus') -and $guardIndex -lt $launchPhase.IndexOf('Enter-SmartM365OrchestratorConcurrencyLease')) 'The local concurrency guard no longer runs before dependencies and the shared lease.'
Check ($launchPhase.Substring($guardIndex, 200) -match '\)\s*\{\s*continue\s*\}') 'A blocked occurrence is no longer deferred.'
Check ($launchPhase -notmatch '\$blockingConcurrencyJobs') 'The launch phase keeps a second copy of the local concurrency guard.'

$stubs = @'
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') $script:Logs.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
'@
$guard = New-Module -ScriptBlock ([scriptblock]::Create((Get-FunctionText 'Write-OrchestratorRuntimeUpdateWarning') + "`n" + (Get-FunctionText 'Test-OrchestratorLocalConcurrencyBlocked') + "`n" + $stubs))
function Reset-Guard([object[]]$Jobs, [string[]]$Running = @()) {
    & $guard {
        param($JobList, $RunningNames)
        $script:Logs = [Collections.Generic.List[object]]::new()
        $script:RuntimeUpdateWarnings = @{}
        $script:Manifest = [pscustomobject]@{ JobsByName = @{} }
        foreach ($j in $JobList) { $script:Manifest.JobsByName[$j.Name] = $j }
        $script:RunningJobs = @{}
        foreach ($n in $RunningNames) { $script:RunningJobs[$n] = [pscustomobject]@{ Name = $n } }
    } $Jobs $Running
}
function Invoke-Guard($Job, [datetime]$Occurrence, [string[]]$Launched = @(), [switch]$IsForced, [datetime]$Now) {
    & $guard { param($j, $o, $l, $f, $n) Test-OrchestratorLocalConcurrencyBlocked -Job $j -Occurrence $o -LaunchedThisTick $l -IsForced:$f -Now $n } $Job $Occurrence $Launched ([bool]$IsForced) $Now
}
function Get-Logs { & $guard { @($script:Logs) } }

$full = [pscustomobject]@{ Name = 'Mailbox-Full'; ConcurrencyKey = 'MailboxKey' }
$fast = [pscustomobject]@{ Name = 'Mailbox-Fast'; ConcurrencyKey = 'MailboxKey' }
$other = [pscustomobject]@{ Name = 'Other-Job'; ConcurrencyKey = 'Other-Job' }
$occurrence = [datetime]'2026-01-10 00:15'
$now = [datetime]'2026-01-10 03:20'

# Key held by a running job: blocked, one INFO line naming the job, occurrence, key and holder.
Reset-Guard -Jobs @($full, $fast, $other) -Running @('Mailbox-Full')
Check (Invoke-Guard $fast $occurrence -Now $now) 'A job was not blocked by a running job sharing its ConcurrencyKey.'
$logs = Get-Logs
Check ($logs.Count -eq 1 -and $logs[0].Level -eq 'INFO') 'The scheduled wait was not logged exactly once at INFO.'
Check ($logs[0].Message -match 'Job Mailbox-Fast: occurrence 2026-01-10 00:15 waits because ConcurrencyKey ''MailboxKey'' is held by Mailbox-Full') 'The wait message does not name the occurrence, key and holder.'

# Later ticks of the same occurrence: still blocked, no new line.
foreach ($minutes in 1..30) { Check (Invoke-Guard $fast $occurrence -Now $now.AddMinutes($minutes)) 'The blocked occurrence was released while the key is held.' }
Check ((Get-Logs).Count -eq 1) 'The wait was logged again on a later tick.'

# Next occurrence still behind the key: one new line.
Check (Invoke-Guard $fast $occurrence.AddDays(1) -Now $now.AddDays(1)) 'The next occurrence was not blocked.'
Check ((Get-Logs).Count -eq 2) 'The next occurrence wait was not logged once.'

# Key released: not blocked, nothing logged.
Reset-Guard -Jobs @($full, $fast, $other)
Check (-not (Invoke-Guard $fast $occurrence -Now $now)) 'A job was blocked without any job holding its key.'
Check ((Get-Logs).Count -eq 0) 'An unblocked occurrence produced a log line.'

# Another key, or the job itself running, never blocks.
Reset-Guard -Jobs @($full, $fast, $other) -Running @('Other-Job', 'Mailbox-Fast')
Check (-not (Invoke-Guard $fast $occurrence -Now $now)) 'A job was blocked by an unrelated key or by itself.'
Check ((Get-Logs).Count -eq 0) 'An unrelated running job produced a log line.'

# Same tick: a job sharing the key that was just launched blocks too.
Reset-Guard -Jobs @($full, $fast, $other)
Check (Invoke-Guard $fast $occurrence -Launched @('Mailbox-Full') -Now $now) 'A job launched in the same tick with the same key did not block.'
Check (@(Get-Logs | Where-Object { $_.Level -eq 'INFO' -and $_.Message -match 'held by Mailbox-Full' }).Count -eq 1) 'The same-tick block was not logged.'

# Forced request: unchanged WARN, no INFO wait line.
Reset-Guard -Jobs @($full, $fast, $other) -Running @('Mailbox-Full')
Check (Invoke-Guard $fast $now -IsForced -Now $now) 'A forced request ignored the held key.'
Check (Invoke-Guard $fast $now.AddMinutes(1) -IsForced -Now $now.AddMinutes(1)) 'A forced request was released while the key is held.'
$logs = Get-Logs
Check ($logs.Count -eq 1 -and $logs[0].Level -eq 'WARN' -and $logs[0].Message -match 'forced launch deferred because ConcurrencyKey ''MailboxKey'' is already held by Mailbox-Full') 'The forced deferral warning changed or repeats.'

Write-Host ("PASS: {0} concurrency wait log offline checks. No collectors, scheduling or live publication." -f $script:passed)
