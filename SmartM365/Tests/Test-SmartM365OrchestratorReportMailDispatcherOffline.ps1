<#
.SYNOPSIS
Offline tests of the report mail queue dispatcher (orchestrator 1.5.93): one server sends the
shared report mail queue, the first Active live server by name; Inactive servers never send it.
.VERSION
1.0.0
.NOTES
Runs the real orchestrator functions extracted from the script with stubbed live servers, logs and
queue worker. No mail, tenant initialization, collector or shared folder is used.
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
foreach ($name in @('Test-OrchestratorLocalServerActive', 'Get-OrchestratorReportMailDispatcher', 'Invoke-OrchestratorReportMailQueue')) {
    . ([scriptblock]::Create((Get-FunctionText $name)))
}

$script:Logs = [Collections.Generic.List[object]]::new()
function Write-OrchestratorLog([string]$Message, [string]$Level = 'INFO') { $script:Logs.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
$script:QueueCalls = 0
function Invoke-SmartM365DeferredMail { $script:QueueCalls++; [pscustomobject]@{ Sent = 1; Pending = 0; Retry = 0; Busy = 0; Invalid = 0; Uncertain = 0 } }
$script:Live = @()
$script:LiveFailure = ''
function Get-LiveOrchestratorServerCapabilities {
    if ($script:LiveFailure) { throw $script:LiveFailure }
    foreach ($server in $script:Live) { [pscustomobject]@{ ServerName = $server } }
}
function Reset-Node([string]$Computer) {
    $env:COMPUTERNAME = $Computer
    $script:Logs.Clear(); $script:QueueCalls = 0
    foreach ($variable in @('ReportMailDispatcher', 'ReportMailDispatcherFallback', 'ReportMailDispatcherError')) { Remove-Variable -Name $variable -Scope Script -ErrorAction SilentlyContinue }
}
function New-Settings([hashtable]$States, [bool]$Explicit = $true, [bool]$Distributed = $true) {
    [pscustomobject]@{ DistributedSchedulingEnabled = $Distributed; ServerStates = $States; ServerStatesExplicit = $Explicit }
}

$originalComputer = $env:COMPUTERNAME
try {
    # Synthetic layout like production: SRV-A1, SRV-A2 and SRV-B1 Inactive, SRV-C1 and SRV-C2 Active.
    $states = @{ 'SRV-A1' = 'Inactive'; 'SRV-A2' = 'Inactive'; 'SRV-B1' = 'Inactive'; 'SRV-C1' = 'Active'; 'SRV-C2' = 'Active' }
    $script:Settings = New-Settings $states
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C1', 'SRV-C2')

    # 1. Exactly one server sends: the first Active live server by name.
    $senders = foreach ($server in $script:Live) {
        Reset-Node $server
        Invoke-OrchestratorReportMailQueue
        Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C1') "Every server must agree on SRV-C1 ($server)."
        if ($script:QueueCalls) { $server }
    }
    Check ((@($senders) -join ',') -eq 'SRV-C1') "Only SRV-C1 must send the queue (senders: $(@($senders) -join ','))."

    # 2. Inactive servers never send, even when first by name; a non-dispatcher logs why once.
    Reset-Node 'SRV-A1'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'An Inactive server must not send the queue.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1; this server does not send*').Count -eq 1) 'A non-dispatcher must log the dispatcher once.'

    # 3. The dispatcher sends at every tick and logs its role once, then the queue counts.
    Reset-Node 'SRV-C1'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2) 'The dispatcher must send the queue at every tick.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: this server*').Count -eq 1) 'The dispatcher role must be logged once.'
    Check (@($script:Logs | Where-Object Message -like 'Report mail queue: {"Sent":1*').Count -eq 2) 'Queue counts must still be logged.'

    # 4. Failover: a stale or deactivated dispatcher is replaced at the next tick, and given back.
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'SRV-C2 must wait while SRV-C1 is live.'
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C2')
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 1 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: this server*').Count -eq 1) 'SRV-C2 must take over when SRV-C1 is no longer live.'
    $script:Live = @('SRV-A1', 'SRV-A2', 'SRV-B1', 'SRV-C1', 'SRV-C2')
    $states['SRV-C1'] = 'Inactive'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2) 'SRV-C2 must keep the queue while SRV-C1 is Inactive.'
    $states['SRV-C1'] = 'Active'
    Invoke-OrchestratorReportMailQueue
    # Second SRV-C1 line of this node: the first one was logged before the failover.
    Check ($script:QueueCalls -eq 2 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1*').Count -eq 2) 'The queue must return to SRV-C1 once it is Active and live again.'

    # 5. Names are compared without case; a server absent from explicit states is Inactive.
    $script:Live = @('srv-c2', 'SRV-Z9')
    Reset-Node 'SRV-C2'
    Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C2') 'Live names must be compared without case, and an unlisted server must be Inactive with explicit states.'
    $script:Settings = New-Settings @{} $false
    Check ((Get-OrchestratorReportMailDispatcher) -eq 'SRV-C2') 'Without explicit states every live server is Active.'
    $script:Settings = New-Settings $states

    # 6. No Active live server: nobody sends and the wait is a warning.
    $script:Live = @('SRV-A1', 'SRV-B1')
    Reset-Node 'SRV-B1'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like '*dispatcher: none*' }).Count -eq 1) 'Without an Active live server nobody sends and a WARN is logged.'

    # 7. Unreadable peer documents: an Active server keeps sending, an Inactive one does not; one WARN each.
    $script:Live = @('SRV-C1', 'SRV-C2')
    $script:LiveFailure = 'Shared folder unreachable'
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 2 -and @($script:Logs | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like '*could not be determined (Shared folder unreachable); this Active server sends it*' }).Count -eq 1) 'An Active server must keep sending with one WARN when peers are unreadable.'
    Reset-Node 'SRV-A1'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object Message -like '*this Inactive server does not send it*').Count -eq 1) 'An Inactive server must not send when peers are unreadable.'
    $script:LiveFailure = ''
    Reset-Node 'SRV-C2'
    Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0 -and @($script:Logs | Where-Object Message -like 'Report mail queue dispatcher: SRV-C1*').Count -eq 1) 'Readable peers must restore the elected dispatcher.'

    # 8. Without distributed scheduling an Active server sends and an Inactive one does not.
    $script:Settings = New-Settings $states $true $false
    Reset-Node 'SRV-C2'; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 1) 'A standalone Active server must send the queue.'
    Reset-Node 'SRV-B1'; Invoke-OrchestratorReportMailQueue
    Check ($script:QueueCalls -eq 0) 'A standalone Inactive server must not send the queue.'

    # 9. The scheduler loop goes through the dispatcher only.
    $loopText = [IO.File]::ReadAllText($orchestratorPath)
    $direct = [regex]::Matches($loopText, '(?m)^(?!\s*#).*\bInvoke-SmartM365DeferredMail\b')
    Check ($direct.Count -eq 1 -and (Get-FunctionText 'Invoke-OrchestratorReportMailQueue').Contains('Invoke-SmartM365DeferredMail')) 'Only Invoke-OrchestratorReportMailQueue may call the queue worker.'
    Check ($loopText -match 'try \{\s*Invoke-OrchestratorReportMailQueue\s*\} catch \{') 'The scheduler loop must call the dispatcher inside its try block.'

    Write-Host ("PASS: report mail queue dispatcher, {0} offline checks." -f $script:passed)
}
finally {
    $env:COMPUTERNAME = $originalComputer
}
