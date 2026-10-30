<#
.SYNOPSIS
Offline tests of the peer monitoring alerts (orchestrator 1.5.78): unconfirmed observations logged as
INFO, every issue confirmed on its own, one alert and one recovery mail per issue for the cluster.
.VERSION
1.0.0
.NOTES
Two observer servers run the real orchestrator functions, extracted from the script, in separate
module scopes with a shared temporary folder. The peer snapshot and the mail are stubbed: no mail is
sent and no real peer is read.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}
$jsonModule = Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1'
$functions = @('Invoke-OrchestratorPeerMonitoring', 'Get-OrchestratorPeerMonitoringState', 'Format-OrchestratorPeerIssue', 'Reset-OrchestratorPeerMonitoringIncident',
    'Get-OrchestratorPeerAlertPaths', 'Read-OrchestratorPeerAlertState', 'Save-OrchestratorPeerAlertState', 'ConvertFrom-OrchestratorPeerMonitorTime',
    'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock')
$stubs = @'
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add("[$Level] $Message") }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $script:Logs.Add("[WARN] $Message") }
function Write-FileAtomically { param($Path, $Content) [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false)) }
function Get-OrchestratorPeerHealthSnapshot { param($Now) [pscustomobject]@{ CheckedAt = $Now; Issues = @($global:PeerTestIssues); Peers = @() } }
function Send-OrchestratorPeerMonitoringEmail { param($Snapshot, [switch]$Recovery, $PreviousSummary, $Category = 'Jobs')
    $global:PeerTestMails.Add([pscustomobject]@{ Observer = $env:COMPUTERNAME; Recovery = [bool]$Recovery; Category = $Category; Keys = @($Snapshot.Issues | ForEach-Object { [string]$_.Key }) })
    $true }
'@
$functionText = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
$originalComputer = $env:COMPUTERNAME
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PeerAlert-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
function New-Observer([string]$Name) {
    $module = New-Module -Name ('Observer-' + $Name) -ScriptBlock ([scriptblock]::Create($functionText))
    & $module {
        param($Root, $JsonModule)
        Import-Module $JsonModule -Force
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:State = @{ Jobs = @{} }
        $script:LastPeerMonitoringCheckTime = [datetime]::MinValue
        $script:Settings = [pscustomobject]@{
            SharedDataFolderPath = $Root; PeerMonitoringEnabled = $true; ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C')
            PeerMonitoringCheckIntervalSeconds = 60; PeerMonitoringConfirmationChecks = 2; PeerAlertReminderMinutes = 240; PeerAlertMailRetryMinutes = 15
            PeerRecoveryEmailEnabled = $true; AtomicWriteRetrySeconds = 5
        }
    } $root $jsonModule
    $module
}
function Invoke-Check([string]$Server, $Module, [datetime]$Now) {
    $env:COMPUTERNAME = $Server
    try { & $Module { param($n) $script:Logs.Clear(); Invoke-OrchestratorPeerMonitoring -Now $n; @($script:Logs) } $Now }
    finally { $env:COMPUTERNAME = $originalComputer }
}
function New-Issue([string]$Key, [string]$Server, [string]$Type, [string]$Job = '') { [pscustomobject]@{ Key = $Key; Server = $Server; Job = $Job; Type = $Type } }

try {
    $global:PeerTestMails = [Collections.Generic.List[object]]::new()
    $a = New-Observer 'SRV-A'; $b = New-Observer 'SRV-B'
    $t = [datetime]'2026-10-30T08:00:00'
    $stale = New-Issue 'HeartbeatStale|SRV-C' 'SRV-C' 'HeartbeatStale'
    $job = New-Issue 'JobNotStarted|SRV-C|Job-1' 'SRV-C' 'JobNotStarted' 'Job-1'

    # 1. First observation: candidate, INFO only, no mail.
    $global:PeerTestIssues = @($stale)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 0) 'An unconfirmed issue sends no mail.'
    Check (@(@($logsA) + @($logsB) | Where-Object { $_ -like '`[INFO`] Peer monitoring candidate issue: SRV-C: HeartbeatStale (1/2)' }).Count -eq 2 -and @(@($logsA) + @($logsB) | Where-Object { $_ -like '`[WARN`]*' }).Count -eq 0) 'Candidate issues are INFO, never WARN.'

    # 2. Confirmed: the first observer sends, the second only acknowledges.
    $t = $t.AddMinutes(1)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 1 -and $global:PeerTestMails[0].Observer -eq 'SRV-A' -and $global:PeerTestMails[0].Category -eq 'Server') 'A confirmed issue sends one alert for the cluster.'
    Check (@($logsB | Where-Object { $_ -like '*alert already sent by another server: SRV-C: HeartbeatStale (by SRV-A)*' }).Count -eq 1) 'The second observer logs that the alert was already sent.'
    Check ((& $b { $script:State.PeerMonitoring.ActiveIssueKey }) -eq 'HeartbeatStale|SRV-C') 'The second observer still tracks the incident for its recovery.'
    $shared = Get-Content -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.json') -Raw | ConvertFrom-Json
    Check ($shared.Issues.'HeartbeatStale|SRV-C'.Observer -eq 'SRV-A') 'The shared state records who sent the alert.'

    # 3. A new issue does not restart the count of the confirmed one.
    $global:PeerTestIssues = @($stale, $job)
    $t = $t.AddMinutes(1)
    $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    Check ($global:PeerTestMails.Count -eq 1 -and @($logsA | Where-Object { $_ -like '*candidate issue: SRV-C/Job-1: JobNotStarted (1/2)' }).Count -eq 1 -and @($logsA | Where-Object { $_ -like '*HeartbeatStale (1/2)*' }).Count -eq 0) 'The confirmed issue keeps its count while a new one is a candidate.'
    $t = $t.AddMinutes(1)
    $null = Invoke-Check 'SRV-B' $b $t; $null = Invoke-Check 'SRV-A' $a $t
    Check ($global:PeerTestMails.Count -eq 2 -and $global:PeerTestMails[1].Observer -eq 'SRV-B' -and $global:PeerTestMails[1].Category -eq 'Jobs' -and (@($global:PeerTestMails[1].Keys) -join ',') -eq 'JobNotStarted|SRV-C|Job-1') 'Only the newly confirmed issue is alerted, once.'

    # 4. Recovery: one mail for the cluster, then the other observer closes silently.
    $global:PeerTestIssues = @()
    foreach ($step in 1..2) {
        $t = $t.AddMinutes(1)
        $logsA = Invoke-Check 'SRV-A' $a $t; $logsB = Invoke-Check 'SRV-B' $b $t
    }
    $recoveries = @($global:PeerTestMails | Where-Object Recovery)
    Check ($recoveries.Count -eq 1 -and $recoveries[0].Observer -eq 'SRV-A') 'One recovery mail for the cluster.'
    Check (@($logsB | Where-Object { $_ -like '*recovery email was already sent by another server*' }).Count -eq 1 -and [string]::IsNullOrEmpty((& $b { $script:State.PeerMonitoring.ActiveIssueKey }))) 'The other observer closes its incident without a mail.'
    $shared = Get-Content -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.json') -Raw | ConvertFrom-Json
    Check (@($shared.Issues.PSObject.Properties).Count -eq 0) 'The recovery clears the shared alerts.'

    # 5. A flapping observation (seen, cleared, seen) never confirms.
    $mailsBefore = $global:PeerTestMails.Count
    foreach ($issuesNow in @(@($stale), @(), @($stale), @())) {
        $global:PeerTestIssues = $issuesNow
        $t = $t.AddMinutes(1)
        $null = Invoke-Check 'SRV-A' $a $t
    }
    Check ($global:PeerTestMails.Count -eq $mailsBefore) 'A flapping observation that never lasts two checks sends nothing.'

    # 6. A stale shared entry (older than two reminder periods) does not block a new alert.
    $old = [ordered]@{ SchemaVersion = 1; Issues = [ordered]@{ 'HeartbeatStale|SRV-C' = [ordered]@{ SentAtUtc = $t.ToUniversalTime().AddHours(-9).ToString('o'); Observer = 'SRV-B' } } }
    [IO.File]::WriteAllText((Join-Path $root 'Orchestrator-PeerAlert.json'), ($old | ConvertTo-Json -Depth 5))
    $global:PeerTestIssues = @($stale)
    foreach ($step in 1..2) { $t = $t.AddMinutes(1); $null = Invoke-Check 'SRV-A' $a $t }
    Check ($global:PeerTestMails.Count -eq $mailsBefore + 1 -and $global:PeerTestMails[-1].Observer -eq 'SRV-A') 'An expired shared entry lets a new alert through.'
    Check (-not (Test-Path -LiteralPath (Join-Path $root 'Orchestrator-PeerAlert.lock'))) 'The shared lock is released.'
}
finally {
    $env:COMPUTERNAME = $originalComputer
    Remove-Variable -Name PeerTestMails, PeerTestIssues -Scope Global -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: peer monitoring alerts, $($script:passed) offline checks."
