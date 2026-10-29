<#
.SYNOPSIS
Offline tests of the per-server Active/Inactive state (orchestrator 1.5.76): cluster rule and
validation, explicit migration of an existing cluster, election, launch gate, retry adoption,
reload with rebalance, heartbeat, Live view and GUI draft helpers.
.VERSION
1.0.0
.NOTES
Runs the real module functions and the real orchestrator and GUI functions extracted from the
scripts, with stubbed logs, in a temporary folder. Fictional server names only. No tenant
initialization, collectors, mail or publication outside the temporary folder.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$guiPath = Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'
$orchestratorAst = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
$guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$null)
function Get-FunctionText($Ast, [string]$Name) {
    $definition = $Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing function: $Name" }
    $definition.Extent.Text
}
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force

$originalComputerName = $env:COMPUTERNAME
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ServerState-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    # 1. Cluster rule: no ServerStates = every server Active; ServerStates present = missing entry is Inactive.
    $legacy = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B') }
    $states = Get-SmartM365OrchestratorServerStateTable -ClusterDocument $legacy
    Check ($states['SRV-A'] -eq 'Active' -and $states['SRV-B'] -eq 'Active') 'A cluster without ServerStates keeps every server Active.'
    Check (@(Get-SmartM365OrchestratorInactiveServers -ClusterDocument $legacy).Count -eq 0) 'No inactive server without ServerStates.'
    $explicit = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'srv-new'); ServerStates = [pscustomobject]@{ 'srv-a' = 'Active'; 'SRV-B' = 'Inactive' } }
    $states = Get-SmartM365OrchestratorServerStateTable -ClusterDocument $explicit
    Check ($states['SRV-A'] -eq 'Active' -and $states['SRV-B'] -eq 'Inactive' -and $states['SRV-NEW'] -eq 'Inactive') 'Explicit states are case-insensitive and a server without entry is Inactive.'
    Check ((@(Get-SmartM365OrchestratorInactiveServers -ClusterDocument $explicit) -join ',') -eq 'SRV-B,SRV-NEW') 'Inactive servers are listed sorted.'

    # 2. Validation.
    $validation = Test-SmartM365OrchestratorClusterDocument -Document $explicit
    Check ($validation.Valid -and @($validation.Warnings | Where-Object { $_ -match "SRV-NEW' has no ServerStates entry" }).Count -eq 1) 'A server without entry is valid with a warning.'
    $bad = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); ServerStates = [pscustomobject]@{ 'SRV-A' = 'Paused' } }
    Check (-not (Test-SmartM365OrchestratorClusterDocument -Document $bad).Valid) 'An unknown state value is rejected.'
    $notObject = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); ServerStates = 'Active' }
    Check (-not (Test-SmartM365OrchestratorClusterDocument -Document $notObject).Valid) 'ServerStates must be an object.'
    $allInactive = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); ServerStates = [pscustomobject]@{ 'SRV-A' = 'Inactive'; 'SRV-GONE' = 'Active' } }
    $validation = Test-SmartM365OrchestratorClusterDocument -Document $allInactive
    Check ($validation.Valid -and @($validation.Warnings | Where-Object { $_ -match 'No server is Active' }).Count -eq 1 -and @($validation.Warnings | Where-Object { $_ -match "non-expected server 'SRV-GONE'" }).Count -eq 1) 'All inactive and orphan entries are warnings.'
    $jobsDocument = [pscustomobject]@{ Jobs = @([pscustomobject]@{ Name = 'Pinned-B'; Enabled = $true; AssignmentMode = 'Pinned'; AllowedServers = @('SRV-B') }) }
    $consistency = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $jobsDocument -ClusterDocument $explicit
    Check ($consistency.Valid -and @($consistency.Warnings | Where-Object { $_ -match "Pinned-B.*SRV-B' is Inactive" }).Count -eq 1) 'A job pinned to an inactive server is a warning.'
    $template = Get-Content -LiteralPath (Join-Path $base 'Orchestrator-Cluster.json.template') -Raw | ConvertFrom-Json
    Check ($template.PSObject.Properties['ServerStates'] -and @($template.ServerStates.PSObject.Properties).Count -eq 0 -and (Test-SmartM365OrchestratorClusterDocument -Document $template).Valid) 'The cluster template carries an empty ServerStates object.'

    # 3. Explicit migration of an existing cluster: declared servers stay Active, jobs file untouched.
    $shared = Join-Path $root 'Orchestrator'
    $paths = Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $shared
    New-Item -ItemType Directory -Path $paths.ConfigFolderPath, $paths.VersionsFolderPath, $paths.AuditFolderPath -Force | Out-Null
    $jobsText = '{"SchemaVersion":1,"Jobs":[]}' + [Environment]::NewLine
    [IO.File]::WriteAllText($paths.JobsPath, $jobsText, [Text.UTF8Encoding]::new($false))
    $jobsHashBefore = Get-SmartM365OrchestratorFileHash $paths.JobsPath
    Write-SmartM365OrchestratorJsonAtomically -Path $paths.ClusterPath -Document ([pscustomobject][ordered]@{ SchemaVersion = 1; ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{} })
    $migration = Invoke-SmartM365OrchestratorServerStatesMigration -SharedDataFolderPath $shared
    $migrated = Read-SmartM365OrchestratorJson $paths.ClusterPath
    Check ($migration.Migrated -and -not $migration.Error -and $migration.VersionId) "The migration must publish a version ($($migration.Error))."
    Check ($migrated.ServerStates.'SRV-A' -eq 'Active' -and $migrated.ServerStates.'SRV-B' -eq 'Active') 'Migrated servers are explicitly Active.'
    Check ((Get-SmartM365OrchestratorFileHash $paths.JobsPath) -eq $jobsHashBefore) 'The migration never rewrites the jobs file.'
    Check (@(Get-ChildItem -LiteralPath $paths.VersionsFolderPath -Directory).Count -eq 1 -and (Test-Path -LiteralPath $paths.AuditPath)) 'The migration is versioned and audited.'
    $again = Invoke-SmartM365OrchestratorServerStatesMigration -SharedDataFolderPath $shared
    Check (-not $again.Migrated -and -not $again.Error) 'A second migration does nothing.'
    $migrated.ExpectedOrchestratorServers = @($migrated.ExpectedOrchestratorServers) + 'SRV-NEW'
    Check ((Get-SmartM365OrchestratorServerStateTable -ClusterDocument $migrated)['SRV-NEW'] -eq 'Inactive') 'A server added after the migration is Inactive.'
    $copy = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A') }
    Check ((Initialize-SmartM365OrchestratorServerStates -ClusterDocument $copy) -and $copy.ServerStates.'SRV-A' -eq 'Active' -and -not (Initialize-SmartM365OrchestratorServerStates -ClusterDocument $copy)) 'Initialize makes the states explicit once.'

    # 4. Election: an inactive server is never a candidate; its published state is enough.
    function New-Capability([string]$Name, [string]$State = '') {
        $capability = [pscustomobject]@{ ServerName = $Name; ReadyCapabilities = @('SharedRuntime', 'Graph'); GraphAppRoles = @(); BlockedJobs = @() }
        if ($State) { $capability | Add-Member -NotePropertyName ServerState -NotePropertyValue $State }
        $capability
    }
    $electionJobs = @(1..6 | ForEach-Object { [pscustomobject]@{ Name = "Job$_"; Enabled = $true; AssignmentMode = 'Elected'; RequiredCapabilities = @('SharedRuntime'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 10; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00') } } })
    $capabilities = @((New-Capability 'SRV-A'), (New-Capability 'SRV-B'), (New-Capability 'SRV-C'))
    $plan = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities $capabilities -InactiveServers @('SRV-C')
    Check (@($plan.Assignments | Where-Object OwnerServer -eq 'SRV-C').Count -eq 0 -and @($plan.Assignments).Count -eq 6) 'An inactive server gets no elected job.'
    Check ((@($plan.EligibleServers) -join ',') -eq 'SRV-A,SRV-B' -and (@($plan.InactiveServers) -join ',') -eq 'SRV-C') 'The plan lists the eligible and inactive servers.'
    $plan = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities @((New-Capability 'SRV-A'), (New-Capability 'SRV-B' 'Inactive'))
    Check (@($plan.Assignments | Where-Object OwnerServer -ne 'SRV-A').Count -eq 0) 'A server publishing Inactive is excluded even before the cluster is reloaded.'
    $plan = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities @((New-Capability 'SRV-A')) -InactiveServers @('SRV-A')
    Check (@($plan.Assignments).Count -eq 0 -and @($plan.UnassignedGroups).Count -eq 6 -and $plan.UnassignedGroups[0].Reason -match 'active server') 'Without an active server every elected job is unassigned.'
    $stable = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities $capabilities -InactiveServers @('SRV-C')
    Check (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $stable -ServerCapabilities $capabilities -InactiveServers @('SRV-C')) 'A live inactive server does not force a rebalance at every refresh.'
    Check (-not (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $stable -ServerCapabilities $capabilities)) 'Activating a server changes the eligible set and rebalances.'
    $preserved = Get-SmartM365OrchestratorElectionPlan -Jobs $electionJobs -ServerCapabilities $capabilities -PreviousPlan $stable -PreservePreviousOwners -InactiveServers @('SRV-B')
    Check (@($preserved.Assignments | Where-Object OwnerServer -eq 'SRV-B').Count -eq 0) 'An owner set Inactive is never preserved.'

    # 5. Resident launch gate and retry adoption (real orchestrator functions).
    function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') $script:Logs.Add("[$Level] $Message") }
    function Get-ElectedJobOwner { param([string]$JobName) 'SRV-A' }
    function Get-JobEffectiveAllowedServers { param($Job) @($Job.AllowedServers) }
    function Get-ManualJobMissingCapabilities { param($Job) @() }
    function Test-JobSelected { param([string]$JobName) $true }
    function Get-OrchestratorRetryOffers { throw 'An inactive server must not read retry offers.' }
    foreach ($name in @('Test-OrchestratorLocalServerActive', 'Get-OrchestratorLocalServerState', 'Test-JobAllowedOnServer', 'Test-JobAllowedOnNamedServer', 'Get-OrchestratorInactiveServers', 'Invoke-RetryHandoffAdoption', 'Set-OrchestratorCentralClusterSettings', 'Write-ServerAllowlistSummary')) {
        . ([scriptblock]::Create((Get-FunctionText $orchestratorAst $name)))
    }
    $script:Logs = [Collections.Generic.List[string]]::new()
    $script:LogReady = $true
    $script:State = @{ Jobs = @{} }
    $script:ElectionPlan = [pscustomobject]@{ PlanId = 'p' }
    $script:LastCapabilityRefreshUtc = [datetime]::UtcNow
    $script:ElectionRebalanceRequested = $false
    $script:Settings = [pscustomobject]@{
        ExpectedOrchestratorServers = @(); ServerJobPolicies = @{}; ServerStates = @{}; ServerStatesExplicit = $false; ElectionWeight = 1.0; DistributedSchedulingEnabled = $true
    }
    $env:COMPUTERNAME = 'SRV-A'
    $legacyJob = [pscustomobject]@{ Name = 'Legacy'; AssignmentMode = 'Legacy'; AllowedServers = @() }
    $pinnedJob = [pscustomobject]@{ Name = 'Pinned'; AssignmentMode = 'Pinned'; AllowedServers = @('SRV-A') }
    $electedJob = [pscustomobject]@{ Name = 'Elected'; AssignmentMode = 'Elected'; AllowedServers = @() }
    $manualJob = [pscustomobject]@{ Name = 'Manual'; AssignmentMode = 'Manual'; AllowedServers = @() }
    Check ((Test-JobAllowedOnServer -Job $legacyJob) -and (Test-JobAllowedOnServer -Job $pinnedJob) -and (Test-JobAllowedOnServer -Job $electedJob) -and (Test-JobAllowedOnServer -Job $manualJob -AllowManual)) 'Without explicit states (local configuration) the server is Active.'

    Set-OrchestratorCentralClusterSettings -ClusterDocument ([pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; ServerStates = [pscustomobject]@{ 'SRV-A' = 'Inactive'; 'SRV-B' = 'Active' } })
    Check (-not (Test-OrchestratorLocalServerActive) -and (Get-OrchestratorLocalServerState) -eq 'Inactive') 'The cluster state applies to the local server.'
    Check (-not (Test-JobAllowedOnServer -Job $legacyJob) -and -not (Test-JobAllowedOnServer -Job $pinnedJob) -and -not (Test-JobAllowedOnServer -Job $electedJob) -and -not (Test-JobAllowedOnServer -Job $manualJob -AllowManual)) 'An inactive server refuses legacy, pinned, elected and forced manual jobs.'
    $script:State.Jobs['Elected'] = @{ PendingRetry = @{ HandoffFrom = 'SRV-B' } }
    Check (-not (Test-JobAllowedOnServer -Job $electedJob)) 'An adopted retry does not bypass the inactive state.'
    Invoke-RetryHandoffAdoption -Now (Get-Date)
    Check ((@(Get-OrchestratorInactiveServers) -join ',') -eq 'SRV-A') 'The resident lists the inactive servers for the election.'
    Check (-not (Test-JobAllowedOnNamedServer -Job $legacyJob -ServerName 'srv-a') -and (Test-JobAllowedOnNamedServer -Job $legacyJob -ServerName 'SRV-B')) 'Peer job monitoring expects no job from an inactive peer.'
    $script:Logs.Clear()
    Write-ServerAllowlistSummary
    Check (@($script:Logs | Where-Object { $_ -match '^\[WARN\] Server inactive: no job will be launched on SRV-A' }).Count -eq 1) 'The resident logs clearly that it launches no job.'

    $script:Logs.Clear()
    $script:LastCapabilityRefreshUtc = [datetime]::UtcNow
    Set-OrchestratorCentralClusterSettings -ClusterDocument ([pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; ServerStates = [pscustomobject]@{ 'SRV-A' = 'Active'; 'SRV-B' = 'Active' } })
    Check ((Test-OrchestratorLocalServerActive) -and (Test-JobAllowedOnServer -Job $legacyJob)) 'Activation takes effect at the reload.'
    Check ($script:ElectionRebalanceRequested -and $script:LastCapabilityRefreshUtc -eq [datetime]::MinValue) 'A local state change republishes the capabilities and rebalances.'
    Check (@($script:Logs | Where-Object { $_ -match 'Server state changed in the cluster configuration: SRV-A=Active' }).Count -eq 1 -and @($script:Logs | Where-Object { $_ -match 'SRV-A is now Active' }).Count -eq 1) 'The state change is logged.'
    $script:ElectionRebalanceRequested = $false
    Set-OrchestratorCentralClusterSettings -ClusterDocument ([pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-NEW'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; ServerStates = [pscustomobject]@{ 'SRV-A' = 'Active'; 'SRV-B' = 'Active' } })
    Check ((Test-OrchestratorLocalServerActive) -and (@(Get-OrchestratorInactiveServers) -join ',') -eq 'SRV-NEW' -and -not $script:ElectionRebalanceRequested) 'Adding a new server leaves the local state and keeps the new one Inactive.'
    $env:COMPUTERNAME = 'SRV-UNDECLARED'
    Check (-not (Test-OrchestratorLocalServerActive)) 'With explicit states, an undeclared server is Inactive.'
    $env:COMPUTERNAME = $originalComputerName

    $orchestratorText = Get-Content -LiteralPath $orchestratorPath -Raw
    Check ($orchestratorText -match 'ServerStateProtocol = 1' -and $orchestratorText -match 'ServerState = Get-OrchestratorLocalServerState') 'The heartbeat publishes the server state.'
    Check ($orchestratorText -match "NotePropertyName ServerState -NotePropertyValue \(Get-OrchestratorLocalServerState\)") 'The capability document publishes the server state.'
    Check ($orchestratorText -match '-InactiveServers \$inactiveServers -NowUtc') 'The election receives the inactive servers.'

    # 6. Live view: Jobs column (Mode) from the cluster, applying while the resident has not reloaded.
    $liveRoot = Join-Path $root 'Live'
    foreach ($server in @('SRV-A', 'SRV-B', 'SRV-C')) {
        New-Item -ItemType Directory -Path (Join-Path $liveRoot $server) -Force | Out-Null
        # Every resident still applies Active: SRV-B and SRV-C have not reloaded the new configuration.
        $heartbeat = [ordered]@{ Timestamp = (Get-Date).ToUniversalTime().ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.76'; ServerStateProtocol = 1; ServerState = 'Active'; RunningJobs = @(); PendingJobs = @() }
        [IO.File]::WriteAllText((Join-Path $liveRoot "$server\Orchestrator-Heartbeat.json"), ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    }
    $liveCluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C'); PeerHeartbeatStaleMinutes = 5; ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; ServerStates = [pscustomobject]@{ 'SRV-A' = 'Active'; 'SRV-B' = 'Inactive' } }
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $liveRoot -ClusterDocument $liveCluster
    $modes = @{}; foreach ($row in @($operations.Servers)) { $modes[[string]$row.Server] = [string]$row.Mode }
    Check ($modes['SRV-A'] -eq 'Active' -and $modes['SRV-B'] -eq 'Inactive (applying)' -and $modes['SRV-C'] -eq 'Inactive (applying)') "Live shows the configured state and a pending reload ($($modes.Values -join ','))."
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $liveRoot -ClusterDocument ([pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); PeerHeartbeatStaleMinutes = 5 })
    Check ([string]$operations.Servers[0].Mode -eq 'Active') 'Live shows Active for an older cluster document.'
    $status = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $liveRoot -ClusterDocument $liveCluster)
    Check (@($status | Where-Object { $_.Server -eq 'SRV-B' -and $_.State -eq 'Inactive' }).Count -eq 1 -and @($status | Where-Object { $_.Server -eq 'SRV-A' -and $_.State -eq 'Active' }).Count -eq 1) 'The Servers grid rows carry the state.'

    # 7. GUI draft helpers: a new server is added Inactive; the state is written explicitly.
    foreach ($name in @('Set-JsonProperty', 'Set-DraftServerState', 'Add-ServerToDraft')) { . ([scriptblock]::Create((Get-FunctionText $guiAst $name))) }
    $script:DraftCluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{} }
    $added = Add-ServerToDraft -Server ' srv-new '
    Check ($added -eq 'SRV-NEW' -and $script:DraftCluster.ServerStates.'SRV-NEW' -eq 'Inactive' -and $script:DraftCluster.ServerStates.'SRV-A' -eq 'Active' -and $script:DraftCluster.ServerStates.'SRV-B' -eq 'Active') 'Add server: new server Inactive, existing servers made explicitly Active.'
    Set-DraftServerState -Server 'srv-new' -State 'Active'
    Check ($script:DraftCluster.ServerStates.'SRV-NEW' -eq 'Active' -and (Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster).Valid) 'The GUI can activate a server.'
    $duplicate = $false; try { [void](Add-ServerToDraft -Server 'SRV-A') } catch { $duplicate = $true }
    Check $duplicate 'A duplicate server is refused.'
    $guiText = Get-Content -LiteralPath $guiPath -Raw
    Check ($guiText -match 'x:Name="ServerActiveCheck"' -and $guiText -match 'Header="State" Binding="\{Binding State\}" Width="70"' -and $guiText -match 'Header="Jobs" Binding="\{Binding Mode\}"') 'The GUI shows the state (Servers grid, Selected server, Live).'
    Check ($guiText -match 'Initialize-SmartM365OrchestratorServerStates -ClusterDocument \$script:DraftCluster' -and $guiText -match 'SMOKE_TEST_SERVER_STATE_OK') 'The GUI draft uses explicit states and the smoke test covers them.'
}
finally {
    $env:COMPUTERNAME = $originalComputerName
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ("PASS Test-SmartM365OrchestratorServerStateOffline: {0} checks" -f $script:passed)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCimgtmbBaiKrhs
# wsYYpa2MpBYRRS5MXC+bJ+5PVxowBaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH/AogAqgXe3YG9xkGR7YkPpUmGA0wDnQ5dWP9xb8vgjMA0GCSqG
# SIb3DQEBAQUABIIBgFQm1aci6DGk4Hu7ifOY3EJnxOyWpSGqyhdYGSIKFvPI35mR
# 3N1s2BqiVSRKJ+yCej7Gvx8lU5ej4vGOcKW+2i6FENsBMfKvGwpeS5beUGi3NCrl
# a4tJkrU3QfqZzLZum2gfJTDDACOLe4kRb/RURHyOJp7Y9bv0n6x2MTnfpT2I6dvH
# wf4j2xTIIqe6Yyxxp80j3biXcZJrwKaTDjGdXX+Ed93t+Rcx7Ogjs9VaDfm+5V9C
# +AKIjBcqpfpcHK1X0VO741twb1VzBhujkNE8AvT3Cn11gI05tYntN5f/UYI+sxub
# RqRi50JZpZ07bonIxUh+pOjz7pSbAGRrRmadB9qX7oZ7++2peb0cmWMIV+licaOE
# uwBGM3CMbHAO91lsvZcc9kxi8jQwUN2dgNv+dbJVFD7tGEQ9bXAO1jjiji077bE0
# 4MYcMvLHArae5irAZyiuk9dt+ERm7MAYltM+8jBjQEGhvuZMo2UhULQMkIBG/QK8
# 54FNHIXJ3FJcjFFlO6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMjAw
# MjFaMC8GCSqGSIb3DQEJBDEiBCDTfNL1y8vANG1CABOdsY4OtyBQ6Yx9H30ou5ic
# SQsYiDANBgkqhkiG9w0BAQEFAASCAgA6EX9JtzZ4FXvdSjKN2AUEWmkNedGBDH4p
# gqRFHvfEdWQTneo46rW6loY+Lw9ldRUnjLmLROwKoXYF6HfeC4WgCI2rTdO68aKN
# NlrENOqCBriPTsiLumxGcR/+F2QUJKYNPdx56ROuFZ7ZR+mC/xIVjj6P26pkuz7/
# aw0cVOFmBG0WcMWjrk7lI1gUQr8jVugiblnubiZbHXEP0o+5Vs0vjSiZKCT9dW6s
# 3v+h2g+rqhp6n8NJ7Z+4b1bmBFUh8i5kMcYaF2HOMDvyQslprWtqo6WBQZC6s4rm
# KkyFX0dvDBeOfIWFJE7UIyzvmDVvzA4EDYurd5dBaZGTJXJ383XauISdA/at0cwh
# Ue7WzRnE6V93YYDtwIb3XPz0dpM7zxL8Eh1QJoozDjBoKSM93+5lJ1yod5Plgh/m
# 1/rYQvBwn9bZ7j6oXHBBnq28qdR/wEV6BD89Rvr1PKPLDzGqByv0olQlAZFN4eL8
# DM0FD4lh5ECPOc86EaL6S4rVEdo47dkApUR3/Y50MvA2H3Ut3huctOt2nkoYb8kE
# TkyBaMaR0uTRwRLt4hdEOHKSQH+0D9zksUA2kTABJBhAq8nx4Nw2p//HYo4sig6m
# bjxdiK+Zhgvh9D35RNUoMoEqdyMxWvomJBaBex2kAM++br1n66a2ARBKynKuV42A
# ohfYYRW5kQ==
# SIG # End signature block
