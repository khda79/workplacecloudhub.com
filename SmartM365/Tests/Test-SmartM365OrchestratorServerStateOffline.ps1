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
