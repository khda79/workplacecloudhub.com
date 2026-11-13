#requires -Version 7.0
$ErrorActionPreference = 'Stop'
$module = Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') -Force -PassThru
$engine = (Get-Process -Id $PID).Path
$progress = [Collections.Generic.List[datetime]]::new()
$callback = { $progress.Add([datetime]::UtcNow); 'ignored output' }.GetNewClosure()
$slow = 'Start-Sleep -Seconds 3; ''SMARTM365_CAPABILITY_JSON:{"Ready":true,"Roles":[],"Detail":"ok"}'''
$result = & $module { param($engine,$slow,$callback) Invoke-IsolatedCapabilityProbe -Name EXO -EnginePath $engine -ScriptText $slow -TimeoutSeconds 8 -OnProgress $callback } $engine $slow $callback
if (@($result).Count -ne 1 -or -not $result.Ready -or $result.Detail -ne 'ok' -or $progress.Count -lt 3) { throw "Slow probe failed or did not signal progress: $($progress.Count)" }
$progress.Clear()
$timeout = & $module { param($engine,$callback) Invoke-IsolatedCapabilityProbe -Name EXO -EnginePath $engine -ScriptText 'Start-Sleep -Seconds 10' -TimeoutSeconds 3 -OnProgress $callback } $engine $callback
if ($timeout.Ready -or $timeout.Detail -notmatch 'timed out' -or $progress.Count -lt 3) { throw 'Timeout did not fail closed or provide progress' }
$progress.Clear()
$pathResult = & $module {
    param($callback)
    $original = $script:PathAccessProbeScript
    try {
        $script:PathAccessProbeScript = 'param($Checks,$Server,$Results) Start-Sleep -Seconds 3; $Results[0] = [pscustomobject]@{Ready=$true;Detail="test"}'
        Test-SmartM365OrchestratorPathAccess -Checks @([pscustomobject]@{Name='test';Path='test';Access='Read';PathType='Container'}) -ServerName 'test' -TimeoutSeconds 8 -OnProgress $callback
    }
    finally { $script:PathAccessProbeScript = $original }
} $callback
if (-not $pathResult.Ready -or $progress.Count -lt 3) { throw 'Slow path probe did not provide progress' }
$failedCallback = { throw 'simulated callback error' }
$result = & $module { param($engine,$slow,$callback) Invoke-IsolatedCapabilityProbe -Name EXO -EnginePath $engine -ScriptText $slow -TimeoutSeconds 8 -OnProgress $callback } $engine $slow $failedCallback
if (-not $result.Ready) { throw 'Progress callback error changed the probe result' }
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
$fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Update-OrchestratorServerCapabilities' }, $true)
. ([scriptblock]::Create($fn.Extent.Text))
$script:Settings = @{ DistributedSchedulingEnabled=$true; CapabilityRefreshMinutes=360; SharedDataFolderPath='x'; ExchangeOnlineOrganization='x'; CapabilityProbeMode='ReadOnly'; CapabilityProbeTimeoutSeconds=90; PrerequisitePathTimeoutSeconds=10; CertificateExpiryWarningDays=30; ElectionWeight=1; ServerJobPolicies=@{}; CapabilitiesPath='x' }
$script:Manifest = @{ OrderedJobs=@([pscustomobject]@{Enabled=$true; AssignmentMode='Elected'; RequiredCapabilities=@('EXO')}) }
$script:LastCapabilityRefreshUtc = [datetime]::MinValue
$script:CapabilityFailureRetryMinutes = 0
$script:heartbeatCalls = 0
$global:TenantId='x'; $global:AppId='x'; $global:Thumbprint='x'
function Get-OrchestratorTenantDataPathCheck { @() }
function Get-OrchestratorLocalServerState { 'Active' }
function Publish-OrchestratorEffectiveCapability { $script:ServerCapabilities=$script:ServerCapabilitiesBase }
function Update-OrchestratorHeartbeatDuringLongOperation { $script:heartbeatCalls++ }
function Write-OrchestratorLog { param($Message, $Level) }
function Get-SmartM365OrchestratorServerCapability {
    param($ServerName,$SharedDataFolderPath,$TenantId,$AppId,$CertificateThumbprint,$Organization,$ProbeMode,$ProbeTimeoutSeconds,$DataPathChecks,$PathTimeoutSeconds,$CertificateExpiryWarningDays,$OnProgress)
    & $OnProgress
    [pscustomobject]@{ Results=@([pscustomobject]@{Name='EXO';Ready=$false;ComponentMissing=$false;Detail='timeout'}); ReadyCapabilities=@(); PrerequisiteBlocks=@() }
}
Update-OrchestratorServerCapabilities
if ($script:heartbeatCalls -ne 1 -or $script:CapabilityFailureRetryMinutes -ne 10) { throw 'Resident callback was not connected or failure retry changed' }
$heartbeatFn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Update-OrchestratorHeartbeatDuringLongOperation' }, $true)
. ([scriptblock]::Create($heartbeatFn.Extent.Text))
$script:LastLongOperationHeartbeat = [datetime]::MinValue
$script:simulatedNow = [datetime]'2026-01-01T00:00:00'
$script:writes = 0
function Get-Date { $script:simulatedNow = $script:simulatedNow.AddSeconds(31); $script:simulatedNow }
function Write-OrchestratorHeartbeat { $script:writes++ }
$realCallback = { Update-OrchestratorHeartbeatDuringLongOperation }
$result = & $module { param($engine,$slow,$callback) Invoke-IsolatedCapabilityProbe -Name EXO -EnginePath $engine -ScriptText $slow -TimeoutSeconds 8 -OnProgress $callback } $engine $slow $realCallback
if (-not $result.Ready -or $script:writes -lt 2) { throw 'Long probe did not trigger repeated throttled heartbeat writes' }
'PASS: slow and timed-out capability probes report progress and preserve results'
