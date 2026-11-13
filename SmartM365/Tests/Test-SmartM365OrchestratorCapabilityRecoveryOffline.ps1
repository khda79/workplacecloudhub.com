#requires -Version 7.0
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$null, [ref]$null)
$fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Update-OrchestratorServerCapabilities' }, $true)
. ([scriptblock]::Create($fn.Extent.Text))
$script:Settings = @{ DistributedSchedulingEnabled=$true; CapabilityRefreshMinutes=360; SharedDataFolderPath='x'; ExchangeOnlineOrganization='x'; CapabilityProbeMode='ReadOnly'; CapabilityProbeTimeoutSeconds=90; PrerequisitePathTimeoutSeconds=10; CertificateExpiryWarningDays=30; ElectionWeight=1; ServerJobPolicies=@{}; CapabilitiesPath='x' }
$script:Manifest = @{ OrderedJobs=@([pscustomobject]@{Enabled=$true; AssignmentMode='Elected'; RequiredCapabilities=@('EXO')}) }
$script:LastCapabilityRefreshUtc = [datetime]::MinValue
$script:CapabilityFailureRetryMinutes = 0
$script:calls = 0
$script:ready = $false
$global:TenantId='x'; $global:AppId='x'; $global:Thumbprint='x'
function Get-OrchestratorTenantDataPathCheck { @() }
function Get-OrchestratorLocalServerState { 'Active' }
function Publish-OrchestratorEffectiveCapability { $script:ServerCapabilities=$script:ServerCapabilitiesBase }
function Write-OrchestratorLog { param($Message, $Level) }
function Get-SmartM365OrchestratorServerCapability {
    $script:calls++
    [pscustomobject]@{ Results=@([pscustomobject]@{Name='EXO'; Ready=$script:ready; ComponentMissing=$false; Detail='test'}); ReadyCapabilities=@($(if ($script:ready) {'EXO'})); PrerequisiteBlocks=@() }
}
function Assert-Equal($actual,$expected,$what) { if ($actual -ne $expected) { throw "$what : expected $expected, got $actual" } }
Update-OrchestratorServerCapabilities
Assert-Equal $script:CapabilityFailureRetryMinutes 10 'first retry'
Update-OrchestratorServerCapabilities
Assert-Equal $script:calls 1 'no premature probe'
$script:LastCapabilityRefreshUtc=[datetime]::UtcNow.AddMinutes(-11)
Update-OrchestratorServerCapabilities
Assert-Equal $script:CapabilityFailureRetryMinutes 20 'second retry'
$script:LastCapabilityRefreshUtc=[datetime]::UtcNow.AddMinutes(-21)
Update-OrchestratorServerCapabilities
Assert-Equal $script:CapabilityFailureRetryMinutes 30 'third retry'
$script:LastCapabilityRefreshUtc=[datetime]::UtcNow.AddMinutes(-31)
Update-OrchestratorServerCapabilities
Assert-Equal $script:CapabilityFailureRetryMinutes 30 'retry cap'
$script:ready=$true
$script:LastCapabilityRefreshUtc=[datetime]::UtcNow.AddMinutes(-31)
Update-OrchestratorServerCapabilities
Assert-Equal $script:CapabilityFailureRetryMinutes 0 'recovery'
Assert-Equal $script:calls 5 'probe count'
'PASS: capability recovery retry and backoff'
