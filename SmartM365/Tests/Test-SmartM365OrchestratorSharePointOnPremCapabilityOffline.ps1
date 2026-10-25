#Requires -Version 7.0
<#
.SYNOPSIS
Checks the SharePointOnPrem orchestrator capability and server policy, offline.
.DESCRIPTION
Verifies that every capability list accepts SharePointOnPrem, that the static probe reports
it from the snap-in registration, that the read-only probe uses the collector snap-in and a
read-only farm call in Windows PowerShell, and that an OnlyJobsRequiring SharePointOnPrem
policy keeps the other jobs away from that server in the election plan.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1') -Force
$script:Checks = 0
function Assert-Capability {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

# Every hard-coded capability list must stay identical and include SharePointOnPrem.
$expected = "'SharedRuntime', 'Graph', 'EXO', 'AD', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell'"
$orchestratorText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'))
$managementText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Management.psm1'))
$guiText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
$distributedText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1'))
Assert-Capability ([regex]::Matches($orchestratorText, [regex]::Escape('$validCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident job capability list.'
Assert-Capability ([regex]::Matches($orchestratorText, [regex]::Escape('$validJobPolicyCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident server policy capability list.'
Assert-Capability ($managementText.Contains('$script:ValidCapabilities = @(' + $expected + ')')) 'Management capability list.'
Assert-Capability ($guiText.Contains('<ComboBoxItem Content="SharePointOnPrem"/>')) 'GUI server policy choice.'

# Publication validation accepts the capability for jobs and server policies.
$jobs = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
$sharePointJob = ($jobs.Jobs[0] | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
$sharePointJob.Name = 'SharePointOnPrem-Sample'
$sharePointJob.RequiredCapabilities = @('SharedRuntime', 'SharePointOnPrem')
$jobs.Jobs = @($jobs.Jobs) + $sharePointJob
$jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $jobs
Assert-Capability ($jobsValidation.Valid -and -not @($jobsValidation.Errors | Where-Object { $_ -like '*SharePointOnPrem*' }).Count) "Jobs validation: $($jobsValidation.Errors -join '; ')"
$cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Cluster.json.template')
$cluster.ExpectedOrchestratorServers = @('SRV01', 'SP01')
$cluster.ServerJobPolicies = [pscustomobject]@{ SP01 = [pscustomobject]@{ OnlyJobsRequiring = @('SharePointOnPrem') } }
$clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $cluster
Assert-Capability $clusterValidation.Valid "Cluster policy validation: $($clusterValidation.Errors -join '; ')"
$cluster.ServerJobPolicies = [pscustomobject]@{ SP01 = [pscustomobject]@{ OnlyJobsRequiring = @('SharePointOnline') } }
Assert-Capability (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'Unknown policy capability is still rejected.'

# Static probe reports SharePointOnPrem from the snap-in registration.
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPCapability-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $root
    $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $root -ProbeMode Static
    $probe = @($static.Results | Where-Object Name -eq 'SharePointOnPrem')
    $registered = [bool](& (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoLogo -NoProfile -NonInteractive -Command "if (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) { 'yes' }")
    Assert-Capability ($probe.Count -eq 1 -and $probe[0].Ready -eq $registered) 'Static probe result matches the snap-in registration.'
    Assert-Capability (($static.ReadyCapabilities -contains 'SharePointOnPrem') -eq $registered) 'ReadyCapabilities follows the probe.'
    Assert-Capability (@(Get-ChildItem -LiteralPath $root -Recurse -File).Count -eq 0) 'Static probe leaves no file behind.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }

# Read-only probe: Windows PowerShell, collector snap-in, read-only farm call, bounded timeout.
$readOnly = [regex]::Match($distributedText, "(?s)\`$sharePointScript = @'(?<Body>.*?)'@\s*\r?\n\s*(?<Call>\`$results\.Add\(\(Invoke-IsolatedCapabilityProbe -Name 'SharePointOnPrem'[^\r\n]*)")
Assert-Capability $readOnly.Success 'Read-only SharePoint probe exists.'
Assert-Capability ($readOnly.Groups['Call'].Value -match '-EnginePath \$windowsPowerShellPath' -and $readOnly.Groups['Call'].Value -match '-TimeoutSeconds \$ProbeTimeoutSeconds') 'Probe runs in bounded Windows PowerShell.'
$body = $readOnly.Groups['Body'].Value
Assert-Capability ($body -match 'Add-PSSnapin Microsoft\.SharePoint\.PowerShell' -and $body -match 'Get-SPFarm' -and $body -match 'SMARTM365_CAPABILITY_JSON') 'Probe loads the collector snap-in and reads the farm.'
Assert-Capability ($body -notmatch '(?i)\b(Set|New|Remove|Add|Update|Install|Enable|Disable|Start|Stop|Mount|Dismount)-SP') 'Probe never changes the farm.'
$tokens = $null; $errors = $null
[void][Management.Automation.Language.Parser]::ParseInput($body, [ref]$tokens, [ref]$errors)
Assert-Capability ($errors.Count -eq 0) 'Probe script parses.'

# Election: a SharePointOnPrem-only server receives only SharePoint on-premises jobs.
$sample = @(
    [pscustomobject]@{ Name = 'SP-Content'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'SP-Content'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime', 'SharePointOnPrem'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('02:00') } }
    [pscustomobject]@{ Name = 'Cloud-A'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'Cloud-A'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime', 'Graph'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('03:00') } }
    [pscustomobject]@{ Name = 'Cloud-B'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'Cloud-B'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('04:00') } }
)
$capabilities = @(
    [pscustomobject]@{ ServerName = 'SRV01'; ReadyCapabilities = @('SharedRuntime', 'Graph'); GraphAppRoles = @() }
    [pscustomobject]@{ ServerName = 'SP01'; ReadyCapabilities = @('SharedRuntime', 'Graph', 'SharePointOnPrem'); GraphAppRoles = @() }
)
$plan = Get-SmartM365OrchestratorElectionPlan -Jobs $sample -ServerCapabilities $capabilities -ServerWeights @{ SRV01 = 1.0; SP01 = 5.0 } -ServerJobPolicies @{ SP01 = @{ OnlyJobsRequiring = @('SharePointOnPrem') } }
$owners = @{}; foreach ($assignment in @($plan.Assignments)) { $owners[[string]$assignment.JobName] = [string]$assignment.OwnerServer }
Assert-Capability ($owners['SP-Content'] -eq 'SP01') 'SharePoint on-premises job elected on the SharePoint server.'
Assert-Capability ($owners['Cloud-A'] -eq 'SRV01' -and $owners['Cloud-B'] -eq 'SRV01') 'Policy keeps other jobs off the SharePoint server despite its weight.'
$noSharePoint = Get-SmartM365OrchestratorElectionPlan -Jobs @($sample[0]) -ServerCapabilities @($capabilities[0]) -ServerWeights @{ SRV01 = 1.0 } -ServerJobPolicies @{}
Assert-Capability (@($noSharePoint.Assignments).Count -eq 0 -and @($noSharePoint.UnassignedGroups).Count -eq 1) 'Without a ready SharePointOnPrem server the job stays unassigned.'

# Forced Manual jobs: the resident launches them only where every required capability is ready.
$orchestratorAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$tokens, [ref]$errors)
foreach ($name in @('Get-ManualJobMissingCapabilities', 'Test-JobAllowedOnServer')) {
    $definition = $orchestratorAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Assert-Capability ($null -ne $definition) "$name exists."
    . ([scriptblock]::Create($definition.Extent.Text))
}
$template = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
foreach ($jobName in @('SharePointOnPrem-Infrastructure-Inventory', 'SharePointOnPrem-Content-Inventory')) {
    $job = @($template.Jobs | Where-Object Name -eq $jobName)[0]
    Assert-Capability ($job.AssignmentMode -eq 'Manual' -and @($job.DependsOn).Count -eq 0 -and @($job.RequiredCapabilities) -contains 'SharePointOnPrem') "$jobName is manual, independent and requires SharePointOnPrem."
    $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD') }
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job -AllowManual)) "$jobName refused by -Force without SharePointOnPrem."
    Assert-Capability ((Get-ManualJobMissingCapabilities -Job $job) -join ',' -eq 'SharePointOnPrem') "$jobName reports the missing capability."
    $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = @('SharedRuntime', 'SharePointOnPrem') }
    Assert-Capability (Test-JobAllowedOnServer -Job $job -AllowManual) "$jobName accepted by -Force on a SharePointOnPrem server."
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job)) "$jobName is never scheduled without -Force."
    $script:ServerCapabilities = $null
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job -AllowManual)) "$jobName fails closed when capabilities are unknown."
}
$script:ServerCapabilities = $null
Assert-Capability (Test-JobAllowedOnServer -Job ([pscustomobject]@{ Name = 'ManualNoCapability'; AssignmentMode = 'Manual'; RequiredCapabilities = @() }) -AllowManual) 'A manual job without required capabilities keeps its previous behaviour.'

"[{0}] PASS: SharePointOnPrem capability and policy, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
