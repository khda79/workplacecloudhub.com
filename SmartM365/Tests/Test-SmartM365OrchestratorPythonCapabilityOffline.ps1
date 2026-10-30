#Requires -Version 7.0
<#
.SYNOPSIS
Checks the Python orchestrator capability, offline.
.DESCRIPTION
Verifies that every capability list accepts Python, that the probe reads PreparedPythonCommand
where WorkplaceEvidence-Prepare reads it, that only Python 3.10+ is ready (a missing command or
the Microsoft Store alias is a missing component), that WorkplaceEvidence-Prepare requires Python
in the template and in an existing central manifest after the template sync, and that the
election keeps the job away from a server without Python. Fake Python commands are temporary
.cmd files; no real Python is needed.
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
$distributedModule = Get-Module -Name 'SmartM365.Orchestrator.Distributed'
$script:Checks = 0
function Assert-Python {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

# Every hard-coded capability list includes Python.
$expected = "'SharedRuntime', 'Graph', 'EXO', 'AD', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes', 'Python'"
$orchestratorText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'))
$managementText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Management.psm1'))
$guiText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
$distributedText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1'))
Assert-Python ([regex]::Matches($orchestratorText, [regex]::Escape('$validCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident job capability list.'
Assert-Python ([regex]::Matches($orchestratorText, [regex]::Escape('$validJobPolicyCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident server policy capability list.'
Assert-Python ($managementText.Contains('$script:ValidCapabilities = @(' + $expected + ')')) 'Management capability list.'
Assert-Python ($guiText.Contains('<ComboBoxItem Content="Python"/>')) 'GUI server policy choice.'

# Publication validation accepts Python for jobs and server policies.
$jobs = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
$jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $jobs
Assert-Python $jobsValidation.Valid "Template validation: $($jobsValidation.Errors -join '; ')"
$prepare = @($jobs.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')
Assert-Python ($prepare.Count -eq 1 -and (@($prepare[0].RequiredCapabilities) -join ',') -ceq 'SharedRuntime,Graph,Python') 'WorkplaceEvidence-Prepare requires SharedRuntime, Graph and Python in the template.'
$cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Cluster.json.template')
$cluster.ExpectedOrchestratorServers = @('SRV01', 'PY01')
$cluster.ServerJobPolicies = [pscustomobject]@{ PY01 = [pscustomobject]@{ OnlyJobsRequiring = @('Python') } }
$clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $cluster
Assert-Python $clusterValidation.Valid "Cluster policy validation: $($clusterValidation.Errors -join '; ')"
$prepare[0].RequiredCapabilities = @('SharedRuntime', 'Python3')
Assert-Python (-not (Test-SmartM365OrchestratorJobsDocument -Document $jobs).Valid) 'Unknown capability is still rejected.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-PyCapability-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $root
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'WindowsApps')
    $fakes = @{
        Python312 = @('@echo 3.12.1')
        Python39 = @('@echo 3.9.18')
        NoVersion = @('@echo Python was not found. 1>&2', '@exit /b 9009')
        Slow = @('@ping -n 6 127.0.0.1 >nul', '@echo 3.12.1')
    }
    foreach ($name in $fakes.Keys) { [IO.File]::WriteAllLines((Join-Path $root "$name.cmd"), [string[]]$fakes[$name]) }
    $storeAlias = Join-Path $root 'WindowsApps\python.cmd'
    [IO.File]::WriteAllLines($storeAlias, [string[]]@('@echo Python was not found; run without arguments to install from the Microsoft Store. 1>&2', '@exit /b 9009'))

    # Version probe: 3.10+ ready; missing command and Store alias are missing components; other failures are warnings.
    $probe = { param($Command, $Timeout) & $distributedModule { param($c, $t) Get-PythonCapabilityResult -Command $c -TimeoutSeconds $t } $Command $Timeout }
    $result = & $probe (Join-Path $root 'Python312.cmd') 10
    Assert-Python ($result.Name -eq 'Python' -and $result.Ready -and $result.Detail -like 'Python 3.12.1 at *') "Python 3.12 is ready: $($result.Detail)"
    $result = & $probe (Join-Path $root 'Python39.cmd') 10
    Assert-Python (-not $result.Ready -and $result.Detail -like '*3.9.18*below 3.10*' -and -not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "Python 3.9 is a warning: $($result.Detail)"
    $result = & $probe (Join-Path $root 'NoVersion.cmd') 10
    Assert-Python (-not $result.Ready -and $result.Detail -like '*gave no Python version (exit code 9009)*' -and -not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "No version is a warning: $($result.Detail)"
    $result = & $probe $storeAlias 10
    Assert-Python (-not $result.Ready -and $result.Detail -like '*Microsoft Store alias*' -and (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "Store alias is a missing component: $($result.Detail)"
    $result = & $probe 'smartm365-no-such-python-command' 10
    Assert-Python (-not $result.Ready -and (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "Missing command is a missing component: $($result.Detail)"
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $result = & $probe (Join-Path $root 'Slow.cmd') 1
    Assert-Python (-not $result.Ready -and $result.Detail -like '*did not answer within 1 seconds*' -and $timer.Elapsed.TotalSeconds -lt 5 -and -not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "Probe is bounded: $($result.Detail)"

    # Capability document: Python result in Static mode, ready only from the probe.
    $election = Join-Path $root 'Shared'
    $null = New-Item -ItemType Directory -Path $election
    $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $election -ProbeMode Static -PythonCommand (Join-Path $root 'Python312.cmd')
    $pythonResults = @($static.Results | Where-Object Name -eq 'Python')
    Assert-Python ($pythonResults.Count -eq 1 -and $pythonResults[0].Ready -and $static.ReadyCapabilities -contains 'Python' -and -not $pythonResults[0].ComponentMissing) 'Static document reports a ready Python.'
    $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $election -ProbeMode Static -PythonCommand 'smartm365-no-such-python-command'
    $pythonResults = @($static.Results | Where-Object Name -eq 'Python')
    Assert-Python ($pythonResults.Count -eq 1 -and -not $pythonResults[0].Ready -and $static.ReadyCapabilities -notcontains 'Python' -and $pythonResults[0].ComponentMissing) 'Static document reports a missing Python as ComponentMissing.'
    $pythonCalls = [regex]::Matches($distributedText, 'Get-PythonCapabilityResult -Command')
    Assert-Python ($pythonCalls.Count -eq 1 -and $distributedText.IndexOf('Get-PythonCapabilityResult -Command') -gt $distributedText.IndexOf("Invoke-IsolatedCapabilityProbe -Name 'Adaxes'")) 'One Python probe after both probe modes (Static and ReadOnly).'

    # PreparedPythonCommand is read where the job reads it (configuration location mapping), read-only.
    $configPath = Join-Path $root 'SmartM365-WorkplaceEvidence-Prepare.local.json'
    & $distributedModule {
        param($path)
        $script:PythonTestMappedFrom = ''
        $script:PythonTestConfigPath = $path
        function script:ConvertTo-SmartM365InventoryConfigLocation { param([string]$Path) $script:PythonTestMappedFrom = $Path; return $script:PythonTestConfigPath }
    } $configPath
    try {
        $command = & $distributedModule { Get-SmartM365OrchestratorPythonCommand }
        $mappedFrom = & $distributedModule { $script:PythonTestMappedFrom }
        Assert-Python ($command -eq 'python' -and $mappedFrom -like '*\SmartInventory\PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.local.json') "Absent configuration gives python from the job configuration path: $mappedFrom"
        Assert-Python (-not (Test-Path -LiteralPath $configPath)) 'Absent configuration is not created.'
        [IO.File]::WriteAllText($configPath, '{ "PreparedPythonCommand": "' + ((Join-Path $root 'Python312.cmd') -replace '\\', '\\') + '" }')
        Assert-Python ((& $distributedModule { Get-SmartM365OrchestratorPythonCommand }) -eq (Join-Path $root 'Python312.cmd')) 'PreparedPythonCommand is used.'
        $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $election -ProbeMode Static
        Assert-Python ($static.ReadyCapabilities -contains 'Python') 'Without -PythonCommand the configured command is probed.'
        [IO.File]::WriteAllText($configPath, '{ "PreparedPythonCommand": "" }')
        Assert-Python ((& $distributedModule { Get-SmartM365OrchestratorPythonCommand }) -eq 'python') 'Empty PreparedPythonCommand gives python.'
        & $distributedModule { function script:ConvertTo-SmartM365InventoryConfigLocation { param([string]$Path) throw 'SmartInventory private file still in the code folder.' } }
        $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $election -ProbeMode Static
        $pythonResults = @($static.Results | Where-Object Name -eq 'Python')
        Assert-Python ($pythonResults.Count -eq 1 -and -not $pythonResults[0].Ready -and $pythonResults[0].Detail -like '*could not be read for PreparedPythonCommand*' -and -not $pythonResults[0].ComponentMissing) 'Unreadable configuration is a warning, not a crash.'
    }
    finally {
        & $distributedModule { Remove-Item -LiteralPath 'function:script:ConvertTo-SmartM365InventoryConfigLocation' -ErrorAction SilentlyContinue }
    }

    # Template sync adds Python to the existing WorkplaceEvidence-Prepare job, once; other capabilities stay local.
    $manifestFolder = Join-Path $root 'Central'
    $null = New-Item -ItemType Directory -Path $manifestFolder
    $manifestPath = Join-Path $manifestFolder 'Orchestrator-Jobs.json'
    $existing = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
    $existingPrepare = @($existing.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')[0]
    $existingPrepare.RequiredCapabilities = @('SharedRuntime', 'Graph')
    $otherJob = @($existing.Jobs | Where-Object { $_.Name -ne 'WorkplaceEvidence-Prepare' -and $_.PSObject.Properties['RequiredCapabilities'] -and @($_.RequiredCapabilities) -contains 'SharedRuntime' -and @($_.RequiredCapabilities) -contains 'Graph' })[0]
    $otherJob.RequiredCapabilities = @('Graph')
    [IO.File]::WriteAllText($manifestPath, ($existing | ConvertTo-Json -Depth 100), [Text.UTF8Encoding]::new($false))
    $sync = Sync-SmartM365OrchestratorJobsManifest -Path $manifestPath -TemplatePath (Join-Path $folder 'Orchestrator-Jobs.json.template')
    $synced = Read-SmartM365OrchestratorJson -Path $manifestPath
    $syncedPrepare = @($synced.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')[0]
    Assert-Python ($sync.Updated -and @($sync.UpdatedJobNames) -contains 'WorkplaceEvidence-Prepare' -and (@($syncedPrepare.RequiredCapabilities) -join ',') -ceq 'SharedRuntime,Graph,Python') 'Sync adds Python to the existing job.'
    Assert-Python (@(@($synced.Jobs | Where-Object Name -eq $otherJob.Name)[0].RequiredCapabilities) -notcontains 'SharedRuntime') 'Sync never re-adds another capability.'
    Assert-Python (@($sync.UpdatedJobNames) -notcontains $otherJob.Name) 'The other job is not reported as updated.'
    $again = Sync-SmartM365OrchestratorJobsManifest -Path $manifestPath -TemplatePath (Join-Path $folder 'Orchestrator-Jobs.json.template')
    Assert-Python (@($again.UpdatedJobNames) -notcontains 'WorkplaceEvidence-Prepare') 'Second sync leaves the job unchanged.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }

# Election: the Prepare job goes only to a server with Python, and stays unassigned without one.
$sample = @(
    [pscustomobject]@{ Name = 'Prepare'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'Prepare'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime', 'Graph', 'Python'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('07:45') } }
)
$capabilities = @(
    [pscustomobject]@{ ServerName = 'SRV01'; ReadyCapabilities = @('SharedRuntime', 'Graph'); GraphAppRoles = @() }
    [pscustomobject]@{ ServerName = 'PY01'; ReadyCapabilities = @('SharedRuntime', 'Graph', 'Python'); GraphAppRoles = @() }
)
$plan = Get-SmartM365OrchestratorElectionPlan -Jobs $sample -ServerCapabilities $capabilities -ServerWeights @{ SRV01 = 10.0; PY01 = 1.0 } -ServerJobPolicies @{}
Assert-Python (@($plan.Assignments | Where-Object { $_.JobName -eq 'Prepare' -and $_.OwnerServer -eq 'PY01' }).Count -eq 1) 'Prepare elected on the Python server despite the weight of the other.'
$noPython = Get-SmartM365OrchestratorElectionPlan -Jobs $sample -ServerCapabilities @($capabilities[0]) -ServerWeights @{ SRV01 = 1.0 } -ServerJobPolicies @{}
Assert-Python (@($noPython.Assignments).Count -eq 0 -and @($noPython.UnassignedGroups).Count -eq 1) 'Without a Python server the job stays unassigned.'

"[{0}] PASS: Python capability, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
