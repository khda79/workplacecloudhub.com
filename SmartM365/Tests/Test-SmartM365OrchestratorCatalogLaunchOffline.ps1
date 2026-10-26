#Requires -Version 7.0
<#
.SYNOPSIS
Validate offline that the orchestrator starts catalog jobs through SmartM365.Inventory.Launch.psm1.
.VERSION
1.0.0
.NOTES
The orchestrator functions are extracted by AST and run against a fictional SmartM365 tree in
%TEMP% (no tenant, no network, no collector). Real supervisor and job processes are started:
- a job declaring Runtime and Execution runs through the catalog engine, in LocalCache by default
  (OnPremExecutionMode), its LauncherPath ignored; the job log shows the launch mode, the effective
  execution mode, the cache folder and the job output; the exit code reaches the result marker;
- Execution.Mode of the runtime manifest overrides the catalog (Direct, no copy);
- an output inside the cache stops the launch with exit code 1;
- a legacy LauncherPath job still runs, with a deprecation warning in the job and orchestrator logs;
- prerequisites check the launch module instead of the launcher; Authenticode includes the module;
- the manifest normalization accepts Execution.Mode and rejects an invalid value.
The LocalCache job uses a unique cache name under the standard cache folders; they are removed.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$orchestratorPath = Join-Path $smartRoot 'SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
$launchModuleSource = Join-Path $smartRoot 'SmartInventory\Orchestrator\SmartM365.Inventory.Launch.psm1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Orchestrator parse failed.' }
$functionNames = @('Start-InventoryJob', 'Get-JobEngine', 'Get-PwshPath', 'Get-OrchestratorWindowsPowerShellModulePathCommand', 'Get-OrchestratorJobLaunchDefinition',
    'Get-OrchestratorJobPrerequisitePathCheck', 'Get-OrchestratorTenantDataPathCheck', 'Get-OrchestratorAuthenticodeFileList', 'ConvertTo-NormalizedJob', 'Resolve-OrchestratorJobArguments')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing orchestrator function: $name" }
    $node.Extent.Text
}

$script:checks = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}

$base = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CatalogLaunch-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
$fixtureRoot = Join-Path $base 'SmartM365'
$inventory = Join-Path $fixtureRoot 'SmartInventory'
$markers = Join-Path $base 'Markers'
$workspace = Join-Path $base 'Workspace'
$cacheName = 'SmartM365OfflineTest-' + [guid]::NewGuid().ToString('N').Substring(0, 12)

function Write-FixtureFile {
    param([string]$Path, [string]$Content)
    New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
    [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false))
}
function Set-FixtureWorkspace {
    param([string]$WorkspaceRootPath)
    Write-FixtureFile (Join-Path $fixtureRoot 'Config\SmartM365.global.local.json') ([ordered]@{ DefaultTenant = 'demo'; SmartM365RootPath = '.'; WorkspaceRootPath = $WorkspaceRootPath; OnPremExecutionMode = 'LocalCache' } | ConvertTo-Json)
}
function Read-Marker {
    param([string]$Name)
    $path = Join-Path $markers ($Name + '.json')
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    return (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)
}

try {
    foreach ($file in @('Config\SmartM365-TenantContext.ps1', 'Config\SmartM365.global.local.json.template', 'Config\SmartM365-JsonTransport.policy.psd1', 'Config\Tenants\tenant.local.json.template', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psd1', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psm1')) {
        $destination = Join-Path $fixtureRoot $file
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $smartRoot $file) -Destination $destination
    }
    New-Item -ItemType Directory -Path $markers, $workspace, (Join-Path $inventory 'Orchestrator') -Force | Out-Null
    Copy-Item -LiteralPath $launchModuleSource -Destination (Join-Path $inventory 'Orchestrator')
    Set-FixtureWorkspace -WorkspaceRootPath $workspace
    Write-FixtureFile (Join-Path $fixtureRoot 'Config\Tenants\demo.local.json') ([ordered]@{ ProfileKey = 'demo'; OrganizationKey = 'contoso'; EnvironmentKey = 'lab'; TenantKey = 'contoso-lab' } | ConvertTo-Json)
    Write-FixtureFile (Join-Path $inventory 'Fixture\SmartM365-Fixture-Probe.ps1') @'
param([string]$Tenant = 'none', [int]$ExitCode = 0, [string]$Marker = 'probe', [switch]$Flag, [switch]$Connect)
'PROBE OUTPUT ' + $Marker
[Console]::Error.WriteLine('PROBE ERROR ' + $Marker)
$data = [ordered]@{ ScriptRoot = $PSScriptRoot; Tenant = $Tenant; Edition = $PSVersionTable.PSEdition; Flag = [bool]$Flag; Connect = [bool]$Connect }
[IO.File]::WriteAllText((Join-Path $env:SMARTM365_TEST_MARKERS ($Marker + '.json')), ($data | ConvertTo-Json))
exit $ExitCode
'@
    Write-FixtureFile (Join-Path $inventory 'Launchers\Start-Fixture-Probe.cmd') "@echo off`r`n`"%ProgramFiles%\PowerShell\7\pwsh.exe`" -NoProfile -ExecutionPolicy Bypass -File `"%~dp0..\Fixture\SmartM365-Fixture-Probe.ps1`" -Tenant `"%SMARTM365_ORCHESTRATOR_TENANT%`" %*`r`nexit /b %ERRORLEVEL%`r`n"
    $catalog = [ordered]@{
        Jobs = @(
            [ordered]@{ Name = 'Engine-Cache'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = '-Marker template'; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd'; Runtime = 'WindowsPowerShell'; LaunchArguments = '-Flag'; Execution = [ordered]@{ CacheName = $cacheName; Include = @() } },
            [ordered]@{ Name = 'Engine-Core'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = ''; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd'; Runtime = 'PowerShell7' },
            [ordered]@{ Name = 'Legacy-Launcher'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; Arguments = ''; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd' }
        )
    }
    Write-FixtureFile (Join-Path $inventory 'Orchestrator\Orchestrator-Jobs.json.template') ($catalog | ConvertTo-Json -Depth 10)
    $env:SMARTM365_TEST_MARKERS = $markers

    $module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
    & $module {
        param($inventoryRoot, $logFolder)
        Import-Module (Join-Path $inventoryRoot 'Orchestrator\SmartM365.Inventory.Launch.psm1') -Force
        $script:LaunchModulePath = Join-Path $inventoryRoot 'Orchestrator\SmartM365.Inventory.Launch.psm1'
        $script:SmartInventoryRoot = $inventoryRoot
        $script:Settings = [pscustomobject]@{ JobLogFolderPath = $logFolder; SharedDataFolderPath = $logFolder; TenantDataRoots = @(); AuthenticodeCheckCoreModule = $false; AuthenticodeCheckWindowsPowerShellModule = $false; CoreModulePath = '' }
        $script:RunningJobs = @{}
        $script:State = @{ Jobs = @{} }
        $script:ScriptName = 'SyntheticOrchestrator'
        $script:Tenant = 'demo'
        $script:Connect = $true
        $script:OrchestratorLog = [Collections.Generic.List[object]]::new()
        function script:Get-JobState { param($JobName) if (-not $script:State.Jobs.ContainsKey($JobName)) { $script:State.Jobs[$JobName] = @{ Running = $null } }; $script:State.Jobs[$JobName] }
        function script:Resolve-OrchestratorJobPath { param($Path) if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path $script:SmartInventoryRoot $Path } }
        function script:Invoke-OrchestratorAuthenticodeValidation { param($Job, $ScriptFullPath) [pscustomobject]@{ Allowed = $true } }
        function script:ConvertTo-StateTime { param($Value) ([datetime]$Value).ToString('o') }
        function script:Save-OrchestratorState {}
        function script:Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:OrchestratorLog.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
        function script:Complete-JobRun { param($JobName, $RunInfo, $StatusHint, $ExitCode, $EndTime, $ErrorText) throw "Launch failed before a child was started: $ErrorText" }
    } $inventory (Join-Path $base 'Logs')

    function Invoke-FixtureJob {
        param([string]$Name, [string]$Arguments, [string]$ExecutionMode = '', [string]$PowerShellEdition = 'PowerShell7')
        $job = [pscustomobject]@{ Name = $Name; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd'; Arguments = $Arguments; PowerShellEdition = $PowerShellEdition; ExecutionMode = $ExecutionMode; TimeoutMinutes = 5 }
        & $module { param($job) Start-InventoryJob -Job $job -Occurrence (Get-Date) } $job
        $process = & $module { param($n) $script:RunningJobs[$n].Process } $Name
        if (-not $process.WaitForExit(120000)) { throw "Fixture job $Name did not exit." }
        $record = & $module { param($n) $script:State.Jobs[$n].Running } $Name
        $result = Get-Content -LiteralPath $record.ResultPath -Raw | ConvertFrom-Json
        $text = [IO.File]::ReadAllText($record.LogPath)
        & $module { param($n) $script:RunningJobs.Remove($n) } $Name
        return [pscustomobject]@{ ExitCode = $result.ExitCode; Log = $text; ProcessName = (& $module { param($n) $script:State.Jobs[$n].Running.ProcessName } $Name) }
    }

    # Catalog engine, LocalCache from OnPremExecutionMode, LauncherPath ignored, Windows PowerShell 5.1 supervisor.
    $run = Invoke-FixtureJob -Name 'Engine-Cache' -Arguments '-Marker cache -ExitCode 4' -PowerShellEdition 'PowerShell7'
    $marker = Read-Marker 'cache'
    Assert-True ($run.ExitCode -eq 4) "Engine-Cache exit code 4 expected, got $($run.ExitCode)."
    Assert-True ($run.Log -match 'Launch mode: CatalogEngine') 'The job log must show the catalog engine launch mode.'
    Assert-True ($run.Log -match 'Execution mode: LocalCache \(requested LocalCache, source OnPremExecutionMode\)') 'The job log must show the effective execution mode and its source.'
    Assert-True ($run.Log -match 'Cache folder: first writable of ') 'The job log must list the cache folder candidates.'
    Assert-True ($run.Log -match '\[INFO\] Execution mode: LocalCache; cache folder: (?<Folder>[^\r\n]+)') 'The engine must log the cache folder it used.'
    $usedCache = $Matches['Folder'].Trim()
    Assert-True ($usedCache -like ('*' + $cacheName + '\SmartM365')) "Unexpected cache folder: $usedCache"
    Assert-True ($run.Log -match 'LauncherPath ignored \(catalog engine\)') 'An ignored LauncherPath must be reported.'
    Assert-True ($run.Log -match 'PowerShell engine: \S*powershell\.exe \(WindowsPowerShell\)' -and $run.ProcessName -eq 'powershell') 'The supervisor must run the catalog runtime (Windows PowerShell 5.1).'
    Assert-True ($run.Log -match 'PROBE OUTPUT cache' -and $run.Log -match 'PROBE ERROR cache') 'The script output and errors must reach the job log.'
    Assert-True ($null -ne $marker -and $marker.Edition -eq 'Desktop' -and $marker.Tenant -eq 'demo' -and $marker.Flag) 'The script must run under 5.1 with the tenant and LaunchArguments.'
    Assert-True ($marker.ScriptRoot -like ('*' + $cacheName + '\SmartM365\SmartInventory\Fixture')) "The script must run from the local cache: $($marker.ScriptRoot)"
    Assert-True ($run.Log -notmatch 'Start-Fixture-Probe\.cmd"') 'The legacy launcher must not be called.'

    # Execution.Mode from the runtime manifest overrides the catalog: Direct, in place.
    $run = Invoke-FixtureJob -Name 'Engine-Cache' -Arguments '-Marker direct' -ExecutionMode 'Direct'
    $marker = Read-Marker 'direct'
    Assert-True ($run.ExitCode -eq 0 -and $run.Log -match 'Execution mode: Direct \(requested Direct, source Job \(runtime manifest\)\)') 'The runtime manifest mode must be applied and logged.'
    Assert-True ($marker.ScriptRoot -ieq (Join-Path $inventory 'Fixture')) 'Direct must run the script in place.'

    # PowerShell 7 catalog job: -Connect follows the orchestrator, exit code transmitted.
    $run = Invoke-FixtureJob -Name 'Engine-Core' -Arguments '-Marker core -ExitCode 6'
    $marker = Read-Marker 'core'
    Assert-True ($run.ExitCode -eq 6 -and $run.ProcessName -eq 'pwsh' -and $marker.Edition -eq 'Core' -and $marker.Connect) 'A PowerShell 7 catalog job must run with -Connect and return its exit code.'
    Assert-True ($run.Log -match 'Execution mode: Direct \(requested Direct, source No Execution definition\)') 'A job without Execution runs Direct.'

    # Output inside the cache: the engine refuses, the job fails with exit code 1.
    Set-FixtureWorkspace -WorkspaceRootPath '{{SmartM365RootPath}}'
    $run = Invoke-FixtureJob -Name 'Engine-Cache' -Arguments '-Marker refused'
    Assert-True ($run.ExitCode -eq 1 -and $run.Log -match 'Catalog engine could not prepare Engine-Cache: .*inside the disposable launcher cache') 'An output inside the cache must stop the launch.'
    Assert-True ($null -eq (Read-Marker 'refused')) 'The script must not run when its output resolves inside the cache.'
    Set-FixtureWorkspace -WorkspaceRootPath $workspace

    # Legacy LauncherPath: still accepted, with a warning.
    $run = Invoke-FixtureJob -Name 'Legacy-Launcher' -Arguments '-Marker legacy -ExitCode 2'
    $warnings = @(& $module { $script:OrchestratorLog } | Where-Object { $_.Level -eq 'WARN' -and $_.Message -match 'Legacy-Launcher: LauncherPath is deprecated' })
    Assert-True ($run.ExitCode -eq 2 -and (Read-Marker 'legacy').Tenant -eq 'demo') 'A LauncherPath job must still run through its launcher.'
    Assert-True ($run.Log -match 'Launch mode: LauncherPath' -and $run.Log -match 'WARNING: LauncherPath is deprecated' -and $warnings.Count -eq 1) 'LauncherPath must be reported as deprecated in the job and orchestrator logs.'

    # Prerequisites and Authenticode.
    $engineChecks = & $module { Get-OrchestratorJobPrerequisitePathCheck -Job ([pscustomobject]@{ Name = 'Engine-Cache'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd'; PowerShellEdition = 'PowerShell7' }) }
    $legacyChecks = & $module { Get-OrchestratorJobPrerequisitePathCheck -Job ([pscustomobject]@{ Name = 'Legacy-Launcher'; ScriptPath = 'Fixture\SmartM365-Fixture-Probe.ps1'; LauncherPath = 'Launchers\Start-Fixture-Probe.cmd'; PowerShellEdition = 'PowerShell7' }) }
    Assert-True ((@($engineChecks.Name) -contains 'Launch module') -and -not (@($engineChecks.Name) -contains 'Job launcher')) 'A catalog job must check the launch module, not its launcher.'
    Assert-True (@($legacyChecks.Name) -contains 'Job launcher') 'A legacy job keeps its launcher check.'
    $files = & $module { Get-OrchestratorAuthenticodeFileList -Job ([pscustomobject]@{ Name = 'Engine-Cache'; PowerShellEdition = 'WindowsPowerShell'; LaunchEngine = $true }) -ScriptFullPath 'C:\probe.ps1' }
    Assert-True (@($files.Role) -contains 'SmartInventory launch module') 'Authenticode must validate the launch module of a catalog job.'

    # Manifest normalization of Execution.Mode.
    $normalized = & $module {
        $errors = [Collections.Generic.List[string]]::new()
        $raw = [pscustomobject]@{ Name = 'Normalized'; ScriptPath = 'x.ps1'; Execution = [pscustomobject]@{ Mode = 'Auto' }; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00') } }
        $job = ConvertTo-NormalizedJob -RawJob $raw -Errors $errors
        $bad = [pscustomobject]@{ Name = 'Bad'; ScriptPath = 'x.ps1'; Execution = [pscustomobject]@{ Mode = 'Remote' }; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00') } }
        $null = ConvertTo-NormalizedJob -RawJob $bad -Errors $errors
        [pscustomobject]@{ Mode = $job.ExecutionMode; Errors = @($errors) }
    }
    Assert-True ($normalized.Mode -eq 'Auto') 'Execution.Mode must be kept by the manifest normalization.'
    Assert-True (@($normalized.Errors | Where-Object { $_ -match "Bad': Execution.Mode must be" }).Count -eq 1 -and @($normalized.Errors).Count -eq 1) 'An invalid Execution.Mode must be rejected.'
}
finally {
    Remove-Item Env:\SMARTM365_TEST_MARKERS -ErrorAction SilentlyContinue
    foreach ($candidate in @((Join-Path $env:ProgramData ('SmartM365\LauncherCache\' + $cacheName)), (Join-Path $env:SystemRoot ('Temp\SmartM365LauncherCache\' + $cacheName)), ('C:\Temp\SmartM365LauncherCache\' + $cacheName), (Join-Path $env:TEMP ('SmartM365LauncherCache\' + $cacheName)))) {
        if (Test-Path -LiteralPath $candidate) { Remove-Item -LiteralPath $candidate -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: orchestrator catalog launch, $($script:checks) offline checks."
