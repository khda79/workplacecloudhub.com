#Requires -Version 7.0
<#
.SYNOPSIS
Validate offline that the orchestrator starts catalog jobs through SmartM365.Inventory.Launch.psm1.
.VERSION
1.1.0
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
- the manifest normalization accepts Execution.Mode and rejects an invalid value;
- a manifest ScriptPath that differs from the catalog script is refused before Authenticode, and
  the engine refuses a catalog script that differs from the validated one.
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
    Assert-True ($run.Log -match 'Cache folder: first writable of ' -and $run.Log -notmatch 'C:\\Temp\\SmartM365LauncherCache') 'The job log must list the orchestrator cache candidates (no C:\Temp).'
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

    # The validated (manifest) script must be the catalog script that runs.
    Write-FixtureFile (Join-Path $inventory 'Fixture\SmartM365-Fixture-Other.ps1') 'exit 0'
    $mismatchJob = [pscustomobject]@{ Name = 'Engine-Core'; ScriptPath = 'Fixture\SmartM365-Fixture-Other.ps1'; LauncherPath = ''; Arguments = ''; PowerShellEdition = 'PowerShell7'; ExecutionMode = ''; TimeoutMinutes = 5 }
    $mismatchError = ''
    try { & $module { param($job) Start-InventoryJob -Job $job -Occurrence (Get-Date) } $mismatchJob } catch { $mismatchError = $_.Exception.Message }
    Assert-True ($mismatchError -match 'differs from the manifest ScriptPath' -and @(& $module { $script:OrchestratorLog } | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -match 'Engine-Core: Catalog script' }).Count -eq 1) 'A manifest ScriptPath different from the catalog must refuse the launch.'
    $expectedLog = Join-Path $base 'expected.log'
    $expectedExit = & $module { param($root, $log) Invoke-SmartM365InventoryOrchestratedLaunch -SmartInventoryRoot $root -Name 'Engine-Core' -Tenant 'demo' -LogPath $log -ExpectedScriptPath (Join-Path $root 'Fixture\SmartM365-Fixture-Other.ps1') } $inventory $expectedLog
    Assert-True ($expectedExit -eq 1 -and (Get-Content -LiteralPath $expectedLog -Raw) -match 'differs from the validated script') 'The engine must refuse a script other than the validated one.'

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCg4KItwnKTNd2E
# RwipmYAmeBQlGvscs5797mma5ZbDr6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPgt61j13+RR2H0bMn5z2fX81F0mwKhTP9RbpTm02y29MA0GCSqG
# SIb3DQEBAQUABIIBgKCUBJNTjqxLCl0zW7nAYC9XtWuOWILMmiytpeFYsxTO/vK9
# 74eDtmwDz1jDK/zwqriEc2u/hFu1mHf3MWuGqU+2foBcaVPPldrBtIY5V2mlDc3b
# ao62NIaQD1hosiz7i3Dr3Ed+kTNM8UmwvMHWs/2a4kaJDEk2+Q/pJUx6hHb2NKry
# 0WiiqvvuelqHyr0jpqKrV3VkzqnRBN8jxG16iGqD94cN6GD2FjW9CVjzNGHt0f1B
# nlufJu/OQXWZQtlmVEJbuqss4Wx/bb1x6aikz4Q6kXHaS1XYFqNinYrao2UpIg8B
# PH8OMuIlHDNFawEUjEIeJ9skwSsrqzIPA1x9DtCVmOMbSg7a5zxmZL88wOcAV2jy
# paPD5+UAc1jOE7AYKPljhNaPpW0r77BvgnvFRUn66aBiUZZnUm1Udmq9/PU6LJvC
# uafLbkjBaDQ8Glf3hHih2FMuFzxhnIW0+QC0fu/+HL22vsQ2HPQH/GRDvFG9zNC9
# nbn4gQJhRyPUb4E196GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYwOTU2
# MTNaMC8GCSqGSIb3DQEJBDEiBCADfbs5Ia44ZiQF7bOMvu6rORSMoIjhm+NEmWxe
# bSTOXjANBgkqhkiG9w0BAQEFAASCAgB/kJK0sZVjXWwHYF8nia/eoRM39k4mY+O3
# pFB4bXYvQB86gem/osUO/wGQOG4nQT6FGqLk/2cKpCE6SA8e0TdSgRoJHEJHWBC7
# KqVPT8BPCshdZXYXbySAn06q1VS+SUy6AwUAh/LOHcZgO2qnvv9bmfLRsXSD42gL
# vXh/ZkqtjFKLH6UnjlEIzHOZjzyoS+LiEEiyeVpkKBFrx+tLP+QtlswQWPsavVJw
# RqGsIqezkkcJ4QzgZz/BiSF3g6idw4honh0VUuTvEVtXmBMD69S82J0gQcjmFq6n
# 6c4mJueB/lWeqyszcacVVsTi0X8dLhTKWhV2TFVwTnuGBhIjSnTe9qmTxDBiWtho
# HwAnbqAU6uTz9hF4xXpCkZ2LON+nvc6Qo4zrCXfUdQgQ884x+SVA//uxvCwCcyWi
# AxBa0q93AtYqLdCiBv7kIKVawr/NcAZ4aPhMG1J9TpOh57sNcQdRrsd0wOOIQz/x
# d4A0DVy2RidQ74mnXr5QgqxseEEcI1G7rGyYiHw6neMJ67Wy5msmhima/z/zfWD9
# 8TzKwFP857oNhp6lhq9NElDEydsf+hJYdFT9ME3UxepgaIO2iIeEMgsunM2CaPf6
# TGI9YEVnj2Q0j2+LcSvKfE5CsGBV6/uDDnq0yaBAmx7GGSJN1JsfUxVeLW+yLtU7
# rKBihUH2gA==
# SIG # End signature block
