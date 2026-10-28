<#
.SYNOPSIS
Offline tests of the server system information: each orchestrator reads OS, CPU, memory and drive
C: once at startup and publishes them in its heartbeat; the GUI Servers and versions view reads
them from the heartbeat and never queries the servers.
.VERSION
1.0
.NOTES
Synthetic heartbeats in a temporary folder. The collector runs read-only CIM queries on this
computer. No tenant initialization, collectors, scheduling or publication.
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

# Startup collection and heartbeat publication.
$source = [IO.File]::ReadAllText($orchestratorPath)
Check ($source -match '\$script:SystemInfo = Get-OrchestratorSystemInfo') 'The orchestrator no longer reads its system information at startup.'
Check ((Get-FunctionText 'Write-OrchestratorHeartbeat') -match 'SystemInfo = \$script:SystemInfo') 'The heartbeat no longer publishes the system information.'
Check (([regex]::Matches($source, 'Get-OrchestratorSystemInfo')).Count -eq 2) 'The system information must be read once, at startup only.'

# Real collector on this computer: values or a named error, never an exception.
. ([scriptblock]::Create((Get-FunctionText 'Get-OrchestratorSystemInfo')))
$info = Get-OrchestratorSystemInfo
Check ($null -ne $info -and $info.SystemDrive -eq 'C:' -and [datetimeoffset]::TryParse($info.CollectedAtUtc, [ref]([datetimeoffset]::MinValue))) 'The collector must return the drive and the collection time.'
Check (($info.Os -and $info.Cpu -and $info.MemoryGB -gt 0 -and $info.LogicalProcessors -gt 0) -or $info.Error) 'OS, CPU and memory must be read or the failure named.'
Check (($info.SystemDriveSizeGB -gt 0 -and $info.SystemDriveFreePercent -ge 0 -and $info.SystemDriveFreePercent -le 100) -or $info.Error -match 'Drive C:') 'Drive C: size and free percent must be read or the failure named.'

# GUI data: read from the heartbeat by the management module.
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SystemInfo-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    foreach ($server in @('SRV-A', 'SRV-B')) { New-Item -ItemType Directory -Path (Join-Path $root $server) -Force | Out-Null }
    $heartbeat = [ordered]@{
        Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.68'
        SystemInfo = [ordered]@{ CollectedAtUtc = '2026-10-28T05:00:00.0000000Z'; Os = 'Windows Server 2022 Standard (10.0.20348, build 20348)'; Cpu = 'Xeon Gold'; LogicalProcessors = 8; MemoryGB = 16; SystemDrive = 'C:'; SystemDriveSizeGB = 127.4; SystemDriveFreePercent = 8.5; Error = '' }
    }
    [IO.File]::WriteAllText((Join-Path $root 'SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    # SRV-B runs an older orchestrator: no SystemInfo in its heartbeat.
    [IO.File]::WriteAllText((Join-Path $root 'SRV-B/Orchestrator-Heartbeat.json'), ([ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.60' } | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $rows = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)
    $a = @($rows | Where-Object Server -eq 'SRV-A')[0]
    $b = @($rows | Where-Object Server -eq 'SRV-B')[0]
    $collectedLocal = ([datetimeoffset]::Parse('2026-10-28T05:00:00Z')).LocalDateTime.ToString('yyyy-MM-dd HH:mm')
    Check ($a.Version -eq '1.5.68' -and $a.Os -like 'Windows Server 2022*' -and $a.Cpu -eq 'Xeon Gold (8 logical)' -and $a.MemoryGB -eq 16) 'Version, OS, CPU and RAM must come from the heartbeat.'
    Check ($a.SystemDriveSizeGB -eq 127.4 -and $a.SystemDriveFreePercent -eq 8.5 -and $a.SystemDriveLow -and $a.SystemInfoAt -eq $collectedLocal) "Drive C: and the collection time must come from the heartbeat ($($a.SystemInfoAt))."
    Check ($b.Version -eq '1.5.60' -and -not $b.Os -and $null -eq $b.MemoryGB -and -not $b.SystemDriveLow -and -not $b.SystemInfoAt) 'An older orchestrator shows its version and empty system values.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# GUI: Servers and versions is a sub-tab of Configuration with the system columns; Activity is removed.
$guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
Check ($guiText -match '<TabItem x:Name="ConfigurationServersTab" Header="Servers and versions">' -and $guiText -notmatch '<TabItem Header="Servers and versions">' -and $guiText -notmatch '<TabItem Header="Activity">') 'Servers and versions must be under Configuration and Activity removed.'
foreach ($binding in @('Version', 'Os', 'Cpu', 'MemoryGB', 'SystemDriveSizeGB', 'SystemDriveFreePercent', 'SystemInfoAt')) {
    Check ($guiText -match ('Binding="\{Binding ' + $binding + '\}"')) "Servers grid column missing: $binding"
}
"PASS: server system information, $($script:passed) offline checks."
