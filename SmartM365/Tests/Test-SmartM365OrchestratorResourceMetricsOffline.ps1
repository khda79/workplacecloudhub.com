<#
.SYNOPSIS
Offline tests of the CPU and memory curves: each orchestrator samples every 5 minutes, keeps only
the last 12 hours in Orchestrator-Metrics.json and publishes the latest sample in its heartbeat;
the GUI reads the file and draws one chart per server in the Dashboard and in Servers and versions.
.VERSION
1.0
.NOTES
Synthetic samples in a temporary folder; one real read-only CIM sample of this computer. No tenant
initialization, collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText($Tree, [string]$Name) {
    $definition = $Tree.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing function: $Name" }
    $definition.Extent.Text
}

# Wiring: sampled in the tick loop before the heartbeat, reloaded at startup, latest sample published.
$source = [IO.File]::ReadAllText($orchestratorPath)
Check ($source -match 'Update-OrchestratorResourceMetrics -Now \$now\s+Write-OrchestratorHeartbeat') 'The tick loop no longer samples CPU and memory before the heartbeat.'
Check ($source -match 'Import-OrchestratorResourceMetrics -Now \(Get-Date\)') 'Startup no longer reloads the previous samples.'
Check ((Get-FunctionText $ast 'Write-OrchestratorHeartbeat') -match 'ResourceUsage = ') 'The heartbeat no longer publishes the latest sample.'
Check ((Get-FunctionText $ast 'Write-FileAtomically') -match 'ElectionPlan\|Metrics') 'The metrics file is not written as a regenerated snapshot.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Metrics-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path (Join-Path $root 'SRV-A') -Force | Out-Null
try {
    $functions = @('Get-OrchestratorMetricsPath', 'Get-OrchestratorResourceSample', 'Remove-OrchestratorExpiredResourceSamples', 'Import-OrchestratorResourceMetrics', 'Update-OrchestratorResourceMetrics')
    $stubs = @'
function Write-FileAtomically { param([string]$Path, [string]$Content) [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false)); $script:Writes++ }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now) $script:Warnings.Add($Message) }
'@
    $text = (@($functions | ForEach-Object { Get-FunctionText $ast $_ }) -join "`n") + "`n" + $stubs
    $module = New-Module -ScriptBlock ([scriptblock]::Create($text))
    Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Global
    function Reset-Orchestrator {
        & $module {
            param($Root)
            $script:Settings = [pscustomobject]@{ HeartbeatPath = (Join-Path $Root 'SRV-A/Orchestrator-Heartbeat.json') }
            $script:ResourceSamples = [Collections.Generic.List[object]]::new()
            $script:LastResourceSampleAt = [datetime]::MinValue
            $script:ResourceSampleMinutes = 5; $script:ResourceRetentionHours = 12
            $script:Writes = 0; $script:Warnings = [Collections.Generic.List[string]]::new()
        } $root
    }
    Reset-Orchestrator
    $now = Get-Date

    # A real sample of this computer: percentages between 0 and 100.
    $sample = & $module { param($n) Get-OrchestratorResourceSample -Now $n } $now
    Check ($null -ne $sample.Cpu -and $sample.Cpu -ge 0 -and $sample.Cpu -le 100 -and $sample.Memory -gt 0 -and $sample.Memory -le 100) "CPU and memory must be read as percentages ($($sample.Cpu), $($sample.Memory))."

    # Samples older than 12 hours are dropped when the file is rewritten.
    & $module {
        param($n)
        $script:ResourceSamples.Add([pscustomobject]@{ TimeUtc = $n.ToUniversalTime().AddHours(-13).ToString('o'); Cpu = 10; Memory = 20 })
        $script:ResourceSamples.Add([pscustomobject]@{ TimeUtc = $n.ToUniversalTime().AddHours(-1).ToString('o'); Cpu = 30; Memory = 40 })
        Update-OrchestratorResourceMetrics -Now $n
    } $now
    $metricsPath = Join-Path $root 'SRV-A/Orchestrator-Metrics.json'
    $document = Get-Content -LiteralPath $metricsPath -Raw | ConvertFrom-Json
    Check (@($document.Samples).Count -eq 2 -and $document.RetentionHours -eq 12 -and $document.SampleMinutes -eq 5 -and (& $module { $script:Writes }) -eq 1) 'The file must keep only the last 12 hours plus the new sample.'
    & $module { param($n) Update-OrchestratorResourceMetrics -Now $n.AddMinutes(4) } $now
    Check ((& $module { $script:Writes }) -eq 1 -and (& $module { $script:ResourceSamples.Count }) -eq 2) 'A sample must not be taken before 5 minutes.'
    & $module { param($n) Update-OrchestratorResourceMetrics -Now $n.AddMinutes(5) } $now
    Check ((& $module { $script:Writes }) -eq 2 -and (& $module { $script:ResourceSamples.Count }) -eq 3) 'A sample must be taken every 5 minutes.'
    Check ((& $module { $script:Warnings.Count }) -eq 0) "Unexpected warning: $(& $module { $script:Warnings -join '; ' })"

    # After a restart the curve continues from the file; stale samples are dropped on import.
    Reset-Orchestrator
    & $module { param($n) Import-OrchestratorResourceMetrics -Now $n.AddHours(11.5) } $now
    Check ((& $module { $script:ResourceSamples.Count }) -eq 2) 'Import must reload the samples still within 12 hours.'

    # Heartbeat values and file reading on the GUI side.
    $heartbeat = [ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.69'; ResourceUsage = [ordered]@{ TimeUtc = $now.ToUniversalTime().ToString('o'); Cpu = 42.5; Memory = 61.2 } }
    [IO.File]::WriteAllText((Join-Path $root 'SRV-A/Orchestrator-Heartbeat.json'), ($heartbeat | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
    Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $status = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)[0]
    Check ($status.CpuPercent -eq 42.5 -and $status.MemoryPercent -eq 61.2) 'Servers and versions must show the latest sample of the heartbeat.'
    $metrics = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-A') -Server 'SRV-A' -Now $now.AddMinutes(6)
    Check ($metrics.Server -eq 'SRV-A' -and @($metrics.Samples).Count -eq 3 -and $metrics.SampleMinutes -eq 5 -and @($metrics.Samples)[0].Time -is [datetime]) 'The GUI must read the samples of the last 12 hours.'
    $later = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-A') -Server 'SRV-A' -Now $now.AddHours(11.5)
    Check (@($later.Samples).Count -eq 2) 'The GUI must not show samples older than 12 hours.'
    $missing = Get-SmartM365OrchestratorServerMetrics -ServerFolder (Join-Path $root 'SRV-X') -Server 'SRV-X' -Now $now
    Check (@($missing.Samples).Count -eq 0) 'A server without samples must give an empty curve.'

    # Drawing: one panel per server, a line per series, broken by a gap.
    $guiPath = Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'
    $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$null)
    foreach ($name in @('ConvertTo-OrchestratorBrush', 'Update-OrchestratorServerMetricsChart')) { . ([scriptblock]::Create((Get-FunctionText $guiAst $name))) }
    $clock = [datetime]'2026-10-18 09:30'
    $samples = @(for ($minute = 720; $minute -ge 0; $minute -= 5) { if ($minute -le 300 -and $minute -gt 240) { continue }; [pscustomobject]@{ Time = $clock.AddMinutes(-$minute); Cpu = 30; Memory = 60 } })
    $canvas = [System.Windows.Controls.Canvas]::new(); $canvas.Width = 900; $canvas.Height = 130
    $canvas.Measure([System.Windows.Size]::new(900, 130)); $canvas.Arrange([System.Windows.Rect]::new(0, 0, 900, 130))
    $canvas.Tag = @([pscustomobject]@{ Server = 'SRV-A'; From = $clock.AddHours(-12); To = $clock; SampleMinutes = 5; Samples = $samples }, [pscustomobject]@{ Server = 'SRV-B'; From = $clock.AddHours(-12); To = $clock; SampleMinutes = 5; Samples = @() })
    Update-OrchestratorServerMetricsChart -Canvas $canvas
    $lines = @($canvas.Children | Where-Object { $_ -is [System.Windows.Shapes.Polyline] })
    Check ($lines.Count -eq 4 -and @($lines | Where-Object { $_.Points.Count -lt 2 }).Count -eq 0) "A one-hour gap must split CPU and RAM into two lines each ($($lines.Count))."
    $titles = @($canvas.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.FontWeight -eq [System.Windows.FontWeights]::SemiBold } | ForEach-Object Text)
    Check ($titles.Count -eq 2 -and $titles[0] -like 'SRV-A  CPU 30% *' -and $titles[1] -eq 'SRV-B  no sample') "Each server needs its own titled panel ($($titles -join ' | '))."
    $canvas.Tag = @()
    Update-OrchestratorServerMetricsChart -Canvas $canvas
    Check (@($canvas.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text -like 'No CPU or memory sample yet*' }).Count -eq 1) 'An empty chart must explain why.'

    # Both views host the chart.
    $guiText = [IO.File]::ReadAllText($guiPath)
    Check ($guiText -match 'x:Name="DashboardMetricsCanvas"' -and $guiText -match 'x:Name="ServersMetricsCanvas"' -and $guiText -match 'DashboardMetricsCanvas\.Add_SizeChanged' -and $guiText -match 'ServersMetricsCanvas\.Add_SizeChanged') 'The Dashboard and Servers and versions must both show the chart.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: CPU and memory curves, $($script:passed) offline checks."
