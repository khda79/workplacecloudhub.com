#Requires -Version 7.0
<#
.SYNOPSIS
Checks the daily server acceptance run by every orchestrator resident, offline.
.DESCRIPTION
Extracts the resident functions and runs them against a fake acceptance script in a temporary
folder: due time with the per-server offset, catch-up after startup when the published result is
older than 24 hours, one run at a time, exit code kept across a recycle (<log>.exitcode), time
limit, disabled mode, cluster keys and validation. Real hidden pwsh child processes are started;
no tenant, collector, mail or publication.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$script:Checks = 0
function Assert-Daily {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$orchestratorText = [IO.File]::ReadAllText($orchestratorPath)
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$errors)
foreach ($name in @('Get-PwshPath', 'Stop-ProcessTree', 'Get-OrchestratorAcceptanceDailyOffsetMinutes', 'Get-OrchestratorAcceptanceLastResultUtc', 'Test-OrchestratorAcceptanceDailyDue', 'Get-OrchestratorAcceptanceDailyRecord', 'Update-OrchestratorAcceptanceDaily')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Assert-Daily ($null -ne $definition) "$name exists."
    . ([scriptblock]::Create($definition.Extent.Text))
}
$script:Log = [Collections.Generic.List[object]]::new()
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') $script:Log.Add([pscustomobject]@{ Message = $Message; Level = $Level }) }
$script:Saves = 0
function Save-OrchestratorState { $script:Saves++ }

# Due time: daily after the local time plus the offset, once; catch-up only at the first check after startup.
$day = [datetime]::new(2026, 10, 30)
$utc = { param([datetime]$Local) $Local.ToUniversalTime() }
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(16).AddMinutes(59) -DailyTime '17:00') -eq '') 'Not due before the daily time.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17) -DailyTime '17:00') -eq 'Daily') 'Due at the daily time.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17).AddMinutes(5) -DailyTime '17:00' -OffsetMinutes 10) -eq '') 'Offset delays the run.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17).AddMinutes(10) -DailyTime '17:00' -OffsetMinutes 10) -eq 'Daily') 'Due at the daily time plus the offset.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(20) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddHours(17).AddMinutes(31))) -eq '') 'Only one run per day.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(20) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddDays(-1).AddHours(17).AddMinutes(31))) -eq 'Daily') 'A day missed while stopped runs later the same day.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-28)) -StartupCheck) -eq 'CatchUp') 'Catch-up at startup when the result is older than 24 h.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -StartupCheck) -eq 'CatchUp') 'Catch-up at startup without any published result.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-2)) -StartupCheck) -eq '') 'No catch-up with a recent result.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddHours(1)) -StartupCheck) -eq '') 'No catch-up loop after a recent start (recycle).'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-28))) -eq '') 'No catch-up after the first check.'

# Offset: 5 minutes per rank in the sorted list of expected servers; 0 when the server is not listed.
$local = ([string]$env:COMPUTERNAME).ToUpperInvariant()
$script:Settings = [pscustomobject]@{ ExpectedOrchestratorServers = @('ZZZZ-LAST', $local.ToLowerInvariant(), '0000-FIRST') }
$expectedIndex = [array]::IndexOf([string[]]@(@('ZZZZ-LAST', $local, '0000-FIRST') | Sort-Object -Unique), $local)
Assert-Daily ((Get-OrchestratorAcceptanceDailyOffsetMinutes) -eq $expectedIndex * 5) 'Offset follows the rank of this server.'
$script:Settings = [pscustomobject]@{ ExpectedOrchestratorServers = @('OTHER-A') }
Assert-Daily ((Get-OrchestratorAcceptanceDailyOffsetMinutes) -eq 0) 'Unlisted server has no offset.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AcceptanceDaily-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'Shared') -Force
    $fakeScript = Join-Path $root 'Fake-Acceptance.ps1'
    function Get-OrchestratorAcceptanceScriptPath { return $fakeScript }
    $Tenant = 'offline'
    $script:Settings = [pscustomobject]@{
        AcceptanceDailyEnabled = $true; AcceptanceDailyTime = '00:00'; AcceptanceDailyTimeoutMinutes = 60
        ExpectedOrchestratorServers = @(); JobLogFolderPath = (Join-Path $root 'Logs'); SharedDataFolderPath = (Join-Path $root 'Shared')
    }

    # Published result: GeneratedAtUtc is read from the service account file of this server.
    Assert-Daily ((Get-OrchestratorAcceptanceLastResultUtc) -eq [datetime]::MinValue) 'No published result.'
    $serverFolder = Join-Path $root "Shared\$local"
    $null = New-Item -ItemType Directory -Path $serverFolder -Force
    [IO.File]::WriteAllText((Join-Path $serverFolder 'Orchestrator-Acceptance.json'), '{ "GeneratedAtUtc": "2026-10-29T03:00:00.0000000Z" }', [Text.UTF8Encoding]::new($false))
    Assert-Daily ((Get-OrchestratorAcceptanceLastResultUtc) -eq [datetime]::new(2026, 10, 29, 3, 0, 0, [DateTimeKind]::Utc)) 'Published GeneratedAtUtc is read.'

    # Disabled: nothing starts.
    $script:State = @{}
    $script:AcceptanceDailyStartupChecked = $false
    $script:Settings.AcceptanceDailyEnabled = $false
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($null -eq (Get-OrchestratorAcceptanceDailyRecord) -and $script:Log.Count -eq 0) 'Disabled: nothing starts.'
    $script:Settings.AcceptanceDailyEnabled = $true

    # Normal run: started hidden with -Tenant, one run at a time, exit code 1 reported as WARN.
    [IO.File]::WriteAllLines($fakeScript, [string[]]@('param([string]$Tenant)', '"fake acceptance tenant=$Tenant"', 'Start-Sleep -Seconds 2', 'exit 1'))
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $record = Get-OrchestratorAcceptanceDailyRecord
    Assert-Daily ($null -ne $record -and [int]$record['Pid'] -gt 0 -and $script:Log[-1].Message -like 'Daily server acceptance started (daily): PID *') "Run started: $($script:Log[-1].Message)"
    $firstPid = [int]$record['Pid']
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ([int](Get-OrchestratorAcceptanceDailyRecord)['Pid'] -eq $firstPid -and @($script:Log | Where-Object Message -like '*started*').Count -eq 1) 'No second run while the first one runs.'
    Wait-Process -Id $firstPid -Timeout 60 -ErrorAction SilentlyContinue
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $finished = $script:Log[-1]
    Assert-Daily ($finished.Message -like 'Daily server acceptance finished: exit code 1 (at least one FAIL)*' -and $finished.Level -eq 'WARN') "Exit code 1 reported as WARN: $($finished.Message)"
    $logPath = [string]$record['LogPath']
    Assert-Daily ((Get-Content -LiteralPath $logPath -Raw) -match 'fake acceptance tenant=offline') 'The acceptance output and -Tenant are in the log.'
    Assert-Daily ((Split-Path $logPath -Leaf) -like "Acceptance-$($env:COMPUTERNAME)_*.log" -and (Test-Path -LiteralPath ($logPath + '.exitcode'))) 'Log and exit code file named for retention.'
    $record = Get-OrchestratorAcceptanceDailyRecord
    Assert-Daily (-not $record.Contains('Pid') -and $record['LastStartUtc']) 'Record cleared, last start kept.'
    $logCount = $script:Log.Count
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($script:Log.Count -eq $logCount) 'No second run the same day.'

    # Recycle: a record left by the previous instance is finished from the exit code file.
    $exitFile = Join-Path $root 'Logs\Acceptance-previous.log'
    [IO.File]::WriteAllText($exitFile + '.exitcode', '0')
    $script:State['AcceptanceDaily'] = @{ Pid = 999999; ProcessStartUtc = (Get-Date).ToUniversalTime().ToString('o'); LastStartUtc = (Get-Date).AddMinutes(-3).ToUniversalTime().ToString('o'); LogPath = $exitFile }
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($script:Log[-1].Message -like 'Daily server acceptance finished: exit code 0 (PASS or WARN)*' -and $script:Log[-1].Level -eq 'INFO') "Previous instance run reported: $($script:Log[-1].Message)"

    # Time limit: the process tree is stopped and reported.
    [IO.File]::WriteAllLines($fakeScript, [string[]]@('param([string]$Tenant)', 'Start-Sleep -Seconds 60'))
    $script:State = @{}
    $script:Settings.AcceptanceDailyTimeoutMinutes = 0
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $slowPid = [int](Get-OrchestratorAcceptanceDailyRecord)['Pid']
    Start-Sleep -Milliseconds 500
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily (@($script:Log | Where-Object { $_.Message -like 'Daily server acceptance stopped after*' -and $_.Level -eq 'WARN' }).Count -eq 1) 'Time limit reported.'
    Start-Sleep -Seconds 1
    Assert-Daily ($null -eq (Get-Process -Id $slowPid -ErrorAction SilentlyContinue)) 'Process tree stopped at the limit.'
    $script:Settings.AcceptanceDailyTimeoutMinutes = 60

    # Missing script: warned once, not retried the same day.
    function Get-OrchestratorAcceptanceScriptPath { return (Join-Path $root 'Missing.ps1') }
    $script:State = @{}
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily (@($script:Log | Where-Object Message -like 'Daily server acceptance not started: script not found*').Count -eq 1) 'Missing script warned once.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# Wiring: main loop outside the launch phase (Active or Inactive, no MaxConcurrency), retention, startup log.
$mainLoop = [regex]::Match($orchestratorText, '(?s)Test-OrchestratorRuntimeUpdate -Now \$now.{0,400}?Update-OrchestratorAcceptanceDaily -Now \$now.{0,500}?Update-OrchestratorDiskSpaceGuard -Now \$now.{0,400}?Invoke-LaunchPhase -Now \$now')
Assert-Daily $mainLoop.Success 'Called every tick, before and outside the launch phase.'
$update = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-OrchestratorAcceptanceDaily' }, $true).Extent.Text
Assert-Daily ($update -notmatch '\$script:Settings\.MaxConcurrency|Get-OrchestratorLocalServerState|Test-OrchestratorLocalServerActive|Send-Orchestrator|Send-SmartM365') 'No concurrency, server state or mail involved.'
Assert-Daily ($orchestratorText.Contains("Filter = 'Acceptance-*.log*'; Days = `$script:Settings.JobLogRetentionDays")) 'Acceptance logs follow the job log retention.'
Assert-Daily ($orchestratorText.Contains('Write-OrchestratorLog -Message ("Daily server acceptance: enabled={0}')) 'Settings logged at startup.'

# Cluster keys: template, central mapping and validation.
$cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $base 'Orchestrator-Cluster.json.template')
Assert-Daily ($cluster.AcceptanceDailyEnabled -eq $true -and $cluster.AcceptanceDailyTime -eq '17:00' -and (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'Template keys are valid.'
foreach ($bad in @('4:30', '24:00', 'noon')) {
    $cluster.AcceptanceDailyTime = $bad
    Assert-Daily (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) "AcceptanceDailyTime '$bad' rejected."
}
$cluster.AcceptanceDailyTime = '23:59'; $cluster.AcceptanceDailyEnabled = 'yes'
Assert-Daily (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'AcceptanceDailyEnabled must be a boolean.'
Assert-Daily ($orchestratorText.Contains("AcceptanceDailyEnabled = 'Boolean'") -and $orchestratorText.Contains("AcceptanceDailyTime = 'String'")) 'Central cluster keys are applied by the resident.'

"[{0}] PASS: daily server acceptance, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
