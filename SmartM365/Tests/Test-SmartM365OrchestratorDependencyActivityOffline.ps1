<#
.SYNOPSIS
Offline tests for the FreshSuccess dependency activity gate: an active dependency expected to end
before the downstream occurrence blocks it; a longer staged run does not.
.VERSION
1.0.0
.NOTES
Synthetic claims and run states under TEMP only. No collector, scheduler, tenant or network access.
#>
#requires -Version 7.0
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic tenant key read by the receipt fallback; the test process ends afterwards.')]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
$transport = Get-Module SmartM365.JsonTransport
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
$defs = foreach ($name in @('Get-JobOccurrencesInWindow', 'Get-LatestPastOccurrence', 'Get-OrchestratorDependencyActivity', 'Get-OrchestratorDependencyMaxAge', 'Get-OrchestratorDependencyFreshStatus', 'Get-OrchestratorClaimOccurrenceUtc', 'ConvertTo-OrchestratorUtcTime')) {
    $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true).Extent.Text
}
$gate = New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root = Join-Path $env:TEMP ('dependency-activity-' + [guid]::NewGuid().ToString('N'))
$shared = Join-Path $root 'DATA-ALL/Orchestrator'
$claims = Join-Path $shared 'Election/Claims'
$latest = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory $claims, $latest | Out-Null
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$managedApps = [pscustomobject]@{ Name = 'Intune-ManagedApps-Inventory'; ScriptPath = 'M365Inventory\WorkplaceScope\SmartM365-IntuneManagedApps-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 240; EstimatedDurationMinutes = 60; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('05:00'); DaysOfWeek = @() } }
$apps = [pscustomobject]@{ Name = 'Intune-DiscoveredApps-Inventory'; ScriptPath = 'M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 8640; EstimatedDurationMinutes = 1500; Schedule = [pscustomobject]@{ Type = 'Weekly'; Times = @('20:00'); DaysOfWeek = @('Saturday') } }
$fast = [pscustomobject]@{ Name = 'EXO-Mailboxes-Inventory-Fast'; ScriptPath = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 1380; EstimatedDurationMinutes = 90; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('00:15'); DaysOfWeek = @() } }
$late = [pscustomobject]@{ Name = 'Late-Producer'; ScriptPath = 'X\SmartM365-Late-Producer.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 60; EstimatedDurationMinutes = 5; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('07:40'); DaysOfWeek = @() } }
$prepare = [pscustomobject]@{ Name = 'WorkplaceEvidence-Prepare'; ScriptPath = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'; Enabled = $true; AssignmentMode = 'Elected'; DependencyMode = 'FreshSuccess'; DependencyMaxAgeHours = 48; DependsOn = @('Intune-ManagedApps-Inventory'); Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('07:45'); DaysOfWeek = @() } }
$jobs = @($managedApps, $apps, $fast, $late, $prepare)

# Claims are written directly so that their safe window follows the simulated clock, not the real one.
function Write-SyntheticClaim($Job, [datetime]$Occurrence, [string]$Status, $SafeUntil = $null) {
    $folder = Join-Path $claims $Job.Name
    New-Item -ItemType Directory $folder -Force | Out-Null
    $until = if ($null -eq $SafeUntil) { $Occurrence.AddDays(10) } else { [datetime]$SafeUntil }
    $document = [ordered]@{ SchemaVersion = 1; ClaimId = [guid]::NewGuid().ToString('N'); JobName = $Job.Name; OccurrenceUtc = $Occurrence.ToUniversalTime().ToString('o'); OwnerServer = 'REMOTE'; PlanId = 'synthetic'; OrchestratorPid = 1; Status = $Status; Attempt = 0; CreatedAtUtc = $Occurrence.ToUniversalTime().ToString('o'); UpdatedAtUtc = $Occurrence.ToUniversalTime().AddMinutes(1).ToString('o'); SafeUntilUtc = $until.ToUniversalTime().ToString('o') }
    $path = Join-Path $folder ($Occurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ') + '.json.txt')
    [IO.File]::WriteAllText($path, ($document | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
}
function Write-SyntheticRunState($Job, [string]$Status, $Staged) {
    $scriptName = [IO.Path]::GetFileName($Job.ScriptPath)
    $document = [ordered]@{ Owner = 'SmartInventory-SourceRun'; ContractVersion = '1.0'; Producer = $scriptName; Status = $Status; RunId = [guid]::NewGuid().ToString('N'); StartedAtUtc = [datetime]::UtcNow.ToString('o'); PublicationStarted = $false; UnqualifiedPublication = $false }
    if ($null -ne $Staged) { $document.StagedPublication = $Staged }
    $path = Join-Path $latest ('SmartInventory_{0}.run.json.txt' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
    [IO.File]::WriteAllText($path, ($document | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
}
function Activity($Job, [datetime]$Downstream, [datetime]$Now) {
    & $gate { param($j, $d, $n) Get-OrchestratorDependencyActivity -DependencyJob $j -DownstreamOccurrence $d -Now $n } $Job $Downstream $Now
}

try {
    & $transport { function script:Get-SmartM365JsonTransportPolicy { @{ Mode = 'JsonText'; QualifiedUncRoots = @(); QualifiedLocalDataRoots = @() } } }
    & $gate { param($c, $s, $j) $script:Settings = [pscustomobject]@{ ElectionClaimsPath = $c; SharedDataFolderPath = $s; ElectionClaimGraceMinutes = 15 }; $script:Manifest = [pscustomobject]@{ OrderedJobs = $j }; $script:DependencyFreshCache = @{} } $claims $shared $jobs
    $global:SmartM365TenantKey = 'synthetic'

    # Weekday: ManagedApps retry at 05:58 blocks the 07:45 preparation, staged or not.
    $monday = [datetime]::new(2026, 10, 19, 0, 0, 0, [DateTimeKind]::Local)
    $d0745 = $monday.AddHours(7).AddMinutes(45)
    Write-SyntheticClaim $managedApps $monday.AddDays(-1).AddHours(5) 'Success'
    foreach ($status in @('Claimed', 'Running', 'RetryScheduled')) {
        Write-SyntheticClaim $managedApps $monday.AddHours(5) $status
        foreach ($staged in @($false, $true)) {
            Write-SyntheticRunState $managedApps 'Collecting' $staged
            $result = Activity $managedApps $d0745 $monday.AddHours(5).AddMinutes(58)
            Check ($result.State -eq 'Waiting' -and $result.Detail -match 'expected to finish before') "Morning dependency $status (staged=$staged) did not block the 07:45 preparation."
        }
    }
    # Final failure today with a fresh success yesterday: activity clear, freshness accepts yesterday.
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'Failed'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A final failure still blocked preparation.'
    $fresh = & $gate { param($j, $n) $script:DependencyFreshCache = @{}; Get-OrchestratorDependencyFreshStatus -DependencyJob $j -MaxAgeHours (Get-OrchestratorDependencyMaxAge -DependencyJob $j -OverrideHours 48 -Now $n) -Now $n } $managedApps $monday.AddHours(7).AddMinutes(45)
    Check ($fresh -eq 'Ready') 'Yesterday success was not accepted after a final failure today.'
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'Success'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A success today still blocked preparation.'
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'CompletedWithWarnings'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A success with warnings today still blocked preparation.'

    # Catch-up of the previous day occurrence while today's run is active: staged only is safe.
    $tuesday = $monday.AddDays(1)
    Write-SyntheticClaim $managedApps $tuesday.AddHours(5) 'Running'
    Write-SyntheticRunState $managedApps 'Collecting' $false
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Waiting') 'A direct-publication run did not block a catch-up preparation.'
    Write-SyntheticRunState $managedApps 'Collecting' $null
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Waiting') 'A run state without the staging flag was trusted.'
    Write-SyntheticRunState $managedApps 'Collecting' $true
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Clear') 'A staged run blocked a catch-up preparation.'
    Write-SyntheticRunState $managedApps 'Publishing' $true
    $result = Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)
    Check ($result.State -eq 'Waiting' -and $result.Detail -match 'promoting') 'A staged run promoting its CSVs did not block.'
    Write-SyntheticClaim $managedApps $tuesday.AddHours(5) 'Success'

    # Saturday 07:45: DiscoveredApps is due at 20:00 only; the previous weekly run succeeded.
    $saturday = [datetime]::new(2026, 10, 24, 0, 0, 0, [DateTimeKind]::Local)
    Write-SyntheticClaim $apps $saturday.AddDays(-7).AddHours(20) 'Success'
    $result = Activity $apps $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)
    Check ($result.State -eq 'Clear') 'The Saturday 20:00 weekly run blocked the Saturday morning preparation.'

    # Sunday 07:45: the 25-hour weekly run is in progress.
    $sunday = $saturday.AddDays(1)
    Write-SyntheticClaim $apps $saturday.AddHours(20) 'Running'
    Write-SyntheticRunState $apps 'Collecting' $true
    $result = Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)
    Check ($result.State -eq 'Clear' -and $result.Detail -match 'staged') 'A staged long weekly run blocked the Sunday preparation.'
    Write-SyntheticRunState $apps 'Collecting' $false
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Waiting') 'A direct-publication long run did not block.'
    Write-SyntheticRunState $apps 'Collecting' $true

    # An expired non-terminal claim (lost owner) does not block; stale archives are ignored.
    Write-SyntheticClaim $apps $saturday.AddHours(20) 'Running' $sunday.AddHours(6)
    Write-SyntheticRunState $apps 'Collecting' $false
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'An expired claim blocked preparation.'
    $folder = Join-Path $claims $apps.Name
    $archive = Join-Path $folder ($sunday.ToUniversalTime().AddHours(1).ToString('yyyyMMddTHHmmssfffZ') + '.json.stale.abc.json.txt')
    [IO.File]::WriteAllText($archive, '{"JobName":"Intune-DiscoveredApps-Inventory","Status":"Running"}')
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A stale claim archive was read as an execution.'

    # Saturday EXO Fast queued behind the full run since 00:15: not started beyond the grace.
    Check ((Activity $fast $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A queued dependency held preparation all Saturday.'
    # A dependency due shortly before the downstream occurrence and not yet claimed blocks during the grace.
    $result = Activity $late $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(46)
    Check ($result.State -eq 'Waiting' -and $result.Detail -match 'not started') 'A due, unclaimed dependency did not block during the claim grace.'
    Check ((Activity $late $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(56)).State -eq 'Clear') 'An unclaimed dependency blocked beyond the claim grace.'

    # A script shared by several jobs cannot prove staging through its single run state.
    $full = [pscustomobject]@{ Name = 'EXO-Mailboxes-Inventory'; ScriptPath = $fast.ScriptPath; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 2160; EstimatedDurationMinutes = 1200; Schedule = [pscustomobject]@{ Type = 'Weekly'; Times = @('20:00'); DaysOfWeek = @('Friday') } }
    & $gate { param($j) $script:Manifest = [pscustomobject]@{ OrderedJobs = $j } } (@($jobs) + $full)
    Write-SyntheticClaim $full $saturday.AddDays(-1).AddHours(20) 'Running'
    Write-SyntheticRunState $full 'Collecting' $true
    Check ((Activity $full $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)).State -eq 'Waiting') 'A shared-script run state was trusted.'

    # Run state reader.
    Write-SyntheticRunState $apps 'Collecting' $true
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job $apps
    Check ($state.Known -and $state.StagedPublication -and $state.Status -eq 'Collecting') 'Staged run state not read.'
    Check (-not (Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job $prepare).Known) 'Absent run state reported as known.'
    Check (-not (Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath (Join-Path $root 'Other') -Job $apps).Known) 'Run state read outside the DATA-ALL/Orchestrator layout.'

    # GUI readiness explains the same wait.
    Write-SyntheticClaim $managedApps $sunday.AddHours(5) 'RetryScheduled'
    Write-SyntheticRunState $managedApps 'Failed' $false
    $readiness = @(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $shared -JobsDocument ([pscustomobject]@{ Jobs = $jobs }) -JobName $prepare.Name -Now $sunday.AddHours(7))
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Waiting' -and $readiness[0].Detail -match 'RetryScheduled') 'GUI readiness ignored a dependency awaiting retry.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Output "PASS: $script:passed dependency activity checks. Synthetic claims only."
