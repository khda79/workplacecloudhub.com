<#
.SYNOPSIS
Offline tests for the elected RunOnce catch-up guard: a due occurrence already covered by a newer or later-started shared claim is not replayed.
.VERSION
1.0
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$null,[ref]$null)
function Get-FunctionText([string]$Name){$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true).Extent.Text}
$defs=foreach($name in @('Get-OrchestratorOccurrenceCoveredByLaterClaim','Get-OrchestratorClaimOccurrenceUtc','ConvertTo-OrchestratorUtcTime')){Get-FunctionText $name}
$gate=New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root=Join-Path $env:TEMP ('elected-catchup-'+[guid]::NewGuid().ToString('N'))
$claims=Join-Path $root 'Claims'
New-Item -ItemType Directory $claims|Out-Null
$script:passed=0
function Check($Value,$Message){if(-not $Value){throw $Message};$script:passed++}
try {
    & $transport {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()}}}
    & $gate {param($c) $script:Settings=[pscustomobject]@{ElectionClaimsPath=$c};$script:DependencyFreshCache=@{}} $claims

    # Incident shape: the schedule moved from 08:30 to 07:45; server A ran the 08:30 occurrence,
    # server B became owner and evaluates the previous day's 07:45 occurrence at 05:54.
    $job=[pscustomobject]@{Name='CatchUpJob';AssignmentMode='Elected';Schedule=[pscustomobject]@{Type='Daily';Times=@('07:45');DaysOfWeek=@()}}
    $now=[datetime]::new(2026,10,20,5,54,0,[DateTimeKind]::Local)
    $due=[datetime]::new(2026,10,19,7,45,0,[DateTimeKind]::Local)
    $previousSchedule=[datetime]::new(2026,10,19,8,30,0,[DateTimeKind]::Local)
    $folder=Join-Path $claims $job.Name
    function Coverage { & $gate {param($j,$o,$n) $script:DependencyFreshCache=@{}; Get-OrchestratorOccurrenceCoveredByLaterClaim -Job $j -ExpectedOccurrence $o -Now $n} $job $due $now }
    function Initialize-ClaimFolder { if(Test-Path $folder){Get-ChildItem $folder -File|Remove-Item -Force} }
    function Write-Claim([datetime]$Occurrence,[string]$Status,[datetime]$Created,[datetime]$Updated,[string]$JobName=$job.Name) {
        New-Item -ItemType Directory $folder -Force|Out-Null
        $document=[ordered]@{SchemaVersion=1;ClaimId=[guid]::NewGuid().ToString('N');JobName=$JobName;OccurrenceUtc=$Occurrence.ToUniversalTime().ToString('o');OwnerServer='SERVER-A';PlanId='synthetic';OrchestratorPid=1;Status=$Status;Attempt=0;CreatedAtUtc=$Created.ToUniversalTime().ToString('o');UpdatedAtUtc=$Updated.ToUniversalTime().ToString('o')}
        $path=Join-Path $folder ($Occurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')
        [IO.File]::WriteAllText($path,($document|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        $path
    }

    Check ($null -eq (Coverage)) 'Missing claims folder covered the due occurrence.'
    New-Item -ItemType Directory $folder -Force|Out-Null
    Check ($null -eq (Coverage)) 'Empty claims folder covered the due occurrence.'

    foreach($status in @('Success','CompletedWithWarnings')) {
        Initialize-ClaimFolder
        Write-Claim $previousSchedule $status $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
        $coverage=Coverage
        Check ($null -ne $coverage -and $coverage.Occurrence -eq $previousSchedule -and $coverage.Status -eq $status -and $coverage.OwnerServer -eq 'SERVER-A') "Newer $status claim did not cover the missed occurrence."
    }
    foreach($status in @('Failed','TimedOut','Interrupted','Cancelled','Claimed','Running','RetryScheduled')) {
        Initialize-ClaimFolder
        Write-Claim $previousSchedule $status $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
        Check ($null -eq (Coverage)) "Newer $status claim covered the missed occurrence."
    }

    # An older occurrence counts only when its run started at or after the due occurrence.
    $older=$due.AddDays(-1)
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $due.AddMinutes(15) $due.AddMinutes(55)|Out-Null
    $coverage=Coverage
    Check ($null -ne $coverage -and $coverage.Occurrence -eq $older) 'Late catch-up started after the due occurrence did not cover it.'
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $older $older.AddMinutes(40)|Out-Null
    Check ($null -eq (Coverage)) 'An older run started before the due occurrence covered it.'
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $due.AddMinutes(-30) $due.AddMinutes(30)|Out-Null
    Check ($null -eq (Coverage)) 'A run in progress at the due occurrence covered it.'

    # The due occurrence itself is left to the existing peer claim check.
    Initialize-ClaimFolder
    Write-Claim $due 'Success' $due $due.AddMinutes(40)|Out-Null
    Check ($null -eq (Coverage)) 'The claim of the due occurrence itself was reported as a later coverage.'

    # Claims that cannot be trusted never cover.
    Initialize-ClaimFolder
    Write-Claim $now.AddHours(2) 'Success' $now.AddHours(2) $now.AddHours(3)|Out-Null
    Check ($null -eq (Coverage)) 'A claim for a future occurrence covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $now.AddHours(1)|Out-Null
    Check ($null -eq (Coverage)) 'A claim completed in the future covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(-5)|Out-Null
    Check ($null -eq (Coverage)) 'A claim completed before it started covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40) 'OtherJob'|Out-Null
    Check ($null -eq (Coverage)) 'A claim of another job covered the due occurrence.'
    Initialize-ClaimFolder
    $corrupt=Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40)
    [IO.File]::WriteAllText($corrupt,'{')
    Check ($null -eq (Coverage)) 'An unreadable claim covered the due occurrence.'
    # A trusted claim still covers when a newer claim is unreadable.
    Write-Claim $due.AddMinutes(30) 'Success' $due.AddMinutes(30) $due.AddMinutes(70)|Out-Null
    $coverage=Coverage
    Check ($null -ne $coverage -and $coverage.Occurrence -eq $due.AddMinutes(30)) 'An unreadable newer claim hid a valid covering claim.'

    # Without coverage the next occurrence is unaffected: normal 07:45 run after the incident day.
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
    $today=[datetime]::new(2026,10,20,7,45,0,[DateTimeKind]::Local)
    $todayCoverage=& $gate {param($j,$o,$n) $script:DependencyFreshCache=@{}; Get-OrchestratorOccurrenceCoveredByLaterClaim -Job $j -ExpectedOccurrence $o -Now $n} $job $today $today.AddSeconds(20)
    Check ($null -eq $todayCoverage) 'The next scheduled occurrence was covered by the previous day.'

    # The launch phase applies the guard to due elected occurrences before dependency gates.
    $launch=Get-FunctionText 'Invoke-LaunchPhase'
    $guardIndex=$launch.IndexOf('Get-OrchestratorOccurrenceCoveredByLaterClaim')
    Check ($guardIndex -gt 0 -and $guardIndex -lt $launch.IndexOf('# Dependency gate')) 'Launch phase does not apply the covered-occurrence guard before dependency gates.'
    Check ($launch -match "\`$reason -eq 'due' -and \`$job\.AssignmentMode -eq 'Elected'") 'Covered-occurrence guard is not limited to due elected occurrences.'
    Check ($launch.Substring($guardIndex,1200) -match 'Set-OccurrenceSkipped') 'Covered occurrence is not recorded as Skipped.'

    # A momentarily unreachable share must mean "not covered" with a warning, never a scheduler crash.
    $launchAst=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-LaunchPhase'},$true)
    $guardIf=$launchAst.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -match "\`$reason -eq 'due' -and \`$job\.AssignmentMode -eq 'Elected'"},$true)
    Check ($null -ne $guardIf) 'Covered-occurrence guard block not found in the launch phase.'
    $harness=New-Module -ScriptBlock ([scriptblock]::Create(($defs -join "`n") + @"

function Get-ChildItem { throw 'Synthetic share outage.' }
function Write-OrchestratorRuntimeUpdateWarning { param(`$Key,`$Message,`$Now) `$script:warnings.Add(`$Message) }
function Set-OccurrenceSkipped { param(`$JobName,`$Occurrence,`$Reason) `$script:skipped=`$true }
function Clear-DependencyWaitLog { param(`$JobName) }
function Invoke-Guard { param(`$job,`$occurrence,`$Now)
    `$name=`$job.Name; `$reason='due'; `$script:reachedLaunch=`$false
    foreach(`$tick in 1) {
$($guardIf.Extent.Text)
        `$script:reachedLaunch=`$true
    }
}
"@))
    try {
        $outage=& $harness {param($c,$j,$o,$n) $script:Settings=[pscustomobject]@{ElectionClaimsPath=$c};$script:DependencyFreshCache=@{};$script:warnings=[Collections.Generic.List[string]]::new();$script:skipped=$false
            Invoke-Guard -job $j -occurrence $o -Now $n
            [pscustomobject]@{Reached=$script:reachedLaunch;Skipped=$script:skipped;Warnings=@($script:warnings)}} $claims $job $due $now
        Check ($outage.Reached -and -not $outage.Skipped) 'A share outage skipped the due occurrence or stopped the launch phase.'
        Check ($outage.Warnings.Count -eq 1 -and $outage.Warnings[0] -match 'Synthetic share outage') 'A share outage was not reported as a warning.'
    } finally { Remove-Module $harness -Force }
    [pscustomobject]@{Passed=$script:passed;Fixture=$root;Scope='Synthetic claims only; no production jobs or network access'}
} finally {
    & $transport {param($p) Set-Item Function:script:Get-SmartM365JsonTransportPolicy $p} $policy
    Remove-Module $gate -Force
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
