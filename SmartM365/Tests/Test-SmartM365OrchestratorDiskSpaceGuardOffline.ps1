<#
.SYNOPSIS
Offline tests of the disk space guard (orchestrator 1.5.75): free space of the DATA share, launch
pause below the block threshold with hysteresis, pending reason, heartbeat, peer monitoring,
cluster configuration, cluster-wide alert mails and GUI data.
.VERSION
1.0.0
.NOTES
Runs the real module functions and the real orchestrator functions extracted from the script with
stubbed measurements, logs and mails, in a temporary folder. The only real measurement reads the
free space of the temporary folder volume. No tenant initialization, collectors or publication.
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
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DiskSpace-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    # 1. Real measurement through GetDiskFreeSpaceEx, hung probe and missing folder.
    $real = Get-SmartM365OrchestratorVolumeFreeSpace -Path $root -TimeoutSeconds 10
    Check ($real.Ready -and $real.TotalGB -gt 0 -and $real.FreePercent -ge 0 -and $real.FreePercent -le 100 -and $real.FreeGB -le $real.TotalGB -and -not $real.Error) "The real measurement must return the free space of the volume ($($real.Error))."
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $hung = Get-SmartM365OrchestratorVolumeFreeSpace -Path $root -TimeoutSeconds 2 -ProbeScript 'param($Path, $Answers) Start-Sleep -Seconds 30'
    $clock.Stop()
    Check (-not $hung.Ready -and $hung.Error -match 'No answer within 2 second' -and $clock.Elapsed.TotalSeconds -lt 8) "A hung share must be abandoned within the timeout ($([math]::Round($clock.Elapsed.TotalSeconds, 1)) s)."
    $missing = Get-SmartM365OrchestratorVolumeFreeSpace -Path (Join-Path $root 'missing\folder') -TimeoutSeconds 10
    Check (-not $missing.Ready -and $missing.Error -match 'Win32 error') "A missing folder must be reported as an error ($($missing.Error))."
    Check (-not (Get-SmartM365OrchestratorVolumeFreeSpace -Path '' -TimeoutSeconds 1).Ready) 'An empty path must not be measured.'

    # 2. Pure decision with hysteresis.
    function New-Measurement([double]$Percent, [string]$Failure = '') {
        if ($Failure) { return [pscustomobject]@{ Path = 'X'; Ready = $false; FreePercent = $null; FreeGB = $null; TotalGB = $null; Error = $Failure } }
        [pscustomobject]@{ Path = 'X'; Ready = $true; FreePercent = $Percent; FreeGB = [math]::Round(135 * $Percent / 100, 1); TotalGB = 135; Error = '' }
    }
    $t0 = [datetime]::new(2026, 10, 29, 10, 0, 0, [DateTimeKind]::Utc)
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 25) -PreviousState $null -NowUtc $t0
    Check ($d.State -eq 'OK' -and -not $d.Blocked -and $d.Detail -eq 'DATA free 25 % (33.8 GiB of 135 GiB).') "25 % must be OK with invariant figures ($($d.Detail))."
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 14) -PreviousState $d -NowUtc $t0.AddMinutes(5)
    Check ($d.State -eq 'Warning' -and -not $d.Blocked -and $d.Detail -match 'below the 15 % warning level') '14 % must warn without pausing.'
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 9.9) -PreviousState $d -NowUtc $t0.AddMinutes(10)
    Check ($d.State -eq 'Low' -and $d.Blocked -and $d.SinceUtc -eq $t0.AddMinutes(10).ToString('o') -and $d.Detail -match 'new launches paused') '9.9 % must pause launches.'
    $lowSince = $d.SinceUtc
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 11) -PreviousState $d -NowUtc $t0.AddMinutes(15)
    Check ($d.Blocked -and $d.SinceUtc -eq $lowSince) '11 % must keep the pause (hysteresis) and its start time.'
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 0 'share hung') -PreviousState $d -NowUtc $t0.AddMinutes(20)
    Check ($d.Blocked -and $d.State -eq 'Low' -and $d.FreePercent -eq 11 -and $d.Detail -match 'launches stay paused') 'A failed measurement must never lift a pause and keeps the last figures.'
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 11.9) -PreviousState $d -NowUtc $t0.AddMinutes(25)
    Check ($d.Blocked) '11.9 % must keep the pause.'
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 12) -PreviousState $d -NowUtc $t0.AddMinutes(30)
    Check (-not $d.Blocked -and $d.State -eq 'Warning') '12 % must resume launches.'
    $d = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 10.5) -PreviousState $d -NowUtc $t0.AddMinutes(35)
    Check (-not $d.Blocked) 'Above the block threshold, an unpaused server must not pause.'
    $u = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 0 'timeout') -PreviousState $d -NowUtc $t0.AddMinutes(40)
    Check ($u.State -eq 'Unknown' -and -not $u.Blocked -and $u.UnknownSinceUtc -eq $t0.AddMinutes(40).ToString('o') -and $u.Detail -match 'Launches continue') 'A failed measurement must not pause launches by itself.'
    $u = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 0 'timeout') -PreviousState $u -NowUtc $t0.AddMinutes(45)
    Check ($u.UnknownSinceUtc -eq $t0.AddMinutes(40).ToString('o')) 'Unknown keeps its start time.'
    $h = Get-SmartM365OrchestratorDiskSpaceDecision -Measurement (New-Measurement 11) -PreviousState @{ State = 'Low'; Blocked = $true; SinceUtc = 'x' } -NowUtc $t0
    Check ($h.Blocked -and $h.SinceUtc -eq 'x') 'A persisted hashtable decision must keep the pause after a restart.'

    # 3. Cluster configuration validation and mapping.
    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A'); DiskSpaceBlockFreePercent = 10; DiskSpaceResumeFreePercent = 12; DiskSpaceWarningFreePercent = 15; DiskSpaceCheckMinutes = 5 }
    Check ((Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'Default disk space thresholds must be valid.'
    foreach ($bad in @(@{ DiskSpaceResumeFreePercent = 10 }, @{ DiskSpaceWarningFreePercent = 11 }, @{ DiskSpaceBlockFreePercent = 0 }, @{ DiskSpaceBlockFreePercent = 60; DiskSpaceResumeFreePercent = 70; DiskSpaceWarningFreePercent = 80 }, @{ DiskSpaceCheckMinutes = 0 })) {
        $copy = $cluster | ConvertTo-Json | ConvertFrom-Json
        foreach ($key in $bad.Keys) { $copy.$key = $bad[$key] }
        Check (-not (Test-SmartM365OrchestratorClusterDocument -Document $copy).Valid) ("Invalid disk space configuration must be rejected: {0}" -f (($bad.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join ', '))
    }
    Check ((Test-SmartM365OrchestratorClusterDocument -Document ([pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A') })).Valid) 'A cluster document without disk space keys stays valid.'
    $clusterSettingsText = Get-FunctionText 'Set-OrchestratorCentralClusterSettings'
    foreach ($key in @('DiskSpaceGuardEnabled', 'DiskSpaceBlockFreePercent', 'DiskSpaceResumeFreePercent', 'DiskSpaceWarningFreePercent', 'DiskSpaceCheckMinutes', 'DiskSpaceCheckTimeoutSeconds', 'DiskSpaceAlertRepeatHours')) {
        Check ($clusterSettingsText -match ("\b{0} = '(Int32|Boolean)'" -f $key)) "The central cluster configuration must map $key."
    }
    $template = Get-Content -Raw (Join-Path $base 'Orchestrator-Cluster.json.template') | ConvertFrom-Json
    Check ($template.DiskSpaceBlockFreePercent -eq 10 -and $template.DiskSpaceResumeFreePercent -eq 12 -and $template.DiskSpaceWarningFreePercent -eq 15 -and (Test-SmartM365OrchestratorClusterDocument -Document $template).Valid) 'The cluster template must carry valid disk space defaults.'

    # 4. Real orchestrator guard update with scripted measurements.
    foreach ($name in @('Get-OrchestratorDiskSpaceThresholds', 'Update-OrchestratorDiskSpaceGuard', 'Invoke-LaunchPhase', 'Get-OrchestratorPendingJobSnapshot', 'Get-JobState', 'Get-DueOccurrence', 'Get-JobOccurrencesInWindow', 'ConvertTo-StateTime', 'ConvertFrom-StateTime', 'Send-OrchestratorDiskSpaceAlert', 'ConvertTo-OrchestratorUtcTime', 'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock')) {
        . ([scriptblock]::Create((Get-FunctionText $name)))
    }
    $script:Logs = [Collections.Generic.List[object]]::new()
    function Write-OrchestratorLog([string]$Message, [string]$Level = 'INFO') { $script:Logs.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
    function Write-OrchestratorRuntimeUpdateWarning([string]$Key, [string]$Message, $Now, $Level) { $script:Logs.Add([pscustomobject]@{ Level = 'WARN'; Message = $Message }) }
    $script:AlertCalls = [Collections.Generic.List[object]]::new()
    $script:RealSendAlert = ${function:Send-OrchestratorDiskSpaceAlert}
    function Send-OrchestratorDiskSpaceAlert([datetime]$Now, [switch]$Resumed) { $script:AlertCalls.Add([bool]$Resumed) }
    $script:NextMeasurements = [Collections.Generic.Queue[object]]::new()
    $script:Measured = 0
    function Get-SmartM365OrchestratorVolumeFreeSpace([string]$Path, [int]$TimeoutSeconds) { $script:Measured++; $script:NextMeasurements.Dequeue() }

    $script:Settings = [pscustomobject]@{
        DiskSpaceGuardEnabled = $true; DiskSpaceDataPath = $root; DiskSpaceBlockFreePercent = 10; DiskSpaceResumeFreePercent = 12; DiskSpaceWarningFreePercent = 15
        DiskSpaceCheckMinutes = 5; DiskSpaceCheckTimeoutSeconds = 10; DiskSpaceAlertRepeatHours = 6; SharedDataFolderPath = $root; MaxConcurrency = 2
    }
    $script:State = @{ Jobs = @{} }
    $script:DiskSpaceGuard = $null
    $script:LastDiskSpaceCheckAt = [datetime]::MinValue
    $now = [datetime]'2026-10-29T10:00:00'
    foreach ($percent in @(25, 9.5)) { $script:NextMeasurements.Enqueue((New-Measurement $percent)) }
    Update-OrchestratorDiskSpaceGuard -Now $now
    Check ($script:DiskSpaceGuard.State -eq 'OK' -and $script:Measured -eq 1) 'The first tick must measure.'
    Update-OrchestratorDiskSpaceGuard -Now $now.AddMinutes(2)
    Check ($script:Measured -eq 1) 'The share must not be measured again before DiskSpaceCheckMinutes.'
    Update-OrchestratorDiskSpaceGuard -Now $now.AddMinutes(5)
    Check ($script:DiskSpaceGuard.Blocked -and $script:State.DiskSpaceGuard.Blocked -and @($script:Logs | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -match '^Disk space guard: new launches paused on' }).Count -eq 1) 'Crossing the threshold must pause, persist and log once.'
    # Recycle: a new instance reads the persisted decision; 11 % keeps the pause.
    $script:DiskSpaceGuard = $null; $script:LastDiskSpaceCheckAt = [datetime]::MinValue
    $script:State = $script:State | ConvertTo-Json -Depth 5 | ConvertFrom-Json -AsHashtable
    $script:NextMeasurements.Enqueue((New-Measurement 11))
    Update-OrchestratorDiskSpaceGuard -Now $now.AddMinutes(10)
    Check ($script:DiskSpaceGuard.Blocked) 'The persisted pause must survive a recycle until the resume threshold.'

    # 5. Launch phase: nothing is evaluated while paused; forced runs are deferred once.
    $script:MaintenanceHealthy = $true
    $script:ForcedPending = @('Job-Forced')
    $script:DiskSpaceDeferredForced = @{}
    $script:HandoffCalls = 0
    function Invoke-RetryHandoffAdoption($Now) { $script:HandoffCalls++ }
    $script:Manifest = [pscustomobject]@{ OrderedJobs = @(); JobsByName = @{} }
    Invoke-LaunchPhase -Now $now.AddMinutes(10)
    Invoke-LaunchPhase -Now $now.AddMinutes(11)
    Check ($script:HandoffCalls -eq 0 -and $script:ForcedPending -contains 'Job-Forced') 'A paused server must not evaluate any launch and must keep forced runs queued.'
    Check (@($script:Logs | Where-Object { $_.Message -match '^Job Job-Forced: -Force deferred until DATA disk space recovers' }).Count -eq 1) 'A deferred forced run must be logged once.'

    # 6. Pending jobs: only what would start now, with the LowDiskSpace reason.
    function Test-JobAllowedOnServer($Job, [switch]$AllowManual) { $true }
    function Get-MaintenanceCatchUpOccurrence { $null }
    function Get-SmartM365OrchestratorConcurrencyLease { $null }
    $script:MaintenanceControl = [pscustomobject]@{ Enabled = $false; Reason = ''; Revision = 1 }
    $script:PipelineWaitState = @{}; $script:DependencyWaitLogState = @{}
    $script:RunningJobs = @{ 'Job-Running' = @{} }
    $script:PipelinePending = @{ 'Job-Request' = @{ BatchId = 'B1' } }
    $daily = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00') }
    $jobs = foreach ($jobName in @('Job-Due', 'Job-Running', 'Job-Retry', 'Job-Request', 'Job-Done', 'Job-Forced')) { [pscustomobject]@{ Name = $jobName; Enabled = $true; Schedule = $daily; ConcurrencyKey = '' } }
    $script:Manifest = [pscustomobject]@{ OrderedJobs = @($jobs); JobsByName = @{} }
    $snapshotNow = [datetime]'2026-10-29T10:20:00'
    $handled = ConvertTo-StateTime -Value ([datetime]'2026-10-29T01:00:00')
    foreach ($jobName in @('Job-Running', 'Job-Retry', 'Job-Request', 'Job-Done', 'Job-Forced')) { (Get-JobState -JobName $jobName).LastScheduledOccurrence = $handled }
    (Get-JobState -JobName 'Job-Retry').PendingRetry = @{ Attempt = 1; ScheduledOccurrence = $handled; NotBefore = $handled }
    $pending = @(Get-OrchestratorPendingJobSnapshot -Now $snapshotNow)
    $names = @($pending | ForEach-Object Name | Sort-Object)
    Check (($names -join ',') -eq 'Job-Due,Job-Forced,Job-Request,Job-Retry') "Pending must list exactly what would start now ($($names -join ','))."
    Check (@($pending | Where-Object Reason -ne 'LowDiskSpace').Count -eq 0 -and @($pending | Where-Object Name -eq 'Job-Request')[0].Details -match '^Request B1 waits for DATA disk space\. DATA free 11 %') 'Every waiting job must show LowDiskSpace and the guard detail.'
    Check ($null -eq (Get-JobState -JobName 'Job-Due').LastScheduledOccurrence) 'A paused occurrence must not be consumed.'
    $script:DiskSpaceGuard = $null
    $unpaused = @(Get-OrchestratorPendingJobSnapshot -Now $snapshotNow)
    Check (@($unpaused | Where-Object Name -eq 'Job-Due')[0].Reason -eq 'DueNotStarted' -and @($unpaused | Where-Object Reason -eq 'LowDiskSpace').Count -eq 0) 'Without a pause the usual reasons come back.'

    # 7. Resume: launches are evaluated again and the alert is told the server resumed.
    $script:DiskSpaceGuard = [pscustomobject]$script:State.DiskSpaceGuard
    $script:LastDiskSpaceCheckAt = [datetime]::MinValue
    $script:NextMeasurements.Enqueue((New-Measurement 12.5))
    $script:AlertCalls.Clear()
    Update-OrchestratorDiskSpaceGuard -Now $now.AddMinutes(30)
    Check (-not $script:DiskSpaceGuard.Blocked -and $script:AlertCalls[0] -eq $true -and @($script:Logs | Where-Object { $_.Message -match '^Disk space guard: launches resumed on' }).Count -eq 1) 'Reaching the resume threshold must resume and report it.'
    $script:Manifest = [pscustomobject]@{ OrderedJobs = @(); JobsByName = @{} }
    Invoke-LaunchPhase -Now $now.AddMinutes(30)
    Check ($script:HandoffCalls -eq 1 -and $script:DiskSpaceDeferredForced.Count -eq 0) 'After the resume the launch phase must run again.'
    $script:Settings.DiskSpaceGuardEnabled = $false
    Update-OrchestratorDiskSpaceGuard -Now $now.AddMinutes(40)
    Check ($null -eq $script:DiskSpaceGuard -and -not $script:State.ContainsKey('DiskSpaceGuard')) 'Disabling the guard must clear the decision.'
    $script:Settings.DiskSpaceGuardEnabled = $true

    # 8. Cluster-wide alert mails through the shared alert file.
    Set-Item -Path function:Send-OrchestratorDiskSpaceAlert -Value $script:RealSendAlert
    $script:Mails = [Collections.Generic.List[object]]::new()
    $Tenant = 'test'
    function New-SmartM365EmailBody([string]$Title, [string]$Category, [string]$Message, $SummaryData) { $Title }
    function Send-OrchestratorMail([string]$Subject, [string]$HtmlBody, [switch]$IsError) { $script:Mails.Add([pscustomobject]@{ Subject = $Subject; IsError = [bool]$IsError }); $true }
    function Get-SmartM365JsonReadPath([string]$Path, [switch]$Optional) { if (Test-Path -LiteralPath $Path) { $Path } }
    function Read-SmartM365JsonDocument([string]$Path) { [pscustomobject]@{ Document = (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json) } }
    function Write-FileAtomically([string]$Path, [string]$Content) { [IO.File]::WriteAllText($Path, $Content) }
    function Set-Guard([string]$State, [double]$Percent, [bool]$Blocked, [string]$UnknownSince = '') {
        $script:DiskSpaceGuard = [pscustomobject]@{ State = $State; Blocked = $Blocked; FreePercent = $Percent; FreeGB = 1; TotalGB = 135; Detail = "detail $Percent"; UnknownSinceUtc = $UnknownSince; BlockFreePercent = 10; ResumeFreePercent = 12; WarningFreePercent = 15 }
    }
    $mailNow = [datetime]'2026-10-29T12:00:00'
    Set-Guard 'Low' 9 $true; Send-OrchestratorDiskSpaceAlert -Now $mailNow
    Check ($script:Mails.Count -eq 1 -and $script:Mails[0].IsError -and $script:Mails[0].Subject -match 'new launches paused$') 'The first paused server must send the pause mail.'
    Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddMinutes(5)
    Check ($script:Mails.Count -eq 1) 'Another paused server must not repeat the pause mail.'
    Set-Guard 'Warning' 10.5 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddMinutes(6)
    Check ($script:Mails.Count -eq 1) 'A server that never paused must not send a resume or warning mail while the cluster alert is a pause.'
    Set-Guard 'Low' 9 $true; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(6)
    Check ($script:Mails.Count -eq 2) 'The pause mail must repeat after DiskSpaceAlertRepeatHours.'
    Set-Guard 'Warning' 12.5 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(7) -Resumed
    Check ($script:Mails.Count -eq 3 -and -not $script:Mails[2].IsError -and $script:Mails[2].Subject -match 'launches resumed$') 'The resumed server must send the resume mail.'
    Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(7).AddMinutes(1) -Resumed
    Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(7).AddMinutes(2)
    Check ($script:Mails.Count -eq 3) 'The resume mail must be sent once and must not be followed by a warning mail.'
    Set-Guard 'OK' 16 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(8)
    Set-Guard 'Warning' 14.5 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(8)
    Check ($script:Mails.Count -eq 3) 'Free space just above the warning level must not re-arm the warning mail.'
    Set-Guard 'OK' 17.5 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(9)
    Set-Guard 'Warning' 14 $false; Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(9)
    Check ($script:Mails.Count -eq 4 -and $script:Mails[3].Subject -match 'warning$') 'After re-arming, a new warning must be mailed once.'
    Set-Guard 'Unknown' $null $false ($mailNow.AddHours(10).ToUniversalTime().ToString('o')); Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(10).AddMinutes(20)
    Check ($script:Mails.Count -eq 4) 'An unknown measurement must wait 30 minutes before mailing.'
    Send-OrchestratorDiskSpaceAlert -Now $mailNow.AddHours(10).AddMinutes(31)
    Check ($script:Mails.Count -eq 5 -and $script:Mails[4].Subject -match 'unknown$') 'An unknown measurement for 30 minutes must be mailed once.'
    Check (-not (Test-Path -LiteralPath (Join-Path $root 'Orchestrator-DiskSpaceAlert.lock'))) 'The alert lock must be released.'

    # 9. Heartbeat, peer monitoring and main loop wiring.
    $heartbeatText = Get-FunctionText 'Write-OrchestratorHeartbeat'
    foreach ($field in @('DiskSpaceProtocol', 'DiskSpaceGuardEnabled', 'DiskSpaceState', 'DiskSpaceBlocked', 'DiskSpaceFreePercent', 'DiskSpaceFreeGB', 'DiskSpaceTotalGB', 'DiskSpaceCheckedAtUtc', 'DiskSpaceDetail')) {
        Check ($heartbeatText -match ("\b{0} = " -f $field)) "The heartbeat must publish $field."
    }
    $peerText = Get-FunctionText 'Get-OrchestratorPeerHealthSnapshot'
    Check ($peerText -match 'if \(\$peerDiskSpaceBlocked\) \{ continue \}' -and $peerText -match "'WaitingConcurrency', 'LowDiskSpace'\)") 'Peer monitoring must not report missed starts of a paused peer.'
    $source = [IO.File]::ReadAllText($orchestratorPath)
    Check ($source -match 'Update-OrchestratorDiskSpaceGuard -Now \$now[\s\S]{0,400}?if \(Test-OrchestratorStatePersistenceReady\) \{\s*Invoke-LaunchPhase -Now \$now') 'The guard must be evaluated right before each launch phase.'

    # 10. GUI data: server row and dashboard item.
    foreach ($server in @('SRV-A', 'SRV-B', 'SRV-C')) { New-Item -ItemType Directory -Path (Join-Path $root "gui/$server") -Force | Out-Null }
    $hbA = [ordered]@{ Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.75'; DiskSpaceProtocol = 1; DiskSpaceGuardEnabled = $true; DiskSpaceState = 'Low'; DiskSpaceBlocked = $true; DiskSpaceFreePercent = 9.4; DiskSpaceDetail = 'DATA free 9.4 %' }
    $hbB = [ordered]@{ Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.75'; DiskSpaceProtocol = 1; DiskSpaceGuardEnabled = $true; DiskSpaceState = 'Warning'; DiskSpaceBlocked = $false; DiskSpaceFreePercent = 13; DiskSpaceDetail = 'DATA free 13 %' }
    $hbC = [ordered]@{ Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; ScriptVersion = '1.5.74' }
    foreach ($pair in @(@('SRV-A', $hbA), @('SRV-B', $hbB), @('SRV-C', $hbC))) { [IO.File]::WriteAllText((Join-Path $root "gui/$($pair[0])/Orchestrator-Heartbeat.json"), ($pair[1] | ConvertTo-Json), [Text.UTF8Encoding]::new($false)) }
    $guiCluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C'); PeerHeartbeatStaleMinutes = 5 }
    $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath (Join-Path $root 'gui') -ClusterDocument $guiCluster
    $rowA = @($operations.Servers | Where-Object Server -eq 'SRV-A')[0]; $rowB = @($operations.Servers | Where-Object Server -eq 'SRV-B')[0]; $rowC = @($operations.Servers | Where-Object Server -eq 'SRV-C')[0]
    Check ($rowA.DiskSpace -eq 'Launches paused' -and $rowA.DataFreePercent -eq 9.4 -and $rowB.DiskSpace -eq 'Warning' -and $rowC.DiskSpace -eq '' -and $null -eq $rowC.DataFreePercent) 'Server rows must show the DATA free space and the guard state.'
    $dashboard = Get-SmartM365OrchestratorDashboard -JobsDocument ([pscustomobject]@{ Jobs = @() }) -Operations $operations
    Check (@($dashboard.Actions | Where-Object { $_.Area -eq 'Disk space' -and $_.Item -eq 'SRV-A' -and $_.Severity -eq 'Critical' }).Count -eq 1 -and @($dashboard.Actions | Where-Object { $_.Area -eq 'Disk space' -and $_.Item -eq 'SRV-B' -and $_.Severity -eq 'Warning' }).Count -eq 1) 'The dashboard must raise the disk space items.'
    $guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
    Check ($guiText -match 'Header="DATA free %" Binding="\{Binding DataFreePercent\}"' -and $guiText -match 'Header="Disk space" Binding="\{Binding DiskSpace\}"' -and $guiText -match 'Low DATA disk space: new launches are paused on') 'The GUI must show the disk space columns and warn before a run request.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ("Disk space guard offline tests passed: {0} checks." -f $script:passed) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDPEUEYdNrYyb0e
# U83Bbab1rmeUGalctvZSSrMnV02t16CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC496ACImWL8MhN+bKXc8IfPzU8bcOavQQdk1TH6M1fSMA0GCSqG
# SIb3DQEBAQUABIIBgFTgQffm1ALPWh1EY0QH27wNgENNlZSK5K8Yab20hAtZpPpW
# Pv1a1HUyQwVHwFgv8quRjRjj4GNQSCA1+MFPgUTYL+OsvnLW1dE0PBfXt0bfU1kg
# bbA30qGEkDNQ/5U4zwllNou6/lAwwHer8AlLJcmhkH6UUzNOWDzAnf+Zg9y87heO
# 0HpxvLIWyyV0kGu6uGtliT+y3ftQHTpp9cmQ4RX0d3mdHqaL4VzCo54WZow0hctH
# u7OM0Z/G153EfvcPJrYxfHAaQ6+RNoBuKscXLq8TkzGlfgENtlM6LKWadaGDIwwf
# QBZqW4q1l3CdM17nuoQHG9A8GXmGCLg11f/7oQqgSVMLa3ocsG4ANuYjeZUqidMI
# /QGIV6xrnCSOyQyEDLNf8hkyaSN5oZM80Snpwqmg8uS/CkpWmtDol0bA7qGydOCi
# N0o8foFWIsr3j928wx52Vnn7QOZBMyAZgmlviFVh2MGhsD0petieF7WGF+mm+w3J
# Qf54aRwnRAyjHpnKcqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxOTEw
# MjVaMC8GCSqGSIb3DQEJBDEiBCCmAYG+Q9HLjz8W6khg85heUuLh+OPndiOdcX/q
# 6fZavzANBgkqhkiG9w0BAQEFAASCAgCuCcyhdu/VLlLVMCUxmMMDkzp3/Ve4jgJf
# i6+HJVLKoTa6wIJ+WJeUmdaymavVMVaV0BFeJ+KRwaTxRPUwFaVAWKHHENHDKovi
# G+UueNRBeCyTd8/A1d/nHy+Ybg0bD3sZ2CFYL2dT+FN2R0ytE0zDjg2nX7+a19+W
# TdTPwT69vjwK6l96Q3n3axHTYGLRs9xL5Bi6nnWSvI+zLR9HhJAQtkA+A10Esx3E
# gyHzlRDmMIp84mL8hnmK2gW5QFbpl2oknpNLgsM/HEsNbC29shBYp9A/kDG+BLLq
# 084ZkBwm5BO3aI1Hc/DUqFUK67xgKKEqgSoaxkqJk3fL5Jw3IsZXOY/zwBP+sVZF
# GeoC6Wpl0iWcBAQGgHXqlh1KhOgvxjyxH/ClTPWqOv3jA+6ubWwZ5lc4ZUzauXLQ
# QcSrbxQuVoGKReRma3m4bsoBEtzTwym5wnQjWp1nDJjqDjpfEFvuHc8QX1KAEyKp
# sAZvtpBCKfZ0jY8Ld4T8tq7UGxGr5TcaEbjkqJnKI/MPrKzXzBdn3+b/XSMYONJT
# joJodXOLQhho3MZnVFCeTjzN7Kf5XpjZ15CAldwMbEgtaKqKxwkSstjkAYGqtuoK
# KBBwzfTuADU1rw9NPFBYie9yTl4FNroVTD7nU3V8h1DOXjCCjOkRWiURD7MC08Jl
# v2xj/rJ+IQ==
# SIG # End signature block
