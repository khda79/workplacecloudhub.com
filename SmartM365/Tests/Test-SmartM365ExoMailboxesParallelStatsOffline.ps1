#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of the parallel live statistics prefetch of the EXO mailboxes producer: parallel values
equal the sequential ones, partial errors, simulated throttling, worker connection failure, deadline,
one persistent session per worker, and unchanged Fast and permissions-only paths.
.VERSION
1.0.0
.NOTES
Synthetic mailboxes and mocked Exchange cmdlets only. No tenant, Exchange, API, mail or synchronized data access.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$producerPath = Join-Path $root 'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1'
$evidencePath = Join-Path $root 'SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$text = [IO.File]::ReadAllText($producerPath)
$ast = [Management.Automation.Language.Parser]::ParseFile($producerPath, [ref]$null, [ref]$null)
$evidenceAst = [Management.Automation.Language.Parser]::ParseFile($evidencePath, [ref]$null, [ref]$null)
function Get-FunctionText($Tree, [string]$Name) { $Tree.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true).Extent.Text }

# ---------- Static wiring ----------
Check ($text -match '\.VERSION\r?\n1\.29\r?\n' -and $text -match '\$ScriptVersion = "1\.29"') 'EXO mailboxes version is not 1.29.'
Check ($text -match "Get-ScriptLocalConfigValue -Config \`$ScriptLocalConfig -Name 'StatsParallelThrottle' -DefaultValue 4") 'StatsParallelThrottle does not default to 4 from the local configuration.'
$template = Get-Content -LiteralPath (Join-Path (Split-Path $producerPath) 'SmartM365-EXO-Mailboxes-Inventory.local.json.template') -Raw | ConvertFrom-Json
Check ($template.StatsParallelThrottle -eq 4) 'The local configuration template does not expose StatsParallelThrottle = 4.'
Check ($text -match 'if \(\$UseLiveStats -and \$StatsParallelThrottleEffective -gt 1 -and \(-not \$InteractiveAuth\)') 'The prefetch is not limited to live stats with a throttle above 1.'
Check ($text -match '\$UseLiveStats = \(\$IncludeStats -and \(-not \$ExcludeStats\) -and \(-not \$PermissionsOnly\)\)') 'Live stats can now run with Fast or permissions-only.'
Check ($text -match '\$prefetchedStats\.ContainsKey\(\$counter - 1\)' -and $text -match '\$statsIndex = -1\r?\n\s*foreach \(\$statsMailbox in \$mailboxes\) \{\r?\n\s*\$statsIndex\+\+') 'Prefetched values are not keyed by the sequential loop position.'
Check ($text -match 'if \(\$null -eq \$liveStats\) \{ \$liveStats = Get-MailboxLiveStatsRecord -Mailbox \$mb') 'A mailbox missing from the prefetch is no longer collected sequentially.'
Check ($text -match "catch \[System\.Management\.Automation\.PipelineStoppedException\] \{ throw \}") 'Cancellation during the prefetch is swallowed.'
Check ($text -match '(?s)Live stats prefetch completed.*?try \{ Invoke-EXOProactiveReconnect \}') 'The main session is not refreshed after the worker sessions.'
$prefetchText = Get-FunctionText $ast 'Invoke-MailboxStatsParallelPrefetch'
Check ($prefetchText -notmatch 'Disconnect-ExchangeOnline') 'A stats worker disconnects and would clear sibling EXO connections.'
Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-ExoStatsExact' }, $true)).Count -eq 2) 'Live stats are collected outside Get-MailboxLiveStatsRecord.'
Check ($text -match 'workers=\{0\}, dispatched=\{1\}, prefetched=\{2\}.*seconds=\{8\}, throughput=\{9\} mailboxes/min') 'Worker count, duration and throughput are not logged.'

# ---------- Runtime with mocked Exchange cmdlets ----------
$state = [System.Collections.Concurrent.ConcurrentDictionary[string,int]]::new()
[AppDomain]::CurrentDomain.SetData('SmartM365ParallelStatsTest', $state)
$mocks = @'
function Get-SyntheticNumber([string]$Identity) { [int]([regex]::Match($Identity, '\d+').Value) }
function New-SyntheticSize([double]$Bytes) {
    $value = [pscustomobject]@{ Bytes = $Bytes }
    $value | Add-Member -MemberType ScriptMethod -Name ToBytes -Value { [double]$this.Bytes }
    [pscustomobject]@{ Value = $value }
}
function Test-SyntheticThrottled([string]$Identity) {
    $state = [AppDomain]::CurrentDomain.GetData('SmartM365ParallelStatsTest')
    $calls = $state.AddOrUpdate("calls|$Identity", 1, { param($k, $v) $v + 1 })
    if ($Identity -like '*always-empty*') { return $true }
    $limit = 0
    if ($state.TryGetValue("throttle|$Identity", [ref]$limit)) { return ($calls -le $limit) }
    return $false
}
function Get-EXOMailboxStatistics {
    [CmdletBinding()]
    param([string]$Identity, [switch]$Archive, [string[]]$Properties)
    if ($Identity -like 'broken*') { throw "Synthetic terminating failure for $Identity" }
    if (Test-SyntheticThrottled $(if ($Archive) { "archive|$Identity" } else { $Identity })) { return $null }
    $n = Get-SyntheticNumber $Identity
    $result = [pscustomobject]@{ TotalItemSize = (New-SyntheticSize (($n + 1) * 123456789.0 * $(if ($Archive) { 3 } else { 1 }))); ItemCount = $n * 10 + $(if ($Archive) { 1 } else { 0 }) }
    if ($n % 5 -ne 0) { $result | Add-Member -NotePropertyName LastLogonTime -NotePropertyValue ([datetime]::new(2026, 1, 1, 8, 0, 0, [DateTimeKind]::Utc).AddMinutes($n)) }
    if ($Properties -contains 'LastUserActionTime') { $result | Add-Member -NotePropertyName LastUserActionTime -NotePropertyValue ([datetime]::new(2026, 2, 1, 8, 0, 0, [DateTimeKind]::Utc).AddMinutes($n)) }
    $result
}
function Get-MailboxStatistics {
    [CmdletBinding()]
    param([string]$Identity, [switch]$Archive)
    if (Test-SyntheticThrottled ("legacy|" + $Identity)) { return $null }
    $n = Get-SyntheticNumber $Identity
    [pscustomobject]@{ LastLogonTime = [datetime]::new(2025, 6, 1, 0, 0, 0, [DateTimeKind]::Utc).AddHours($n); ItemCount = $n; TotalItemSize = '1.5 GB (1,610,612,736 bytes)' }
}
'@
$definitions = (@(
        (Get-FunctionText $ast 'Convert-QuotaToGB'),
        (Get-FunctionText $evidenceAst 'ConvertTo-IdentityNativeIso'),
        (Get-FunctionText $ast 'Get-ExoStatsExact'),
        (Get-FunctionText $ast 'Get-MailboxLiveStatsRecord')
    ) -join "`n") + "`n" + $mocks
. ([scriptblock]::Create($definitions))
. ([scriptblock]::Create($prefetchText))
$connect = @'
param($Label)
$state = [AppDomain]::CurrentDomain.GetData('SmartM365ParallelStatsTest')
[void]$state.AddOrUpdate('connect', 1, { param($k, $v) $v + 1 })
if ($Label -eq 'fail') { throw 'Synthetic EXO connection refusal' }
'@
function New-Items([string[]]$Identities) {
    $index = -1
    foreach ($identity in $Identities) {
        $index++
        $status = if ($index % 3 -eq 0) { 'Active' } else { 'None' }
        [pscustomobject]@{ Index = $index; Mailbox = [pscustomobject]@{ UserPrincipalName = $identity; ArchiveStatus = $status } }
    }
}
function Get-ValueJson($Values) { $Values | Select-Object ArchiveSizeGB, ArchiveCount, ArchiveLastLogon, ArchiveLastLogonIso, TotalSizeGB, ItemCount, LastLogon, LastLogonIso, LastAction, LastActionIso | ConvertTo-Json -Compress }
function Invoke-Prefetch($Items, [int]$Throttle = 4, [string]$Label = 'ok', [datetime]$Deadline = [datetime]::MaxValue, [int]$EmptyRetries = 2, [switch]$Luat) {
    Invoke-MailboxStatsParallelPrefetch -Mailboxes $Items -ThrottleLimit $Throttle -WorkerDefinitions $definitions -ConnectScript $connect -ConnectArguments @($Label) `
        -ConnectRetries 1 -IncludeLastUserActionTime:$Luat -EmptyResultRetries $EmptyRetries -RetryDelaySeconds 0 -DeadlineUtc $Deadline -ProgressInterval 10 `
        -OnProgress { param($done, $all, $elapsed) $script:progress.Add($done) } -OnWorkerFailure { param($id, $message) $script:workerFailures.Add("$id|$message") }
}
function Reset-State { $state.Clear(); $script:progress = [Collections.Generic.List[int]]::new(); $script:workerFailures = [Collections.Generic.List[string]]::new() }

# Parallel values equal the sequential values, keyed by loop position, with one session per worker.
$identities = @(1..41 | ForEach-Object { 'user{0:000}@contoso.test' -f $_ })
foreach ($luat in @($false, $true)) {
    Reset-State
    $items = @(New-Items $identities)
    $sequential = @{}
    foreach ($item in $items) { $sequential[$item.Index] = Get-ValueJson (Get-MailboxLiveStatsRecord -Mailbox $item.Mailbox -IncludeLastUserActionTime:$luat) }
    Reset-State
    $summary = Invoke-Prefetch $items -Luat:$luat
    Check ($summary.WorkerCount -eq 4 -and $summary.Dispatched -eq 41 -and $summary.Prefetched -eq 41 -and $summary.Results.Count -eq 41) "Parallel prefetch did not return every mailbox (LUAT=$luat)."
    foreach ($item in $items) {
        Check ((Get-ValueJson $summary.Results[$item.Index]) -ceq $sequential[$item.Index]) "Parallel values differ from sequential values for position $($item.Index) (LUAT=$luat)."
    }
    $connects = 0; [void]$state.TryGetValue('connect', [ref]$connects)
    Check ($connects -eq 4) 'Workers did not open exactly one persistent session each.'
    Check (($summary.Retried + $summary.Failed + $summary.EmptyAfterRetries + $summary.ConnectFailedWorkers + $summary.DeadlineSkipped) -eq 0) 'A clean run reported retries or failures.'
    Check ((@($script:progress) -join ',') -eq '10,20,30,40') 'Prefetch progress was not reported at the interval.'
}
$sample = $summary.Results[0]
Check ($sample.TotalSizeGB -eq 0.23 -and $sample.ArchiveCount -eq 11 -and $sample.LastLogonIso -like '2026-01-01T08:01:00*' -and $sample.LastActionIso -like '2026-02-01T08:01:00*') 'Synthetic values were not converted as expected.'
Check ($summary.Results[4].LastLogonIso -like '2025-06-01T05:00:00*') 'The legacy LastLogonTime fallback was not used.'

# A single worker request (throttle 2 over one mailbox) still uses one worker.
Reset-State
$one = Invoke-Prefetch @(New-Items @('user007@contoso.test')) -Throttle 2
Check ($one.WorkerCount -eq 1 -and $one.Results.Count -eq 1) 'One mailbox did not use a single worker.'

# Partial errors and simulated throttling: retried values are kept, failures are left to the sequential loop.
Reset-State
[void]$state.TryAdd('throttle|user002@contoso.test', 2)
[void]$state.TryAdd('throttle|legacy|user002@contoso.test', 2)
$mixed = @(New-Items @('user001@contoso.test', 'user002@contoso.test', 'broken003@contoso.test', 'always-empty004@contoso.test', 'user005@contoso.test'))
$summary = Invoke-Prefetch $mixed -Throttle 2
Check ($summary.Results.ContainsKey(1) -and $summary.Retried -ge 1) 'A throttled mailbox was not retried into a complete value.'
Check (-not $summary.Results.ContainsKey(2) -and $summary.Failed -eq 1) 'A failing mailbox was not left to the sequential loop.'
Check (-not $summary.Results.ContainsKey(3) -and $summary.EmptyAfterRetries -eq 1) 'An always-empty mailbox was not left to the sequential loop after its retries.'
Check ($summary.Results.ContainsKey(0) -and $summary.Results.ContainsKey(4) -and $summary.Prefetched -eq 3) 'Healthy mailboxes were lost next to failures.'
$calls = 0; [void]$state.TryGetValue('calls|always-empty004@contoso.test', [ref]$calls)
Check ($calls -eq 3) 'The empty-result retry count is not EmptyResultRetries + 1 primary calls.'

# Worker connection failure: no values, one failure callback per worker, every mailbox stays sequential.
Reset-State
$summary = Invoke-Prefetch @(New-Items $identities) -Label 'fail'
Check ($summary.ConnectFailedWorkers -eq 4 -and $summary.Results.Count -eq 0 -and $script:workerFailures.Count -eq 4) 'Failed worker sessions were not reported or produced values.'
Check (@($script:workerFailures | Where-Object { $_ -match 'Synthetic EXO connection refusal' }).Count -eq 4) 'The worker connection error was not reported.'

# Deadline (MaxRunMinutes) reached: workers stop before calling Exchange.
Reset-State
$summary = Invoke-Prefetch @(New-Items $identities) -Deadline ([datetime]::UtcNow.AddMinutes(-1))
$anyCall = @($state.Keys | Where-Object { $_ -like 'calls|*' }).Count
Check ($summary.DeadlineSkipped -eq 41 -and $summary.Results.Count -eq 0 -and $anyCall -eq 0) 'Workers kept calling Exchange after the deadline.'

# Empty input returns an empty summary without workers.
$empty = Invoke-MailboxStatsParallelPrefetch -Mailboxes @() -ThrottleLimit 4 -WorkerDefinitions $definitions -ConnectScript $connect
Check ($empty.WorkerCount -eq 0 -and $empty.Results.Count -eq 0) 'An empty prefetch started workers.'

[AppDomain]::CurrentDomain.SetData('SmartM365ParallelStatsTest', $null)
Write-Host ("PASS: {0} parallel live stats offline checks. Synthetic mailboxes only; no tenant, Exchange or publication." -f $script:checks)
