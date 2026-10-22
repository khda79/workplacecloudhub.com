#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the retry backoff shared by the Intune export collectors: exponential waits with jitter that
honour Retry-After, a time budget for the exportJobs creation with the Graph SDK retry turned off, progress
lines while a job waits, a stalled job created again once, and explicit final errors.
.VERSION
1.0.0
.NOTES
Graph, the SDK request context, storage, the clock and the waits are recorded stand-ins inside the modules.
No tenant, Graph, storage, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic module test state only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Retry waits are recorded instead of slept.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.103' -Force -PassThru -Global
$evidence = Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.13' -Force -PassThru
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function New-HttpFailure([int]$Status, [string]$RetryAfter) {
    $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
    if ($RetryAfter) { $response.Headers.TryAddWithoutValidation('Retry-After', $RetryAfter) | Out-Null }
    [Microsoft.PowerShell.Commands.HttpResponseException]::new(("Response status code does not indicate success: {0} ({1})." -f $Status, $response.ReasonPhrase), $response)
}
function New-Record($Exception) { [Management.Automation.ErrorRecord]::new($Exception, 'Synthetic', 'NotSpecified', $null) }

$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-ExportBackoff-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
try {
    # ---- Core delay helper ----
    $d = Get-SmartM365RetryDelay -Attempt 1 -BaseSeconds 5 -JitterSample 0
    Check ($d.Seconds -eq 3 -and $d.Source -eq 'Backoff') "Attempt 1 lower bound is not half of 5s: $($d.Seconds)"
    Check ((Get-SmartM365RetryDelay -Attempt 1 -BaseSeconds 5 -JitterSample 1).Seconds -eq 5) 'Attempt 1 upper bound is not 5s.'
    Check ((Get-SmartM365RetryDelay -Attempt 4 -BaseSeconds 5 -JitterSample 0.5).Seconds -eq 30) 'Attempt 4 is not exponential (5*2^3 with jitter 0.5 = 30s).'
    Check ((Get-SmartM365RetryDelay -Attempt 12 -BaseSeconds 5 -MaximumSeconds 300 -JitterSample 1).Seconds -eq 300) 'The backoff is not capped.'
    $samples = @(1..60 | ForEach-Object { (Get-SmartM365RetryDelay -Attempt 3 -BaseSeconds 5).Seconds })
    Check (@($samples | Where-Object { $_ -lt 10 -or $_ -gt 20 }).Count -eq 0 -and @($samples | Sort-Object -Unique).Count -gt 1) "Jittered attempt 3 waits leave 10..20s or never vary: $($samples -join ',')"
    $d = Get-SmartM365RetryDelay -Attempt 1 -ErrorRecord (New-Record (New-HttpFailure 503 '17')) -JitterSample 0
    Check ($d.Seconds -eq 17 -and $d.Source -eq 'Retry-After') "Retry-After in seconds was not honoured: $($d.Seconds) $($d.Source)"
    $d = Get-SmartM365RetryDelay -Attempt 1 -ErrorRecord (New-Record (New-HttpFailure 429 ([DateTimeOffset]::UtcNow.AddSeconds(40).ToString('r'))))
    Check ($d.Seconds -ge 38 -and $d.Seconds -le 41 -and $d.Source -eq 'Retry-After') "Retry-After as an HTTP date was not honoured: $($d.Seconds)"
    Check ((Get-SmartM365RetryDelay -Attempt 1 -ErrorRecord (New-Record (New-HttpFailure 503 '900')) -MaximumSeconds 300).Seconds -eq 300) 'Retry-After is not capped.'
    $dataOnly = [Exception]::new('HTTP 503'); $dataOnly.Data['Retry-After'] = '9'
    Check ((Get-SmartM365RetryDelay -Attempt 5 -ErrorRecord (New-Record $dataOnly)).Seconds -eq 9) 'Retry-After carried in Exception.Data was not honoured.'

    # ---- Core export download: storage 503 without Retry-After, then the file ----
    $download = & $core {
        param($Answers, $OutFile)
        $script:BackoffAnswers = [Collections.Generic.Queue[object]]::new(); foreach ($a in $Answers) { $script:BackoffAnswers.Enqueue($a) }
        $script:BackoffSleeps = [Collections.Generic.List[int]]::new(); $script:BackoffLogs = [Collections.Generic.List[string]]::new()
        function script:Invoke-WebRequest { [CmdletBinding()] param($Method, $Uri, $OutFile, $TimeoutSec, $Headers) $answer = $script:BackoffAnswers.Dequeue(); if ($answer -is [Exception]) { throw $answer }; [IO.File]::WriteAllText($OutFile, 'zip') }
        function script:Start-Sleep { param([int]$Seconds) $script:BackoffSleeps.Add($Seconds) }
        function script:WriteLog { param($Message, $Level) $script:BackoffLogs.Add("$Level $Message") }
        try { $null = Invoke-SmartM365ExportFileDownload -Url 'https://intunestorage.example.invalid/r.zip?sig=SECRET' -OutFile $OutFile -Operation 'Intune export download (FirewallUnhealthyStatus)' }
        finally { Remove-Item Function:\Invoke-WebRequest, Function:\Start-Sleep, Function:\WriteLog -ErrorAction SilentlyContinue }
        [pscustomobject]@{ Saved = Test-Path -LiteralPath $OutFile; Sleeps = $script:BackoffSleeps.ToArray(); Logs = ($script:BackoffLogs -join "`n") }
    } @((New-HttpFailure 503), (New-HttpFailure 503), 'ok') (Join-Path $temporary 'download.zip')
    Check ($download.Saved -and $download.Sleeps.Count -eq 2 -and $download.Sleeps[0] -ge 5 -and $download.Sleeps[0] -le 10 -and $download.Sleeps[1] -ge 10 -and $download.Sleeps[1] -le 20) "Storage 503 without Retry-After did not back off exponentially: $($download.Sleeps -join ',')"
    Check ($download.Logs -match 'attempt 1/6 failed\. Status=503.*\(Backoff\)' -and $download.Logs -notmatch 'SECRET') 'The storage retry is not logged with its source, or the URL signature is.'

    # ---- EvidenceCollector stand-ins: fake clock, recorded waits/logs, SDK request context and Graph ----
    & $evidence {
        $script:FakeNow = [datetime]::SpecifyKind([datetime]'2026-10-22T01:00:00', 'Utc')
        $script:Sleeps = [Collections.Generic.List[int]]::new(); $script:Logs = [Collections.Generic.List[string]]::new()
        $script:Requests = [Collections.Generic.List[object]]::new(); $script:Post = [Collections.Generic.Queue[object]]::new(); $script:PostDefault = $null
        $script:JobStates = @{}; $script:Sdk = @{ MaxRetry = 3; History = [Collections.Generic.List[int]]::new() }
        function script:Get-SmartM365EvidenceUtcNow { $script:FakeNow }
        function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds); $script:FakeNow = $script:FakeNow.AddSeconds($Seconds) }
        function script:WriteLog { param($Message, $Level) $script:Logs.Add("$Level $Message") }
        function script:Get-MgRequestContext { [pscustomobject]@{ MaxRetry = $script:Sdk.MaxRetry } }
        function script:Set-MgRequestContext { param([int]$MaxRetry) $script:Sdk.MaxRetry = $MaxRetry; $script:Sdk.History.Add($MaxRetry) }
        function script:Invoke-MgGraphRequest {
            [CmdletBinding()] param($Method, $Uri, $Body, $ContentType, $OutputType)
            $script:Requests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; SdkMaxRetry = $script:Sdk.MaxRetry })
            $script:FakeNow = $script:FakeNow.AddSeconds(1)
            $answer = if ($Method -eq 'POST') { if ($script:Post.Count) { $script:Post.Dequeue() } else { $script:PostDefault } }
            else { $state = $script:JobStates[($Uri -split '/')[-1]]; if ($state -is [Collections.Generic.Queue[object]]) { if ($state.Count -gt 1) { $state.Dequeue() } else { $state.Peek() } } else { $state } }
            if ($answer -is [Exception]) { throw $answer }
            $answer
        }
        function script:Invoke-SmartM365ExportFileDownload {
            param($Url, $OutFile, $Operation)
            $folder = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-BackoffZip-' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $folder | Out-Null
            try { [IO.File]::WriteAllText((Join-Path $folder 'r.csv'), "DeviceId,DeviceName`r`nd1,PC1"); Compress-Archive -Path (Join-Path $folder 'r.csv') -DestinationPath $OutFile -Force } finally { Remove-Item -LiteralPath $folder -Recurse -Force }
            $OutFile
        }
        function script:Reset-Fake {
            param([object[]]$Post, $PostDefault, [hashtable]$JobStates)
            $script:Sleeps.Clear(); $script:Logs.Clear(); $script:Requests.Clear(); $script:Post.Clear(); foreach ($p in @($Post | Where-Object { $null -ne $_ })) { $script:Post.Enqueue($p) }
            $script:PostDefault = $PostDefault; $script:JobStates = if ($JobStates) { $JobStates } else { @{} }; $script:Sdk.MaxRetry = 3; $script:Sdk.History.Clear()
        }
        function script:Get-Fake { [pscustomobject]@{ Sleeps = $script:Sleeps.ToArray(); Logs = $script:Logs.ToArray(); Requests = $script:Requests.ToArray(); SdkMaxRetry = $script:Sdk.MaxRetry; SdkHistory = ($script:Sdk.History -join ',') } }
    }

    # 503 with Retry-After, then 503 without, then success, within the creation budget.
    $r = & $evidence {
        param($first, $second)
        Reset-Fake -Post @($first, $second, [pscustomobject]@{ id = 'job-1' })
        $value = Invoke-SmartM365EvidenceGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs' -Body '{}' -RetryBudgetSeconds 1080 -Operation 'Intune export job creation (FirewallStatus)'
        [pscustomobject]@{ Value = $value; Fake = Get-Fake }
    } (New-HttpFailure 503 '4') (New-HttpFailure 503)
    Check ($r.Value.id -eq 'job-1' -and $r.Fake.Requests.Count -eq 3) '503 then success did not return the job after three requests.'
    Check ($r.Fake.Sleeps[0] -eq 4 -and $r.Fake.Sleeps[1] -ge 5 -and $r.Fake.Sleeps[1] -le 10) "Waits ignore Retry-After or the backoff: $($r.Fake.Sleeps -join ',')"
    $retries = @($r.Fake.Logs | Where-Object { $_ -match ' returned HTTP 503 on attempt ' })
    Check ($retries.Count -eq 2 -and $retries[0] -match '^INFO Intune export job creation \(FirewallStatus\) returned HTTP 503 on attempt 1; retrying in 4s \(Retry-After; elapsed 1s, budget 1080s\)' -and $retries[1] -match '^INFO .*attempt 2; retrying in \d+s \(Backoff') "Retries are not logged with operation, wait and source: $($retries -join ' | ')"
    Check (@($r.Fake.Logs | Where-Object { $_ -like 'WARNING *' -or $_ -like 'ERROR *' }).Count -eq 0) 'A retry that recovered was logged as a warning (the run would end CompletedWithWarnings).'

    # The creation budget runs out: one explicit error with operation, attempts, elapsed time and status.
    $r = & $evidence {
        param($failure)
        Reset-Fake -PostDefault $failure
        $result = [pscustomobject]@{ Error = $null; Fake = $null }
        try { Invoke-SmartM365EvidenceGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs' -Body '{}' -RetryBudgetSeconds 600 -Operation 'Intune export job creation (FirewallStatus)' | Out-Null }
        catch { $result.Error = $_ }
        $result.Fake = Get-Fake
        $result
    } (New-HttpFailure 503)
    $match = [regex]::Match([string]$r.Error.Exception.Message, 'Intune export job creation \(FirewallStatus\) failed after (\d+) attempt\(s\) over (\d+)s \(HTTP 503\)\. Last error: Response status code does not indicate success: 503')
    Check ($match.Success) "Budget exhaustion is not reported explicitly: $($r.Error.Exception.Message)"
    Check ([int]$match.Groups[1].Value -eq $r.Fake.Requests.Count -and [int]$match.Groups[1].Value -ge 6 -and [int]$match.Groups[2].Value -le 600 -and ($r.Fake.Sleeps | Measure-Object -Sum).Sum -le 600) "The budget is not respected: $($r.Error.Exception.Message); waits $($r.Fake.Sleeps -join ',')"
    Check ((Get-SmartM365EvidenceStatusCode $r.Error) -eq 503 -and $r.Error.Exception.Data['Attempts'] -eq $r.Fake.Requests.Count) 'The final error lost its HTTP status or attempt count.'

    # Budgeted export: SDK retry off during the calls and restored; a stuck job logs progress, then times out.
    $r = & $evidence {
        param($failure)
        Reset-Fake -Post @($failure, [pscustomobject]@{ id = 'job-1' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'inProgress' } }
        $result = [pscustomobject]@{ Error = $null; Fake = $null }
        try { Invoke-SmartM365EvidenceExportReport -ReportName 'DefenderAgents' -TimeoutSeconds 300 -PollSeconds 30 -CreateRetryBudgetSeconds 1080 -ProgressLogSeconds 90 | Out-Null }
        catch { $result.Error = $_ }
        $result.Fake = Get-Fake
        $result
    } (New-HttpFailure 503 '2')
    Check ($null -ne $r.Error -and $r.Error.Exception.Message -match "^Intune export job job-1 for DefenderAgents timed out after \d+s \(limit 300s\); last status 'inProgress'\.$") "A stuck job does not time out with its report, id and status: $($r.Error)"
    $progress = @($r.Fake.Logs | Where-Object { $_ -match "^INFO Intune export job job-1 for DefenderAgents is 'inProgress' after \d+s \(report elapsed \d+s of 300s\)\." })
    Check ($progress.Count -ge 3) "Progress was not logged regularly while the job waited: $($r.Fake.Logs -join ' | ')"
    Check (@($r.Fake.Requests | Where-Object SdkMaxRetry -ne 0).Count -eq 0 -and $r.Fake.SdkMaxRetry -eq 3 -and $r.Fake.SdkHistory -eq '0,3') "The SDK retry was not off during the export or not restored: $($r.Fake.SdkHistory)"
    Check ($r.Fake.Sleeps[0] -eq 2) 'The creation 503 Retry-After was not honoured inside the export.'

    # A job still inProgress past StallSeconds is abandoned and created again once; the new job completes.
    $second = [Collections.Generic.Queue[object]]::new(); $second.Enqueue([pscustomobject]@{ status = 'inProgress' }); $second.Enqueue([pscustomobject]@{ status = 'completed'; url = 'https://intunestorage.example.invalid/r.zip?sig=S' })
    $r = & $evidence {
        param($second)
        Reset-Fake -Post @([pscustomobject]@{ id = 'job-1' }, [pscustomobject]@{ id = 'job-2' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'inProgress' }; 'job-2' = $second }
        $rows = @(Invoke-SmartM365EvidenceExportReport -ReportName 'DefenderAgents' -TimeoutSeconds 1200 -PollSeconds 60 -CreateRetryBudgetSeconds 1080 -StallSeconds 600)
        [pscustomobject]@{ Rows = $rows; Fake = Get-Fake }
    } $second
    Check ($r.Rows.Count -eq 1 -and $r.Rows[0].DeviceId -eq 'd1') 'The recreated export did not return its rows.'
    Check (@($r.Fake.Requests | Where-Object Method -eq 'POST').Count -eq 2) 'The stalled job was not created again exactly once.'
    Check (@($r.Fake.Logs | Where-Object { $_ -match "^WARNING Intune export job job-1 for DefenderAgents is still 'inProgress' after \d+s; abandoning it" }).Count -eq 1) 'The stalled job abandonment is not logged.'

    # A second stall is not recreated again: the timeout names the job and says it was already recreated.
    $r = & $evidence {
        Reset-Fake -Post @([pscustomobject]@{ id = 'job-1' }, [pscustomobject]@{ id = 'job-2' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'notStarted' }; 'job-2' = [pscustomobject]@{ status = 'notStarted' } }
        $result = [pscustomobject]@{ Error = $null; Fake = $null }
        try { Invoke-SmartM365EvidenceExportReport -ReportName 'DefenderAgents' -TimeoutSeconds 900 -PollSeconds 60 -CreateRetryBudgetSeconds 1080 -StallSeconds 300 | Out-Null } catch { $result.Error = $_ }
        $result.Fake = Get-Fake
        $result
    }
    Check (@($r.Fake.Requests | Where-Object Method -eq 'POST').Count -eq 2 -and $r.Error.Exception.Message -match "^Intune export job job-2 for DefenderAgents timed out after \d+s \(limit 900s\); last status 'notStarted'\. The job had already been created again once\.$") "A second stall was recreated or not reported: $($r.Error)"

    # Without a creation budget the previous behaviour stays: SDK retry untouched, attempt-limited retries.
    $r = & $evidence {
        Reset-Fake -Post @([pscustomobject]@{ id = 'job-1' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'completed'; url = 'https://intunestorage.example.invalid/r.zip?sig=S' } }
        $rows = @(Invoke-SmartM365EvidenceExportReport -ReportName 'AppInvRawData' -PollSeconds 1)
        [pscustomobject]@{ Rows = $rows; Fake = Get-Fake }
    }
    Check ($r.Rows.Count -eq 1 -and $r.Fake.SdkHistory -eq '' -and @($r.Fake.Requests | Where-Object SdkMaxRetry -ne 3).Count -eq 0) 'An export without creation budget changed the SDK retry.'
}
finally {
    Remove-Module $evidence -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks export retry backoff checks; PowerShell $($PSVersionTable.PSVersion)."
