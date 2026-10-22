#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the retry backoff shared by the Intune export collectors: exponential waits with jitter that
honour Retry-After, a time budget for the exportJobs creation that never changes the Graph SDK request context
(turned off at connection only), progress lines while a job waits, a stalled job created again once, and explicit final errors.
.VERSION
1.0.1
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
$evidence = Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.14' -Force -PassThru
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
            param([object[]]$Post, $PostDefault, [hashtable]$JobStates, [int]$SdkMaxRetry = 3)
            $script:Sleeps.Clear(); $script:Logs.Clear(); $script:Requests.Clear(); $script:Post.Clear(); foreach ($p in @($Post | Where-Object { $null -ne $_ })) { $script:Post.Enqueue($p) }
            $script:PostDefault = $PostDefault; $script:JobStates = if ($JobStates) { $JobStates } else { @{} }; $script:Sdk.MaxRetry = $SdkMaxRetry; $script:Sdk.History.Clear()
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

    # Budgeted export after a -DisableSdkRetry connection: the request context is never changed (SDK 2.30 rebuilds
    # its client from an emptied environment and app-only auth fails with Invalid URI); a stuck job logs progress.
    $r = & $evidence {
        param($failure)
        Reset-Fake -Post @($failure, [pscustomobject]@{ id = 'job-1' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'inProgress' } } -SdkMaxRetry 0
        $result = [pscustomobject]@{ Error = $null; Fake = $null }
        try { Invoke-SmartM365EvidenceExportReport -ReportName 'DefenderAgents' -TimeoutSeconds 300 -PollSeconds 30 -CreateRetryBudgetSeconds 1080 -ProgressLogSeconds 90 | Out-Null }
        catch { $result.Error = $_ }
        $result.Fake = Get-Fake
        $result
    } (New-HttpFailure 503 '2')
    Check ($null -ne $r.Error -and $r.Error.Exception.Message -match "^Intune export job job-1 for DefenderAgents timed out after \d+s \(limit 300s\); last status 'inProgress'\.$") "A stuck job does not time out with its report, id and status: $($r.Error)"
    $progress = @($r.Fake.Logs | Where-Object { $_ -match "^INFO Intune export job job-1 for DefenderAgents is 'inProgress' after \d+s \(report elapsed \d+s of 300s\)\." })
    Check ($progress.Count -ge 3) "Progress was not logged regularly while the job waited: $($r.Fake.Logs -join ' | ')"
    Check (@($r.Fake.Requests | Where-Object SdkMaxRetry -ne 0).Count -eq 0 -and $r.Fake.SdkHistory -eq '') "The export changed the SDK request context: $($r.Fake.SdkHistory)"
    Check (@($r.Fake.Logs | Where-Object { $_ -match 'Graph SDK retry on' }).Count -eq 0) 'An export with the SDK retry already off warned about it.'
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

    # A budgeted export on a connection without -DisableSdkRetry warns once and still leaves the context alone.
    $r = & $evidence {
        Reset-Fake -Post @([pscustomobject]@{ id = 'job-1' }) -JobStates @{ 'job-1' = [pscustomobject]@{ status = 'completed'; url = 'https://intunestorage.example.invalid/r.zip?sig=S' } }
        $rows = @(Invoke-SmartM365EvidenceExportReport -ReportName 'FirewallStatus' -PollSeconds 1 -CreateRetryBudgetSeconds 1080)
        [pscustomobject]@{ Rows = $rows; Fake = Get-Fake }
    }
    Check ($r.Rows.Count -eq 1 -and $r.Fake.SdkHistory -eq '' -and @($r.Fake.Logs | Where-Object { $_ -match '^WARNING Intune export job creation \(FirewallStatus\) runs with the Graph SDK retry on \(MaxRetry 3\)' }).Count -eq 1) "A budgeted export with the SDK retry on did not warn once or changed the context: $($r.Fake.SdkHistory)"

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtaW1mrKnIB/9P
# RnMWd3t2y00EkCEQWviq4qmm27vxzqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID1tRZdenAZG145KExaMPk9309wonc462saSRYtCxlD+MA0GCSqG
# SIb3DQEBAQUABIIBgA4oxbUTyBf2zw/eE+zI5QJLt6n7M9LsXhMzloCrnysgxGtr
# A8zSmW8DJCsnFrSlSCQ0SgF3pZijZPVfmoAyQ2HKrPHYSqNCzXeb0/bWbdjB6yCW
# pZMLyNRfthMhqRXtuwiQggnQlEz+PKnxcONyY/LmKU7LJ1igRYqLM6uj52BYINtC
# xuX/rnLdeKb6kr//6gh7PMUdPFX28CdUUaXr23ZJGNCcxh0OLVLr3igTvWL/7cwU
# v+5t2ux6Rpc4LjFRIZHLBxowWI/GrlnU7pMA5HrEh6NBW27hrf5q4vJg5+TWU4Z0
# ByR+3vMGt3r4PCsBwRO845G/mPldpmfdtv22Vd9CxMu8BDkDNQo4HHZCaDhllLrb
# wgBq9iiwYjm7EYJa+zAdBK6XcrQCsCTm11oQPzku4mdM14OspA8NefIDllEy3IEP
# BvUOB+JqZgJDU3XBAla7h3l6I54OyNHQP7/a8x4odECY/EmjeI2SrAVsw5Ms3qb3
# PwP2PhgVA9uc4AkN0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNzI5
# NTlaMC8GCSqGSIb3DQEJBDEiBCAx7RAECdO+RfxkADwoyRujSPUvrIdyaAvuLavw
# japTPDANBgkqhkiG9w0BAQEFAASCAgAPqKILuQdtEPvRBX334zFUsiC/aVwCpSTd
# ne/0sTXYbWTYKpOb4LasD09zGO03qYN1RkcXXrQjhJyFKgyNVdu69eZhgSXuTPet
# fQnwDrxMZdl2IeYEl32rz6J4RQjls1Kacb/GgDq+y22voTm4evCYh3VKPNovZdoA
# 70ZWVrF/QOBxRVySDJhiobHuAmCGNxJ1gtvTWtoV+fNlF1yowzs9/HNo/TRZI7Zq
# rSObSwLoxdbjb2Ibinu3ocJbJt889JurVoe0ejjb47wqIwuZaJa62Si9F00Xim9O
# l4zPz7bOhOfJn8dUxcZndrRqph+H11FcOhiM4fH2uI0BqbS1JVJekh5IBUs58igb
# eLhgY7LMK1O/QGK9rtogBZF2zhE5rmrBsdsGok15PJUjAOPKaDkPJbfAv6AJdGh4
# veh32SrMFpYCrXipnaKsyhTqpCBoQh7AmSleNCoF95R+UlsTTAAHoTv48mmElT7w
# p/3rt8nD5fwUe7bFZUKP1RW+JL9mGLio3hQggOsjxOYhsThSnpoG60m/M2oM9KMb
# ftrZT8sJndA/CKPTfDsOt2a1dX7LTZwRAH6mlvMSsf3TF0O1kdTx1AdIl8KDWI/4
# qs4lA2KpLDZ1maTGzYbn07g8uFUXVPy7DypLklqFWKztoL8655uHrYIPYysMaA/p
# GU7mIb6cqg==
# SIG # End signature block
