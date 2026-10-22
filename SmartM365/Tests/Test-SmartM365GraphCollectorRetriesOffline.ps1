#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the collector Graph retries and of the Intune export download: 401/403 fail at once,
408/429/5xx, timeouts and connection resets are retried, and a pre-signed export URL is downloaded
without an Authorization header, its real error logged without the URL signature.
.VERSION
1.0.2
.NOTES
Collector functions are loaded from the AST with recorded stand-ins for Graph and storage. No tenant,
Graph, storage, mail or synchronized data access.
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
# Endpoint Analytics and Autopatch import Core with -Prefix Core: their functions call the prefixed names.
Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.103' -Prefix Core -Global -WarningAction SilentlyContinue
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function New-HttpFailure([int]$Status, [string]$RetryAfter) {
    $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
    if ($RetryAfter) { $response.Headers.Add('Retry-After', $RetryAfter) }
    [Microsoft.PowerShell.Commands.HttpResponseException]::new(("Response status code does not indicate success: {0} ({1})." -f $Status, $response.ReasonPhrase), $response)
}
function New-Timeout { [System.Threading.Tasks.TaskCanceledException]::new('The request was canceled due to the configured HttpClient.Timeout of 100 seconds elapsing.') }
function New-Record($Exception) { [Management.Automation.ErrorRecord]::new($Exception, 'Synthetic', 'NotSpecified', $null) }
function Import-CollectorFunction([string]$RelativePath, [string[]]$Names, [string]$Prelude) {
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root $RelativePath), [ref]$null, [ref]$null)
    $text = foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) { throw "$name is missing in $RelativePath." }
        $node.Extent.Text
    }
    New-Module -ScriptBlock ([scriptblock]::Create($Prelude + "`n" + ($text -join "`n")))
}
# Common stand-ins: an answer queue (exception to throw or value to return) and recorded waits.
$standIns = @'
$script:Answers = [Collections.Generic.Queue[object]]::new()
$script:Calls = 0
$script:Slept = [Collections.Generic.List[int]]::new()
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $script:Slept.Add($Seconds) }
function WriteLog { param($Message, $Level) $null = $Message, $Level }
function Write-Log { param($Message, $Level, $Category) $null = $Message, $Level, $Category }
function Write-Warning { param($Message) $null = $Message }
function Next-Answer { $script:Calls++; $answer = $script:Answers.Dequeue(); if ($answer -is [Exception]) { throw $answer }; $answer }
function Set-Answers([object[]]$Items) { $script:Answers.Clear(); $script:Calls = 0; $script:Slept.Clear(); foreach ($item in $Items) { $script:Answers.Enqueue($item) } }
'@

$temporary = $null
try {
    # ---- Core classification and export download ----
    $classify = { param($e, [switch]$Conflict) (Test-SmartM365GraphRetryableError -ErrorRecord (New-Record $e) -RetryConflict:$Conflict) }
    Check ((& $classify (New-HttpFailure 500)).Retry -and (& $classify (New-HttpFailure 503)).StatusCode -eq 503) 'A 5xx is not retryable.'
    Check (-not (& $classify (New-HttpFailure 403)).Retry -and -not (& $classify (New-HttpFailure 401)).Retry) 'A 401/403 is retryable.'
    Check ((& $classify (New-Timeout)).Retry -and (& $classify ([Net.Http.HttpRequestException]::new('An error occurred while sending the request.'))).Retry) 'A timeout or connection error is not retryable.'
    Check ((& $classify ([Exception]::new('InternalServerError (Internal Server Error)'))).Retry -and (& $classify ([Exception]::new('Graph returned HTTP 503 for the page'))).StatusCode -eq 503) 'A message-only Graph SDK status is not recognized.'
    Check (-not (& $classify ([Exception]::new('Forbidden (Forbidden): gateway timed out'))).Retry) 'A 403 mentioning a timeout is retried.'
    Check (-not (& $classify (New-HttpFailure 409)).Retry -and (& $classify (New-HttpFailure 409) -Conflict).Retry) 'A 409 retry is not opt-in.'
    Check ($null -eq (Get-SmartM365GraphErrorStatusCode -ErrorRecord (New-Record ([Exception]::new('Failed on https://graph.microsoft.com/v1.0/reports/404/rows'))))) 'A number in a URL is read as a status.'

    $temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-GraphRetries-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $temporary | Out-Null
    & $core {
        $script:DownloadAnswers = [Collections.Generic.Queue[object]]::new()
        $script:DownloadCalls = [Collections.Generic.List[object]]::new()
        $script:DownloadLogs = [Collections.Generic.List[string]]::new()
        $script:DownloadSleeps = [Collections.Generic.List[int]]::new()
        function script:Invoke-WebRequest {
            [CmdletBinding()]
            param($Method, $Uri, $OutFile, $TimeoutSec, $Headers, [switch]$UseBasicParsing)
            $script:DownloadCalls.Add([pscustomobject]@{ HasHeaders = $PSBoundParameters.ContainsKey('Headers'); TimeoutSec = $TimeoutSec })
            $answer = $script:DownloadAnswers.Dequeue()
            if ($answer -is [Exception]) { throw $answer }
            [IO.File]::WriteAllText($OutFile, 'zip-bytes')
        }
        function script:Start-Sleep { param([int]$Seconds) $script:DownloadSleeps.Add($Seconds) }
        function script:WriteLog { param($Message, $Level) $script:DownloadLogs.Add("$Level $Message") }
    }
    $sasUrl = 'https://intunestorage.example.invalid/exports/report.zip?sv=2024&sig=SECRETSIGNATURE'
    function Invoke-Download([object[]]$Answers) {
        & $core {
            param($Answers, $Url, $OutFile)
            $script:DownloadAnswers.Clear(); $script:DownloadCalls.Clear(); $script:DownloadLogs.Clear(); $script:DownloadSleeps.Clear()
            foreach ($answer in $Answers) { $script:DownloadAnswers.Enqueue($answer) }
            $result = [pscustomobject]@{ Error = ''; Saved = $false }
            try { $null = Invoke-SmartM365ExportFileDownload -Url $Url -OutFile $OutFile; $result.Saved = Test-Path -LiteralPath $OutFile }
            catch { $result.Error = $_.Exception.Message }
            [pscustomobject]@{ Error = $result.Error; Saved = $result.Saved; Calls = $script:DownloadCalls.ToArray(); Logs = ($script:DownloadLogs -join "`n"); Sleeps = ($script:DownloadSleeps -join ',') }
        } $Answers $sasUrl (Join-Path $temporary ('export-{0}.zip' -f [guid]::NewGuid().ToString('N')))
    }
    # Blob not yet available right after "completed", then the file.
    $r = Invoke-Download @((New-HttpFailure 404), (New-HttpFailure 403), 'ok')
    Check ($r.Saved -and @($r.Calls).Count -eq 3 -and $r.Sleeps -eq '5,10') "A not-yet-available export was not retried briefly: $($r.Error)"
    Check (@($r.Calls | Where-Object HasHeaders).Count -eq 0) 'The pre-signed export URL was sent an Authorization header.'
    Check ($r.Logs -match 'Status=404' -and $r.Logs -match 'Status=403' -and $r.Logs -notmatch 'SECRETSIGNATURE') 'The real download error is not logged, or the URL signature is.'
    # A 403 that persists is reported with its real status, without the signature, after the short grace.
    $r = Invoke-Download @((New-HttpFailure 403), (New-HttpFailure 403), (New-HttpFailure 403), (New-HttpFailure 403))
    Check (-not $r.Saved -and @($r.Calls).Count -eq 4 -and $r.Error -match 'failed after 4 attempt\(s\)\. Status=403' -and $r.Error -notmatch 'SECRETSIGNATURE') "A persistent 403 is not reported clearly: $($r.Error)"
    # Transient storage errors and timeouts are retried with progressive waits; Retry-After is honoured.
    $r = Invoke-Download @((New-HttpFailure 503 '7'), (New-Timeout), 'ok')
    # Second wait: exponential backoff 10*2^1 with equal jitter, so 10 to 20 seconds.
    $waits = @($r.Sleeps -split ',' | ForEach-Object { [int]$_ })
    Check ($r.Saved -and $waits.Count -eq 2 -and $waits[0] -eq 7 -and $waits[1] -ge 10 -and $waits[1] -le 20 -and @($r.Calls | Where-Object { -not $_.TimeoutSec }).Count -eq 0) "Transient download errors were not retried within a time limit: $($r.Sleeps)"
    # A 401 is not a lag: immediate failure.
    $r = Invoke-Download @((New-HttpFailure 401))
    Check (-not $r.Saved -and @($r.Calls).Count -eq 1 -and $r.Error -match 'Status=401') 'A 401 on the export download was retried.'

    # ---- EvidenceCollector (ManagedApps, Configurations, EnrollmentFailures...) ----
    $evidence = Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.13' -Force -PassThru
    $invoke = { param($answers) & $evidence { param($answers) $script:n = 0; $script:q = [Collections.Generic.Queue[object]]::new(); foreach ($a in $answers) { $script:q.Enqueue($a) }
        $result = [pscustomobject]@{ Value = $null; Error = ''; Calls = 0 }
        try { $result.Value = Invoke-SmartM365EvidenceGraphRequest -Uri 'https://graph.microsoft.com/v1.0/x' -MaxRetryCount 4 -NoSleep -Invoker { $script:n++; $a = $script:q.Dequeue(); if ($a -is [Exception]) { throw $a }; $a } } catch { $result.Error = $_.Exception.Message }
        $result.Calls = $script:n; $result } $answers }
    $r = & $invoke @((New-HttpFailure 500), 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'EvidenceCollector did not retry a 500.'
    $r = & $invoke @((New-Timeout), 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'EvidenceCollector did not retry a timeout.'
    $r = & $invoke @((New-HttpFailure 403), 'ok'); Check ($r.Calls -eq 1 -and $r.Error) 'EvidenceCollector retried a 403.'
    $r = & $invoke @(([Exception]::new('Failed for https://graph.microsoft.com/beta/reports/500/items')), 'ok'); Check ($r.Calls -eq 1) 'EvidenceCollector read a number in a URL as a 500.'
    # Completed job without URL (20/10): polled again, then downloaded without Authorization.
    $rows = & $evidence {
        $script:jobPolls = 0
        function script:Invoke-MgGraphRequest { [CmdletBinding()] param($Method, $Uri, $Body, $ContentType, $OutputType)
            if ($Method -eq 'POST') { return [pscustomobject]@{ id = 'job-1' } }
            $script:jobPolls++
            if ($script:jobPolls -eq 1) { return [pscustomobject]@{ status = 'completed'; url = '' } }
            [pscustomobject]@{ status = 'completed'; url = 'https://intunestorage.example.invalid/x.zip?sig=S' } }
        function script:Invoke-SmartM365ExportFileDownload { param($Url, $OutFile, $Operation)
            $folder = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-EvidenceZip-' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $folder | Out-Null
            try { [IO.File]::WriteAllText((Join-Path $folder 'r.csv'), "Id,Name`r`n1,App") ; Compress-Archive -Path (Join-Path $folder 'r.csv') -DestinationPath $OutFile -Force } finally { Remove-Item -LiteralPath $folder -Recurse -Force }
            $OutFile }
        function script:Start-Sleep { param([int]$Seconds) $null = $Seconds }
        function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
        $result = @(Invoke-SmartM365EvidenceExportReport -ReportName 'AppInvRawData' -PollSeconds 1)
        [pscustomobject]@{ Rows = $result.Count; Polls = $script:jobPolls }
    }
    Check ($rows.Rows -eq 1 -and $rows.Polls -eq 2) 'A completed export without URL was not polled again.'
    $evidenceText = [IO.File]::ReadAllText((Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psm1'))
    Check ($evidenceText -notmatch 'Invoke-WebRequest -Uri \$downloadUrl') 'EvidenceCollector downloads the export outside the Core helper.'

    # ---- Endpoint Analytics ----
    $ea = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1' -Names @('Get-EAStatusCode', 'Get-EARetryAfterSeconds', 'Invoke-EAGraphRequest', 'Wait-EAExportJob') -Prelude ($standIns + "`nfunction Write-EALog { param(`$Message, `$Level) `$null = `$Message, `$Level }`n`$MaxRetryCount = 4`n`$script:GraphApiBase = 'https://graph.microsoft.com'")
    try {
        $run = { param($answers) & $ea { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Value = $null; Error = ''; Calls = 0 }; try { $r.Value = Invoke-EAGraphRequest GET 'https://graph.microsoft.com/beta/x' -NoSleep -Invoker { Next-Answer } } catch { $r.Error = $_.Exception.Message }; $r.Calls = $script:Calls; $r } $answers }
        $r = & $run @((New-Timeout), 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Endpoint Analytics did not retry a timeout.'
        $r = & $run @((New-HttpFailure 403), 'ok'); Check ($r.Calls -eq 1) 'Endpoint Analytics retried a 403.'
        $grain = [Exception]::new('Grain mismatch: 404 rows timed out'); $grain.Data['EndpointAnalyticsGrain'] = $true
        $r = & $run @($grain, 'ok'); Check ($r.Calls -eq 1) 'Endpoint Analytics retried a data grain failure.'
        $waited = & $ea { $script:poll = 0; Wait-EAExportJob -JobId 'j' -ApiVersion 'beta' -TimeoutSeconds 60 -PollSeconds 1 -NoSleep -StatusProvider { param($n) if ($n -lt 3) { [pscustomobject]@{ status = 'completed'; url = '' } } else { [pscustomobject]@{ status = 'completed'; url = 'https://x.invalid/a.zip' } } } }
        Check ([string]$waited.url -eq 'https://x.invalid/a.zip') 'Endpoint Analytics returned a completed job without waiting for its URL.'
    } finally { Remove-Module $ea -Force }

    # ---- Autopatch ----
    $ap = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1' -Names @('Get-AutopatchGraphStatusCode', 'Get-AutopatchGraphRetryDelaySeconds', 'Invoke-AutopatchGraphRequest') -Prelude ($standIns + "`n`$script:GraphTransientRetryCount = 0`nfunction Invoke-MgGraphRequest { [CmdletBinding()] param(`$Method, `$Uri, `$Body, `$ContentType, `$OutputType) Next-Answer }")
    try {
        $run = { param($answers) & $ap { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Value = $null; Calls = 0 }; try { $r.Value = Invoke-AutopatchGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/x' } catch { $null = $_ }; $r.Calls = $script:Calls; $r } $answers }
        $r = & $run @((New-Timeout), 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Autopatch did not retry a timeout.'
        $r = & $run @(([Exception]::new('InternalServerError (Internal Server Error)')), 'ok'); Check ($r.Value -eq 'ok') 'Autopatch did not retry a Graph SDK 500.'
        $r = & $run @((New-HttpFailure 403), 'ok'); Check ($r.Calls -eq 1) 'Autopatch retried a 403.'
    } finally { Remove-Module $ap -Force }

    # ---- BIOS (status was read as text: 5xx/408 were permanent) ----
    $bios = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1' -Names @('Get-GraphErrorDetails', 'Invoke-GraphRequestWithRetry') -Prelude ($standIns + "`nfunction Invoke-MgGraphRequest { [CmdletBinding()] param(`$Method, `$Uri, `$Body, `$ContentType, `$ResponseHeadersVariable, `$StatusCodeVariable) Next-Answer }")
    try {
        $run = { param($answers) & $bios { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Value = $null; Calls = 0 }; try { $r.Value = (Invoke-GraphRequestWithRetry -Method GET -Uri 'https://graph.microsoft.com/beta/x' -MaxRetries 3).Body } catch { $null = $_ }; $r.Calls = $script:Calls; [pscustomobject]@{ Value = $r.Value; Calls = $r.Calls; Slept = ($script:Slept -join ',') } } $answers }
        $sdk503 = [Exception]::new('Response status code does not indicate success: ServiceUnavailable (Service Unavailable).')
        $r = & $run @($sdk503, 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'BIOS did not retry a Graph SDK 503.'
        $r = & $run @((New-HttpFailure 429 '9'), 'ok'); Check ($r.Value -eq 'ok' -and $r.Slept -eq '9') "BIOS did not honour Retry-After: $($r.Slept)"
        $r = & $run @((New-HttpFailure 403), 'ok'); Check ($r.Calls -eq 1) 'BIOS retried a 403.'
    } finally { Remove-Module $bios -Force }

    # ---- Backup protected mailboxes (no retry before) ----
    $backup = Import-CollectorFunction -RelativePath 'SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1' -Names @('Get-SmartM365GraphErrorBody', 'Invoke-SmartM365GraphCollectionRequest') -Prelude ($standIns + "`nfunction Invoke-MgGraphRequest { [CmdletBinding()] param(`$Method, `$Uri, `$OutputType) Next-Answer }")
    try {
        $page = '{"value":[{"id":"m1"}]}'
        $run = { param($answers) & $backup { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Count = -1; Calls = 0 }; try { $r.Count = @(Invoke-SmartM365GraphCollectionRequest -Uri 'https://graph.microsoft.com/beta/x').Count } catch { $null = $_ }; $r.Calls = $script:Calls; $r } $answers }
        $r = & $run @((New-HttpFailure 503), (New-Timeout), $page); Check ($r.Count -eq 1 -and $r.Calls -eq 3) 'Backup did not retry transient Graph errors.'
        $r = & $run @((New-HttpFailure 403), $page); Check ($r.Count -eq -1 -and $r.Calls -eq 1) 'Backup retried a 403.'
    } finally { Remove-Module $backup -Force }

    # ---- Teams batch transport (retried 401/403 before) ----
    $teams = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1' -Names @('Get-TeamsRetryDelay', 'Invoke-TeamsGraphBatch') -Prelude ($standIns + "`nfunction Invoke-MgGraphRequest { [CmdletBinding()] param(`$Method, `$Uri, `$Body, `$ContentType) Next-Answer }")
    try {
        $ok = [pscustomobject]@{ responses = @([pscustomobject]@{ id = '1'; status = 200; body = [pscustomobject]@{ id = 't' }; headers = @{} }) }
        $run = { param($answers) & $teams { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Count = -1; Calls = 0 }; try { $r.Count = @(Invoke-TeamsGraphBatch -Requests @([pscustomobject]@{ id = '1'; method = 'GET'; url = '/teams/t' }) -MaxAttempts 3).Count } catch { $null = $_ }; $r.Calls = $script:Calls; $r } $answers }
        $r = & $run @((New-HttpFailure 403), $ok); Check ($r.Calls -eq 1 -and $r.Count -eq -1) 'Teams retried a 403 batch transport failure.'
        $r = & $run @((New-HttpFailure 503), $ok); Check ($r.Calls -eq 2 -and $r.Count -eq 1) 'Teams did not retry a 503 batch transport failure.'
    } finally { Remove-Module $teams -Force }

    # ---- Discovered Apps (Retry-After ignored on the HTTP-status path) ----
    $discovered = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1' -Names @('Get-DiscoveredAppsGraphHttpErrorMessage', 'Get-ShortGraphErrorMessage', 'Get-GraphRetryDelaySeconds', 'Invoke-GraphPagedRequest') -Prelude ($standIns + @'

$script:Stat_GraphCalls = 0; $script:Stat_ThrottleRetries = 0; $script:GraphRetryMaxSeconds = 300; $script:GraphMaxRetryAttempts = 3; $script:GraphRetryDefaultSeconds = 2
function Invoke-MgGraphRequest {
    [CmdletBinding()] param($Method, $Uri, $OutputType, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $ResponseHeadersVariable)
    $answer = Next-Answer
    Set-Variable -Name $StatusCodeVariable -Value $answer.Status -Scope 1
    Set-Variable -Name $ResponseHeadersVariable -Value $answer.Headers -Scope 1
    $answer.Body
}
'@)
    try {
        $state = & $discovered {
            Set-Answers @([pscustomobject]@{ Status = 429; Headers = @{ 'Retry-After' = @('11') }; Body = [pscustomobject]@{ error = [pscustomobject]@{ message = 'TooManyRequests' } } }, [pscustomobject]@{ Status = 200; Headers = @{}; Body = [pscustomobject]@{ value = [Collections.Generic.List[object]]@([pscustomobject]@{ id = 'a' }) } })
            $items = @(Invoke-GraphPagedRequest -InitialUri 'https://graph.microsoft.com/beta/x')
            [pscustomobject]@{ Count = $items.Count; Slept = ($script:Slept -join ',') }
        }
        Check ($state.Count -eq 1 -and $state.Slept -eq '11') "Discovered Apps ignored Retry-After on the HTTP-status path: $($state.Slept)"
        $state = & $discovered {
            Set-Answers @([pscustomobject]@{ Status = 403; Headers = @{}; Body = [pscustomobject]@{ error = [pscustomobject]@{ message = 'Forbidden' } } })
            $failed = $false; try { $null = Invoke-GraphPagedRequest -InitialUri 'https://graph.microsoft.com/beta/x' } catch { $failed = $true }
            [pscustomobject]@{ Failed = $failed; Calls = $script:Calls }
        }
        Check ($state.Failed -and $state.Calls -eq 1) 'Discovered Apps retried a 403.'
    } finally { Remove-Module $discovered -Force }

    # ---- Compliance (timeouts, connection resets) ----
    $compliance = Import-CollectorFunction -RelativePath 'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1' -Names @('Get-ShortGraphErrorMessage', 'Get-GraphRetryDelaySeconds', 'Invoke-WithRetry') -Prelude ($standIns + "`n`$script:GraphRequestDelayMs = 0; `$script:GraphMaxRetryAttempts = 3; `$script:GraphRetryMaxSeconds = 60`nfunction Write-ComplianceWarning { param(`$Message) `$null = `$Message }")
    try {
        $run = { param($answers) & $compliance { param($answers) Set-Answers $answers; $r = [pscustomobject]@{ Value = $null; Calls = 0 }; try { $r.Value = Invoke-WithRetry -Script { Next-Answer } } catch { $null = $_ }; $r.Calls = $script:Calls; $r } $answers }
        $r = & $run @(([Net.Http.HttpRequestException]::new('An error occurred while sending the request.')), 'ok'); Check ($r.Value -eq 'ok' -and $r.Calls -eq 2) 'Compliance did not retry a connection error.'
        $r = & $run @((New-HttpFailure 403), 'ok'); Check ($r.Calls -eq 1) 'Compliance retried a 403.'
    } finally { Remove-Module $compliance -Force }

    # Static: export downloads go through the Core helper, versions agree, collectors require Core 1.0.98.
    foreach ($relative in @(
        'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1',
        'SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1',
        'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1')) {
        $text = [IO.File]::ReadAllText((Join-Path $root $relative))
        $name = [IO.Path]::GetFileName($relative)
        $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
        Check ($version -and $text -match ('\$(?:script:)?ScriptVersion ?= ?[''"]' + [regex]::Escape($version) + '[''"]')) "$name header and runtime versions differ."
        Check ($text -match "-MinimumVersion '1\.0\.(?:9[89]|[1-9]\d\d)'") "$name does not require Core 1.0.98."
        Check ($text -notmatch 'Invoke-WebRequest\s+-Uri\s+\$(?:downloadUrl|completedJob\.url|Uri)\s+-OutFile') "$name downloads an export outside the Core helper."
    }
}
finally {
    Remove-Module SmartM365.EvidenceCollector.Common -Force -ErrorAction SilentlyContinue
    if ($temporary -and (Test-Path -LiteralPath $temporary)) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks collector Graph retry and export download checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDSH5lYlu8TJsGp
# C2/g1TBcqyERXewofHj2HK7qcXl9EqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG85WHbl9fKq+ao6qCnDufcgXYL11AdhTObb/ok1x7d5MA0GCSqG
# SIb3DQEBAQUABIIBgLFs5Xu3j8ZgxBaYb20Uhc9vJUtz7wrHRh1XzDJhICnrEwpJ
# vzY9mroBbWQJayS+05vORUvPJIWPra5zAN/y3praudhqyahzlZoddGBV7QZTRS48
# SMGgkZt77CSY+4YNzWUEBk8sjgzKdJUnrTUnH/Lz0BvtQs5WE1Imj7OfaUBGJCa7
# v1tMJNYcyzdaSU9F5l1wuQibDqbMh28hch5eVmLQDkACLrjIv0NW79rM5dijxGiw
# /bGpaCtMIgkPN7V4eCCDj+4SA/GLQf9sF6Gd/rpic0wF4kQYd/dEmzWR7F2PXHER
# /7qScTRqX1ot5wbXycmwQ2kDxHrEVKrV7wxduiW1Zoh5RLPuZBNAD/a/ScC7jcS4
# w2kFWxdM8k4gx8aYwPioGpbFqTwujfRM2aJTeUIhYS4bm+crn5VI7H59LQVPfEcA
# 7dWJS3Ynlq8v3CLolR5DERLaJEx+oTxPL3XK5ksrqP/6tkxhY/zxHb69naElDC8b
# 6SW13arVjKLiK6ZwAKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIwOTI4
# NTFaMC8GCSqGSIb3DQEJBDEiBCDek3xNe/3F7IP+hpU4z+oI247ej+z3CJs+3c6j
# Bp3m5TANBgkqhkiG9w0BAQEFAASCAgAGXibNPDTSAhEuU7GC2vAqLvo8X0xerm9c
# MgZoUmiIr6q6KPOGn2QmLrsv7/53OAM1H2s1KjayQtC8HJ/unHKpoYOUXTBNWsv1
# 08PjBPB2lgLjdbKN0i22crA0FVsPgetIb/S8EXR05laMf08ZTZAYkfT+UEi1fgaS
# lyNy7cmClR9lbFPslZTyvat4tavqOh42vCEG4Rf61jScfHbym2I9AtBe+mXAjfi6
# qdTEQ+sKg303pF4/PSowv3Y2SuYf8TFJWdr0sVSbWyZJ8DytPK34CIZ/5NzVylIX
# ZP5X8yE3C88vH3+g64GQgtLD8o8zIUvP2pWtIVDQDvOpbRBeAZCE+gQ3+INwqQmC
# zOEsfJVVsO7ovQ6sKfoLnAcPM6U2769g57eZaCH0zhlFtyNj47r1Us2sNxZTWmSF
# /kfiHQxI4fKG9AZpRr34cjwLK9tnZQo3kHu9s+0AyiP07+x84BAtu0q2cH02VSRs
# Fc95Uh1XJjZoFl26lZM0MLskGbrfT5b/1CQn/Li9QzpGyWY0x6ug+zi5FGd5w9Nf
# gdYEz94lEWpV9OotqNhctLLqOzNt4pdra2juAVQsR82GTj+OlqgWX0B2KzZMwBMb
# 30pnAbbx/fJmHqhRXXFDI2NGGwtlxJYgN8cSvuiIoML/SIwJAv8rRjkgmNKrxZxU
# KRX5/IrKtQ==
# SIG # End signature block
