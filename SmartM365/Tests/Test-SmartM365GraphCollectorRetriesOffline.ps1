#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the collector Graph retries and of the Intune export download: 401/403 fail at once,
408/429/5xx, timeouts and connection resets are retried, and a pre-signed export URL is downloaded
without an Authorization header, its real error logged without the URL signature.
.VERSION
1.0.0
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
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.99' -Force -PassThru -Global
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
    Check ($r.Saved -and $r.Sleeps -eq '7,20' -and @($r.Calls | Where-Object { -not $_.TimeoutSec }).Count -eq 0) "Transient download errors were not retried within a time limit: $($r.Sleeps)"
    # A 401 is not a lag: immediate failure.
    $r = Invoke-Download @((New-HttpFailure 401))
    Check (-not $r.Saved -and @($r.Calls).Count -eq 1 -and $r.Error -match 'Status=401') 'A 401 on the export download was retried.'

    # ---- EvidenceCollector (ManagedApps, Configurations, EnrollmentFailures...) ----
    $evidence = Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.12' -Force -PassThru
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
