#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the preflight Graph probe and of the orchestrator cloud capability probes: transient
Graph errors (5xx, 429 with Retry-After) are retried, a persistent one is reported as a transient
service error distinct from a permission failure, and 401/403 fail at once without retry.
.VERSION
1.0.0
.NOTES
Graph calls are replaced by synthetic HTTP responses. No tenant, Graph, certificate or network access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Recorders accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic module test state only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='The probe attempt setting is saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.97' -Force -PassThru
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$probeUri = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=1'
$savedAttempts = Get-Variable -Name GraphProbeMaxAttempts -Scope Global -ErrorAction Ignore
$global:GraphProbeMaxAttempts = $null

& $core {
    $script:Calls = New-Object 'Collections.Generic.List[object]'
    $script:Sleeps = New-Object 'Collections.Generic.List[int]'
    $script:Logs = New-Object 'Collections.Generic.List[string]'
    $script:Answers = New-Object 'Collections.Generic.Queue[object]'
    # Each answer is a status code (0 = success) or an exception to throw.
    function script:New-SyntheticHttpFailure {
        param([int]$Status, [string]$RetryAfter)
        $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
        if ($RetryAfter) { $response.Headers.Add('Retry-After', $RetryAfter) }
        [Microsoft.PowerShell.Commands.HttpResponseException]::new(("Response status code does not indicate success: {0} ({1})." -f $Status, $response.ReasonPhrase), $response)
    }
    function script:Invoke-SyntheticAnswer {
        param([string]$Kind, [string]$Uri, $TimeoutSec)
        $script:Calls.Add([pscustomobject]@{ Kind = $Kind; Uri = $Uri; TimeoutSec = $TimeoutSec })
        $answer = $script:Answers.Dequeue()
        if ($answer -is [Exception]) { throw $answer }
        [pscustomobject]@{ value = @() }
    }
    function script:Invoke-RestMethod { [CmdletBinding()] param($Method, $Uri, $Headers, $TimeoutSec) Invoke-SyntheticAnswer -Kind 'Rest' -Uri $Uri -TimeoutSec $TimeoutSec }
    function script:Invoke-MgGraphRequest { [CmdletBinding()] param($Method, $Uri, $OutputType) Invoke-SyntheticAnswer -Kind 'Mg' -Uri $Uri }
    function script:Get-MgContext { [CmdletBinding()] param() [pscustomobject]@{ ClientId = ''; Scopes = @() } }
    function script:Start-Sleep { param([int]$Seconds) $script:Sleeps.Add($Seconds) }
    function script:WriteLog { param($Message, $Level) $script:Logs.Add(("{0} {1}" -f $Level, $Message)) }
}
function Invoke-Preflight([object[]]$Answers, [switch]$UseGraphSdk) {
    & $core {
        param($Answers, $UseGraphSdk, $Uri)
        $script:Calls.Clear(); $script:Sleeps.Clear(); $script:Logs.Clear(); $script:Answers.Clear()
        foreach ($answer in $Answers) {
            $script:Answers.Enqueue($(if ($answer -is [int] -and $answer -gt 0) { New-SyntheticHttpFailure -Status $answer } else { $answer }))
        }
        $token = if ($UseGraphSdk) { '' } else { 'synthetic.token.value' }
        $outcome = [pscustomobject]@{ Passed = $false; Error = '' }
        try { $outcome.Passed = [bool](Invoke-SmartM365Preflight -ScriptName 'Synthetic' -GraphProbeUris @($Uri) -GraphAccessToken $token) }
        catch { $outcome.Error = $_.Exception.Message }
        [pscustomobject]@{ Passed = $outcome.Passed; Error = $outcome.Error; Calls = $script:Calls.Count; Sleeps = ($script:Sleeps -join ','); Logs = $script:Logs.ToArray(); Timeouts = @($script:Calls | ForEach-Object TimeoutSec) }
    } $Answers ([bool]$UseGraphSdk) $probeUri
}

try {
    # 1. 500 then success: the preflight passes after one progressive wait.
    $r = Invoke-Preflight @(500, 0)
    Check ($r.Passed -and $r.Calls -eq 2 -and $r.Sleeps -eq '15') "A transient 500 failed the preflight: $($r.Error)"
    Check (@($r.Logs | Where-Object { $_ -like 'WARNING Preflight Graph probe transient failure*Status=500; attempt 1/4*' }).Count -eq 1 -and @($r.Logs | Where-Object { $_ -like 'INFO Preflight Graph permission OK:*(after 2 attempts)' }).Count -eq 1) 'The retried probe is not logged.'
    Check (@($r.Timeouts | Where-Object { $_ -ne 180 }).Count -eq 0) 'The probe lost its request time limit.'

    # 2. 500 on every attempt: a persistent transient error, not a permission failure.
    $r = Invoke-Preflight @(500, 500, 500, 500)
    Check (-not $r.Passed -and $r.Calls -eq 4 -and $r.Sleeps -eq '15,30,45') 'A persistent 500 was not retried with progressive waits.'
    Check ($r.Error -match [regex]::Escape("Persistent transient Microsoft Graph error for '$probeUri' after 4 attempt(s) (Status=500); not a permission failure") -and $r.Error -notmatch 'Graph permission probe failed') "A persistent 500 is reported as a permission failure: $($r.Error)"

    # 3. 403: the permission failure is immediate and keeps its message.
    $r = Invoke-Preflight @(403)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Sleeps -eq '' -and $r.Error -match [regex]::Escape("Graph permission probe failed for '$probeUri'")) "A 403 was retried or reworded: $($r.Error)"
    $r = Invoke-Preflight @(401)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A 401 was retried.'

    # 4. 429 with Retry-After: the wait follows the service.
    $throttled = & $core { New-SyntheticHttpFailure -Status 429 -RetryAfter '7' }
    $r = Invoke-Preflight @($throttled, 0)
    Check ($r.Passed -and $r.Calls -eq 2 -and $r.Sleeps -eq '7') "A 429 did not honour Retry-After: $($r.Sleeps)"

    # 5. Graph SDK path: the status exists only in the message, as in the 17/10 failure.
    $sdk500 = [Exception]::new('Response status code does not indicate success: InternalServerError (Internal Server Error).')
    $r = Invoke-Preflight @($sdk500, 0) -UseGraphSdk
    Check ($r.Passed -and $r.Calls -eq 2) "A Graph SDK 500 failed the preflight: $($r.Error)"
    $sdk403 = [Exception]::new('Forbidden (Forbidden). Request timed out while authorizing.')
    $r = Invoke-Preflight @($sdk403) -UseGraphSdk
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A Graph SDK 403 was retried because its text mentions a timeout.'

    # 6. A 404 (wrong endpoint) is not transient: unchanged permission-probe failure.
    $r = Invoke-Preflight @(404)
    Check (-not $r.Passed -and $r.Calls -eq 1 -and $r.Error -match 'Graph permission probe failed') 'A 404 was retried.'

    # 7. Orchestrator cloud capability probes: transient errors retried, authorization errors not.
    $distributed = [IO.File]::ReadAllText((Join-Path $root 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1'))
    $retryText = [regex]::Match($distributed, "(?s)\`$cloudProbeRetry = @'\r?\n(?<Body>.*?)\r?\n'@").Groups['Body'].Value
    Check ($retryText -match 'function Invoke-ProbeWithRetry') 'The capability probe retry helper is missing.'
    Check ($distributed -match "\`$graphScript = \`$cloudProbeRetry \+" -and $distributed -match "\`$exoScript = \`$cloudProbeRetry \+" -and $distributed -match 'Invoke-ProbeWithRetry \{ Connect-MgGraph' -and $distributed -match '(?s)Invoke-ProbeWithRetry \{\s+Connect-ExchangeOnline') 'Graph or EXO capability probes do not retry transient errors.'
    $probe = New-Module -ScriptBlock ([scriptblock]::Create($retryText + "`nfunction Start-Sleep { param([int]`$Seconds) `$script:Slept += `$Seconds }`n`$script:Slept = 0"))
    try {
        $state = & $probe { $script:n = 0; $value = Invoke-ProbeWithRetry { $script:n++; if ($script:n -lt 3) { throw 'Response status code does not indicate success: ServiceUnavailable (Service Unavailable).' }; 'ok' }; [pscustomobject]@{ Value = $value; Calls = $script:n; Slept = $script:Slept } }
        Check ($state.Value -eq 'ok' -and $state.Calls -eq 3 -and $state.Slept -eq 15) 'A transient capability probe error was not retried.'
        $state = & $probe { $script:n = 0; $message = ''; try { Invoke-ProbeWithRetry { $script:n++; throw 'An error occurred while sending the request.' } } catch { $message = $_.Exception.Message }; [pscustomobject]@{ Calls = $script:n; Message = $message } }
        Check ($state.Calls -eq 3 -and $state.Message -like 'Persistent transient error after 3 attempt(s): *') 'A persistent capability probe error is not reported as such.'
        $state = & $probe { $script:n = 0; try { Invoke-ProbeWithRetry { $script:n++; throw 'AADSTS700027: Client assertion failed; request timed out.' } } catch { $null = $_ }; $script:n }
        Check ($state -eq 1) 'An authorization error of a capability probe was retried.'
    }
    finally { Remove-Module $probe -Force }
}
finally {
    if ($savedAttempts) { $global:GraphProbeMaxAttempts = $savedAttempts.Value } else { Remove-Variable -Name GraphProbeMaxAttempts -Scope Global -ErrorAction Ignore }
    Remove-Module $core -Force -ErrorAction SilentlyContinue
}
"PASS: $script:checks preflight Graph probe checks; PowerShell $($PSVersionTable.PSVersion)."
