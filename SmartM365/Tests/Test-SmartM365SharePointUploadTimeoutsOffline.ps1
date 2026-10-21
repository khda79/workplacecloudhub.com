#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the SharePoint upload time limits: a token request, an upload session and a chunk
that never answer are bounded by their timeout, retried, and a later attempt succeeds; an endpoint
that stays silent ends with a clear error, and the per-file limit stops the upload with a warning.
.VERSION
1.0.0
.NOTES
The silent endpoint is a local TCP listener on the loopback address that accepts connections and never
answers. Every other Graph and SharePoint call is replaced by recorders. No tenant, Graph, SharePoint,
certificate or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic upload globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Recorders accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic fixtures in TEMP and module test state only.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.96' -Force -PassThru
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$globalNames = @('AppId', 'TenantId', 'Thumb', 'Thumbprint', 'EnableSharePointUpload', 'SharePointSiteHostname', 'SharePointSitePath', 'SharePointLibraryDisplayName', 'SharePointTargetFolderPath',
    'GraphTokenTimeoutSeconds', 'GraphTokenMaxAttempts', 'GraphRequestTimeoutSeconds', 'SharePointUploadTransferTimeoutSeconds', 'SharePointUploadChunkTimeoutSeconds', 'SharePointUploadFileTimeoutMinutes', 'SmartM365SharePointUploadedFiles')
$saved = @{}
foreach ($name in $globalNames) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:AppId = 'synthetic-app'; $global:TenantId = 'synthetic-tenant'; $global:Thumb = 'SYNTHETIC'; $global:Thumbprint = 'SYNTHETIC'
$global:EnableSharePointUpload = $true; $global:SharePointSiteHostname = 'contoso.test'; $global:SharePointSitePath = '/sites/synthetic'
$global:SharePointLibraryDisplayName = 'Documents'; $global:SharePointTargetFolderPath = 'SMART-M365/DATA'
# Shortest accepted limits keep the silent-endpoint cases within seconds.
$global:GraphTokenTimeoutSeconds = 5; $global:GraphRequestTimeoutSeconds = 10; $global:SharePointUploadChunkTimeoutSeconds = 30
$global:GraphTokenMaxAttempts = $null; $global:SharePointUploadTransferTimeoutSeconds = $null; $global:SharePointUploadFileTimeoutMinutes = $null

$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
$listener.Start()
$silentUri = 'http://127.0.0.1:{0}/silent' -f ([Net.IPEndPoint]$listener.LocalEndpoint).Port
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-UploadTimeouts-' + [guid]::NewGuid().ToString('N'))
$dataAll = Join-Path $temporary 'DATA-ALL/Intune'
New-Item -ItemType Directory -Path $dataAll | Out-Null

& $core {
    param($SilentUri)
    $script:SilentUri = $SilentUri
    $script:Calls = New-Object 'Collections.Generic.List[object]'
    $script:Logs = New-Object 'Collections.Generic.List[object]'
    $script:Handler = $null
    # Each call is recorded with its timeout; the handler answers, throws, or sends it to the silent endpoint.
    function script:Invoke-SyntheticTimeout { param($Seconds) throw [System.Threading.Tasks.TaskCanceledException]::new(("The request was canceled due to the configured HttpClient.Timeout of {0} seconds elapsing." -f $Seconds)) }
    function script:Invoke-SilentRequest { param($TimeoutSec) $request = @{ Method = 'GET'; Uri = $script:SilentUri; ErrorAction = 'Stop' }; if ($TimeoutSec) { $request.TimeoutSec = $TimeoutSec }; Microsoft.PowerShell.Utility\Invoke-RestMethod @request }
    function script:Invoke-RestMethod {
        [CmdletBinding()]
        param($Method, $Uri, $Body, $ContentType, $Headers, $TimeoutSec, [switch]$SkipHttpErrorCheck, $StatusCodeVariable, $ResponseHeadersVariable)
        $call = [pscustomobject]@{ Kind = 'Rest'; Method = $Method; Uri = [string]$Uri; TimeoutSec = $TimeoutSec; Index = $script:Calls.Count }
        $script:Calls.Add($call)
        & $script:Handler $call
    }
    function script:Invoke-WebRequest {
        [CmdletBinding()]
        param($Method, $Uri, $Body, $ContentType, $Headers, $TimeoutSec, [switch]$UseBasicParsing, [switch]$SkipHttpErrorCheck, $OutFile)
        $call = [pscustomobject]@{ Kind = 'Web'; Method = $Method; Uri = [string]$Uri; TimeoutSec = $TimeoutSec; Index = $script:Calls.Count; Range = [string]$Headers['Content-Range'] }
        $script:Calls.Add($call)
        & $script:Handler $call
    }
    function script:Start-Sleep { param([int]$Seconds) $script:Logs.Add([pscustomobject]@{ Level = 'SLEEP'; Message = [string]$Seconds }) }
    function script:WriteLog { param($Message, $Level) $script:Logs.Add([pscustomobject]@{ Level = [string]$Level; Message = [string]$Message }) }
    function script:Get-SmartM365CertificateByThumbprint { param($Thumbprint) [pscustomobject]@{ Thumbprint = $Thumbprint } }
    function script:New-SmartM365GraphClientAssertion { param($TenantId, $ClientId, $Certificate) 'synthetic.assertion.value' }
    function script:Test-SmartM365InterruptedRun { $false }
} $silentUri
function Reset-Recording([scriptblock]$Handler) {
    & $core {
        param($Handler)
        $script:Calls.Clear(); $script:Logs.Clear(); $script:Handler = $Handler
        $script:SmartM365GraphRestTokenCache = @{}; $script:SmartM365SharePointUploadConnectionKey = $null
        $script:SmartM365SharePointDriveIdCache = @{}; $script:SmartM365NetworkDeadline = $null
    } $core.NewBoundScriptBlock($Handler)
}
function Get-RecordedCall { & $core { $script:Calls.ToArray() } }
function Get-RecordedLog { & $core { $script:Logs.ToArray() } }
function Set-CachedToken { & $core { $script:SmartM365GraphRestTokenCache = @{ 'synthetic-app|synthetic-tenant|SYNTHETIC' = [pscustomobject]@{ AccessToken = 'cached'; ExpiresOn = (Get-Date).ToUniversalTime().AddHours(1) } } } }
function New-SyntheticFile([string]$Name, [int]$Size) { $path = Join-Path $dataAll $Name; $bytes = [byte[]]::new($Size); (New-Object Random 7).NextBytes($bytes); [IO.File]::WriteAllBytes($path, $bytes); $path }

try {
    # Defaults and bounds of the settings.
    $defaults = & $core {
        $global:GraphTokenMaxAttempts = 'invalid'
        $result = [pscustomobject]@{
            Attempts = Get-SmartM365NetworkSetting -Name 'GraphTokenMaxAttempts'
            Transfer = Get-SmartM365NetworkSetting -Name 'SharePointUploadTransferTimeoutSeconds'
            File = Get-SmartM365NetworkSetting -Name 'SharePointUploadFileTimeoutMinutes'
            Token = Get-SmartM365NetworkSetting -Name 'GraphTokenTimeoutSeconds'
        }
        $global:GraphTokenMaxAttempts = $null
        $result
    }
    Check ($defaults.Attempts -eq 3 -and $defaults.Transfer -eq 900 -and $defaults.File -eq 60 -and $defaults.Token -eq 5) 'Network settings do not apply their defaults, bounds or overrides.'

    # 1. Token endpoint silent on the first attempt: the real HTTP call ends at its timeout, the retry succeeds.
    Reset-Recording {
        param($call)
        if ($call.Index -eq 0) { Invoke-SilentRequest -TimeoutSec $call.TimeoutSec }
        [pscustomobject]@{ access_token = 'synthetic-token'; expires_in = 3600 }
    }
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $token = & $core { Get-SmartM365GraphAccessToken -Purpose 'Create SharePoint upload session' }
    $watch.Stop()
    $calls = @(Get-RecordedCall)
    Check ($token -eq 'synthetic-token' -and $calls.Count -eq 2) 'A silent token endpoint was not retried to success.'
    Check (@($calls | Where-Object { $_.TimeoutSec -ne 5 -or $_.Uri -notlike 'https://login.microsoftonline.com/*' }).Count -eq 0) 'The token request has no time limit.'
    Check ($watch.Elapsed.TotalSeconds -ge 4 -and $watch.Elapsed.TotalSeconds -lt 30) "The silent token endpoint was not bounded by its timeout ($([int]$watch.Elapsed.TotalSeconds)s)."
    Check (@(Get-RecordedLog | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -match 'token request .*transient failure.*attempt 1/3' }).Count -eq 1) 'The token retry is not logged.'

    # 2. Token endpoint that never answers: three bounded attempts, then a clear error.
    Reset-Recording { param($call) Invoke-SyntheticTimeout $call.TimeoutSec }
    $failure = $null
    try { & $core { Get-SmartM365GraphAccessToken -Purpose 'SharePoint upload' } | Out-Null } catch { $failure = $_.Exception.Message }
    Check (@(Get-RecordedCall).Count -eq 3 -and $failure -match 'SharePoint upload: Microsoft Graph token request failed after 3 attempt\(s\) \(limit 5s per attempt\)') "A silent token endpoint did not end with a clear error: $failure"
    Check (@(Get-RecordedLog | Where-Object Level -eq 'SLEEP' | ForEach-Object { [int]$_.Message }) -join ',' -eq '5,10') 'Token retries do not wait progressively.'

    # 3. Upload session silent once, then a chunk silent once: both are retried and the file completes.
    $file = New-SyntheticFile 'Synthetic_ManagedApps_20261021_082842.csv' 700000
    Set-CachedToken
    & $core {
        $script:Calls.Clear(); $script:Logs.Clear()
        $script:Handler = {
            param($call)
            if ($call.Uri -like '*createUploadSession') {
                if (@($script:Calls | Where-Object { $_.Uri -like '*createUploadSession' }).Count -eq 1) { Invoke-SilentRequest -TimeoutSec $call.TimeoutSec }
                return [pscustomobject]@{ uploadUrl = 'https://contoso.test/upload/session' }
            }
            if ($call.Kind -eq 'Web') {
                if ($call.Range -like 'bytes 327680-*' -and @($script:Calls | Where-Object { $_.Range -like 'bytes 327680-*' }).Count -eq 1) { Invoke-SyntheticTimeout $call.TimeoutSec }
                $last = $call.Range -match '-699999/700000$'
                return [pscustomobject]@{ StatusCode = $(if ($last) { 201 } else { 202 }); Content = $(if ($last) { '{"id":"item-1","webUrl":"https://contoso.test/item-1"}' } else { '{"nextExpectedRanges":["0-"]}' }) }
            }
            throw "Unexpected call $($call.Uri)"
        }
    }
    $item = & $core { param($Path) Invoke-SmartM365SharePointLargeFileUpload -LocalFilePath $Path -DriveId 'drive-1' -TargetPath 'SMART-M365/DATA/DATA-ALL/Intune/file.csv' -ChunkSizeBytes 327680 } $file
    $calls = @(Get-RecordedCall)
    $sessions = @($calls | Where-Object { $_.Uri -like '*createUploadSession' })
    $chunks = @($calls | Where-Object Kind -eq 'Web')
    Check ($item.id -eq 'item-1') 'The upload did not complete after the session and chunk retries.'
    Check ($sessions.Count -eq 2 -and @($sessions | Where-Object TimeoutSec -ne 10).Count -eq 0) 'The silent upload session was not bounded and retried.'
    Check ($chunks.Count -eq 4 -and @($chunks | Where-Object TimeoutSec -ne 30).Count -eq 0 -and ($chunks.Range -join '|') -eq 'bytes 0-327679/700000|bytes 327680-655359/700000|bytes 327680-655359/700000|bytes 655360-699999/700000') 'The silent chunk was not bounded and retried at the same range.'

    # 4. A chunk that never answers: four bounded attempts, then an error naming the range.
    Set-CachedToken
    & $core {
        $script:Calls.Clear(); $script:Logs.Clear()
        $script:Handler = {
            param($call)
            if ($call.Uri -like '*createUploadSession') { return [pscustomobject]@{ uploadUrl = 'https://contoso.test/upload/session' } }
            Invoke-SyntheticTimeout $call.TimeoutSec
        }
    }
    $failure = $null
    try { & $core { param($Path) Invoke-SmartM365SharePointLargeFileUpload -LocalFilePath $Path -DriveId 'drive-1' -TargetPath 'x.csv' -ChunkSizeBytes 327680 } $file | Out-Null } catch { $failure = $_.Exception.Message }
    Check (@(Get-RecordedCall | Where-Object Kind -eq 'Web').Count -eq 4 -and $failure -match 'SharePoint chunk upload failed\. Status=unknown\. Range=bytes 0-327679/700000\. Error=The request was canceled due to the configured HttpClient\.Timeout') "A silent chunk did not end with a clear error: $failure"

    # 5. Whole file: the upload of a file that outlives its limit stops with a warning, without hanging
    # or retrying past the limit, and the limit does not leak to later calls.
    $small = New-SyntheticFile 'Synthetic_Small_20261021_082842.csv' 2048
    Reset-Recording {
        param($call)
        if ($call.Uri -like 'https://login.microsoftonline.com/*') { return [pscustomobject]@{ access_token = 'synthetic-token'; expires_in = 3600 } }
        if ($call.Uri -like '*/sites/contoso.test:*') { return [pscustomobject]@{ id = 'site-1' } }
        if ($call.Uri -like '*/sites/site-1/drives') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'drive-1'; name = 'Documents' }) } }
        # The file upload stays silent until the per-file limit has passed.
        $script:SmartM365NetworkDeadline.ExpiresAtUtc = [datetime]::UtcNow.AddSeconds(-1)
        Invoke-SyntheticTimeout $call.TimeoutSec
    }
    $record = & $core { param($Path) Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Path } $small
    $calls = @(Get-RecordedCall)
    $put = @($calls | Where-Object Method -eq 'PUT')
    $warning = @(Get-RecordedLog | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -like 'SharePoint upload failed but script continues:*' })
    Check ($null -eq $record -and $put.Count -eq 1 -and $put[0].TimeoutSec -eq 900) 'The single-request upload has no transfer limit or retried past the file limit.'
    Check ($warning.Count -eq 1 -and $warning[0].Message -match 'Upload SharePoint file stopped: SharePoint upload of Synthetic_Small_20261021_082842\.csv exceeded its 60-minute limit') "The per-file limit did not end with a clear warning: $($warning.Message)"
    Check ($null -eq (& $core { $script:SmartM365NetworkDeadline })) 'The per-file limit leaked after the upload.'
    Check ((& $core { Get-SmartM365NetworkCallTimeout -TimeoutSeconds 42 }) -eq 42) 'Calls after the upload are still bounded by its limit.'

    # 6. A passed limit is never treated as transient, while timeouts and connection resets are.
    $classification = & $core {
        $deadline = [System.OperationCanceledException]::new('limit'); $deadline.Data['SmartM365DeadlineExceeded'] = $true
        $records = @(
            [Management.Automation.ErrorRecord]::new($deadline, 'x', 'OperationTimeout', $null),
            [Management.Automation.ErrorRecord]::new([System.Threading.Tasks.TaskCanceledException]::new('The request was canceled due to the configured HttpClient.Timeout of 5 seconds elapsing.'), 'x', 'OperationTimeout', $null),
            [Management.Automation.ErrorRecord]::new([Net.Http.HttpRequestException]::new('An error occurred while sending the request.'), 'x', 'ConnectionError', $null),
            [Management.Automation.ErrorRecord]::new([InvalidOperationException]::new('Forbidden'), 'x', 'PermissionDenied', $null)
        )
        @($records | ForEach-Object { Test-SmartM365GraphTransientError -ErrorRecord $_ -StatusCode $null }) -join ','
    }
    Check ($classification -eq 'False,True,True,False') "Transient classification changed: $classification"

    # Static wiring: every raw network call of Core carries a time limit.
    $text = [IO.File]::ReadAllText((Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1'))
    $ast = [Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$null)
    $raw = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('Invoke-RestMethod', 'Invoke-WebRequest') }, $true))
    # A splatted call is bounded when its function sets TimeoutSec in the splatted table.
    $unbounded = @($raw | Where-Object {
        $call = $_
        $function = $call
        while ($function -and $function -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $function = $function.Parent }
        $call.Extent.Text -notmatch '-TimeoutSec\b' -and -not ($call.Extent.Text -match '@(params|putParams)\b' -and $function -and $function.Extent.Text -match 'TimeoutSec\s*=')
    })
    Check ($raw.Count -ge 7 -and $unbounded.Count -eq 0) ("Core network calls without a time limit: " + (($unbounded | ForEach-Object { $_.Extent.StartLineNumber }) -join ','))
}
finally {
    $listener.Stop()
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    Remove-Module $core -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks SharePoint upload time limit checks; PowerShell $($PSVersionTable.PSVersion)."
