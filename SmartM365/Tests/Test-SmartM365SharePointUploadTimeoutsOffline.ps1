#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks of the SharePoint upload time limits: a token request, an upload session and a chunk
that never answer are bounded by their timeout, retried, and a later attempt succeeds; an endpoint
that stays silent ends with a clear error, and the per-file limit stops the upload with a warning.
A SharePoint maintenance 403 (serviceReadOnly) is retried while other 403s are final, and a source
receipt and its run state are not mirrored when a file they reference failed to upload in the run.
.VERSION
1.2.0
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
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.112' -Force -PassThru
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
    # Graph error answer with its JSON body, as Invoke-RestMethod raises it.
    function script:New-SyntheticGraphFailure {
        param([int]$Status, [string]$Code, [string]$InnerCode)
        $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$Status)
        $response.Content = [Net.Http.StringContent]::new((@{ error = @{ code = $Code; message = 'Synthetic'; innerError = @{ code = $InnerCode } } } | ConvertTo-Json -Depth 4))
        [Microsoft.PowerShell.Commands.HttpResponseException]::new(("Response status code does not indicate success: {0} (Forbidden)." -f $Status), $response)
    }
} $silentUri
function Reset-Recording([scriptblock]$Handler) {
    & $core {
        param($Handler)
        $script:Calls.Clear(); $script:Logs.Clear(); $script:Handler = $Handler
        $script:SmartM365GraphRestTokenCache = @{}; $script:SmartM365SharePointUploadConnectionKey = $null
        $script:SmartM365SharePointDriveIdCache = @{}; $script:SmartM365NetworkDeadline = $null; $script:SmartM365SharePointFailedUploads = $null
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

    # 7. SharePoint maintenance: a 403 serviceReadOnly is retried and the upload completes; any other 403 is final.
    $readOnlyFile = New-SyntheticFile 'Synthetic_ReadOnly_20261028_002030.csv' 1024
    Set-CachedToken
    & $core {
        $script:Calls.Clear(); $script:Logs.Clear(); $script:SmartM365SharePointUploadConnectionKey = $null
        $script:SmartM365SharePointDriveIdCache = @{ 'contoso.test|/sites/synthetic|Documents' = 'drive-1' }
        $script:Handler = {
            param($call)
            if ($call.Uri -like 'https://login.microsoftonline.com/*') { return [pscustomobject]@{ access_token = 'synthetic-token'; expires_in = 3600 } }
            if ($call.Method -eq 'PUT' -and @($script:Calls | Where-Object Method -eq 'PUT').Count -eq 1) { throw (New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'serviceReadOnly') }
            [pscustomobject]@{ id = 'item-1'; webUrl = 'https://contoso.test/item-1' }
        }
    }
    $record = & $core { param($Path) Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Path } $readOnlyFile
    Check ($record -and $record.ItemId -eq 'item-1' -and @(Get-RecordedCall | Where-Object Method -eq 'PUT').Count -eq 2) 'A SharePoint serviceReadOnly 403 was not retried to success.'
    Check (@(Get-RecordedLog | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -match '^Upload SharePoint file transient failure\. Status=403; attempt 1/4' }).Count -eq 1) 'The serviceReadOnly retry is not logged.'
    $forbidden = & $core {
        $readOnly = [Management.Automation.ErrorRecord]::new((New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'serviceReadOnly'), 'x', 'InvalidOperation', $null)
        $denied = [Management.Automation.ErrorRecord]::new((New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'accessDenied'), 'x', 'InvalidOperation', $null)
        @(
            (Test-SmartM365GraphTransientError -ErrorRecord $readOnly -StatusCode 403),
            (Test-SmartM365GraphTransientError -ErrorRecord $denied -StatusCode 403),
            (Test-SmartM365GraphRetryableError -ErrorRecord $readOnly).Retry
        ) -join ','
    }
    Check ($forbidden -eq 'True,False,False') "403 classification changed (serviceReadOnly, other 403, collector rule): $forbidden"
    # Windows PowerShell 5.1 module: same rule, evaluated from its own definition.
    $compatText = [IO.File]::ReadAllText((Join-Path $root 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'))
    $compatAst = [Management.Automation.Language.Parser]::ParseInput($compatText, [ref]$null, [ref]$null)
    $compatReadOnly = $compatAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-SmartM365SharePointReadOnlyError' }, $true)
    Check ($null -ne $compatReadOnly) 'The 5.1 module has no serviceReadOnly rule.'
    $compatClassification = & $core {
        param($Definition)
        . ([scriptblock]::Create($Definition))
        @(
            (Test-SmartM365SharePointReadOnlyError -ErrorRecord ([Management.Automation.ErrorRecord]::new((New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'serviceReadOnly'), 'x', 'InvalidOperation', $null)) -StatusCode 403),
            (Test-SmartM365SharePointReadOnlyError -ErrorRecord ([Management.Automation.ErrorRecord]::new((New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'accessDenied'), 'x', 'InvalidOperation', $null)) -StatusCode 403),
            (Test-SmartM365SharePointReadOnlyError -ErrorRecord ([Management.Automation.ErrorRecord]::new((New-SyntheticGraphFailure -Status 503 -Code 'serviceNotAvailable' -InnerCode 'serviceReadOnly'), 'x', 'InvalidOperation', $null)) -StatusCode 503)
        ) -join ','
    } $compatReadOnly.Extent.Text
    Check ($compatClassification -eq 'True,False,False') "5.1 serviceReadOnly classification changed: $compatClassification"
    Check (([regex]::Matches($compatText, 'Test-SmartM365SharePointReadOnlyError -ErrorRecord \$_ -StatusCode \$statusCode')).Count -eq 2) 'The 5.1 Graph request and chunk retries do not apply the serviceReadOnly rule.'

    # 8. Mirror coherence: when a CSV of the receipt failed to upload, neither the receipt nor the run state
    # is mirrored; a later successful upload of the same CSV lifts the block; files outside the receipt pass.
    $dataLast = Join-Path $temporary 'DATA-LAST'
    New-Item -ItemType Directory -Path $dataLast | Out-Null
    $csv = Join-Path $dataLast 'Synthetic_Users_Active.csv'
    [IO.File]::WriteAllText($csv, "Id`r`n1`r`n")
    $receipt = Join-Path $dataLast 'SmartInventory_Synthetic-Producer.current.json.txt'
    $runState = Join-Path $dataLast 'SmartInventory_Synthetic-Producer.run.json.txt'
    [IO.File]::WriteAllText($receipt, (@{ Status = 'Completed'; Files = @(@{ File = 'Synthetic_Users_Active.csv'; SHA256 = (Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash }) } | ConvertTo-Json -Depth 4))
    [IO.File]::WriteAllText($runState, '{"Status":"Completed"}')
    $other = New-SyntheticFile 'Synthetic_Other.log' 16
    Set-CachedToken
    & $core {
        $script:Calls.Clear(); $script:Logs.Clear(); $script:SmartM365SharePointFailedUploads = $null; $script:FailCsv = $true
        $script:SmartM365SharePointDriveIdCache = @{ 'contoso.test|/sites/synthetic|Documents' = 'drive-1' }
        $script:Handler = {
            param($call)
            if ($call.Uri -like 'https://login.microsoftonline.com/*') { return [pscustomobject]@{ access_token = 'synthetic-token'; expires_in = 3600 } }
            if ($call.Method -eq 'PUT' -and $call.Uri -like '*Synthetic_Users_Active.csv*' -and $script:FailCsv) { throw (New-SyntheticGraphFailure -Status 403 -Code 'accessDenied' -InnerCode 'accessDenied') }
            [pscustomobject]@{ id = 'item-2'; webUrl = 'https://contoso.test/item-2' }
        }
    }
    $blocked = & $core {
        param($Csv, $Receipt, $RunState, $Other)
        $null = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Csv
        @((Test-SmartM365SourceReceiptMirrorBlocked -Path $Receipt), (Test-SmartM365SourceReceiptMirrorBlocked -Path $RunState), (Test-SmartM365SourceReceiptMirrorBlocked -Path $Other)) -join ','
    } $csv $receipt $runState $other
    Check ($blocked -eq 'True,True,False') "A receipt or run state referencing a failed upload was not held back: $blocked"
    $mirrorWarnings = @(Get-RecordedLog | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -like 'SharePoint mirror of SmartInventory_Synthetic-Producer.*skipped: referenced file(s) not uploaded in this run: Synthetic_Users_Active.csv.*' })
    Check ($mirrorWarnings.Count -eq 2) 'The held-back receipt and run state are not reported as warnings.'
    $released = & $core {
        param($Csv, $Receipt, $RunState)
        $script:FailCsv = $false
        $null = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Csv
        @((Test-SmartM365SourceReceiptMirrorBlocked -Path $Receipt), (Test-SmartM365SourceReceiptMirrorBlocked -Path $RunState)) -join ','
    } $csv $receipt $runState
    Check ($released -eq 'False,False') "A successful re-upload did not release the receipt: $released"
    # Both execution summaries check the guard before each run artifact upload, and every outcome is recorded.
    $coreText = [IO.File]::ReadAllText((Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1'))
    $guardPattern = 'foreach \(\$logUploadCandidate in \$logUploadCandidates\) \{\s+if \(Test-SmartM365SourceReceiptMirrorBlocked -Path \$logUploadCandidate\) \{ continue \}\s+Invoke-SmartM365SharePointCsvUpload'
    Check ($coreText -match $guardPattern -and $compatText -match $guardPattern) 'A run artifact upload loop does not check the receipt mirror guard.'
    Check (([regex]::Matches($compatText, 'Set-SmartM365SharePointUploadOutcome -Path \$LocalFilePath')).Count -eq 4 -and ([regex]::Matches($coreText, 'Set-SmartM365SharePointUploadOutcome -Path \$LocalFilePath')).Count -eq 5) 'An upload outcome is not recorded in Core or the 5.1 module.'

    # Static wiring: every raw network call of Core carries a time limit.
    $text = [IO.File]::ReadAllText((Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1'))
    $ast = [Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$null)
    $raw = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('Invoke-RestMethod', 'Invoke-WebRequest') }, $true))
    # A call is bounded by a splatted table from Get-SmartM365WebRequestTimeout: TimeoutSec alone leaves the
    # read of a stalled response body unbounded (see Test-SmartM365NetworkReadTimeoutOffline.ps1).
    $unbounded = @($raw | Where-Object {
        $call = $_
        $function = $call
        while ($function -and $function -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $function = $function.Parent }
        $splatted = @($call.CommandElements | Where-Object { $_ -is [Management.Automation.Language.VariableExpressionAst] -and $_.Splatted }).Count -gt 0
        $call.Extent.Text -match '-TimeoutSec\b' -or -not ($splatted -and $function -and $function.Extent.Text -match 'Get-SmartM365WebRequestTimeout -Command')
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC59uoX4J3n4NkH
# FDe+fVwu78AX+OYj/6wNpvR7Kb6Jd6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK/fPSMlpAIN48EPkHakSkFv/wztuwSRXK/WiERc1klgMA0GCSqG
# SIb3DQEBAQUABIIBgDqnSDKZcjHFUm73H99yuMYMVOZLDL32zJDz23qXHu3flRzp
# mFhE+Uvr31W9sVnZGNh3MSji3RpzuI7iLOF/2RAWtxopnxVC/7Tzt50uuHHBDTi6
# iXQE13tv7R9Y/u49LgcvVx57Dh3AwQDtKLt6BAU6U+Vv7lRRHDXSIj6Q0S7WdkRq
# eOwhITVeq8fYRtfJnY1rfX/eMAH6SZsX5lCBdezAeTpALzRq9MgyPhtLGP+vQWVp
# WJwzLA8TN/wNsBiTSyV8rliskk1qWtS/5c/GhXdLTnpzw5WU0bj1bKpPr5Qf83eG
# ajGLIiydXqIWMH00c6Cfk8QOyCgLlxoWRuAS7gKRI+Iu73ni0jcAnLQr0nFwkYXk
# CQxNrbHG0WOxSpx8XDZ5VK4wffpsaohV4ARHalGy+V1UPNSLsjbZbpdeCqWb9oOd
# W+v+jhwFlw+W9kgXelv6Aa8UhsNTy0u1BWFzqJ86kXJAyZqv5ZJ43u/4fdq2a0bC
# V3txdZIOkNF5P7JIfqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwOTI3
# NThaMC8GCSqGSIb3DQEJBDEiBCB3gS1Lidm0lNhggX3ufj+pl0GJeFN9olvGacN/
# Sd8QYTANBgkqhkiG9w0BAQEFAASCAgAEE7Kika636hTpjFGOIpxwykxU69MGKjax
# dHJJe3UX7XiXpawIBY4LSayp6HE8kF0FXxBadMFkIoDt6BH9vRdybTtM0RLIRKcE
# Y1EaMs6nGEdAI2mNTlDKYrDQTmkqjlnRr1H1u4ndTQKn+DqlpJEvl6gS1rBSn/hf
# xek2F4e2ezZMhcuG0FOTNo+hFPf12EzoerlZ+e1exdvVb2neYryp/NRt2DBIjFIs
# jmLGPIRlFkFIhTB/tJ7ylRsFa3v80nFs6X0eDa8z0U7OYIHxFcnnDnBEcx8HWsxO
# cyfediEM0KixTr3YtqqpyeGUa9lcufrpcBzboo9R1wwB80iCOCjCpIxdYET765SB
# N8+eTwAvr/Mdt4hEiOKlj6TsWu+PL0VwNHkJsLOH+tLe1BPNyBetSNzB8tmvehyN
# QjpVXjcNBolvaFP9fjZLfpR0TgHM+LaPsJjvxoC8A9lfVUdT+Mx0nXPMtMijmTOH
# v1GekTFQ92h+n17HwmTBdftgBHbkLb9lLQ3tJgGE8MgQ0kLl/79Mlj3b1E4v16fs
# jIrF8VSH8CaposaNI3RLwL4NerwTu01ISbhH7ZBUbqprkFH05q8Yqdf5i2Xqlo90
# 1pCiMaimCd7s/uZNPArwqfU1TNGFLYqrQbfN6i8LFRAi9/K71DFPXusLBBz3HfhG
# WkmPdyV6wQ==
# SIG # End signature block
