<#
.SYNOPSIS
    Real-transport retention check: the installed Invoke-MgGraphRequest against a local HTTP listener.
.DESCRIPTION
    Uses a synthetic access token and a localhost listener. No tenant, Graph endpoint or network
    access outside this machine. Requires PowerShell 7 and Microsoft.Graph.Authentication; reports
    NOT RUN with exit code 3 when the SDK is absent instead of passing.
    -GraphAuthenticationVersion selects an installed SDK version (for example 2.30.0, the server
    version). When the SDK session type is reachable (2.30.0), the 2026-10-22 incident is replayed:
    after an absolute-URI request the session is switched to an app-only certificate context with
    an absent thumbprint, so a client rebuild fails exactly as on the server, without any network.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param([string]$GraphAuthenticationVersion)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'PowerShell 7 is required.'}
$sdk=@(Get-Module -ListAvailable Microsoft.Graph.Authentication | Where-Object {-not $GraphAuthenticationVersion -or $_.Version -eq [version]$GraphAuthenticationVersion})
if(-not $sdk.Count){Write-Host "NOT RUN: Microsoft.Graph.Authentication $GraphAuthenticationVersion is not installed.";exit 3}
if($GraphAuthenticationVersion){Import-Module Microsoft.Graph.Authentication -RequiredVersion $GraphAuthenticationVersion -ErrorAction Stop}
else{Import-Module Microsoft.Graph.Authentication -ErrorAction Stop}
$modulePath=Join-Path $PSScriptRoot 'PreparedSharePointRetention.psm1'
Import-Module $modulePath -Force
$module=Get-Module PreparedSharePointRetention
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}

$port=Get-Random -Minimum 20000 -Maximum 45000
$prefix="http://localhost:$port/"
$listener=[Net.HttpListener]::new();$listener.Prefixes.Add($prefix);$listener.Start()
$responses=[Collections.Concurrent.ConcurrentQueue[int]]::new()
$seen=[Collections.Concurrent.ConcurrentQueue[object]]::new()
$runspace=[runspacefactory]::CreateRunspace();$runspace.Open()
$server=[powershell]::Create();$server.Runspace=$runspace
$null=$server.AddScript({param($Listener,$Responses,$Seen)
    while($Listener.IsListening){
        try {$context=$Listener.GetContext()} catch {break}
        $status=0;if(-not $Responses.TryDequeue([ref]$status)){$status=500}
        $Seen.Enqueue([pscustomobject]@{Method=$context.Request.HttpMethod;Path=$context.Request.RawUrl;IfMatch=$context.Request.Headers['If-Match']})
        $context.Response.StatusCode=$status
        if($status -eq 200){
            $bytes=[Text.Encoding]::UTF8.GetBytes('{"id":"item"}');$context.Response.ContentType='application/json'
            $context.Response.OutputStream.Write($bytes,0,$bytes.Length)
        }
        $context.Response.Close()
    }
}).AddArgument($listener).AddArgument($responses).AddArgument($seen)
$handle=$server.BeginInvoke()
function Get-NextRequest {$item=$null;if($seen.TryDequeue([ref]$item)){$item}else{$null}}
function Invoke-Delete([string]$Uri,[string]$ETag){& $module {param($u,$e) Invoke-RetentionGraphDelete $u $e} $Uri $ETag}
try {
    # Synthetic unsigned token: Connect-MgGraph -AccessToken performs no network call.
    $claims=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"tid":"00000000-0000-0000-0000-000000000001","exp":4102444800}')).TrimEnd('=').Replace('+','-').Replace('/','_')
    $token=ConvertTo-SecureString ('eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.'+$claims+'.') -AsPlainText -Force
    Connect-MgGraph -AccessToken $token | Out-Null
    # Test setup only: one listener response per request, so SDK retries must not consume queued statuses.
    Set-MgRequestContext -MaxRetry 0 | Out-Null
    $version=(Get-Module Microsoft.Graph.Authentication).Version
    & $module {Assert-RetentionGraphDeleteSupport}
    Check $true 'Installed Invoke-MgGraphRequest exposes the retention DELETE parameters'
    $etag='"{6A5B1C2D-0000-0000-0000-000000000001},3"'
    $uri=$prefix+'v1.0/drives/b%21synthetic/items/01SYNTHETICITEM'
    # Same order as the incident run: an absolute-URI GET first, then the DELETE, without touching the request context.
    $responses.Enqueue(200)
    $null=Invoke-MgGraphRequest -Method GET -Uri ($prefix+'v1.0/drives/b%21synthetic/items/01SYNTHETICITEM') -OutputType PSObject -ErrorAction Stop
    $null=Get-NextRequest
    $environment='unavailable'
    try {$environment=[string][Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.Environment.AzureADEndpoint} catch {Write-Verbose 'GraphSession is not reachable from this load context.'}
    foreach($case in @(
        @{Status=204;Outcome='Recycled';Reached=$false},
        @{Status=404;Outcome='NotFound';Reached=$false},
        @{Status=412;Outcome='PreconditionFailed';Reached=$false},
        @{Status=429;Outcome='HttpError';Reached=$false},
        @{Status=500;Outcome='HttpError';Reached=$true})){
        $responses.Enqueue($case.Status)
        $result=Invoke-Delete $uri $etag
        $request=Get-NextRequest
        Check ($null -ne $request -and $request.Method -ceq 'DELETE' -and $request.Path -ceq '/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM' -and $request.IfMatch -ceq $etag) "HTTP $($case.Status): the real cmdlet did not send the exact conditional DELETE"
        Check ($result.Outcome -ceq $case.Outcome -and $result.HttpStatus -eq $case.Status -and $result.MayHaveReachedService -eq $case.Reached) "HTTP $($case.Status): -StatusCodeVariable outcome was $($result.Outcome)/$($result.HttpStatus) instead of $($case.Outcome)/$($case.Status)"
    }
    # Regression: a URI the SDK cannot use fails before sending, is reported explicitly and is never 'in flight'.
    # A relative string is unsafe here: the SDK resolves it against the host of the previous absolute request.
    $result=Invoke-Delete 'https://graph.microsoft.com:99999/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM' $etag
    Check ($result.Outcome -ceq 'RequestFailed' -and -not $result.MayHaveReachedService -and $result.Error -match 'UriFormatException|Invalid URI|URI') 'Invalid URI was not reported as a pre-send failure'
    Check ($null -eq (Get-NextRequest)) 'Invalid URI request reached the listener'
    # A refused connection stays uncertain: the module cannot prove the request was not accepted.
    $closed=Get-Random -Minimum 45001 -Maximum 60000
    $result=Invoke-Delete ("http://localhost:$closed/v1.0/drives/b%21synthetic/items/01SYNTHETICITEM") $etag
    Check ($result.Outcome -ceq 'RequestFailed' -and $result.MayHaveReachedService) 'Connection failure was not kept uncertain'
    # Incident replay, last because it leaves the SDK session unusable.
    $sessionType='Microsoft.Graph.PowerShell.Authentication.GraphSession' -as [type]
    $replay='not run (SDK session type not reachable in this version)'
    if($sessionType){
        $session=$sessionType::Instance
        Check ([string]::IsNullOrEmpty([string]$session.Environment.AzureADEndpoint)) 'Expected the SDK session environment to lose its authority after an absolute-URI request'
        $context=$session.AuthContext
        $context.AuthType=[Microsoft.Graph.PowerShell.Authentication.AuthenticationType]::AppOnly
        $context.TokenCredentialType=[Microsoft.Graph.PowerShell.Authentication.TokenCredentialType]::ClientCertificate
        $context.ClientId='00000000-0000-0000-0000-0000000000aa';$context.CertificateThumbprint=('0'*40)
        # Fixed path: the cached client is kept, so the DELETE is sent.
        $responses.Enqueue(204)
        $result=Invoke-Delete $uri $etag
        Check ($result.Outcome -ceq 'Recycled' -and $null -ne (Get-NextRequest)) 'Fixed DELETE path did not reach the listener after the session lost its authority'
        # Former path: a request-context change discards the client; the rebuild must fail before sending.
        Set-MgRequestContext -MaxRetry 0 | Out-Null
        $result=Invoke-Delete $uri $etag
        Check ($result.Outcome -ceq 'RequestFailed' -and -not $result.MayHaveReachedService -and $result.Error -match 'UriFormatException: Invalid URI') "Incident replay did not reproduce a pre-send Invalid URI: $($result.Error)"
        Check ($null -eq (Get-NextRequest)) 'Incident replay sent a request'
        $replay='reproduced: Set-MgRequestContext + client rebuild raises Invalid URI before sending; the fixed path sends the DELETE'
    }
    Write-Host "PASS: $script:checks real-transport checks with Microsoft.Graph.Authentication $version (synthetic token, localhost listener). Session AzureADEndpoint after an absolute-URI request: '$environment'."
    Write-Host "Incident replay: $replay."
} finally {
    try {Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null} catch {Write-Verbose 'No Graph context to disconnect.'}
    $listener.Stop();$listener.Close()
    try {$null=$server.EndInvoke($handle)} catch {Write-Verbose 'Listener stopped.'}
    $server.Dispose();$runspace.Dispose()
}
