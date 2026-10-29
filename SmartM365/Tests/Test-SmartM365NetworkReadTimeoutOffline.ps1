#Requires -Version 7.0
<#
.SYNOPSIS
Offline checks that a response body that stops arriving ends a network call with a retried, logged error
instead of hanging the collector.
.DESCRIPTION
2026-10-29: an Intune export download and the licensing SharePoint cleanup stayed silent for hours. In
PowerShell 7.4+ -TimeoutSec (ConnectionTimeoutSeconds) bounds a web call only until the response headers
arrive; the read of the body was unbounded.
1. Get-SmartM365WebRequestTimeout adds OperationTimeoutSeconds (NetworkReadIdleTimeoutSeconds, default 120,
   capped by the call timeout) when the command offers it.
2. Real transport: a loopback listener sends the headers and part of the body, then stalls.
   a. Invoke-SmartM365ExportFileDownload: a stalled attempt is logged and retried; the next one downloads.
   b. Invoke-SmartM365ExportFileDownload: stalled on every attempt, it ends with a clear error.
   c. Invoke-SmartM365GraphRestWithRetry and Invoke-SmartM365GraphDeleteQuietly: a stalled answer is retried.
3. Child process with the installed Microsoft.Graph.Authentication (-GraphAuthenticationVersion, default 2.30.0):
   Invoke-SmartM365EvidenceGraphRequest against a stalled body ends at the SDK ClientTimeout with a clear error.
   NOT RUN when that SDK version is absent.
.VERSION
1.0.0
.NOTES
No tenant, Graph, storage, mail or synchronized data access. Tokens and logs are recorded stand-ins.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Network settings are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText','',Justification='Synthetic unsigned token for a loopback listener.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Loopback test server and recorded logs only.')]
param([string]$GraphAuthenticationVersion='2.30.0',[switch]$TransportChild)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$root = Split-Path $PSScriptRoot -Parent

# Loopback HTTP server: each connection takes the next mode ('ok' answers in full); otherwise it stalls after
# the headers and part of the body, and the connection is held open until the server stops.
function Start-StallServer {
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0); $listener.Start()
    $modes = [Collections.Concurrent.ConcurrentQueue[string]]::new()
    $held = [Collections.Concurrent.ConcurrentBag[object]]::new()
    $runspace = [runspacefactory]::CreateRunspace(); $runspace.Open()
    $shell = [powershell]::Create(); $shell.Runspace = $runspace
    $null = $shell.AddScript({ param($Listener, $Modes, $Held)
        while ($true) {
            try { $client = $Listener.AcceptTcpClient() } catch { break }
            $stream = $client.GetStream(); $buffer = [byte[]]::new(16384); [void]$stream.Read($buffer, 0, $buffer.Length)
            $next = $null
            $mode = if ($Modes.TryDequeue([ref]$next)) { $next } else { 'stall' }
            if ($mode -eq 'ok') {
                $body = [Text.Encoding]::ASCII.GetBytes('{"id":"synthetic-ok"}')
                $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: application/json`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
                $stream.Write($head, 0, $head.Length); $stream.Write($body, 0, $body.Length); $stream.Flush(); $client.Close()
            }
            else {
                $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: application/json`r`nContent-Length: 100000`r`n`r`n{`"value`":`"")
                $stream.Write($head, 0, $head.Length); $stream.Flush(); $Held.Add($client)
            }
        } }).AddArgument($listener).AddArgument($modes).AddArgument($held)
    $handle = $shell.BeginInvoke()
    [pscustomobject]@{ Listener = $listener; Modes = $modes; Held = $held; Shell = $shell; Runspace = $runspace; Handle = $handle; Uri = ('http://127.0.0.1:{0}/' -f ([Net.IPEndPoint]$listener.LocalEndpoint).Port) }
}
function Stop-StallServer {
    param($Server)
    $Server.Listener.Stop()
    foreach ($client in $Server.Held) { try { $client.Close() } catch { Write-Verbose 'Connection already closed.' } }
    try { $null = $Server.Shell.EndInvoke($Server.Handle) } catch { Write-Verbose 'Listener stopped.' }
    $Server.Shell.Dispose(); $Server.Runspace.Dispose()
}

if ($TransportChild) {
    # ---- 3. Child process: the SDK client timeout ends a stalled Invoke-MgGraphRequest ----
    Import-Module Microsoft.Graph.Authentication -RequiredVersion $GraphAuthenticationVersion -ErrorAction Stop
    Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.112' -Global -WarningAction SilentlyContinue
    Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.17' -Force
    $server = Start-StallServer
    try {
        $claims = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"tid":"00000000-0000-0000-0000-000000000001","exp":4102444800}')).TrimEnd('=').Replace('+','-').Replace('/','_')
        $token = ConvertTo-SecureString ('eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.' + $claims + '.') -AsPlainText -Force
        Connect-MgGraph -AccessToken $token -NoWelcome | Out-Null
        # As Connect-SmartM365EvidenceGraph -DisableSdkRetry does, with the shortest practical client timeout.
        Set-MgRequestContext -MaxRetry 0 -ClientTimeout 10 | Out-Null
        $watch = [Diagnostics.Stopwatch]::StartNew(); $failure = $null
        try { $null = Invoke-SmartM365EvidenceGraphRequest -Uri ($server.Uri + 'v1.0/deviceManagement/reports/exportJobs/job-1') -MaxRetryCount 1 -Operation 'Synthetic export job status' } catch { $failure = $_.Exception }
        $watch.Stop()
        Check ($null -ne $failure -and $failure.Message -match '^Synthetic export job status failed after 1 attempt\(s\) .*\(a transient error\)\. Last error: .*Timeout') "A stalled Graph body did not end with a clear, transient timeout error: $failure"
        Check ($watch.Elapsed.TotalSeconds -ge 8 -and $watch.Elapsed.TotalSeconds -lt 60) "A stalled Graph body was not bounded by ClientTimeout ($([int]$watch.Elapsed.TotalSeconds)s)."
        "PASS: $script:checks SDK client-timeout checks with Microsoft.Graph.Authentication $GraphAuthenticationVersion."
    }
    finally {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'No Graph context to disconnect.' }
        Stop-StallServer $server
    }
    return
}

$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.112' -Force -PassThru -WarningAction SilentlyContinue
$globalNames = @('NetworkReadIdleTimeoutSeconds', 'GraphRequestTimeoutSeconds', 'SharePointUploadTransferTimeoutSeconds', 'AppId', 'TenantId', 'Thumb', 'Thumbprint')
$saved = @{}
foreach ($name in $globalNames) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-ReadTimeout-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$server = $null
try {
    $global:AppId = 'synthetic-app'; $global:TenantId = 'synthetic-tenant'; $global:Thumb = 'SYNTHETIC'; $global:Thumbprint = 'SYNTHETIC'
    $global:GraphRequestTimeoutSeconds = $null; $global:SharePointUploadTransferTimeoutSeconds = $null

    # ---- 1. Timeout parameters ----
    $global:NetworkReadIdleTimeoutSeconds = $null
    $default = & $core { Get-SmartM365WebRequestTimeout -Command 'Invoke-WebRequest' -TimeoutSeconds 900 }
    Check ($default.TimeoutSec -eq 900 -and $default.OperationTimeoutSeconds -eq 120) "The default read idle timeout is not 120 s: $($default | ConvertTo-Json -Compress)"
    $global:NetworkReadIdleTimeoutSeconds = 5
    $short = & $core { Get-SmartM365WebRequestTimeout -Command 'Invoke-RestMethod' -TimeoutSeconds 3 }
    $configured = & $core { Get-SmartM365WebRequestTimeout -Command 'Invoke-RestMethod' -TimeoutSeconds 180 }
    Check ($short.OperationTimeoutSeconds -eq 3 -and $configured.TimeoutSec -eq 180 -and $configured.OperationTimeoutSeconds -eq 5) 'The read idle timeout is not configurable or not capped by the call timeout.'
    $offered = & $core {
        function script:Invoke-WebRequest { param($Uri, $TimeoutSec) }
        try { Get-SmartM365WebRequestTimeout -Command 'Invoke-WebRequest' -TimeoutSeconds 60 } finally { Remove-Item -Path Function:\Invoke-WebRequest }
    }
    Check (-not $offered.ContainsKey('OperationTimeoutSeconds') -and $offered.TimeoutSec -eq 60) 'OperationTimeoutSeconds is passed to a command that does not offer it.'

    # ---- 2. Real transport against a stalled body ----
    $server = Start-StallServer
    & $core {
        $script:Logs = [Collections.Generic.List[object]]::new()
        function script:WriteLog { param($Message, $Level) $script:Logs.Add([pscustomobject]@{ Level = [string]$Level; Message = [string]$Message }) }
        function script:Start-Sleep { param([int]$Seconds) $script:Logs.Add([pscustomobject]@{ Level = 'SLEEP'; Message = [string]$Seconds }) }
        $script:SmartM365NetworkDeadline = $null
        $script:SmartM365GraphRestTokenCache = @{ 'synthetic-app|synthetic-tenant|SYNTHETIC' = [pscustomobject]@{ AccessToken = 'synthetic'; ExpiresOn = (Get-Date).ToUniversalTime().AddHours(1) } }
    }
    function Reset-Log { & $core { $script:Logs.Clear() } }
    function Get-Log { & $core { $script:Logs.ToArray() } }

    # a. Stalled once, then answered: retried with a WARNING, the file is downloaded.
    Reset-Log; $server.Modes.Enqueue('stall'); $server.Modes.Enqueue('ok')
    $out = Join-Path $temporary 'export-a.zip'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $result = & $core { param($Uri, $Out) Invoke-SmartM365ExportFileDownload -Url ($Uri + 'export.zip?sig=SECRET') -OutFile $Out -Operation 'Synthetic export download' -MaxAttempts 2 } $server.Uri $out
    $watch.Stop()
    Check ($result -eq $out -and (Get-Content -LiteralPath $out -Raw) -eq '{"id":"synthetic-ok"}') 'The download did not recover after a stalled attempt.'
    Check ($watch.Elapsed.TotalSeconds -ge 4 -and $watch.Elapsed.TotalSeconds -lt 30) "The stalled download attempt was not bounded by the read idle timeout ($([int]$watch.Elapsed.TotalSeconds)s)."
    $warnings = @(Get-Log | Where-Object Level -eq 'WARNING')
    Check ($warnings.Count -eq 1 -and $warnings[0].Message -match 'Synthetic export download attempt 1/2 failed\..*OperationTimeout' -and $warnings[0].Message -notmatch 'SECRET') "The stalled download attempt was not logged as a retried timeout: $($warnings.Message -join ' | ')"
    Check (@(Get-ChildItem -LiteralPath $temporary -Filter '*.partial').Count -eq 0) 'A partial download file was left behind.'

    # b. Stalled on every attempt: a clear error after the attempts, no hang.
    Reset-Log
    $watch = [Diagnostics.Stopwatch]::StartNew(); $failure = $null
    try { & $core { param($Uri, $Out) Invoke-SmartM365ExportFileDownload -Url ($Uri + 'export.zip?sig=SECRET') -OutFile $Out -Operation 'Synthetic export download' -MaxAttempts 2 } $server.Uri (Join-Path $temporary 'export-b.zip') } catch { $failure = $_.Exception }
    $watch.Stop()
    Check ($null -ne $failure -and $failure -is [IO.IOException] -and $failure.Message -match '^Synthetic export download failed after 2 attempt\(s\)\..*OperationTimeout' -and $failure.Message -notmatch 'SECRET') "A download stalled on every attempt did not end with a clear error: $failure"
    Check ($watch.Elapsed.TotalSeconds -ge 8 -and $watch.Elapsed.TotalSeconds -lt 40) "Stalled download attempts were not bounded ($([int]$watch.Elapsed.TotalSeconds)s)."
    Check (-not (Test-Path -LiteralPath (Join-Path $temporary 'export-b.zip'))) 'A failed download left an output file.'

    # c. Graph REST request and SharePoint delete: a stalled answer is retried, the next one succeeds.
    Reset-Log; $server.Modes.Enqueue('stall'); $server.Modes.Enqueue('ok')
    $reply = & $core { param($Uri) Invoke-SmartM365GraphRestWithRetry -Method GET -Uri ($Uri + 'v1.0/sites/synthetic') -Operation 'Synthetic Graph request' -MaxAttempts 2 -DefaultRetrySeconds 1 } $server.Uri
    Check ($reply.id -eq 'synthetic-ok') 'A stalled Graph REST answer was not retried to success.'
    Check (@(Get-Log | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -match 'Synthetic Graph request transient failure.*attempt 1/2' }).Count -eq 1) 'The stalled Graph REST attempt was not logged.'
    Reset-Log; $server.Modes.Enqueue('stall'); $server.Modes.Enqueue('ok')
    $delete = & $core { param($Uri) Invoke-SmartM365GraphDeleteQuietly -Uri ($Uri + 'v1.0/drives/d/items/i') -Operation 'Synthetic SharePoint delete' -MaxAttempts 2 -DefaultRetrySeconds 1 } $server.Uri
    Check ($delete.Success -and -not $delete.NotFound) "A stalled SharePoint delete answer was not retried to success: $($delete | ConvertTo-Json -Compress)"
    Check (@(Get-Log | Where-Object { $_.Level -eq 'WARNING' -and $_.Message -match 'Synthetic SharePoint delete' }).Count -ge 1) 'The stalled SharePoint delete attempt was not logged.'
}
finally {
    if ($server) { Stop-StallServer $server }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    Remove-Module $core -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}

# ---- 3. SDK client timeout in a child process ----
$transport = "NOT RUN: Microsoft.Graph.Authentication $GraphAuthenticationVersion is not installed."
if (@(Get-Module -ListAvailable Microsoft.Graph.Authentication | Where-Object { $_.Version -eq [version]$GraphAuthenticationVersion }).Count) {
    $output = @(& (Join-Path $PSHOME 'pwsh') -NoProfile -NonInteractive -File $PSCommandPath -TransportChild -GraphAuthenticationVersion $GraphAuthenticationVersion 2>&1)
    $transport = [string]($output | Select-Object -Last 1)
    Check ($LASTEXITCODE -eq 0 -and $transport -match '^(PASS|NOT RUN):') "SDK client-timeout replay failed: $($output -join ' | ')"
}
"PASS: $script:checks network read timeout checks; PowerShell $($PSVersionTable.PSVersion). SDK client timeout: $transport"
