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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDGgqiu+4pTknIf
# B5nTuf++e8aYR2giB/wvVfgvskjYN6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICwF7BtnWjXETvsbrceu1HPh8TtGtMsOpVd58AIvdUx2MA0GCSqG
# SIb3DQEBAQUABIIBgH3pBoGhnT1SqjZkeJ3eV5dtZ7WqGB6/yhk0bP42a2ppuEUG
# OmgTpA8op5W9cXvSAR9baFiCY55JxglNRg/L7MvQRNceaGuPbeXZSdGfOObJrOWI
# 5RjOcp/wTX/WOBI2RQ8GDmu4cBrzMRPXICGho0sPHKPba092CuCsRdCv51gIgqkd
# uUZ/ok9I7W8Hw/bcpCJNooiuZ6xfhCbVqEgsYp+Vx+pJHo6ty04LsUmy4vCYZCTC
# Jnvgv1st6UIcx/YH3AGRre+N9lhm5Z8patjMn490K+v8EJ68hAs+4ldJdPJngj3p
# sCnysmxcOG552ySdQAhyLGojeiEm35dwp9XcgtaVcMPMJgWTCFsiai5R1lAUu3vf
# hn11/7ADtB9xZnTXLr1mLGe9szJgvcEGX0U+VuQ7VpcjPqG1oMTcGg++vSRQtWe1
# K9Quh7MK/4GTvlsFD1aTXhhVm5nqxJ2gl+xiTzq402WZf893DjO5F8N5Mi9/1f3X
# 8WibDX5BUB7G+bWalaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkwOTI3
# NTdaMC8GCSqGSIb3DQEJBDEiBCDzAEqEuEYMhing3zr9o7Z+pr5JIE02K5efpqdO
# jxLNUjANBgkqhkiG9w0BAQEFAASCAgA45wpoa+0lTe4rgtcvf7DmesBW0EqjnFwo
# bfLdWkEsWZT8NrAHy6PvIxzwSTyAlNUHUEsGNzziHj3lPMpo0MvWiNORISjv5jlI
# NLDN8HZmXkD7rYUIbOzX/cNrzPem7ST7JeVyR5plw4Fh21dctaUXExOlqaEQJtre
# h/Z9i+9Bkzng04CxqVhqvSFTuEA0C722ta4l7rWUmn6AraxD+/IitKW3Zc+HZtcL
# h9VN5JlensADicaQ18wFpVSCQyYSG51fS6VKbbRaV8Q4ZGr57Q0Llzjffes/SInk
# BqsbojRAKZuyiyd5MmmH0IqloRI/lN1h/BeXKDbuQ45UYSpvbvv5wRkyWO70jUBQ
# +HO1KLQdMosYkMRnXeY1Q1SAVohPxA5i6qSTT34RDXUdDu0trcBnpfCgwa+AfdUr
# ATMGXPPobh48hN5ujrcp8DxnNKFb9OtsHLEWoopUbGw7Nr64l1lxhoeJrlTavplQ
# ynuP1g/cSlXyWWJguU/NP32AGgB7HMHYfxw9EG8w3pfByRsijQmoFD/zmSpWUS1Z
# KZGqmsX7d0p2RL/NHjzj9/c9hxfkE2L50PCL98nTigXk5YwZa8ZVnI7nMFwCopgP
# T5qh+f1Kap5uaZen0gN+RzE+0GwMTdIYrhkKJYhiDy2NpFl8A9vY/Ne/YjHH/HnC
# yxb3G5dmhw==
# SIG # End signature block
