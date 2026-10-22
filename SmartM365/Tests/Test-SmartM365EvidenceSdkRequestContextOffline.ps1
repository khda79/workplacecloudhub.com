#Requires -Version 7.0
<#
.SYNOPSIS
Regression checks: the Graph SDK request context is changed only at connection, never on the export paths.
.DESCRIPTION
With Microsoft.Graph.Authentication 2.30 an Invoke-MgGraphRequest with an absolute URI empties the session's
AzureADEndpoint. Set-MgRequestContext then drops the cached client, and its rebuild for app-only certificate
auth fails with "Invalid URI" before anything is sent (endpoint-security exports, 2026-10-22).
1. Recorded stand-ins: Connect-SmartM365EvidenceGraph -DisableSdkRetry and three budgeted exports, as the
   endpoint-security collector runs them. Set-MgRequestContext must be called once, before the first request.
2. Static scan: Set-MgRequestContext appears in SmartM365 runtime code only inside Connect-SmartM365EvidenceGraph,
   and the endpoint-security collector connects with -DisableSdkRetry.
3. Real transport (child pwsh): the installed SDK (-GraphAuthenticationVersion, default 2.30.0) against a localhost
   listener with a synthetic token. The fixed order sends every request; the former order (a request-context
   change after an absolute-URI request) fails before sending. NOT RUN when that SDK version is absent.
.VERSION
1.0.0
.NOTES
No tenant, Graph, storage, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Stand-ins accept the parameters of the commands they replace.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Import-Module and the waits are recorded stand-ins inside the module.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText','',Justification='Synthetic unsigned token for a localhost listener.')]
param([string]$GraphAuthenticationVersion='2.30.0',[switch]$TransportChild)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

if ($TransportChild) {
    # ---- 3. Child process: the SDK session is left unusable by the former-order replay ----
    Import-Module Microsoft.Graph.Authentication -RequiredVersion $GraphAuthenticationVersion -ErrorAction Stop
    $sessionType = 'Microsoft.Graph.PowerShell.Authentication.GraphSession' -as [type]
    if (-not $sessionType) { 'NOT RUN: the SDK session type is not reachable in this version.'; return }
    $port = Get-Random -Minimum 20000 -Maximum 45000; $prefix = "http://localhost:$port/"
    $listener = [Net.HttpListener]::new(); $listener.Prefixes.Add($prefix); $listener.Start()
    $seen = [Collections.Concurrent.ConcurrentQueue[string]]::new()
    $runspace = [runspacefactory]::CreateRunspace(); $runspace.Open(); $server = [powershell]::Create(); $server.Runspace = $runspace
    $null = $server.AddScript({ param($Listener, $Seen)
        while ($Listener.IsListening) {
            try { $context = $Listener.GetContext() } catch { break }
            $Seen.Enqueue("$($context.Request.HttpMethod) $($context.Request.RawUrl)")
            $bytes = [Text.Encoding]::UTF8.GetBytes('{"id":"job-1","status":"completed"}')
            $context.Response.StatusCode = 200; $context.Response.ContentType = 'application/json'
            $context.Response.OutputStream.Write($bytes, 0, $bytes.Length); $context.Response.Close()
        } }).AddArgument($listener).AddArgument($seen)
    $handle = $server.BeginInvoke()
    try {
        $claims = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"tid":"00000000-0000-0000-0000-000000000001","exp":4102444800}')).TrimEnd('=').Replace('+','-').Replace('/','_')
        $token = ConvertTo-SecureString ('eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.' + $claims + '.') -AsPlainText -Force
        $uri = $prefix + 'v1.0/deviceManagement/reports/exportJobs'
        function Send { param([string]$Method, [string]$Uri) if ($Method -eq 'POST') { $null = Invoke-MgGraphRequest -Method POST -Uri $Uri -Body '{}' -ContentType 'application/json' -OutputType PSObject } else { $null = Invoke-MgGraphRequest -Uri $Uri -OutputType PSObject } }
        function Use-CertificateContext {
            # Production auth type with an absent thumbprint: a client rebuild must build the authority URI first.
            $context = $sessionType::Instance.AuthContext
            $context.AuthType = [Microsoft.Graph.PowerShell.Authentication.AuthenticationType]::AppOnly
            $context.TokenCredentialType = [Microsoft.Graph.PowerShell.Authentication.TokenCredentialType]::ClientCertificate
            $context.ClientId = '00000000-0000-0000-0000-0000000000aa'; $context.CertificateThumbprint = ('0' * 40)
        }
        # Fixed order: request context set right after connecting, then three reports without touching it.
        Connect-MgGraph -AccessToken $token -NoWelcome | Out-Null
        Set-MgRequestContext -MaxRetry 0 | Out-Null
        Send POST $uri; Send GET "$uri/job-1"
        Check ([string]::IsNullOrEmpty([string]$sessionType::Instance.Environment.AzureADEndpoint)) 'Premise: an absolute-URI request no longer empties the session AzureADEndpoint (review the SDK version).'
        Use-CertificateContext
        foreach ($report in 2, 3) { Send POST $uri; Send GET "$uri/job-1" }
        Check ($seen.Count -eq 6 -and (Get-MgRequestContext).MaxRetry -eq 0) "Fixed order: $($seen.Count) of 6 requests reached the listener."
        # Former order: a request-context change after an absolute-URI request; the rebuild fails before sending.
        Set-MgRequestContext -MaxRetry 3 | Out-Null
        $failure = $null; try { Send POST $uri } catch { $failure = $_.Exception }
        Check ($null -ne $failure -and $failure.GetType().Name -eq 'UriFormatException' -and $seen.Count -eq 6) "Former order did not fail with Invalid URI before sending: $failure"
        "PASS: $script:checks real-transport checks with Microsoft.Graph.Authentication $GraphAuthenticationVersion."
    }
    finally {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'No Graph context to disconnect.' }
        $listener.Stop(); $listener.Close()
        try { $null = $server.EndInvoke($handle) } catch { Write-Verbose 'Listener stopped.' }
        $server.Dispose(); $runspace.Dispose()
    }
    return
}

$root = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.103' -Force -Global -WarningAction SilentlyContinue
$evidence = Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.14' -Force -PassThru
$collector = Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/Security/SmartM365-Intune-EndpointSecurityHealth-Inventory.ps1'
try {
    # ---- 1. Recorded stand-ins: connection, then three budgeted exports ----
    $events = & $evidence {
        $script:Events = [Collections.Generic.List[string]]::new(); $script:MaxRetry = 3
        function script:Import-Module { param([Parameter(ValueFromRemainingArguments)]$Arguments) }
        function script:Get-MgContext { param([Parameter(ValueFromRemainingArguments)]$Arguments) $null }
        function script:Disconnect-MgGraph { param([Parameter(ValueFromRemainingArguments)]$Arguments) }
        function script:Connect-MgGraph { param([Parameter(ValueFromRemainingArguments)]$Arguments) $script:Events.Add('Connect') }
        function script:Set-SmartM365InventoryConnectionOwnership { param($Connection, $Owned) }
        function script:Get-MgRequestContext { [pscustomobject]@{ MaxRetry = $script:MaxRetry } }
        function script:Set-MgRequestContext { param([int]$MaxRetry) $script:MaxRetry = $MaxRetry; $script:Events.Add("Set:$MaxRetry") }
        function script:WriteLog { param($Message, $Level) if ($Level -eq 'WARNING') { $script:Events.Add("Warning:$Message") } }
        function script:Start-Sleep { param([int]$Seconds) }
        function script:Invoke-MgGraphRequest {
            [CmdletBinding()] param($Method, $Uri, $Body, $ContentType, $OutputType)
            $script:Events.Add("Request:$Method")
            if ($Method -eq 'POST') { [pscustomobject]@{ id = 'job-1' } } else { [pscustomobject]@{ status = 'completed'; url = 'https://intunestorage.example.invalid/r.zip?sig=S' } }
        }
        function script:Invoke-SmartM365ExportFileDownload {
            param($Url, $OutFile, $Operation)
            $folder = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-SdkContextZip-' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $folder | Out-Null
            try { [IO.File]::WriteAllText((Join-Path $folder 'r.csv'), "DeviceId,DeviceName`r`nd1,PC1"); Compress-Archive -Path (Join-Path $folder 'r.csv') -DestinationPath $OutFile -Force } finally { Remove-Item -LiteralPath $folder -Recurse -Force }
            $OutFile
        }
        $runtime = [pscustomobject]@{ Config = [pscustomobject]@{ AppId = '00000000-0000-0000-0000-0000000000aa'; TenantId = '00000000-0000-0000-0000-000000000001'; Thumbprint = ('0' * 40) } }
        Connect-SmartM365EvidenceGraph $runtime @('DeviceManagementManagedDevices.ReadWrite.All') -DisableSdkRetry
        foreach ($report in 'DefenderAgents', 'FirewallStatus', 'FirewallUnhealthyStatus') {
            $null = @(Invoke-SmartM365EvidenceExportReport -ReportName $report -PollSeconds 1 -CreateRetryBudgetSeconds 1080 -StallSeconds 600)
        }
        $script:Events.ToArray()
    }
    $sets = @(for ($i = 0; $i -lt $events.Count; $i++) { if ($events[$i] -like 'Set:*') { $i } })
    $requests = @(for ($i = 0; $i -lt $events.Count; $i++) { if ($events[$i] -like 'Request:*') { $i } })
    Check ($events[0] -eq 'Connect' -and $sets.Count -eq 1 -and $events[$sets[0]] -eq 'Set:0') "The SDK retry is not turned off exactly once at connection: $($events -join ', ')"
    Check ($requests.Count -eq 6 -and $sets[0] -lt $requests[0]) "Set-MgRequestContext runs after a Graph request: $($events -join ', ')"
    Check (@($events | Where-Object { $_ -like 'Warning:*' }).Count -eq 0) "Budgeted exports warned although the SDK retry was off: $($events -join ', ')"

    # Without -DisableSdkRetry the connection leaves the request context alone (other collectors).
    $plain = & $evidence {
        $script:Events.Clear(); $script:MaxRetry = 3
        Connect-SmartM365EvidenceGraph ([pscustomobject]@{ Config = [pscustomobject]@{ AppId = 'a'; TenantId = 't'; Thumbprint = 'x' } }) @('User.Read.All')
        $script:Events.ToArray()
    }
    Check (($plain -join ',') -eq 'Connect') "A connection without -DisableSdkRetry changed the request context: $($plain -join ', ')"

    # ---- 2. Static scan of the SmartM365 runtime code ----
    $files = @(Get-ChildItem -LiteralPath $root -Recurse -File -Include '*.ps1', '*.psm1' | Where-Object { $_.FullName -notmatch '[\\/]Tests[\\/]' })
    $offenders = foreach ($file in $files) {
        $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
        foreach ($command in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Set-MgRequestContext' }, $true)) {
            $owner = $command.Parent; while ($owner -and $owner -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $owner = $owner.Parent }
            if (-not ($file.Name -eq 'SmartM365.EvidenceCollector.Common.psm1' -and $owner -and $owner.Name -eq 'Connect-SmartM365EvidenceGraph')) { "$($file.Name):$($command.Extent.StartLineNumber)" }
        }
    }
    Check (@($offenders).Count -eq 0) "Set-MgRequestContext outside Connect-SmartM365EvidenceGraph: $(@($offenders) -join ', ')"
    $collectorAst = [Management.Automation.Language.Parser]::ParseFile($collector, [ref]$null, [ref]$null)
    $connects = @($collectorAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Connect-SmartM365EvidenceGraph' }, $true))
    Check ($connects.Count -eq 1 -and @($connects[0].CommandElements | Where-Object { $_ -is [Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'DisableSdkRetry' }).Count -eq 1) 'The endpoint-security collector does not connect with -DisableSdkRetry.'
}
finally {
    Remove-Module $evidence -Force -ErrorAction SilentlyContinue
}

# ---- 3. Real transport in a child process ----
$transport = "NOT RUN: Microsoft.Graph.Authentication $GraphAuthenticationVersion is not installed."
if (@(Get-Module -ListAvailable Microsoft.Graph.Authentication | Where-Object { $_.Version -eq [version]$GraphAuthenticationVersion }).Count) {
    $output = @(& (Join-Path $PSHOME 'pwsh') -NoProfile -NonInteractive -File $PSCommandPath -TransportChild -GraphAuthenticationVersion $GraphAuthenticationVersion 2>&1)
    $transport = [string]($output | Select-Object -Last 1)
    Check ($LASTEXITCODE -eq 0 -and $transport -match '^(PASS|NOT RUN):') "Real-transport replay failed: $($output -join ' | ')"
}
"PASS: $script:checks SDK request-context checks; PowerShell $($PSVersionTable.PSVersion). Real transport: $transport"

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCoWxaEw4i0fMjh
# ra81W/bwn3xGcZwXvBuEK1J8j5eH5aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMac/4m1EU2fhH0EM5gMxTiW4giAaymTHGAa8SbWfAO7MA0GCSqG
# SIb3DQEBAQUABIIBgKfGtPZ3DGebWVqR6jLHhBpNH2z0qyxx+V5jcsHbvtzKVbOc
# 6nON5uRrYVhkpDdHs5GOcxBXA0fRRfmH2iKrmbLlH2VgVi83sp41OqutDNqLeek6
# 80vsnL1uTLV1rUspXd5+QnUrSpg/Lr5gmDTj2XSo/yh95/A+kJUZ6nXKN68KHKmm
# pDhzc/caPLj8R5dl3SK5Hfwk6GZ5QXlItO/a/EhuOlLO8dbQIBexaU/VC7ZprrRl
# kVImPsHD/2vWFPEVpcl7B1DSqyWS7dlcJXpoFmle5+4KjW4Tf4XOtXUs4zs8ra3d
# KM/jV4cxmlrZzA3MuSuyQ2vszhxTMaTwEt5ZuyKeEa1+6DO1hVMQfMHFjmNkIOW6
# xmKgt8M6HHghR0wp/ykt5RZ+wjnWACxLyb9a5E29kInYkeFXDtyCJf6zskNWTKau
# OG1ZoBWbZ3usHkjUBH9Eeh2IGeHyU+XCGAiMeHz13DoSLG1g2qX2HYwTwdmNHH2a
# +hozhsEb+cvhN4c1IaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNzI5
# NTlaMC8GCSqGSIb3DQEJBDEiBCDSnnIZduI5Y/euwGs3YBgVbthslxjxWYxrcQSU
# adNiJjANBgkqhkiG9w0BAQEFAASCAgClP0Ncqn3L1lXWH2xNKAke1uD7H7/pO3Y8
# kWdTdg4rYZjHOdf7EWPKmPRUWa1JP4//3Pqy1FavYbqN6I6I9vm4qBk9sWXvosKk
# v7O1HXNWSrqi5/WT6cf/WOhPf1Ot96S1rIZp09ahJE6zMJaFDdtHNfoII+ATrY7u
# p10nvnDEvA+3ktnOehw6xAJcOWjRoa2G/FcibGeaktYROfhJ7mpc8C6IUAKc9N6Q
# obKII0HFKiNk9nKj3Rq5D6BDS04FEcjmzQKZSJVJQCy1Xa73+QVtQDoAENfnBp9n
# TDmiImBqiPLvZtAYEtCpD+a2XYTev/8i7QlV7AAUQZLb+7A+1EGsw6XoucDjNrpW
# Fr2/2QnaRIiunqH5wCsH/OIqLBMrUW5TQ89F/wkDzCWCq/t+o6JVePfWlq2eYJUd
# 9/5NxatHc1GYf7Pu8Z+bk1eL3GIMqEU5Qfyl+R51RmmpzV6h+S2xQbqx4GvwB6Yu
# fmqSHxr5SI5Vh2UF5Chnd7sXVca6Uuyp4tXpYQHKZ18AyrBZJrsALv6gQSZy+ZG6
# m0IcN/CLq6mfohNKnaWkdLkVAdUBwTJn+bujZd798OsyYPnHTfjfiVqEwkvTC8tp
# qgfQKn3GQzHgZLlV9spqFpzMIHfkV0vMc22pvIbIe9EdySxbV7ifqd4gzLZ8CN3b
# Z7FQ9lpVuw==
# SIG # End signature block
