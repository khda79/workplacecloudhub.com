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
