<#
.SYNOPSIS
Offline tests of the Exchange hybrid free/busy collector and its runner.

.DESCRIPTION
Extracts functions from the collector and the runner without running them. Microsoft's Free/Busy
Checker is never executed; Exchange Online and module installation are mocked. Covers:
- work-folder checks identical to the Exchange infrastructure collector, integrity identical in the runner;
- vendored file integrity (valid, SHA256 mismatch, non-Microsoft signer) and report structure still parsed;
- text-only HTML parsing: sections, checks, statuses, masking, command cells and headings skipped;
- NotAuthorized checks, summary rows and run status for every outcome;
- runner: Exchange Management Shell loaded first (ExchangeShellNotLoaded otherwise), no interactive sign-in,
  app-only certificate session proven, CurrentUser module installation, stand-ins for Exchange Online cmdlets
  the app cannot run,
  Free/Busy Checker started with -SkipVersionCheck only, nothing deserialized;
- child process: environment hand-off, time limit with process tree stop;
- Free/Busy Checker run without the runner's strict mode and stop-on-error, which the runner keeps;
- report: title and status in the mail banner, runs that did not happen never shown as healthy;
- OAuth only (Intra Organization Connector enabled): DAuth sides NotApplicable, not counted against the run.

.VERSION
1.0.7
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Mocks keep the parameters of the commands they replace.')]
param([string]$SourceRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$folder = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/Hybrid/FreeBusy'
$collectorPath = Join-Path $folder 'SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1'
$runnerPath = Join-Path $folder 'SmartM365-Exchange-Hybrid-FreeBusy-Runner.ps1'
$vendoredPath = Join-Path $folder 'ThirdParty/CSS-Exchange/FreeBusyChecker/FreeBusyChecker.ps1'
$infrastructurePath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$nonMicrosoftSignedPath = Join-Path (Split-Path $SourceRoot -Parent) 'Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1'

function Get-ScriptAst {
    param([string]$Path)
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Parse failed: $Path $parseErrors" }
    return $ast
}
function Get-FunctionText {
    param($Ast, [string]$Name)
    $node = $Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $node) { throw "Missing function: $Name" }
    return $node.Extent.Text
}
$collectorAst = Get-ScriptAst $collectorPath
$runnerAst = Get-ScriptAst $runnerPath
$infrastructureAst = Get-ScriptAst $infrastructurePath

$script:Checks = 0
function Assert-FreeBusy {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$constants = @{}
foreach ($name in @('FreeBusyCheckerVersion', 'FreeBusyCheckerExpectedSha256', 'FreeBusyCheckerRelativePath', 'FreeBusyRunnerFileName')) {
    $assignment = $collectorAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$name" }, $true)
    if ($null -eq $assignment) { throw "Missing collector constant: $name" }
    $constants[$name] = $assignment.Right.Extent.Text.Trim("'", '"')
}

$collectorFunctions = @('Test-FreeBusyOAuthOnly', 'Ensure-Directory', 'Format-HtmlValue', 'Get-FreeBusyCsvColumn', 'Get-FreeBusyCmdletScope', 'Get-FreeBusySetting', 'Test-FreeBusyCheckerIntegrity',
    'Test-FreeBusyRunnerIntegrity', 'Test-FreeBusyElevation', 'Get-HealthCheckerUnsafeAclEntry', 'Get-HealthCheckerWorkFolderAncestor', 'Set-HealthCheckerWorkFolderAclRule',
    'Initialize-HealthCheckerWorkFolder', 'Invoke-FreeBusyRunnerProcess', 'Read-FreeBusyRunnerStatus', 'ConvertTo-FreeBusySafeValue', 'Get-FreeBusyMaskedItem',
    'ConvertFrom-FreeBusyCheckerHtml', 'Add-FreeBusyNotAuthorizedFinding', 'New-FreeBusySummaryRow', 'Resolve-FreeBusyRunStatus', 'ConvertTo-FreeBusyReportHtml', 'Clear-FreeBusyWorkFolder')
$definitions = @("`$FreeBusyCheckerExpectedSha256 = '$($constants.FreeBusyCheckerExpectedSha256)'", "`$FreeBusyCheckerVersion = '$($constants.FreeBusyCheckerVersion)'", "`$RunId = '20261025-050000'", "`$FreeBusyDAuthNotApplicableReason = 'Intra Organization Connector enabled: free/busy uses OAuth, DAuth not checked'") + @($collectorFunctions | ForEach-Object { Get-FunctionText $collectorAst $_ })
# The collector runs under Set-StrictMode -Version Latest: its functions are tested the same way.
$module = New-Module -ScriptBlock ([scriptblock]::Create(((@('Set-StrictMode -Version Latest') + $definitions) -join "`n")))
& $module {
    Set-StrictMode -Version Latest
    $script:LogLines = New-Object 'System.Collections.Generic.List[string]'
    $script:Config = @{}
    function script:Write-Log { param($Message, $Level = 'INFO') $script:LogLines.Add(("[{0}] {1}" -f $Level, $Message)) | Out-Null }
    function script:Get-SmartM365EffectiveConfigValue { param($Name, $DefaultValue) if ($script:Config.ContainsKey($Name)) { return $script:Config[$Name] } return $DefaultValue }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-FreeBusy-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

# Synthetic report with the structure Free/Busy Checker writes (headers, summary rows, colored values).
$syntheticHtml = @"
<!DOCTYPE html><html><head><title>Hybrid Free Busy Configuration Checker</title></head><body>
<div class='Black'><h2><b>Parameters:</b></h2><span style='color:green'>user.one@contoso.example</span></div>
<div class='Black'><h2><b>Exchange On Premise Free Busy Configuration:</b></h2></div>
<table style='width:100%'>
<tr><th ColSpan='2' style='text-align:center; color:white;'><b>Exchange On Premise DAuth Configuration</b></th></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Get-OrganizationRelationship</th></tr>
<tr><td><b>Get-OrganizationRelationship</b></td><td>
    <div> <b>Domain Names: </b> <span style='color: green'>contoso.mail.onmicrosoft.com</span></div>
    <div> <b>FreeBusyAccessEnabled: </b> <span style='color: green'>True</span></div>
    <div> <b>FreeBusyAccessLevel: </b> <span style='color: red'>AvailabilityOnly</span></div>
</td></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - On-Premise Get-AvailabilityAddressSpace</th></tr>
<tr><td><b> Get-AvailabilityAddressSpace contoso.mail.onmicrosoft.com | fl ForestName</b></td><td>
<div> <b>Forest Name: </b> contoso.mail.onmicrosoft.com</div>
<div> <b>Access Method: </b> <span style='color:orange'>InternalProxy</span></div>
<div> <b>ProxyUrl: </b> <span style='color:green'>https://mail.contoso.example/EWS/Exchange.asmx</span></div>
</td></tr>
<tr><th ColSpan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th></tr>
<tr><td><b>Test-OrganizationRelationship -Identity Rel -UserIdentity user.two@contoso.example</b></td><td>
<div class='green'> <b>No Significant Issues to Report</b><div>
</td></tr>
</table>
<div class='Black'><h2><b>Exchange Online Free Busy Configuration:</b></h2></div>
<table style='width:100%'>
<tr><th ColSpan='2' style='text-align:center; color:white;'><b>Exchange Online OAuth Configuration</b></th></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Get-IntraOrganizationConnector</th></tr>
<tr><td><b>Get-IntraOrganizationConnector</b></td><td>
<div> <b>Enabled: </b> <span style='color:green'>True</span></div>
<div> <b>DiscoveryEndpoint: </b> <span style='color:green'>https://autodiscover.contoso.example/autodiscover/autodiscover.svc</span></div>
<div> <b>Note: </b> <span style='color:red'>Token issued for 0f8fad5b-d9cb-469f-a165-70867728950e by host mail.contoso.example at 192.0.2.10, thumbprint 0123456789ABCDEF0123456789ABCDEF01234567</span></div>
</td></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Test-OAuthConnectivity</th></tr>
<tr><td><b>Test-OAuthConnectivity -Service EWS -TargetUri x -Mailbox user.one@contoso.example</b></td><td>
<div class='red'> <b>OAuth Test was completed with Error.</b></div>
</td></tr>
</table>
</body></html>
"@

try {
    # 1. Shared code stays identical.
    foreach ($name in @('Get-HealthCheckerUnsafeAclEntry', 'Get-HealthCheckerWorkFolderAncestor', 'Set-HealthCheckerWorkFolderAclRule', 'Initialize-HealthCheckerWorkFolder')) {
        Assert-FreeBusy ((Get-FunctionText $collectorAst $name) -ceq (Get-FunctionText $infrastructureAst $name)) "The work-folder check $name is identical to the infrastructure collector."
    }
    Assert-FreeBusy ((Get-FunctionText $collectorAst 'Test-FreeBusyCheckerIntegrity') -ceq (Get-FunctionText $runnerAst 'Test-FreeBusyCheckerIntegrity')) 'The runner checks Free/Busy Checker exactly like the collector.'

    # 2. Vendored file and expected report structure.
    Assert-FreeBusy ($constants.FreeBusyCheckerRelativePath -eq 'ThirdParty\CSS-Exchange\FreeBusyChecker\FreeBusyChecker.ps1' -and $constants.FreeBusyRunnerFileName -eq 'SmartM365-Exchange-Hybrid-FreeBusy-Runner.ps1') 'The collector points at the vendored file and the runner.'
    Assert-FreeBusy ((Get-FileHash -LiteralPath $vendoredPath -Algorithm SHA256).Hash -eq $constants.FreeBusyCheckerExpectedSha256) 'The expected SHA256 matches the vendored file.'
    $valid = & $module { param($p, $h) Test-FreeBusyCheckerIntegrity -Path $p -ExpectedSha256 $h } $vendoredPath $constants.FreeBusyCheckerExpectedSha256
    Assert-FreeBusy ($valid.IsValid -and $valid.Signer -match 'CN=Microsoft Corporation') "The vendored file passes the integrity check ($($valid.Reason))."
    $tampered = Join-Path $tempRoot 'Tampered.ps1'
    $bytes = [IO.File]::ReadAllBytes($vendoredPath); $bytes[2000] = if ($bytes[2000] -eq 0x41) { [byte]0x42 } else { [byte]0x41 }; [IO.File]::WriteAllBytes($tampered, $bytes)
    Assert-FreeBusy ((& $module { param($p, $h) Test-FreeBusyCheckerIntegrity -Path $p -ExpectedSha256 $h } $tampered $constants.FreeBusyCheckerExpectedSha256).Reason -match 'SHA256 mismatch') 'A modified file fails on SHA256.'
    $otherHash = (Get-FileHash -LiteralPath $nonMicrosoftSignedPath -Algorithm SHA256).Hash
    Assert-FreeBusy (-not (& $module { param($p, $h) Test-FreeBusyCheckerIntegrity -Path $p -ExpectedSha256 $h } $nonMicrosoftSignedPath $otherHash).IsValid) 'A file signed by another publisher is refused.'
    $vendoredText = [IO.File]::ReadAllText($vendoredPath)
    foreach ($title in @('Exchange On Premise DAuth Configuration', 'Exchange On Premise OAuth Configuration', 'Exchange Online DAuth Configuration', 'Exchange Online OAuth Configuration')) {
        Assert-FreeBusy ($vendoredText.Contains($title)) "Free/Busy Checker still writes the section header parsed by the collector: $title"
    }
    Assert-FreeBusy ([regex]::Matches($vendoredText, "<th ColSpan='2' style='color:white;'>(?:<b>)?Summary - ").Count -ge 20) 'Free/Busy Checker still writes the Summary rows parsed by the collector.'
    Assert-FreeBusy ($vendoredText.Contains('FBCheckerOutput_') -and $vendoredText.Contains("ParameterSetName = `"SkipVersionCheck`"")) 'Report name and the separate SkipVersionCheck parameter set are unchanged.'

    # 3. Text-only parsing.
    $findings = @(& $module { param($h) ConvertFrom-FreeBusyCheckerHtml -Html $h -ServerName 'EXSRV01' -FreeBusyCheckerVersion '26.06.06.1238' -GenerationTime '2026-10-25T05:00:00+01:00' } $syntheticHtml)
    $columns = & $module { Get-FreeBusyCsvColumn -Kind Findings }
    Assert-FreeBusy ((@($findings[0].PSObject.Properties.Name) -join ',') -eq ($columns -join ',')) 'Finding rows follow the Findings CSV columns.'
    Assert-FreeBusy ($findings.Count -eq 11) "Every result item becomes a finding; headings and command cells are skipped (found $($findings.Count))."
    $onPrem = @($findings | Where-Object { $_.Side -eq 'ExchangeOnPremises' -and $_.Auth -eq 'DAuth' })
    Assert-FreeBusy ($onPrem.Count -eq 7 -and @($findings | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.Auth -eq 'OAuth' }).Count -eq 4) 'Sections map to side and authentication.'
    Assert-FreeBusy (@($onPrem | Where-Object Category -eq 'Get-AvailabilityAddressSpace').Count -eq 3) 'The on-premises prefix is removed from check names.'
    $domainNames = @($findings | Where-Object Item -eq 'Domain Names')[0]
    Assert-FreeBusy ($domainNames.Value -eq '[masked]' -and $domainNames.ValueMasked -and $domainNames.Status -eq 'OK') 'Identifying items are masked and keep their status.'
    Assert-FreeBusy (@($findings | Where-Object { $_.Item -eq 'FreeBusyAccessEnabled' -and $_.Status -eq 'OK' -and $_.Value -eq 'True' }).Count -eq 1) 'Green values are OK.'
    Assert-FreeBusy (@($findings | Where-Object { $_.Item -eq 'FreeBusyAccessLevel' -and $_.Status -eq 'Error' }).Count -eq 1) 'Red values are errors.'
    Assert-FreeBusy (@($findings | Where-Object { $_.Item -eq 'Access Method' -and $_.Status -eq 'Warning' }).Count -eq 1) 'Orange values are warnings.'
    Assert-FreeBusy (@($findings | Where-Object { $_.Category -eq 'Test-OrganizationRelationship' -and $_.Status -eq 'OK' -and $_.Value -eq 'No Significant Issues to Report' }).Count -eq 1) 'A colored result without label is kept.'
    $note = @($findings | Where-Object Item -eq 'Note')[0]
    Assert-FreeBusy ($note.Value -match '\[guid\]' -and $note.Value -match '\[domain\]' -and $note.Value -match '\[ip\]' -and $note.Value -match '\[thumbprint\]') "Free text is scrubbed: $($note.Value)"
    $allValues = ($findings | ForEach-Object { "$($_.Category) $($_.Item) $($_.Value)" }) -join ' '
    Assert-FreeBusy ($allValues -notmatch 'contoso|@|192\.0\.2|0f8fad5b|https?://') 'No address, domain, URL or identifier reaches the findings.'
    Assert-FreeBusy (@(& $module { ConvertFrom-FreeBusyCheckerHtml -Html '<html>no sections</html>' -ServerName 'S' -FreeBusyCheckerVersion 'v' -GenerationTime 't' }).Count -eq 0) 'A report without sections gives no findings.'

    # 4. NotAuthorized checks.
    $authorized = @(& $module { param($f) Add-FreeBusyNotAuthorizedFinding -Findings $f -NotAuthorizedCmdlets @('Test-EOOAuthConnectivity', 'Get-EOSharingPolicy') -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' } $findings)
    $oauthTest = @($authorized | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.Category -eq 'Test-OAuthConnectivity' })
    Assert-FreeBusy ($oauthTest.Count -eq 1 -and $oauthTest[0].Status -eq 'NotAuthorized' -and $oauthTest[0].Value -match 'Mail Recipients') 'A check the app cannot run is NotAuthorized with the missing role, not an error.'
    $added = @($authorized | Where-Object { $_.Item -eq 'Get-EOSharingPolicy' })
    Assert-FreeBusy ($added.Count -eq 1 -and $added[0].Status -eq 'NotAuthorized' -and $added[0].Auth -eq 'DAuth' -and $added[0].RowNumber -eq 12) 'A missing cmdlet without reported rows gets its own NotAuthorized row.'

    # 5. Summary rows and run status.
    $runner = { param($ExchangeOnline, [string[]]$Missing) [pscustomobject]@{ Stage = 'FreeBusyCheckerReturned'; RunStatus = $null; Message = $null; ExchangeOnline = $ExchangeOnline; ExchangeOnlineMessage = $null; NotAuthorizedCmdlets = $Missing } }
    $partial = & $runner 'NotAuthorizedPartial' @('Test-EOOAuthConnectivity')
    $summary = @(& $module { param($f, $r) New-FreeBusySummaryRow -Findings $f -RunStatus 'EXONotAuthorized' -RunMessage 'm' -RunnerStatus $r -DurationSeconds 60 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' } $authorized $partial)
    $summaryColumns = & $module { Get-FreeBusyCsvColumn -Kind Summary }
    Assert-FreeBusy ($summary.Count -eq 4 -and (@($summary[0].PSObject.Properties.Name) -join ',') -eq ($summaryColumns -join ',')) 'Four summary rows with the Summary CSV columns.'
    $exoOAuth = @($summary | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.Auth -eq 'OAuth' })[0]
    Assert-FreeBusy ($exoOAuth.SideStatus -eq 'NotAuthorizedPartial' -and $exoOAuth.NotAuthorizedCmdlets -eq 'Test-EOOAuthConnectivity' -and $exoOAuth.MissingRoles -match 'Mail Recipients' -and $exoOAuth.NotAuthorizedCount -eq 1) 'Exchange Online OAuth is partial with the missing cmdlet and role.'
    Assert-FreeBusy (@($summary | Where-Object { $_.Side -eq 'ExchangeOnPremises' -and $_.Auth -eq 'DAuth' })[0].SideStatus -eq 'Completed') 'On-premises DAuth is completed.'
    Assert-FreeBusy (@($summary | Where-Object { $_.Side -eq 'ExchangeOnPremises' -and $_.Auth -eq 'OAuth' })[0].SideStatus -eq 'NotRun') 'A side without results is NotRun.'
    $notConnected = @(& $module { param($f, $r) New-FreeBusySummaryRow -Findings $f -RunStatus 'EXONotConnected' -RunMessage 'm' -RunnerStatus $r -DurationSeconds 60 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' } @($findings | Where-Object Side -eq 'ExchangeOnPremises') (& $runner 'NotConnected' @()))
    Assert-FreeBusy (@($notConnected | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.SideStatus -eq 'NotConnected' -and $null -eq $_.ErrorCount }).Count -eq 2) 'Exchange Online not connected has no measured counts.'
    $blocked = @(& $module { New-FreeBusySummaryRow -Findings @() -RunStatus 'UnsafeWorkFolder' -RunMessage 'm' -RunnerStatus $null -DurationSeconds $null -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' })
    Assert-FreeBusy (@($blocked | Where-Object { $_.SideStatus -eq 'NotRun' -and $null -eq $_.OkCount }).Count -eq 4) 'A run that did not start is NotRun everywhere.'
    $resolve = { param($p, $r, $h, $x) & $module { param($a, $b, $c, $d) Resolve-FreeBusyRunStatus -Process $a -RunnerStatus $b -HtmlFound $c -ParseFailed $d } $p $r $h $x }
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Timeout' }) $null $false $false) -eq 'Timeout') 'Timeout wins.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) ([pscustomobject]@{ RunStatus = 'ModuleInstallFailed'; ExchangeOnline = 'NotConnected'; NotAuthorizedCmdlets = @() }) $false $false) -eq 'ModuleInstallFailed') 'The runner status is kept.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) (& $runner 'Connected' @()) $false $false) -eq 'NoOutput') 'No report is NoOutput.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) (& $runner 'Connected' @()) $true $true) -eq 'ParseFailed') 'An unreadable report is ParseFailed.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) (& $runner 'NotConnected' @()) $true $false) -eq 'EXONotConnected') 'Exchange Online not connected.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) $partial $true $false) -eq 'EXONotAuthorized') 'Exchange Online partly not authorized.'
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) (& $runner 'Connected' @()) $true $false) -eq 'Completed') 'Completed.'

    # 6. Runner status file: plain JSON, cmdlet names filtered.
    $statusFile = Join-Path $tempRoot 'status.json'
    [IO.File]::WriteAllText($statusFile, '{"Stage":"FreeBusyCheckerReturned","RunStatus":null,"ExchangeOnline":"NotAuthorizedPartial","NotAuthorizedCmdlets":["Test-EOOAuthConnectivity","<script>"]}')
    $read = & $module { param($p) Read-FreeBusyRunnerStatus -Path $p } $statusFile
    Assert-FreeBusy ((@($read.NotAuthorizedCmdlets) -join ',') -eq 'Test-EOOAuthConnectivity' -and $read.ExchangeOnline -eq 'NotAuthorizedPartial') 'Only cmdlet names are taken from the status file.'
    Assert-FreeBusy ($null -eq (& $module { param($p) Read-FreeBusyRunnerStatus -Path $p } (Join-Path $tempRoot 'none.json'))) 'A missing status file is no status.'
    foreach ($ast in @($collectorAst, $runnerAst)) {
        Assert-FreeBusy (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('Import-Clixml', 'Read-Host', 'Get-Credential') }, $true)).Count -eq 0) 'Nothing is deserialized and nothing prompts.'
    }

    # 7. Runner: app-only session only, CurrentUser installation, -SkipVersionCheck only.
    $connects = @($runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Connect-ExchangeOnline' }, $true))
    Assert-FreeBusy ($connects.Count -eq 1 -and $connects[0].Extent.Text -match '-AppId \$AppId' -and $connects[0].Extent.Text -match '-CertificateThumbprint \$CertificateThumbprint' -and $connects[0].Extent.Text -match '-Prefix EO') 'Exchange Online is opened once, with the app and its certificate, prefix EO.'
    $checkerFunction = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-FreeBusyCheckerScript' }, $true)
    $fbcCalls = @($runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.CommandElements[0].Extent.Text -in @('$FreeBusyCheckerPath', '$Path') }, $true))
    Assert-FreeBusy ($null -ne $checkerFunction -and $fbcCalls.Count -eq 1 -and $checkerFunction.Extent.Text.Contains($fbcCalls[0].Extent.Text) -and
        $fbcCalls[0].CommandElements.Count -eq 2 -and $fbcCalls[0].CommandElements[1].Extent.Text -eq '-SkipVersionCheck') 'Free/Busy Checker runs only through its function, with -SkipVersionCheck only.'
    Assert-FreeBusy (@($runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-FreeBusyCheckerScript' }, $true)).Count -eq 1) 'The runner starts Free/Busy Checker once.'

    # Free/Busy Checker reads variables it never sets (for example $versionsUrl with -SkipVersionCheck) and
    # relies on non-terminating errors: a synthetic checker is run for real, in a separate process, under
    # the runner's settings.
    $fakeChecker = Join-Path $tempRoot 'FakeFreeBusyChecker.ps1'
    [IO.File]::WriteAllText($fakeChecker, (@(
        'param([switch]$SkipVersionCheck)',
        'if (-not ([string]::IsNullOrEmpty($versionsUrl))) { ''unexpected'' }',
        'Write-Error ''Synthetic non-terminating service error''',
        '"FAKE-CHECKER-DONE SkipVersionCheck=$SkipVersionCheck"'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $strictHost = Join-Path $tempRoot 'StrictHost.ps1'
    [IO.File]::WriteAllText($strictHost, (@(
        'param([string]$Checker, [switch]$Direct)',
        'Set-StrictMode -Version Latest',
        '$ErrorActionPreference = ''Stop''',
        $checkerFunction.Extent.Text,
        'try { if ($Direct) { & $Checker -SkipVersionCheck } else { Invoke-FreeBusyCheckerScript -Path $Checker } } catch { "RUN-FAILED: $($_.Exception.Message)"; exit 1 }',
        'try { $null = $notSetAfterRun; ''STRICT-OFF-AFTER'' } catch { ''STRICT-STILL-ON'' }',
        'if ($ErrorActionPreference -eq ''Stop'') { ''STOP-STILL-ON'' }'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $ErrorActionPreference = 'Continue'
    $strictRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $strictHost -Checker $fakeChecker 2>&1 | Out-String)
    $strictExit = $LASTEXITCODE
    $directRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $strictHost -Checker $fakeChecker -Direct 2>&1 | Out-String)
    $directExit = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    Assert-FreeBusy ($directExit -eq 1 -and $directRun -match 'RUN-FAILED') 'The synthetic checker fails when called directly under the runner settings (the call before runner 1.0.2).'
    # Windows PowerShell 5.1 wraps the error record it writes to stderr at its buffer width, inside the message.
    $strictRunText = $strictRun -replace '\s+', ' '
    Assert-FreeBusy ($strictExit -eq 0 -and $strictRun -match 'FAKE-CHECKER-DONE SkipVersionCheck=True' -and $strictRunText -match 'Synthetic non-terminating service error') "Free/Busy Checker runs to its end without strict mode and keeps its non-terminating errors: $($strictRun.Trim())"
    Assert-FreeBusy ($strictRun -match 'STRICT-STILL-ON' -and $strictRun -match 'STOP-STILL-ON') 'The runner keeps strict mode and stop-on-error for its own code.'

    # Exchange Management Shell: Free/Busy Checker does not load it, the runner does, before Exchange Online.
    $shellFunction = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Import-FreeBusyExchangeShell' }, $true)
    $shellCall = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Import-FreeBusyExchangeShell' }, $true)
    $exoCall = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Connect-FreeBusyExchangeOnline' }, $true)
    $checkerCall = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-FreeBusyCheckerScript' }, $true)
    Assert-FreeBusy ($null -ne $shellFunction -and $null -ne $shellCall -and $shellCall.Extent.StartOffset -lt $exoCall.Extent.StartOffset -and $exoCall.Extent.StartOffset -lt $checkerCall.Extent.StartOffset) 'The runner loads the Exchange Management Shell before Exchange Online and Free/Busy Checker.'
    Assert-FreeBusy ($runnerAst.Extent.Text -match "\`$FreeBusyExchangeShellCmdlets = @\('Get-RemoteMailbox', 'Get-OrganizationRelationship', 'Test-OAuthConnectivity'\)" -and
        @([regex]::Matches($runnerAst.Extent.Text, "RunStatus = 'ExchangeShellNotLoaded'")).Count -eq 2) 'The on-premises cmdlets are checked after loading and again after the Exchange Online connection.'
    $fakeExchange = Join-Path $tempRoot 'FakeExchange'
    New-Item -ItemType Directory -Path (Join-Path $fakeExchange 'bin') -Force | Out-Null
    # Written for default settings like the real loader: it reads a variable it never sets.
    [IO.File]::WriteAllText((Join-Path $fakeExchange 'bin\RemoteExchange.ps1'), (@(
        'if ($exchangeLoaderNotSet) { ''unexpected'' }',
        'function global:Connect-ExchangeServer { param([switch]$auto, $ClientApplication) foreach ($name in @(''Get-RemoteMailbox'', ''Get-OrganizationRelationship'', ''Test-OAuthConnectivity'')) { Set-Item -Path "function:global:$name" -Value { ''on-premises'' } } }'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $shellHost = Join-Path $tempRoot 'ShellHost.ps1'
    [IO.File]::WriteAllText($shellHost, (@(
        'param([string]$InstallPath)',
        'Set-StrictMode -Version Latest',
        '$ErrorActionPreference = ''Stop''',
        $shellFunction.Extent.Text,
        'if ($InstallPath) { $env:ExchangeInstallPath = $InstallPath } else { Remove-Item Env:\ExchangeInstallPath -ErrorAction SilentlyContinue }',
        '$result = Import-FreeBusyExchangeShell -RequiredCmdlets @(''Get-RemoteMailbox'', ''Get-OrganizationRelationship'', ''Test-OAuthConnectivity'')',
        '"LOADED=$($result.Loaded)|$($result.Message)"',
        'if (Get-Command Get-RemoteMailbox -ErrorAction SilentlyContinue) { ''VISIBLE-TO-RUNNER'' }',
        'try { $null = $notSetAfterShell; ''STRICT-OFF-AFTER'' } catch { ''STRICT-STILL-ON'' }'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $ErrorActionPreference = 'Continue'
    $loadedRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $shellHost -InstallPath $fakeExchange 2>&1 | Out-String)
    $noToolsRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $shellHost 2>&1 | Out-String)
    $emptyExchange = Join-Path $tempRoot 'EmptyExchange'
    New-Item -ItemType Directory -Path $emptyExchange -Force | Out-Null
    $noLoaderRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $shellHost -InstallPath $emptyExchange 2>&1 | Out-String)
    $ErrorActionPreference = 'Stop'
    Assert-FreeBusy ($loadedRun -match 'LOADED=True\|' -and $loadedRun -match 'VISIBLE-TO-RUNNER' -and $loadedRun -match 'STRICT-STILL-ON') "The shell loader runs under the runner settings and its cmdlets are visible afterwards: $($loadedRun.Trim())"
    Assert-FreeBusy ($noToolsRun -match 'LOADED=False\|ExchangeInstallPath is not set') 'Without the Exchange Management Tools the shell is reported not loaded.'
    Assert-FreeBusy ($noLoaderRun -match 'LOADED=False\|RemoteExchange\.ps1 was not found') 'Without RemoteExchange.ps1 the shell is reported not loaded.'
    $shellStatus = [pscustomobject]@{ Stage = 'ExchangeShell'; RunStatus = 'ExchangeShellNotLoaded'; Message = 'm'; ExchangeOnline = 'NotConnected'; ExchangeOnlineMessage = $null; NotAuthorizedCmdlets = @('Test-EOOAuthConnectivity') }
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) $shellStatus $false $false) -eq 'ExchangeShellNotLoaded') 'A shell that did not load is ExchangeShellNotLoaded, not a generic runner failure.'
    $shellSummary = @(& $module { param($r) New-FreeBusySummaryRow -Findings @() -RunStatus 'ExchangeShellNotLoaded' -RunMessage 'm' -RunnerStatus $r -DurationSeconds 5 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' } $shellStatus)
    Assert-FreeBusy (@($shellSummary | Where-Object { $_.SideStatus -ne 'NotRun' -or $_.MissingRoles }).Count -eq 0) 'A shell that did not load leaves every side NotRun.'

    # Exchange Online cmdlets the app cannot run: Free/Busy Checker calls them unguarded, so each gets a stand-in.
    $standInRoles = $runnerAst.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$FreeBusyStandInRoles' }, $true)
    $fakeExoChecker = Join-Path $tempRoot 'FakeExoChecker.ps1'
    [IO.File]::WriteAllText($fakeExoChecker, (@(
        'param([switch]$SkipVersionCheck)',
        '$oauth = Test-EOOAuthConnectivity -Service EWS -TarGetUri $null -Mailbox ''user'' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue',
        'if ($oauth.ResultType.Value -like ''Success'') { ''OAUTH-OK'' } else { ''OAUTH-NOT-OK'' }',
        '"OAUTH-ERROR=" + $Error[0].Exception.Message',
        '$policy = Get-EOSharingPolicy | Select-Object Name',
        '"FAKE-EXO-END oauthNull=$($null -eq $oauth) policyNull=$($null -eq $policy)"'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $catchClause = @($runnerAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Body.Extent.Text.Contains('Invoke-FreeBusyCheckerScript -Path') })[0].CatchClauses[0].Extent.Text
    $exoHost = Join-Path $tempRoot 'ExoHost.ps1'
    [IO.File]::WriteAllText($exoHost, (@(
        'param([string]$Checker, [switch]$NoStandIn)',
        'Set-StrictMode -Version Latest',
        '$ErrorActionPreference = ''Stop''',
        $standInRoles.Extent.Text,
        $checkerFunction.Extent.Text,
        '$status = [ordered]@{ Stage = ''FreeBusyCheckerStarted''; RunStatus = $null; Message = $null; ExchangeOnline = ''NotAuthorizedPartial''; ExchangeOnlineMessage = $null; NotAuthorizedCmdlets = @(''Test-EOOAuthConnectivity'', ''Get-EOSharingPolicy'') }',
        'try { Invoke-FreeBusyCheckerScript -Path $Checker -UnavailableCmdlets $(if ($NoStandIn) { @() } else { @($status.NotAuthorizedCmdlets) }) }',
        $catchClause,
        '"STATUS=" + ($status | ConvertTo-Json -Compress)',
        'if (Get-Command Test-EOOAuthConnectivity -ErrorAction SilentlyContinue) { ''STAND-IN-LEAKED'' }'
    ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $ErrorActionPreference = 'Continue'
    $standInRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $exoHost -Checker $fakeExoChecker 2>&1 | Out-String -Width 400)
    $stoppedRun = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $exoHost -Checker $fakeExoChecker -NoStandIn 2>&1 | Out-String -Width 400)
    $ErrorActionPreference = 'Stop'
    $standInStatus = ([regex]::Match($standInRun, 'STATUS=(\{.*\})').Groups[1].Value) | ConvertFrom-Json
    Assert-FreeBusy ($standInRun -match 'OAUTH-NOT-OK' -and $standInRun -match 'FAKE-EXO-END oauthNull=True policyNull=True' -and $standInRun -match 'OAUTH-ERROR=Test-EOOAuthConnectivity is not authorized for the tenant app \(Mail Recipients or Organization Transport Settings role required\)' -and $standInRun -match 'Get-EOSharingPolicy is not authorized for the tenant app \(role not assigned\)') "With stand-ins Free/Busy Checker runs to its end and the unauthorized check is never OK: $($standInRun.Trim())"
    Assert-FreeBusy ($null -eq $standInStatus.RunStatus -and $standInStatus.Stage -eq 'FreeBusyCheckerStarted' -and $standInRun -notmatch 'STAND-IN-LEAKED') 'The stand-ins stay local to the Free/Busy Checker call and the run is not a failure.'
    $stoppedStatus = ([regex]::Match($stoppedRun, 'STATUS=(\{.*\})').Groups[1].Value) | ConvertFrom-Json
    Assert-FreeBusy ($stoppedRun -notmatch 'FAKE-EXO-END' -and $null -eq $stoppedStatus.RunStatus -and $stoppedStatus.Stage -eq 'FreeBusyCheckerStoppedNotAuthorized' -and $stoppedStatus.ExchangeOnline -eq 'NotAuthorizedPartial') "A stop caused only by an Exchange Online cmdlet the app cannot run is not a runner failure: $($stoppedRun.Trim())"
    $stoppedRead = [pscustomobject]@{ Stage = $stoppedStatus.Stage; RunStatus = $stoppedStatus.RunStatus; ExchangeOnline = $stoppedStatus.ExchangeOnline; NotAuthorizedCmdlets = @($stoppedStatus.NotAuthorizedCmdlets) }
    Assert-FreeBusy ((& $resolve ([pscustomobject]@{ Status = 'Exited' }) $stoppedRead $true $false) -eq 'EXONotAuthorized') 'That stop resolves to EXONotAuthorized and keeps the reported findings.'
    $installs = @($runnerAst.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -in @('Install-Module', 'Install-PackageProvider') }, $true))
    Assert-FreeBusy ($installs.Count -eq 2 -and @($installs | Where-Object { $_.Extent.Text -match '-Scope CurrentUser' -and $_.Extent.Text -match '-Force' }).Count -eq 2) 'Installations are CurrentUser and without prompt.'
    Assert-FreeBusy ($runnerAst.Extent.Text -match "\[Parameter\(Mandatory\)\]\[string\]\`$FreeBusyCheckerPath" -and $runnerAst.Extent.Text -notmatch 'param\([^)]*AppId') 'App identifiers do not travel on the runner command line.'

    $runnerModule = New-Module -ScriptBlock ([scriptblock]::Create(((@('Connect-FreeBusyExchangeOnline', 'Install-FreeBusyExchangeOnlineModule') | ForEach-Object { Get-FunctionText $runnerAst $_ }) -join "`n")))
    & $runnerModule {
        $script:Calls = New-Object 'System.Collections.Generic.List[string]'
        $script:Connection = $null
        $script:Modules = @()
        function script:Connect-ExchangeOnline { param($AppId, $CertificateThumbprint, $Organization, $Prefix, [switch]$ShowBanner, $ErrorAction) $script:Calls.Add("connect:${AppId}:$Prefix") }
        function script:Get-ConnectionInformation { $script:Connection }
        function script:Disconnect-ExchangeOnline { param([switch]$Confirm, $ErrorAction) $script:Calls.Add('disconnect') }
        function script:Get-Module { param([switch]$ListAvailable, $Name) $script:Modules }
        function script:Get-Command { param($Name, $ErrorAction) [pscustomobject]@{ Name = $Name } }
        function script:Get-PackageProvider { param([switch]$ListAvailable, $Name, $ErrorAction) $null }
        function script:Install-PackageProvider { param($Name, $MinimumVersion, $Scope, [switch]$Force, $ErrorAction) $script:Calls.Add("nuget:$Scope") }
        function script:Install-Module { param($Name, $MinimumVersion, $Scope, $Repository, [switch]$Force, [switch]$AllowClobber, $ErrorAction) $script:Calls.Add("install:${Name}:${Scope}:$Repository"); $script:Modules = @([pscustomobject]@{ Name = $Name; Version = [version]'3.6.0' }) }
    }
    $none = & $runnerModule { Connect-FreeBusyExchangeOnline -AppId '' -CertificateThumbprint 'T' -Organization 'O' }
    Assert-FreeBusy (-not $none.Connected -and $none.Message -match 'no Exchange Online connection was attempted' -and (& $runnerModule { $script:Calls.Count }) -eq 0) 'Without app settings nothing is attempted.'
    & $runnerModule { $script:Connection = [pscustomobject]@{ ModulePrefix = 'EO'; State = 'Connected'; AppId = 'other'; CertificateAuthentication = $false } }
    $wrong = & $runnerModule { Connect-FreeBusyExchangeOnline -AppId 'app' -CertificateThumbprint 'T' -Organization 'O' }
    Assert-FreeBusy (-not $wrong.Connected -and ((& $runnerModule { $script:Calls }) -join ',') -eq 'connect:app:EO,disconnect') 'A session that is not the app certificate session is closed and not used.'
    & $runnerModule { $script:Calls.Clear(); $script:Connection = [pscustomobject]@{ ModulePrefix = 'EO'; State = 'Connected'; AppId = 'app'; CertificateAuthentication = $true } }
    Assert-FreeBusy ((& $runnerModule { Connect-FreeBusyExchangeOnline -AppId 'app' -CertificateThumbprint 'T' -Organization 'O' }).Connected) 'The app-only certificate session is accepted.'
    & $runnerModule { $script:Calls.Clear() }
    $installed = & $runnerModule { Install-FreeBusyExchangeOnlineModule -MinimumVersion ([version]'3.0.0') }
    Assert-FreeBusy ($installed.Version -eq [version]'3.6.0' -and ((& $runnerModule { $script:Calls }) -join ',') -eq 'nuget:CurrentUser,install:ExchangeOnlineManagement:CurrentUser:PSGallery') 'A missing module is installed for the current user from PSGallery after NuGet.'
    & $runnerModule { $script:Calls.Clear() }
    $null = & $runnerModule { Install-FreeBusyExchangeOnlineModule -MinimumVersion ([version]'3.0.0') }
    Assert-FreeBusy ((& $runnerModule { $script:Calls.Count }) -eq 0) 'An installed module is not reinstalled.'

    # 8. Settings and runner check.
    $settings = & $module { $script:Config = @{}; Get-FreeBusySetting }
    Assert-FreeBusy ($settings.WorkFolderPath -eq (Join-Path $env:ProgramFiles 'SmartM365\FreeBusy\Work') -and $settings.TimeoutMinutes -eq 30) 'Default work folder under Program Files, 30 minute limit.'
    $refused = $false
    try { & $module { $script:Config = @{ FreeBusyWorkFolderPath = '\\server\share\fb' }; Get-FreeBusySetting } } catch { $refused = $_.Exception.Message -match 'local absolute path' }
    Assert-FreeBusy $refused 'A share as work folder is refused.'
    $refused = $false
    try { & $module { $script:Config = @{ FreeBusyCheckerTimeoutMinutes = 200 }; Get-FreeBusySetting } } catch { $refused = $true }
    Assert-FreeBusy $refused 'A limit beyond the job timeout is refused.'
    & $module { $script:Config = @{} }
    $runnerHash = (Get-FileHash -LiteralPath $runnerPath -Algorithm SHA256).Hash
    Assert-FreeBusy ($null -eq (& $module { param($p, $h) Test-FreeBusyRunnerIntegrity -Path $p -ExpectedSha256 $h -CollectorSignerThumbprint $null } $runnerPath $runnerHash)) 'The repository runner passes.'
    Assert-FreeBusy ((& $module { param($p) Test-FreeBusyRunnerIntegrity -Path $p -ExpectedSha256 'AA' -CollectorSignerThumbprint $null } $runnerPath) -match 'differs') 'A different runner is refused.'
    Assert-FreeBusy ((& $module { param($p, $h) Test-FreeBusyRunnerIntegrity -Path $p -ExpectedSha256 $h -CollectorSignerThumbprint 'NOT-THE-SIGNER' } $runnerPath $runnerHash) -match 'not signed by') 'A runner not signed like the collector is refused.'
    $unsafeRoot = Join-Path $tempRoot 'unsafe\Work'
    $unsafe = $false
    try { & $module { param($r) Initialize-HealthCheckerWorkFolder -WorkRoot $r } $unsafeRoot } catch { $unsafe = $_.Exception.Message -match 'not restricted to administrators' }
    Assert-FreeBusy $unsafe 'A work folder below a user-writable parent is refused.'

    # 9. Child process: environment hand-off and time limit.
    $fakeRunner = Join-Path $tempRoot 'FakeRunner.ps1'
    [IO.File]::WriteAllText($fakeRunner, @'
param([string]$FreeBusyCheckerPath, [string]$ExpectedSha256, [string]$StatusPath)
if ($env:SMARTM365_FREEBUSY_APP_ID -eq 'sleep') { Set-Content -LiteralPath "$StatusPath.pid" -Value $PID; Start-Sleep -Seconds 120 }
$state = @{ Stage = 'FreeBusyCheckerReturned'; RunStatus = $null; ExchangeOnline = 'Connected'; NotAuthorizedCmdlets = @(); Message = ("{0}|{1}|{2}" -f $env:SMARTM365_FREEBUSY_APP_ID, $env:SMARTM365_FREEBUSY_CERT_THUMBPRINT, $env:SMARTM365_FREEBUSY_ORGANIZATION) }
[IO.File]::WriteAllText($StatusPath, ($state | ConvertTo-Json))
'@, [Text.UTF8Encoding]::new($false))
    $processSettings = [pscustomobject]@{ AppId = 'app-1'; CertificateThumbprint = 'THUMB'; Organization = 'org.example' }
    $okStatus = Join-Path $tempRoot 'ok-status.json'
    $ok = & $module { param($r, $s, $set) Invoke-FreeBusyRunnerProcess -RunnerPath $r -FreeBusyCheckerPath 'x.ps1' -StatusPath $s -TimeoutSeconds 120 -Settings $set } $fakeRunner $okStatus $processSettings
    $okRead = & $module { param($p) Read-FreeBusyRunnerStatus -Path $p } $okStatus
    Assert-FreeBusy ($ok.Status -eq 'Exited' -and $okRead.Message -eq 'app-1|THUMB|org.example') "The runner receives the app settings through its environment ($($ok.Status))."
    Assert-FreeBusy ($null -eq [Environment]::GetEnvironmentVariable('SMARTM365_FREEBUSY_APP_ID', 'Process')) 'The app settings are removed from the collector environment.'
    $sleepStatus = Join-Path $tempRoot 'sleep-status.json'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $timeout = & $module { param($r, $s) Invoke-FreeBusyRunnerProcess -RunnerPath $r -FreeBusyCheckerPath 'x.ps1' -StatusPath $s -TimeoutSeconds 5 -Settings ([pscustomobject]@{ AppId = 'sleep'; CertificateThumbprint = 'T'; Organization = 'O' }) } $fakeRunner $sleepStatus
    $watch.Stop()
    Assert-FreeBusy ($timeout.Status -eq 'Timeout' -and $watch.Elapsed.TotalSeconds -lt 60) "A hung run is stopped at its limit ($($timeout.Status))."
    $sleepPid = [int](Get-Content -LiteralPath "$sleepStatus.pid" -Raw)
    Assert-FreeBusy ($null -eq (Get-Process -Id $sleepPid -ErrorAction SilentlyContinue)) 'The stopped run leaves no process behind.'

    # 10. Report.
    $report = & $module { param($s, $f) ConvertTo-FreeBusyReportHtml -Summary $s -Findings $f -RunStatus 'EXONotAuthorized' -RunMessage '<b>m</b>' } $summary $authorized
    Assert-FreeBusy ($report -match 'EXONotAuthorized' -and $report -match 'NotAuthorizedPartial' -and $report -match '&lt;b&gt;m&lt;/b&gt;') 'The report shows the statuses and encodes text.'
    Assert-FreeBusy ($report.IndexOf('>Error<') -lt $report.IndexOf('>NotAuthorized<')) 'Errors are listed before not authorized checks.'

    # Mail banner: the shared branding puts the Power BI link beside the report title, in its hero banner.
    $coreAst = Get-ScriptAst (Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1')
    $brandingModule = New-Module -ScriptBlock ([scriptblock]::Create(((@('Add-SmartM365MailBranding', 'ConvertTo-SmartM365EmailHtmlText') | ForEach-Object { Get-FunctionText $coreAst $_ }) + @(
        'function Get-SmartM365MailBrandingConfig { [pscustomobject]@{ Enabled = $true; ClientName = ''''; ClientLogoPath = ''''; ClientLogoMaxKB = 64; PowerBiEnabled = $true; PowerBiUrl = ''https://app.powerbi.com/groups/me/reports/synthetic''; PowerBiLinkText = ''Click here to open the Power BI dashboard''; PowerBiLogoPath = ''''; FooterProductName = ''SmartM365''; FooterWebsiteUrl = ''https://example.org''; FooterGitHubLabel = ''GitHub''; FooterGitHubUrl = ''https://example.org/git'' } }',
        'function ConvertTo-SmartM365MailLogoDataUri { param($Path, $MaxKB) '''' }',
        'function Resolve-SmartM365ConfigValue { param($Value) $Value }',
        'function Add-SmartM365MailExecutionFooter { param($BodyHtml) $BodyHtml }',
        'function WriteLog { param($Message, $Level) }'
    )) -join "`n"))
    $branded = & $brandingModule { param($h) Add-SmartM365MailBranding -BodyHtml $h } $report
    $titleCell = [regex]::Match($branded, '(?is)<td align="left" valign="middle"[^>]*>(?<Title>.*?)</td>\s*<td align="right" valign="middle"[^>]*>(?<Link>.*?)</td>')
    Assert-FreeBusy ($titleCell.Success -and $titleCell.Groups['Title'].Value -match 'Exchange Hybrid Free/Busy Configuration' -and $titleCell.Groups['Title'].Value -match 'Run status: <strong>EXONotAuthorized</strong>' -and $titleCell.Groups['Link'].Value -match 'Click here to open the Power BI dashboard') 'The mail banner shows the title and the run status beside the Power BI link.'
    Assert-FreeBusy ($branded -notmatch 'background:#0f172a;text-align:right;') 'The Power BI link is not put in a separate empty banner.'
    Assert-FreeBusy ($report -match '>Sides tested<' -and $report -match '>3 / 4<' -and $report -match '>Not authorized<') 'The report shows the tested sides and the counts.'

    # A run that did not happen never reads as "no problem found" and shows no missing role.
    $runnerFailedSummary = @(& $module { param($r) New-FreeBusySummaryRow -Findings @() -RunStatus 'RunnerFailed' -RunMessage 'Free/Busy Checker runner failed: synthetic.' -RunnerStatus $r -DurationSeconds 24 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' } $partial)
    Assert-FreeBusy (@($runnerFailedSummary | Where-Object { $_.SideStatus -ne 'NotRun' -or $_.MissingRoles -or $_.NotAuthorizedCmdlets }).Count -eq 0) 'A run that did not happen lists no missing cmdlet or role.'
    $failedReport = & $module { param($s) ConvertTo-FreeBusyReportHtml -Summary $s -Findings @() -RunStatus 'RunnerFailed' -RunMessage 'Free/Busy Checker runner failed: synthetic.' } $runnerFailedSummary
    Assert-FreeBusy ($failedReport -match 'The check did not run: Free/Busy Checker runner failed: synthetic\.' -and $failedReport -notmatch 'No error, warning or not authorized check was reported') 'A run that did not happen says so and does not report "no problem".'
    Assert-FreeBusy ($failedReport -match '>N/A<' -and $failedReport -match '>0 / 4<' -and $failedReport -notmatch 'Mail Recipients') 'A run that did not happen shows unmeasured counts and no missing role.'
    $partialReport = & $module { param($s) ConvertTo-FreeBusyReportHtml -Summary $s -Findings @() -RunStatus 'EXONotConnected' -RunMessage 'm' } $notConnected
    Assert-FreeBusy ($partialReport -match 'on the 1 side\(s\) that were measured' -and $partialReport -notmatch 'No error, warning or not authorized check was reported\.<') 'Unmeasured sides are not reported as healthy.'

    # OAuth only: with the Intra Organization Connector enabled, Free/Busy Checker skips its DAuth checks.
    $oauthOnlyText = 'Checking OAuth only as Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled'
    Assert-FreeBusy ($vendoredText.Contains($oauthOnlyText)) 'Free/Busy Checker still writes the OAuth-only message the collector detects.'
    Assert-FreeBusy ((& $module { param($h) Test-FreeBusyOAuthOnly -Html $h } ("<div><p> <span style='color: green;'>$oauthOnlyText</span></p></div>" + $syntheticHtml)) -and -not (& $module { param($h) Test-FreeBusyOAuthOnly -Html $h } $syntheticHtml)) 'The OAuth-only message is detected only when present.'
    $oauthFindings = @($findings | Where-Object Auth -eq 'OAuth')
    $oauthOnlyRows = @(& $module { param($f) Add-FreeBusyNotAuthorizedFinding -Findings $f -NotAuthorizedCmdlets @('Test-EOOAuthConnectivity', 'Get-EOSharingPolicy') -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' -OAuthOnly } $oauthFindings)
    Assert-FreeBusy (@($oauthOnlyRows | Where-Object Item -eq 'Get-EOSharingPolicy').Count -eq 0 -and @($oauthOnlyRows | Where-Object { $_.Category -eq 'Test-OAuthConnectivity' -and $_.Status -eq 'NotAuthorized' }).Count -eq 1) 'A DAuth check that does not apply gets no NotAuthorized row; OAuth ones still do.'
    $oauthOnlySummary = @(& $module { param($f, $r) New-FreeBusySummaryRow -Findings $f -RunStatus 'EXONotAuthorized' -RunMessage 'm' -RunnerStatus $r -DurationSeconds 51 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' -OAuthOnly } $oauthOnlyRows (& $runner 'NotAuthorizedPartial' @('Test-EOOAuthConnectivity', 'Get-EOSharingPolicy')))
    $dauthRows = @($oauthOnlySummary | Where-Object Auth -eq 'DAuth')
    Assert-FreeBusy ($dauthRows.Count -eq 2 -and @($dauthRows | Where-Object { $_.SideStatus -ne 'NotApplicable' -or $_.NotAuthorizedCmdlets -or $_.MissingRoles -or $null -ne $_.ErrorCount -or $_.RunMessage -ne 'Intra Organization Connector enabled: free/busy uses OAuth, DAuth not checked' }).Count -eq 0) 'Both DAuth sides are NotApplicable with the reason, no counts and no missing role.'
    Assert-FreeBusy (@($oauthOnlySummary | Where-Object { $_.Side -eq 'ExchangeOnPremises' -and $_.Auth -eq 'OAuth' })[0].SideStatus -eq 'NotRun' -and @($oauthOnlySummary | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.Auth -eq 'OAuth' })[0].SideStatus -eq 'NotAuthorizedPartial') 'OAuth sides keep their own status.'
    $failedOAuthOnly = @(& $module { New-FreeBusySummaryRow -Findings @() -RunStatus 'RunnerFailed' -RunMessage 'm' -RunnerStatus $null -DurationSeconds 1 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' -OAuthOnly })
    Assert-FreeBusy (@($failedOAuthOnly | Where-Object SideStatus -ne 'NotRun').Count -eq 0) 'A run that did not happen stays NotRun everywhere, even when OAuth only was detected.'
    $dauthOnlyMissing = & $runner 'NotAuthorizedPartial' @('Get-EOSharingPolicy')
    Assert-FreeBusy ((& $module { param($r) Resolve-FreeBusyRunStatus -Process ([pscustomobject]@{ Status = 'Exited' }) -RunnerStatus $r -HtmlFound $true -ParseFailed $false -OAuthOnly } $dauthOnlyMissing) -eq 'Completed' -and
        (& $module { param($r) Resolve-FreeBusyRunStatus -Process ([pscustomobject]@{ Status = 'Exited' }) -RunnerStatus $r -HtmlFound $true -ParseFailed $false } $dauthOnlyMissing) -eq 'EXONotAuthorized') 'A missing DAuth cmdlet does not degrade the run status when DAuth does not apply.'
    $oauthOnlyReport = & $module { param($s, $f) ConvertTo-FreeBusyReportHtml -Summary $s -Findings $f -RunStatus 'EXONotAuthorized' -RunMessage 'm' } $oauthOnlySummary $oauthOnlyRows
    Assert-FreeBusy ($oauthOnlyReport -match '>1 / 2<' -and $oauthOnlyReport -match "color:#64748b;'>NotApplicable<" -and $oauthOnlyReport -match 'Not applicable: Intra Organization Connector enabled: free/busy uses OAuth, DAuth not checked \(2 side\(s\)\)') 'The mail counts applicable sides only and shows NotApplicable in gray with its reason.'
    $completedOAuthOnly = @(& $module { param($f, $r) New-FreeBusySummaryRow -Findings $f -RunStatus 'Completed' -RunMessage $null -RunnerStatus $r -DurationSeconds 51 -ServerName 'EXSRV01' -FreeBusyCheckerVersion 'v' -GenerationTime 't' -OAuthOnly } @($findings | Where-Object { $_.Auth -eq 'OAuth' -or $_.Side -eq 'ExchangeOnPremises' } | ForEach-Object { $copy = $_.PSObject.Copy(); $copy.Auth = 'OAuth'; $copy; $copy2 = $_.PSObject.Copy(); $copy2.Side = 'ExchangeOnline'; $copy2.Auth = 'OAuth'; $copy2.Status = 'OK'; $copy2 }) (& $runner 'Connected' @()))
    $completedReport = & $module { param($s) ConvertTo-FreeBusyReportHtml -Summary $s -Findings @() -RunStatus 'Completed' -RunMessage $null } $completedOAuthOnly
    Assert-FreeBusy ($completedReport -match '>2 / 2<' -and $completedReport -match 'No error, warning or not authorized check was reported\.<' -and $completedReport -notmatch 'side\(s\) that were measured') 'Not applicable sides do not make a complete run look partial.'

    # The real run of 22 October (strict mode, Windows PowerShell 5.1): OAuth only, one Exchange Online cmdlet
    # not authorized, real findings on the OAuth sides only and none on the DAuth sides.
    $realHtml = "<div><p> <span style='color: green;'>$oauthOnlyText</span></p></div>" + @"
<table>
<tr><th ColSpan='2' style='text-align:center; color:white;'>Exchange On Premise OAuth Configuration</th></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Get-User ApplicationAccount</th></tr>
<tr><td><b>Get-User ApplicationAccount</b></td><td>
<div> <b>UserAccountControl: </b> <span style='color:red'>AccountDisabled</span></div>
<div> <b>RecipientTypeDetails: </b> <span style='color:green'>LinkedUser</span></div>
</td></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Get-AuthServer</th></tr>
<tr><td><b>Get-AuthServer</b></td><td>
<div> <b>Enabled: </b> <span style='color:green'>True</span></div>
<div> <b>Note: </b> informational</div>
</td></tr>
</table>
<table>
<tr><th ColSpan='2' style='text-align:center; color:white;'><b>Exchange Online OAuth Configuration</b></th></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Get-IntraOrganizationConnector</th></tr>
<tr><td><b>Get-IntraOrganizationConnector</b></td><td>
<div> <b>Enabled: </b> <span style='color:green'>True</span></div>
</td></tr>
<tr><th ColSpan='2' style='color:white;'>Summary - Test-OAuthConnectivity</th></tr>
<tr><td><b>Test-OAuthConnectivity</b></td><td>
<div class='red'> <b> - OAuth Test was completed with Error.</b></div>
</td></tr>
</table>
"@
    $realRunner = [pscustomobject]@{ Stage = 'FreeBusyCheckerReturned'; RunStatus = $null; Message = $null; ExchangeOnline = 'NotAuthorizedPartial'; ExchangeOnlineMessage = 'Exchange Online cmdlets not available to the app: Test-EOOAuthConnectivity'; NotAuthorizedCmdlets = @('Test-EOOAuthConnectivity') }
    $realResult = & $module {
        param($h, $r)
        $oauthOnly = Test-FreeBusyOAuthOnly -Html $h
        $found = @(ConvertFrom-FreeBusyCheckerHtml -Html $h -ServerName 'EXSRV01' -FreeBusyCheckerVersion '26.06.06.1238' -GenerationTime 't')
        $found = @(Add-FreeBusyNotAuthorizedFinding -Findings $found -NotAuthorizedCmdlets @($r.NotAuthorizedCmdlets) -ServerName 'EXSRV01' -FreeBusyCheckerVersion '26.06.06.1238' -GenerationTime 't' -OAuthOnly:$oauthOnly)
        $status = Resolve-FreeBusyRunStatus -Process ([pscustomobject]@{ Status = 'Exited' }) -RunnerStatus $r -HtmlFound $true -ParseFailed $false -OAuthOnly:$oauthOnly
        $rows = @(New-FreeBusySummaryRow -Findings $found -RunStatus $status -RunMessage $r.ExchangeOnlineMessage -RunnerStatus $r -DurationSeconds 51 -ServerName 'EXSRV01' -FreeBusyCheckerVersion '26.06.06.1238' -GenerationTime 't' -OAuthOnly:$oauthOnly)
        $html = ConvertTo-FreeBusyReportHtml -Summary $rows -Findings $found -RunStatus $status -RunMessage $r.ExchangeOnlineMessage
        [pscustomobject]@{ OAuthOnly = $oauthOnly; Findings = $found; Status = $status; Summary = $rows; Html = $html }
    } $realHtml $realRunner
    $byKey = @{}; foreach ($row in $realResult.Summary) { $byKey["$($row.Side)/$($row.Auth)"] = $row }
    Assert-FreeBusy ($realResult.OAuthOnly -and $realResult.Status -eq 'EXONotAuthorized') 'The real run: OAuth only and EXONotAuthorized.'
    Assert-FreeBusy (@($realResult.Findings | Where-Object Status -eq 'Error').Count -eq 1 -and @($realResult.Findings | Where-Object { $_.Status -eq 'NotAuthorized' -and $_.Category -eq 'Test-OAuthConnectivity' }).Count -eq 1) 'The real run keeps its single Error and one NotAuthorized check.'
    Assert-FreeBusy ($byKey['ExchangeOnPremises/DAuth'].SideStatus -eq 'NotApplicable' -and $byKey['ExchangeOnline/DAuth'].SideStatus -eq 'NotApplicable' -and
        $byKey['ExchangeOnPremises/OAuth'].SideStatus -eq 'Completed' -and $byKey['ExchangeOnPremises/OAuth'].ErrorCount -eq 1 -and
        $byKey['ExchangeOnline/OAuth'].SideStatus -eq 'NotAuthorizedPartial' -and $byKey['ExchangeOnline/OAuth'].NotAuthorizedCmdlets -eq 'Test-EOOAuthConnectivity') 'The real run: DAuth NotApplicable, on-premises OAuth Completed with its error, Exchange Online OAuth partial.'
    Assert-FreeBusy ($realResult.Html -match '>2 / 2<' -and $realResult.Html -match 'Run status: <strong>EXONotAuthorized</strong>') 'The real run report shows both applicable sides tested.'
    $collectorText = [IO.File]::ReadAllText($collectorPath)
    Assert-FreeBusy ($collectorText -match 'Exchange_Hybrid_FreeBusy_Summary\.csv' -and $collectorText -match 'Exchange_Hybrid_FreeBusy_Findings\.csv') 'Both CSVs are published.'
    $receipts = [IO.File]::ReadAllText((Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365-SourceReceipts.json.txt')) | ConvertFrom-Json
    $producer = @($receipts.Producers | Where-Object Script -eq 'SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1')
    Assert-FreeBusy ($producer.Count -eq 1 -and (@($producer[0].Files) -join ',') -eq 'Exchange_Hybrid_FreeBusy_Summary.csv,Exchange_Hybrid_FreeBusy_Findings.csv') 'Both CSVs are receipt outputs.'

    # 11. Interrupted and completed runs: the collector's own catch/finally blocks, in a separate process.
    $outerTry = @($collectorAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State') })[0]
    $innerTry = $outerTry.Body.Find({ param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $n.Finally -and $n.Finally.Extent.Text.Contains('CompletionBanner') }, $true)
    # The tenant context banner is the one the entry script sees; it resolves Cancelled through the lifecycle.
    $tenantContextBanner = Get-FunctionText (Get-ScriptAst (Join-Path $SourceRoot 'Config/SmartM365-TenantContext.ps1')) 'Write-SmartM365CompletionBanner'
    $runBodies = @{
        Interrupted = 'throw [Management.Automation.PipelineStoppedException]::new()'
        Completed   = "`$script:CompletionStatus = 'Success'; Complete-FreeBusyRun -Status 'Success'; `$smartM365RunGuard.ReachedTerminal=`$true"
    }
    $runs = @{}
    foreach ($scenario in @('Interrupted', 'Completed')) {
        $eventsPath = Join-Path $tempRoot "$scenario-events.txt"
        $scenarioScript = Join-Path $tempRoot "$scenario.ps1"
        [IO.File]::WriteAllText($scenarioScript, (@(
            ". '$((Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1').Replace("'", "''"))'",
            $tenantContextBanner,
            (Get-FunctionText $collectorAst 'Write-Log'),
            "function Complete-FreeBusyRun { param(`$Status) Add-Content -LiteralPath '$($eventsPath.Replace("'", "''"))' -Value `"Receipt:`$Status`" }",
            "`$LogFile = '$((Join-Path $tempRoot "$scenario.log").Replace("'", "''"))'",
            '$global:EnableSharePointUpload = $true; $global:EnableTeamsNotifications = $true; $global:LogTextFile = $null',
            "`$ScriptName = 'SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1'; `$StartTime = Get-Date; `$script:CompletionStatus = 'Auto'; `$script:FreeBusyWarningCount = 0; `$script:FreeBusyErrorCount = 0",
            '$smartM365RunGuard = Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn',
            'try {',
            "try { $($runBodies[$scenario]) }",
            $innerTry.CatchClauses[0].Extent.Text,
            "finally $($innerTry.Finally.Extent.Text)",
            '$smartM365RunGuard.ReachedTerminal=$true',
            '$smartM365RunGuard.ReturnedNormally=$true',
            "} $($outerTry.CatchClauses[0].Extent.Text) finally $($outerTry.Finally.Extent.Text)",
            'if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}'
        ) -join "`r`n"), [Text.UTF8Encoding]::new($false))
        $ErrorActionPreference = 'Continue'
        $output = (& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $scenarioScript 2>&1 | Out-String)
        $exitCode = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        $runs[$scenario] = [pscustomobject]@{ Output = $output; ExitCode = $exitCode; Events = $(if (Test-Path -LiteralPath $eventsPath) { [IO.File]::ReadAllText($eventsPath) } else { '' }) }
    }
    $interrupted = $runs.Interrupted
    Assert-FreeBusy ($interrupted.ExitCode -eq 130) "An interrupted run exits with 130 (exit $($interrupted.ExitCode))."
    Assert-FreeBusy ($interrupted.Output -match 'Status\s+:\s+CANCELLED' -and $interrupted.Output -match 'Execution cancelled' -and $interrupted.Output -match 'Warnings\s+:\s+1' -and $interrupted.Output -match 'Errors\s+:\s+0' -and $interrupted.Output -cnotmatch 'SUCCESS') "An interrupted run shows CANCELLED with a warning, never SUCCESS: $($interrupted.Output.Trim())"
    Assert-FreeBusy ($interrupted.Output -match '\[WARN\] Run interrupted before completion \(cancelled\): no receipt or success status was written\.') 'An interrupted run logs why it is not a success.'
    Assert-FreeBusy ($interrupted.Output -notmatch 'Cleanup failed' -and $interrupted.Events -notmatch 'Receipt:') 'An interrupted run cleans up and writes no receipt.'
    $completed = $runs.Completed
    Assert-FreeBusy ($completed.ExitCode -eq 0 -and $completed.Output -match 'Status\s+:\s+SUCCESS' -and $completed.Output -match 'Errors\s+:\s+0' -and $completed.Events -match 'Receipt:Success') "A completed run stays SUCCESS with exit 0 and its receipt (exit $($completed.ExitCode))."
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("PASS: {0} Exchange hybrid free/busy checks; PowerShell {1}." -f $script:Checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCe+aUiBZQNu/s8
# D4PV22twkssgq263GlB1FG20aKkOJqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHTElcb6pI1vo0owvQAuqilk5qNHRdS45oRRY3RcSA7CMA0GCSqG
# SIb3DQEBAQUABIIBgG8ClierxBTbmUA68MFbK7Kd/ki9eI9WXk8yxi/ksKCdjBxe
# ibnvTMtsmNGjuP2iT3q0REcsFdHJsPZsMw7LD2LdY7j0xPxcwGBYoE2BhZBkdfeW
# sZMMO7VL7ucCK4fDRlYNoHWs1EUoOs4YPu6zpKHX3jblLuFVHRZYiRL1Qh1klvRK
# t76SsBGZvE1SaNiKQy/vSlmPGNG1kdp8FyWvuQDwOb2mjhZkV5EJRzTRibOwvlK4
# jRp6RA2zatcbUp2X5TzRSGkhkyq8jsN0KpNpgnAqtiIk7cXkDxAwpYkORtBgCQmc
# 8C4CiesrKJTT9G9Iucgq1JorGYX0O40AXlSyHTLlIzX+iCJ58au7oJV5v0KKeJC1
# i5ZX8BfqEkzYwwrLvL42ZzCqkWgHgDTqcPrCjs4pLv1KXMG5kV8mEo6CORTmRLth
# 4jZKXmsUfiEYhfwGlA34cYJC5WhlB8jZV04M4TDgw4z6RRKYwPlpIGqRfdV4lYYC
# sl0Bqj3NrMiPC+4RLqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEyMjMx
# MThaMC8GCSqGSIb3DQEJBDEiBCC5zvHh2eoaArov3f/qGC6i4Vf9aPO1GCE76rYe
# mcSsETANBgkqhkiG9w0BAQEFAASCAgB+PkrFT/MgmldKElZJeihIftfZ3kOpkwf7
# OV7sHPDo2DmSn2M4Mrzyl8uWYgHn9ZQNcQ4/pAEEhsy+h9ivFLZ/FjDcpVaSx/xr
# 3cHrpyIPHJwZAx26UfmLxn4XLcmA8FnagOgP7UBeTre3xVNiwImmuLdja37LMpG7
# rMo0pGsalhcZVzh4dNFy46f/+UAKJGqwIIgY0k2kYIYi6DLVb3x6BcL8uJ6eOBKv
# Rqdzc+nVh/OJsZHPQi0GFNRpnzvrsALGW1zjDZoUpLhQZ34uvN2SRrR8DP1twiiE
# RXmTb1KB/lghgtTNTz8ZcmpizUc+KTHbclTos2hpXeQ2P6ZKGJNscTgLZyFf2lHk
# IRrO3PA1znRoFoPL2R6TO2W7JdDAz6pclaPIJ5hTxvQgRC4p4roOzlrBo3FRTj6U
# 8sS5Uw4tvYFwkSl/EwRlFIyiW1irH6qBq4ZQXZsBzQ1GiSIUmAUtVJ83czVh/mDq
# 72Grvh6azZAQw1S1zkfuqWGPnORcLelM5qeTFf07ZYdTpbBkxiXwsC3JJE9hJN/u
# +18gLHTL70Bc/dsAhhjRizKSGiwnflrcIRjRdlkb1oo1mbg13KI73n5RyccrNLGf
# I3dIDNAM+fTg7s4TDNBY9256jQFftPzLPd7TNLdvGagjd0WxLBTMBsqkESs1dhLd
# eqjGm0XV1g==
# SIG # End signature block
