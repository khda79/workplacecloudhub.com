#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic maintenance mail admission and transition tests. Production entrypoints,
authentication and mail transports are never executed. Run in PS7 and PS5.1.
.VERSION
1.0.0
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Tests emulate the existing Core global runtime context and restore it afterwards.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Transport and branding mock signatures match production helpers without performing external actions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns','',Justification='Mock names match existing production helpers; local assertion names describe a test collection.')]
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-MailMaintenance-Test-'+[guid]::NewGuid().ToString('N'))
$repo=Split-Path $PSScriptRoot -Parent
$core=Join-Path $repo 'Modules/SmartM365.Core'
$orch=Join-Path $repo 'SmartInventory/Orchestrator'
$script:Cases=0
$runtimes=@()
$savedEnvRoot=$env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER
$savedEnvTenant=$env:SMARTM365_ORCHESTRATOR_TENANT
$savedContext=Get-Variable SmartM365MaintenanceMailContext -Scope Global -ErrorAction SilentlyContinue
$savedFailure=Get-Variable ScriptFailed -Scope Global -ErrorAction SilentlyContinue
$savedGlobals=@{}
foreach($name in @('AppId','TenantId','Thumbprint','Thumb','SmartM365MailHtmlFiles')) {$savedGlobals[$name]=Get-Variable $name -Scope Global -ErrorAction SilentlyContinue}
function Assert-Case {param([bool]$Condition,[string]$Message) $script:Cases++;if(-not $Condition){throw $Message}}
function Assert-Throws {param([scriptblock]$Action) $caught=$false;try{& $Action|Out-Null}catch{$caught=$true};Assert-Case $caught 'Expected safe refusal.'}
function Write-Fixture {param($Path,$Document) $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force; [IO.File]::WriteAllText($Path,($Document|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))}
function Get-Definitions {param($Path,[string[]]$Names) $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$t,[ref]$e);if($e.Count){throw ($e.Message -join ';')};foreach($name in $Names){$n=$ast.Find({param($x)$x -is [Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $name},$true);if(-not $n){throw "Missing $name"};$n.Extent.Text}}
try {
    $null=New-Item -ItemType Directory -Path $root
    $env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER=$null;$env:SMARTM365_ORCHESTRATOR_TENANT=$null
    Remove-Variable SmartM365MaintenanceMailContext -Scope Global -ErrorAction SilentlyContinue
    $global:ScriptFailed=$false
    Import-Module (Join-Path $orch 'SmartM365.Orchestrator.Maintenance.psm1') -Force
    $shared=Join-Path $root 'Orchestrator';$null=New-Item -ItemType Directory -Path $shared
    $zero=Initialize-SmartM365OrchestratorMaintenance $shared
    $paths=Get-SmartM365OrchestratorMaintenancePaths $shared
    foreach($edition in @('SmartM365.Core.psm1','Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $source=Join-Path $core $edition
        $names=@('Send-SmartM365GraphMail','SendEmailHtmlReport','ConvertToRecipientArray','ConvertTo-SmartM365GraphRecipient','ConvertTo-SmartM365GraphFileAttachment')
        if($edition -eq 'SmartM365.Core.psm1'){$names+='Send-SmartM365Mail'}
        $definitions=@(Get-Definitions $source $names)+@(Get-Content (Join-Path $core 'SmartM365-MailMaintenance.ps1'))
        $runtime=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"));$runtimes+=$runtime
        & $runtime {
            param($Shared)
            $script:Cfg=[pscustomobject]@{TenantKey='synthetic';ProfileKey='synthetic';DataAllRootPath=(Split-Path $Shared -Parent);To='global-report@example.invalid';ErrorMailTo='global-error@example.invalid'}
            $script:Caller=[pscustomobject]@{To='local-report@example.invalid';ErrorMailTo='local-error@example.invalid'}
            $script:Graph=@();$script:Smtp=@();$script:FailGraph=$false;$script:BeforeAuth=$null
            $global:AppId='synthetic';$global:TenantId='synthetic';$global:Thumbprint='synthetic';$global:Thumb='synthetic'
            function script:Get-SmartM365EffectiveModuleGlobalConfig {$script:Cfg}
            function script:Resolve-SmartM365ConfigValue {param($Value) $Value}
            function script:Get-SmartM365CallerLocalConfig {$script:Caller}
            function script:Get-ModuleLocalConfig {$script:Caller}
            function script:Get-ModuleLocalConfigValue {param($Config,$Name,$DefaultValue) if($Config.PSObject.Properties[$Name]){$Config.$Name}else{$DefaultValue}}
            function script:WriteLog {param($Message,$Level)}
            function script:Format-SmartM365MailSubject {param($Subject) $Subject}
            function script:Convert-SmartM365MailBodyLocalPathsToSharePointLinks {param($BodyHtml) $BodyHtml}
            function script:ConvertTo-SmartM365EmailBody {param($BodyHtml,$Subject,$Category) $BodyHtml}
            function script:Add-SmartM365MaxItemsSubjectPrefix {param($Subject) $Subject}
            function script:Add-SmartM365MaxItemsMailBanner {param($BodyHtml) $BodyHtml}
            function script:Save-SmartM365MailHtmlCopy {param($Subject,$BodyHtml)}
            function script:Get-SmartM365GraphAccessToken {param($AppId,$TenantId,$Thumbprint,$Purpose) if($script:BeforeAuth){& $script:BeforeAuth};'synthetic'}
            function script:Connect-SmartM365GraphAppOnly {param($AppId,$TenantId,$Thumbprint,$Purpose) if($script:BeforeAuth){& $script:BeforeAuth};$true}
            function script:Test-SmartM365ConfiguredValue {param($Value) [bool]$Value}
            function script:Invoke-MgGraphRequest {param($Method,$Uri,$Body,$ContentType) if($script:FailGraph){throw 'Synthetic transport outage'};$script:Graph+=($Body|ConvertFrom-Json)}
            function script:Invoke-SmartM365GraphRestWithRetry {param($Method,$Uri,$Body,$ContentType,$Operation) Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType $ContentType}
            function script:Send-MailMessage {param($SmtpServer,$Port,$From,$To,$Cc,$Subject,$Body,$BodyAsHtml,$Attachments,$ErrorAction) $script:Smtp+=@{To=@($To);Cc=@($Cc)}}
            function script:Send-TestGraph {param($Purpose='Report',$To='local-report@example.invalid') Send-SmartM365GraphMail -From 'sender@example.invalid' -To $To -Cc 'local-cc@example.invalid' -Subject 'Synthetic' -BodyHtml '<p>Test</p>' -SkipHtmlCopy -MailPurpose $Purpose}
            function script:Send-TestReport {param($Mode='SMTP',$Purpose='Report') SendEmailHtmlReport -From 'sender@example.invalid' -To 'local-report@example.invalid' -Cc 'local-cc@example.invalid' -SmtpServer 'synthetic.invalid' -SendMailMode $Mode -Subject 'Synthetic' -BodyHtml '<p>Test</p>' -MailPurpose $Purpose}
        } $shared
        Write-Fixture $paths.State $zero
        & $runtime {Send-TestGraph;Send-TestReport}
        $normal=& $runtime {[pscustomobject]@{Graph=$script:Graph[-1];Smtp=$script:Smtp[-1]}}
        Assert-Case ($normal.Graph.message.toRecipients[0].emailAddress.address -eq 'local-report@example.invalid') 'Inactive Graph changed To.'
        Assert-Case ($normal.Graph.message.ccRecipients[0].emailAddress.address -eq 'local-cc@example.invalid') 'Inactive Graph changed Cc.'
        Assert-Case ($normal.Smtp.To[0] -eq 'local-report@example.invalid' -and $normal.Smtp.Cc[0] -eq 'local-cc@example.invalid') 'Inactive SMTP changed recipients.'
        $active=$zero.PSObject.Copy();$active.Enabled=$true;$active.Revision=1
        $active.ChangedAtUtc=[datetime]::UtcNow.ToString('o');$active.ChangedBy='synthetic';$active.ChangedFromServer='SERVER-A';$active.Reason='Synthetic policy test'
        Write-Fixture $paths.State $active
        foreach($purpose in @('Report','Error','Maintenance','Auto')) {
            $to=if($purpose -eq 'Auto'){'local-error@example.invalid'}else{'local-report@example.invalid'}
            & $runtime {param($P,$To) Send-TestGraph $P $To} $purpose $to
            $mail=& $runtime {$script:Graph[-1]}
            $expected=if($purpose -in @('Error','Auto')){'global-error@example.invalid'}else{'global-report@example.invalid'}
            Assert-Case ($mail.message.toRecipients[0].emailAddress.address -eq $expected) "Active $purpose Graph recipient is not global."
            Assert-Case (-not $mail.message.PSObject.Properties['ccRecipients'] -and @($mail.message.toRecipients).Count -eq 1) 'Restricted Graph leaked a Cc or extra recipient.'
        }
        foreach($mode in @('SMTP','Both','Graph')) {
            & $runtime {param($Mode) $script:FailGraph=($Mode -eq 'Both');Send-TestReport $Mode Error;$script:FailGraph=$false} $mode
            $actual=& $runtime {param($Mode) if($Mode -eq 'Graph'){$script:Graph[-1].message.toRecipients[0].emailAddress.address}else{$script:Smtp[-1].To[0]}} $mode
            Assert-Case ($actual -eq 'global-error@example.invalid') "Restricted $mode report did not use global ErrorMailTo."
        }
        $global:ScriptFailed=$true
        & $runtime {Send-TestGraph Auto}
        Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'global-error@example.invalid') 'Legacy terminal failure did not use global ErrorMailTo.'
        $global:ScriptFailed=$false
        if($edition -eq 'SmartM365.Core.psm1') {
            & $runtime {Send-SmartM365Mail -From 'sender@example.invalid' -To 'local-report@example.invalid' -Cc 'local-cc@example.invalid' -Subject 'Synthetic wrapper' -BodyHtml '<p>Test</p>' -SendMailMode Graph -MailPurpose Error}
            Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'global-error@example.invalid') 'Core mail wrapper lost explicit error purpose.'
        }
        $smtp=& $runtime {$script:Smtp[-1]}
        Assert-Case (@($smtp.Cc|Where-Object {$_}).Count -eq 0) 'Restricted SMTP retained Cc.'
        $params=@{To=@('local-report@example.invalid');Cc=@('cc@example.invalid');Bcc=@('hidden@example.invalid')}
        & $runtime {param($P) Set-SmartM365SmtpMaintenanceRecipient $P Report} $params
        Assert-Case (-not $params.ContainsKey('Cc') -and -not $params.ContainsKey('Bcc')) 'Restricted SMTP retained blind recipients.'
        Write-Fixture $paths.State $zero
        & $runtime {Send-TestGraph Maintenance;Send-TestGraph Report}
        $exit=& $runtime {$script:Graph[-2]};$restored=& $runtime {$script:Graph[-1]}
        Assert-Case ($exit.message.toRecipients[0].emailAddress.address -eq 'global-report@example.invalid') 'Maintenance exit notification used local recipients.'
        Assert-Case ($restored.message.toRecipients[0].emailAddress.address -eq 'local-report@example.invalid') 'Normal routing did not return on exit.'
        # A collector launched before maintenance must re-read after preparation/authentication.
        & $runtime {param($P,$A) $script:BeforeAuth={[IO.File]::WriteAllText($P,($A|ConvertTo-Json),[Text.UTF8Encoding]::new($false))}.GetNewClosure();Send-TestGraph;$script:BeforeAuth=$null} $paths.State $active
        $late=& $runtime {$script:Graph[-1]}
        Assert-Case ($late.message.toRecipients[0].emailAddress.address -eq 'global-report@example.invalid') 'Late activation was not honored.'
        $before=& $runtime {@($script:Graph).Count+@($script:Smtp).Count}
        Write-Fixture $paths.State ([pscustomobject]@{SchemaVersion=1;Revision=1;Enabled='false'})
        Assert-Throws {& $runtime {Send-TestGraph}}
        Assert-Throws {& $runtime {Send-TestReport Both}}
        Write-Fixture $paths.State ([pscustomobject]@{SchemaVersion=1;Revision=0;Enabled=$false})
        Assert-Throws {& $runtime {Send-TestGraph}}
        Remove-Item -LiteralPath $paths.State
        Assert-Throws {& $runtime {Send-TestGraph}}
        Write-Fixture $paths.State $active
        & $runtime {$script:Cfg.To=''}
        Assert-Throws {& $runtime {Send-TestGraph}}
        & $runtime {$script:Cfg.To='global-report@example.invalid';$script:Cfg.ErrorMailTo='__USE_GLOBAL__'}
        Assert-Throws {& $runtime {Send-TestGraph Error}}
        & $runtime {$script:Cfg.ErrorMailTo='global-error@example.invalid'}
        $env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER=$shared;$env:SMARTM365_ORCHESTRATOR_TENANT='different'
        Assert-Throws {& $runtime {Send-TestGraph}}
        $env:SMARTM365_ORCHESTRATOR_TENANT='synthetic'
        Assert-Case ((& $runtime {Get-SmartM365MaintenanceMailRoot $script:Cfg}) -eq $shared) 'Child process live shared root was not used.'
        $env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER=$null;$env:SMARTM365_ORCHESTRATOR_TENANT=$null
        $after=& $runtime {@($script:Graph).Count+@($script:Smtp).Count}
        Assert-Case ($after -eq $before) 'Unsafe admission reached a transport.'
        Assert-Case ((& $runtime {$script:Caller.To}) -eq 'local-report@example.invalid') 'Caller configuration was rewritten.'
        $global:SmartM365MaintenanceMailContext=[pscustomobject]@{TenantKey='synthetic';ProfileKey='synthetic';SharedDataFolderPath=$shared}
        Assert-Case ((& $runtime {Get-SmartM365MaintenanceMailRoot $script:Cfg}) -eq $shared) 'Orchestrator live context was not used.'
        Remove-Variable SmartM365MaintenanceMailContext -Scope Global
        & $runtime {param($Root) $script:Cfg|Add-Member -NotePropertyName OrchestratorSharedDataFolderPath -NotePropertyValue $Root -Force} $shared
        Assert-Case ((& $runtime {Get-SmartM365MaintenanceMailRoot $script:Cfg}) -eq $shared) 'Direct collector custom shared root was not used.'
    }
    # Real maintenance setters and audit, but an in-memory Boolean delivery callback.
    Write-Fixture $paths.State $zero
    $notifications=New-Object Collections.Generic.List[object]
    $send={param($Transition) $gate=Enter-SmartM365OrchestratorMaintenanceGate $shared -TimeoutSeconds 0;$gate.Dispose();$notifications.Add($Transition);$true}.GetNewClosure()
    $init=Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send -Initialize
    Assert-Case ($init.Status -eq 'Initialized' -and $init.Sent -eq 0) 'Initial deployment replayed old mail.'
    $cluster=[pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5}
    Write-Fixture (Join-Path $shared 'Config/Orchestrator-Cluster.json.txt') $cluster
    function Write-TestHeartbeat {param($State) Write-Fixture (Join-Path $shared 'SERVER-A/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{ScriptVersion='1.5.40';RunningJobs=@();Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';MaintenanceProtocol=1;MaintenanceHealthy=$true;MaintenanceRevision=$State.Revision;MaintenanceEnabled=$State.Enabled})}
    Write-TestHeartbeat $zero
    $one=Set-SmartM365OrchestratorMaintenance $shared $true 0 'Synthetic enable'
    Write-TestHeartbeat $one
    $two=Set-SmartM365OrchestratorMaintenance $shared $false 1 'Synthetic disable'
    $result=Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send
    Assert-Case ($result.Sent -eq 2 -and $notifications.Count -eq 2 -and $notifications[0].Enabled -and -not $notifications[1].Enabled) 'Rapid enable/disable did not send both notifications.'
    Assert-Case ((Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send).Sent -eq 0) 'Another worker replayed a delivered notification.'
    $noop=Set-SmartM365OrchestratorMaintenance $shared $false 2 'Same state'
    Assert-Case ($noop.Revision -eq 2 -and (Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send).Sent -eq 0) 'Idempotent control produced mail.'
    $mailGate=[IO.File]::Open($paths.MailGate,'Open','ReadWrite','None')
    try{Assert-Case ((Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send).Status -eq 'Busy') 'Concurrent worker bypassed mail lock.'}finally{$mailGate.Dispose()}
    Write-TestHeartbeat $two
    $null=Set-SmartM365OrchestratorMaintenance $shared $true 2 'Synthetic failure'
    $now=[datetime]::UtcNow
    $retry=Invoke-SmartM365OrchestratorMaintenanceNotification $shared {param($Transition) $false} -Now $now
    Assert-Case ($retry.Status -eq 'RetryPending' -and (Get-SmartM365OrchestratorMaintenanceState $shared).Enabled) 'Transport failure undid maintenance.'
    $retry=Invoke-SmartM365OrchestratorMaintenanceNotification $shared {throw 'Should not retry yet'} -Now $now.AddSeconds(10)
    Assert-Case ($retry.Status -eq 'RetryPending') 'Retry ignored backoff.'
    Assert-Throws {Invoke-SmartM365OrchestratorMaintenanceNotification $shared {throw 'Synthetic outage'} -Now $now.AddSeconds(61)}
    Assert-Case ((Get-SmartM365OrchestratorMaintenanceState $shared).Revision -eq 3) 'Thrown mail error changed control.'
    Assert-Case ((Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send -Now $now.AddSeconds(122)).Sent -eq 1) 'Failed transition was not retried.'
    $null=Set-SmartM365OrchestratorMaintenance $shared $false 3 'Synthetic missing final audit'
    $audit=@(Import-Csv $paths.Audit|Where-Object { -not ($_.Revision -eq '4' -and $_.Outcome -eq 'Published') })
    $audit|Export-Csv -LiteralPath $paths.Audit -NoTypeInformation
    Assert-Case ((Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send).Sent -eq 1) 'Authoritative latest transition was lost after failed audit append.'
    $checkpoint=(Read-SmartM365JsonDocument $paths.MailCheckpoint).Document
    $checkpoint.LastSentRevision=3
    Write-Fixture $paths.MailCheckpoint $checkpoint
    Assert-Throws {Invoke-SmartM365OrchestratorMaintenanceNotification $shared {'Unexpected output';$true} -RetrySeconds 0}
    Assert-Case ((Read-SmartM365JsonDocument $paths.MailCheckpoint).Document.LastSentRevision -eq 3) 'Non-Boolean callback falsely recorded delivery.'
    # Exercise the actual resident notification callback and purpose propagation.
    $residentDefinitions=@(Get-Definitions (Join-Path $orch 'SmartM365-Inventory-Orchestrator.ps1') @('Update-OrchestratorMaintenance','Send-OrchestratorMail'))
    $resident=New-Module -ScriptBlock ([scriptblock]::Create($residentDefinitions -join "`n"));$runtimes+=$resident
    $global:SmartM365MailHtmlFiles=@()
    & $resident {
        param($Root)
        $script:Tenant='synthetic';$script:MaintenanceHealthy=$true;$script:MaintenanceControl=$null;$script:StatePersistenceHealthy=$true
        $script:Manifest=@{OrderedJobs=@()};$script:Mail=@();$script:BodySummary=$null
        $script:Settings=@{SharedDataFolderPath=$Root;MailEnabled=$true;MailTo='local-report@example.invalid';ErrorMailTo='local-error@example.invalid';MailCc='cc@example.invalid';MailBcc='';MailFrom='sender@example.invalid';SmtpServer='';SmtpPort=25;SendMailMode='Graph'}
        function script:Write-OrchestratorLog {param($Message,$Level)}
        function script:Write-OrchestratorRuntimeUpdateWarning {param($Key,$Message,$Now) throw $Message}
        function script:Format-SmartM365MailSubject {param($Subject,[switch]$Orchestrator) $Subject}
        function script:New-SmartM365EmailBody {param($Title,$Category,$Message,[hashtable]$SummaryData) $script:BodySummary=$SummaryData;'<p>Synthetic resident mail</p>'}
        function script:Send-SmartM365Mail {param($SmtpServer,$SmtpPort,$SendMailMode,$From,$To,$Cc,$Subject,$BodyHtml,[switch]$BodyAsHtml,[switch]$HighPriority,$MailPurpose,$ErrorAction) $script:Mail+=@{Subject=$Subject;Purpose=$MailPurpose}}
        function script:Invoke-OrchestratorMailHtmlUploads {param($PreviousCount)}
    } $shared
    & $resident {Update-OrchestratorMaintenance -Now ([datetime]::UtcNow.AddSeconds(61))}
    $callback=& $resident {[pscustomobject]@{Mail=$script:Mail;Summary=$script:BodySummary}}
    Assert-Case ($callback.Mail.Count -eq 1 -and $callback.Mail[0].Subject -eq 'Orchestrator maintenance disabled' -and $callback.Mail[0].Purpose -eq 'Maintenance') 'Resident callback lost transition purpose or failed to send.'
    Assert-Case ($callback.Summary.Revision -eq 4 -and $callback.Summary.Reason -eq 'Synthetic missing final audit') 'Resident mail omitted transition evidence.'
    $checkpoint.LastSentRevision=0
    Write-Fixture $paths.MailCheckpoint $checkpoint
    $audit=@(Import-Csv $paths.Audit|Where-Object { -not ($_.Revision -eq '1' -and $_.Outcome -eq 'Published') });$audit|Export-Csv $paths.Audit -NoTypeInformation
    Assert-Throws {Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send}
    Assert-Case ($notifications.Count -eq 4) 'Revision gap sent an unsupported notification.'
    Write-Fixture $paths.MailCheckpoint ([pscustomobject]@{SchemaVersion=1;LastSentRevision='bad'})
    Assert-Throws {Invoke-SmartM365OrchestratorMaintenanceNotification $shared $send}
    Write-Output ("PASS: {0} offline maintenance mail checks; PS {1}; no real mail or maintenance change." -f $script:Cases,$PSVersionTable.PSVersion)
}
finally {
    foreach($runtime in $runtimes){Remove-Module $runtime -ErrorAction SilentlyContinue}
    $env:SMARTM365_ORCHESTRATOR_SHARED_DATA_FOLDER=$savedEnvRoot;$env:SMARTM365_ORCHESTRATOR_TENANT=$savedEnvTenant
    if($savedContext){$global:SmartM365MaintenanceMailContext=$savedContext.Value}else{Remove-Variable SmartM365MaintenanceMailContext -Scope Global -ErrorAction SilentlyContinue}
    if($savedFailure){$global:ScriptFailed=$savedFailure.Value}else{Remove-Variable ScriptFailed -Scope Global -ErrorAction SilentlyContinue}
    foreach($name in $savedGlobals.Keys){if($savedGlobals[$name]){Set-Variable $name -Value $savedGlobals[$name].Value -Scope Global}else{Remove-Variable $name -Scope Global -ErrorAction SilentlyContinue}}
    $resolved=[IO.Path]::GetFullPath($root);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if(-not $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartM365-MailMaintenance-Test-*'){throw 'Unsafe fixture cleanup target.'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6CId2iLCZcFzC
# q/+lgcKEvZFQa3XpPcK78Ysi0rvlUaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDBSXNHbOFv/aDCOZa5+FCt2bYU3yO5gaDyYWREVvMYRMA0GCSqG
# SIb3DQEBAQUABIIBgDGwq6nppb2fjF0WVcF2uyx1g5SHWB03y2gp7BhmJCu+5yos
# xB118yQvDfK1ipvSCb7PbarIsM6xVoFydn91DBQ2J4H2MIFjD053KyOhiGy7o19/
# rHYsDMeJIBy0OWrpLh8Fm7X/j3VAcd7RDpW6aUAdpiIo1abaKGuO00RpwQSroDIE
# QDfows/at4MSCxO7hOtY9DqnEnsASDHH0lZchE/jnHs6o4/1Y5ze5bLDt61ak0LV
# uoaXlIhd3SLLl6vMnvl7UPth7deMTgAMxJOAzMuoPzchve0CuuS9+VxHT9X+7Hsn
# lwaH6m2cWoTqrqPOp8A6IxROTDR8Y93WrNE8HmZicOE3miVQC4lL71KY35jPGGVF
# 5YJfHJCnBm15wG0KVmnoSoJFQVLFY2mMmYU/im5TyAKWYvNqSMfrsrBJa949c7+T
# oc6UUgvjXgG8srsDRh6mHmG+GSWpeE5j3Chd4HCX51DKS2b9zOnHFGo0CG2Rk7Yi
# +/osCBMaVwKXFYzJRKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQxODM3
# NDVaMC8GCSqGSIb3DQEJBDEiBCDFLhIlYFw1JiImHUCaZ+FVYZ9uSFD+eTodf/lr
# stjQ3zANBgkqhkiG9w0BAQEFAASCAgB966HjTQeR+xqIweGJCeocbpEUK8KzV7Tt
# sWcvPkuKnX4h18Xo7COSAkdaLyTPqudQQ59yNwklBFU2pQ4LsOB1yAiJTCWMbA8O
# dPtrSeSBLRhU7Yw/DS1h+w9yY6E6WwoInB8X7G2soIgJEQk1tTRvCFtrw5LrBgwe
# QR3/BdIxIqgruuxqOVvMBWMSyLT+HaZ23gfUffonErkrkoxxy7QY0alLUAfPXNKW
# 45aP7vjQX58PGtRkCpVXUunanCk3G6tfjJx+Gc+gak3VtE/IyCb8iAeximM52/ij
# zBnA4Z0cQrSztOJ2PYJscrVMBDbjyDiQz7Lybl54Z8A8J7Q8Na+5ZJv4W06V+xJS
# OMc0L/+F2fNEO2XI+gaSQoZLxHdtKuak8m4oVF2xPH8EEqNUjgyKRXX5TKpqXptp
# Qncr9Pk4Q+5cfhD0FAAGeYga/d0R/w+jUFIDcMOdz2oWlpZfURzWAV884C3sGlPw
# efFykOgd4TPL/4E+JRL9CLynii2bgUVJkUI3o5ha+9fVkydLN98rskcaKCERXedV
# 1p3s2OVn7BBhdOUI80V8BmDRtdxxsm+yPntw2EwKAx0l/C2JCXTj4VOK2PSSYeGh
# F56Agm6FIRl+ysO2o83LuPNBSOht5axEscHzspj52kVBEv7zT1hhG46BRsP/7i5I
# Myp6eDw78w==
# SIG # End signature block
