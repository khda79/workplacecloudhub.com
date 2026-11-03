#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic maintenance mail admission and transition tests. Production entrypoints,
authentication and mail transports are never executed. Run in PS7 and PS5.1.
.VERSION
1.0.4
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
        $names=@('Send-SmartM365GraphMail','SendEmailHtmlReport','ConvertToRecipientArray','ConvertTo-SmartM365GraphRecipient','ConvertTo-SmartM365GraphFileAttachment','Get-ModuleLocalConfigValue')
        if($edition -eq 'SmartM365.Core.psm1'){$names+='Send-SmartM365Mail'}
        $definitions=@(Get-Definitions $source $names)+@(Get-Definitions (Join-Path $core 'SmartM365-MailAdminRouting.ps1') @('Resolve-SmartM365AdminMailRoute','Get-SmartM365AdminMailControl','Set-SmartM365AdminMailControl'))+@(Get-Definitions (Join-Path $core 'SmartM365-MailMaintenance.ps1') @('Get-SmartM365MaintenanceMailRoot','Get-SmartM365MaintenanceMailState','Add-SmartM365AdminRecipients','Resolve-SmartM365MaintenanceMailRecipientRoute','Set-SmartM365SmtpMaintenanceRecipient','Resolve-SmartM365UnifiedMailRoute'))
        $runtime=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"));$runtimes+=$runtime
        & $runtime {
            param($Shared)
            # Preserve the pre-existing maintenance contract independently of the
            # administrator-routing suite, which exercises the real new policy.
            $script:UseAdminPolicy=$false;$script:AdminOnly=$false;$script:OriginalAdminAdd=${function:Add-SmartM365AdminRecipients};$script:OriginalUnified=${function:Resolve-SmartM365UnifiedMailRoute}
            function script:Get-SmartM365AdminMailControl { param($SharedDataFolderPath) [pscustomobject]@{AdministratorsOnly=$script:AdminOnly} }
            function script:Add-SmartM365AdminRecipients { param($Route,$Config,$Control) if($script:UseAdminPolicy){& $script:OriginalAdminAdd -Route $Route -Config $Config -Control $Control}else{$Route} }
            function script:Resolve-SmartM365UnifiedMailRoute {param($To,$Cc,$Bcc,$MailPurpose) if($script:UseAdminPolicy){& $script:OriginalUnified -To $To -Cc $Cc -Bcc $Bcc -MailPurpose $MailPurpose}else{$r=Resolve-SmartM365MaintenanceMailRecipientRoute -To $To -Cc $Cc -MailPurpose $MailPurpose;[pscustomobject]@{To=$r.To;Cc=$r.Cc;Bcc=$(if($r.Restricted){''}else{$Bcc})}}}
            $script:Cfg=[pscustomobject]@{TenantKey='synthetic';ProfileKey='synthetic';DataAllRootPath=(Split-Path $Shared -Parent);To='global-report@example.invalid';ErrorMailTo='global-error@example.invalid';Smart365Administrators='admin@example.invalid'}
            $script:Caller=[pscustomobject]@{To='local-report@example.invalid';ErrorMailTo='local-error@example.invalid'}
            $script:Graph=@();$script:Smtp=@();$script:FailGraph=$false;$script:BeforeAuth=$null
            $global:AppId='synthetic';$global:TenantId='synthetic';$global:Thumbprint='synthetic';$global:Thumb='synthetic'
            function script:Get-SmartM365EffectiveModuleGlobalConfig {$script:Cfg}
            function script:Resolve-SmartM365ConfigValue {param($Value) $Value}
            function script:Get-SmartM365CallerLocalConfig {$script:Caller}
            function script:Get-SmartM365ConfigurationValue {param($Name,$LocalConfig,$EffectiveConfig,$DefaultValue) $p=$LocalConfig.PSObject.Properties[$Name]; if($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value) -and $p.Value -notin @('__USE_GLOBAL__','USE_GLOBAL')){$p.Value}else{$q=$EffectiveConfig.PSObject.Properties[$Name];if($q){$q.Value}else{$DefaultValue}}}
            function script:Get-ModuleLocalConfig {$script:Caller}
            function script:WriteLog {param($Message,$Level)}
            function script:Test-SmartM365InterruptedRun {$false}
            function script:Add-SmartM365DeferredReport {param($Message,$Attachments,$MailPurpose,$DeliveryClass) $false}
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
        & $runtime {$script:UseAdminPolicy=$true;$script:AdminOnly=$false;Send-TestGraph;Send-TestReport}
        $adminNormal=& $runtime {[pscustomobject]@{Graph=$script:Graph[-1];Smtp=$script:Smtp[-1]}}
        Assert-Case ((@($adminNormal.Graph.message.toRecipients | Where-Object {$_.emailAddress.address -eq 'admin@example.invalid'}).Count -eq 1) -and @($adminNormal.Smtp.To | Where-Object {$_ -eq 'admin@example.invalid'}).Count -eq 1) 'Normal Graph/SMTP missed administrator.'
        & $runtime {$script:AdminOnly=$true;Send-TestGraph;Send-TestReport}
        $adminOnly=& $runtime {[pscustomobject]@{Graph=$script:Graph[-1];Smtp=$script:Smtp[-1]}}
        Assert-Case (@($adminOnly.Graph.message.toRecipients).Count -eq 1 -and $adminOnly.Graph.message.toRecipients[0].emailAddress.address -eq 'admin@example.invalid' -and @($adminOnly.Smtp.To).Count -eq 1 -and $adminOnly.Smtp.To[0] -eq 'admin@example.invalid' -and @($adminOnly.Smtp.Cc | Where-Object {$_}).Count -eq 0) 'Exclusive Graph/SMTP leaked recipients.'
        & $runtime {$script:UseAdminPolicy=$false;$script:AdminOnly=$false}
        # Error routing applies outside maintenance as well, even with explicit report recipients.
        foreach($mode in @('SMTP','Graph','Both')) {
            & $runtime {param($Mode) $script:FailGraph=($Mode -eq 'Both');Send-TestReport $Mode Error;$script:FailGraph=$false} $mode
            $mail=& $runtime {param($Mode) if($Mode -eq 'Graph'){$script:Graph[-1]}else{$script:Smtp[-1]}} $mode
            if($mode -eq 'Graph') {
                Assert-Case ($mail.message.toRecipients[0].emailAddress.address -eq 'local-error@example.invalid' -and @($mail.message.toRecipients).Count -eq 1) 'Inactive error Graph used report To.'
                Assert-Case (-not $mail.message.PSObject.Properties['ccRecipients']) 'Inactive error Graph leaked Cc.'
            } else {
                Assert-Case ($mail.To[0] -eq 'local-error@example.invalid' -and @($mail.To).Count -eq 1) "Inactive error $mode used report To."
                Assert-Case (@($mail.Cc|Where-Object {$_}).Count -eq 0) "Inactive error $mode leaked Cc."
            }
        }
        & $runtime {Send-TestGraph Error}
        Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'local-error@example.invalid') 'Direct Graph error ignored local ErrorMailTo.'
        $global:ScriptFailed=$true
        & $runtime {Send-TestGraph Auto}
        Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'local-error@example.invalid') 'Legacy failure outside maintenance used report To.'
        $global:ScriptFailed=$false
        foreach($marker in @('__USE_GLOBAL__','USE_GLOBAL','', ' ', $null)) {
            & $runtime {param($Marker) $script:Caller.ErrorMailTo=$Marker;Send-TestGraph Error} $marker
            Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'global-error@example.invalid') 'Local ErrorMailTo marker did not inherit the tenant ErrorMailTo.'
        }
        $beforeErrorRefusal=& $runtime {@($script:Graph).Count+@($script:Smtp).Count}
        & $runtime {$script:Caller.ErrorMailTo='';$script:Cfg.ErrorMailTo=''}
        Assert-Throws {& $runtime {Send-TestReport Graph Error}}
        Assert-Throws {& $runtime {Send-TestReport SMTP Error}}
        Assert-Throws {& $runtime {Send-TestReport Both Error}}
        Assert-Case ((& $runtime {@($script:Graph).Count+@($script:Smtp).Count}) -eq $beforeErrorRefusal) 'Missing ErrorMailTo fell back to report recipients.'
        & $runtime {$script:Caller.ErrorMailTo='local-error@example.invalid';$script:Cfg.ErrorMailTo='global-error@example.invalid'}
        $params=@{To=@('report@example.invalid');Cc=@('cc@example.invalid');Bcc=@('hidden@example.invalid')}
        & $runtime {param($P) Set-SmartM365SmtpMaintenanceRecipient $P Error} $params
        Assert-Case ($params.To[0] -eq 'local-error@example.invalid' -and -not $params.ContainsKey('Cc') -and -not $params.ContainsKey('Bcc')) 'Inactive SMTP error retained visible or blind copies.'
        if($edition -eq 'SmartM365.Core.psm1') {
            & $runtime {Send-SmartM365Mail -From 'sender@example.invalid' -To 'report@example.invalid' -Cc 'copy@example.invalid' -Subject 'Synthetic' -BodyHtml '<p>Test</p>' -SendMailMode Graph -MailPurpose Error}
            Assert-Case ((& $runtime {$script:Graph[-1].message.toRecipients[0].emailAddress.address}) -eq 'local-error@example.invalid') 'Core wrapper error outside maintenance ignored ErrorMailTo.'
        }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCConeGXI7DWZXIM
# CCOpwZdr1i5z9PYwY4s6vunN5jYLoKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD/mleWNECsZe0+ahv8wBHZ
# UfWZTFMMWz2Pga8kW6og+zANBgkqhkiG9w0BAQEFAASCAYBO8kWaZ0jHpppkIZNx
# 1aNl2jQwPBKc93xPoqdYhAn4H34gSkK4tuPcfbgztr06wwJiv0/EBcr5Dz1cxcE3
# bCMscBoUvWEBQTMkBQ7hFCr1f6ga89Vgbl7RrJSTbMq7Ywkx7cwEqriWJmLPvIe0
# mZkUVquWzXM2EDdvLykP6LoVbztlCAanIR0eqH7vjgC/2CCAt0+2nr+OU5K6MLRv
# k6YmlkMQLel5pHj7CifcGeQry0PgGNG5Ln6471xs4JaQuzvSeiv4Glycr3MeWipf
# 0nP9Jq/8ibSUyx95LKKSo5Xrs0wnBOXxCneJUwW1MZVT8IEGdQAP4l1+PoeCEZNt
# 2aV756JdxAeO39R+ytUu68E9bkgIyzgp8QYdOGsl4jXy3g7JvgSwyP6KMDo7xVQ/
# VS8PooVn8zDEWFR0dllx2/Gs0mcjP5/R8C8kYpzBwUI8wVqW9wqoYT6wtnVqLl3+
# U+RV6ajEFoUNfPJlS6KFTu5w5pa6HEgL9Z02qp8Dbsv3D5M=
# SIG # End signature block
