#Requires -Version 7.0
<#
.SYNOPSIS
Offline migration schedules, recipient provenance, worker and WPF integration checks.
.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$TestRoot=(Join-Path $env:TEMP ('SmartM365-MigrationSelection-'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
$app=Join-Path (Split-Path $PSScriptRoot -Parent) 'Exchange/ExchangeUserNotifications'
Import-Module (Join-Path $app 'SmartM365.ExchangeMigrationSelection.psm1') -Force
[void][IO.Directory]::CreateDirectory($TestRoot)
$script:checks=0
function Assert-Selection {param([bool]$Condition,[string]$Message) if(-not $Condition){throw "FAILED: $Message"};$script:checks++}
function Assert-SelectionThrows {param([scriptblock]$Action,[string]$Pattern) $caught=$null;try{& $Action | Out-Null}catch{$caught=$_};Assert-Selection ($null -ne $caught -and $caught.Exception.Message -match $Pattern) "Expected failure: $Pattern"}
$zone='Romance Standard Time'
$tenant='11111111-1111-4111-8111-111111111111'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{CompleteAfterUTC=[datetime]'2030-01-01T23:30:00'}) 'Fixture' $zone
Assert-Selection ($s.Known -and $s.EffectiveDate -eq '2030-01-02') 'UTC property converts to the next Paris calendar day'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{CompleteAfter='2030-07-01T23:30:00Z'}) 'Fixture' $zone
Assert-Selection ($s.Known -and $s.ScheduledLocal -like '*01:30*') 'Summer offset is applied'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{CompleteAfter='2030-01-01T23:30:00'}) 'Fixture' $zone
Assert-Selection (-not $s.Known) 'An unspecified time zone must remain unknown'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{StartAfter='2030-01-01T23:30:00Z'}) 'Fixture' $zone
Assert-Selection (-not $s.Known -and -not $s.EffectiveDate) 'Sync start is never a completion date'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{CompleteAfterUTC='9999-12-31T00:00:00Z'}) 'Fixture' $zone
Assert-Selection (-not $s.Known) 'Infinite scheduling sentinel is not a notification date'
$s=Get-ExoSelectionSchedule ([pscustomobject]@{CompleteAfter='2030-10-27T02:30:00';TimeZone=$zone}) 'Fixture' $zone
Assert-Selection (-not $s.Known) 'Ambiguous daylight-saving wall time stays unknown'
function New-SelectionFixture {
    param([string]$Email='user@example.invalid',[string]$Batch='Batch-A',[string]$Date='2030-01-01T23:30:00Z',[string]$Status='Synced',$Stats=([pscustomobject]@{CompleteAfter=$null}))
    [pscustomobject]@{Batch=[pscustomobject]@{Identity=$Batch;CompleteAfterUTC=$Date};User=[pscustomobject]@{EmailAddress=$Email;Status=$Status};Statistics=$Stats}
}
$rows=@(ConvertTo-ExoSelectionRecipient -Records @((New-SelectionFixture)) -DisplayTimeZoneId $zone)
Assert-Selection ($rows.Count -eq 1 -and $rows[0].Selected -and $rows[0].DateSource -eq 'Batch CompleteAfter') 'Verified absence of user override inherits batch schedule'
$override=New-SelectionFixture -Stats ([pscustomobject]@{CompleteAfterUTC='2030-01-04T20:00:00Z'})
$overrideRows=@(ConvertTo-ExoSelectionRecipient @($override) $zone)
Assert-Selection ($overrideRows[0].EffectiveDate -eq '2030-01-04' -and $overrideRows[0].DateSource -eq 'User CompleteAfter') 'User schedule takes precedence over batch'
$unknown=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture -Stats $null)) $zone)
Assert-Selection (-not $unknown[0].DateConfirmed -and -not $unknown[0].Selected -and $unknown[0].EffectiveDate) 'Unverified individual override is only a proposal'
$duplicates=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture),(New-SelectionFixture -Batch 'Batch-B')) $zone)
Assert-Selection ($duplicates.Count -eq 1 -and $duplicates[0].BatchIds -eq 'Batch-A;Batch-B' -and -not $duplicates[0].Conflict) 'Same-date duplicates merge with batch provenance'
$conflict=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture),(New-SelectionFixture -Batch 'Batch-B' -Date '2030-02-01T12:00:00Z')) $zone)
Assert-Selection ($conflict.Count -eq 1 -and $conflict[0].Conflict -and -not $conflict[0].Selected) 'Conflicting dates require an explicit decision'
$done=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture -Status 'Completed')) $zone)
Assert-Selection ($done[0].Completed -and -not $done[0].Selected) 'Completed migrations are excluded'
$sent=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture)) $zone @{'user@example.invalid'=$true})
Assert-Selection ($sent[0].AlreadySent -and -not $sent[0].Selected) 'Existing registry is reflected before selection'
Assert-SelectionThrows {Assert-ExoMigrationSelection $unknown $tenant} 'Confirm'
Assert-SelectionThrows {Assert-ExoMigrationSelection $done $tenant} 'Completed'
Assert-SelectionThrows {Assert-ExoMigrationSelection $sent $tenant} 'already notified'
Assert-ExoMigrationSelection $sent $tenant -ForceSend
$csv=Join-Path $TestRoot 'fictional.csv'
Export-ExoMigrationSelection -Rows $rows -TenantId $tenant -TenantKey 'fixture' -Path $csv -SourceReadAtUtc ([datetime]::UtcNow.ToString('o')) | Out-Null
$csvRows=@(Import-Csv $csv)
Assert-ExoMigrationSelectionCsv -Rows $csvRows -TenantId $tenant -TenantKey 'fixture'
Assert-Selection ($csvRows[0].EffectiveDate -eq '2030-01-02' -and $csvRows[0].BatchIds -eq 'Batch-A') 'Export preserves per-row dates and batch provenance'
Assert-SelectionThrows {Assert-ExoMigrationSelectionCsv $csvRows '22222222-2222-4222-8222-222222222222'} 'tenant'
Assert-SelectionThrows {Assert-ExoMigrationSelectionCsv $csvRows $tenant -TenantKey 'other'} 'profile'
$csvRows[0].SourceReadAtUtc=[datetime]::UtcNow.AddMinutes(-16).ToString('o')
Assert-SelectionThrows {Assert-ExoMigrationSelectionCsv $csvRows $tenant} 'expired'
$csvRows[0].SourceReadAtUtc=[datetime]::UtcNow.ToString('o');$csvRows[0].EffectiveDate=''
Assert-SelectionThrows {Assert-ExoMigrationSelectionCsv $csvRows $tenant} 'date'
Assert-SelectionThrows {ConvertTo-ExoSelectionRecipient @((New-SelectionFixture -Email 'bad')) $zone} 'invalid email'

# Run the actual worker control flow with all tenant commands replaced locally.
$tokens=$null;$errors=$null
$workerAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $app 'SmartM365-ExchangeMigration-SelectionWorker.ps1'),[ref]$tokens,[ref]$errors)
$workerTry=$workerAst.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst]}|Select-Object -Last 1
& {
    $TenantId=$tenant;$connected=$false;$responses=[Collections.Generic.List[object]]::new();$disconnects=0
    function Import-Module {param($Name,$MinimumVersion,$ErrorAction)}
    function Connect-ExchangeOnline {param($ShowBanner,$ErrorAction)}
    function Disconnect-ExchangeOnline {param($Confirm,$ErrorAction) $script:disconnectFixture=$true}
    function Get-ConnectionInformation { [pscustomobject]@{State='Connected';TenantID=$TenantId;UserPrincipalName='operator@example.invalid';AppCertificateAuthentication=$false} }
    function Get-MigrationBatch {param($ResultSize,$ErrorAction) [pscustomobject]@{Identity='Batch-A';Status='Synced';TotalCount=1;CompleteAfterUTC=[datetime]'2030-01-01T23:30:00Z'} }
    function Get-MigrationUser {param($BatchId,$ResultSize,$ErrorAction) [pscustomobject]@{Identity='user@example.invalid';EmailAddress='user@example.invalid';Status='Synced'} }
    function Get-MigrationUserStatistics {param($Identity,$ErrorAction) [pscustomobject]@{CompleteAfter=$null} }
    function Write-SelectionResponse {param($Kind,$Data) $responses.Add([pscustomobject]@{Kind=$Kind;Data=$Data})}
    $originalIn=[Console]::In
    $inputLines=@{Operation='Batches'}|ConvertTo-Json -Compress
    $inputLines+="`n"+(@{Operation='Recipients';BatchIds=@('Batch-A');TimeZoneId=$zone}|ConvertTo-Json -Compress)+"`n"+(@{Operation='Close'}|ConvertTo-Json -Compress)
    try {
        [Console]::SetIn([IO.StringReader]::new($inputLines))
        . ([scriptblock]::Create($workerTry.Extent.Text))
    }
    finally {[Console]::SetIn($originalIn)}
    Assert-Selection (($responses.Kind -join ',') -eq 'Connected,Batches,Recipients') 'Worker connects, lists and loads through read-only fixture commands'
    Assert-Selection ($responses[2].Data.Rows.Count -eq 1 -and $responses[2].Data.Rows[0].EffectiveDate -eq '2030-01-02') 'Worker passes authoritative date normalization to the GUI'
    Assert-Selection $script:disconnectFixture 'Owned worker disconnects during cleanup'
}
foreach($failure in @('WrongTenant','MissingMembership','CountMismatch','StatisticsDenied','PartialBatchFailure')) {
    & {
        $TenantId=$tenant;$connected=$false;$responses=[Collections.Generic.List[object]]::new()
        function Import-Module {param($Name,$MinimumVersion,$ErrorAction)}
        function Connect-ExchangeOnline {param($ShowBanner,$ErrorAction)}
        function Disconnect-ExchangeOnline {param($Confirm,$ErrorAction)}
        function Get-ConnectionInformation { [pscustomobject]@{State='Connected';TenantID=$(if($failure -eq 'WrongTenant'){'22222222-2222-4222-8222-222222222222'}else{$TenantId});UserPrincipalName='operator@example.invalid'} }
        function Get-MigrationBatch {param($ResultSize,$ErrorAction) foreach($id in @('Batch-A','Batch-B')){[pscustomobject]@{Identity=$id;Status='Synced';TotalCount=$(if($failure -eq 'CountMismatch'){2}else{1});CompleteAfterUTC=[datetime]'2030-01-01T23:30:00Z'}} }
        function Get-MigrationUser {
            param($BatchId,$ResultSize,$ErrorAction)
            if($failure -eq 'MissingMembership'){return}
            if($failure -eq 'PartialBatchFailure' -and $BatchId -eq 'Batch-B'){throw 'Synthetic second-batch query failure'}
            [pscustomobject]@{Identity='user@example.invalid';EmailAddress='user@example.invalid';Status='Synced'}
        }
        function Get-MigrationUserStatistics {param($Identity,$ErrorAction) if($failure -eq 'StatisticsDenied'){throw 'Synthetic denied schedule'};[pscustomobject]@{CompleteAfter=$null} }
        function Write-SelectionResponse {param($Kind,$Data) $responses.Add([pscustomobject]@{Kind=$Kind;Data=$Data})}
        $inputLines=(@{Operation='Recipients';BatchIds=@('Batch-A','Batch-B');TimeZoneId=$zone}|ConvertTo-Json -Compress)+"`n"+(@{Operation='Close'}|ConvertTo-Json -Compress)
        $originalIn=[Console]::In
        try {
            [Console]::SetIn([IO.StringReader]::new($inputLines))
            # Replace only the process exit boundary so tenant rejection cannot terminate this suite.
            try {. ([scriptblock]::Create($workerTry.Extent.Text.Replace('exit 1',"throw 'Fixture worker terminated'")))}catch{if($_.Exception.Message -ne 'Fixture worker terminated'){throw}}
        }
        finally {[Console]::SetIn($originalIn)}
        if($failure -eq 'StatisticsDenied') {
            $result=@($responses|Where-Object Kind -EQ 'Recipients')
            Assert-Selection ($result.Count -eq 1 -and -not $result[0].Data.Rows[0].DateConfirmed -and -not $result[0].Data.Rows[0].Selected) 'Denied individual schedules require confirmation, never automatic selection'
        }
        else {
            Assert-Selection (@($responses|Where-Object Kind -EQ 'Error').Count -ge 1 -and @($responses|Where-Object Kind -EQ 'Recipients').Count -eq 0) "$failure cannot yield a usable partial audience"
        }
    }
}

# Test WPF/XAML and real command construction without running GUI startup.
Add-Type -AssemblyName PresentationFramework
. (Join-Path $app 'SmartM365.GuiDemo.ps1')
$gui=Join-Path $app 'SmartM365-ExchangeUserNotifications-GUI.ps1'
$markup=Get-SmartM365DemoMarkup $gui
$window=New-SmartM365DemoView $markup
try {
    Assert-Selection ($window.WindowState -eq 'Maximized') 'Operational window starts maximized'
    foreach($name in @('ExoBatchesGrid','ExoRecipientsGrid','ExoConnectButton','ExoUseSelectionButton')) {Assert-Selection ($null -ne $window.FindName($name)) "Original XAML contains $name"}
    $ast=[Management.Automation.Language.Parser]::ParseFile($gui,[ref]$tokens,[ref]$errors)
    $definition=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-CommandArgument'},$true)
    & {
        Set-StrictMode -Off
        function Get-CommandArgument {}
        Set-Item Function:Get-CommandArgument $definition.Body.GetScriptBlock()
        function Get-SelectedCampaign {[pscustomobject]@{ScriptPath='fixture.ps1';SupportsFromList=$false;RecipientPathParameter='-RecipientsPath';SupportsEffectiveDate=$true;SupportsSkipConfirmation=$false}}
        function Get-SelectedTenantProfile {'fixture'}
        function Test-MigrationSelectionSource {$true}
        function Get-RunMode {'DryRun'}
        $script:RecipientPathBox=[pscustomobject]@{Text='legacy.csv'}
        $script:MigrationSelectionPath=$csv;$script:ForceLanguageCombo=[pscustomobject]@{SelectedItem='(Auto)'};$script:AutoLanguageLabel='(Auto)'
        $script:EffectiveDatePicker=[pscustomobject]@{SelectedDate=[nullable[datetime]]([datetime]'2030-05-01')}
        $script:MailSendModeCombo=[pscustomobject]@{SelectedItem='Disabled'};$script:ExchangeManagementModeCombo=[pscustomobject]@{SelectedItem='Auto'}
        foreach($name in @('ForceSendCheck','NoSummaryCheck','TeamsUserMessageCheck','SkipConfirmationCheck')){Set-Variable -Scope Script -Name $name -Value ([pscustomobject]@{IsChecked=$false})}
        $argsOut=Get-CommandArgument
        Assert-Selection ($argsOut -contains '-MigrationBatchSelection' -and $argsOut -contains $csv -and $argsOut -contains '-WhatIf') 'EXO selection uses the reviewed CSV and retains Dry Run'
        Assert-Selection ($argsOut -notcontains '-ForceEffectiveDate') 'A global date cannot erase distinct batch dates'
        function Test-MigrationSelectionSource {$false}
        $argsOut=Get-CommandArgument
        Assert-Selection ($argsOut -contains '-ForceEffectiveDate' -and $argsOut -contains '2030-05-01') 'Standard CSV mode retains its explicit global date option'
        Assert-Selection ($argsOut -notcontains '-MigrationBatchSelection' -and $argsOut -contains 'legacy.csv') 'Legacy CSV source does not acquire EXO selection semantics'
    }
}
finally {$window.Close()}

& {
    Add-Type -AssemblyName PresentationFramework
    Import-Module (Join-Path $app 'SmartM365.Communications.psm1') -Force -DisableNameChecking
    . (Join-Path $app 'SmartM365.ExchangeMigrationSelection.Gui.ps1')
    $view=New-SmartM365DemoView $markup
    $controls=Get-SmartM365DemoControls $view $markup
    foreach($name in $controls.Keys){Set-Variable -Scope Script -Name $name -Value $controls[$name]}
    $script:CurrentProcess=$null;$script:SelectedCampaignIndex=0
    $script:Campaigns=@([pscustomobject]@{Name='Fixture'})
    $fixtureContext=[pscustomobject]@{TenantId=$tenant;TenantKey='fixture'}
    function Get-MigrationSelectionContext {$fixtureContext}
    function Get-CampaignConfigValue {param($Campaign,$Name) if($Name -eq 'SentRegistryPath'){Join-Path $TestRoot 'ui-registry.csv'}else{Join-Path $TestRoot 'ui-output'}}
    function Resolve-GuiTokenizedPath {param($Value,$Campaign) $Value}
    function Show-CommandPreview {}
    function Add-OutputLine {param($Text)}
    function Show-Status {param($Text,$Color)}
    $state=$script:ExoMigration
    $state.TenantId=$tenant;$state.TenantKey='fixture';$state.Connected=$true;$state.Account='operator@example.invalid'
    $now=[datetime]::UtcNow.ToString('o')
    try {
        Receive-MigrationWorkerResponse ([pscustomobject]@{Kind='Batches';Data=[pscustomobject]@{Rows=@([pscustomobject]@{Identity='Batch-A';Status='Synced';Selected=$true},[pscustomobject]@{Identity='Old';Status='Completed';Selected=$false});ReadAtUtc=$now}})
        Assert-Selection ($script:ExoBatchesGrid.ItemsSource.Count -eq 1) 'Completed batches are hidden by default'
        $script:ExoShowCompletedCheck.IsChecked=$true;Show-MigrationBatchList
        Assert-Selection ($script:ExoBatchesGrid.ItemsSource.Count -eq 2) 'Completed batches can be inspected explicitly'
        $newRows=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture)) $zone)
        Receive-MigrationWorkerResponse ([pscustomobject]@{Kind='Recipients';Data=[pscustomobject]@{Rows=$newRows;ReadAtUtc=$now}})
        $script:MigrationRecipientSourceCombo.SelectedIndex=1
        Show-MigrationSelectionSource
        Assert-Selection (-not $script:RecipientPathBox.IsEnabled -and -not $script:EffectiveDatePicker.IsEnabled) 'EXO mode prevents legacy path and global date overrides'
        Save-MigrationRecipientSelection
        $prepared=$script:MigrationSelectionPath
        Assert-Selection ((Test-Path $prepared) -and @(Import-Csv $prepared).Count -eq 1) 'GUI prepares only reviewed recipients'
        $script:ExoOverrideDate.SelectedDate=[datetime]'2030-03-05';Set-MigrationRecipientDate
        Assert-Selection ($state.Rows[0].EffectiveDate -eq '2030-03-05' -and $state.Rows[0].DateSource -eq 'Operator override' -and -not $script:MigrationSelectionPath) 'Explicit date updates selected rows and invalidates prior CSV'
        $state.SelectionReadAtUtc=[datetime]::UtcNow.AddMinutes(-16).ToString('o')
        Assert-SelectionThrows {Save-MigrationRecipientSelection} 'expired'
        $state.SelectionReadAtUtc=$now
        $fixtureContext.TenantKey='other'
        Assert-SelectionThrows {Save-MigrationRecipientSelection} 'Tenant changed'
        $fixtureContext.TenantKey='fixture'
        $script:CurrentProcess=[pscustomobject]@{Fixture=$true};Show-MigrationWorkerState
        Assert-Selection (-not $script:ExoConnectButton.IsEnabled -and -not $script:ExoUseSelectionButton.IsEnabled) 'A running campaign disables selection changes'
        $script:CurrentProcess=$null
        Receive-MigrationWorkerResponse ([pscustomobject]@{Kind='Error';Data=[pscustomobject]@{Message='Synthetic partial read failure'}})
        Assert-Selection ($state.Rows.Count -eq 0 -and -not $script:MigrationSelectionPath) 'Read errors cannot reuse a previous audience'
        Assert-Selection (-not (Prepare-MigrationSelectionRun)) 'A rejected preparation returns only false and blocks campaign launch'
        Stop-MigrationSelectionWorker
        Assert-Selection (-not $state.Connected -and $state.Batches.Count -eq 0) 'Disconnect clears source provenance and selections'
    }
    finally {$view.Close()}
}

# Execute the real migration entry point in a self-contained fictional tree.
# No operational tenant configuration is copied or loaded.
$fixture=Join-Path $TestRoot 'standalone'
$campaign=Join-Path $fixture 'ExchangeUserNotifications/ExchangeMigration'
foreach($folder in @($campaign,(Join-Path $fixture 'Config'),(Join-Path $fixture 'Modules/SmartM365.Core'),(Join-Path $fixture 'ExchangeUserNotifications/Config/Campaigns'))) {[void][IO.Directory]::CreateDirectory($folder)}
foreach($file in @('SmartM365.Communications.psm1','SmartM365.ExchangeMigrationSelection.psm1')) {Copy-Item (Join-Path $app $file) (Join-Path $fixture ('ExchangeUserNotifications/'+$file))}
Copy-Item (Join-Path $app 'ExchangeMigration/SmartM365-ExchangeMigration-NotifyUsers.ps1') $campaign
Copy-Item (Join-Path $app 'ExchangeMigration/Templates') (Join-Path $campaign 'Templates') -Recurse
$core=Join-Path $fixture 'Modules/SmartM365.Core'
@'
function Set-SmartM365CoreContext { }
function Send-SmartM365TeamsNotification { throw 'Real summary notification was reached by an offline test' }
Export-ModuleMember -Function Set-SmartM365CoreContext,Send-SmartM365TeamsNotification
'@ | Set-Content (Join-Path $core 'Fixture.psm1') -Encoding utf8
"@{RootModule='Fixture.psm1';ModuleVersion='1.0.24';GUID='33333333-3333-4333-8333-333333333333';FunctionsToExport=@('Set-SmartM365CoreContext','Send-SmartM365TeamsNotification')}" | Set-Content (Join-Path $core 'SmartM365.Core.psd1') -Encoding utf8
@'
function Initialize-SmartM365TenantContext {
    $root=Split-Path $PSScriptRoot -Parent
    return @{TenantId='11111111-1111-4111-8111-111111111111';SmartM365RootPath=$root;WorkspaceRootPath=$root;DataAllRootPath=(Join-Path $root 'data');LatestCsvFolderPath=(Join-Path $root 'latest');LogAllRootPath=(Join-Path $root 'logs');From='helpdesk@example.invalid'}
}
'@ | Set-Content (Join-Path $fixture 'Config/SmartM365-TenantContext.ps1') -Encoding utf8
$config=@{CampaignName='Fixture';RunOutputPath=(Join-Path $fixture 'data');LogOutputPath=(Join-Path $fixture 'logs');SentRegistryPath=(Join-Path $fixture 'data/registry.csv');MailSendMode='Disabled';EnableAdLookupForLanguage=$false;ExchangeManagementMode='Auto';RequireExchangeManagement=$true;DefaultLanguageTag='en';From='helpdesk@example.invalid';PreventResendAcrossRuns=$true;DefaultEffectiveDate='';TermsPortalUrl='https://terms.example.invalid';OldWebmailUrl='https://old.example.invalid';NewWebmailUrl='https://new.example.invalid'}
$json=$config|ConvertTo-Json -Depth 5
foreach($relative in @('ExchangeUserNotifications/Config/ExchangeUserNotifications.local.json','ExchangeUserNotifications/Config/ExchangeUserNotifications.local.json.template','ExchangeUserNotifications/Config/Campaigns/ExchangeMigration.local.json','ExchangeUserNotifications/Config/Campaigns/ExchangeMigration.local.json.template')) {$json|Set-Content (Join-Path $fixture $relative) -Encoding utf8}
$selection=Join-Path $fixture 'selected.csv'
$freshRows=@(ConvertTo-ExoSelectionRecipient @((New-SelectionFixture)) $zone)
Export-ExoMigrationSelection $freshRows $tenant 'fixture' $selection ([datetime]::UtcNow.ToString('o')) | Out-Null
$invokePath=Join-Path $campaign 'SmartM365-ExchangeMigration-NotifyUsers.ps1'
& (Get-Command pwsh).Source -NoProfile -ExecutionPolicy Bypass -File $invokePath -Tenant fixture -RecipientsPath $selection -MigrationBatchSelection -WhatIf *> (Join-Path $TestRoot 'campaign-simulation.log')
Assert-Selection ($LASTEXITCODE -eq 0) 'Real campaign accepts reviewed recipients without EXO authentication in selection mode'
$output=[IO.File]::ReadAllText((Join-Path $TestRoot 'campaign-simulation.log'))
Assert-Selection ($output -match 'DryRun=1' -and $output -match 'Sent=0' -and $output -notmatch 'Failed=1') 'Real campaign previews one selected message with no delivery'
$badRows=@(Import-Csv $selection);$badRows[0].EffectiveDate='';$badRows|Export-Csv $selection -NoTypeInformation -Encoding utf8
& (Get-Command pwsh).Source -NoProfile -ExecutionPolicy Bypass -File $invokePath -Tenant fixture -RecipientsPath $selection -MigrationBatchSelection -WhatIf *> (Join-Path $TestRoot 'campaign-rejected.log')
Assert-Selection ($LASTEXITCODE -ne 0) 'Real campaign rejects missing dates before any source connection or message processing'
[pscustomobject]@{Status='PASS';Checks=$script:checks;Mode='Offline only';TestRoot=$TestRoot}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJcE2uDemDxUUp
# +OsMDTNx7faRcY3spXMu19Nh22jLj6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILRJiNdXs7u+pX39K7/K5XK/aXPZexaF8Kgzap9OnSaaMA0GCSqG
# SIb3DQEBAQUABIIBgFUGRhh7IDFEKDdohpZgzuLzdMBFMz/JOoHRBeEA4kVpVEoF
# 6nVYMLJla/9zMFQaZzf4IzFwx+TPM9elCDSPysm6PRHzUIu2OFcaTCWwL2KzjT3m
# I3HXIqYdMrQPGgKlftkyzeqZsLysYEy3E+m/6/VMHGo4i9Y7H4BYCEtcROKuf63J
# 83V14X3zZ+p5dVhSL9l7w1PP3naf7KfPDxDTKqOU6+QA1C5zIHOmCxqRhAoI/8ZK
# Y/XsrZMd3LxbAkcu60fD+HKpbvJU+35vi2839HHqHTLY9Ums/X5Udc1kKF2sGRTU
# UCqKf+hQW053M/6dEJ0sbRXH8bO0YnmezjKWEVeo7q6S8nuUdaLa+i1M9uNOVsfF
# h8Ed/XGYwYuUy4j7771fUlIKASd8A6kze4CVqGn6++e2W/IYttYQlphacOYtDc7i
# 93U/bJVABM6Labq8VaemlDjg7wQ1QkvlUwK6/a2nNxoUlxVztd2kfAmSLG3e0764
# eHwvsHsiJ8Ow7RcWjKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzUz
# NTVaMC8GCSqGSIb3DQEJBDEiBCBDX2Wyu1Bmu1+X0pUORxwGAcetD1Zjk8yelLGV
# EzzHBjANBgkqhkiG9w0BAQEFAASCAgAVFRjLKa8j4aAL3FL4L0zmZ8QNyUjBogVz
# ebcpxY2g6ar+t6L93NAnqnlZYi0tQPhcXtra735Z+ZPjixIewJ5q+lpOeC3uthi8
# 6t5tKfhF8xjqpAGJPGyC2sWZ62+Xwd1XDWZmipaNspNwkwifLtVl/TIpL8PTfqwN
# mVVsPRHK7Saq5JK8LHmoNI4nTQG40a0fhsjyNHt3XzGHgSZMpaa4Vzp6ScttYKxn
# UUkajTgl4j3dIMMmRLf9pddl0I6CtKVTrym3yrNzcXpUZ/dVqReD8mQYTzfu41KH
# 8iI5VEWTUHLCnaqQyjR4fBcVxP8+cYI4beOrFGADLLhfn7h/m6495px6tmbYrRiW
# 8nqA2j64OHpYwOSWnf+Q92dXytwa12dA7bWHZ4ZJARSkaw1ZHz5JZ5FtsTiue89B
# af/IG5624xIQ5riklC3ArckVSPjtqBch+bVrhYbjbluu91sRw34XkN+OgKke5Vh1
# Um1B3YhiatekFJ8i0WMbfOpW2kRjSO9K7Pk2Ycz46pgG9w/ROGLzksLFdf3Fq0UR
# TiK8CxySd+TBBymEBvUiG9vkccGC4SsmuDoFVPSEho2SGtEmv3NKM/ToO50EVnMW
# xc6Jjo4XHwwCHD6DlCHuM4mdYhtKlLfU4868jRR1bWg1LBHAbCy8uYKrCofzBDLU
# VE9Co04ziw==
# SIG # End signature block
