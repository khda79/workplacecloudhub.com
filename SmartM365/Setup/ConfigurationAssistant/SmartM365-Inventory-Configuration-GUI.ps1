<# SmartInventory configuration assistant, version 0.99.0. Offline tested; test qualification pending. #>
[CmdletBinding()]
param([string]$Tenant = '', [switch]$ValidateOnly, [switch]$OfflineSmokeTest, [switch]$OfflineEmptyProfileSmokeTest, [switch]$OfflineMultipleProfilesSmokeTest)
$ErrorActionPreference = 'Stop'
if ($OfflineEmptyProfileSmokeTest -and $OfflineMultipleProfilesSmokeTest) {throw 'Select only one offline profile fixture.'}
if ($OfflineEmptyProfileSmokeTest -or $OfflineMultipleProfilesSmokeTest) {$OfflineSmokeTest=$true}
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationAudit.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ProfileDiscovery.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationSteps.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.Workflow.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.SimpleWorkflow.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationEdit.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.PendingEdits.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.PendingActions.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.SigningCertificate.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.TenantIdentityEdit.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationChecks.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationServiceAccount.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ServiceReadiness.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ServiceShareAccess.Network.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.DirectoryServiceAccount.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ServiceAccountRights.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ConfigurationCatalog.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.OrchestratorPlan.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.OrchestratorSettings.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.OrchestratorConsole.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.OrchestratorRegistration.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.TaskInstallerOffer.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.PermissionPlan.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.PermissionRunner.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.BootstrapPlan.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.BootstrapReconcile.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.AutomaticReconciliation.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.CertificateExport.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.ApplyProgress.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.BootstrapStage.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.AutomaticBootstrap.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.BootstrapProposal.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.BootstrapRun.Interactive.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.TenantSignInUi.InitialDomain.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.PrerequisiteInstall.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.MailConfigurationCheck.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.WorkflowAssessment.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.AcceptanceEvidence.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.LocalPrerequisiteScan.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.MailTest.ps1')
. (Join-Path $PSScriptRoot 'Lib/SmartM365.AssistantLog.ps1')
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Set-SiAssistantRepositoryRoot -Root $root
$runtimeRoot = $root
$assistantRoot = Join-Path $PSScriptRoot 'Views'
$assistantLibRoot = Join-Path $PSScriptRoot 'Lib'
$xamlPath = Join-Path $PSScriptRoot 'Views/SmartM365-Inventory-Configuration-GUI.xaml'
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
if ($ValidateOnly -and -not $OfflineSmokeTest) {
    $null = Get-SiProfileDiscovery -Root $root
    if ($Tenant) {$null = Get-SiStepAudit -SmartM365Root $root -Tenant $Tenant}
    Write-Output 'Assistant preview: XML and passive profile discovery validated.'
    return
}
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$splash=$null; $smokeRoot=$null;$assistantLog=$null;$assistantLogUnavailable=$false
if (-not $OfflineSmokeTest) {
    try {$assistantLog=New-SiAssistantLog -Directory (Join-Path $PSScriptRoot 'Log');Write-SiAssistantLog -Log $assistantLog -Event Start} catch {$assistantLogUnavailable=$true}
}
try {
    . (Join-Path $runtimeRoot 'Config/SmartM365-TenantContext.ps1')
    if ($OfflineSmokeTest) {
        $smokeRoot=Join-Path ([IO.Path]::GetTempPath()) ('smartinventory-ui-smoke-'+[guid]::NewGuid().ToString('N'))
        $null=New-Item -ItemType Directory -Path (Join-Path $smokeRoot 'Config/Tenants') -Force
        $null=New-Item -ItemType Directory -Path (Join-Path $smokeRoot 'Setup') -Force
        [IO.File]::WriteAllText((Join-Path $smokeRoot 'Setup/Install-SmartM365-SmartInventoryPrerequisites.ps1'),'throw "synthetic installer must not run"')
        foreach ($synthetic in @('Setup/SmartM365-Create-AppRegistration.ps1','Config/SmartM365-TenantContext.ps1','Config/SmartM365-JsonTransport.policy.psd1','Config/Tenants/tenant.mail.local.json.template','Modules/SmartM365.Core/SmartM365.JsonTransport.psd1','Modules/SmartM365.Core/SmartM365.JsonTransport.psm1')) {
            $file=Join-Path $smokeRoot $synthetic;$null=New-Item -ItemType Directory -Path (Split-Path -Parent $file) -Force;[IO.File]::WriteAllText($file,'synthetic offline fixture; never execute')
        }
        $assistantLog=New-SiAssistantLog -Directory (Join-Path $smokeRoot 'Log');Write-SiAssistantLog -Log $assistantLog -Event Start
        if (-not $OfflineEmptyProfileSmokeTest) {[IO.File]::WriteAllText((Join-Path $smokeRoot 'Config/SmartM365.global.local.json'),$(if ($OfflineMultipleProfilesSmokeTest) {'{"DefaultTenant":"beta","WorkspaceRootPath":"\\\\synthetic.invalid\\old","Preserved":"keep"}'} else {'{}'}))}
        [IO.File]::WriteAllText((Join-Path $smokeRoot 'Config/SmartM365.global.local.json.template'),'{"DefaultTenant":"test","WorkspaceRootPath":"synthetic"}')
        if (-not $OfflineEmptyProfileSmokeTest) {[IO.File]::WriteAllText((Join-Path $smokeRoot 'Config/Tenants/test.local.json'),'{"ProfileKey":"test","OrganizationKey":"synthetic","EnvironmentKey":"test","TenantKey":"synthetic-test","EnableSharePointUpload":true,"EnableTeamsNotifications":false,"TeamsWebhookUrl":"https://workflow.invalid/legacy","TeamsAlertsWebhookUrl":"https://workflow.invalid/alerts"}')}
        if ($OfflineMultipleProfilesSmokeTest) {[IO.File]::WriteAllText((Join-Path $smokeRoot 'Config/Tenants/beta.local.json.txt'),'{"ProfileKey":"beta"}')}
        [IO.File]::WriteAllText((Join-Path $smokeRoot 'Config/Tenants/tenant.local.json.template'),'{"ProfileKey":"tenant","OrganizationKey":"example","EnvironmentKey":"test","TenantKey":"example-test","ServiceAccount":"","Environment":"Test","OrgDomain":"example.invalid","Smart365Administrators":"","ErrorMailTo":"admin@example.com","To":"","SharePointSiteHostname":"contoso.sharepoint.com","TeamsWebhookUrl":"","TeamsAlertsWebhookUrl":"","TeamsInfosWebhookUrl":"","Thumb":"","Thumbprint":"","EnableTeamsNotifications":false,"EnableSharePointUpload":false,"EnableSharePointUploadDataAll":"__USE_GLOBAL__","EnableSharePointUploadDataLast":"__USE_GLOBAL__","EnableSharePointUploadDataPowerBI":"__USE_GLOBAL__","EnableSharePointUploadLogAll":"__USE_GLOBAL__"}')
        $root=$smokeRoot; $Tenant=if ($OfflineMultipleProfilesSmokeTest) {''} else {'test'}
    } else {
        . (Join-Path $PSScriptRoot 'Lib/SmartM365.GuiSplash.ps1')
        $splash=Start-SmartM365GuiSplash -Framework Wpf -ProductName 'SmartInventory Configuration' -LogoPath (Join-Path $PSScriptRoot 'Assets/WorkplaceCloudHub-lockup-WPF.png') -WindowIconPath (Join-Path $PSScriptRoot 'Assets/WorkplaceCloudHub.ico')
    }
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $tenantBox = $window.FindName('Tenant'); $profileChoices=$window.FindName('ProfileChoices')
    $discovery=Get-SiProfileDiscovery -Root $root
    $window.FindName('ProfileDiscoveryStatus').Text=$discovery.Message
    foreach ($profile in @($discovery.Profiles | Where-Object State -eq 'Present')) {[void]$profileChoices.Items.Add($profile.Name)}
    if ($discovery.State -in @('Unavailable','CorrectionRequired')) {$tenantBox.IsEnabled=$false}
    else {
        $selected=[string]$discovery.Selected
        if ($Tenant -and @($discovery.Profiles | Where-Object {$_.State -eq 'Present' -and $_.Name -ieq $Tenant}).Count -eq 1) {$selected=[string]@($discovery.Profiles | Where-Object {$_.Name -ieq $Tenant})[0].Name}
        if ($selected) {$profileChoices.SelectedItem=$selected;$tenantBox.Text=$selected}
    }
    $steps = $window.FindName('Steps'); $status = $window.FindName('Status')
    if ($assistantLogUnavailable) {$window.Title+=' [log unavailable]'}
    $phases=$window.FindName('Phases');$phaseOptions=$window.FindName('PhaseOptions');$phaseState=@{Changing=$false}
    $phases.ItemsSource=@(Get-SiSimpleWorkflow);$phases.DisplayMemberPath='Title'
    $field = $window.FindName('Field'); $valueBox = $window.FindName('Value')
    $window.FindName('ReviewDecision').Visibility='Collapsed';$window.FindName('MarkReviewed').Visibility='Collapsed'
    $identityButton=[Windows.Controls.Button]::new();$identityButton.Content='Review tenant identity';$identityButton.HorizontalAlignment='Left';$identityButton.Visibility='Collapsed'
    $null=$window.FindName('Next').Parent.Children.Insert(0,$identityButton)
    $createGlobalButton=[Windows.Controls.Button]::new();$createGlobalButton.Content='Create global configuration';$createGlobalButton.Visibility='Collapsed'
    $null=$window.FindName('Next').Parent.Children.Insert(0,$createGlobalButton)
    $serviceShareButton=[Windows.Controls.Button]::new();$serviceShareButton.Content='Verify service access';$serviceShareButton.Visibility='Collapsed'
    $null=$window.FindName('Next').Parent.Children.Insert(0,$serviceShareButton)
    $storageButton=[Windows.Controls.Button]::new();$storageButton.Content='Change CSV and log share';$storageButton.Visibility='Collapsed'
    $null=$window.FindName('Next').Parent.Children.Insert(0,$storageButton)
    $applyPendingButton=[Windows.Controls.Button]::new();$applyPendingButton.Content='Apply reviewed settings';$applyPendingButton.FontWeight=[Windows.FontWeights]::Bold;$applyPendingButton.Visibility='Collapsed'
    $discardPendingButton=[Windows.Controls.Button]::new();$discardPendingButton.Content='Discard pending settings';$discardPendingButton.Visibility='Collapsed'
    $null=$window.FindName('Next').Parent.Children.Insert(0,$applyPendingButton);$null=$window.FindName('Next').Parent.Children.Insert(0,$discardPendingButton)
    $pendingEdits=[Collections.Generic.List[object]]::new();$pendingActions=[Collections.Generic.List[object]]::new();$window.FindName('SaveField').Content='Add to review'
    $window.FindName('PlanTask').Content='Add task to review'
    $smokeUiState=@{PassiveRuns=0;SaveError='';ProgressSeen=$false;ProgressAfterOne=$false;ShareProbeDenied=$false}
    $editState = @{ Snapshot = $null }
    $observations = @{}
    $serviceState=@{Identity=$null;Created=$false}
    $serviceShareState=@{Ready=[bool]$OfflineSmokeTest;Sid='';Path='';State='Not checked'}
    $exchangeRoleState=@{Ready=[bool]$OfflineSmokeTest;Sid=''}
    $shareState=@{Path='';Ready=[bool]$OfflineSmokeTest;Checking=$false;Process=$null;Timer=$null;Result='';Started=$null}
    $reviewed = @{}
    $taskPlanState=@{Plan=$null}
    $taskInstallerState=@{Run=$null;Timer=$null;Action=$null}
    $taskOfferState=@{HadChanges=$false;Offers=0}
    $machineState=@{Ready=$false;Rows=@();Known=$false;SimulatedInstalls=0;LastInstall='';LastInstallDetail='';ValidationFailed=$false}
    $installState=@{Run=$null;Timer=$null;Action=$null;Confirmed=$false}
    $mailState=@{Run=$null;Timer=$null}
    $bootstrapState=@{Run=$null;Timer=$null}
    $applyProgressState=@{Progress=$null;Active=$false;Starts=0;Finishes=0;LastResult=''}
    $setApplyProgress={param([string]$message) if ($applyProgressState.Progress) {Set-SiApplyProgress -Progress $applyProgressState.Progress -Message $message}}.GetNewClosure()
    $finishApplyProgress={param([string]$message,[string]$kind)
        $wasActive=$applyProgressState.Active
        if ($applyProgressState.Progress) {try {Stop-SiApplyProgress -Progress $applyProgressState.Progress} catch {$message+=' The progress window could not be confirmed closed.';if ($kind -eq 'Success') {$kind='Warning'}};$applyProgressState.Progress=$null}
        $applyProgressState.Active=$false
        if ($wasActive) {$applyProgressState.Finishes++;$applyProgressState.LastResult=$kind;if (-not $OfflineSmokeTest) {[void][Windows.MessageBox]::Show($window,$message,'Configuration assistant','OK',$(if ($kind -in @('Success','Next')) {'Information'} else {'Warning'}))}}
    }.GetNewClosure()
    $certificateExportState=@{Thumb='';Profile='';Exported=$false}
    $certificateExportButton=[Windows.Controls.Button]::new();$certificateExportButton.Content='Export certificate';$certificateExportButton.Visibility='Collapsed';$null=$window.FindName('Next').Parent.Children.Insert(0,$certificateExportButton)
    $consoleState=@{Run=$null;Timer=$null}
    $permissionState=@{Run=$null;Timer=$null}
    $assessmentState=@{Generation=0;Active=$null;Async=$null;Timer=$null;Popup=$null;Pending=$false;Tenant='';StartedGeneration=0;Closed=$false}
    $scope=@{Set=$true;Cloud=$true;AD=$true;Exchange=$true;SharePoint=$false;Teams=$false;Mail=$true}
    $prereqList=[Windows.Controls.StackPanel]::new();$prereqList.Visibility='Collapsed';$null=$window.FindName('Action').Parent.Children.Insert(4,$prereqList)
    $profileSelectionState=@{Preparing=$false}
    $tenantSession=@{Tenant='';Account='';Mail='';SiteHost='';InitialDomain='';Profile='';PreflightReady=$false}
    $prereqUi=@{Owner=$null}
    $window.FindName('Logo').Source = [Windows.Media.Imaging.BitmapImage]::new([uri](Join-Path $PSScriptRoot 'Assets/WorkplaceCloudHub-lockup-WPF.png'))
    $updateDetail = {
        $row = $steps.SelectedItem
        if ($null -eq $row) { return }
        $window.FindName('Title').Text = $row.Title
        $window.FindName('StepState').Text = $row.State
        $window.FindName('Detail').Text = $row.Detail
        $window.FindName('Action').Text = $row.Action
        if ($observations.ContainsKey($row.Id)) {
            $window.FindName('StepState').Text = $observations[$row.Id].State
            $window.FindName('Action').Text = $observations[$row.Id].Detail
        }
        $prereqList.Visibility=if ($row.Id -eq 'prerequisites') {'Visible'} else {'Collapsed'}
        if ($row.Id -eq 'prerequisites') {
            $prereqList.Children.Clear()
            foreach ($item in $machineState.Rows) {
                $line=[Windows.Controls.TextBlock]::new();$line.Text=if ($item.Present) {'✓  '+$item.Name+'  —  '+$item.Detail} else {'✗  '+$item.Name+'  —  '+$item.Detail}
                $line.Foreground=if ($item.Present) {[Windows.Media.Brushes]::DarkGreen} else {[Windows.Media.Brushes]::DarkRed};$line.Margin='0,3,0,3';$line.TextWrapping='Wrap';[void]$prereqList.Children.Add($line)
            }
            $window.FindName('StepState').Text=if (-not $machineState.Known) {'Checking machine prerequisites'} elseif ($machineState.Ready) {'All required prerequisites detected'} else {'Prerequisites missing or need attention'}
            $window.FindName('Detail').Text='Required: PowerShell modules available to all users and the signing certificate trusted on this machine.'
            $window.FindName('Action').Text=if ($machineState.Ready) {'Continue to the next step.'} elseif ($machineState.ValidationFailed) {'Installation validation failed. Open installation details for the error.'} else {'Install prerequisites, then wait for the updated status. A Windows restart may be required for ActiveDirectory RSAT.'}
        }
        $serviceReady=Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions
        $serviceShareReady=($serviceReady.Ready -and $serviceShareState.Ready -and ($OfflineSmokeTest -or ($serviceShareState.Sid -ceq [string]$serviceState.Identity.Sid -and $serviceShareState.Path -ieq $shareState.Path)))
        $exchangeRoleReady=(-not $scope.Exchange -or ($exchangeRoleState.Ready -and ($OfflineSmokeTest -or $exchangeRoleState.Sid -ceq [string]$serviceState.Identity.Sid)))
        if ($row.Id -eq 'service') {
            $window.FindName('Detail').Text='Use an existing account or create a domain service account.'
            switch ($serviceReady.Kind) {
                'Existing' {$window.FindName('StepState').Text=if ($serviceState.Created) {'Service account created'} else {'Existing account selected'};$window.FindName('Action').Text='Service account: '+$serviceReady.Label}
                'Pending' {$window.FindName('StepState').Text='Service account required';$window.FindName('Action').Text='The pending account has not been created. Discard it, then create or select an account here.'}
                'Uncertain' {$window.FindName('StepState').Text='Account creation unconfirmed';$window.FindName('Action').Text='Check the directory before retrying.'}
                default {$window.FindName('StepState').Text='Service account required';$window.FindName('Action').Text='Create a domain account or select an existing account before continuing.'}
            }
        }
        if ($row.Id -eq 'service' -and $serviceReady.Ready) {
            if ($serviceShareReady -and $exchangeRoleReady) {$window.FindName('StepState').Text='Service account and share ready';$window.FindName('Action').Text='Service account: '+$serviceReady.Label+'; share read and write verified.'}
            elseif ($serviceShareReady) {$window.FindName('StepState').Text='Exchange role group membership required';$window.FindName('Action').Text='Verify the service account is in View-Only Organization Management.'}
            else {$window.FindName('StepState').Text='Service account share access required';$window.FindName('Action').Text='Verify this account can read and write the CSV and log share before continuing.'}
        }
        if ($row.Id -eq 'identity') {
            $bootstrapPending=@($pendingActions | Where-Object Kind -eq 'Bootstrap')
            if ($bootstrapPending.Count -eq 1) {
                $bootstrapPlan=$bootstrapPending[0].Data.Plan.BootstrapPlan
                $window.FindName('StepState').Text='Bootstrap included in final review'
                $window.FindName('Detail').Text="Tenant ID: $($bootstrapPlan.Directory)`nApplication: $($bootstrapPlan.DisplayName)`nTeams workspace and SharePoint: $($bootstrapPlan.ConfigureTeams)"
                $window.FindName('Action').Text='Continue to review. The bootstrap will run only after final confirmation.'
            }
        }
        if ($row.Id -eq 'sharepoint') {
            try {
                $site=Get-SiSimpleSharePointDefaults -Root $root -Tenant $tenantBox.Text
                $siteHost=if ($site.ProfileState -eq 'Missing') {if ($tenantSession.SiteHost) {$tenantSession.SiteHost} else {'(not detected yet)'}} else {$site.Values.SharePointSiteHostname}
                $window.FindName('Detail').Text="SharePoint upload: $($site.Enabled)`nSite host: $siteHost`nSite path: $($site.Values.SharePointSitePath)`nLibrary: $($site.Values.SharePointLibraryDisplayName)`nFolder: $($site.Values.SharePointTargetFolderPath)`n`nExisting profile values take priority; missing values are shown from the template. Review placeholders before applying."
            } catch {$window.FindName('Detail').Text='SharePoint settings could not be displayed. Review the profile and tenant template.'}
        }
        if ($row.Id -eq 'share') {
            try {
                $storage=Get-SiSimpleStoragePreview -Root $root
                $workspace=[string]$storage.Values.WorkspaceRootPath;$workspaceNote=''
                $globalPending=@($pendingActions | Where-Object Kind -eq 'Global' | Select-Object -First 1)
                if ($globalPending.Count) {$pendingWorkspace=@($globalPending[0].Data.PendingEdits | Where-Object Field -eq 'WorkspaceRootPath' | Select-Object -Last 1);if ($pendingWorkspace.Count) {$workspace=[string]$pendingWorkspace[0].Value;$workspaceNote=' (pending review)'}}
                $pendingWorkspaceEdit=@($pendingEdits | Where-Object {$_.Snapshot.Step -eq 'share' -and $_.Snapshot.Field -eq 'WorkspaceRootPath'} | Select-Object -Last 1)
                if ($pendingWorkspaceEdit.Count) {$workspace=[string]$pendingWorkspaceEdit[0].Value;$workspaceNote=' (pending review)'}
                $window.FindName('Detail').Text="Product location: $root`nCSV and log share: $workspace$workspaceNote`n`nThe share is global. Existing individual paths can override it."
                $window.FindName('Action').Text='Change the share if needed, then continue. Changes are applied from the final review.'
                if (-not $OfflineSmokeTest -and $shareState.Path -ine $workspace) {$shareState.Ready=$false}
                $window.FindName('StepState').Text=if ($shareState.Checking) {'Checking share availability'} elseif ($shareState.Ready) {'Share available'} else {'Share unavailable or not checked'}
                $window.FindName('Action').Text=if ($shareState.Checking) {'Checking access under the current Windows account...'} elseif ($shareState.Ready) {'Share is accessible under the current Windows account. Continue to the next step.'} else {'The share must be accessible before continuing. Change it or use Refresh to retry.'}
            } catch {$window.FindName('Detail').Text='Storage settings could not be displayed. Review or create the global configuration.'}
        }
        $sectionPending=@($pendingEdits | Where-Object {$_.Snapshot.Step -eq $row.Id})
        if ($sectionPending.Count -and $row.Id -ne 'final') {$window.FindName('Detail').Text+="`n`nPending changes in this section: "+(($sectionPending | ForEach-Object {$_.Snapshot.Field}) -join ', ')+'. Review them on the final page.'}
        if ($reviewed.ContainsKey($row.Id)) { $window.FindName('StepState').Text += '  |  Decision: '+$reviewed[$row.Id] }
        if ($row.Id -eq 'final') {
            $window.FindName('Action').Text='Review the planned changes, then select Apply reviewed settings.'
            try {
                $sections=[Collections.Generic.List[string]]::new()
                if ($null -ne $serviceState.Identity) {$sections.Add('Service account: '+[string]$serviceState.Identity.Account)}
                $sections.Add('Pending actions:');$sections.Add((Get-SiPendingCreationReview -Queue $pendingActions))
                if ($pendingEdits.Count) {$sections.Add('');$sections.Add('Pending configuration changes:');$sections.Add((Get-SiConcisePendingEditReview -Queue $pendingEdits))}
                $window.FindName('Detail').Text=$sections -join [Environment]::NewLine
            } catch {$window.FindName('Detail').Text='Review unavailable; refresh the selected profile.'}
        }
        $window.FindName('CreateProfile').Visibility = 'Collapsed'
        $createGlobalButton.Visibility = 'Collapsed'
        $storageButton.Visibility = if ($row.Id -eq 'share') { 'Visible' } else { 'Collapsed' }
        $serviceShareButton.Visibility=if ($row.Id -eq 'service' -and $serviceReady.Ready -and (-not $serviceShareReady -or -not $exchangeRoleReady)) {'Visible'} else {'Collapsed'}
        $applyPendingButton.Visibility=if ($row.Id -eq 'final') {'Visible'} else {'Collapsed'};$applyPendingButton.IsEnabled=($pendingEdits.Count -gt 0 -or $pendingActions.Count -gt 0)
        $discardPendingButton.Visibility=if ($row.Id -eq 'final') {'Visible'} else {'Collapsed'};$discardPendingButton.IsEnabled=($pendingEdits.Count -gt 0 -or $pendingActions.Count -gt 0)
        $window.FindName('AuditPrerequisites').Visibility = 'Collapsed'
        $window.FindName('InstallPrerequisites').Visibility = if ($row.Id -eq 'prerequisites') { 'Visible' } else { 'Collapsed' }
        $window.FindName('InstallSigningCertificate').Visibility = 'Collapsed'
        $window.FindName('StopPrerequisites').Visibility = 'Collapsed'
        $window.FindName('StopPrerequisites').IsEnabled = ($null -ne $installState.Run)
        $window.FindName('ConfigureService').Visibility = if ($row.Id -eq 'service') { 'Visible' } else { 'Collapsed' }
        $window.FindName('CreateService').Visibility = 'Collapsed'
        $window.FindName('CreateDirectoryService').Visibility = 'Collapsed'
        $window.FindName('CheckShare').Visibility = 'Collapsed'
        $window.FindName('CheckShareAsService').Visibility = 'Collapsed'
        $window.FindName('CheckShareAsService').IsEnabled = ($observations.ContainsKey('service') -and [bool]$observations['service'].Sid -and -not [bool]$observations['service'].Managed)
        $window.FindName('CheckShareWrite').Visibility = 'Collapsed'
        $window.FindName('CheckTask').Visibility = if ($row.Id -eq 'orchestrator') { 'Visible' } else { 'Collapsed' }
        $window.FindName('CheckCertificate').Visibility = 'Collapsed'
        $window.FindName('SetCertificate').Visibility = if ($row.Id -eq 'identity' -and -not @($pendingActions | Where-Object Kind -eq 'Profile').Count) { 'Visible' } else { 'Collapsed' }
        $window.FindName('CheckMail').Visibility = if ($row.Id -eq 'mail') { 'Visible' } else { 'Collapsed' }
        $window.FindName('TestMail').Visibility = 'Collapsed'
        $window.FindName('StopMail').Visibility = 'Collapsed'
        $window.FindName('StopMail').IsEnabled = ($null -ne $mailState.Run)
        $window.FindName('CheckCatalog').Visibility = if ($row.Id -eq 'collectors') { 'Visible' } else { 'Collapsed' }
        $window.FindName('SetEnabled').Visibility = 'Collapsed'
        $window.FindName('SetUploadCategory').Visibility = 'Collapsed'
        $window.FindName('SetScope').Visibility = 'Collapsed'
        $window.FindName('CheckService').Visibility = 'Collapsed'
        $window.FindName('CheckServiceRights').Visibility = 'Collapsed'
        $window.FindName('CheckServiceRights').IsEnabled = ($observations.ContainsKey('service') -and [bool]$observations['service'].Sid)
        $window.FindName('PlanTask').Visibility = if ($row.Id -eq 'orchestrator') { 'Visible' } else { 'Collapsed' }
        $window.FindName('EditOrchestratorSettings').Visibility = if ($row.Id -eq 'orchestrator') { 'Visible' } else { 'Collapsed' }
        $window.FindName('OpenOrchestratorConsole').Visibility = 'Collapsed'
        $window.FindName('PlanTask').IsEnabled = (($observations.ContainsKey('service') -and [bool]$observations['service'].Sid) -or @($pendingActions | Where-Object {$_.Kind -in @('LocalAccount','DomainAccount') -and $_.State -eq 'Pending'}).Count -eq 1)
        $window.FindName('RegisterTask').Visibility = 'Collapsed'
        $window.FindName('RegisterTask').IsEnabled = $false
        $window.FindName('PlanBootstrap').Visibility = if ($row.Id -eq 'identity' -and -not @($pendingActions | Where-Object Kind -eq 'Bootstrap').Count) { 'Visible' } else { 'Collapsed' }
        $window.FindName('CheckPermissions').Visibility = 'Collapsed'
        $window.FindName('StopPermissions').Visibility = 'Collapsed'
        $window.FindName('StopPermissions').IsEnabled = ($null -ne $permissionState.Run)
        $window.FindName('ReconcileBootstrap').Visibility = if ($row.Id -eq 'identity' -and $bootstrapState.ProposalAwaiting) { 'Visible' } else { 'Collapsed' }
        if ($certificateExportState.Profile -and $certificateExportState.Profile -cne $tenantBox.Text) {$certificateExportState.Thumb='';$certificateExportState.Profile='';$certificateExportState.Exported=$false}
        if ($row.Id -eq 'identity' -and -not $certificateExportState.Thumb) {
            try {
                $storedProfile=Read-SiPrivateJson (Join-Path $root ('Config/Tenants/'+$tenantBox.Text+'.local'))
                if ($storedProfile.State -eq 'Present') {
                    $storedThumb=[string]$storedProfile.Value.Thumbprint
                    if ($storedThumb -match '^[0-9a-fA-F]{40}$' -and $storedThumb -notmatch '^0+$' -and $storedThumb -ieq [string]$storedProfile.Value.Thumb -and (Test-Path -LiteralPath ('Cert:\CurrentUser\My\'+$storedThumb))) {$certificateExportState.Thumb=$storedThumb;$certificateExportState.Profile=$tenantBox.Text}
                }
            } catch { }
        }
        $certificateExportButton.Visibility=if ($row.Id -eq 'identity' -and $certificateExportState.Thumb -and $certificateExportState.Profile -ceq $tenantBox.Text -and -not $certificateExportState.Exported) {'Visible'} else {'Collapsed'}
        $editState.Snapshot = $null
        $valueBox.Text = ''
        $fields = @(Get-SiEditableFields $row.Id)
        $field.ItemsSource = $fields
        $window.FindName('Editor').Visibility = 'Collapsed'
        $identityButton.Visibility=if ($row.Id -eq 'tenant' -and -not @($pendingActions | Where-Object Kind -eq 'Profile').Count) {'Visible'} else {'Collapsed'}
        if ($fields.Count) { $field.SelectedIndex = 0 }
        $window.FindName('Previous').IsEnabled = $steps.SelectedIndex -gt 0
        $window.FindName('Next').IsEnabled = ($steps.SelectedIndex -lt ($steps.Items.Count - 1) -and ($row.Id -ne 'share' -or $shareState.Ready) -and ($row.Id -ne 'prerequisites' -or $machineState.Ready) -and ($row.Id -ne 'service' -or ($serviceShareReady -and $exchangeRoleReady)))
        $window.FindName('MarkReviewed').IsEnabled = ($row.Id -ne 'final')
        $decisionIndex=if (-not $reviewed.ContainsKey($row.Id)) { 0 } elseif ($reviewed[$row.Id] -eq 'Needs action') { 1 } elseif ($reviewed[$row.Id] -eq 'Not applicable') { 2 } else { 0 }
        $window.FindName('ReviewDecision').SelectedIndex=$decisionIndex
        $window.FindName('ReviewDecision').IsEnabled=($row.Id -ne 'final')
        $profileState='Present'
        if ($row.Id -eq 'tenant') { $profileState=(Read-SiPrivateJson (Join-Path $root ("Config/Tenants/$($tenantBox.Text).local"))).State }
        $serviceResolved=($observations.ContainsKey('service') -and [bool]$observations['service'].Sid)
        $guide=Get-SiWorkflowNextAction -Step $row.Id -ProfileState $profileState -ServiceResolved $serviceResolved
        $window.FindName('GuidedAction').Visibility=if ($row.Id -in @('prerequisites','service','share','teams','tenant','identity','permissions','administrators','sharepoint','mail','final')) {'Collapsed'} else {'Visible'}
        $window.FindName('GuidedAction').Content=$guide.Label
        $window.FindName('GuidedAction').IsEnabled=([bool]$guide.Control)
    }.GetNewClosure()
    $startShareCheck={
        if ($OfflineSmokeTest) {return}
        try {
            $path=[string](Get-SiSimpleStoragePreview -Root $root).Values.WorkspaceRootPath
            $globalPending=@($pendingActions | Where-Object Kind -eq 'Global' | Select-Object -First 1)
            if ($globalPending.Count) {$v=@($globalPending[0].Data.PendingEdits | Where-Object Field -eq 'WorkspaceRootPath' | Select-Object -Last 1);if ($v.Count) {$path=[string]$v[0].Value}}
            $v=@($pendingEdits | Where-Object {$_.Snapshot.Step -eq 'share' -and $_.Snapshot.Field -eq 'WorkspaceRootPath'} | Select-Object -Last 1);if ($v.Count) {$path=[string]$v[0].Value}
            $path=ConvertTo-SiWorkspaceSharePath -Path $path
            if ($shareState.Checking) {return}
            $shareState.Path=$path;$shareState.Ready=$false;$shareState.Checking=$true;$shareState.Started=[datetime]::UtcNow
            $quoted="'"+$path.Replace("'","''")+"'"
            $command='try { $null=Get-Item -LiteralPath '+$quoted+' -ErrorAction Stop; $null=Get-ChildItem -LiteralPath '+$quoted+' -ErrorAction Stop | Select-Object -First 1; exit 0 } catch { exit 1 }'
            $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
            $shareState.Process=Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -ErrorAction Stop
            $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(500);$shareState.Timer=$timer
            $timer.Add_Tick({
                if (-not $shareState.Process.HasExited -and ([datetime]::UtcNow-$shareState.Started).TotalSeconds -lt 20) {return}
                $shareState.Timer.Stop()
                try {if (-not $shareState.Process.HasExited) {$shareState.Process.Kill();$shareState.Ready=$false} else {$shareState.Ready=($shareState.Process.ExitCode -eq 0)}} catch {$shareState.Ready=$false}
                finally {$shareState.Process.Dispose();$shareState.Process=$null;$shareState.Timer=$null;$shareState.Checking=$false}
                & $updateDetail
            }.GetNewClosure())
            $timer.Start();& $updateDetail
        } catch {$shareState.Ready=$false;$shareState.Checking=$false;$status.Text='Share check could not start: '+$_.Exception.Message;& $updateDetail}
    }.GetNewClosure()
    $steps.Add_SelectionChanged($updateDetail)
    $steps.Add_SelectionChanged({if ($null -ne $steps.SelectedItem -and $steps.SelectedItem.Id -eq 'share') {& $startShareCheck}}.GetNewClosure())
    $steps.Add_SelectionChanged({
        if ($phaseState.Changing -or $null -eq $steps.SelectedItem) {return}
        $phaseState.Changing=$true
        try {
            $index=Get-SiSimpleWorkflowIndex -Step ([string]$steps.SelectedItem.Id)
            $phases.SelectedIndex=$index
            $group=$phases.Items[$index]
            $phaseOptions.ItemsSource=@(foreach ($item in $steps.Items) {if ($item.Id -in $group.Steps -and $item.Id -notin @('discovery','tenant','permissions','administrators','teams','mail','collectors','orchestrator')) {[pscustomobject]@{Id=$item.Id;Title=($item.Title -replace '^\d+\.\s*','');Row=$item}}})
            $phaseOptions.SelectedItem=@($phaseOptions.Items | Where-Object {$_.Id -eq $steps.SelectedItem.Id})[0]
            $phaseOptions.Visibility=if ($phaseOptions.Items.Count -gt 1) {'Visible'} else {'Collapsed'}
            $window.FindName('Title').Text=$group.Title
        } finally {$phaseState.Changing=$false}
    }.GetNewClosure())
    $phases.Add_SelectionChanged({
        if ($phaseState.Changing -or $null -eq $phases.SelectedItem) {return}
        if ($phases.SelectedIndex -gt 1 -and -not $shareState.Ready) {
            $phaseState.Changing=$true
            try {$phases.SelectedIndex=1;$shareRow=@($steps.Items | Where-Object Id -eq 'share')[0];$steps.SelectedItem=$shareRow;$phaseOptions.ItemsSource=@([pscustomobject]@{Id='share';Title='Shared locations';Row=$shareRow});$phaseOptions.SelectedIndex=0;$phaseOptions.Visibility='Collapsed'} finally {$phaseState.Changing=$false}
            $status.Text='The CSV and log share must be accessible before continuing.';& $startShareCheck;& $updateDetail;return
        }
        if ($phases.SelectedIndex -gt 2 -and -not $machineState.Ready) {
            $phaseState.Changing=$true
            try {$phases.SelectedIndex=2;$prereqRow=@($steps.Items | Where-Object Id -eq 'prerequisites')[0];$steps.SelectedItem=$prereqRow;$phaseOptions.ItemsSource=@([pscustomobject]@{Id=$prereqRow.Id;Title='Prerequisites';Row=$prereqRow});$phaseOptions.SelectedIndex=0;$phaseOptions.Visibility='Collapsed'} finally {$phaseState.Changing=$false}
            $status.Text='Install and verify all prerequisites before continuing.';& $updateDetail;return
        }
        if ($phases.SelectedIndex -gt 3 -and (-not (Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions).Ready -or -not $serviceShareState.Ready -or (-not $OfflineSmokeTest -and ($serviceShareState.Sid -cne [string]$serviceState.Identity.Sid -or $serviceShareState.Path -ine $shareState.Path)) -or ($scope.Exchange -and (-not $exchangeRoleState.Ready -or (-not $OfflineSmokeTest -and $exchangeRoleState.Sid -cne [string]$serviceState.Identity.Sid))))) {
            $phaseState.Changing=$true
            try {$phases.SelectedIndex=3;$serviceRow=@($steps.Items | Where-Object Id -eq 'service')[0];$steps.SelectedItem=$serviceRow;$phaseOptions.ItemsSource=@([pscustomobject]@{Id=$serviceRow.Id;Title='Service account';Row=$serviceRow});$phaseOptions.SelectedIndex=0;$phaseOptions.Visibility='Collapsed'} finally {$phaseState.Changing=$false}
            $status.Text='Create or select a service account before continuing.';& $updateDetail;return
        }
        $group=$phases.SelectedItem
        $firstId=if ($group.Steps[0] -eq 'discovery') {'share'} elseif ($group.Steps[0] -eq 'tenant') {'identity'} elseif ($group.Steps[0] -eq 'teams') {'sharepoint'} else {$group.Steps[0]}
        $first=@($steps.Items | Where-Object {$_.Id -eq $firstId})
        if ($first.Count) {$steps.SelectedItem=$first[0]}
    }.GetNewClosure())
    $phaseOptions.Add_SelectionChanged({
        if (-not $phaseState.Changing -and $null -ne $phaseOptions.SelectedItem) {$steps.SelectedItem=$phaseOptions.SelectedItem.Row}
    }.GetNewClosure())
    $runPassiveChecks = {
        if (-not [string]::IsNullOrWhiteSpace($tenantBox.Text)) {$window.FindName('RunLocalBatch').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
    }.GetNewClosure()
    $refresh = {
        try {
            $assessmentState.Generation++
            if ([string]::IsNullOrWhiteSpace($tenantBox.Text)) {
                $steps.ItemsSource=$null
                $window.FindName('Title').Text='Select or create a profile'
                $window.FindName('StepState').Text='Profile required'
                $window.FindName('Detail').Text=$discovery.Message
                $window.FindName('GuidedAction').Visibility='Collapsed'
                $window.FindName('Action').Text=if ($discovery.State -in @('Unavailable','CorrectionRequired')) {'Correct the configuration issue, then reopen the assistant.'} else {'Select an existing profile or choose Create new profile.'}
                $status.Text=$discovery.Message
                return
            }
            $observations.Clear()
            $reviewed.Clear()
            $taskPlanState.Plan=$null
            if ($scope.Set) {
                $chosen=@(foreach ($name in @('AD','Exchange','SharePoint','Mail')) { if ($scope[$name]) { $name } })
            $observations['welcome']=@{State='Scope selected for this session';Detail=('Selection: ' + ($chosen -join ', ') + '. No jobs enabled automatically.')}
            }
            $index = [Math]::Max(0, $steps.SelectedIndex)
            $steps.ItemsSource = @(foreach ($step in Get-SiConfigurationSteps) {[pscustomobject]@{Id=$step.Id;Title=$step.Title;State='Local checks pending';Detail=$step.Detail;Action=$step.Action}})
            $steps.SelectedIndex = [Math]::Min($index, $steps.Items.Count - 1)
            $status.Text = 'Configuration assessment completed. External checks remain.'
        } catch {
            $steps.ItemsSource = $null
            $status.Text = 'Assessment failed: check the profile name and local access.'
        }
    }.GetNewClosure()
    $profileChoices.Add_SelectionChanged({
        if ($null -ne $profileChoices.SelectedItem -and $tenantBox.IsEnabled) {
            $tenantBox.Text=[string]$profileChoices.SelectedItem
            $window.FindName('ProfileDiscoveryStatus').Text='Profile selected: '+$tenantBox.Text+'. '+$(if ($discovery.GlobalState -eq 'Missing') {'Global configuration is missing.'} else {'Global configuration found.'})
            & $refresh
            if (-not $profileSelectionState.Preparing) {& $runPassiveChecks}
        }
    }.GetNewClosure())
    $tenantBox.Add_TextChanged({
        $assessmentState.Generation++
        $serviceState.Identity=$null;$serviceState.Created=$false
        $observations.Clear(); $reviewed.Clear(); $editState.Snapshot = $null; $scope.Set=$true; $taskPlanState.Plan=$null; $steps.ItemsSource = $null
        foreach ($name in @('Cloud','AD','Exchange','SharePoint','Teams','Mail')) { $scope[$name]=$false }
        $scope.Cloud=$true;$scope.AD=$true;$scope.Exchange=$true;$scope.Teams=$false;$scope.Mail=$true
        foreach ($controlName in @('Title','StepState','Detail','Action')) { $window.FindName($controlName).Text = '' }
        $window.FindName('Editor').Visibility = 'Collapsed';$certificateExportButton.Visibility='Collapsed';$identityButton.Visibility='Collapsed';$createGlobalButton.Visibility='Collapsed';$storageButton.Visibility='Collapsed';$applyPendingButton.Visibility='Collapsed';$discardPendingButton.Visibility='Collapsed'
        $window.FindName('GuidedAction').Visibility='Collapsed'
        foreach ($controlName in @('CreateProfile','AuditPrerequisites','InstallPrerequisites','InstallSigningCertificate','StopPrerequisites','ConfigureService','CreateService','CreateDirectoryService','CheckShare','CheckShareAsService','CheckShareWrite','CheckTask','CheckCertificate','SetCertificate','CheckMail','TestMail','StopMail','CheckCatalog','SetEnabled','SetUploadCategory','SetScope','CheckService','CheckServiceRights','PlanTask','RegisterTask','EditOrchestratorSettings','OpenOrchestratorConsole','PlanBootstrap','ReconcileBootstrap','CheckPermissions','StopPermissions')) { $window.FindName($controlName).Visibility = 'Collapsed' }
        $status.Text='Profile changed: refresh the assessment before continuing.'
    }.GetNewClosure())
    $window.FindName('Refresh').Add_Click({& $refresh;& $runPassiveChecks}.GetNewClosure())
    $discardPendingButton.Add_Click({
        if (($pendingEdits.Count -or $pendingActions.Count) -and [Windows.MessageBox]::Show($window,'Discard all pending actions and settings?','Discard pending changes','YesNo','Question') -eq 'Yes') {
            $discardedProfile=@($pendingActions | Where-Object Kind -eq 'Profile' | Select-Object -First 1)
            Clear-SiPendingEdits -Queue $pendingEdits;Clear-SiPendingCreations -Queue $pendingActions;$tenantBox.IsEnabled=$true
            if ($discardedProfile.Count -and (Read-SiPrivateJson (Join-Path $root ('Config/Tenants/'+$discardedProfile[0].Data.ProfileKey+'.local'))).State -eq 'Missing') {
                $profileChoices.Items.Remove($discardedProfile[0].Data.ProfileKey);$tenantBox.Text='';$window.FindName('ProfileDiscoveryStatus').Text='Profile creation discarded; select or create a profile.';& $refresh
            } else {
                if ($phases.SelectedIndex -gt 3 -and -not (Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions).Ready) {$phases.SelectedIndex=3}
                & $updateDetail
            }
            $status.Text='Pending changes discarded.'
        }
    }.GetNewClosure())
    $applyPendingButton.Add_Click({
        try {
            if ($null -ne $assessmentState.Active -or $assessmentState.Pending) {throw 'Wait for local checks to finish before applying changes.'}
            if (-not $machineState.Ready) {throw 'Install and verify all prerequisites before applying configuration.'}
            if (-not $shareState.Ready) {throw 'The CSV and log share must be accessible before applying configuration.'}
            if (-not $serviceShareState.Ready -or (-not $OfflineSmokeTest -and ($serviceShareState.Sid -cne [string]$serviceState.Identity.Sid -or $serviceShareState.Path -ine $shareState.Path))) {throw 'Verify service-account share access before applying configuration.'}
            if (-not (Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions).Ready) {throw 'Create or select a service account before applying configuration.'}
            if ($scope.Exchange -and (-not $exchangeRoleState.Ready -or (-not $OfflineSmokeTest -and $exchangeRoleState.Sid -cne [string]$serviceState.Identity.Sid))) {throw 'Verify Exchange role-group membership before applying configuration.'}
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) {throw 'Another operation is already running.'}
            if ($bootstrapState.ProposalAwaiting) {throw 'Review the bootstrap proposal before applying the remaining actions.'}
            if (@($pendingActions | Where-Object Kind -eq 'Bootstrap').Count -and (-not $tenantSession.PreflightReady -or $tenantSession.Profile -cne $tenantBox.Text)) {throw 'Connect to Microsoft 365 and complete the account preflight before applying.'}
            if (@($pendingEdits | Where-Object {$_.Snapshot.Tenant -cne $tenantBox.Text}).Count) {throw 'The selected profile differs from pending changes; return to that profile or discard them.'}
            if (@($pendingActions | Where-Object {$_.Kind -eq 'Profile' -and $_.Data.ProfileKey -cne $tenantBox.Text}).Count) {throw 'The selected profile differs from the pending new profile.'}
            if ($OfflineSmokeTest -and @($pendingActions | Where-Object {$_.Kind -in @('LocalAccount','DomainAccount','Task','Prerequisites','SigningCertificate','Bootstrap','MailTest','ShareWrite','Console','PermissionTest')}).Count) {throw 'Real setup, message, console, and external checks are disabled in offline smoke tests.'}
            if (@($pendingActions | Where-Object State -ne 'Pending').Count) {throw 'An earlier action has an uncertain result. Inspect it and discard it before retrying.'}
            if ($pendingActions.Count -or $pendingEdits.Count) {$taskOfferState.HadChanges=$true}
            foreach ($item in $pendingActions) {Write-SiAssistantLog -Log $assistantLog -Event Queue -Step 'final' -Action ('Action_'+$item.Kind) -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}
            foreach ($item in $pendingEdits) {Write-SiAssistantLog -Log $assistantLog -Event Queue -Step 'final' -Action ('Setting_'+$item.Snapshot.Field) -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}
            if (@($pendingActions | Where-Object {$_.Kind -eq 'Task' -and $_.Data.EnableAfterCreation}).Count) {Write-SiAssistantLog -Log $assistantLog -Event Queue -Step 'final' -Action 'TaskEnableRequested' -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}
            foreach ($fileAction in @($pendingActions | Where-Object Kind -eq 'FileEdit')) {
                if (@($pendingEdits | Where-Object {$_.Snapshot.Path -ieq $fileAction.Data.Path}).Count) {throw 'A file has both a general edit and a separate file edit. Apply one group first and review the other again.'}
                $d=$fileAction.Data
                $null=New-SiPendingFileAction -Type $d.Type -Plan $d.Plan -Value $d.Value -Worker $d.Worker -Root $d.Root -ScriptPath $d.ScriptPath
            }
            $details=(Get-SiPendingCreationReview -Queue $pendingActions)+"`n`n"+(Get-SiPendingEditReview -Queue $pendingEdits)
            if ($details.Length -gt 24000) {throw 'Pending review is too large. Apply fewer settings at a time.'}
            if (-not $OfflineSmokeTest -and -not $installState.Confirmed) {
                $summary=[Collections.Generic.List[string]]::new()
                foreach ($action in $pendingActions) {
                    switch ($action.Kind) {
                        'Profile' {$summary.Add('Create profile: '+$action.Data.ProfileKey+'  |  Environment: '+$action.Data.EnvironmentKey);foreach ($edit in $action.Data.PendingEdits) {$summary.Add('  '+$edit.Field+': '+$edit.Value)}}
                        'Global' {$summary.Add('Create global configuration');foreach ($edit in $action.Data.PendingEdits) {$summary.Add('  '+$edit.Field+': '+$edit.Value)}}
                        'Bootstrap' {$summary.Add('Configure Microsoft 365 application and Teams / SharePoint: '+$action.Data.Plan.BootstrapPlan.DisplayName);$summary.Add('  This creates or changes Microsoft 365 resources; generated settings are saved automatically when conflict-free.')}
                        default {$summary.Add($action.Label)}
                    }
                    $summary.Add('')
                }
                if ($pendingEdits.Count) {$summary.Add('Configuration changes');$summary.Add((Get-SiConcisePendingEditReview -Queue $pendingEdits))}
                [xml]$confirmXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ConfirmConfigurationDialog.xaml') -Raw
                $confirm=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($confirmXaml));$confirm.Owner=$window
                $confirm.FindName('Summary').Text=$summary -join [Environment]::NewLine
                $confirm.FindName('Apply').Add_Click({$confirm.DialogResult=$true}.GetNewClosure())
                if ($confirm.ShowDialog() -ne $true) {return}
            }
            if (-not $applyProgressState.Active) {
                if (-not $OfflineSmokeTest) {$applyProgressState.Progress=Start-SiApplyProgress -LibraryRoot $assistantLibRoot -Message 'Preparing the reviewed configuration...'}
                $applyProgressState.Active=$true;$applyProgressState.Starts++
            }
            & $setApplyProgress 'Preparing the reviewed configuration...'
            $signingActions=@($pendingActions | Where-Object Kind -eq 'SigningCertificate')
            if ($signingActions.Count -gt 1) {throw 'Multiple code-signing certificate actions require review.'}
            if ($signingActions.Count) {
                & $setApplyProgress 'Installing the code-signing certificate...'
                $certificateAction=$signingActions[0]
                $result=Invoke-SiPendingSigningCertificate -Action $certificateAction
                [void]$pendingActions.Remove($certificateAction)
                $observations['prerequisites']=@{State=$result.State;Detail=$result.Detail}
            }
            $prerequisiteActions=@($pendingActions | Where-Object Kind -eq 'Prerequisites')
            if ($prerequisiteActions.Count -gt 1) {throw 'Multiple prerequisite actions require review.'}
            if ($prerequisiteActions.Count) {
                & $setApplyProgress 'Installing prerequisites...'
                $installState.Action=$prerequisiteActions[0]
                $installState.Run=Start-SiPendingPrerequisiteAction -Action $installState.Action
                $installState.Confirmed=$true;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$installState.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $installState.Run -or -not $installState.Run.Process.HasExited) {return}
                    $installState.Timer.Stop();$continue=$false
                    try {
                        $installation=Complete-SiPrerequisiteInstallation -Run $installState.Run
                        $observations['prerequisites']=@{State=$installation.State;Detail=$installation.Detail}
                        if ($installation.State -eq 'InstalledAuditPassed') {[void]$pendingActions.Remove($installState.Action);$continue=$true}
                        else {$installState.Action.State='Uncertain';$status.Text='Prerequisite installation needs review; remaining actions were not started.'}
                    } catch {$installState.Action.State='Uncertain';$status.Text='Prerequisite result unconfirmed; remaining actions were not started.'}
                    finally {$installState.Run=$null;$installState.Timer=$null;$installState.Action=$null;$window.IsEnabled=$true}
                    if ($continue) {$applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} else {$installState.Confirmed=$false;& $updateDetail;& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
                }.GetNewClosure())
                $status.Text='Installing prerequisites; remaining actions will follow after a successful audit.';$timer.Start();return
            }
            & $setApplyProgress 'Creating the reviewed local configuration...'
            $creations=Invoke-SiPendingCreations -Queue $pendingActions
            foreach ($item in $creations.Completed) {
                if ($item.Kind -eq 'LocalAccount') {$identity=[pscustomobject]@{Account=($env:COMPUTERNAME+'\'+$item.Result.Name);Sid=[string]$item.Result.SID;Managed=$false;Detail='Account created; rights and access are not verified.'};$serviceState.Identity=$identity;Set-SiServiceIdentityObservation -Observations $observations -Identity $identity}
                elseif ($item.Kind -eq 'DomainAccount') {$identity=[pscustomobject]@{Account=($item.Result.Name+'@'+$item.Result.Domain);Sid=[string]$item.Result.Sid;Managed=$false;Detail='Account created; rights and access are not verified.'};$serviceState.Identity=$identity;Set-SiServiceIdentityObservation -Observations $observations -Identity $identity}
                elseif ($item.Kind -eq 'Profile') {$window.FindName('ProfileDiscoveryStatus').Text='Profile created: '+$tenantBox.Text+'.'}
                elseif ($item.Kind -eq 'Global') {$discovery.GlobalState='Present'}
            }
            $bootstrapActions=@($pendingActions | Where-Object Kind -eq 'Bootstrap')
            if ($bootstrapActions.Count -gt 1) {throw 'Multiple bootstrap actions require review.'}
            if ($bootstrapActions.Count) {
                & $setApplyProgress 'Preparing Microsoft 365 setup. Follow the sign-in window if prompted...'
                $action=$bootstrapActions[0]
                if ($action.State -ne 'Pending') {throw 'The bootstrap action needs review.'}
                try {
                    $null=New-SiPendingBootstrapIntent -Plan $action.Data.Plan -Parent $action.Data.Parent
                    $stage=New-SiBootstrapStage -Plan $action.Data.Plan -Parent $action.Data.Parent
                    $bootstrapState.Run=Start-SiBootstrapRun -Stage $stage;$action.State='Running'
                } catch {$action.State='Uncertain';throw 'Bootstrap preparation or launch was interrupted. Inspect its local folder and tenant before retrying.'}
                $bootstrapState.Action=$action;$installState.Confirmed=$true;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$bootstrapState.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $bootstrapState.Run -or -not $bootstrapState.Run.Process.HasExited) {return}
                    $bootstrapState.Timer.Stop();$continue=$false;$outcome=$null
                    try {
                        $outcome=Complete-SiBootstrapRun -Run $bootstrapState.Run
                        $detail=$outcome.Detail+"`nStaging directory: "+$bootstrapState.Run.Stage.Path+"`nProposal: "+$outcome.ProposalPath
                        $observations['identity']=@{State=$outcome.State;Detail=$detail}
                        if ($outcome.State -eq 'ReviewRequired') {
                            $bootstrapState.Proposal=$outcome.ProposalPath
                            $bootstrapState.ProposalAwaiting=$true
                            $candidate=$bootstrapState.Action.Data.Plan.BootstrapPlan
                            $bootstrapState.NewProfilePlan=if ($candidate.ProfileSHA256 -ceq 'ABSENT') {$candidate} else {$null}
                            [void]$pendingActions.Remove($bootstrapState.Action)
                            $status.Text='Bootstrap completed. Review its proposal before applying remaining actions.'
                            try {
                                $plan=New-SiBootstrapReconciliation -Root $root -Tenant $tenantBox.Text -ProposalPath $bootstrapState.Proposal -NewProfilePlan $bootstrapState.NewProfilePlan
                                if (Add-SiAutomaticBootstrapReconciliation -Plan $plan -Queue $pendingEdits) {
                                    $bootstrapState.ProposalAwaiting=$false
                                    $proposalCert=Read-SiProposal -Path $bootstrapState.Proposal
                                    if ([string]$proposalCert.Document.Thumbprint -match '^[0-9a-fA-F]{40}$' -and [string]$proposalCert.Document.Thumbprint -ieq [string]$proposalCert.Document.Thumb) {$certificateExportState.Thumb=[string]$proposalCert.Document.Thumbprint;$certificateExportState.Profile=$tenantBox.Text;$certificateExportState.Exported=$false;$bootstrapState.OfferCertificate=$true}
                                    $status.Text='Microsoft 365 setup finished. Saving the generated configuration...';$continue=$true
                                }
                            } catch {$status.Text='Automatic proposal preparation failed. Review the proposal before applying remaining actions.'}
                        } else {$bootstrapState.Action.State='Uncertain';$status.Text=$detail}
                    } catch {$bootstrapState.Action.State='Uncertain';$status.Text='Bootstrap result is uncertain; inspect the tenant and staging logs.'}
                    finally {$bootstrapState.Run=$null;$bootstrapState.Timer=$null;$bootstrapState.Action=$null;$window.IsEnabled=$true}
                    if ($continue) {$applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} else {
                        $installState.Confirmed=$false;& $updateDetail
                        if ($outcome.State -eq 'ReviewRequired') {
                            if ($bootstrapState.ProposalAwaiting) {& $finishApplyProgress 'Microsoft 365 setup finished. Review the bootstrap proposal before applying its changes.' 'Next'}
                            else {& $finishApplyProgress 'Microsoft 365 setup finished. The proposed changes are in the final review and must still be applied.' 'Next'}
                        } else {& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
                        if ($bootstrapState.OfferCertificate -and -not $OfflineSmokeTest) {$bootstrapState.OfferCertificate=$false;$certificateExportButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
                    }
                }.GetNewClosure())
                & $setApplyProgress 'Configuring Microsoft 365. Follow the sign-in window if prompted...'
                $status.Text='Microsoft 365 setup is running; generated settings will be applied when conflict-free.';$timer.Start();return
            }
            & $setApplyProgress 'Saving the reviewed configuration changes...'
            $result=Invoke-SiPendingEdits -Queue $pendingEdits
            Invoke-SiPendingFileActions -Queue $pendingActions
            $writeActions=@($pendingActions | Where-Object Kind -eq 'ShareWrite')
            if ($writeActions.Count) {
                $probe=$writeActions[0]
                try {$probe.State='Running';$outcome=Test-SiSharedPathWrite -Path $probe.Data.Path -MarkerName $probe.Data.Marker -ExpectedSid $probe.Data.Sid
                    if ($outcome.State -notin @('WriteVerified','WriteVerifiedAsExpected')) {throw $outcome.Detail}
                    $observations['share']=@{State=$outcome.State;Detail=$outcome.Detail};[void]$pendingActions.Remove($probe)
                } catch {$probe.State='Uncertain';throw 'Share write result is uncertain. Inspect the reviewed marker before retrying.'}
            }
            $testAction=@($pendingActions | Where-Object {$_.Kind -in @('MailTest')} | Select-Object -First 1)
            if ($testAction.Count) {
                & $setApplyProgress 'Running the reviewed test...'
                $action=$testAction[0];$state=$mailState
                try {$state.Run=Start-SiMailTest -Plan $action.Data;$action.State='Running'} catch {$action.State='Uncertain';throw 'The test may have started. Inspect delivery or workflow history before retrying.'}
                $state.Action=$action;$installState.Confirmed=$true;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$state.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $state.Run -or -not $state.Run.Process.HasExited) {return}
                    $state.Timer.Stop();$continue=$false
                    try {
                        $outcome=Complete-SiMailTest -Run $state.Run
                        $key='mail'
                        $observations[$key]=@{State=$outcome.State;Detail=$outcome.Detail}
                        if ($outcome.State -in @('Completed','Accepted')) {[void]$pendingActions.Remove($state.Action);$continue=$true} else {$state.Action.State='Uncertain';$status.Text=$outcome.Detail}
                    } catch {$state.Action.State='Uncertain';$status.Text='Test result is uncertain; inspect delivery or workflow history.'}
                    finally {$state.Run=$null;$state.Timer=$null;$state.Action=$null;$window.IsEnabled=$true}
                    if ($continue) {$applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} else {$installState.Confirmed=$false;& $updateDetail;& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
                }.GetNewClosure())
                $status.Text='Running the reviewed test; remaining actions will follow on success.';$timer.Start();return
            }
            $permissionActions=@($pendingActions | Where-Object Kind -eq 'PermissionTest')
            if ($permissionActions.Count) {
                & $setApplyProgress 'Checking Microsoft 365 permissions...'
                $action=$permissionActions[0]
                try {$permissionState.Run=Start-SiPermissionCheck -Plan $action.Data;$action.State='Running'} catch {$action.State='Uncertain';throw 'Permission check launch is uncertain; inspect authentication and temporary files.'}
                $permissionState.Action=$action;$installState.Confirmed=$true;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$permissionState.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $permissionState.Run -or -not $permissionState.Run.Process.HasExited) {return}
                    $permissionState.Timer.Stop();$continue=$false
                    try {$outcome=Complete-SiPermissionCheck -Run $permissionState.Run;$observations['permissions']=@{State=$outcome.State;Detail=$outcome.Detail}
                        if ($outcome.State -ne 'Uncertain') {[void]$pendingActions.Remove($permissionState.Action);$continue=$true} else {$permissionState.Action.State='Uncertain';$status.Text=$outcome.Detail}
                    } catch {$permissionState.Action.State='Uncertain';$status.Text='Permission result is uncertain; inspect the retained private folder.'}
                    finally {$permissionState.Run=$null;$permissionState.Timer=$null;$permissionState.Action=$null;$window.IsEnabled=$true}
                    if ($continue) {$applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} else {$installState.Confirmed=$false;& $updateDetail;& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
                }.GetNewClosure())
                $status.Text='Permission check running; authentication or consent may be requested.';$timer.Start();return
            }
            $consoleActions=@($pendingActions | Where-Object Kind -eq 'Console')
            if ($consoleActions.Count) {
                & $setApplyProgress 'Waiting for the management console to close...'
                $action=$consoleActions[0]
                try {$consoleState.Run=Start-SiOrchestratorConsole -Plan $action.Data;$action.State='Running'} catch {$action.State='Uncertain';throw 'Console launch could not be confirmed; inspect running processes.'}
                $consoleState.Action=$action;$installState.Confirmed=$true;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$consoleState.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $consoleState.Run -or -not $consoleState.Run.Process.HasExited) {return}
                    $consoleState.Timer.Stop();$continue=$false
                    try {$outcome=Complete-SiOrchestratorConsole -Run $consoleState.Run
                        if ($outcome.ExitCode -eq 0) {[void]$pendingActions.Remove($consoleState.Action);$continue=$true} else {$consoleState.Action.State='Uncertain';$status.Text=$outcome.Detail}
                    } catch {$consoleState.Action.State='Uncertain';$status.Text='Console result is uncertain; review its changes.'}
                    finally {$consoleState.Run=$null;$consoleState.Timer=$null;$consoleState.Action=$null;$window.IsEnabled=$true}
                    if ($continue) {$applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} else {$installState.Confirmed=$false;& $updateDetail;& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
                }.GetNewClosure())
                $status.Text='Management console open; close it to continue.';$timer.Start();return
            }
            $installState.Confirmed=$false
            $taskActions=@($pendingActions | Where-Object {$_.Kind -eq 'Task'})
            if ($taskActions.Count -gt 1) {throw 'Multiple task actions require review.'}
            if ($taskActions.Count) {
                if ($OfflineSmokeTest) {throw 'Real task installation is disabled in offline smoke tests.'}
                if (-not $observations.ContainsKey('service') -or -not $observations.service.Sid) {throw 'The selected service account must be resolved before the task can be created.'}
                & $setApplyProgress 'Creating the scheduled task...'
                $taskAction=$taskActions[0]
                $identity=[pscustomobject]@{Account=$observations.service.Account;Sid=$observations.service.Sid;Managed=$false}
                $taskInstallerState.Run=Start-SiPendingTaskAction -Action $taskAction -Identity $identity
                $taskInstallerState.Action=$taskAction;$window.IsEnabled=$false
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromSeconds(1);$taskInstallerState.Timer=$timer
                $timer.Add_Tick({
                    if ($null -eq $taskInstallerState.Run -or -not $taskInstallerState.Run.Process.HasExited) {return}
                    $taskInstallerState.Timer.Stop();$taskResult=$null
                    try {
                        $taskResult=Complete-SiOrchestratorRegistration -Run $taskInstallerState.Run
                        if ($taskResult.State -eq 'CreatedDisabled' -and $taskInstallerState.Action.Data.EnableAfterCreation) {$taskResult=Enable-SiCreatedOrchestratorTask -Plan $taskInstallerState.Run.Plan}
                        if ($taskResult.State -in @('CreatedDisabled','CreatedDisabledReview','CreatedEnabled')) {[void]$pendingActions.Remove($taskInstallerState.Action)} else {$taskInstallerState.Action.State='Uncertain'}
                        $observations['orchestrator']=@{State=$taskResult.State;Detail=$taskResult.Detail};$status.Text=$taskResult.Detail
                    } catch {$taskInstallerState.Action.State='Uncertain';$status.Text='Task result unconfirmed. Inspect Task Scheduler before retrying.'}
                    finally {$taskInstallerState.Run=$null;$taskInstallerState.Timer=$null;$taskInstallerState.Action=$null;$window.IsEnabled=$true;$tenantBox.IsEnabled=($pendingEdits.Count -eq 0 -and $pendingActions.Count -eq 0);& $updateDetail}
                    if ($taskResult -and $taskResult.State -in @('CreatedDisabled','CreatedEnabled')) {
                        if ($pendingActions.Count -or $pendingEdits.Count) {& $finishApplyProgress ($status.Text+' Other changes still need review and application.') 'Next'}
                        else {& $finishApplyProgress $status.Text 'Success'}
                    } else {& $finishApplyProgress $status.Text 'Error'}
                }.GetNewClosure())
                $status.Text='Task installer running; complete its credential prompt.';$timer.Start();return
            }
            if (-not $pendingEdits.Count -and -not $pendingActions.Count) {$tenantBox.IsEnabled=$true}
            $savedIdentity=if ($observations.ContainsKey('service')) {$observations['service']} else {$null}
            $bootstrapReview=if ($observations.ContainsKey('identity') -and $observations['identity'].State -eq 'ReviewRequired') {$observations['identity']} else {$null}
            & $refresh
            if ($null -ne $savedIdentity) {$observations['service']=$savedIdentity}
            if ($null -ne $bootstrapReview) {$observations['identity']=$bootstrapReview}
            if ($null -ne $savedIdentity -or $null -ne $bootstrapReview) {& $updateDetail}
            $status.Text=('Applied '+$creations.Completed.Count+' creation action(s) and '+$result.Applied+' setting(s) in '+$result.Files+' file(s).')
            if ($pendingActions.Count -or $pendingEdits.Count) {& $finishApplyProgress 'Some changes were applied. Other actions remain in the final review.' 'Next'}
            else {& $finishApplyProgress 'The reviewed configuration changes were applied. No scheduled task has been created by this step.' 'Success'}
            if ($bootstrapState.OfferCertificate -and -not $pendingActions.Count -and -not $pendingEdits.Count) {
                $bootstrapState.OfferCertificate=$false
                if (-not $OfflineSmokeTest) {$certificateExportButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
            }
            $closeAfterSuccess=(-not $pendingActions.Count -and -not $pendingEdits.Count)
            if ($taskOfferState.HadChanges -and $closeAfterSuccess) {
                $taskOfferState.HadChanges=$false
                if ($OfflineSmokeTest) {$taskOfferState.Offers++}
                else {
                    try {
                        $confirmTask={param($message) [Windows.MessageBox]::Show($window,$message,'Scheduled task installer','YesNo','Question') -eq 'Yes'}.GetNewClosure()
                        $outcome=Invoke-SiTaskInstallerOffer -Root $root -Tenant $tenantBox.Text -PowerShell (Join-Path $PSHOME 'pwsh.exe') -Confirm $confirmTask
                        $status.Text=if ($outcome.State -eq 'Launched') {'Configuration applied. Scheduled-task installer started in a separate console; complete its prompts.'} else {'Configuration applied. Scheduled-task installer not opened.'}
                    } catch {$closeAfterSuccess=$false;$status.Text='Configuration applied. Task installer was not opened: '+$_.Exception.Message;[void][Windows.MessageBox]::Show($window,$status.Text,'Scheduled task installer','OK','Warning')}
                }
            }
            if ($closeAfterSuccess -and -not $OfflineSmokeTest) {$window.Close()}
        } catch {
            $installState.Confirmed=$false;$taskOfferState.HadChanges=$false;$status.Text='Apply stopped: '+$_.Exception.Message
            if ($applyProgressState.Active) {& $finishApplyProgress ($status.Text+' Some changes may already have been applied; review the result before retrying.') 'Error'}
            elseif (-not $OfflineSmokeTest) {[void][Windows.MessageBox]::Show($window,$status.Text,'Configuration assistant','OK','Warning')}
            & $updateDetail
        }
    }.GetNewClosure())
    $storageButton.Add_Click({
        if ($steps.SelectedItem.Id -ne 'share') {return}
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-WorkspaceShareDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=$window
            $current=[string](Get-SiSimpleStoragePreview -Root $root).Values.WorkspaceRootPath
            $globalPending=@($pendingActions | Where-Object Kind -eq 'Global' | Select-Object -First 1)
            if ($globalPending.Count) {$v=@($globalPending[0].Data.PendingEdits | Where-Object Field -eq 'WorkspaceRootPath' | Select-Object -Last 1);if ($v.Count) {$current=[string]$v[0].Value}}
            $edit=@($pendingEdits | Where-Object {$_.Snapshot.Step -eq 'share' -and $_.Snapshot.Field -eq 'WorkspaceRootPath'} | Select-Object -Last 1)
            if ($edit.Count) {$current=[string]$edit[0].Value}
            if ($current -match '^\\\\[^\\]+\\[^\\]+') {$dialog.FindName('Workspace').Text=$current}
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Save').Add_Click({
                try {
                    $workspace=ConvertTo-SiWorkspaceSharePath -Path $dialog.FindName('Workspace').Text
                    $creation=@($pendingActions | Where-Object Kind -eq 'Global' | Select-Object -First 1)
                    if ($creation.Count) {
                        $snapshot=Get-SiPendingCreationSnapshot -Action $creation[0] -Tenant $tenantBox.Text -Step share -Field WorkspaceRootPath
                        Set-SiPendingCreationSetting -Snapshot $snapshot -Value $workspace
                    } else {
                        $globalCurrent=Read-SiPrivateJson (Join-Path $root 'Config/SmartM365.global.local')
                        if ($globalCurrent.State -eq 'Missing') {
                            $action=New-SiPendingCreation -Kind Global -Data (Get-SiGlobalCreationPlan -Root $root)
                            $default=Get-SiPendingCreationSnapshot -Action $action -Tenant $tenantBox.Text -Step share -Field DefaultTenant
                            Set-SiPendingCreationSetting -Snapshot $default -Value $tenantBox.Text
                            $snapshot=Get-SiPendingCreationSnapshot -Action $action -Tenant $tenantBox.Text -Step share -Field WorkspaceRootPath
                            Set-SiPendingCreationSetting -Snapshot $snapshot -Value $workspace
                            Add-SiPendingCreation -Queue $pendingActions -Action $action
                        } elseif ($globalCurrent.State -eq 'Present') {
                            $snapshot=Get-SiEditSnapshot -Root $root -Tenant $tenantBox.Text -Step share -Field WorkspaceRootPath
                            if ($snapshot.Present -and $snapshot.Value -ceq $workspace) {
                                foreach ($old in @($pendingEdits | Where-Object {$_.Snapshot.Path -ieq $snapshot.Path -and $_.Snapshot.Field -ieq $snapshot.Field})) {[void]$pendingEdits.Remove($old)}
                            } else {Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingEdit -Snapshot $snapshot -Value $workspace)}
                        } else {throw 'The global configuration cannot be read safely.'}
                    }
                    $tenantBox.IsEnabled=(-not $pendingActions.Count -and -not $pendingEdits.Count)
                    $serviceShareState.Ready=[bool]$OfflineSmokeTest;$shareState.Ready=[bool]$OfflineSmokeTest;$shareState.Path='';$dialog.Close();& $startShareCheck;& $updateDetail;$status.Text='CSV and log share prepared for final review; no configuration was written.'
                } catch {[void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Share unavailable','OK','Warning')}
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch {$status.Text='Cannot open share settings: '+$_.Exception.Message}
    }.GetNewClosure())
    $createGlobalButton.Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) {throw 'Another operation is already running.'}
            $plan=Get-SiGlobalCreationPlan -Root $root
            $action=New-SiPendingCreation -Kind Global -Data $plan
            $snapshot=Get-SiPendingCreationSnapshot -Action $action -Tenant $tenantBox.Text -Step share -Field DefaultTenant
            Set-SiPendingCreationSetting -Snapshot $snapshot -Value $tenantBox.Text
            Add-SiPendingCreation -Queue $pendingActions -Action $action
            $tenantBox.IsEnabled=$false;$status.Text='Global configuration added to final review; no file was written.'
        } catch {$status.Text='Global configuration not created: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('RunLocalBatch').Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) {throw 'Another operation is already running.'}
            if ([string]::IsNullOrWhiteSpace($tenantBox.Text)) {return}
            if ($null -ne $assessmentState.Active) {$assessmentState.Pending=$true;return}
            if ($null -eq $assessmentState.Popup) {
                [xml]$progressXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-LocalChecksProgress.xaml') -Raw
                $assessmentState.Popup=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($progressXaml))
                $assessmentState.Popup.Owner=$window;$assessmentState.Popup.ShowActivated=$false
                if ($OfflineSmokeTest) {$assessmentState.Popup.ShowInTaskbar=$false;$assessmentState.Popup.Opacity=0}
                $assessmentState.Popup.Show()
            }
            $status.Text='Running local checks...'
            $assessmentState.Tenant=$tenantBox.Text;$assessmentState.StartedGeneration=$assessmentState.Generation
            $worker=[PowerShell]::Create()
            [void]$worker.AddCommand((Join-Path $assistantLibRoot 'SmartM365.LocalAssessmentWorker.ps1')).AddParameter('Root',$root).AddParameter('Tenant',$assessmentState.Tenant).AddParameter('LibraryRoot',$assistantLibRoot).AddParameter('OfflineFixture',[bool]$OfflineSmokeTest)
            $assessmentState.Active=$worker;$assessmentState.Async=$worker.BeginInvoke()
            $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(100);$assessmentState.Timer=$timer
            $timer.Add_Tick({
                if ($null -eq $assessmentState.Active -or -not $assessmentState.Async.IsCompleted) {return}
                $assessmentState.Timer.Stop();$worker=$assessmentState.Active;$result=$null;$failure=''
                try {$output=@($worker.EndInvoke($assessmentState.Async));if ($worker.HadErrors -or $output.Count -ne 1) {throw 'The local checks did not return a complete result.'};$result=$output[0]}
                catch {$failure=$_.Exception.Message}
                finally {$worker.Dispose();$assessmentState.Active=$null;$assessmentState.Async=$null;$assessmentState.Timer=$null}
                if ($assessmentState.Closed) {return}
                if ($assessmentState.Pending) {
                    $assessmentState.Pending=$false
                    if (-not [string]::IsNullOrWhiteSpace($tenantBox.Text)) {$window.FindName('RunLocalBatch').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));return}
                }
                $current=($assessmentState.StartedGeneration -eq $assessmentState.Generation -and $assessmentState.Tenant -ceq $tenantBox.Text)
                if ($null -ne $assessmentState.Popup) {$assessmentState.Popup.Close();$assessmentState.Popup=$null}
                if (-not $current) {return}
                if ($failure) {$machineState.Ready=$false;$machineState.Known=$false;$status.Text='Local checks failed: '+$failure;& $updateDetail;return}
                if ($OfflineSmokeTest) {$smokeUiState.PassiveRuns++}
                $machineState.Ready=([bool]$result.MachinePrerequisites.Ready -and -not $machineState.ValidationFailed);$machineState.Rows=@($result.MachinePrerequisites.Rows);$machineState.Known=$true
                $observations.Clear();$reviewed.Clear();$taskPlanState.Plan=$null
                if ($scope.Set) {$chosen=@(foreach ($name in @('AD','Exchange','SharePoint','Mail')) {if ($scope[$name]) {$name}});$observations['welcome']=@{State='Scope selected for this session';Detail=('Selection: '+($chosen -join ', ')+'. No jobs enabled automatically.')}}
                foreach ($item in $result.Results) {if ($item.Id -notin @('welcome','final')) {$observations[$item.Id]=@{State=$item.State;Detail=$item.Detail}}}
                if ($null -ne $serviceState.Identity) {Set-SiServiceIdentityObservation -Observations $observations -Identity $serviceState.Identity}
                if (@($pendingActions | Where-Object {$_.Kind -eq 'Profile' -and $_.Data.ProfileKey -ceq $tenantBox.Text}).Count) {$observations['tenant']=@{State='New profile pending review';Detail='The profile has not been created; local checks use existing files only.'}}
                if (-not $machineState.Ready -and $phases.SelectedIndex -gt 2) {$phases.SelectedIndex=2}
                & $updateDetail
                $status.Text=if ($machineState.LastInstall) {$machineState.LastInstall} else {'Configuration checks finished. The repository may have been read over the network; no tenant connection, account or task action occurred.'}
            }.GetNewClosure())
            $timer.Start()
        } catch {
            if ($null -ne $assessmentState.Active) {$assessmentState.Active.Dispose();$assessmentState.Active=$null}
            if ($null -ne $assessmentState.Popup) {$assessmentState.Popup.Close();$assessmentState.Popup=$null}
            $status.Text='Local checks failed: '+$_.Exception.Message
        }
    }.GetNewClosure())
    $window.FindName('ReconcileBootstrap').Add_Click({
        try {
            $proposalPath=[string]$bootstrapState.Proposal;$newProfilePlan=$bootstrapState.NewProfilePlan
            if (-not $proposalPath -or ($null -ne $newProfilePlan -and $newProfilePlan.ProfileKey -cne $tenantBox.Text)) {
                $picker=[Microsoft.Win32.OpenFileDialog]::new(); $picker.Title='Select a proposed profile to review'
                $picker.Filter='JSON files (*.json;*.json.txt)|*.json;*.json.txt|All files (*.*)|*.*'; $picker.CheckFileExists=$true
                if ($picker.ShowDialog($window) -ne $true) { return };$proposalPath=$picker.FileName;$newProfilePlan=$null
            }
            $plan=New-SiBootstrapReconciliation -Root $root -Tenant $tenantBox.Text -ProposalPath $proposalPath -NewProfilePlan $newProfilePlan
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-BootstrapReconcileDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $lines=[Collections.Generic.List[string]]::new()
            $lines.Add('Profile: '+$plan.Tenant); $lines.Add('Proposal: '+$plan.ProposalPath); $lines.Add('')
            $lines.Add('Proposed additions:')
            if ($plan.Additions.Count -eq 0) { $lines.Add('(none)') }
            foreach ($item in $plan.Additions) { $lines.Add(('  {0} = {1}' -f $item.Field,$item.Value)) }
            $lines.Add('');$lines.Add('New-profile defaults to replace:')
            if ($plan.Replacements.Count -eq 0) {$lines.Add('(none)')}
            foreach ($item in $plan.Replacements) {$lines.Add(('  {0}: {1} -> {2}' -f $item.Field,$item.Current,$item.Value))}
            $lines.Add(''); $lines.Add('Blocked groups: '+($plan.BlockedGroups -join ', ')); $lines.Add('Conflicting fields: '+($plan.Conflicts -join ', '))
            foreach ($conflict in $plan.ConflictValues) { $lines.Add(('  {0}: current {1} -> proposed {2}' -f $conflict.Field,$conflict.CurrentJson,$conflict.ProposedJson)) }
            $lines.Add(''); $lines.Add('Other values added from the template:')
            if ($plan.TemplateAdditions.Count -eq 0) { $lines.Add('(none)') }
            foreach ($item in $plan.TemplateAdditions) { $lines.Add(('  {0} = {1}' -f $item.Field,$item.ValueJson)) }
            if ($plan.BlockedTemplateKeys.Count -gt 0) { $lines.Add('Addition blocked: the template would add fields in a conflicting group: '+($plan.BlockedTemplateKeys -join ', ')) }
            $lines.Add(''); $lines.Add('Only unchanged defaults from this newly created profile can be replaced. Other existing values are preserved.')
            $dialog.FindName('Summary').Text=$lines -join [Environment]::NewLine
            $dialog.FindName('Apply').IsEnabled=$true
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Apply').Add_Click({
                try {
                    if ($plan.BlockedTemplateKeys.Count -eq 0 -and ($plan.Additions.Count+$plan.Replacements.Count) -gt 0) {Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingReconciliationEdit -Plan $plan)}
                    $bootstrapState.ProposalAwaiting=$false
                    $tenantBox.IsEnabled=$false;$dialog.Close();& $updateDetail
                    $status.Text=if ($plan.BlockedTemplateKeys.Count) {'Proposal reviewed; blocked template changes were not added.'} else {'Bootstrap changes added to final review; other existing values are preserved.'}
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Additions refused','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Review failed: '+$_.Exception.Message }
    }.GetNewClosure())
    $certificateExportButton.Add_Click({
        try {
            if (-not $certificateExportState.Thumb -or $certificateExportState.Profile -cne $tenantBox.Text) {throw 'The application certificate is not available for this profile.'}
            [xml]$exportXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-CertificateExportDialog.xaml') -Raw
            $exportDialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($exportXaml));$exportDialog.Owner=$window
            $exportDialog.FindName('Destination').Text='Destination: '+(Join-Path $root 'Config/Certificats')
            $exportDialog.FindName('Export').Add_Click({
                $first=$null;$second=$null
                try {
                    $first=$exportDialog.FindName('Password').SecurePassword;$second=$exportDialog.FindName('Confirm').SecurePassword
                    if (-not (Test-SiSecurePasswordEqual -First $first -Second $second)) {throw 'The passwords do not match.'}
                    $path=Export-SiApplicationCertificate -Root $root -Thumbprint $certificateExportState.Thumb -Password $first
                    $certificateExportState.Exported=$true;$status.Text='Application certificate exported: '+$path
                    $exportDialog.DialogResult=$true
                } catch {$exportDialog.FindName('Message').Text=$_.Exception.Message}
                finally {if ($first) {$first.Dispose()};if ($second) {$second.Dispose()}}
            }.GetNewClosure())
            $exportDialog.Add_Closed({$exportDialog.FindName('Password').Clear();$exportDialog.FindName('Confirm').Clear()}.GetNewClosure())
            [void]$exportDialog.ShowDialog();& $updateDetail
        } catch {$status.Text='Certificate export unavailable: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('CheckShareWrite').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-SharedWriteDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $hasService=($observations.ContainsKey('service') -and [bool]$observations['service'].Sid -and -not [bool]$observations['service'].Managed)
            $dialog.FindName('RequireService').IsEnabled=$hasService
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Check').Add_Click({
                $markerPath=''
                try {
                    $checkedPath=[string]$dialog.FindName('Path').Text
                    if (-not [IO.Path]::IsPathFullyQualified($checkedPath)) { throw 'Enter an absolute local or UNC path.' }
                    $marker='.smartm365-setup-probe-'+[guid]::NewGuid().ToString('N')+'.tmp'
                    $markerPath=Join-Path $checkedPath $marker
                    $expectedSid=''
                    if ($dialog.FindName('RequireService').IsChecked -eq $true) {
                        if (-not $observations.ContainsKey('service') -or -not $observations['service'].Sid) {throw 'Resolve the service account first.'}
                        if ([bool]$observations['service'].Managed) {throw 'The managed-account share probe has been removed.'}
                        $expectedSid=[string]$observations['service'].Sid
                    }
                    $intent=New-SiPendingTestAction -Kind ShareWrite -Data ([pscustomobject]@{Path=$checkedPath;Marker=$marker;Sid=$expectedSid})
                    Add-SiPendingCreation -Queue $pendingActions -Action $intent
                    $tenantBox.IsEnabled=$false;$status.Text='Share write test added to final review; no marker was created.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,($_.Exception.Message + [Environment]::NewLine + 'If the test started, inspect this marker before retrying: ' + $markerPath),'Test unconfirmed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckMail').Add_Click({
        try {
            $result=Get-SiTenantMailObservation -Root $root -Tenant $tenantBox.Text
            $label=if ($result.State -eq 'ReviewNeeded') { 'Local settings need review' } else { 'Partial local check' }
            $detail='Fields to review: '+($result.Issues -join ', ')+[Environment]::NewLine+'Missing or unresolved inherited fields: '+($result.Unresolved -join ', ')+[Environment]::NewLine+$result.Detail
            $observations['mail']=@{State=$label;Detail=$detail}
            $window.FindName('StepState').Text=$label; $window.FindName('Action').Text=$detail
            $status.Text='Local read only; no email sent.'
        } catch { $status.Text='Check failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('SetCertificate').Add_Click({
        try {
            $snapshot=Get-SiCertificateEditSnapshot -Root $root -Tenant $tenantBox.Text
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-CertificateEditDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $dialog.FindName('Current').Text="Current Thumb: $($snapshot.Thumb)`nCurrent Thumbprint: $($snapshot.Thumbprint)"
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Save').Add_Click({
                try {
                    $thumb=$dialog.FindName('Thumbprint').Text.Replace(' ','').ToUpperInvariant()
                    if ($thumb -notmatch '^[0-9A-F]{40}$' -or $thumb -match '^0+$') { throw 'The thumbprint must contain 40 nonzero hexadecimal characters.' }
                    $pending=New-SiPendingCertificateEdit -Snapshot $snapshot -Thumbprint $thumb
                    Add-SiPendingEdit -Queue $pendingEdits -Edit $pending
                    $tenantBox.IsEnabled=$false;$dialog.Close()
                    $status.Text='Certificate thumbprints added to final review; no configuration was written.'
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Save refused','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('StopPrerequisites').Add_Click({
        try {
            if ($null -eq $installState.Run) { return }
            if ([Windows.MessageBox]::Show($window,'Stop the process? Installed modules and repository trust will not be rolled back.','Stop installation','YesNo','Warning') -ne 'Yes') { return }
            Stop-SiPrerequisiteInstallation -Run $installState.Run
            $window.FindName('StopPrerequisites').IsEnabled=$false
            $status.Text='Stop requested; installation may be partial. Wait for the check.'
        } catch { $status.Text='Stop unconfirmed; wait for the process or inspect its state.' }
    }.GetNewClosure())
    $window.Add_Closing({ param($sender,$eventArgs) if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) { $eventArgs.Cancel=$true; $status.Text='An operation or management console is open; wait for it to finish before closing.';return };if (-not $OfflineSmokeTest -and ($pendingEdits.Count -gt 0 -or $pendingActions.Count -gt 0) -and [Windows.MessageBox]::Show($window,'Pending changes have not been applied. Discard them and close?','Pending changes','YesNo','Warning') -ne 'Yes') {$eventArgs.Cancel=$true} }.GetNewClosure())
    $window.Add_Closed({
        $assessmentState.Closed=$true
        if ($null -ne $assessmentState.Timer) {$assessmentState.Timer.Stop()}
        if ($null -ne $assessmentState.Popup) {try {$assessmentState.Popup.Close()} catch {}; $assessmentState.Popup=$null}
        if ($null -ne $assessmentState.Active) {try {[void]$assessmentState.Active.BeginStop($null,$null)} catch {}}
        Clear-SiPendingEdits -Queue $pendingEdits;Clear-SiPendingCreations -Queue $pendingActions
    }.GetNewClosure())
    $window.FindName('StopMail').Add_Click({
        try {
            if ($null -eq $mailState.Run) { return }
            if ([Windows.MessageBox]::Show($window,'Stop the test? The message may already have been sent. Check delivery before retrying.','Stop test','YesNo','Warning') -ne 'Yes') { return }
            Stop-SiMailTest -Run $mailState.Run
            $window.FindName('StopMail').IsEnabled=$false
            $status.Text='Stop requested; the result may be uncertain.'
        } catch { $status.Text='Stop unconfirmed; check the recipient and delivery.' }
    }.GetNewClosure())
    $window.FindName('TestMail').Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run) { throw 'Another operation is already running.' }
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-MailTestDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Send').Add_Click({
                try {
                    $transport=[string]$dialog.FindName('Transport').SelectedItem.Content
                    $plan=New-SiMailTestPlan -Root $root -Tenant $tenantBox.Text -From $dialog.FindName('From').Text -To $dialog.FindName('To').Text -Transport $transport
                    Add-SiPendingCreation -Queue $pendingActions -Action (New-SiPendingTestAction -Kind MailTest -Data $plan)
                    $tenantBox.IsEnabled=$false;$status.Text='Mail test added to final review; nothing was sent.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,($_.Exception.Message + [Environment]::NewLine + 'If the test started, check recipient and delivery before retrying.'),'Test unconfirmed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('InstallSigningCertificate').Add_Click({
        try {
            $intent=New-SiSigningCertificateIntent -RepositoryRoot (Split-Path -Parent $runtimeRoot)
            Add-SiPendingCreation -Queue $pendingActions -Action $intent
            $tenantBox.IsEnabled=$false;& $updateDetail
            $status.Text='Code-signing certificate installation added to final review; no certificate was installed.'
        } catch {$status.Text='Cannot prepare code-signing certificate installation: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('ShowInstallDetails').Add_Click({
        try {
            [xml]$detailXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-PrerequisiteDetailsDialog.xaml') -Raw
            $detailWindow=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($detailXaml));$detailWindow.Owner=$window
            $detailWindow.FindName('Details').Text=[string]$machineState.LastInstallDetail
            $detailWindow.FindName('Close').Add_Click({$detailWindow.Close()}.GetNewClosure())
            [void]$detailWindow.ShowDialog()
        } catch {$status.Text='Installation details could not be opened: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('InstallPrerequisites').Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $assessmentState.Active) {throw 'Wait for the current operation to finish.'}
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-MachinePrerequisitesProgressDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=if ($null -ne $prereqUi.Owner) {$prereqUi.Owner} else {$window}
            $progressState=@{Active=$false;Refresh=$false;Worker=$null;Async=$null;Timer=$null;Phase='';Plan=$null;Clock=$null;Shared=[hashtable]::Synchronized(@{Current='';Completed=0;Total=0})}
            $dialog.Add_Closing({param($sender,$e) if ($progressState.Active) {$e.Cancel=$true}}.GetNewClosure())
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Install').Add_Click({
                if ($progressState.Active) {return}
                Write-SiAssistantLog -Log $assistantLog -Event Action -Step 'prerequisites' -Action 'ConfirmPrerequisites' -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count
                try {
                    if (-not $OfflineSmokeTest) {
                        $principal=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
                        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {throw 'Restart the assistant as administrator to install machine-wide prerequisites.'}
                    }
                    $progressState.Active=$true;$progressState.Phase='Certificate';$progressState.Clock=[Diagnostics.Stopwatch]::StartNew();$machineState.Ready=$false
                    $dialog.FindName('Install').IsEnabled=$false;$dialog.FindName('Cancel').IsEnabled=$false;$dialog.FindName('Trust').IsEnabled=$false
                    $dialog.FindName('ProgressText').Text='Installing the signing certificate...';$dialog.FindName('ProgressText').Visibility='Visible';$dialog.FindName('Progress').Visibility='Visible'
                    $status.Text='Installing the signing certificate...'
                    Write-SiAssistantLog -Log $assistantLog -Event Status -Step 'prerequisites' -StatusText 'Certificate installation running' -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count
                    $worker=[PowerShell]::Create();[void]$worker.AddCommand((Join-Path $assistantLibRoot 'SmartM365.PrerequisiteWorker.Progress.ps1')).AddParameter('Mode','Certificate').AddParameter('LibraryRoot',$assistantLibRoot).AddParameter('RepositoryRoot',(Split-Path -Parent $runtimeRoot)).AddParameter('Root',$root).AddParameter('TrustRepository',[bool]$dialog.FindName('Trust').IsChecked).AddParameter('OfflineFixture',[bool]$OfflineSmokeTest)
                    $progressState.Worker=$worker;$progressState.Async=$worker.BeginInvoke()
                    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(150);$progressState.Timer=$timer
                    $prereqProgressRef=$progressState;$prereqMachineRef=$machineState;$prereqDialogRef=$dialog;$prereqStatusRef=$status
                    $prereqOfflineRef=[bool]$OfflineSmokeTest;$prereqUpdateRef=$updateDetail;$prereqLogRef=$assistantLog;$prereqActionsRef=$pendingActions;$prereqEditsRef=$pendingEdits
                    $prereqInstallRef=$installState;$prereqLibRef=$assistantLibRoot;$prereqSmokeRef=$smokeUiState;$prereqWindowRef=$window
                    $timer.Add_Tick({
                        try {
                            $elapsed=$prereqProgressRef.Clock.Elapsed
                            $elapsedText='{0:00}:{1:00}:{2:00}' -f [int][Math]::Floor($elapsed.TotalHours),$elapsed.Minutes,$elapsed.Seconds
                            $phaseText=switch ($prereqProgressRef.Phase) {
                                'Certificate' {'Installing the signing certificate...'}
                                'Start' {'Preparing the module installer...'}
                                'Install' {'Installing Windows features and PowerShell modules...'}
                                'Audit' {
                                    $total=[int]$prereqProgressRef.Shared.Total;$done=[int]$prereqProgressRef.Shared.Completed;$current=[string]$prereqProgressRef.Shared.Current
                                    if ($total -gt 0 -and $current) {"Checking module: $current`nVerified: $done / $total"}
                                    elseif ($total -gt 0) {"Finalizing verification`nVerified: $done / $total"}
                                    else {'Checking the installation...'}
                                }
                                default {'Preparing...'}
                            }
                            $prereqDialogRef.FindName('ProgressText').Text=$phaseText+"`nElapsed: "+$elapsedText
                            if ($prereqProgressRef.Phase -in @('Certificate','Start','Audit')) {
                                if (-not $prereqProgressRef.Async.IsCompleted) {return}
                                $phase=$prereqProgressRef.Phase;$worker=$prereqProgressRef.Worker
                                try {$output=@($worker.EndInvoke($prereqProgressRef.Async));if ($worker.HadErrors -or $output.Count -ne 1) {throw 'The background operation did not return a complete result.'};$outcome=$output[0]}
                                finally {$worker.Dispose();$prereqProgressRef.Worker=$null;$prereqProgressRef.Async=$null}
                                if ($phase -eq 'Certificate') {
                                    $prereqProgressRef.Plan=$outcome.Plan
                                    $prereqDialogRef.FindName('ProgressText').Text='Preparing the module installer...';$prereqStatusRef.Text='Preparing the module installer...'
                                    $worker=[PowerShell]::Create();[void]$worker.AddCommand((Join-Path $prereqLibRef 'SmartM365.PrerequisiteWorker.Progress.ps1')).AddParameter('Mode','Start').AddParameter('LibraryRoot',$prereqLibRef).AddParameter('Plan',$prereqProgressRef.Plan).AddParameter('OfflineFixture',$prereqOfflineRef)
                                    $prereqProgressRef.Worker=$worker;$prereqProgressRef.Async=$worker.BeginInvoke();$prereqProgressRef.Phase='Start';return
                                }
                                if ($phase -eq 'Start') {
                                    $prereqInstallRef.Run=$outcome;$prereqProgressRef.Phase='Install'
                                    $prereqDialogRef.FindName('ProgressText').Text='Installing Windows features and PowerShell modules. This may take several minutes...'
                                    $prereqStatusRef.Text='Installing Windows features and PowerShell modules...'
                                    Write-SiAssistantLog -Log $prereqLogRef -Event Status -Step 'prerequisites' -StatusText 'Prerequisite installation running' -PendingActions $prereqActionsRef.Count -PendingEdits $prereqEditsRef.Count
                                    return
                                }
                                if ($prereqOfflineRef) {
                                    $prereqMachineRef.SimulatedInstalls++;$prereqMachineRef.ValidationFailed=$false;$prereqMachineRef.Ready=$true;$prereqMachineRef.Rows=@([pscustomobject]@{Name='Synthetic module';Present=$true;Detail='Installed for all users'})
                                    $prereqInstallRef.Run=$null;$prereqStatusRef.Text='Offline simulation: no installation occurred.';$prereqProgressRef.Timer.Stop();$prereqProgressRef.Active=$false;$prereqDialogRef.Close();& $prereqUpdateRef;return
                                }
                                $prereqMachineRef.ValidationFailed=($outcome.State -ne 'InstalledAuditPassed')
                                $prereqMachineRef.LastInstall=if ($outcome.State -eq 'InstalledAuditPassed') {'Module installation completed. Review the updated prerequisite status.'} elseif ($outcome.State -eq 'ImportValidationFailed') {'Modules were detected, but an import failed. Open installation details.'} else {'Installation or audit incomplete. Open installation details.'}
                                $prereqMachineRef.LastInstallDetail=if ($outcome.State -eq 'InstalledAuditPassed') {''} else {'Installer exit: '+$outcome.InstallExit+'; audit exit: '+$outcome.AuditExit+"`n`n"+$outcome.Detail}
                                $prereqWindowRef.FindName('ShowInstallDetails').Visibility=if ($prereqMachineRef.LastInstallDetail) {'Visible'} else {'Collapsed'}
                                $prereqInstallRef.Run=$null;$prereqProgressRef.Timer.Stop();$prereqProgressRef.Active=$false;$prereqProgressRef.Refresh=$true;$prereqDialogRef.Close();return
                            }
                            if ($prereqProgressRef.Phase -eq 'Install' -and $prereqInstallRef.Run.Process.HasExited) {
                                $prereqDialogRef.FindName('ProgressText').Text='Checking the installation...';$prereqStatusRef.Text='Checking the installation...'
                                $worker=[PowerShell]::Create();[void]$worker.AddCommand((Join-Path $prereqLibRef 'SmartM365.PrerequisiteWorker.Progress.ps1')).AddParameter('Mode','Complete').AddParameter('LibraryRoot',$prereqLibRef).AddParameter('Run',$prereqInstallRef.Run).AddParameter('Progress',$prereqProgressRef.Shared).AddParameter('OfflineFixture',$prereqOfflineRef)
                                $prereqProgressRef.Worker=$worker;$prereqProgressRef.Async=$worker.BeginInvoke();$prereqProgressRef.Phase='Audit'
                            }
                        } catch {
                            if ($null -ne $prereqProgressRef.Worker) {try {$prereqProgressRef.Worker.Dispose()} catch {};$prereqProgressRef.Worker=$null}
                            $prereqProgressRef.Timer.Stop();$prereqProgressRef.Active=$false;$prereqMachineRef.Ready=$false;$prereqMachineRef.ValidationFailed=$true;$prereqMachineRef.LastInstall='Installation incomplete or uncertain. Open installation details.';$prereqMachineRef.LastInstallDetail=$_.Exception.ToString();$prereqWindowRef.FindName('ShowInstallDetails').Visibility='Visible'
                            Write-SiAssistantLog -Log $prereqLogRef -Event Status -Step 'prerequisites' -StatusText 'Prerequisite installation failed or uncertain' -PendingActions $prereqActionsRef.Count -PendingEdits $prereqEditsRef.Count
                            $prereqInstallRef.Run=$null;$prereqProgressRef.Refresh=$true;$prereqDialogRef.Close()
                            if ($prereqOfflineRef) {$prereqSmokeRef.SaveError=$prereqMachineRef.LastInstall} else {[void][Windows.MessageBox]::Show($prereqWindowRef,$prereqMachineRef.LastInstall,'Installation failed','OK','Warning')}
                        }
                    }.GetNewClosure())
                    $timer.Start()
                } catch {
                    $progressState.Active=$false;$machineState.Ready=$false;$machineState.ValidationFailed=$true;$machineState.LastInstall='Installation incomplete or not started. Open installation details.';$machineState.LastInstallDetail=$_.Exception.ToString();$window.FindName('ShowInstallDetails').Visibility='Visible'
                    Write-SiAssistantLog -Log $assistantLog -Event Status -Step 'prerequisites' -StatusText 'Prerequisite installation failed' -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count
                    $dialog.Close();if ($OfflineSmokeTest) {$smokeUiState.SaveError=$machineState.LastInstall} else {[void][Windows.MessageBox]::Show($window,$machineState.LastInstall,'Installation failed','OK','Warning')};$progressState.Refresh=$true
                }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
            if ($progressState.Refresh -and -not $OfflineSmokeTest) {
                $refreshTarget=$window
                [void]$window.Dispatcher.BeginInvoke([Action]{$refreshTarget.FindName('RunLocalBatch').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}.GetNewClosure(),[Windows.Threading.DispatcherPriority]::Background)
            }
        } catch {$status.Text='Installation unavailable: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('StopPermissions').Add_Click({
        try {
            if ($null -eq $permissionState.Run) {return}
            if ([Windows.MessageBox]::Show($window,'Stop the read-only check? Authentication or consent may already have occurred. The private temporary folder will be retained for review.','Stop permission check','YesNo','Warning') -ne 'Yes') {return}
            Stop-SiPermissionCheck -Run $permissionState.Run; $window.FindName('StopPermissions').IsEnabled=$false; $status.Text='Stop requested; wait for the result.'
        } catch { $status.Text='Stop could not be confirmed; inspect the child process.' }
    }.GetNewClosure())
    $window.FindName('CheckPermissions').Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) {throw 'Another operation is already running.'}
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-PermissionCheckDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=$window
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Browse').Add_Click({$picker=[Microsoft.Win32.OpenFileDialog]::new();$picker.Title='Choose the catalog to compare';$picker.Filter='JSON files (*.json;*.json.txt)|*.json;*.json.txt';$picker.CheckFileExists=$true;if ($picker.ShowDialog($dialog) -eq $true) {$dialog.FindName('Catalog').Text=$picker.FileName}}.GetNewClosure())
            $dialog.FindName('Check').Add_Click({
                try {
                    $plan=New-SiPermissionCheckPlan -Root $root -Tenant $tenantBox.Text -CatalogPath $dialog.FindName('Catalog').Text -Outside $dialog.FindName('Outside').Text.Trim() -ExpectedScope $dialog.FindName('Scope').Text.Trim()
                    Add-SiPendingCreation -Queue $pendingActions -Action (New-SiPendingTestAction -Kind PermissionTest -Data $plan)
                    $tenantBox.IsEnabled=$false;$status.Text='Permission check added to final review; no connection was made.';$dialog.Close()
                } catch {[void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Permission check unavailable','OK','Warning')}
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch {$status.Text='Dialog unavailable: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('PlanBootstrap').Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run) { throw 'Another operation is already running.' }
            $stageId=[string]$steps.SelectedItem.Id
            $newProfileActions=@($pendingActions | Where-Object {$_.Kind -eq 'Profile' -and $_.Data.ProfileKey -ceq $tenantBox.Text})
            $pendingProfile=if ($newProfileActions.Count -eq 1) {$newProfileActions[0].Data} else {$null}
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-BootstrapAutoTenantDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $currentProfile=Read-SiPrivateJson (Join-Path $root ("Config/Tenants/$($tenantBox.Text).local"))
            if ($currentProfile.State -eq 'Present' -and $currentProfile.Value.EnableSharePointUpload -is [bool]) {$dialog.FindName('ConfigureTeams').IsChecked=[bool]$currentProfile.Value.EnableSharePointUpload}
            elseif ($null -ne $pendingProfile) {
                $spEdit=@($pendingProfile.PendingEdits | Where-Object {$_.Field -eq 'EnableSharePointUpload'} | Select-Object -Last 1)
                $dialog.FindName('ConfigureTeams').IsChecked=if ($spEdit.Count) {[bool](ConvertTo-SiEditValue -Snapshot $spEdit[0] -Value $spEdit[0].Value)} else {[bool]$pendingProfile.Model.EnableSharePointUpload}
            }
            $detectState=@{Tenant='';Account='';PreflightReady=$false;PreflightMessage='';Process=$null;Timer=$null;Path=''}
            if ($tenantSession.Profile -ceq $tenantBox.Text) {$detectState.Tenant=$tenantSession.Tenant;$detectState.Account=$tenantSession.Account;$detectState.PreflightReady=$tenantSession.PreflightReady}
            if ($detectState.Tenant) {$dialog.FindName('TenantStatus').Text="Tenant ID: $($detectState.Tenant)`nSigned in as: $($detectState.Account)`nBootstrap may require a new sign-in."}
            $bootstrapChanged={ $dialog.FindName('Stage').IsEnabled=[bool]$detectState.Tenant; if ($detectState.Tenant) {$tenantSession.Tenant=$detectState.Tenant;$tenantSession.Account=$detectState.Account;$tenantSession.PreflightReady=$detectState.PreflightReady;$tenantSession.Profile=$tenantBox.Text} elseif ($tenantSession.Profile -ceq $tenantBox.Text) {$tenantSession.Tenant='';$tenantSession.Account='';$tenantSession.Profile=''} }.GetNewClosure()
            $bootstrapInstall={ $prereqUi.Owner=$dialog;try {$window.FindName('InstallPrerequisites').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$prereqUi.Owner=$null} }.GetNewClosure()
            Initialize-SiTenantSignInDialog -Dialog $dialog -State $detectState -LibraryRoot $assistantLibRoot -Changed $bootstrapChanged -InstallPrerequisites $bootstrapInstall -OfflineFixture ([bool]$OfflineSmokeTest)
            & $bootstrapChanged
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Stage').Add_Click({
                try {
                    if (-not $detectState.Tenant) {throw 'Connect and confirm the Microsoft 365 tenant first.'}
                    $review=New-SiBootstrapPlan -Root $root -ProfileKey $tenantBox.Text -Directory $detectState.Tenant -DisplayName $dialog.FindName('DisplayName').Text -AutoSelectOutsideAddress $true -UpdateExisting ([bool]$dialog.FindName('UpdateExisting').IsChecked) -ConfigureTeams ([bool]$dialog.FindName('ConfigureTeams').IsChecked) -PendingProfile $pendingProfile
                    $stagePlan=New-SiBootstrapStagePlan -Root $root -BootstrapPlan $review
                    $picker=[Microsoft.Win32.OpenFolderDialog]::new(); $picker.Title='Choose a local folder for isolated preparation'
                    if ($picker.ShowDialog($dialog) -ne $true) { return }
                    $intent=New-SiPendingBootstrapIntent -Plan $stagePlan -Parent $picker.FolderName
                    Add-SiPendingCreation -Queue $pendingActions -Action $intent
                    $tenantBox.IsEnabled=$false
                    $status.Text='Bootstrap added to final review; no folder or tenant resource was created.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Preparation failed','OK','Warning') }
            }.GetNewClosure())
            $dialog.FindName('Prepare').Add_Click({
                try {
                    $plan=New-SiBootstrapPlan -Root $root -ProfileKey $tenantBox.Text -Directory $detectState.Tenant -DisplayName $dialog.FindName('DisplayName').Text -AutoSelectOutsideAddress $true -UpdateExisting ([bool]$dialog.FindName('UpdateExisting').IsChecked) -ConfigureTeams ([bool]$dialog.FindName('ConfigureTeams').IsChecked) -PendingProfile $pendingProfile
                    $detail="Profile: $($plan.ProfileKey)`nRequested tenant: $($plan.Directory)`nApplication : $($plan.DisplayName)`nExisting application update: $($plan.UpdateExisting)`nTeams / SharePoint : $($plan.ConfigureTeams)`nNonempty local fields bootstrap may replace (including placeholders): $($plan.PresentFields -join ', ')`nScript hash: $($plan.SetupSHA256)`nProfile hash: $($plan.ProfileSHA256)`n`n$($plan.Effects)`n`nLocal review only: effective settings, identities and tenant permissions are not verified. Execution requires isolated preparation and separate confirmation."
                    $observations[$stageId]=@{State='Plan requires review; not run';Detail=$detail}
                    $window.FindName('StepState').Text=$observations[$stageId].State
                    $window.FindName('Action').Text=$detail
                    $status.Text='Plan prepared without connection or changes.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Plan failed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('OpenOrchestratorConsole').Add_Click({
        try {
            if ($null -ne $consoleState.Run -or $null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run) { throw 'Another operation is already running.' }
            $plan=New-SiOrchestratorConsolePlan -Root $root -Tenant $tenantBox.Text
            Add-SiPendingCreation -Queue $pendingActions -Action (New-SiPendingTestAction -Kind Console -Data $plan)
            $tenantBox.IsEnabled=$false;& $updateDetail
            $status.Text='Management console added to final review; it has not been opened.'
        } catch { $status.Text='Console not opened: '+$_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('EditOrchestratorSettings').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-OrchestratorSettingsDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $localState=@{Snapshot=$null}; $dialog.FindName('Field').ItemsSource=@(Get-SiOrchestratorSettingNames); $dialog.FindName('Field').SelectedIndex=0
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Field').Add_SelectionChanged({$localState.Snapshot=$null; $dialog.FindName('Current').Text=''; $dialog.FindName('Value').Text=''}.GetNewClosure())
            $dialog.FindName('Browse').Add_Click({
                $picker=[Microsoft.Win32.OpenFileDialog]::new(); $picker.Title='Select the orchestrator local configuration'; $picker.Filter='JSON files (*.json;*.json.txt)|*.json;*.json.txt'; $picker.CheckFileExists=$true
                if ($picker.ShowDialog($dialog) -eq $true) { $dialog.FindName('Path').Text=$picker.FileName; $localState.Snapshot=$null; $dialog.FindName('Current').Text=''; $dialog.FindName('Value').Text='' }
            }.GetNewClosure())
            $dialog.FindName('Load').Add_Click({
                try {
                    $template=Join-Path $runtimeRoot 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.local.json.template'
                    $localState.Snapshot=Get-SiOrchestratorSettingSnapshot -Path $dialog.FindName('Path').Text -Field ([string]$dialog.FindName('Field').SelectedItem) -TemplatePath $template
                    $dialog.FindName('Current').Text=if ($localState.Snapshot.Present) { $localState.Snapshot.Value } else { '(missing; template default) '+$localState.Snapshot.Value }
                    $dialog.FindName('Value').Text=$localState.Snapshot.Value
                }
                catch { $localState.Snapshot=$null; [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Read failed','OK','Warning') }
            }.GetNewClosure())
            $dialog.FindName('Save').Add_Click({
                try {
                    $snapshot=$localState.Snapshot
                    if ($null -eq $snapshot -or $snapshot.Path -ine $dialog.FindName('Path').Text -or $snapshot.Field -cne [string]$dialog.FindName('Field').SelectedItem) { throw 'Load the setting again before saving.' }
                    $value=$dialog.FindName('Value').Text
                    if ($snapshot.Present -and $value -ceq $snapshot.Value) { return }
                    Add-SiPendingCreation -Queue $pendingActions -Action (New-SiPendingFileAction -Type Orchestrator -Plan $snapshot -Value $value)
                    $localState.Snapshot=$null;$tenantBox.IsEnabled=$false;$dialog.Close();& $updateDetail
                    $status.Text='Orchestrator setting added to final review; no file was written.'
                } catch { $localState.Snapshot=$null; [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Save failed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: '+$_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('PlanTask').Add_Click({
        try {
            $pendingAccount=@($pendingActions | Where-Object {$_.Kind -in @('LocalAccount','DomainAccount') -and $_.State -eq 'Pending'})
            if ($pendingAccount.Count -eq 1) {$accountLabel='Pending account: '+$pendingAccount[0].Data.Name}
            elseif ($observations.ContainsKey('service') -and $observations['service'].Sid -and -not $observations['service'].Managed) {$accountLabel=[string]$observations['service'].Account}
            else {throw 'Select or prepare a standard service account first.'}
            $candidates=@(foreach ($base in @($env:ProgramFiles,[Environment]::GetEnvironmentVariable('ProgramFiles(x86)'))) { if ($base) { Join-Path $base 'PowerShell/7/pwsh.exe' } })
            $executable=@($candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)
            if ($executable.Count -eq 0) { throw 'PowerShell 7 was not found in the expected locations.' }
            $scriptPath=Join-Path $root 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
            if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw 'Orchestrator script not found in this copy.' }
            $choice=[Windows.MessageBox]::Show($window,'Enable the new task after it is created and verified? It may run according to its schedule. No immediate start will be requested. Choose No to leave it disabled.','Task activation','YesNoCancel','Question','No')
            if ($choice -eq 'Cancel') {return}
            $intent=New-SiPendingTaskIntent -Root $root -Tenant $tenantBox.Text -Executable $executable[0] -AccountLabel $accountLabel -EnableAfterCreation ($choice -eq 'Yes')
            Add-SiPendingCreation -Queue $pendingActions -Action $intent;$tenantBox.IsEnabled=$false
            $detail="Account: $accountLabel`nInstaller SHA-256: $($intent.Data.InstallerHash)`nEnable after creation: $($intent.Data.EnableAfterCreation)`nCreate only, without replacement or an immediate start. The installer will request the account password interactively after final confirmation."
            $observations['orchestrator']=@{State='Plan prepared; not run';Detail=$detail}
            $window.FindName('StepState').Text=$observations['orchestrator'].State
            $window.FindName('Action').Text=$detail
            $status.Text='Task added to final review; no task was created.'
        } catch { $status.Text='Plan failed : ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckServiceRights').Add_Click({
        try {
            if (-not $observations.ContainsKey('service') -or -not $observations['service'].Sid) { throw 'Resolve the service account first.' }
            $account=$observations['service'].Account; $sid=$observations['service'].Sid; $managed=[bool]$observations['service'].Managed
            $result=Get-SiDirectAccountRights -Sid $sid
            $label=if ($result.DirectDeny) { 'Direct deny detected' } elseif ($result.DirectGrant) { 'Direct grant observed' } else { 'No direct grant observed' }
            $observations['service']=@{State=$label;Detail=('Account: '+$account+[Environment]::NewLine+$result.Detail);Sid=$sid;Account=$account;Managed=$managed}
            $window.FindName('StepState').Text=$label
            $window.FindName('Action').Text=$observations['service'].Detail
            $status.Text='Local read only; group rights and effective access not verified.'
        } catch { $status.Text='Rights check failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $serviceShareButton.Add_Click({
        try {
            $identity=$serviceState.Identity
            if (-not (Get-SiServiceReadiness -Identity $identity -PendingActions $pendingActions).Ready -or -not $shareState.Ready) {throw 'Select a service account and an accessible share first.'}
            $target=ConvertTo-SiServiceShareTarget -Path $shareState.Path
            $sourceTarget=ConvertTo-SiServiceShareTarget -Path $(if ($OfflineSmokeTest) {'\\synthetic.invalid\source$\SmartM365'} else {$root})
            [xml]$x=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ServiceAccessExchangeDialog.xaml') -Raw
            $d=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($x));$d.Owner=$window
            $d.FindName('Identity').Text='Account: '+$identity.Account;$d.FindName('Share').Text='CSV and log share: '+$target.Path
            $op=@{Active=$false;Worker=$null;Async=$null;Timer=$null;Mode='';Credential=$null;Password=$null;Start=$null;GrantAttempted=$false;ExchangeSnapshot=$null;ExchangeAdded=$false}
            $d.Add_Closing({param($sender,$e) if ($op.Active) {$e.Cancel=$true}}.GetNewClosure())
            $d.Add_Closed({$d.FindName('Password').Clear();if ($null -ne $op.Password) {$op.Password.Dispose();$op.Password=$null};$op.Credential=$null}.GetNewClosure())
            $d.FindName('Cancel').Add_Click({$d.Close()}.GetNewClosure())
            $op.Start={param($mode)
                $op.Active=$true;$op.Mode=$mode;$d.FindName('Check').IsEnabled=$false;$d.FindName('Cancel').IsEnabled=$false;$d.FindName('Password').IsEnabled=$false;$d.FindName('Progress').Visibility='Visible'
                $d.FindName('ProgressText').Text=switch ($mode) {'Grant' {'Applying share and folder permissions...'} 'ExchangeCheck' {'Checking Exchange role group...'} 'ExchangeAdd' {'Adding the account to the Exchange role group...'} default {'Checking access as the service account...'}}
                $worker=[PowerShell]::Create()
                if ($mode -in @('ExchangeCheck','ExchangeAdd')) {
                    [void]$worker.AddCommand((Join-Path $assistantLibRoot 'SmartM365.ExchangeRoleWorker.Ou.ps1')).AddParameter('Mode',$(if ($mode -eq 'ExchangeCheck') {'Check'} else {'Add'})).AddParameter('AccountSid',[string]$identity.Sid).AddParameter('LibraryRoot',$assistantLibRoot).AddParameter('Snapshot',$op.ExchangeSnapshot).AddParameter('OfflineFixture',[bool]$OfflineSmokeTest).AddParameter('FixtureMember',[bool]$op.ExchangeAdded)
                } else {
                    [void]$worker.AddCommand((Join-Path $assistantLibRoot 'SmartM365.ServiceShareWorker.Network.ps1')).AddParameter('Mode',$mode).AddParameter('Target',$target).AddParameter('SourceTarget',$sourceTarget).AddParameter('Sid',[string]$identity.Sid).AddParameter('Credential',$op.Credential).AddParameter('LibraryRoot',$assistantLibRoot).AddParameter('OfflineFixture',[bool]$OfflineSmokeTest).AddParameter('FixtureState',$(if ($OfflineSmokeTest -and $smokeUiState.ShareProbeDenied -and -not $op.GrantAttempted) {'Denied'} else {'Verified'}))
                }
                $op.Worker=$worker;$op.Async=$worker.BeginInvoke()
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(150);$op.Timer=$timer
                $shareOpRef=$op;$shareDialogRef=$d;$shareIdentityRef=$identity;$shareTargetRef=$target;$shareSourceRef=$sourceTarget;$shareResultRef=$serviceShareState;$shareUpdateRef=$updateDetail;$shareStatusRef=$status;$shareOfflineRef=[bool]$OfflineSmokeTest;$roleResultRef=$exchangeRoleState;$roleEnabledRef=[bool]$scope.Exchange
                $timer.Add_Tick({
                    if (-not $shareOpRef.Async.IsCompleted) {return}
                    $shareOpRef.Timer.Stop();$next=''
                    try {
                        $worker=$shareOpRef.Worker;try {try {$output=@($worker.EndInvoke($shareOpRef.Async))} catch {$cause=$_.Exception;while ($null -ne $cause.InnerException) {$cause=$cause.InnerException};throw $cause.Message};if ($worker.HadErrors -or $output.Count -ne 1) {throw 'The operation did not complete. Permissions or membership may have changed; review the file server or directory before retrying.'};$result=$output[0]} finally {$worker.Dispose();$shareOpRef.Worker=$null}
                        if ($shareOpRef.Mode -eq 'Grant') {$shareOpRef.GrantAttempted=$true;$next='Check'}
                        elseif ($shareOpRef.Mode -eq 'ExchangeAdd') {if (-not $result.Member) {throw 'Exchange membership was not confirmed.'};$shareOpRef.ExchangeAdded=$true;$next='ExchangeCheck'}
                        elseif ($shareOpRef.Mode -eq 'ExchangeCheck') {
                            if ($result.Member) {$roleResultRef.Ready=$true;$roleResultRef.Sid=[string]$shareIdentityRef.Sid;$shareOpRef.Active=$false;$shareDialogRef.Close();& $shareUpdateRef;$shareStatusRef.Text='Service account share access and Exchange role group verified.'}
                            else {$shareOpRef.ExchangeSnapshot=$result;$message='Add '+$shareIdentityRef.Account+' to the Exchange role group View-Only Organization Management?';if ($shareOfflineRef -or [Windows.MessageBox]::Show($shareDialogRef,$message,'Confirm Exchange role group','YesNo','Question') -eq 'Yes') {$next='ExchangeAdd'} else {$shareDialogRef.FindName('ProgressText').Text='Exchange role group membership is required.'}}
                        }
                        elseif ($result.State -eq 'Verified') {
                            if (-not $shareOpRef.GrantAttempted) {$next='Grant'} else {
                                $shareResultRef.Ready=$true;$shareResultRef.Sid=[string]$shareIdentityRef.Sid;$shareResultRef.Path=$shareTargetRef.Path;$shareResultRef.State='Verified'
                                if ($roleEnabledRef) {$next='ExchangeCheck'} else {$shareOpRef.Active=$false;$shareDialogRef.Close();& $shareUpdateRef;$shareStatusRef.Text='Service account share access verified.'}
                            }
                        } elseif ($result.State -eq 'Denied' -and -not $shareOpRef.GrantAttempted) {
                            $next='Grant'
                        } else {$shareDialogRef.FindName('ProgressText').Text=[string]$result.Detail}
                    } catch {if ($shareOpRef.Mode -like 'Exchange*') {$shareDialogRef.FindName('ProgressText').Text='Exchange role group check failed: '+$_.Exception.Message;$roleResultRef.Ready=$false} else {$shareDialogRef.FindName('ProgressText').Text='Service account share check failed: '+$_.Exception.Message;$shareResultRef.Ready=$false}}
                    if ($next -eq 'Grant') {
                        try {
                            $review=Get-SiServiceShareGrantReview -SourceTarget $shareSourceRef -DataTarget $shareTargetRef -Account $shareIdentityRef.Account
                            $accepted=if ($shareOfflineRef) {$true} else {[Windows.MessageBox]::Show($shareDialogRef,$review.Message,'Confirm share permissions','YesNo','Warning') -eq 'Yes'}
                            if (-not $accepted) {$next='';$shareDialogRef.FindName('ProgressText').Text='Share permissions were not changed.'}
                        } catch {$next='';$shareDialogRef.FindName('ProgressText').Text='Share permission review failed: '+$_.Exception.Message}
                    }
                    if ($next) {& $shareOpRef.Start $next;return}
                    $shareOpRef.Active=$false;if ($shareDialogRef.IsVisible) {$shareDialogRef.FindName('Progress').Visibility='Collapsed';$shareDialogRef.FindName('Check').IsEnabled=$true;$shareDialogRef.FindName('Cancel').IsEnabled=$true;$shareDialogRef.FindName('Password').IsEnabled=$true}
                }.GetNewClosure())
                $timer.Start()
            }.GetNewClosure()
            $d.FindName('Check').Add_Click({
                try {
                    if ($op.Active) {return}
                    $entered=$d.FindName('Password').SecurePassword
                    try {if ($entered.Length -eq 0) {throw 'Enter the service account password.'};if ($null -ne $op.Password) {$op.Password.Dispose()};$op.Password=$entered.Copy()} finally {$entered.Dispose()}
                    $op.Credential=[pscredential]::new([string]$identity.Account,$op.Password);$op.GrantAttempted=$false;$op.ExchangeAdded=$false;$op.ExchangeSnapshot=$null;$serviceShareState.Ready=$false;if ($scope.Exchange) {$exchangeRoleState.Ready=$false}
                    & $op.Start 'Check'
                } catch {$d.FindName('ProgressText').Text=$_.Exception.Message}
            }.GetNewClosure())
            [void]$d.ShowDialog()
        } catch {$status.Text='Service account share access unavailable: '+$_.Exception.Message;& $updateDetail}
    }.GetNewClosure())
    $window.FindName('ConfigureService').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ServiceChoiceNowDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=$window
            $selection=@{Control=''}
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Continue').Add_Click({
                $selection.Control=if ($dialog.FindName('Domain').IsChecked) {'CreateDirectoryService'} else {'CheckService'}
                $dialog.Close()
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
            if ($selection.Control) {$window.FindName($selection.Control).RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
            if ($selection.Control -and (Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions).Ready) {
                Set-SiPendingServiceAccount -Root $root -Tenant $tenantBox.Text -Account ([string]$serviceState.Identity.Account) -Actions $pendingActions -Edits $pendingEdits
                if (-not $OfflineSmokeTest) {$serviceShareButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
                & $updateDetail
            }
        } catch {$status.Text='Service account selection unavailable: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('CheckService').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ServiceIdentityDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Check').Add_Click({
                try {
                    $accountName=$dialog.FindName('Account').Text.Trim()
                    if ($accountName.EndsWith('$')) {throw 'This version supports standard service accounts only.'}
                    if (@($pendingActions | Where-Object Kind -eq 'DomainAccount').Count) {throw 'Discard the pending account creation before selecting an existing account.'}
                    $result=if ($OfflineSmokeTest) {Resolve-SiServiceAccount -Account $accountName -Managed:$false -Resolver {param($name) 'S-1-5-21-1-2-3-1001'}} else {Resolve-SiServiceAccount -Account $accountName -Managed:$false}
                    $serviceState.Identity=$result;$serviceState.Created=$false;if (-not $OfflineSmokeTest) {$serviceShareState.Ready=$false;$exchangeRoleState.Ready=$false}
                    $taskPlanState.Plan=$null
                    Set-SiServiceIdentityObservation -Observations $observations -Identity $result
                    $window.FindName('CheckServiceRights').IsEnabled=$true
                    $window.FindName('StepState').Text=$observations['service'].State
                    $window.FindName('Action').Text=$observations['service'].Detail
                    $status.Text='Identity resolved. Permissions still need verification.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Resolution failed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $identityButton.Add_Click({
        try {
            if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) {throw 'Another operation is already running.'}
            $snapshot=Get-SiTenantIdentitySnapshot -Root $root -Tenant $tenantBox.Text
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-TenantIdentityDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=$window
            $dialog.FindName('Current').Text="Current ProfileKey: $($snapshot.Current.ProfileKey)`nCurrent TenantKey: $($snapshot.Current.TenantKey)"
            $dialog.FindName('Organization').Text=$snapshot.Current.OrganizationKey;$dialog.FindName('Environment').Text=$snapshot.Current.EnvironmentKey
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Save').Add_Click({
                try {
                    $plan=New-SiTenantIdentityEditPlan -Snapshot $snapshot -OrganizationKey $dialog.FindName('Organization').Text.Trim() -EnvironmentKey $dialog.FindName('Environment').Text.Trim()
                    Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingIdentityEdit -Plan $plan)
                    $tenantBox.IsEnabled=$false;$dialog.Close();& $updateDetail
                    $status.Text='Tenant identity added to final review; existing data will not be moved.'
                } catch {[void][Windows.MessageBox]::Show($dialog,($_.Exception.Message+"`n`nIf saving began, review the local profile before retrying."),'Identity change unconfirmed','OK','Warning')}
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch {$status.Text='Tenant identity review unavailable: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('SetScope').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ScopeDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            foreach ($name in @('Cloud','AD','Exchange','SharePoint','Teams','Mail')) { $dialog.FindName($name).IsChecked=[bool]$scope[$name] };$dialog.FindName('Teams').Visibility='Collapsed';$dialog.FindName('Teams').IsChecked=$false;$dialog.FindName('Cloud').Visibility='Collapsed';$dialog.FindName('Cloud').IsChecked=$true
            $dialog.FindName('Cancel').Add_Click({ $dialog.Close() }.GetNewClosure())
            $dialog.FindName('Apply').Add_Click({
                foreach ($name in @('Cloud','AD','Exchange','SharePoint','Teams','Mail')) { $scope[$name]=[bool]$dialog.FindName($name).IsChecked }
                $scope.Set=$true; $dialog.Close(); & $refresh; & $runPassiveChecks
                $status.Text='Scope retained for this session; no job changed.'
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('SetUploadCategory').Add_Click({
        try {
            $snapshots=@{}
            foreach ($category in @('DataAll','DataLast','DataPowerBI','LogAll')) { $snapshots[$category]=Get-SiUploadCategorySnapshot -Root $root -Tenant $tenantBox.Text -Category $category }
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-SharePointCategoriesDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $updateCategory={
                $category=[string]$dialog.FindName('Category').SelectedItem.Content
            $source=if ($snapshots[$category].Present) {'Current local value'} else {'Template default; not yet saved'}
            $dialog.FindName('Current').Text=$source+': ' + $snapshots[$category].Choice
            }.GetNewClosure()
            $dialog.FindName('Category').Add_SelectionChanged($updateCategory)
            & $updateCategory
            $dialog.FindName('Cancel').Add_Click({ $dialog.Close() }.GetNewClosure())
            $dialog.FindName('Save').Add_Click({
                try {
                    $category=[string]$dialog.FindName('Category').SelectedItem.Content
                    $choice=[string]$dialog.FindName('Choice').SelectedItem.Content
                    $snapshot=$snapshots[$category]
                    $current=Get-SiEditSnapshot -Root $root -Tenant $tenantBox.Text -Step sharepoint -Field $snapshot.Field
                    if ($current.Path -cne $snapshot.Path -or $current.Hash -cne $snapshot.Hash) {throw 'Configuration changed; reopen the category dialog.'}
                    $value=switch ($choice) {'Enabled' {'true'} 'Disabled' {'false'} default {'__USE_GLOBAL__'}}
                    Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingEdit -Snapshot $current -Value $value)
                    $tenantBox.IsEnabled=$false;$dialog.Close()
                    $status.Text='SharePoint category added to final review; no configuration was written.'
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Save refused','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('SetEnabled').Add_Click({
        try {
            $step=[string]$steps.SelectedItem.Id
            $snapshot=Get-SiBooleanSnapshot -Root $root -Tenant $tenantBox.Text -Step $step
            $current=Get-SiEditSnapshot -Root $root -Tenant $tenantBox.Text -Step $step -Field $snapshot.Field
            if ($current.Path -cne $snapshot.Path -or $current.Hash -cne $snapshot.Hash) {throw 'Configuration changed; retry the edit.'}
            $previous=@($pendingEdits | Where-Object {$_.Snapshot.Path -ieq $current.Path -and $_.Snapshot.Field -ieq $current.Field} | Select-Object -Last 1)
            $currentValue=if ($previous.Count) {$previous[0].Value -ieq 'true'} else {[bool]$snapshot.Value}
            $next=if ($currentValue) {'false'} else {'true'}
            if ($current.Present -and $next -ieq $current.Value) {
                foreach ($old in $previous) {[void]$pendingEdits.Remove($old)}
                if (-not $pendingEdits.Count -and -not $pendingActions.Count) {$tenantBox.IsEnabled=$true}
                $status.Text='Pending setting removed; no configuration was written.';return
            }
            Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingEdit -Snapshot $current -Value $next)
            $tenantBox.IsEnabled=$false
            $status.Text='Setting added to final review; no configuration was written.'
        } catch { $status.Text='Change refused: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckCatalog').Add_Click({
        try {
            $picker=[Microsoft.Win32.OpenFileDialog]::new()
            $picker.Title='Select a catalog to analyze'
            $picker.Filter='JSON files (*.json;*.json.txt)|*.json;*.json.txt|All files (*.*)|*.*'
            $picker.CheckFileExists=$true
            if ($picker.ShowDialog($window) -ne $true) { return }
            $result=Get-SiCatalogObservation -CatalogPath $picker.FileName -InventoryRoot (Join-Path $root 'SmartInventory')
            $label=if ($result.Issues.Count -gt 0) { 'Issues detected in the selected catalog' } else { 'Partial local checks completed' }
            $detail="Catalog: $($picker.FileName)`nEnabled jobs: $($result.EnabledCount)`nExternal or unchecked scripts: $($result.UncheckedScripts)`nDisabled or manual dependencies ignored by the scheduler: $($result.IgnoredDependencies)`n"
            if ($result.Issues.Count -gt 0) { $detail += [Environment]::NewLine + ($result.Issues -join [Environment]::NewLine) + [Environment]::NewLine }
            $detail += 'Check limited to this catalog and local copy. Effective catalog, access, configuration and servers remain to be qualified.'
            $observations['collectors']=@{State=$label;Detail=$detail}
            $window.FindName('StepState').Text=$label
            $window.FindName('Action').Text=$detail
            $status.Text='Read-only analysis completed; no collection started.'
        } catch { $status.Text='Analysis failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckCertificate').Add_Click({
        try {
            $result=Get-SiCertificateObservation -Root $root -Tenant $tenantBox.Text
            $label = switch ($result.State) { 'Observed' { 'Certificate found locally' } 'NotObserved' { 'Certificate not found' } 'Multiple' { 'Multiple certificates' } 'OutsideValidity' { 'Invalid certificate dates' } default { 'No private key associated' } }
            $observations['identity'] = @{State=$label; Detail=$result.Detail}
            $window.FindName('StepState').Text = $label
            $window.FindName('Action').Text=$result.Detail
            $status.Text='Local observation only; access and permissions remain to be checked.'
        } catch { $status.Text='Inspection failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckTask').Add_Click({
        try {
            $result=Get-SiOrchestratorTaskObservation -Tenant $tenantBox.Text -SmartM365Root $root
            $label = switch ($result.State) { 'Observed' { 'Task found' } 'ActionMismatch' { 'Action requires review' } 'Conflict' { 'Conflict requires review' } default { 'Standard task not found' } }
            $observations['orchestrator'] = @{State=$label; Detail=$result.Detail}
            $window.FindName('StepState').Text = $label
            $window.FindName('Action').Text=$result.Detail
            $status.Text='Local observation only; no task started or changed.'
        } catch { $status.Text='Detection failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckShareAsService').Add_Click({
        try {
            if (-not $observations.ContainsKey('service') -or -not $observations['service'].Sid) { throw 'Resolve the service account first.' }
            $selectedSid=[string]$observations['service'].Sid; $selectedAccount=[string]$observations['service'].Account; $selectedManaged=[bool]$observations['service'].Managed
            if ($selectedManaged) {throw 'The selected managed account is not supported by this share check.'}
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-ServiceShareDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml)); $dialog.Owner=$window
            $dialog.FindName('Identity').Text='Selected account: '+$selectedAccount+'  |  SID: '+$selectedSid
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('ReviewAcceptance').Add_Click({
                try {
                    $picker=[Microsoft.Win32.OpenFileDialog]::new();$picker.Title='Select a local acceptance result';$picker.Filter='JSON files (*.json;*.json.txt)|*.json;*.json.txt|All files (*.*)|*.*'
                    if ($picker.ShowDialog($dialog) -ne $true) {return}
                    $result=Get-SiAcceptanceShareEvidence -Path $picker.FileName -ExpectedServer $env:COMPUTERNAME -ExpectedAccount $selectedAccount
                    $observations['share']=@{State=$result.State;Detail=$result.Detail}
                    $window.FindName('StepState').Text=$result.State;$window.FindName('Action').Text=$result.Detail;$status.Text='Local historical evidence reviewed; access is not qualified.'
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,'Local acceptance evidence could not be reviewed. Select a valid local file.','Evidence unavailable','OK','Warning') }
            }.GetNewClosure())
            $dialog.FindName('Check').Add_Click({
                try {
                    $path=$dialog.FindName('Path').Text.Trim()
                    if (-not [IO.Path]::IsPathFullyQualified($path)) {throw 'Enter an absolute local or UNC path.'}
                    $message="Check read access to $path only if the child process runs under SID $selectedSid? The share may be contacted. No file is created."
                    if ([Windows.MessageBox]::Show($dialog,$message,'Confirm read check','YesNo','Question') -ne 'Yes') {return}
                    $result=Test-SiSharedPathRead -Path $path -ExpectedSid $selectedSid
                    $observations['share']=@{State=$result.State;Detail=('Path: '+$path+[Environment]::NewLine+$result.Detail)}
                    $window.FindName('StepState').Text=$result.State; $window.FindName('Action').Text=$observations['share'].Detail; $status.Text=$result.Detail
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Read check failed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Dialog unavailable: '+$_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CheckShare').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-SharedPathDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml))
            $dialog.Owner=$window
            $dialog.FindName('Cancel').Add_Click({ $dialog.Close() }.GetNewClosure())
            $dialog.FindName('Check').Add_Click({
                try {
                    $checkedPath=$dialog.FindName('Path').Text
                    $result=Test-SiSharedPathRead -Path $checkedPath
                    $label=if ($result.State -eq 'Readable') { 'Read as current account' } else { 'Read access unverified' }
                    $observations['share'] = @{State=$label; Detail=('Path: ' + $checkedPath + [Environment]::NewLine + $result.Detail)}
                    $window.FindName('StepState').Text = $label
                    $window.FindName('Action').Text=$observations['share'].Detail
                    $status.Text=$result.Detail
                    $dialog.Close()
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Check failed','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Cannot open dialog: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CreateDirectoryService').Add_Click({
        try {
            if (@($pendingActions | Where-Object Kind -eq 'DomainAccount').Count) {throw 'A previously prepared account is pending. Discard it before creating a new account.'}
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-CreateDomainAccountNowDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml));$dialog.Owner=$window
            $accountProgress=@{Active=$false;Worker=$null;Async=$null;Timer=$null;Action=$null}
            $dialog.FindName('AccountName').Text='svc-smartinv'
            if (-not $OfflineSmokeTest) {
                try {
                    $domain=[System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
                    $controller=$domain.FindDomainController().Name;$dialog.FindName('DomainController').Text=$controller
                    if (Get-Command Get-ADDomain -ErrorAction SilentlyContinue) {$directory=Get-ADDomain -Server $controller -ErrorAction Stop;$dialog.FindName('OrganizationalUnit').Text=[string]$directory.UsersContainer}
                    else {$dn=[string]$domain.GetDirectoryEntry().Properties['distinguishedName'].Value;$users=[ADSI]('LDAP://'+$controller+'/<WKGUID=A9D1CA15768811D1ADED00C04FD8D5CD,'+$dn+'>');$dialog.FindName('OrganizationalUnit').Text=[string]$users.Properties['distinguishedName'].Value}
                } catch {$status.Text='Domain details could not be detected; complete them before continuing.'}
            } else {$dialog.FindName('DomainController').Text='dc.synthetic.invalid';$dialog.FindName('OrganizationalUnit').Text='CN=Users,DC=synthetic,DC=invalid'}
            $dialog.Add_Closing({param($sender,$e) if ($accountProgress.Active) {$e.Cancel=$true}}.GetNewClosure())
            $dialog.Add_Closed({$dialog.FindName('AccountPassword').Clear()}.GetNewClosure())
            $dialog.FindName('Cancel').Add_Click({$dialog.Close()}.GetNewClosure())
            $dialog.FindName('Create').Add_Click({
                if ($accountProgress.Active) {return}
                $secret=$null;$action=$null;$started=$false
                try {
                    $data=[pscustomobject]@{Name=$dialog.FindName('AccountName').Text.Trim();DomainController=$dialog.FindName('DomainController').Text.Trim();OrganizationalUnit=$dialog.FindName('OrganizationalUnit').Text.Trim()}
                    $secret=$dialog.FindName('AccountPassword').SecurePassword
                    $action=New-SiPendingCreation -Kind DomainAccount -Data $data -Password $secret
                    if (-not $OfflineSmokeTest -and [Windows.MessageBox]::Show($dialog,('Create domain service account '+$data.Name+' on '+$data.DomainController+' now? Existing accounts will not be changed.'),'Confirm account creation','YesNo','Question') -ne 'Yes') {return}
                    Write-SiAssistantLog -Log $assistantLog -Event Action -Step 'service' -Action 'ConfirmDomainAccountCreation' -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count
                    $accountProgress.Active=$true;$accountProgress.Action=$action
                    foreach ($name in @('Create','Cancel','AccountName','DomainController','OrganizationalUnit','AccountPassword')) {$dialog.FindName($name).IsEnabled=$false}
                    $dialog.FindName('ProgressText').Visibility='Visible';$dialog.FindName('Progress').Visibility='Visible'
                    $worker=[PowerShell]::Create();[void]$worker.AddCommand((Join-Path $assistantLibRoot 'SmartM365.DirectoryAccountWorker.ps1')).AddParameter('Data',$data).AddParameter('Password',$action.Secret).AddParameter('LibraryRoot',$assistantLibRoot).AddParameter('OfflineFixture',[bool]$OfflineSmokeTest)
                    $accountProgress.Worker=$worker;$accountProgress.Async=$worker.BeginInvoke();$started=$true
                    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(150);$accountProgress.Timer=$timer
                    $accountRef=$accountProgress;$accountDialogRef=$dialog;$accountStateRef=$serviceState;$accountShareRef=$serviceShareState;$accountExchangeRef=$exchangeRoleState;$accountObservationsRef=$observations;$accountUpdateRef=$updateDetail;$accountStatusRef=$status;$accountLogRef=$assistantLog;$accountActionsRef=$pendingActions;$accountEditsRef=$pendingEdits;$accountOfflineRef=[bool]$OfflineSmokeTest;$accountSmokeRef=$smokeUiState
                    $timer.Add_Tick({
                        if (-not $accountRef.Async.IsCompleted) {return}
                        $accountRef.Timer.Stop()
                        try {
                            $worker=$accountRef.Worker
                            try {$output=@($worker.EndInvoke($accountRef.Async));if ($worker.HadErrors -or $output.Count -ne 1) {throw 'Account creation did not return a complete result.'};$result=$output[0]} finally {$worker.Dispose();$accountRef.Worker=$null}
                            if ($result.State -ne 'CreatedReviewRequired' -or [string]$result.Name -ine [string]$accountRef.Action.Data.Name -or [string]$result.Sid -notmatch '^S-1-(?:\d+-)+\d+$' -or [Uri]::CheckHostName([string]$result.Domain) -ne [UriHostNameType]::Dns) {throw 'Created account identity could not be confirmed.'}
                            $identity=[pscustomobject]@{Account=([string]$result.Name+'@'+[string]$result.Domain);Sid=[string]$result.Sid;Managed=$false;Detail='Account created and identity confirmed.'}
                            $accountStateRef.Identity=$identity;$accountStateRef.Created=$true;if (-not $accountOfflineRef) {$accountShareRef.Ready=$false;$accountExchangeRef.Ready=$false};Set-SiServiceIdentityObservation -Observations $accountObservationsRef -Identity $identity
                            Write-SiAssistantLog -Log $accountLogRef -Event Status -Step 'service' -StatusText 'Domain account created' -PendingActions $accountActionsRef.Count -PendingEdits $accountEditsRef.Count
                            $accountRef.Active=$false;$accountDialogRef.Close();& $accountUpdateRef;$accountStatusRef.Text='Domain service account created.'
                        } catch {
                            $accountRef.Active=$false;$accountDialogRef.FindName('Progress').Visibility='Collapsed';$accountDialogRef.FindName('ProgressText').Text='Creation was not confirmed. Check the directory before retrying.';$accountDialogRef.FindName('Cancel').IsEnabled=$true
                            Write-SiAssistantLog -Log $accountLogRef -Event Status -Step 'service' -StatusText 'Domain account creation uncertain' -PendingActions $accountActionsRef.Count -PendingEdits $accountEditsRef.Count
                            if ($accountOfflineRef) {$accountSmokeRef.SaveError=$_.Exception.Message;$accountDialogRef.Close()} else {[void][Windows.MessageBox]::Show($accountDialogRef,('Creation was not confirmed. Check the directory before retrying. '+$_.Exception.Message),'Account creation unconfirmed','OK','Warning')}
                        } finally {if ($null -ne $accountRef.Action -and $null -ne $accountRef.Action.Secret) {$accountRef.Action.Secret.Dispose();$accountRef.Action.Secret=$null}}
                    }.GetNewClosure())
                    $timer.Start()
                } catch {
                    $accountProgress.Active=$false
                    if ($OfflineSmokeTest) {$smokeUiState.SaveError=$_.Exception.Message;$dialog.Close()} else {[void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Account creation failed','OK','Warning')}
                } finally {
                    if ($null -ne $secret) {$secret.Dispose()}
                    if (-not $started -and $null -ne $action -and $null -ne $action.Secret) {$action.Secret.Dispose();$action.Secret=$null}
                }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch {$status.Text='Cannot open dialog: '+$_.Exception.Message}
    }.GetNewClosure())
    $window.FindName('CreateService').Add_Click({
        try {
            [xml]$dialogXaml=Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-NewLocalServiceDialog.xaml') -Raw
            $dialog=[Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml))
            $dialog.Owner=$window
            $dialog.Add_Closed({ $dialog.FindName('AccountPassword').Clear() }.GetNewClosure())
            $dialog.FindName('Cancel').Add_Click({ $dialog.FindName('AccountPassword').Clear(); $dialog.Close() }.GetNewClosure())
            $dialog.FindName('Create').Add_Click({
                $secret=$null;$pendingAction=$null;$queued=$false
                try {
                    $name=$dialog.FindName('AccountName').Text.Trim()
                    $secret=$dialog.FindName('AccountPassword').SecurePassword
                    $pendingAction=New-SiPendingCreation -Kind LocalAccount -Data ([pscustomobject]@{Name=$name}) -Password $secret
                    Add-SiPendingCreation -Queue $pendingActions -Action $pendingAction;$queued=$true;$tenantBox.IsEnabled=$false
                    $dialog.FindName('AccountPassword').Clear()
                    $dialog.Close()
                    $status.Text='Local account creation added to final review; no account was created.'
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Creation unconfirmed','OK','Warning') }
                finally { if ($null -ne $secret) { $secret.Dispose() };if (-not $queued -and $null -ne $pendingAction -and $null -ne $pendingAction.Secret) {$pendingAction.Secret.Dispose()} }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Cannot open dialog: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('AuditPrerequisites').Add_Click({
        try {
            $status.Text = 'Auditing modules for this account...'
            $result = Invoke-SiPrerequisiteAudit -SmartM365Root $root
            $observations['prerequisites'] = @{State=$result.State; Detail=$result.Detail}
            $window.FindName('StepState').Text = $result.State
            $window.FindName('Action').Text = $result.Detail
            $status.Text = 'Audit limited to this account and PowerShell edition. Check other accounts and hosts.'
        } catch { $status.Text = 'Audit failed: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('CreateProfile').Add_Click({
        try {
            [xml]$dialogXaml = Get-Content -LiteralPath (Join-Path $assistantRoot 'SmartM365-NewTenantAdminDialog.xaml') -Raw
            $dialog = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($dialogXaml))
            $dialog.Owner = $window
            $detectState=@{Tenant='';Account='';Mail='';SiteHost='';InitialDomain='';MissingScopes=@();PreflightReady=$false;PreflightMessage='';Process=$null;Timer=$null;Path=''}
            $adminUi=@{AutoFillKey='';ProfileSuggestionTenant=''}
            $parseAdmins={param([string]$value)
                if ([string]::IsNullOrWhiteSpace($value) -or $value.Length -gt 2048 -or $value -match '[\r\n]') {throw 'Enter at least one administrator email address.'}
                $addresses=@($value -split '[;,]' | ForEach-Object {$_.Trim()} | Where-Object {$_})
                if (-not $addresses.Count) {throw 'Enter at least one administrator email address.'}
                foreach ($address in $addresses) {try {$parsed=[Net.Mail.MailAddress]::new($address)} catch {throw 'Enter valid administrator email addresses.'};if ($parsed.Address -cne $address -or $address -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {throw 'Enter valid administrator email addresses.'}}
                $addresses -join ';'
            }.GetNewClosure()
            $updateProfilePreview={
                if ($detectState.Tenant -and $detectState.InitialDomain -and $adminUi.ProfileSuggestionTenant -cne $detectState.Tenant) {
                    $adminUi.ProfileSuggestionTenant=[string]$detectState.Tenant
                    if ([string]::IsNullOrWhiteSpace($dialog.FindName('ProfileKey').Text)) {
                        $suggestedProfile=([string]$detectState.InitialDomain -split '\.')[0]
                        if ($suggestedProfile -match '^[a-zA-Z0-9][a-zA-Z0-9-]{0,62}$') {$dialog.FindName('ProfileKey').Text=$suggestedProfile}
                    }
                }
                if ($detectState.Tenant -and $detectState.Account -and $adminUi.AutoFillKey -cne ($detectState.Tenant+'|'+$detectState.Account)) {
                    $adminUi.AutoFillKey=$detectState.Tenant+'|'+$detectState.Account
                    if ($detectState.Mail) {$dialog.FindName('Administrators').Text=$detectState.Mail;$dialog.FindName('AdministratorHint').Text='Email from the connected Microsoft 365 account.'} else {$dialog.FindName('Administrators').Text='';$dialog.FindName('AdministratorHint').Text='No email was returned for this account. Enter an administrator email.'}
                }
                try {
                    $suggestion=Get-SiSuggestedProfileIdentity -ProfileName $dialog.FindName('ProfileKey').Text -EnvironmentName $dialog.FindName('EnvironmentKey').Text
                    $null=ConvertTo-SiWorkspaceSharePath -Path $dialog.FindName('WorkspaceRootPath').Text
                    $dialog.FindName('ProfilePreview').Text='Profile: '+$suggestion.ProfileKey
                    $dialog.FindName('OrganizationPreview').Text='Organization: '+$suggestion.OrganizationKey
                    $dialog.FindName('EnvironmentPreview').Text='Environment: '+$suggestion.EnvironmentKey
                    $dialog.FindName('TenantPreview').Text='Data key: '+$suggestion.TenantKey
                    $emailValid=$true;try {$null=& $parseAdmins $dialog.FindName('Administrators').Text} catch {$emailValid=$false}
                    $dialog.FindName('Create').IsEnabled=([bool]$detectState.Tenant -and [bool]$detectState.SiteHost -and [bool]$detectState.InitialDomain -and $detectState.MissingScopes.Count -eq 0 -and [bool]$detectState.PreflightReady -and $emailValid)
                } catch {
                    $dialog.FindName('ProfilePreview').Text='Enter a valid profile, environment and UNC workspace share.'
                    foreach ($name in @('OrganizationPreview','EnvironmentPreview','TenantPreview')) {$dialog.FindName($name).Text=''}
                    $dialog.FindName('Create').IsEnabled=$false
                }
            }.GetNewClosure()
            $profileInstall={ $prereqUi.Owner=$dialog;try {$window.FindName('InstallPrerequisites').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$prereqUi.Owner=$null} }.GetNewClosure()
            Initialize-SiTenantSignInDialog -Dialog $dialog -State $detectState -LibraryRoot $assistantLibRoot -Changed $updateProfilePreview -InstallPrerequisites $profileInstall -OfflineFixture ([bool]$OfflineSmokeTest)
            $dialog.FindName('ProfileKey').Add_TextChanged($updateProfilePreview)
            $dialog.FindName('EnvironmentKey').Add_TextChanged($updateProfilePreview)
            $dialog.FindName('WorkspaceRootPath').Add_TextChanged($updateProfilePreview)
            $dialog.FindName('Administrators').Add_TextChanged($updateProfilePreview)
            $dialog.FindName('ProfileKey').Text = $tenantBox.Text
            $globalAtOpen=Read-SiPrivateJson (Join-Path $root 'Config/SmartM365.global.local')
            if ($globalAtOpen.State -eq 'Present' -and $null -ne $globalAtOpen.Value.PSObject.Properties['WorkspaceRootPath']) {$dialog.FindName('WorkspaceRootPath').Text=[string]$globalAtOpen.Value.WorkspaceRootPath}
            & $updateProfilePreview
            $dialog.FindName('Cancel').Add_Click({ $dialog.Close() }.GetNewClosure())
            $dialog.FindName('Create').Add_Click({
                try {
                    if (-not $detectState.Tenant -or -not $detectState.SiteHost -or -not $detectState.InitialDomain -or $detectState.MissingScopes.Count -or -not $detectState.PreflightReady) {throw 'Connect to Microsoft 365 and resolve the account preflight results.'}
                    $suggestion=Get-SiSuggestedProfileIdentity -ProfileName $dialog.FindName('ProfileKey').Text -EnvironmentName $dialog.FindName('EnvironmentKey').Text
                    $profileKey=$suggestion.ProfileKey
                    $organization=$suggestion.OrganizationKey
                    $environment=$suggestion.EnvironmentKey
                    $workspace=ConvertTo-SiWorkspaceSharePath -Path $dialog.FindName('WorkspaceRootPath').Text
                    if (@($pendingActions | Where-Object {$_.Kind -in @('Profile','Global','Bootstrap')}).Count -or @($pendingEdits | Where-Object {$_.Snapshot.Step -eq 'share'}).Count) {throw 'Review or discard the existing pending profile or shared configuration changes first.'}
                    $plan=Get-SiTenantCreationPlan -Root $root -ProfileKey $profileKey -OrganizationKey $organization -EnvironmentKey $environment
                    $profileAction=New-SiPendingCreation -Kind Profile -Data $plan
                    $adminValue=& $parseAdmins $dialog.FindName('Administrators').Text
                    $adminSnapshot=Get-SiPendingCreationSnapshot -Action $profileAction -Tenant $profileKey -Step administrators -Field Smart365Administrators
                    Set-SiPendingCreationSetting -Snapshot $adminSnapshot -Value $adminValue
                    $errorMailSnapshot=Get-SiPendingCreationSnapshot -Action $profileAction -Tenant $profileKey -Step administrators -Field ErrorMailTo
                    Set-SiPendingCreationSetting -Snapshot $errorMailSnapshot -Value $adminValue
                    $reportMailSnapshot=Get-SiPendingCreationSnapshot -Action $profileAction -Tenant $profileKey -Step mail -Field To
                    Set-SiPendingCreationSetting -Snapshot $reportMailSnapshot -Value $adminValue
                    $domainSnapshot=Get-SiPendingCreationSnapshot -Action $profileAction -Tenant $profileKey -Step tenant -Field OrgDomain
                    Set-SiPendingCreationSetting -Snapshot $domainSnapshot -Value $detectState.InitialDomain
                    $siteSnapshot=Get-SiPendingCreationSnapshot -Action $profileAction -Tenant $profileKey -Step sharepoint -Field SharePointSiteHostname
                    Set-SiPendingCreationSetting -Snapshot $siteSnapshot -Value $detectState.SiteHost
                    $bootstrapIntent=New-SiAutomaticBootstrapIntent -Root $root -PendingProfile $profileAction.Data -TenantId $detectState.Tenant
                    $globalAction=$null;$workspaceEdit=$null
                    $globalCurrent=Read-SiPrivateJson (Join-Path $root 'Config/SmartM365.global.local')
                    if ($globalCurrent.State -eq 'Missing') {
                        $globalAction=New-SiPendingCreation -Kind Global -Data (Get-SiGlobalCreationPlan -Root $root)
                        $defaultSnapshot=Get-SiPendingCreationSnapshot -Action $globalAction -Tenant $profileKey -Step share -Field DefaultTenant
                        Set-SiPendingCreationSetting -Snapshot $defaultSnapshot -Value $profileKey
                        $workspaceSnapshot=Get-SiPendingCreationSnapshot -Action $globalAction -Tenant $profileKey -Step share -Field WorkspaceRootPath
                        Set-SiPendingCreationSetting -Snapshot $workspaceSnapshot -Value $workspace
                    } elseif ($globalCurrent.State -eq 'Present') {
                        $workspaceSnapshot=Get-SiEditSnapshot -Root $root -Tenant $profileKey -Step share -Field WorkspaceRootPath
                        if (-not $workspaceSnapshot.Present -or $workspaceSnapshot.Value -cne $workspace) {$workspaceEdit=New-SiPendingEdit -Snapshot $workspaceSnapshot -Value $workspace}
                    } else {throw 'The existing global configuration is unreadable or conflicting.'}
                    if ($tenantBox.Text -cne $profileKey) {
                        $profileSelectionState.Preparing=$true
                        try {
                            if (-not $profileChoices.Items.Contains($profileKey)) {[void]$profileChoices.Items.Add($profileKey)}
                            $profileChoices.SelectedItem=$profileKey
                        } finally {$profileSelectionState.Preparing=$false}
                    }
                    Add-SiPendingCreation -Queue $pendingActions -Action $profileAction
                    Add-SiPendingCreation -Queue $pendingActions -Action $bootstrapIntent
                    if ($null -ne $globalAction) {Add-SiPendingCreation -Queue $pendingActions -Action $globalAction}
                    if ($null -ne $workspaceEdit) {Add-SiPendingEdit -Queue $pendingEdits -Edit $workspaceEdit}
                    $tenantSession.Tenant=$detectState.Tenant;$tenantSession.Account=$detectState.Account;$tenantSession.Mail=$detectState.Mail;$tenantSession.SiteHost=$detectState.SiteHost;$tenantSession.InitialDomain=$detectState.InitialDomain;$tenantSession.PreflightReady=$detectState.PreflightReady;$tenantSession.Profile=$profileKey
                    $tenantBox.IsEnabled=$false
                    $window.FindName('ProfileDiscoveryStatus').Text='New profile pending review: '+$profileKey+'.'
                    $dialog.Close()
                    & $runPassiveChecks
                    $status.Text='Profile, workspace and bootstrap added to final review; no configuration was written.'
                } catch { [void][Windows.MessageBox]::Show($dialog,$_.Exception.Message,'Creation refused','OK','Warning') }
            }.GetNewClosure())
            [void]$dialog.ShowDialog()
        } catch { $status.Text='Cannot open dialog: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('NewProfile').Add_Click({$window.FindName('CreateProfile').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}.GetNewClosure())
    $window.FindName('LoadField').Add_Click({
        try {
            $step=[string]$steps.SelectedItem.Id
            $creation=@($pendingActions | Where-Object {($_.Kind -eq 'Global' -and $step -eq 'share') -or ($_.Kind -eq 'Profile' -and $step -ne 'share' -and $_.Data.ProfileKey -ceq $tenantBox.Text)} | Select-Object -First 1)
            $editState.Snapshot = if ($creation.Count) {Get-SiPendingCreationSnapshot -Action $creation[0] -Tenant $tenantBox.Text -Step $step -Field ([string]$field.SelectedItem)} else {Get-SiEditSnapshot -Root $root -Tenant $tenantBox.Text -Step $step -Field ([string]$field.SelectedItem)}
            $valueBox.Text = $editState.Snapshot.Value
            $pendingValue=@($pendingEdits | Where-Object {$_.Snapshot.Path -ieq $editState.Snapshot.Path -and $_.Snapshot.Field -ieq $editState.Snapshot.Field} | Select-Object -Last 1)
            if ($pendingValue.Count) {$valueBox.Text=$pendingValue[0].Value}
            if ($editState.Snapshot.Virtual) {$virtualValue=@($editState.Snapshot.Action.Data.PendingEdits | Where-Object {$_.Step -eq $step -and $_.Field -eq $editState.Snapshot.Field} | Select-Object -Last 1);if ($virtualValue.Count) {$valueBox.Text=$virtualValue[0].Value}}
            $valueBox.ToolTip=if ($editState.Snapshot.Field -eq 'AdComputerNameCountryPrefixes') {'Enter a JSON object, for example {"FR":"France"}.'} elseif ($editState.Snapshot.Field -eq 'AdDomainCountryMappings') {'Enter a JSON object, for example {"north.example.test":"France"}.'} elseif ($editState.Snapshot.Field -eq 'DeviceSystemMaxObservationAgeHours') {'Enter null or a finite positive number of hours.'} elseif ($editState.Snapshot.Field -eq 'ExpectedThirdPartyAV') {'Enter a JSON object with DisplayName and a RegisteredNames array.'} elseif ($editState.Snapshot.Field -eq 'EndpointProtectionPolicy') {'Enter the policy as a JSON object. Existing scopes and requirements must be preserved or explicitly changed.'} elseif ($editState.Snapshot.Field -eq 'ReportMailSchedule') {'Enter the schedule as a JSON object. Enabling it can defer report delivery and requires trusted SIDs and a separately provisioned queue.'} elseif ($editState.Snapshot.Field -in @('MailBrandingEnabled','MailPowerBiEnabled') -and $editState.Snapshot.Step -eq 'share') {'Enter true or false.'} elseif ($editState.Snapshot.Field -in @('MailBrandingEnabled','MailPowerBiEnabled') -or $editState.Snapshot.Field -like 'EnableSharePointUpload*') {'Enter true, false or __USE_GLOBAL__.'} elseif ($editState.Snapshot.Field -eq 'MailClientLogoMaxKB') {'Enter a positive size in KB; tenant settings can also inherit.'} else {$null}
            $status.Text = if ($pendingValue.Count) {'Pending value loaded; the configuration is unchanged.'} elseif ($editState.Snapshot.Present) {'Value loaded; no changes saved.'} else {'Setting absent; template default loaded. Review additions before saving.'}
        } catch { $editState.Snapshot = $null; $valueBox.Text = ''; $status.Text = 'Read failed: file missing, invalid, conflicting or setting not editable.' }
    }.GetNewClosure())
    $window.FindName('SaveField').Add_Click({
        try {
            $snapshot = $editState.Snapshot
            if ($null -eq $snapshot -or $snapshot.Tenant -cne $tenantBox.Text -or $snapshot.Step -cne $steps.SelectedItem.Id -or $snapshot.Field -cne [string]$field.SelectedItem) { throw 'Reload the setting before saving.' }
            if ($snapshot.Virtual) {Set-SiPendingCreationSetting -Snapshot $snapshot -Value $valueBox.Text;$editState.Snapshot=$null;$tenantBox.IsEnabled=$false;$status.Text='Setting added to the pending configuration creation; no file was written.';return}
            if ($snapshot.Present -and $snapshot.Value -ceq $valueBox.Text) {
                $removed=0
                foreach ($old in @($pendingEdits | Where-Object {$_.Snapshot.Path -ieq $snapshot.Path -and $_.Snapshot.Field -ieq $snapshot.Field})) {if ($pendingEdits.Remove($old)) {$removed++}}
                if (-not $pendingEdits.Count -and -not $pendingActions.Count) {$tenantBox.IsEnabled=$true}
                $status.Text=if ($removed) {'Pending setting removed; no configuration was written.'} else {'No change.'}
                return
            }
            $pending=New-SiPendingEdit -Snapshot $snapshot -Value $valueBox.Text
            Add-SiPendingEdit -Queue $pendingEdits -Edit $pending
            $editState.Snapshot = $null
            $tenantBox.IsEnabled=$false
            $status.Text = 'Setting added to final review; no configuration was written.'
        } catch { $editState.Snapshot = $null; $status.Text = 'Save refused: ' + $_.Exception.Message }
    }.GetNewClosure())
    $window.FindName('Previous').Add_Click({
        if ($phaseOptions.SelectedIndex -gt 0) {$phaseOptions.SelectedIndex--}
        elseif ($phases.SelectedIndex -gt 0) {$phases.SelectedIndex--;if ($phaseOptions.Items.Count -gt 0) {$phaseOptions.SelectedIndex=$phaseOptions.Items.Count-1}}
    }.GetNewClosure())
    $window.FindName('MarkReviewed').Add_Click({
        if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) { return }
        $row=$steps.SelectedItem
        if ($null -eq $row -or $row.Id -eq 'final') { return }
        $decision=[string]$window.FindName('ReviewDecision').SelectedItem.Content
        if ($decision -notin @('Reviewed','Needs action','Not applicable')) { throw 'Choose a valid review decision.' }
        $reviewed[[string]$row.Id]=$decision
        if ($steps.SelectedIndex -lt ($steps.Items.Count-1)) { $steps.SelectedIndex++ } else { & $updateDetail }
        $status.Text='Decision recorded for this session; actual access and operations are not implied.'
    }.GetNewClosure())
    $window.FindName('GuidedAction').Add_Click({
        if ($null -ne $installState.Run -or $null -ne $mailState.Run -or $null -ne $bootstrapState.Run -or $null -ne $consoleState.Run -or $null -ne $taskInstallerState.Run -or $null -ne $permissionState.Run) { return }
        $row=$steps.SelectedItem
        if ($null -eq $row) { return }
        $profileState='Present'
        if ($row.Id -eq 'tenant') { $profileState=(Read-SiPrivateJson (Join-Path $root ("Config/Tenants/$($tenantBox.Text).local"))).State }
        $serviceResolved=($observations.ContainsKey('service') -and [bool]$observations['service'].Sid)
        $action=Get-SiWorkflowNextAction -Step $row.Id -ProfileState $profileState -ServiceResolved $serviceResolved
        if (-not $action.Control) { $status.Text=$action.Label; return }
        $control=$window.FindName($action.Control)
        if ($null -eq $control -or -not $control.IsEnabled -or ($control.Visibility -ne 'Visible' -and -not ($row.Id -eq 'welcome' -and $action.Control -eq 'SetScope'))) { $status.Text='The guided action is unavailable; review its prerequisites.'; return }
        $control.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }.GetNewClosure())
    $window.FindName('Refresh').Add_Click({if ($null -ne $steps.SelectedItem -and $steps.SelectedItem.Id -eq 'share') {& $startShareCheck}}.GetNewClosure())
    $window.FindName('Next').Add_Click({
        if ($steps.SelectedItem.Id -eq 'share' -and -not $shareState.Ready) {$status.Text='The CSV and log share must be accessible before continuing.';return}
        if ($steps.SelectedItem.Id -eq 'service' -and (-not (Get-SiServiceReadiness -Identity $serviceState.Identity -PendingActions $pendingActions).Ready -or -not $serviceShareState.Ready -or (-not $OfflineSmokeTest -and ($serviceShareState.Sid -cne [string]$serviceState.Identity.Sid -or $serviceShareState.Path -ine $shareState.Path)) -or ($scope.Exchange -and (-not $exchangeRoleState.Ready -or (-not $OfflineSmokeTest -and $exchangeRoleState.Sid -cne [string]$serviceState.Identity.Sid))))) {$status.Text='Create or select a service account before continuing.';return}
        if ($phaseOptions.SelectedIndex -lt ($phaseOptions.Items.Count-1)) {$phaseOptions.SelectedIndex++}
        elseif ($phases.SelectedIndex -lt $phases.Items.Count-1) {$phases.SelectedIndex++}
    }.GetNewClosure())
    $reviewInvalidation=@{
        AuditPrerequisites='prerequisites';InstallPrerequisites='prerequisites';InstallSigningCertificate='prerequisites';ConfigureService='service';CheckService='service';CheckServiceRights='service';CreateService='service';CreateDirectoryService='service'
        CheckShare='share';CheckShareAsService='share';CheckShareWrite='share';CreateProfile='tenant';CheckCertificate='identity';SetCertificate='identity';PlanBootstrap='identity';ReconcileBootstrap='identity';CheckPermissions='permissions'
        SetUploadCategory='sharepoint';CheckMail='mail';TestMail='mail';CheckCatalog='collectors';CheckTask='orchestrator';PlanTask='orchestrator';RegisterTask='orchestrator';EditOrchestratorSettings='orchestrator';OpenOrchestratorConsole='orchestrator'
    }
    foreach ($controlName in $reviewInvalidation.Keys) {
        $invalidateStep=[string]$reviewInvalidation[$controlName]
        $invalidateControl=[string]$controlName
        $window.FindName($controlName).Add_Click({
            if ($invalidateControl -ne 'CreateProfile' -and $null -ne $assessmentState.Active) {$assessmentState.Generation++;$assessmentState.Pending=$false}
            if ($steps.Items.Count -eq 15) {
                Clear-SiWorkflowReviewsFromStep -Steps @($steps.Items) -Decisions $reviewed -Step $invalidateStep
                if ($null -eq $installState.Run -and $null -eq $mailState.Run -and $null -eq $bootstrapState.Run -and $null -eq $consoleState.Run -and $null -eq $taskInstallerState.Run -and $null -eq $permissionState.Run) { & $updateDetail }
            }
        }.GetNewClosure())
    }
    $window.FindName('SetEnabled').Add_Click({
        if ($null -ne $assessmentState.Active) {$assessmentState.Generation++;$assessmentState.Pending=$false}
        if ($steps.Items.Count -eq 15 -and $steps.SelectedItem.Id -in @('teams','sharepoint')) { Clear-SiWorkflowReviewsFromStep -Steps @($steps.Items) -Decisions $reviewed -Step ([string]$steps.SelectedItem.Id); & $updateDetail }
    }.GetNewClosure())
    if ($null -ne $assistantLog) {
        $logStatus=[ComponentModel.DependencyPropertyDescriptor]::FromProperty([Windows.Controls.TextBlock]::TextProperty,[Windows.Controls.TextBlock])
        $logStatus.AddValueChanged($status,{
            $step=if ($null -ne $steps.SelectedItem) {[string]$steps.SelectedItem.Id} else {''}
            Write-SiAssistantLog -Log $assistantLog -Event Status -Step $step -StatusText $status.Text -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count
            if ($assistantLog.Failed -and $window.Title -notlike '*[[]log unavailable[]]*') {$window.Title+=' [log unavailable]'}
        }.GetNewClosure())
        $steps.Add_SelectionChanged({if ($null -ne $steps.SelectedItem) {Write-SiAssistantLog -Log $assistantLog -Event Step -Step ([string]$steps.SelectedItem.Id) -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}}.GetNewClosure())
        $logControls=@('Refresh','RunLocalBatch','CreateProfile','AuditPrerequisites','InstallPrerequisites','InstallSigningCertificate','StopPrerequisites','ConfigureService','CreateService','CreateDirectoryService','CheckShare','CheckShareAsService','CheckShareWrite','CheckTask','CheckCertificate','SetCertificate','CheckMail','TestMail','StopMail','CheckCatalog','SetEnabled','SetUploadCategory','SetScope','CheckService','CheckServiceRights','PlanTask','EditOrchestratorSettings','OpenOrchestratorConsole','PlanBootstrap','ReconcileBootstrap','CheckPermissions','StopPermissions','LoadField','SaveField','Previous','Next','GuidedAction','MarkReviewed')
        foreach ($name in $logControls) {
            $button=$window.FindName($name);if ($null -eq $button) {continue}
            $logName=$name
            $button.Add_Click({$step=if ($null -ne $steps.SelectedItem) {[string]$steps.SelectedItem.Id} else {''};Write-SiAssistantLog -Log $assistantLog -Event Action -Step $step -Action $logName -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}.GetNewClosure())
        }
        foreach ($entry in @(@($createGlobalButton,'CreateGlobal'),@($storageButton,'SetStorage'),@($applyPendingButton,'ApplyReviewed'),@($discardPendingButton,'DiscardReviewed'),@($identityButton,'EditIdentity'))) {
            $logName=[string]$entry[1];$entry[0].Add_Click({$step=if ($null -ne $steps.SelectedItem) {[string]$steps.SelectedItem.Id} else {''};Write-SiAssistantLog -Log $assistantLog -Event Action -Step $step -Action $logName -PendingActions $pendingActions.Count -PendingEdits $pendingEdits.Count}.GetNewClosure())
        }
    }
    & $refresh
    if ($OfflineSmokeTest) {
        if ($window.WindowState -ne 'Maximized') {throw 'The main window is not configured to open maximized.'}
        $window.WindowState='Normal'
        $window.ShowActivated=$false; $window.ShowInTaskbar=$false; $window.Opacity=0
        $window.Show()
        & $runPassiveChecks
        $waitAssessment={
            $limit=[DateTime]::UtcNow.AddSeconds(25)
            while ($null -ne $assessmentState.Active -or $assessmentState.Pending) {
                if ([DateTime]::UtcNow -gt $limit) {throw 'Offline local checks timed out.'}
                $frame=[Windows.Threading.DispatcherFrame]::new();$pulse=[Windows.Threading.DispatcherTimer]::new();$pulse.Interval=[TimeSpan]::FromMilliseconds(50)
                $pulse.Add_Tick({$pulse.Stop();$frame.Continue=$false}.GetNewClosure());$pulse.Start();[Windows.Threading.Dispatcher]::PushFrame($frame)
            }
            if ($status.Text -like 'Local checks failed:*') {throw $status.Text}
        }.GetNewClosure()
        try {
            if ($OfflineMultipleProfilesSmokeTest) {
                if ($profileChoices.Items.Count -ne 2 -or $tenantBox.Text -ne 'beta' -or $profileChoices.SelectedItem -ne 'beta') {throw 'Existing global default was not selected in the GUI.'}
                if ($null -eq $assessmentState.Popup -or -not $assessmentState.Popup.IsVisible -or $null -eq $assessmentState.Active) {throw 'The progress popup was not shown while checks run.'}
                $profileChoices.SelectedItem='test'
                & $waitAssessment
                if ($tenantBox.Text -ne 'test' -or $steps.Items.Count -ne 15 -or $smokeUiState.PassiveRuns -ne 1 -or $null -ne $assessmentState.Popup -or -not $observations.ContainsKey('prerequisites')) {throw 'Latest profile checks were not applied after discarding the stale result.'}
                $existingProfile=Join-Path $root 'Config/Tenants/test.local.json';$beforeProfile=(Get-FileHash -LiteralPath $existingProfile).Hash
                Set-SiPendingServiceAccount -Root $root -Tenant 'test' -Account 'TEST\first' -Actions $pendingActions -Edits $pendingEdits
                Set-SiPendingServiceAccount -Root $root -Tenant 'test' -Account 'TEST\service' -Actions $pendingActions -Edits $pendingEdits
                if ($pendingEdits.Count -ne 1 -or $pendingEdits[0].Value -cne 'TEST\service' -or (Get-FileHash -LiteralPath $existingProfile).Hash -cne $beforeProfile) {throw 'Existing profile service account was not staged safely.'}
                $null=Invoke-SiPendingEdits -Queue $pendingEdits
                if ((Get-Content -LiteralPath $existingProfile -Raw | ConvertFrom-Json).ServiceAccount -cne 'TEST\service') {throw 'Existing profile service account was not saved.'}
                Set-SiPendingServiceAccount -Root $root -Tenant 'test' -Account 'TEST\service' -Actions $pendingActions -Edits $pendingEdits
                if ($pendingEdits.Count) {throw 'An unchanged service account was staged again.'}
                $existingGlobal=Join-Path $root 'Config/SmartM365.global.local.json';$beforeGlobal=(Get-FileHash -LiteralPath $existingGlobal).Hash
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(100)
                $timer.Add_Tick({
                    $dialog=@($window.OwnedWindows | Where-Object {$_.IsVisible} | Select-Object -First 1)
                    if ($dialog.Count) {$timer.Stop();if ($dialog[0].FindName('WorkspaceRootPath').Text -cne '\\synthetic.invalid\old') {throw 'Existing global workspace was not prefilled.'};$dialog[0].FindName('ProfileKey').Text='Another';$dialog[0].FindName('EnvironmentKey').Text='Prod';$dialog[0].FindName('WorkspaceRootPath').Text='\\synthetic.invalid\new';$dialog[0].FindName('DetectTenant').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));$dialog[0].FindName('Create').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
                }.GetNewClosure())
                $timer.Start();try {$window.FindName('NewProfile').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$timer.Stop()}
                & $waitAssessment
                if ($tenantBox.Text -ne 'another' -or @($pendingActions | Where-Object Kind -eq 'Profile').Count -ne 1 -or @($pendingActions | Where-Object Kind -eq 'Global').Count -ne 0 -or @($pendingEdits | Where-Object {$_.Snapshot.Field -eq 'WorkspaceRootPath' -and $_.Value -ceq '\\synthetic.invalid\new'}).Count -ne 1 -or (Get-FileHash -LiteralPath $existingGlobal).Hash -cne $beforeGlobal) {throw 'Existing global workspace was not staged without writing.'}
                $steps.SelectedIndex=4
                if ($window.FindName('Detail').Text -notlike '*CSV and log share: \\synthetic.invalid\new (pending review)*') {throw 'The pending workspace was not shown in shared locations.'}
                $steps.SelectedIndex=14
                if ($window.FindName('Detail').Text -notmatch 'WorkspaceRootPath' -or $window.FindName('Detail').Text -notmatch 'another') {throw 'Existing global workspace change was not shown in final review.'}
                $automatic=@($pendingActions | Where-Object Kind -eq 'Bootstrap')
                if ($automatic.Count -ne 1 -or $automatic[0].Data.Plan.BootstrapPlan.DisplayName -ne 'SMART-M365-INVENTORY' -or -not $automatic[0].Data.Plan.BootstrapPlan.ConfigureTeams) {throw 'Bootstrap was not automatically added for the new profile.'}
                [void]$pendingActions.Remove($automatic[0])
                $machineState.Ready=$true;$machineState.Known=$true
                $serviceState.Identity=[pscustomobject]@{Account='TEST\service';Sid='S-1-5-21-1-2-3-1001';Managed=$false}
                $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                $savedGlobal=Get-Content -LiteralPath $existingGlobal -Raw | ConvertFrom-Json
                if ($pendingActions.Count -ne 0 -or $pendingEdits.Count -ne 0 -or $savedGlobal.WorkspaceRootPath -cne '\\synthetic.invalid\new' -or $savedGlobal.DefaultTenant -cne 'beta' -or $savedGlobal.Preserved -cne 'keep') {throw 'Existing global configuration was not safely updated.'}
                'Offline multiple-profile WPF smoke test passed.'
                return
            }
            if ($OfflineEmptyProfileSmokeTest) {
                if ($tenantBox.Text -ne '' -or $profileChoices.Items.Count -ne 0 -or $discovery.State -ne 'New' -or $smokeUiState.PassiveRuns -ne 0 -or $window.FindName('GuidedAction').Visibility -ne 'Collapsed' -or -not $window.FindName('NewProfile').IsEnabled) {throw 'Empty startup did not offer new profile creation without a guided action.'}
                $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(100)
                $timer.Add_Tick({
                    $dialog=@($window.OwnedWindows | Where-Object {$_.IsVisible} | Select-Object -First 1)
                    if ($dialog.Count) {$timer.Stop();$dialog[0].FindName('DetectTenant').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($dialog[0].FindName('ProfileKey').Text -cne 'synthetic') {throw 'Initial domain was not suggested as the profile name.'};$dialog[0].FindName('ProfileKey').Text='';if ($dialog[0].FindName('ProfileKey').Text -ne '') {throw 'Cleared profile name was unexpectedly restored.'};$dialog[0].FindName('ProfileKey').Text='NewProfile';if ($dialog[0].FindName('EnvironmentKey').Text -ne '' -or $dialog[0].FindName('Create').IsEnabled) {throw 'An environment was assumed without user input.'};$dialog[0].FindName('EnvironmentKey').Text='Lab';if ($dialog[0].FindName('Create').IsEnabled) {throw 'A workspace share was assumed without user input.'};$dialog[0].FindName('WorkspaceRootPath').Text='\\synthetic.invalid\wch-data$';if ($dialog[0].FindName('OrganizationPreview').Text -ne 'Organization: newprofile' -or $dialog[0].FindName('EnvironmentPreview').Text -ne 'Environment: lab' -or $dialog[0].FindName('ProfilePreview').Text -ne 'Profile: newprofile') {throw 'Suggested profile values were not displayed.'};$dialog[0].FindName('DetectTenant').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($dialog[0].FindName('ProfileKey').Text -cne 'NewProfile') {throw 'A user-entered profile name was overwritten.'};if ($dialog[0].FindName('Administrators').Text -ne 'admin@synthetic.invalid') {throw 'Graph administrator email was not prefilled.'};$dialog[0].FindName('Administrators').Text='invalid';if ($dialog[0].FindName('Create').IsEnabled) {throw 'Invalid administrator email was accepted.'};$dialog[0].FindName('Administrators').Text='ops@synthetic.invalid';$dialog[0].FindName('Create').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
                }.GetNewClosure())
                $timer.Start();try {$window.FindName('NewProfile').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$timer.Stop()}
                if ($tenantBox.Text -ne 'newprofile' -or $window.FindName('GuidedAction').Visibility -ne 'Visible' -or @($pendingActions | Where-Object {$_.Kind -eq 'Profile' -and $_.Data.ProfileKey -eq 'newprofile'}).Count -ne 1 -or (Test-Path -LiteralPath (Join-Path $root 'Config/Tenants/newprofile.local.json.txt'))) {throw 'New profile was not staged without writing.'}
                if ($null -eq $assessmentState.Popup -or -not $assessmentState.Popup.IsVisible) {throw 'The pending profile did not show the progress popup.'}
                & $waitAssessment
                if ($smokeUiState.PassiveRuns -ne 1 -or -not $observations.ContainsKey('prerequisites') -or $observations['tenant'].State -ne 'New profile pending review') {throw 'Local checks did not run for the pending profile.'}
                if ($tenantSession.Tenant -ne '11111111-2222-3333-4444-555555555555' -or $tenantSession.Profile -ne 'newprofile') {throw 'Confirmed tenant was not retained for the new profile.'}
                $preparedProfile=@($pendingActions | Where-Object Kind -eq 'Profile')[0].Data;if (@($preparedProfile.PendingEdits | Where-Object {$_.Field -eq 'Smart365Administrators' -and $_.Value -eq 'ops@synthetic.invalid'}).Count -ne 1) {throw 'Administrator email was not staged in the new profile.'}
                if (@($preparedProfile.PendingEdits | Where-Object {$_.Field -eq 'ErrorMailTo' -and $_.Value -eq 'ops@synthetic.invalid'}).Count -ne 1) {throw 'Error-mail recipient was not staged from the administrator email.'}
                if (@($preparedProfile.PendingEdits | Where-Object {$_.Field -eq 'To' -and $_.Value -eq 'ops@synthetic.invalid'}).Count -ne 1) {throw 'Report-mail recipient was not staged from the administrator email.'}
                $steps.SelectedIndex=6;$automatic=@($pendingActions | Where-Object Kind -eq 'Bootstrap')
                if ($automatic.Count -ne 1 -or $automatic[0].Data.Plan.BootstrapPlan.Directory -ne $tenantSession.Tenant -or $automatic[0].Data.Plan.BootstrapPlan.DisplayName -ne 'SMART-M365-INVENTORY' -or -not $automatic[0].Data.Plan.BootstrapPlan.ConfigureTeams -or $window.FindName('PlanBootstrap').Visibility -ne 'Collapsed') {throw 'Bootstrap was not automatically queued with the confirmed tenant and defaults.'}
                [void]$pendingActions.Remove($automatic[0])
                $globalPending=@($pendingActions | Where-Object Kind -eq 'Global')
                if ($globalPending.Count -ne 1 -or @($globalPending[0].Data.PendingEdits | Where-Object {$_.Field -eq 'DefaultTenant' -and $_.Value -eq 'newprofile'}).Count -ne 1 -or @($globalPending[0].Data.PendingEdits | Where-Object {$_.Field -eq 'WorkspaceRootPath' -and $_.Value -ceq '\\synthetic.invalid\wch-data$'}).Count -ne 1) {throw 'The new global configuration did not retain the selected profile and workspace.'}
                $steps.SelectedIndex=4
                if ($window.FindName('Detail').Text -notlike '*CSV and log share: \\synthetic.invalid\wch-data$ (pending review)*' -or $createGlobalButton.Visibility -ne 'Collapsed') {throw 'The pending workspace or global creation was not shown correctly in shared locations.'}
                $shareTimer=[Windows.Threading.DispatcherTimer]::new();$shareTimer.Interval=[TimeSpan]::FromMilliseconds(100)
                $shareTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'CSV and log share'} | Select-Object -First 1);if ($d.Count) {$shareTimer.Stop();if ($d[0].FindName('Workspace').Text -cne '\\synthetic.invalid\wch-data$') {throw 'Pending workspace was not prefilled.'};$d[0].FindName('Save').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}}.GetNewClosure())
                $shareTimer.Start();try {$storageButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$shareTimer.Stop()}
                if (@($globalPending[0].Data.PendingEdits | Where-Object {$_.Field -eq 'WorkspaceRootPath' -and $_.Value -ceq '\\synthetic.invalid\wch-data$'}).Count -ne 1) {throw 'Pending global workspace was not preserved.'}
                $machineState.Ready=$true;$machineState.Known=$true
                $steps.SelectedIndex=14
                if ($window.FindName('Detail').Text -notmatch 'SharePointSiteHostname: synthetic\.sharepoint\.invalid' -or $window.FindName('Detail').Text -match 'contoso\.sharepoint\.com') {throw 'The final review did not use the detected SharePoint host.'}
                if ($window.FindName('Detail').Text -notmatch 'OrgDomain: synthetic\.onmicrosoft\.com' -or $window.FindName('Detail').Text -match 'OrgDomain: example\.invalid') {throw 'The final review did not use the detected initial domain.'}
                if ($window.FindName('Detail').Text -notmatch 'DefaultTenant: newprofile' -or $window.FindName('Detail').Text -match 'Template defaults:|"DefaultTenant"|DefaultTenant: test') {throw 'The final review displayed raw or overridden template values.'}
                if ($window.FindName('Detail').Text -notmatch 'organization: newprofile' -or $window.FindName('Detail').Text -notmatch 'environment: lab' -or $window.FindName('Detail').Text -notmatch 'DefaultTenant' -or $window.FindName('Detail').Text -notmatch 'WorkspaceRootPath') {throw 'The proposed identity and global settings were not shown in final review.'}
                $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                if ($pendingActions.Count -ne 2 -or (Test-Path -LiteralPath (Join-Path $root 'Config/Tenants/newprofile.local.json.txt')) -or $status.Text -notmatch 'service account') {throw 'Apply proceeded without a service account.'}
                $serviceState.Identity=[pscustomobject]@{Account='TEST\service';Sid='S-1-5-21-1-2-3-1001';Managed=$false}
                Set-SiPendingServiceAccount -Root $root -Tenant $tenantBox.Text -Account $serviceState.Identity.Account -Actions $pendingActions -Edits $pendingEdits
                $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                $created=Join-Path $root 'Config/Tenants/newprofile.local.json.txt'
                $createdGlobal=Join-Path $root 'Config/SmartM365.global.local.json.txt'
                $createdProfile=if (Test-Path -LiteralPath $created) {Get-Content -LiteralPath $created -Raw | ConvertFrom-Json} else {$null}
                if ($createdProfile.ServiceAccount -cne 'TEST\service') {throw 'The selected service account was not saved in the new profile.'}
                if ($createdProfile.ErrorMailTo -cne $createdProfile.Smart365Administrators) {throw 'Error-mail recipient does not match the selected administrator email.'}
                if ($createdProfile.To -cne $createdProfile.Smart365Administrators) {throw 'Report-mail recipient does not match the selected administrator email.'}
                if ($createdProfile.OrgDomain -ne 'synthetic.onmicrosoft.com') {throw 'The detected initial domain was not saved.'}
                if ($createdProfile.SharePointSiteHostname -ne 'synthetic.sharepoint.invalid') {throw 'The detected SharePoint host was not saved.'}
                if ($applyProgressState.Active -or $applyProgressState.Starts -ne 1 -or $applyProgressState.Finishes -ne 1 -or $applyProgressState.LastResult -ne 'Success') {throw 'Apply progress or completion status was not completed.'}
                $createdGlobalModel=Get-Content -LiteralPath $createdGlobal -Raw | ConvertFrom-Json
                if ($taskOfferState.Offers -ne 1 -or $pendingActions.Count -ne 0 -or $null -eq $createdProfile -or $createdProfile.ProfileKey -ne 'newprofile' -or $createdProfile.OrganizationKey -ne 'newprofile' -or $createdProfile.EnvironmentKey -ne 'lab' -or $createdProfile.Smart365Administrators -ne 'ops@synthetic.invalid' -or $createdGlobalModel.DefaultTenant -ne 'newprofile' -or $createdGlobalModel.WorkspaceRootPath -cne '\\synthetic.invalid\wch-data$') {throw 'Reviewed synthetic configuration creation did not complete.'}
                $snapshot=Get-SiEditSnapshot -Root $root -Tenant 'newprofile' -Step share -Field DefaultTenant
                Add-SiPendingEdit -Queue $pendingEdits -Edit (New-SiPendingEdit -Snapshot $snapshot -Value 'otherprofile')
                [IO.File]::AppendAllText($createdGlobal,' ')
                $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                if ($applyProgressState.Active -or $applyProgressState.Starts -ne 2 -or $applyProgressState.Finishes -ne 2 -or $applyProgressState.LastResult -ne 'Error') {throw 'A failed write did not close progress with an error result.'}
                'Offline empty-profile WPF smoke test passed.'
                return
            }
            if ($null -eq $assessmentState.Popup -or -not $assessmentState.Popup.IsVisible) {throw 'Startup did not show the progress popup.'}
            & $waitAssessment
            if ($steps.Items.Count -ne 15) { throw 'Expected 15 steps.' }
            if ($phases.Items.Count -ne 7) {throw 'Expected seven setup stages.'}
            $status.Text=(('synthetic line'+[Environment]::NewLine)*2000);$window.UpdateLayout()
            if ($status.ActualHeight -gt 36 -or $status.Parent.Parent.MaxHeight -gt 64) {throw 'Long status text can hide the navigation.'}
            $machineState.LastInstallDetail="Synthetic detail`nLast error";$window.FindName('ShowInstallDetails').Visibility='Visible'
            $detailTimer=[Windows.Threading.DispatcherTimer]::new();$detailTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $detailTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Prerequisite installation details'} | Select-Object -First 1);if ($d.Count) {$detailTimer.Stop();if ($d[0].FindName('Details').Text -ne $machineState.LastInstallDetail -or -not $d[0].FindName('Details').IsReadOnly) {throw 'Details could not be read.'};$d[0].Close()}}.GetNewClosure())
            $detailTimer.Start();try {$window.FindName('ShowInstallDetails').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$detailTimer.Stop()}
            $machineState.LastInstallDetail='';$window.FindName('ShowInstallDetails').Visibility='Collapsed';$status.Text='Offline test status.'

            $phases.SelectedIndex=1;$shareState.Ready=$false;& $updateDetail
            if ($window.FindName('Next').IsEnabled) {throw 'Unavailable share did not block Next.'}
            $phases.SelectedIndex=4
            if ($phases.SelectedIndex -ne 1 -or $steps.SelectedItem.Id -ne 'share') {throw 'Stage navigation bypassed unavailable share.'}
            $shareState.Ready=$true;& $updateDetail
            if (-not $window.FindName('Next').IsEnabled) {throw 'Accessible synthetic share did not unblock Next.'}
            $phases.SelectedIndex=2
            if ($window.FindName('Next').IsEnabled -or $prereqList.Children.Count -ne 1) {throw 'Missing prerequisites did not block Next or show their status.'}
            $phases.SelectedIndex=4
            if ($phases.SelectedIndex -ne 2 -or $steps.SelectedItem.Id -ne 'prerequisites' -or $phaseOptions.SelectedItem.Id -ne 'prerequisites') {throw 'Stage navigation bypassed missing prerequisites.'}
            $machineState.Ready=$true;$machineState.Known=$true;& $updateDetail
            $phases.SelectedIndex=3
            if ($window.FindName('Next').IsEnabled -or $window.FindName('StepState').Text -ne 'Service account required') {throw 'Missing service account did not block Next.'}
            $phases.SelectedIndex=4
            if ($phases.SelectedIndex -ne 3 -or $steps.SelectedItem.Id -ne 'service') {throw 'Stage navigation bypassed the service account.'}
            $accountTimer=[Windows.Threading.DispatcherTimer]::new();$accountTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $accountTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Existing service account'} | Select-Object -First 1);if ($d.Count) {$accountTimer.Stop();$d[0].FindName('Account').Text='TEST\service';$d[0].FindName('Check').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}}.GetNewClosure())
            $accountTimer.Start();try {$window.FindName('CheckService').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$accountTimer.Stop()}
            & $updateDetail
            if (-not $window.FindName('Next').IsEnabled -or $window.FindName('StepState').Text -ne 'Service account and share ready') {throw 'Selecting an existing account did not unblock navigation.'}
            $shareState.Path='\\synthetic.invalid\data$';$serviceShareState.Ready=$false;$smokeUiState.ShareProbeDenied=$true;& $updateDetail
            if ($window.FindName('Next').IsEnabled -or $serviceShareButton.Visibility -ne 'Visible') {throw 'Unverified service share access did not block navigation.'}
            $phases.SelectedIndex=4;if ($phases.SelectedIndex -ne 3 -or $steps.SelectedItem.Id -ne 'service') {throw 'Stage navigation bypassed service share access.'}
            $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($status.Text -notmatch 'Verify service-account share access') {throw 'Apply bypassed service share access.'}
            $serviceShareTimer=[Windows.Threading.DispatcherTimer]::new();$serviceShareTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $serviceShareTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Service account share access'} | Select-Object -First 1);if ($d.Count) {$serviceShareTimer.Stop();$d[0].FindName('Password').Password='synthetic-only';$d[0].FindName('Check').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}}.GetNewClosure())
            $shareWatch=[Windows.Threading.DispatcherTimer]::new();$shareWatch.Interval=[TimeSpan]::FromSeconds(5)
            $shareWatch.Add_Tick({$shareWatch.Stop();$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Service account share access'} | Select-Object -First 1);if ($d.Count) {$smokeUiState.SaveError=$d[0].FindName('ProgressText').Text;$d[0].Close()}}.GetNewClosure())
            $serviceShareTimer.Start();$shareWatch.Start();try {$serviceShareButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$serviceShareTimer.Stop();$shareWatch.Stop()}
            if (-not $serviceShareState.Ready -or $serviceShareState.Sid -ne 'S-1-5-21-1-2-3-1001' -or -not $exchangeRoleState.Ready -or $exchangeRoleState.Sid -ne 'S-1-5-21-1-2-3-1001' -or -not $window.FindName('Next').IsEnabled) {throw ('Synthetic share and Exchange role path did not finish successfully: '+$smokeUiState.SaveError)}
            $exchangeRoleState.Ready=$false;& $updateDetail;if ($window.FindName('Next').IsEnabled) {throw 'Missing Exchange role membership did not block Next.'};$exchangeRoleState.Ready=$true;& $updateDetail
            $smokeUiState.ShareProbeDenied=$false
            $savedServiceIdentity=$serviceState.Identity;$serviceState.Identity=$null;[void]$observations.Remove('service');& $updateDetail
            $failureTimer=[Windows.Threading.DispatcherTimer]::new();$failureTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $failureTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Create directory service account'} | Select-Object -First 1);if ($d.Count) {$failureTimer.Stop();$d[0].FindName('AccountName').Text='svc-fail';$d[0].FindName('AccountPassword').Password='synthetic-only';$d[0].FindName('Create').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}}.GetNewClosure())
            $failureTimer.Start();try {$window.FindName('CreateDirectoryService').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$failureTimer.Stop()}
            & $updateDetail
            if ($null -ne $serviceState.Identity -or $window.FindName('Next').IsEnabled -or $smokeUiState.SaveError -notmatch 'Synthetic directory failure') {throw 'An unconfirmed creation incorrectly allowed navigation.'}
            $smokeUiState.SaveError=''
            $domainTimer=[Windows.Threading.DispatcherTimer]::new();$domainTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $domainTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Create directory service account'} | Select-Object -First 1);if ($d.Count) {$domainTimer.Stop();$d[0].FindName('AccountName').Text='svc-fixture';$d[0].FindName('AccountPassword').Password='synthetic-only';$d[0].FindName('Create').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($d[0].FindName('Progress').Visibility -ne 'Visible' -or $d[0].FindName('Create').IsEnabled) {throw 'Account creation progress was not displayed.'}}}.GetNewClosure())
            $domainTimer.Start();try {$window.FindName('CreateDirectoryService').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$domainTimer.Stop()}
            if (@($pendingActions | Where-Object Kind -eq 'DomainAccount').Count -ne 0 -or -not $window.FindName('Next').IsEnabled -or $window.FindName('StepState').Text -ne 'Service account and share ready' -or $serviceState.Identity.Account -ne 'svc-fixture@synthetic.invalid') {throw ('Simulated creation did not unblock navigation: '+$smokeUiState.SaveError)}
            $serviceState.Identity=$savedServiceIdentity;$serviceState.Created=$false;Set-SiServiceIdentityObservation -Observations $observations -Identity $savedServiceIdentity;& $updateDetail
            $phases.SelectedIndex=5
            if ($steps.SelectedItem.Id -ne 'sharepoint' -or @($phaseOptions.Items | Where-Object {$_.Id -in @('teams','mail')}).Count) {throw 'Teams notifications or the mail page remained in navigation.'}
            if ($window.FindName('RunLocalBatch').Visibility -ne 'Collapsed' -or $window.FindName('Refresh').Content -ne 'Refresh' -or $smokeUiState.PassiveRuns -ne 1 -or -not $observations.ContainsKey('prerequisites')) {throw 'Startup did not run passive checks with a single Refresh button.'}
            if (-not $observations.ContainsKey('welcome') -or $observations['welcome'].Detail -cne 'Selection: AD, Exchange, Mail. No jobs enabled automatically.') {throw 'Default scope was not retained without opening the dialog.'}
            foreach ($view in @('SmartM365-ServiceIdentityDialog.xaml','SmartM365-CreateDomainAccountNowDialog.xaml')) {
                $viewText=Get-Content -LiteralPath (Join-Path $assistantRoot $view) -Raw
                if ($viewText -match '(?i)gmsa|group managed') {throw 'Managed-account choices remained in the service dialogs.'}
            }
            for ($phaseIndex=0;$phaseIndex -lt $phases.Items.Count;$phaseIndex++) {
                $phases.SelectedIndex=$phaseIndex
                $expectedFirst=if ($phases.SelectedItem.Steps[0] -eq 'discovery') {'share'} elseif ($phases.SelectedItem.Steps[0] -eq 'tenant') {'identity'} elseif ($phases.SelectedItem.Steps[0] -eq 'teams') {'sharepoint'} else {$phases.SelectedItem.Steps[0]};if ($steps.SelectedItem.Id -ne $expectedFirst) {throw 'Setup stage did not select its first section.'}
                foreach ($optionId in @($phaseOptions.Items | ForEach-Object {$_.Id})) {$option=@($phaseOptions.Items | Where-Object {$_.Id -eq $optionId})[0];$phaseOptions.SelectedItem=$option;if ($steps.SelectedItem.Id -ne $optionId) {throw 'Setup section navigation failed.'}}
            }
            $phases.SelectedIndex=4
            if ($steps.SelectedItem.Id -ne 'identity' -or @($phaseOptions.Items | Where-Object {$_.Id -in @('tenant','permissions','administrators')}).Count) {throw 'Informational tenant, permission or administrator page remained accessible.'}
            $phases.SelectedIndex=1
            if ($steps.SelectedItem.Id -ne 'share' -or @($phaseOptions.Items | Where-Object Id -eq 'discovery').Count) {throw 'Initial assessment remained accessible.'}
            $window.FindName('Next').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($steps.SelectedItem.Id -ne 'prerequisites') {throw 'Next did not advance to the next setup stage.'}
            $window.FindName('Previous').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($steps.SelectedItem.Id -ne 'share') {throw 'Previous did not return to the last setup section.'}
            $checksBeforeRefresh=$smokeUiState.PassiveRuns
            $window.FindName('Refresh').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($null -eq $assessmentState.Popup -or -not $assessmentState.Popup.IsVisible) {throw 'Refresh did not show the progress popup.'}
            & $waitAssessment
            if ($smokeUiState.PassiveRuns -ne ($checksBeforeRefresh+1)) {throw 'Refresh did not rerun passive checks.'}
            if ($serviceState.Identity.Account -ne 'TEST\service' -or $observations.service.Sid -ne 'S-1-5-21-1-2-3-1001') {throw 'Refreshing lost the selected service account.'}
            $machineState.Ready=$true;$machineState.Known=$true
            if (-not $observations.ContainsKey('mail') -or $observations['permissions'].State -ne 'Manual or external check required') { throw ('Passive workflow batch did not report its limits: '+$status.Text) }
            for ($index=0; $index -lt $steps.Items.Count; $index++) {
                $steps.SelectedIndex=$index
                if ([string]::IsNullOrWhiteSpace($window.FindName('Title').Text)) { throw 'Step navigation did not update the title.' }
            }
            $steps.SelectedIndex=0
            $window.FindName('MarkReviewed').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($steps.SelectedIndex -ne 1 -or -not $reviewed.ContainsKey('welcome')) { throw 'Session review navigation failed.' }
            $window.FindName('ReviewDecision').SelectedIndex=1
            $window.FindName('MarkReviewed').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($reviewed['discovery'] -ne 'Needs action') { throw 'Needs-action decision was not recorded.' }
            $window.FindName('ReviewDecision').SelectedIndex=2
            $window.FindName('MarkReviewed').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($reviewed['prerequisites'] -ne 'Not applicable') { throw 'Not-applicable decision was not recorded.' }
            $steps.SelectedIndex=14
            if ($window.FindName('Action').Text -notmatch 'Review the planned changes') { throw 'Final review instruction is missing.' }
            for ($index=1; $index -lt 14; $index++) {
                $steps.SelectedIndex=$index
                $window.FindName('ReviewDecision').SelectedIndex=0
                $window.FindName('MarkReviewed').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                if ($steps.SelectedIndex -ne ($index+1)) { throw 'Review navigation did not advance.' }
            }
            if ($reviewed.Count -ne 14 -or $window.FindName('Action').Text -notmatch 'Review the planned changes') { throw 'Final review instruction is missing.' }
            $steps.SelectedIndex=1; $bootstrapState.Run=@{}
            try {
                $window.FindName('ReviewDecision').SelectedIndex=1
                $window.FindName('MarkReviewed').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                $window.FindName('GuidedAction').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                if ($steps.SelectedIndex -ne 1 -or $reviewed.Count -ne 14 -or $reviewed['discovery'] -ne 'Reviewed') { throw 'A workflow action ran while another operation was active.' }
            } finally { $bootstrapState.Run=$null }
            $steps.SelectedIndex=11
            $window.FindName('CheckMail').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($reviewed.Count -ne 11 -or $reviewed.ContainsKey('mail') -or -not $reviewed.ContainsKey('sharepoint')) { throw 'A changed check did not invalidate downstream decisions.' }
            $steps.SelectedIndex=0
            if ($window.FindName('SetScope').Visibility -ne 'Collapsed' -or $window.FindName('GuidedAction').Content -ne 'Choose workloads' -or -not $window.FindName('GuidedAction').IsEnabled) {throw 'Welcome must show only the Choose workloads action.'}
            $expectedControls=@{2='InstallPrerequisites';3='ConfigureService';6='SetCertificate';11='CheckMail';12='CheckCatalog';13='CheckTask'}
            foreach ($index in $expectedControls.Keys) {
                $steps.SelectedIndex=$index
                if ($window.FindName($expectedControls[$index]).Visibility -ne 'Visible') { throw 'A step action was not made visible.' }
            }
            $steps.SelectedIndex=6
            if ($window.FindName('PlanBootstrap').Visibility -ne 'Visible' -or $window.FindName('SetCertificate').Visibility -ne 'Visible' -or @('CheckCertificate','GuidedAction','ReconcileBootstrap' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Identity actions were not simplified.'}
            $bootstrapState.ProposalAwaiting=$true;& $updateDetail
            if ($window.FindName('ReconcileBootstrap').Visibility -ne 'Visible') {throw 'A bootstrap proposal cannot be reviewed.'}
            $bootstrapState.ProposalAwaiting=$false
            $certificateExportState.Thumb=('A'*40);$certificateExportState.Profile=$tenantBox.Text;$certificateExportState.Exported=$false;& $updateDetail
            if ($certificateExportButton.Visibility -ne 'Visible') {throw 'Certificate export retry is unavailable.'}
            $exportTimer=[Windows.Threading.DispatcherTimer]::new();$exportTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $exportTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Export application certificate'} | Select-Object -First 1);if ($d.Count) {$exportTimer.Stop();$d[0].FindName('Password').Password='synthetic-password-one';$d[0].FindName('Confirm').Password='synthetic-password-two';$d[0].FindName('Export').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($d[0].FindName('Message').Text -ne 'The passwords do not match.') {throw 'Password mismatch was not refused.'};$d[0].Close()}}.GetNewClosure())
            $exportTimer.Start();try {$certificateExportButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$exportTimer.Stop()}
            if ($certificateExportState.Exported) {throw 'A cancelled export was reported as complete.'}
            $certificateExportState.Thumb='';$certificateExportState.Profile=''
            $steps.SelectedIndex=7
            if (@('PlanBootstrap','CheckPermissions','StopPermissions','GuidedAction' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Advanced permission checks remained visible.'}
            $steps.SelectedIndex=10
            if (@('SetUploadCategory','SetEnabled','GuidedAction' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'SharePoint actions were not simplified.'}
            $steps.SelectedIndex=14
            if (@('OpenOrchestratorConsole','GuidedAction' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Extra actions must be hidden on the final review.'}
            if ($applyPendingButton.FontWeight -ne [Windows.FontWeights]::Bold) {throw 'The apply action must be bold.'}
            $steps.SelectedIndex=11
            if ($window.FindName('CheckMail').Visibility -ne 'Visible' -or @('TestMail','StopMail','GuidedAction' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Mail actions were not simplified.'}
            $steps.SelectedIndex=2
            if ($window.FindName('InstallPrerequisites').Visibility -ne 'Visible' -or $window.FindName('InstallPrerequisites').Content -ne 'Install prerequisites' -or $window.FindName('AuditPrerequisites').Visibility -ne 'Collapsed' -or $window.FindName('InstallSigningCertificate').Visibility -ne 'Collapsed' -or $window.FindName('StopPrerequisites').Visibility -ne 'Collapsed' -or $window.FindName('GuidedAction').Visibility -ne 'Collapsed') {throw 'Prerequisites must show only one install action.'}
            $prereqTimer=[Windows.Threading.DispatcherTimer]::new();$prereqTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $prereqTimer.Add_Tick({
                $dialog=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Install prerequisites'} | Select-Object -First 1)
                if ($dialog.Count) {$prereqTimer.Stop();if ([bool]$dialog[0].FindName('Trust').IsChecked) {throw 'Repository trust default changed.'};$dialog[0].FindName('Install').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));if ($dialog[0].FindName('Progress').Visibility -ne 'Visible' -or $dialog[0].FindName('Install').IsEnabled -or $dialog[0].FindName('ProgressText').Text -notmatch 'certificate') {throw 'Installation progress was not visible immediately.'}}
            }.GetNewClosure())
            $watcher=[Windows.Threading.DispatcherTimer]::new();$watcher.Interval=[TimeSpan]::FromMilliseconds(100)
            $watcher.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'Install prerequisites'} | Select-Object -First 1);if ($d.Count) {$t=$d[0].FindName('ProgressText').Text;if ($t -match 'Checking module: Synthetic\.' -and $t -match 'Verified: [0-2] / 3' -and $t -match 'Elapsed: \d{2}:\d{2}:\d{2}') {$smokeUiState.ProgressSeen=$true};if ($t -match 'Checking module: Synthetic\.' -and $t -match 'Verified: [1-2] / 3') {$smokeUiState.ProgressAfterOne=$true}}}.GetNewClosure())
            $prereqTimer.Start();$watcher.Start();try {$window.FindName('InstallPrerequisites').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$prereqTimer.Stop();$watcher.Stop()}
            if (-not $smokeUiState.ProgressSeen -or -not $smokeUiState.ProgressAfterOne) {throw 'Module name, progressing count and elapsed time were not shown during verification.'}
            if ($machineState.SimulatedInstalls -ne 1 -or -not $machineState.Ready -or -not $window.FindName('Next').IsEnabled -or @($pendingActions | Where-Object Kind -in @('Prerequisites','SigningCertificate')).Count) {throw ('Offline confirmation must simulate installation without queuing or installing: '+$smokeUiState.SaveError)}
            if (-not (Get-Content -LiteralPath $assistantLog.Path -Raw).Contains('ConfirmPrerequisites')) {throw 'The installation confirmation was not logged.'}
            Clear-SiPendingCreations -Queue $pendingActions;$tenantBox.IsEnabled=$true
            $steps.SelectedIndex=3
            if ($window.FindName('ConfigureService').Visibility -ne 'Visible' -or $window.FindName('GuidedAction').Visibility -ne 'Collapsed' -or @('CheckService','CheckServiceRights','CreateService','CreateDirectoryService' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Service account must show only one configuration action.'}
            foreach ($choice in @(@('Existing','Existing service account'),@('Domain','Create directory service account'))) {
                $routeState=@{Opened=$false;Closed=$false};$routeTimer=[Windows.Threading.DispatcherTimer]::new();$routeTimer.Interval=[TimeSpan]::FromMilliseconds(100)
                $routeTimer.Add_Tick({
                    foreach ($owned in @($window.OwnedWindows | Where-Object IsVisible)) {
                        if ($owned.Title -eq 'Configure service account') {if ($null -ne $owned.FindName('Local')) {throw 'Local account creation remained available.'};$owned.FindName([string]$choice[0]).IsChecked=$true;$owned.FindName('Continue').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}
                        elseif ($owned.Title -eq [string]$choice[1]) {if ($owned.Title -eq 'Create directory service account' -and ($owned.FindName('AccountName').Text -ne 'svc-smartinv' -or $owned.FindName('DomainController').Text -ne 'dc.synthetic.invalid' -or $owned.FindName('OrganizationalUnit').Text -ne 'CN=Users,DC=synthetic,DC=invalid' -or $owned.FindName('AccountPassword').Password.Length -ne 0)) {throw 'Domain account defaults were not populated safely.'};$routeState.Opened=$true;$routeState.Closed=$true;$routeTimer.Stop();$owned.Close()}
                    }
                }.GetNewClosure())
                $routeTimer.Start();try {$window.FindName('ConfigureService').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$routeTimer.Stop()}
                if (-not $routeState.Opened -or -not $routeState.Closed -or $pendingActions.Count) {throw ('Service account selection did not safely open '+$choice[0]+'.')}
            }
            $observations['service']=@{State='Synthetic identity';Detail='Offline fixture';Sid='S-1-5-21-1-2-3-1001';Account='TEST\service$';Managed=$true}
            $steps.SelectedIndex=4; if ($storageButton.Visibility -ne 'Visible' -or $window.FindName('Editor').Visibility -ne 'Collapsed' -or @('CheckShare','CheckShareAsService','CheckShareWrite','GuidedAction' | Where-Object {$window.FindName($_).Visibility -ne 'Collapsed'}).Count) {throw 'Share page contains extra controls.'}
            $steps.SelectedIndex=13; if ($null -ne $window.FindName('CheckGmsaTask') -or $window.FindName('GuidedAction').Content -ne 'Review local orchestrator settings') {throw 'Managed task controls remained available.'}
            $observations['service'].Managed=$false
            $dialogs=@(@(0,'GuidedAction'),@(2,'InstallPrerequisites'),@(3,'CheckService'),@(3,'CreateService'),@(3,'CreateDirectoryService'),@(5,$identityButton),@(6,'SetCertificate'),@(6,'PlanBootstrap'),@(7,'CheckPermissions'),@(10,'SetUploadCategory'),@(11,'TestMail'),@(13,'EditOrchestratorSettings'))
            foreach ($entry in $dialogs) {
                $steps.SelectedIndex=[int]$entry[0]
                $modalState=@{Closed=0;BootstrapSharePoint=$false;ScopeDefaults=$false}
                $closer=[Windows.Threading.DispatcherTimer]::new(); $closer.Interval=[TimeSpan]::FromMilliseconds(100)
                $closer.Add_Tick({
                    foreach ($owned in @($window.OwnedWindows)) {
                        if ($owned.IsVisible) { if ($owned.Title -eq 'Prepare bootstrap') {$modalState.BootstrapSharePoint=([bool]$owned.FindName('ConfigureTeams').IsChecked -and $null -eq $owned.FindName('Directory') -and -not $owned.FindName('Stage').IsEnabled -and $null -ne $owned.FindName('DetectTenant') -and $owned.FindName('DisplayName').Text -ceq 'SMART-M365-INVENTORY' -and $owned.FindName('Prepare').Visibility -eq 'Collapsed' -and $owned.FindName('Stage').Content -eq 'Add to review')};if ($owned.Title -eq 'Workflow scope') {$modalState.ScopeDefaults=([bool]$owned.FindName('Cloud').IsChecked -and $owned.FindName('Cloud').Visibility -eq 'Collapsed' -and [bool]$owned.FindName('AD').IsChecked -and [bool]$owned.FindName('Exchange').IsChecked -and [bool]$owned.FindName('Mail').IsChecked -and -not [bool]$owned.FindName('Teams').IsChecked -and $owned.FindName('Teams').Visibility -eq 'Collapsed')};$modalState.Closed++; $owned.Close() }
                    }
                }.GetNewClosure())
                $closer.Start()
                try { $button=if ($entry[1] -is [Windows.Controls.Button]) {$entry[1]} else {$window.FindName([string]$entry[1])};$button.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
                finally { $closer.Stop() }
                if ($modalState.Closed -ne 1) { throw ('Dialog did not open and close: ' + [string]$entry[1]) }
                if ($entry[1] -eq 'PlanBootstrap' -and -not $modalState.BootstrapSharePoint) {throw 'Bootstrap did not reflect the synthetic SharePoint setting.'}
                if ($entry[1] -eq 'GuidedAction' -and -not $modalState.ScopeDefaults) {throw 'Choose workloads did not show scope defaults.'}
            }
            $steps.SelectedIndex=13
            if ($window.FindName('RegisterTask').IsEnabled) { throw 'Task registration must require a plan.' }
            $steps.SelectedIndex=14
            $window.FindName('Previous').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($phases.SelectedIndex -ne 5 -or $steps.SelectedItem.Id -ne 'sharepoint') { throw 'Previous stage navigation failed.' }
            $window.FindName('Next').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($phases.SelectedIndex -ne 6 -or $steps.SelectedItem.Id -ne 'final') { throw 'Next stage navigation failed.' }
            $steps.SelectedIndex=4
            $shareTimer=[Windows.Threading.DispatcherTimer]::new();$shareTimer.Interval=[TimeSpan]::FromMilliseconds(100)
            $shareTimer.Add_Tick({$d=@($window.OwnedWindows | Where-Object {$_.IsVisible -and $_.Title -eq 'CSV and log share'} | Select-Object -First 1);if ($d.Count) {$shareTimer.Stop();$d[0].FindName('Workspace').Text='\\synthetic.invalid\storage';$d[0].FindName('Save').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))}}.GetNewClosure())
            $shareTimer.Start();try {$storageButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))} finally {$shareTimer.Stop()}
            if (@($pendingEdits | Where-Object {$_.Snapshot.Field -ne 'ServiceAccount'}).Count -ne 1 -or (Get-Content -LiteralPath (Join-Path $root 'Config/SmartM365.global.local.json') -Raw).Trim() -cne '{}') {throw 'A setting was not staged without writing.'}
            $steps.SelectedIndex=14
            if ($window.FindName('Detail').Text -notmatch 'WorkspaceRootPath' -or -not $applyPendingButton.IsEnabled) {throw 'Pending settings were not shown in final review.'}
            $blockedAction=New-SiPendingTestAction -Kind MailTest -Data ([pscustomobject]@{From='synthetic@example.invalid';To='recipient@example.invalid';Transport='Graph'})
            Add-SiPendingCreation -Queue $pendingActions -Action $blockedAction
            $beforeBlocked=(Get-FileHash -LiteralPath (Join-Path $root 'Config/SmartM365.global.local.json')).Hash
            $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if (@($pendingEdits | Where-Object {$_.Snapshot.Field -ne 'ServiceAccount'}).Count -ne 1 -or $blockedAction.State -ne 'Pending' -or $taskOfferState.Offers -ne 0 -or (Get-FileHash -LiteralPath (Join-Path $root 'Config/SmartM365.global.local.json')).Hash -cne $beforeBlocked) {throw 'Offline mode did not block an external action before file writes.'}
            [void]$pendingActions.Remove($blockedAction)
            $applyPendingButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            $appliedProfile=Get-Content -LiteralPath (Join-Path $root 'Config/Tenants/test.local.json') -Raw | ConvertFrom-Json
            if ($appliedProfile.ServiceAccount -cne 'TEST\service' -or $taskOfferState.Offers -ne 1 -or $pendingEdits.Count -ne 0 -or (Get-Content -LiteralPath (Join-Path $root 'Config/SmartM365.global.local.json') -Raw | ConvertFrom-Json).WorkspaceRootPath -ne '\\synthetic.invalid\storage') {throw 'Reviewed synthetic settings were not applied.'}
            $tenantBox.Text='newprofile';$window.FindName('Refresh').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent));& $waitAssessment
            $newPlan=Get-SiTenantCreationPlan -Root $root -ProfileKey newprofile -OrganizationKey synthetic -EnvironmentKey test
            Add-SiPendingCreation -Queue $pendingActions -Action (New-SiPendingCreation -Kind Profile -Data $newPlan)
            $steps.SelectedIndex=5;$field.SelectedItem='OrgDomain'
            $window.FindName('LoadField').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            $valueBox.Text='new.example.invalid';$window.FindName('SaveField').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            if ($newPlan.PendingEdits.Count -ne 1 -or $newPlan.PendingEdits[0].Value -ne 'new.example.invalid' -or (Test-Path -LiteralPath $newPlan.Destination)) {throw 'New profile setting was not prepared through the GUI.'}
            Clear-SiPendingCreations -Queue $pendingActions;$tenantBox.IsEnabled=$true
            $tenantBox.Text='other'
            if ($steps.Items.Count -ne 0 -or $reviewed.Count -ne 0) { throw 'Changing the profile did not clear stale results.' }
            $tenantBox.Text='test'
            $window.FindName('Refresh').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
            & $waitAssessment
            if ($steps.Items.Count -ne 15) { throw 'Refresh did not restore the steps.' }
            if (-not $observations.ContainsKey('welcome') -or $observations['welcome'].Detail -cne 'Selection: AD, Exchange, Mail. No jobs enabled automatically.') {throw 'Default scope was not restored after changing the profile.'}
            $entries=@(Get-Content -LiteralPath $assistantLog.Path | ForEach-Object {$_ | ConvertFrom-Json})
            if ($assistantLog.Failed -or @($entries | Where-Object Event -eq 'Action').Count -eq 0 -or @($entries | Where-Object Event -eq 'Queue').Count -eq 0 -or ((Get-Content -LiteralPath $assistantLog.Path -Raw) -match 'workflow\.invalid|smoke@example|private-synthetic|synthetic-only')) {throw 'Synthetic GUI event logging failed or exposed a private value.'}
            'Offline WPF navigation smoke test passed.'
        } finally { $window.Close() }
    } else {
        Hide-SmartM365GuiSplash -Splash $splash
        $startupChecks=@{Started=$false}
        $window.Add_ContentRendered({if (-not $startupChecks.Started) {$startupChecks.Started=$true;& $runPassiveChecks}}.GetNewClosure())
        [void]$window.ShowDialog()
    }
} finally {
    Close-SiAssistantLog -Log $assistantLog
    if ($null -ne $splash) { Close-SmartM365GuiSplash -Splash $splash }
    if ($smokeRoot -and (Test-Path -LiteralPath $smokeRoot)) {
        $resolved=(Resolve-Path -LiteralPath $smokeRoot).Path
        $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
        if ($resolved -ine [IO.Path]::GetFullPath($smokeRoot) -or -not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected smoke test directory; refusing cleanup.' }
        $config=Join-Path $resolved 'Config'; $tenants=Join-Path $config 'Tenants';$setup=Join-Path $resolved 'Setup'
        foreach ($path in @($resolved,$config,$tenants,$setup)) {
            if (([IO.File]::GetAttributes($path) -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Unexpected linked smoke test directory; refusing cleanup.' }
        }
        [IO.File]::Delete((Join-Path $tenants 'test.local.json'))
        [IO.File]::Delete((Join-Path $tenants 'test.local.json.transport.lock'))
        [IO.File]::Delete((Join-Path $tenants 'beta.local.json.txt'))
        [IO.File]::Delete((Join-Path $tenants 'another.local.json.txt'))
        [IO.File]::Delete((Join-Path $tenants 'another.local.json.transport.lock'))
        [IO.File]::Delete((Join-Path $tenants 'another.local.json.txt.transport.lock'))
        [IO.File]::Delete((Join-Path $tenants 'newprofile.local.json'))
        [IO.File]::Delete((Join-Path $tenants 'newprofile.local.json.transport.lock'))
        [IO.File]::Delete((Join-Path $tenants 'newprofile.local.json.txt'))
        [IO.File]::Delete((Join-Path $tenants 'newprofile.local.json.txt.transport.lock'))
        [IO.File]::Delete((Join-Path $tenants 'tenant.local.json.template'))
        [IO.File]::Delete((Join-Path $config 'SmartM365.global.local.json'))
        [IO.File]::Delete((Join-Path $config 'SmartM365.global.local.json.transport.lock'))
        [IO.File]::Delete((Join-Path $config 'SmartM365.global.local.json.txt'))
        [IO.File]::Delete((Join-Path $config 'SmartM365.global.local.json.txt.transport.lock'))
        [IO.File]::Delete((Join-Path $config 'SmartM365.global.local.json.template'))
        [IO.File]::Delete((Join-Path $tenants 'tenant.mail.local.json.template'))
        [IO.File]::Delete((Join-Path $config 'SmartM365-TenantContext.ps1'))
        [IO.File]::Delete((Join-Path $config 'SmartM365-JsonTransport.policy.psd1'))
        [IO.File]::Delete((Join-Path $setup 'SmartM365-Create-AppRegistration.ps1'))
        $syntheticModule=Join-Path $resolved 'Modules/SmartM365.Core'
        foreach ($path in @((Join-Path $resolved 'Modules'),$syntheticModule)) {if (([IO.File]::GetAttributes($path) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {throw 'Unexpected linked synthetic module directory; refusing cleanup.'}}
        [IO.File]::Delete((Join-Path $syntheticModule 'SmartM365.JsonTransport.psd1'))
        [IO.File]::Delete((Join-Path $syntheticModule 'SmartM365.JsonTransport.psm1'))
        [IO.Directory]::Delete($syntheticModule,$false);[IO.Directory]::Delete((Join-Path $resolved 'Modules'),$false)
        [IO.File]::Delete((Join-Path $setup 'Install-SmartM365-SmartInventoryPrerequisites.ps1'))
        [IO.Directory]::Delete($setup,$false)
        if ($null -ne $assistantLog) {[IO.File]::Delete($assistantLog.Path);[IO.Directory]::Delete((Join-Path $resolved 'Log'),$false)}
        $unexpectedTenantFiles=@(Get-ChildItem -LiteralPath $tenants -Force -ErrorAction Stop)
        if ($unexpectedTenantFiles.Count) {throw ('Unexpected synthetic tenant files; refusing cleanup: '+(($unexpectedTenantFiles | ForEach-Object Name) -join ', '))}
        [IO.Directory]::Delete($tenants,$false); [IO.Directory]::Delete($config,$false); [IO.Directory]::Delete($resolved,$false)
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB9hXSrmB9Z0M/n
# smRH4JUwsftwbTP4auVJjBn8L6xpw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII7ABPvz7GPILBV9svusVHAywoVSqjD5L+51/bqs8BORMA0GCSqG
# SIb3DQEBAQUABIIBgAchX6gCscXZPPVR3Gnn3fHKHFJgzDnb4Q07JKnpEPah/UDB
# L3+VV1heIFRRYjvgIiFqyj6vcxbqjnkmFxVCc+xEBWCwQzzJPOEKpInGy/nRyhAd
# B2GjLNIQ9//PX6T4XPRi3/z5GueZX1EH2RZLPcWxeYCGYHy3WkPAV+ENgh203kRU
# xz0J6CLgSa2Wp8C99bC1WIdB3T06arPhYmBdWcSJUHOYXvwH7vmmTbYjDx7Ed+R7
# 3GfkSRnh7T7Tc7JjqNugF29LSIG/UaDV0QRhG0vtZq6x9Fe7FgHEyu4lBT8YdGcw
# 6M5hn2u8HQywLSzQ59fC8U2ueopAbq751VzNS46jmiq66dxPzj0ptVVfqi2K13z4
# wB37+0LO9RnWmK79gZ/8MGm/tLfjDc6Lewdrn+1zcMjqQ/el2Kzi1x/MlmlEa4Fj
# PKze5niVRnzysIdu9VCRZLmE+fysojqk9b/8tCQVxZEBAUGEg2M+HCRRnEnxgtkp
# CsNnLn7kSZF+sR9N1KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTVaMC8GCSqGSIb3DQEJBDEiBCD2nijHl9CBk87znvr2DaBjM9p4dyeGBgHCgr7a
# ce+PXTANBgkqhkiG9w0BAQEFAASCAgCjBRjrMAwGJRblvArVoxZEhU/z9APhgZlI
# Dvv6jAiv5CfsTJRFXuox8YZnCdHGfhAEPoPhNHmhDs8l3Cea5SXp9VamXx6ZB6Da
# YNl/mRkMWHrgZI4Xn17TWcAgoXipGDXnLG7kyaFEiqbXhA7SlkIoyoULz7JnrTrg
# 2HoCh07RctU260DjsMSH/Sz6HOMuBge4ZVB+rLnZdHfsYhjckzGyr/qeMk0Hvtra
# YwBNfTfeDOfcfHTne6XQo7KlIPQZng2me6/OMpx83jheesxtCLf1Hn1F6bknHUxD
# meNJ8sAB/NsjSTXbmnl9Y9149Li2YZftWrg2jPIkGHlmkD55ygw6f7D/072uDHJu
# tt0UCc52/vXFaI7xudtJI68XpBvNqCWNdf554ZnJNWaK+6UliwxP2iKGoTHus3nd
# z2b9V4aDxhl1kfi+jMtpQddjNJzrcCK3e9X81g2cdh/lf2EknvROkeKSkn5AiYPm
# 0icB//TeyI5UDIzQJuFCJ3h5HnLx3gzNkmGgWzYf3VHWKki05Ckj92t6w53Ltgof
# M2B15a6jQjKRpH3+gACzOk2S+JvU82HPns9w7TwLXoIGsk7v9dm4MQtToLOBU6no
# VZD9CQbNK/5qeum7l7MGPUFYQnWO2LaXG7gyEBEbzl43nh3WTgkWfQUAwNNi3am7
# X778lefZtQ==
# SIG # End signature block
