#Requires -Version 7.4
<#
.SYNOPSIS
Fictional migration portfolio, all dashboard panels, batch UI, wizard and reports.
.VERSION
1.0.4
#>
function script:Switch-SharePointDemoPanel {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    Invoke-SmartM365SharePointLoading -Window $s.Window -Message ("Ouverture de l'onglet : " + $Name) -Action {
    foreach($panel in @('panelSummary','panelArchivedOverview','panelWorkflows','panelDiagnostics','panelBatches','panelOperations','panelLogs','panelConfig')){$c[$panel].Visibility='Collapsed'}
    $target=@{tabSummary='panelSummary';tabArchivedOverview='panelArchivedOverview';tabFiles='panelWorkflows';tabDiagnostics='panelDiagnostics';tabBatches='panelBatches';tabOperations='panelOperations';tabLogs='panelLogs';tabConfig='panelConfig'}[$Name]
    if($target){$c[$target].Visibility='Visible'}
    foreach($tab in @('tabSummary','tabArchivedOverview','tabFiles','tabDiagnostics','tabBatches','tabOperations','tabLogs','tabConfig')){$c[$tab].IsChecked=$tab -eq $Name}
    $s.Data.SelectedPanel=$Name
    }
}
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    . (Join-Path $s.AppRoot 'SmartM365-SharePointMigration-Loading.ps1')
    $presentationSource=Join-Path $s.AppRoot 'SmartM365-SharePointMigration-Summary.ps1'
    . ([scriptblock]::Create((Get-SmartM365DemoFunctionText $presentationSource @('Get-SmartM365PortfolioGapVisual','Get-SmartM365PortfolioRateVisual','Get-SmartM365ComplexityVisual')) -join [Environment]::NewLine))
    if(-not $s.Data.ContainsKey('Migrations')){$s.Data.Migrations=@('Contoso-HR','Contoso-Finance','Contoso-Intranet')}
    Set-SmartM365DemoCombo cmbMigration $s.Data.Migrations
    Set-SmartM365DemoCombo cmbAuthMode @('Interactive (simulated)','Device login (simulated)','Certificate (simulated)')
    $s.Data.SourceUrl='https://sharepoint.contoso.example/hr';$s.Data.TargetUrl='https://contoso.sharepoint.example/sites/hr'
    $complexity = Get-Content -LiteralPath (Join-Path $s.AppRoot 'Config/migration-complexity.demo.json.template') -Raw | ConvertFrom-Json
    if ($s.Scenario -eq 'Issues') { $complexity = $complexity.PartialResult }
    $s.Data.Complexity = $complexity
    $c.lblSourceType.Text='SP2019';$c.lblTargetType.Text='SPO';$c.lblSourceUrl.Text=$s.Data.SourceUrl;$c.lblTargetUrl.Text=$s.Data.TargetUrl
    $portfolio=@(for($i=0;$i -lt $s.Data.Migrations.Count;$i++){
        $fields=@('Name','Migration','MigrationName','Source','Destination','SourceScan','SourcePermissionScan','SourceInventoryDisplay','TargetScan','TargetPermissionScan','TargetInventoryDisplay','FilesComparison','PermissionsComparison','FilesStatus','PermissionsStatus','Complexity','ComplexityDisplay','ComplexityLabel','ComplexityScore','ComplexityTooltip','SourceTooltip','TargetTooltip','SourceScansTooltip','TargetScansTooltip','SourceFileScanIsToday','SourcePermissionScanIsToday','TargetFileScanIsToday','TargetPermissionScanIsToday')
        $row=New-SmartM365DemoRow $fields -Index $i -Values @{Name=$s.Data.Migrations[$i];Migration=$s.Data.Migrations[$i];MigrationName=$s.Data.Migrations[$i];Source=$s.Data.SourceUrl;Destination=$s.Data.TargetUrl;SourceScan='2026-10-18 08:00';SourcePermissionScan='2026-10-18 08:10';TargetScan='2026-10-18 08:20';TargetPermissionScan='2026-10-18 08:30';SourceInventoryDisplay='1,200 files | 35 folders | 4.8 GB';TargetInventoryDisplay='1,188 files | 35 folders | 4.7 GB';FilesComparison='98.33% | 20 missing';PermissionsComparison='98.0% | 10 missing';FilesStatus='Review';PermissionsStatus='Review';Complexity='Medium';ComplexityDisplay='Medium - 35/100';ComplexityLabel='Medium';ComplexityScore=35;ComplexityTooltip='DEMO - Fictional source complexity';SourceFileScanIsToday=$true;SourcePermissionScanIsToday=$true;TargetFileScanIsToday=$true;TargetPermissionScanIsToday=$true}
        $fileRate=if($s.Scenario -eq 'Completed'){100}else{@(98.33,100,94.5)[$i%3]}
        $permissionRate=if($s.Scenario -eq 'Completed'){100}else{@(98.0,100,96.0)[$i%3]}
        $row.FilesComparison=('{0:N2}% matched (sample)' -f $fileRate)
        $row.PermissionsComparison=('{0:N2}% matched (sample)' -f $permissionRate)
        $row.TargetInventoryDisplay=('{0:N0} files | 35 folders | 4.7 GB' -f ([Math]::Round(1200*$fileRate/100)+8))
        $extra=@{ScanGapText='0.02';ScanGapTooltip='DEMO - Source/target evidence gap: 20 minutes';ScanGapVisual=(Get-SmartM365PortfolioGapVisual 0.02);Status='Review';StatusTooltip='DEMO - Fictional differences';ComplexityText='35 / 100';ComplexityCaption='Medium';ComplexityVisual=(Get-SmartM365ComplexityVisual -Color '#9A6700' -Background '#FFF8E8');ComparisonVisual=(Get-SmartM365PortfolioRateVisual -Rate $fileRate -Caption '09:30 (demo)');PermissionComparisonVisual=(Get-SmartM365PortfolioRateVisual -Rate $permissionRate -Caption '09:30 (demo)');GlobalComparisonVisual=(Get-SmartM365PortfolioRateVisual -Rate (($fileRate+$permissionRate)/2) -Caption 'Equal weighting');ComparisonTooltip='DEMO - Source keys matched against target';PermissionComparisonTooltip='DEMO - Source permission keys matched';GlobalComparisonTooltip='DEMO - Mean of files and permissions';SourceInventoryTooltip='DEMO - 1,200 files, 35 folders, 4.8 GB';TargetInventoryTooltip='DEMO - 1,188 files, 35 folders, 4.7 GB'}
        foreach($key in $extra.Keys){$row|Add-Member -NotePropertyName $key -NotePropertyValue $extra[$key] -Force}
        $row.ComplexityText = '{0} / 100' -f $complexity.Score
        $row.ComplexityCaption = $complexity.Level
        $row.ComplexityScore = $complexity.Score
        $row.Complexity = $complexity.Level; $row.ComplexityLabel = $complexity.Level
        $row.ComplexityDisplay = '{0} - {1}/100' -f $complexity.Level,$complexity.Score
        $row.ComplexityCaption += ' - auto'
        $row.ComplexityTooltip = 'DEMO - Automatic source assessment; {0}% coverage. Collection and qualification are separate. Workflow replacement risk remains visible even when qualification is partial.' -f $complexity.CoveragePercent
        $row.ComplexityVisual = Get-SmartM365ComplexityVisual -Color $complexity.LevelColor -Background $complexity.LevelBackground
        $row
    })
    Set-SmartM365DemoGrid gridSummary $portfolio
    Set-SmartM365DemoGrid gridArchivedOverview @($portfolio[0].PSObject.Copy())
    foreach($pair in @(@('lblPortfolioMigrations',[string]$portfolio.Count),@('lblPortfolioFilesCompared',[string]$portfolio.Count),@('lblPortfolioPermissionsCompared',[string]$portfolio.Count),@('lblPortfolioActions','24'),@('lblPortfolioRefreshErrors','0'),@('lblArchivedMigrations','1'),@('lblArchivedFilesCompared','1'),@('lblArchivedPermissionsCompared','1'),@('lblArchivedGlobalCompared','1'),@('lblArchivedRefreshErrors','0'))){$c[$pair[0]].Text=$pair[1]}
    $c.lblSummaryStatus.Text='DEMO - All migration rows are fictional';$c.lblArchivedStatus.Text='DEMO - Archived sample retained for presentation'
    $c.lblWorkflowMigration.Text='Contoso-HR';$c.lblWorkflowSourceUrl.Text=$s.Data.SourceUrl;$c.lblWorkflowTargetUrl.Text=$s.Data.TargetUrl
    $c.lblWorkflowComplexity.Text = '{0} - {1}/100' -f $complexity.Level, $complexity.Score
    $c.lblWorkflowComplexityCaption.Text='DEMO - automatic estimate; {0}% coverage; {1}' -f $complexity.CoveragePercent, $complexity.Status
    $c.lblWorkflowComplexitySummary.Text='13 criteria; automatic source evidence. Collection and qualification tracked separately; effort is estimated.'
    $complexityRows = @($complexity.Criteria | ForEach-Object {
        [pscustomobject]@{Axis=$_.Axis;Criterion=$_.Name;Value=$_.ValueDisplay;Band=$_.Band;Points=$(if($_.Evaluated){'{0} / 3' -f $_.Points}else{'N/A'});Weighted=$(if($_.Evaluated){'{0} / {1}' -f $_.WeightedPoints,$_.MaxWeightedPoints}else{'N/A'});Detail=('DEMO - '+$_.Detail);Thresholds=($_.Bands -join '; ')}
    })
    Set-SmartM365DemoGrid gridWorkflowComplexity $complexityRows
    foreach($name in @('lblScanSrcAge','lblScanTgtAge','lblScanSrcPermAge','lblScanTgtPermAge','lblHistoryAge','lblPermHistoryAge','lblCmpFilesAge','lblCmpPermsAge')){$c[$name].Text='2026-10-18 09:30 (demo)'}
    foreach($name in @('cmbScanSrcFile','cmbScanTgtFile','cmbScanSrcPermFile','cmbScanTgtPermFile','cmbHistoryOldFile','cmbHistoryNewFile','cmbPermHistoryOldFile','cmbPermHistoryNewFile')){Set-SmartM365DemoCombo $name @([pscustomobject]@{Display='Contoso-HR-Inventory-20261018-DEMO.csv'},[pscustomobject]@{Display='Contoso-HR-Inventory-20261011-DEMO.csv'})}
    $c.lblCmpFilesDir.Text='Demo/Comparisons/Files';$c.lblCmpPermsDir.Text='Demo/Comparisons/Permissions'
    $c.lblCmpFilesAvailability.Text='98.33% matched - sample comparison';$c.lblCmpPermsAvailability.Text='98.0% matched - sample comparison'
    $c.bannerCmpFilesRefresh.Visibility='Collapsed';$c.bannerCmpPermsRefresh.Visibility='Collapsed'
    $c.lblDiagMigration.Text='Contoso-HR';$c.lblDiagScope.Text='DEMO - Fictional ShareGate session and scan evidence';$c.lblDiagInputPath.Text='Demo/ShareGate/MigrationReport';$c.lblDiagLatestReport.Text='MigrationDiagnostics-Report-DEMO.html'
    foreach($name in @('lblDiagReportState','lblDiagAnalysisState','lblDiagHtmlState')){$c[$name].Text='Ready (demo)'}
    foreach($name in @('lblDiagReportEvidence','lblDiagAnalysisEvidence','lblDiagHtmlEvidence')){$c[$name].Text='Synthetic evidence generated for screenshots'}
    foreach($pair in @(@('lblSummaryShareGateValue','1,445 lines'),@('lblSummaryFilesValue','98.33%'),@('lblSummaryPermissionsValue','98.0%'),@('lblSummaryCrossCheckValue','2 residual items'),@('lblDiagSourceFiles','1,200 files'),@('lblDiagSourceFolders','35 folders'),@('lblDiagSourceVolume','4.8 GB'),@('lblDiagTargetFiles','1,188 files'),@('lblDiagTargetFolders','35 folders'),@('lblDiagTargetVolume','4.7 GB'),@('lblDiagLines','1445'),@('lblDiagKeyedItems','1200'),@('lblDiagSuccess','1400'),@('lblDiagError','15'),@('lblDiagWarning','30'),@('lblDiagAccepted','43'),@('lblDiagToFixLines','2'),@('lblDiagUnkeyedLines','0'),@('lblDiagItemsToFix','2'),@('lblDiagResidualLines','2'),@('lblDiagResidualItems','2'))){$c[$pair[0]].Text=$pair[1]}
    $c.lblDiagInterpretation.Text='DEMO - Two fictional items remain for operator review.';$c.lblDiagNextAction.Text='Review the sample patterns and cross-check. No ShareGate report was read.'
    $c.lblFarmResult.Text='DEMO - Fictional farm evidence available';$c.lblFarmPeaks.Text='2 sample access windows';$c.lblFarmPrerequisites.Text='DEMO - No SharePoint snap-in, server access or elevation required'
    $c.txtFarmDryRun.Text='DEMO - Farm diagnostic preview; no command is executed';$c.txtFarmRun.Text='DEMO - Farm collection is simulated'
    Set-SmartM365DemoGrid gridDiagPatterns @([pscustomobject]@{Category='Access';State='To fix';Lines=2;Items=2;Pattern='Fictional access warning';AccessEvidence=$s.Data.SourceUrl+'/Documents';Remediation='Review sample permissions';HelpLink='https://workplacecloudhub.com'},[pscustomobject]@{Category='Version';State='Accepted';Lines=43;Items=40;Pattern='Fictional retry/version warnings';AccessEvidence='Contoso sample';Remediation='Accepted in demo';HelpLink='https://workplacecloudhub.com'}) @('Category','State','Lines','Items','Pattern','AccessEvidence','Remediation')
    Set-SmartM365DemoGrid gridDiagRows @([pscustomobject]@{Status='Warning';Source=$s.Data.SourceUrl+'/Documents/Policy.docx';Target=$s.Data.TargetUrl+'/Documents/Policy.docx';Message='DEMO - Sample file requires review.';Date='2026-10-18 09:20'}) @('Status','Source','Target','Message','Date')
    foreach($name in @('gridCrossCheckEvidence','gridCrossCheckScopes')){Set-SmartM365DemoGrid $name @([pscustomobject]@{Scope='Documents';Source='Synthetic scan';Target='Synthetic scan';Status='Review';Details='Two fictional residual differences.'}) @('Scope','Source','Target','Status','Details')}
    $ops=@([pscustomobject]@{DisplayName='Destination site lock state';CmdFile='DEMO-SiteLockState.cmd'},[pscustomobject]@{DisplayName='Disable page comments';CmdFile='DEMO-PageCommentsDisable.cmd'},[pscustomobject]@{DisplayName='Clean extra files';CmdFile='DEMO-ExtraFilesRemove.cmd'},[pscustomobject]@{DisplayName='Clean extra folders';CmdFile='DEMO-ExtraFoldersRemove.cmd'},[pscustomobject]@{DisplayName='Clean extra libraries';CmdFile='DEMO-ExtraLibrariesRemove.cmd'})
    $c.listOps.ItemsSource=$ops
    $c.lblConfigPath.Text='Demo/in-memory/migration.config.json';$c.lblConfigStatus.Text='DEMO - Edits are never written to operational configuration'
    $c.txtConfigContent.Text=(@{Mode='DEMO';Name='Contoso-HR';SourceUrl=$s.Data.SourceUrl;TargetUrl=$s.Data.TargetUrl;Comparison=@{CompareModified=$true;ModifiedDateToleranceMinutes=5;MaxScanAgeDifferenceHours=24}}|ConvertTo-Json -Depth 5)
    $c.txtLogContent.Text='[2026-10-18 09:30:00] DEMO - Source scan, destination scan and comparisons simulated.'
    $c.listLogFiles.ItemsSource=@('DEMO-SourceScan.log','DEMO-TargetScan.log','DEMO-Comparison.log')
    $c.lblLogName.Text='DEMO-Comparison.log';$c.lblLastRefresh.Text='2026-10-18 09:30 (frozen)';$c.lblAppVersion.Text='DEMO'
    if(-not $s.Data.ContainsKey('BatchView')){
        $source=Join-Path $s.AppRoot 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-BatchGui.ps1'
        . ([scriptblock]::Create((Get-SmartM365DemoFunctionText $source @('New-SmartM365BatchView')) -join [Environment]::NewLine))
        $batch=New-SmartM365BatchView;$c.panelBatches.Content=$batch;$s.Data.BatchView=$batch
        $batchNames=@('batchAll','batchNames','batchScope','batchSourceRoot','batchSourcePrerequisites','batchSourceCheck','batchSourceCopy')
        foreach($kind in @('Source','Target','Comparison')){foreach($suffix in @('Mode','Auth','Preview','Run','Status','Latest','Counts','Logs','Summary')){$batchNames+='batch'+$kind+$suffix}}
        foreach($name in $batchNames){$control=$batch.FindName($name);if($control){$c[$name]=$control}}
    }
    $c.batchNames.ItemsSource=$s.Data.Migrations;$c.batchScope.Text=([string]$s.Data.Migrations.Count+' fictional migrations in scope')
    $c.batchSourceRoot.Text='Demo/Toolkit';$c.batchSourcePrerequisites.Text='DEMO - Prerequisites simulated. No farm or server is contacted.'
    foreach($kind in @('Source','Target','Comparison')){$c['batch'+$kind+'Status'].Text='DEMO - Ready';$c['batch'+$kind+'Latest'].Text='2026-10-18 09:30 (sample)';$c['batch'+$kind+'Counts'].Text='Successful actions: 6 | Failed actions: 0 (simulated)'}
    Switch-SharePointDemoPanel tabSummary
}
function script:Connect-SmartM365DemoProfileEvents {
    $s=$script:SmartM365Demo;$c=$s.Controls
    foreach($name in @('tabSummary','tabArchivedOverview','tabFiles','tabDiagnostics','tabBatches','tabOperations','tabLogs','tabConfig')){[void]$s.ViewNames.Add($name)}
    $c.cmbMigration.Add_SelectionChanged({
        $s=$script:SmartM365Demo
        Invoke-SmartM365SharePointLoading -Window $s.Window -Message ('Chargement de la migration : '+[string]$s.Controls.cmbMigration.SelectedItem) -Action {
            $s.Controls.lblWorkflowMigration.Text=[string]$s.Controls.cmbMigration.SelectedItem
            $s.Controls.lblDiagMigration.Text=[string]$s.Controls.cmbMigration.SelectedItem
        }
    })
    $c.listOps.AddHandler([Windows.Controls.Button]::ClickEvent,[Windows.RoutedEventHandler]{param($hostControl,$eventArgs)Write-SmartM365DemoActivity ('Operation simulated: '+[string]$eventArgs.OriginalSource.Tag)},$true)
    $c.batchAll.Add_Click({$c=$script:SmartM365Demo.Controls;$c.batchNames.Visibility=if($c.batchAll.IsChecked){'Collapsed'}else{'Visible'}})
    foreach($pair in @(@('txtDiagFilter','gridDiagPatterns'),@('txtDiagRowFilter','gridDiagRows'))){$c[$pair[0]].Tag=$pair[1];$c[$pair[0]].Add_TextChanged({param($box)$s=$script:SmartM365Demo;$s.Controls[[string]$box.Tag].ItemsSource=@($s.Rows[[string]$box.Tag]|Where-Object {($_.PSObject.Properties.Value -join ' ') -like ('*'+$box.Text+'*')})})}
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    if($Name -like 'tab*'){Switch-SharePointDemoPanel $Name;return $true}
    if($Name -match '^batch'){
        foreach($kind in @('Source','Target','Comparison')){if($Name -like ('batch'+$kind+'*')){$c['batch'+$kind+'Status'].Text='DEMO - Batch completed in memory; no process or connection was started.'}}
        Write-SmartM365DemoActivity ($Name+' simulated. Fictional batch selection and outcomes only.');return $true
    }
    switch($Name){
        'btnNewMigration' {
            $markup=Get-SmartM365DemoMarkup (Join-Path $s.AppRoot 'SmartM365-SharePointMigration-NewWizard.ps1')
            $window=New-SmartM365DemoView $markup;$window.Title+=' | DEMO - Sample data'
            $window.FindName('txtName').Text='Contoso-New-DEMO';$window.FindName('txtSourceUrl').Text=$s.Data.SourceUrl;$window.FindName('txtTargetUrl').Text=$s.Data.TargetUrl;$window.FindName('txtTargetAdmin').Text='https://admin.contoso.example'
            foreach($name in @('txtSourceApp','txtTargetApp')){$window.FindName($name).Text='00000000-0000-4000-8000-000000000001'}
            foreach($name in @('txtSourceRoot','txtTargetRoot')){$window.FindName($name).Text='Demo/Toolkit'}
            $window.FindName('txtAdditionalMappings').Text='[]'
            $window.FindName('lblStatus').Text='DEMO - URL validation and creation are simulated.'
            foreach($buttonName in @('btnCancel','btnCreate')){$button=$window.FindName($buttonName);$button.Tag=$window;$button.Add_Click({param($button)$s=$script:SmartM365Demo;if($button.Name -eq 'btnCreate'){$name=$button.Tag.FindName('txtName').Text;$s.Data.Migrations+=@($name);Set-SmartM365DemoCombo cmbMigration $s.Data.Migrations;Write-SmartM365DemoActivity ('Migration created in memory: '+$name)};$button.Tag.Close()})}
            [void]$s.ChildWindows.Add($window);if(-not $s.ValidateOnly){$window.Owner=$s.Window;$window.Show()};return $true
        }
        'btnArchiveMigration' {Write-SmartM365DemoActivity 'Selected migration archived in demo memory only.';return $true}
        'btnFarmWindow' {
            if($s.Data.ContainsKey('FarmWindow') -and $s.Data.FarmWindow.IsVisible){$s.Data.FarmWindow.Activate();return $true}
            . ([scriptblock]::Create((Get-SmartM365DemoFunctionText $s.GuiPath @('New-FarmDiagnosticsWindow')) -join [Environment]::NewLine))
            $card=$c.farmDiagnosticsHost.Child
            if(-not $card){return $true}
            $c.farmDiagnosticsHost.Child=$null
            $window=New-FarmDiagnosticsWindow -Owner $s.Window -MigrationName 'Contoso-HR | DEMO - Sample data' -Content $card
            $window.Add_Closed({param($window)$card=$window.Content.Children[1].Content;$window.Content.Children[1].Content=$null;$script:SmartM365Demo.Controls.farmDiagnosticsHost.Child=$card})
            $s.Data.FarmWindow=$window;[void]$s.ChildWindows.Add($window)
            if(-not $s.ValidateOnly){$window.Show()}
            return $true
        }
        'btnSaveConfig' {Write-SmartM365DemoActivity 'Configuration edits retained in memory. No migration config was written.';return $true}
    }
    if($Name -match 'Report|btnDiagOpenReport|btnFarmOpenReport'){
        $paths=@(Export-SmartM365DemoProfileOutputs)
        $key=if($Name -match 'Farm'){'Farm'}elseif($Name -match 'Diag'){'Diagnostics'}elseif($Name -match 'Permissions|Perms'){'Permissions'}else{'Files'}
        Show-SmartM365DemoFile $s.Reports[$key];return $true
    }
    return $false
}
function script:Save-SharePointDemoReport {
    param([string]$Name,[string]$Title,[string]$Kind,[Collections.IDictionary]$Metrics,[object[]]$Rows)
    $s=$script:SmartM365Demo
    # Reuse the shipped report renderer's HTML shell and CSS. Do not load branding config.
    $source=[IO.File]::ReadAllText((Join-Path $s.AppRoot 'Scripts/Compare/report_html.py'))
    $css=[regex]::Match($source,'(?s)REPORT_CSS = """(?<CSS>.*?)"""').Groups['CSS'].Value
    if ($Kind -eq 'Source complexity') {
        $css += ' table{table-layout:fixed;width:100%} th:nth-child(1){width:22%} th:nth-child(2){width:10%} th:nth-child(3){width:7%} th:nth-child(4){width:8%} th:nth-child(5){width:53%} td{overflow-wrap:anywhere}'
    }
    $shell=[regex]::Match($source.Substring($source.IndexOf('def render_report(')),"(?s)return f'''(?<HTML>.*?)'''").Groups['HTML'].Value
    if(-not $css -or -not $shell){throw 'The shipped SharePoint report presentation could not be loaded.'}
    $cards=foreach($key in $Metrics.Keys){'<div class="metric"><div class="metric-label">'+[Net.WebUtility]::HtmlEncode($key)+'</div><div class="metric-value">'+[Net.WebUtility]::HtmlEncode([string]$Metrics[$key])+'</div></div>'}
    $table=($Rows|ConvertTo-Html -Fragment) -join [Environment]::NewLine
    $body='<section class="section"><div class="section-heading"><h2>Contoso sample evidence</h2></div><div class="table-scroll">'+$table+'</div></section>'
    $logoPath=Join-Path $s.AssetRoot 'WorkplaceCloudHub-lockup-WPF.png';$logo=if(Test-Path -LiteralPath $logoPath){'<img alt="WorkplaceCloudHub" src="data:image/png;base64,'+[Convert]::ToBase64String([IO.File]::ReadAllBytes($logoPath))+'">'}else{''}
    $values=@{'{escape(title)}'=[Net.WebUtility]::HtmlEncode($Title);'{escape(report_kind)}'=$Kind;'{escape(generated_at)}'='2026-10-18 09:30';'{escape(status_text)}'='DEMO - Sample data';'{escape(status_class)}'='warn';'{cards_html}'=($cards -join '');'{body_html}'=$body;'{escape(footer)}'='DEMO - Synthetic Contoso evidence';'{escape(intro_text)}'='Fictional data generated for demonstration. No scan, comparison, collection or live validation was performed.';'{alert_html}'='';'{download_html}'='';'{logo_html()}'=$logo;'{REPORT_CSS}'=$css}
    foreach($key in $values.Keys){$shell=$shell.Replace($key,[string]$values[$key])}
    $path=Get-SmartM365DemoOutputPath $Name
    [IO.File]::WriteAllText($path,$shell,[Text.UTF8Encoding]::new($false))
    return $path
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo;$paths=[Collections.Generic.List[string]]::new()
    $files=@([pscustomobject]@{Library='Documents';Source=1200;Target=1188;Matched=1180;Missing=20;Extra=8;ChangedVersion=5;TargetOlder=3;Success='98.33%'})
    $permissions=@([pscustomobject]@{Scope='Documents';SourcePermissions=500;TargetPermissions=495;Matched=490;Missing=10;Extra=5;Success='98.0%'})
    $definitions=@(
        @('Complexity','Complexity-DEMO.html','Contoso-HR source migration complexity','Source complexity',[ordered]@{Score=$s.Data.Complexity.Score;Criteria=13;Coverage=('{0}%' -f $s.Data.Complexity.CoveragePercent);Level=$s.Data.Complexity.Level},@($s.Rows.gridWorkflowComplexity|Select-Object Criterion,Value,Points,Weighted,Detail)),
        @('Files','Files-Comparison-DEMO.html','Contoso-HR file comparison','File comparison',[ordered]@{Matched=1180;Missing=20;Extra=8;'Changed version'=5;'Target older'=3;'Success rate'='98.33%'},$files),
        @('Permissions','Permissions-Comparison-DEMO.html','Contoso-HR permissions comparison','Permission comparison',[ordered]@{Matched=490;Missing=10;Extra=5;'Success rate'='98.0%'},$permissions),
        @('FileHistory','File-History-DEMO.html','Contoso-HR file scan history','File history',[ordered]@{Previous=1150;Current=1200;Added=60;Removed=10;Changed=25},$files),
        @('PermissionHistory','Permission-History-DEMO.html','Contoso-HR permissions history','Permission history',[ordered]@{Previous=480;Current=500;Added=25;Removed=5;Changed=12},$permissions),
        @('GlobalFiles','Global-Files-DEMO.html','Contoso migration portfolio - files','Global file comparison',[ordered]@{Migrations=3;Source=3600;Matched=3514;Missing=86;Extra=24},@($s.Rows.gridSummary|Select-Object Migration,Source,Destination,FilesComparison,PermissionsComparison)),
        @('GlobalPermissions','Global-Permissions-DEMO.html','Contoso migration portfolio - permissions','Global permission comparison',[ordered]@{Migrations=3;Source=1500;Matched=1470;Missing=30;Extra=15},$permissions),
        @('Diagnostics','MigrationDiagnostics-Report-DEMO.html','Contoso-HR migration diagnostics','ShareGate diagnostics',[ordered]@{Lines=1445;Success=1400;Warnings=30;Errors=15;'Residual items'=2},$s.Rows.gridDiagPatterns)
    )
    foreach($definition in $definitions){$path=Save-SharePointDemoReport -Name $definition[1] -Title $definition[2] -Kind $definition[3] -Metrics $definition[4] -Rows $definition[5];$s.Reports[$definition[0]]=$path;[void]$paths.Add($path)}
    # Use the actual farm HTML writer, without its collection/startup functions.
    . ([scriptblock]::Create((Get-SmartM365DemoFunctionText (Join-Path $s.AppRoot 'Scripts/Diagnostics/SmartM365-SharePointMigration-FarmDiagnostic.ps1') @('Write-FarmDiagHtml')) -join [Environment]::NewLine))
    function Write-FarmDiagText {param($Path,$Value)[IO.File]::WriteAllText($Path,$Value,[Text.UTF8Encoding]::new($false))}
    $farmPath=Get-SmartM365DemoOutputPath 'Farm-Report-DEMO.html'
    $coverage=@([pscustomobject]@{Server='CTSO-SP-01';Source='WAS/IIS/ULS';Status='DEMO';Detail='Synthetic source farm evidence';DurationSeconds=0},[pscustomobject]@{Server='CTSO-SP-02';Source='WAS/IIS/ULS';Status='DEMO';Detail='Synthetic source farm evidence';DurationSeconds=0})
    $summary=[pscustomobject]@{Project='Contoso-HR (DEMO)';Status='DEMO';GeneratedAtUtc='2026-10-18T07:30:00Z';Coverage=$coverage;PeakCount=2}
    $correlation=@([pscustomobject]@{StartUtc='2026-10-18 07:20';StartLocal='2026-10-18 09:20';EndUtc='2026-10-18 07:25';EndLocal='2026-10-18 09:25';ShareGateLines=8;WasEvents=1;SecurityEvents=0;Iis401Count=2;UlsEvents=1;Servers='CTSO-SP-01'})
    $pools=@([pscustomobject]@{Server='CTSO-SP-01';Pool='Contoso-Content-DEMO';CpuLimitPercent=60;CpuAction='Throttle'})
    $recurrences=@([pscustomobject]@{Server='CTSO-SP-01';EventId=5011;LocalFiveMinuteBucket='02:00';Nights=2;Events=2;ScanDays=7})
    $timeline=@([pscustomobject]@{Utc='2026-10-18 07:21';Local='2026-10-18 09:21';Server='CTSO-SP-01';Source='WAS (sample)';Event='Synthetic worker restart';Count=1;Detail='DEMO - Fictional access window evidence'})
    $risks=@([pscustomobject]@{Server='CTSO-SP-01';Pool='Contoso-Content-DEMO';Risk='Sample throttle overlap';Detail='DEMO - Review fictional timing. No setting was inspected or changed.'})
    Write-FarmDiagHtml -Path $farmPath -Summary $summary -Pools $pools -Events @([pscustomobject]@{Mode='DEMO'}) -IisRows @([pscustomobject]@{Mode='DEMO'}) -Correlation $correlation -Recurrences $recurrences -Timeline $timeline -Risks $risks
    Set-SmartM365DemoMarking $farmPath;$s.Reports.Farm=$farmPath;[void]$paths.Add($farmPath)
    return $paths.ToArray()
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDoaf3GW800PmGk
# e75f9iljGbMq61IPhnKw65wnkgIxeaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGt6AR2spzsKluUEKEin1NGG7KcfnNp3p5QyKef96aX8MA0GCSqG
# SIb3DQEBAQUABIIBgHd5foYBWJn4x/wQDXyVVySLSIjur3Iap+xROGAIQetDtpab
# 2vG9ylPHuLbT8s7LGLGeJAr77lACPGuS+etGFl78SWqJ9bBlg7jeV5QFIeUztoLM
# jmMoHiOpBtcO8LC6YPQIwcIiUKHNxz0VBXm3g0gEt+36AbH6sGX+ODC2Zj5M8ns7
# 8xe2+6NZErPIN6mWbM3cDnh4FdajP1fRR5OFK0YRidvaRz+M2af92xmFTRce8g63
# oyzt4kiJudRTldc+/XDolFf/SpFso15dYG+q5GmcL0QG6qke6JxbuQM3tHojuXLv
# klHe96ExhllLzGf72+rl6OI112quhdtvcKLeeTWfbH7PxawMwXVis3S9/83Updzb
# TM0NTJosn5kWTJNkPlFlssSr9bTj46bRfjYluFHw0MqN+WrvJNm0T5W1SBod6//p
# xAQSOqJ3uJWzzzmnAk20q898xIgREC6cCCx5bbiQLyFZE+jRKTgS/4uQXSpshz7/
# 4p+I36XWqX9/LlRSJaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxMTUz
# NDBaMC8GCSqGSIb3DQEJBDEiBCDjW9IsNraEQw3m6/VKfjUwP90p+OayFSxWwjx9
# /9dHWTANBgkqhkiG9w0BAQEFAASCAgAQH/tolpf01cVUvlzRfjXyFHGk3L8IQhQN
# SO876WUQOjc253yTVYqd8blZmiGMBCm+lq+16yVgGVHnfi3FRuAfqzMlRyA+LAL7
# UmUeSA0b5Cp8JxHhUKOSch9+jx/9KFNvGfdG3+DEHmjLhZRoSseGdvUYuFqAAt0G
# ZLVB/XkkGGVV1NPhvaM/msuslAZG2LVgXKcBlosjUZn3xE07SUcCbyYybV9eKv6H
# AkbB2LlGQgm9kAZuYlKw8MwucI6oDgrmEHGSvf2t2A4htWV2yGJPz9pec6EXxKNh
# xtZp3NfeWMbPswqZl4La4Vcy0oUiCXD/tyb3/FatbP1syNvy4LzViCOUzHY/urJw
# 2AD2UwhGiEvsCXkmLq7YWHXs6FPLeCmhOmJz/B0s3Se2SbmvRSew4smQLwbMwBPT
# /T+6vTTIZWDhUhrpdGUtbEacZbEx8ReQfZtWw1h2uoOkgXNOisREeqhLcIRIqZWx
# Xjbza4nfufDKL69NkAx9DzAgRRVCfQ2ZH7PeAG10qgs2FiVMr3tOYTw2bCSWbFZe
# 31yX3DbSkyvqiiR3V3WrX8CEpYvys9pfQADREv8dZQvHvfSA3tzZsgolxTHcL6EI
# 4OFc9XIDmmbVcz/XuSEMOxm4vC0Yx/3inQP+6qUXeZ3H3QS89tMd5SnmopOHj0we
# X45+Lg7RPA==
# SIG # End signature block
