#Requires -Version 7.0
<#
.SYNOPSIS
All three campaigns and shipped language templates, with fictional recipients.
.VERSION
1.1.0
#>
function script:Get-NotificationsDemoHtml {
    param($Template)
    $s=$script:SmartM365Demo
    $path=Join-Path (Join-Path (Join-Path $s.AppRoot $Template.Campaign) 'Templates') $Template.FileName
    $html=[IO.File]::ReadAllText($path)
    $logoPath=Join-Path $s.AssetRoot 'WorkplaceCloudHub-lockup-WPF.png'
    $logo=if(Test-Path -LiteralPath $logoPath){'<img alt="WorkplaceCloudHub" style="max-height:64px" src="data:image/png;base64,'+[Convert]::ToBase64String([IO.File]::ReadAllBytes($logoPath))+'">'}else{'WorkplaceCloudHub'}
    $tokens=@{Date='2026-11-04';LogoImgTag=$logo;FooterLogoImgTag=$logo;FreeTextBlockStyle='display:block;';FreeTextBlockTextHtml='DEMO - Fictional Contoso campaign. No email will be sent.';Hotline='demo.helpdesk@contoso.example';MailboxQuotaGB='100';MailboxSizeGB='24';MailboxUsageText='24 GB / 100 GB (sample)';NewWebmailUrl='https://mail.contoso.example';OldWebmailUrl='https://legacy-mail.contoso.example';WebmailUrl='https://mail.contoso.example';TermsPortalBlockStyle='display:block;';TermsPortalUrl='https://terms.contoso.example'}
    foreach($match in [regex]::Matches($html,'\{\{(?<Name>[^}]+)\}\}')){
        $name=$match.Groups['Name'].Value
        if(-not $tokens.ContainsKey($name)){throw ('Unknown shipped email token in demo: '+$name)}
        $html=$html.Replace($match.Value,[string]$tokens[$name])
    }
    $html=[regex]::Replace($html,'(?i)<body([^>]*)>','<body$1><div style="background:#fff3cd;padding:12px;font:14px Segoe UI">DEMO - Sample data | Contoso | Fictional recipient: user01@contoso.example</div>',1)
    return $html
}
function script:Update-NotificationsDemoTemplates {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $campaign=[string]$c.TemplateCampaignCombo.SelectedItem
    if(-not $campaign){$campaign='ExchangeMigration'}
    $rows=@($s.Data.Templates|Where-Object Campaign -EQ $campaign)
    $c.TemplateList.ItemsSource=$rows
    $c.TemplateFolderPathText.Text='Shipped '+$campaign+'/Templates (demo substitutions only)'
    $c.TemplateSummaryText.Text=('{0} shipped language templates available. All preview tokens are fictional.' -f $rows.Count)
    $english=@($rows|Where-Object Language -EQ 'en'|Select-Object -First 1)
    $c.TemplateList.SelectedItem=if($english.Count){$english[0]}else{$rows[0]}
    Update-NotificationsDemoPreview
}
function script:Update-NotificationsDemoPreview {
    $s=$script:SmartM365Demo;$c=$s.Controls;$template=$c.TemplateList.SelectedItem
    if(-not $template){return}
    $s.Data.PreviewHtml=Get-NotificationsDemoHtml $template
    if(-not $s.ValidateOnly){$c.TemplatePreviewBrowser.NavigateToString($s.Data.PreviewHtml)}
}
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $campaigns=@('ExchangeMigration','ExchangeArchive','ExchangeMigrationMailboxSizeReduction')
    $templates=@(foreach($campaign in $campaigns){
        $root=Join-Path (Join-Path $s.AppRoot $campaign) 'Templates'
        foreach($file in Get-ChildItem -LiteralPath $root -Filter '*.html' -File|Sort-Object Name){
            $language=if($file.Name -match '\.([a-zA-Z-]+)\.html$'){$matches[1]}else{'default'}
            [pscustomobject]@{Campaign=$campaign;Language=$language;FileName=$file.Name;Source='Shipped/demo';Modified='2026-10-18'}
        }
    })
    $s.Data.Templates=$templates
    $s.Data.Campaign='ExchangeMigration'
    Set-SmartM365DemoCombo TenantProfileCombo @('Contoso - DEMO')
    Set-SmartM365DemoCombo TemplateCampaignCombo $campaigns
    Set-SmartM365DemoCombo ForceLanguageCombo (@('(Auto)')+@($templates.Language|Sort-Object -Unique))
    Set-SmartM365DemoCombo MailSendModeCombo @('Graph (simulated)','SMTP (simulated)')
    Set-SmartM365DemoCombo ExchangeManagementModeCombo @('Exchange Online (simulated)','Exchange on-premises (simulated)')
    $c.CampaignBadgeText.Text='EXCHANGE MIGRATION | DEMO';$c.CampaignTitleText.Text='Contoso mailbox migration';$c.CampaignDescriptionText.Text='Preview the complete campaign experience using fictional recipients and shipped templates.'
    $c.RecipientPathBox.Text='Demo/Contoso-Recipients.csv';$c.RecipientCountText.Text='12 fictional recipients'
    $c.RecipientPathLabel.Text='Recipients CSV';$c.RecipientPathStatusText.Text='DEMO - Sample recipient list ready';$c.RecipientPathHelpText.Text='Browse actions choose generated examples; no operational list is read.'
    $c.ScriptPathText.Text='Shipped campaign script (not executed in demo)';$c.LocalConfigPathText.Text='In-memory Contoso demo configuration';$c.TemplateConfigPathText.Text='Shipped neutral templates only'
    $c.EffectiveDatePicker.SelectedDate=[datetime]'2026-11-04';$c.DryRunRadio.IsChecked=$true
    $c.TeamsUserMessageCheck.IsEnabled=$true
    $rows=[Collections.Generic.List[object]]::new()
    function Add-DemoConfigRows {param($Object,[string]$Prefix='')
        foreach($property in $Object.PSObject.Properties){
            $name=if($Prefix){$Prefix+'.'+$property.Name}else{$property.Name}
            if($property.Value -is [pscustomobject]){Add-DemoConfigRows $property.Value $name}
            else{
                $value=if($property.Name -match 'Tenant|AppId|Thumbprint|Secret|Webhook'){'Not used in demo'}elseif($property.Name -match 'From|To|Email|Mail'){'demo.helpdesk@contoso.example'}elseif($property.Name -match 'Path|Folder|Root'){'Demo/'+$property.Name}else{[string]$property.Value}
                [void]$rows.Add([pscustomobject]@{Name=$name;Value=$value;Source='Demo/template'})
            }
        }
    }
    foreach($relative in @('Config/ExchangeUserNotifications.local.json.template','Config/Campaigns/ExchangeMigration.local.json.template','Config/Campaigns/ExchangeArchive.local.json.template','Config/Campaigns/ExchangeMigrationMailboxSizeReduction.local.json.template')){
        Add-DemoConfigRows (Get-Content -LiteralPath (Join-Path $s.AppRoot $relative) -Raw|ConvertFrom-Json) $relative
    }
    $c.ConfigParameterList.ItemsSource=$rows.ToArray()
    Initialize-NotificationsExoDemo
    if (-not $s.ValidateOnly) { $s.Window.WindowState='Maximized' }
    Update-NotificationsDemoTemplates
}
function script:Initialize-NotificationsExoDemo {
    $s=$script:SmartM365Demo; $c=$s.Controls
    Set-SmartM365DemoCombo ExoTimeZoneCombo @('Romance Standard Time','UTC')
    $c.MigrationRecipientSourceCombo.SelectedIndex=0
    $c.ExoConnectionText.Text='DEMO - Fictional EXO session. Connect and refresh are simulated; no authentication is possible.'
    $c.ExoOverrideDate.SelectedDate=[datetime]'2026-11-04'
    $s.Data.ExoBatches=@(
        [pscustomobject]@{Selected=$true;Identity='DEMO-Paris-Nov04';Status='Synced';MigrationType='ExchangeRemoteMove';TotalCount=6;StartAfter='2026-11-01 00:00';CompleteAfterUTC='2026-11-04T21:00:00Z'},
        [pscustomobject]@{Selected=$true;Identity='DEMO-Lyon-Nov05';Status='Syncing';MigrationType='ExchangeRemoteMove';TotalCount=6;StartAfter='2026-11-02 00:00';CompleteAfterUTC='2026-11-05T21:00:00Z'},
        [pscustomobject]@{Selected=$false;Identity='DEMO-Completed';Status='Completed';MigrationType='ExchangeRemoteMove';TotalCount=4;StartAfter='2026-10-01';CompleteAfterUTC='2026-10-02T21:00:00Z'}
    )
    $s.Data.ExoRows=@(for($i=1;$i -le 12;$i++){
        $day=if($i -le 6){'04'}else{'05'}
        [pscustomobject]@{Selected=$true;EmailAddress=('user{0:00}@contoso.example' -f $i);BatchIds=$(if($i -le 6){'DEMO-Paris-Nov04'}else{'DEMO-Lyon-Nov05'});MigrationStatus='Synced';EffectiveDate=('2026-11-'+$day);DateConfirmed=$true;DateSource='Batch CompleteAfter';ScheduledLocal=('2026-11-'+$day+' 22:00 [Romance Standard Time]');ScheduledUtc=('2026-11-'+$day+'T21:00:00Z');Exclusion='';ScheduleReason='DEMO - Scheduled earliest completion; actual completion may be later'}
    })
    if ($s.Scenario -eq 'Issues') {
        $s.Data.ExoRows[0].Selected=$false;$s.Data.ExoRows[0].EffectiveDate='';$s.Data.ExoRows[0].ScheduledLocal='';$s.Data.ExoRows[0].ScheduledUtc='';$s.Data.ExoRows[0].DateSource='Unavailable';$s.Data.ExoRows[0].DateConfirmed=$false;$s.Data.ExoRows[0].ScheduleReason='DEMO - Missing completion schedule: choose an explicit date'
        $s.Data.ExoRows[1].Selected=$false;$s.Data.ExoRows[1].DateConfirmed=$false;$s.Data.ExoRows[1].DateSource='Batch proposal (individual override unverified)'
        $s.Data.ExoRows[2].Selected=$false;$s.Data.ExoRows[2].Exclusion='Already notified (Force send required)'
        $s.Data.ExoRows[3].Selected=$false;$s.Data.ExoRows[3].DateConfirmed=$false;$s.Data.ExoRows[3].ScheduleReason='DEMO - Conflicting dates across batches'
    }
    if ($s.Scenario -eq 'Completed') {foreach($row in $s.Data.ExoRows){$row.Selected=$false;$row.MigrationStatus='Completed';$row.Exclusion='Migration completed'}}
    $c.ExoBatchesGrid.ItemsSource=@($s.Data.ExoBatches | Where-Object {$c.ExoShowCompletedCheck.IsChecked -or $_.Status -ne 'Completed'})
    $c.ExoRecipientsGrid.ItemsSource=$s.Data.ExoRows
    foreach($name in @('ExoRefreshButton','ExoLoadButton','ExoCancelButton','ExoApplyDateButton','ExoUseSelectionButton')){$c[$name].IsEnabled=$true}
}
function script:Connect-SmartM365DemoProfileEvents {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $c.MigrationRecipientSourceCombo.Add_SelectionChanged({
        $s=$script:SmartM365Demo
        if($s.Controls.MigrationRecipientSourceCombo.SelectedIndex -eq 1){$s.Controls.MainTabs.SelectedItem=$s.Controls.ExoMigrationTab}
    })
    $c.ExoShowCompletedCheck.Add_Checked({$script:SmartM365Demo.Controls.ExoBatchesGrid.ItemsSource=$script:SmartM365Demo.Data.ExoBatches})
    $c.ExoShowCompletedCheck.Add_Unchecked({$script:SmartM365Demo.Controls.ExoBatchesGrid.ItemsSource=@($script:SmartM365Demo.Data.ExoBatches|Where-Object Status -NE 'Completed')})
    $c.TemplateCampaignCombo.Add_SelectionChanged({Update-NotificationsDemoTemplates})
    $c.TemplateList.Add_SelectionChanged({Update-NotificationsDemoPreview})
    $s.Window.Add_ContentRendered({Update-NotificationsDemoPreview})
    $c.MainTabs.Add_SelectionChanged({
        $s=$script:SmartM365Demo;$c=$s.Controls;$selected=$c.MainTabs.SelectedItem
        $campaign=if($selected -eq $c.ExchangeMigrationCampaignTab){'ExchangeMigration'}elseif($selected -eq $c.ExchangeArchiveCampaignTab){'ExchangeArchive'}elseif($selected -eq $c.ExchangeSizeReductionCampaignTab){'ExchangeMigrationMailboxSizeReduction'}else{''}
        if($selected -eq $c.ExoMigrationTab){$c.MigrationRecipientSourceCombo.SelectedIndex=1;return}
        if(-not $campaign){return}
        if($selected.Content -ne $c.RunCampaignContent){$parent=$c.RunCampaignContent.Parent;if($parent -is [Windows.Controls.ContentControl]){$parent.Content=$null};$selected.Content=$c.RunCampaignContent}
        $s.Data.Campaign=$campaign;$c.CampaignTitleText.Text=$selected.Header+' - Contoso demo';$c.CampaignBadgeText.Text=$campaign+' | DEMO'
        $c.TemplateCampaignCombo.SelectedItem=$campaign
    })
    $c.ForceLanguageCombo.Add_SelectionChanged({
        $s=$script:SmartM365Demo;$c=$s.Controls;$language=[string]$c.ForceLanguageCombo.SelectedItem
        $row=@($c.TemplateList.ItemsSource|Where-Object Language -EQ $language|Select-Object -First 1)
        if($row.Count){$c.TemplateList.SelectedItem=$row[0]}
    })
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    switch($Name){
        'ExoConnectButton' {$c.ExoConnectionText.Text='DEMO - Connected as demo.operator@contoso.example | Fictional Contoso tenant | No real connection';return $true}
        'ExoRefreshButton' {$c.ExoBatchesGrid.ItemsSource=@($s.Data.ExoBatches|Where-Object {$c.ExoShowCompletedCheck.IsChecked -or $_.Status -ne 'Completed'});return $true}
        'ExoLoadButton' {$c.ExoRecipientsGrid.ItemsSource=$s.Data.ExoRows;$c.ExoConnectionText.Text='DEMO - 12 unique recipients loaded; review each date and exclusion';return $true}
        'ExoCancelButton' {$c.ExoRecipientsGrid.ItemsSource=@();$c.ExoConnectionText.Text='DEMO - Disconnected; no authentication or delivery occurred';return $true}
        'ExoApplyDateButton' {
            if($null -ne $c.ExoOverrideDate.SelectedDate){foreach($row in $s.Data.ExoRows|Where-Object Selected){$row.EffectiveDate=([datetime]$c.ExoOverrideDate.SelectedDate).ToString('yyyy-MM-dd');$row.DateSource='Operator override';$row.DateConfirmed=$true;$row.ScheduledLocal='';$row.ScheduledUtc=''}}
            $c.ExoRecipientsGrid.Items.Refresh();return $true
        }
        'ExoUseSelectionButton' {$paths=@(Export-SmartM365DemoProfileOutputs);$c.ExoConnectionText.Text='DEMO - Fictional CSV prepared. No email or Teams message was sent.';return $true}
        'RunButton' {
            $paths=@(Export-SmartM365DemoProfileOutputs)
            Write-SmartM365DemoActivity ($s.Data.Campaign+': 12 fictional recipients processed; 12 simulated messages; zero emails or Teams messages sent.')
            [void](Show-SmartM365DemoDetail -Title 'Campaign result' -Text ('DEMO - Contoso campaign completed. All 12 sample recipients succeeded. No email or Teams message was sent. Preview files: '+($paths.Count)))
            return $true
        }
        'ClearOutputButton' {$c.OutputBox.Clear();return $true}
        'OpenLocalConfigButton' {[void](Show-SmartM365DemoDetail -Title 'Demo campaign configuration' -Text (($c.ConfigParameterList.ItemsSource|ForEach-Object {$_.Name+' = '+$_.Value}) -join [Environment]::NewLine));return $true}
        'OpenTemplateFolderButton' {$paths=@(Export-SmartM365DemoProfileOutputs);Show-SmartM365DemoFile $paths[0];return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo;$paths=[Collections.Generic.List[string]]::new()
    foreach($template in $s.Data.Templates){
        $path=Get-SmartM365DemoOutputPath ($template.Campaign+'-'+$template.FileName)
        [IO.File]::WriteAllText($path,(Get-NotificationsDemoHtml $template),[Text.UTF8Encoding]::new($false))
        [void]$paths.Add($path)
    }
    $summary=Save-SmartM365DemoHtml -Name 'Campaign-Summary-DEMO.html' -Title 'Contoso campaign summary email preview' -Body '<p>To: demo.helpdesk@contoso.example</p><table><tr><th>Campaign</th><th>Recipients</th><th>Simulated success</th><th>Messages actually sent</th></tr><tr><td>Exchange migration</td><td>12</td><td>12</td><td>0</td></tr><tr><td>Exchange archive</td><td>12</td><td>12</td><td>0</td></tr><tr><td>Mailbox size reduction</td><td>12</td><td>12</td><td>0</td></tr></table>'
    [void]$paths.Add($summary);$s.Reports.EmailPreviews=$paths.ToArray()
    $csv=Get-SmartM365DemoOutputPath 'DEMO-MigrationSelection.csv'
    $s.Data.ExoRows|Where-Object Selected|Select-Object EmailAddress,BatchIds,MigrationStatus,EffectiveDate,DateSource,ScheduledUtc|Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
    [void]$paths.Add($csv)
    $body='<p>DEMO - Fictional EXO batches and selected migration recipients. No tenant connection or messages.</p><table><tr><th>Email</th><th>Batch</th><th>Mail date</th><th>Schedule source</th><th>Review</th></tr>'
    foreach($row in $s.Data.ExoRows){$body+='<tr><td>'+[Net.WebUtility]::HtmlEncode($row.EmailAddress)+'</td><td>'+[Net.WebUtility]::HtmlEncode($row.BatchIds)+'</td><td>'+[Net.WebUtility]::HtmlEncode($row.EffectiveDate)+'</td><td>'+[Net.WebUtility]::HtmlEncode($row.DateSource)+'</td><td>'+[Net.WebUtility]::HtmlEncode($row.Exclusion)+'</td></tr>'}
    $body+='</table>'
    [void]$paths.Add((Save-SmartM365DemoHtml -Name 'DEMO-MigrationSelection.html' -Title 'Contoso EXO migration selection - DEMO' -Body $body))
    return $paths.ToArray()
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKGKQzQQExmSJs
# 04ks5jIGgJsPMYQwDS97WPq78Z2fn6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM3YJR+y8G2PpKdmw23bHG1Ssfb5rAGE9onBvRS/Xj4XMA0GCSqG
# SIb3DQEBAQUABIIBgBHAjdiM+Fi3XyFbHdZsncisSIaJLYRW/d87aL3VNAQQI+T/
# Wk/bHvhWiaYhkTWcc2Yhxft0k81T3meMZB737NueMs21y9M8x7F6uYwBii4EnEqh
# h9jDGLZDTfcgJ9wR/1DKBDT1vRckilCtgrDwexPbA9qlmQ+xKORc8cmeDa4SeTtP
# xY+MysyDTROac+ejnWbQmabz4RsFSJ5HGKKiidO/YC4miahNZd5HsI/kzHz1Txyl
# 5GJYYemvWQis5xsRaDeCgxc+FnRvyXYHamUIgLrarWDYj/9zxj6BzRasxoZXcA0L
# DaPhZqRjGQ8Ds1rCZqwbP+RB+CwbDHqaK4PncYI3w7ZZdWHCNbOd7xTXelB7TFJd
# BN0ZNygfITKqay6f6t553+VQEa6+V6cOPgJ381FsW/kQ/gokVc7hCpo5RJ2KY/8w
# rJm0t1k2ZrUwgMV59uKWpWwI4/daGLgj8mQsQ3R2wln/vGJjiiXdOWf2RMUnBCmc
# TwxQK8u51FP1+FaTcqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzUz
# NTNaMC8GCSqGSIb3DQEJBDEiBCBIEvM0IYMSlep4MoXSpWrCDZu/OV/7z+dfz1V+
# H6lYYDANBgkqhkiG9w0BAQEFAASCAgB+GMXxFWWxlLL7+vUly0ditwSsZC6Eg5QW
# +Kjdr1AMxEhtGRnEFXY1ZIA854ml9bEEUdWlOMdoNar30MqTvmTgfRoiKHQdWCtj
# ZK0VlggAqCzBtTGy6Y++EccDyrRXOaktJULU444LEEyU/eozXcPV1e2xt+PL4Fzu
# pO4SMASwIL9KOr/EEpNDT2up3U9T5F5CRTXjRe7g0XQ84L8sSxqhbW/SU9W7UH6O
# 51KmMUeYrf6r+CKk6pHubvuGtJ0pjYx4xefz5fw38+0lb8J473lgqik17uytqyGN
# XQ2/RiXO5tfTuRZqrymJcWFyxITZFcXkkImNwaii0/69OnuUiLeLGYdYkumaE+6e
# bRkcrJq8/f2P0P4bDxFNwpOOT1OzOu1KcbKbwjutVvbG1hL14hs7GcKVO5AACxOB
# Lvtyr8ZxtsVPDWSX/hPavkdaOFS9gTUPriI06FLUqhLOX7c3ST7B1g6+QPHQacN5
# 5whoy3s2PPxj7akyc2650dlE7DpX/u0GNHHYvsqu8jlONBpmIwb0BWjEiX4LLNup
# wYp1R2yUF3Xqf6JZMBG3XJ4jRH2FgNHg1DTyidTlZzRIuH4AWyL45CbXNnXzYmZq
# f9oU5dcYW4LtG9oJwZ0hB/VdsC9ty5yvieQGXI4aTTWIAbbuOMMLu+Di9daG8hF+
# KwF39oco6w==
# SIG # End signature block
