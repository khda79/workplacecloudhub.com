#Requires -Version 5.1
<#
.SYNOPSIS
Fictional diagnostic evidence covering every Analyzer tab and native HTML section.
.VERSION
1.0.0
#>
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $device=New-SmartM365DemoRow @('ComputerName','IPAddress','OSVersion','Proxy','LastUser','ImeVersion') -Values @{ComputerName='CTSO-LAP-0142';IPAddress='192.0.2.42';OSVersion='Windows 11 Enterprise 24H2';Proxy='Direct (sample)';LastUser='user01@contoso.example';ImeVersion='1.85.100.0'}
    $issues=@([pscustomobject]@{Severity='Warning';Category='Windows Update';Area='Windows Update';Title='Restart pending';Detail='A fictional update requires a restart.';Details='A fictional update requires a restart.';Recommendation='Schedule a restart (demo only).';Source='Synthetic evidence';TechnicalDetail='DEMO - RebootRequired=True in the generated fixture.'})
    $issues+=@([pscustomobject]@{Severity='Error';Category='Application';Area='Applications';Title='Sample application retry';Detail='A fictional application needs another attempt.';Details='A fictional application needs another attempt.';Recommendation='Review the sample application log.';Source='Synthetic evidence';TechnicalDetail='DEMO - No application was evaluated.'},[pscustomobject]@{Severity='Warning';Category='Policy';Area='MDM';Title='Sample policy conflict';Detail='A fictional legacy policy overlaps MDM.';Details='A fictional legacy policy overlaps MDM.';Recommendation='Review the sample policy assignment.';Source='Synthetic evidence';TechnicalDetail='DEMO - No policy was read.'})
    if($s.Scenario -eq 'Completed'){$issues=@()}
    $policies=@([pscustomobject]@{Area='DeviceLock';Policy='MinDevicePasswordLength';Value='12'},[pscustomobject]@{Area='Update';Policy='AllowAutoUpdate';Value='4'})
    $events=@([pscustomobject]@{Log='DeviceManagement-Enterprise-Diagnostics-Provider';Level='Warning';TimeCreated='2026-10-18 09:25';Id=813;Provider='MDM';Message='DEMO - Sample policy evaluation warning.';File='Demo/MDM.evtx'})
    $apps=@([pscustomobject]@{Name='Microsoft 365 Apps';Version='16.0.18000';Publisher='Microsoft';InstallDate='2026-10-01';Arch='x64'},[pscustomobject]@{Name='Contoso Support';Version='2.4.0';Publisher='Contoso';InstallDate='2026-10-04';Arch='x64'})
    $drivers=@([pscustomobject]@{OriginalName='contoso-network.inf';Provider='Contoso Demo';ClassName='Net';DriverVersion='12.3.4';Signer='Demo fixture'})
    $health=@([pscustomobject]@{Area='Enrollment';Severity='Information';Title='Enrollment observed';Details='Fictional Intune enrollment signals are present.';Recommendation='No action in this sample.';Source='Synthetic evidence'})
    $coverage=@([pscustomobject]@{Area='MDM';State='Available';Detail='Synthetic diagnostics fixture'},[pscustomobject]@{Area='Events';State='Available';Detail='Synthetic event fixture'})
    $components=@([pscustomobject]@{Category='Windows Update';EvidenceCount=1;RawPenalty=8;Cap=20;AppliedPenalty=8;Rationale='Fictional pending restart.'})
    $ime=@([pscustomobject]@{Severity='Warning';Timestamp='2026-10-18 09:24';Theme='Application';Category='Detection';ErrorCode='0x00000000';KnownCode='Success';Message='DEMO - Application detection completed.';SourceFile='Demo/IntuneManagementExtension.log'})
    $battery=[pscustomobject]@{ComputerName='CTSO-LAP-0142';SystemProduct='Contoso Laptop';BIOS='1.12';DesignCapacity='56000 mWh';FullChargeCapacity='51520 mWh';HealthPct=92;CycleCount=120}
    $firewall=@([pscustomobject]@{Name='Domain';State='ON';FirewallPolicy='BlockInbound,AllowOutbound';RemoteManagement='Disabled';LogAllowed='Yes';LogDropped='Yes';LogFileName='Demo/Firewall.log'})
    $certs=@([pscustomobject]@{Status='Valid (sample)';Subject='CN=CTSO-LAP-0142';Issuer='CN=Contoso Demo CA';NotAfter='2027-10-18';DaysToExpiry=365;Store='LocalMachine/My (synthetic)'})
    $a=[pscustomobject]@{
        ZipInfo=[pscustomobject]@{ZipName='Contoso-CTSO-LAP-0142-DEMO.zip';TotalFiles=42}
        DeviceSummary=@($device);CriticalIssues=$issues;Applications=$apps;Drivers=$drivers;ImeEvents=$ime
        Insights=[pscustomobject]@{ScoreDisplay='92 / 100 (DEMO)';Status='Attention';Coverage=[pscustomobject]@{State='Synthetic';Detail='All evidence shown was generated, never collected.';Areas=$coverage};ScoreComponents=$components;TopActions=$issues}
        Compliance=[pscustomobject]@{OverallStatus='DEMO - Pending restart';PolicyStatuses=@([pscustomobject]@{Area='Update';Status='Pending';Details='Fictional restart pending.';SourceFile='Demo/MDM.html'})}
        ErrorSummary=[pscustomobject]@{ErrorCount=1;WarningCount=2}
        WindowsUpdate=[pscustomobject]@{Info=@{TargetReleaseVersion='24H2';UpdateRing='Contoso Pilot';RestartPending='True (sample)'};ReportingGroups=@([pscustomobject]@{Level='Warning';FirstSeen='2026-10-18 09:20';LastSeen='2026-10-18 09:25';Occurrences=2;Source='Synthetic WU';ErrorCode='0';Meaning='Restart pending';Message='DEMO - Update installed.';Recommendation='Restart in the demo.'});EtlStatus='Synthetic log available';GeneratedLogPath='Demo/WindowsUpdate.log';EtlEvents=@()}
        Hardware=[pscustomobject]@{Battery=$battery;FirewallProfiles=$firewall;Certificates=$certs}
        Health=[pscustomobject]@{Findings=$health}
        MdmDiagnostics=[pscustomobject]@{CabStatus='DEMO - Synthetic CAB content loaded';DeviceInfo=@{ComputerName='CTSO-LAP-0142';OS='Windows 11 24H2'};ConnectionInfo=@{Tenant='Contoso (DEMO)';Enrollment='Intune (sample)'};ManagedPolicies=$policies;BlockedGps=@([pscustomobject]@{Path='Demo/Policies';Name='Legacy policy conflict (sample)'});Issues=$issues}
        EventLogs=[pscustomobject]@{Summary=@([pscustomobject]@{Log='MDM (sample)';Critical=0;Error=1;Warning=2;Information=15;Other=0;Scanned=18});Events=$events}
        Win11Compatibility=[pscustomobject]@{Indicators=@([pscustomobject]@{TargetVersion='Windows 11 24H2';IsBlocking=$false;ReasonText='Compatible in this fictional scenario.';SourceFile='Demo/Readiness.xml'})}
    }
    $s.Data.Analysis=$a
    foreach($name in @('BannerComputer','SummaryComputer')){$c[$name].Text='CTSO-LAP-0142'}
    $c.BannerUser.Text='user01@contoso.example';$c.BannerAccount.Text='CONTOSO\user01';$c.BannerOS.Text=$device.OSVersion;$c.BannerPowerShell.Text='PowerShell - Demo'
    $c.SummaryIP.Text='192.0.2.42';$c.SummaryOS.Text=$device.OSVersion;$c.SummaryErrors.Text='1';$c.SummaryWarnings.Text='2';$c.SummaryFiles.Text='42'
    $c.TxtFile.Text=$a.ZipInfo.ZipName;$c.SummaryText.Text='DEMO - Synthetic diagnostic loaded. MDM evidence present; a sample update requires a restart. No local device or ZIP was inspected.'
    Set-SmartM365DemoGrid SummaryDeviceGrid @($device) @('ComputerName','IPAddress','OSVersion','LastUser')
    Set-SmartM365DemoGrid SummaryConnectionGrid @([pscustomobject]@{Signal='AzureAdJoined';Value='YES (sample)'},[pscustomobject]@{Signal='MDM enrolled';Value='YES (sample)'}) @('Signal','Value')
    Set-SmartM365DemoGrid AppsGrid $apps @('Name','Version','Publisher','InstallDate','Arch')
    Set-SmartM365DemoGrid DriversGrid $drivers @('OriginalName','Provider','ClassName','DriverVersion','Signer')
    Set-SmartM365DemoGrid BatteryGrid @($battery) @('ComputerName','DesignCapacity','FullChargeCapacity','HealthPct','CycleCount')
    Set-SmartM365DemoGrid FirewallGrid $firewall @('Name','State','FirewallPolicy','LogAllowed','LogDropped')
    Set-SmartM365DemoGrid CertGrid $certs @('Status','Subject','Issuer','NotAfter','DaysToExpiry')
    Set-SmartM365DemoGrid MdmPoliciesGrid $policies @('Area','Policy','Value')
    Set-SmartM365DemoGrid EventLogGrid $events @('Level','TimeCreated','Id','Provider','Message')
    Set-SmartM365DemoGrid ImeGrid $ime @('Severity','Timestamp','Theme','Category','Message')
    Set-SmartM365DemoGrid InsightsGrid $issues @('Severity','Title','Detail','Recommendation')
    Set-SmartM365DemoGrid HealthGrid $health @('Area','Severity','Title','Details','Recommendation')
    Set-SmartM365DemoGrid ComplianceGrid $a.Compliance.PolicyStatuses @('Area','Status','Details','SourceFile')
    Set-SmartM365DemoGrid Win11Grid $a.Win11Compatibility.Indicators @('TargetVersion','IsBlocking','ReasonText','SourceFile')
    Set-SmartM365DemoGrid DeviceGrid @($device) @('ComputerName','IPAddress','OSVersion','LastUser','ImeVersion')
    Set-SmartM365DemoGrid MdmDeviceGrid @([pscustomobject]@{Property='ComputerName';Value='CTSO-LAP-0142'},[pscustomobject]@{Property='Enrollment';Value='Intune (sample)'}) @('Property','Value')
    Set-SmartM365DemoGrid MdmConnectionGrid @([pscustomobject]@{Property='Tenant';Value='Contoso (DEMO)'},[pscustomobject]@{Property='AzureAdJoined';Value='YES (sample)'}) @('Property','Value')
    Set-SmartM365DemoGrid MdmBlockedGpsGrid $a.MdmDiagnostics.BlockedGps @('Path','Name')
    Set-SmartM365DemoGrid MdmUnmanagedGrid @([pscustomobject]@{Area='DemoPolicy';Policy='Sample unmanaged setting';Value='Not assigned (sample)'}) @('Area','Policy','Value')
    Set-SmartM365DemoGrid MdmLapsGrid @([pscustomobject]@{Policy='Windows LAPS';Status='Configured (sample)';BackupDirectory='Entra (sample)';Password='Not included in demo'}) @('Policy','Status','BackupDirectory','Password')
    Set-SmartM365DemoGrid MdmSourcesGrid $coverage @('Area','State','Detail')
    Set-SmartM365DemoGrid MdmCabFilesGrid @([pscustomobject]@{Name='MDMDiagReport.html';Category='MDM report';SizeBytes=18240;Source='Synthetic CAB'}) @('Name','Category','SizeBytes','Source')
    Set-SmartM365DemoGrid WuInfoGrid @([pscustomobject]@{Property='Update ring';Value='Contoso Pilot'},[pscustomobject]@{Property='Target release';Value='24H2'}) @('Property','Value')
    Set-SmartM365DemoGrid WuPoliciesGrid @([pscustomobject]@{Policy='TargetReleaseVersion';Value='24H2';Source='Intune (sample)'}) @('Policy','Value','Source')
    Set-SmartM365DemoGrid WuHistoryGrid @([pscustomobject]@{Title='Contoso sample cumulative update';Result='Installed (sample)';Date='2026-10-18 09:20';Restart='Pending (sample)'}) @('Title','Result','Date','Restart')
    Set-SmartM365DemoGrid WuEventsGrid $a.WindowsUpdate.ReportingGroups @('Level','FirstSeen','Occurrences','Message','Recommendation')
    Set-SmartM365DemoGrid EventLogSummaryGrid $a.EventLogs.Summary @('Log','Critical','Error','Warning','Information','Scanned')
    Set-SmartM365DemoGrid AutopatchGrid @([pscustomobject]@{Name='Contoso Pilot';Status='Registered (sample)';Ring='Test';LastSync='2026-10-18 09:20'}) @('Name','Status','Ring','LastSync')
    Set-SmartM365DemoGrid CollectionErrorsGrid @([pscustomobject]@{Component='Demo fixture';Status='Synthetic';Message='No local collection was performed.'}) @('Component','Status','Message')
    Set-SmartM365DemoGrid WingetGrid @([pscustomobject]@{Name='Contoso Support';Id='Contoso.Demo.Support';Version='2.4.0';Available='2.4.1 (sample)';Source='Demo'}) @('Name','Id','Version','Available','Source')
    Set-SmartM365DemoGrid WifiGrid @([pscustomobject]@{Name='Contoso-Corporate-DEMO';Authentication='WPA2-Enterprise (sample)';Signal='92% (sample)';Address='192.0.2.42';Credentials='Not collected'}) @('Name','Authentication','Signal','Address','Credentials')
    Set-SmartM365DemoGrid EnrollmentsGrid @([pscustomobject]@{Provider='Microsoft Intune (sample)';State='Enrolled (sample)';User='user01@contoso.example';Date='2026-10-01';Device='CTSO-LAP-0142'}) @('Provider','State','User','Date','Device')
    $zipRows=@(1..42|ForEach-Object {[pscustomobject]@{Name=('Evidence-{0:00}.txt' -f $_);Path=('Demo/Diagnostics/Evidence-{0:00}.txt' -f $_);Category=$(if($_%2){'Diagnostics'}else{'Logs'});SizeBytes=(1024+$_*20);Modified='2026-10-18 09:30';Text='DEMO - Synthetic Contoso evidence.'}})
    Set-SmartM365DemoGrid ZipFilesGrid $zipRows @('Name','Path','Category','SizeBytes','Modified')
    Set-SmartM365DemoGrid SearchGrid $zipRows @('Path','Category','Text')
    Set-SmartM365DemoCombo EventLogSelector (@('All')+@($events.Log|Select-Object -Unique))
    Set-SmartM365DemoCombo ImeThemeFilter @('All','Application')
    Set-SmartM365DemoCombo ZipCategoryFilter @('All','Diagnostics','Logs')
    foreach($name in @('MdmIssuesText','EventDetailText','ImeText','InsightsText','ComplianceText','EnrollmentsText','ZipDetailText','SearchDetailText','Win11Text','WuInfoText')){$c[$name].Text='DEMO - Fictional Contoso evidence. Select a sample row to inspect its details; no source was collected.'}
    $c.WuStatusText.Text='DEMO - Fictional Windows Update evidence loaded'
    $c.MdmCabStatus.Text=$a.MdmDiagnostics.CabStatus;$c.BatteryHealthText.Text='92% (sample)';$c.HealthBadge.Text='DEMO - Healthy';$c.ComplianceBadge.Text='DEMO - Pending restart';$c.Win11StatusText.Text='Windows 11 24H2 - compatible (sample)'
    $c.AiCredential.Text='';$c.AiModel.Text='Demo analysis';$c.AiTokens.Text='2048'
    $c.AiText.Text="DEMO - Simulated AI explanation`r`nThe fictional device is enrolled and its firewall is enabled. A Windows Update restart is pending. Verify the update state after a scheduled restart. No AI provider was called."
    foreach($name in @('MdmPolicyCount','EventCount','ImeCount','AppsCount','ZipCount','SearchCount')){$c[$name].Text='DEMO - sample records loaded'}
    if($s.Scenario -eq 'Completed'){$a.ErrorSummary.ErrorCount=0;$a.ErrorSummary.WarningCount=0;$c.SummaryErrors.Text='0';$c.SummaryWarnings.Text='0';$a.Compliance.OverallStatus='DEMO - Ready';$c.ComplianceBadge.Text='DEMO - Ready'}
}
function script:Update-AnalyzerDemoFilters {
    $s=$script:SmartM365Demo;$c=$s.Controls
    foreach($spec in @(@('EventLogGrid','EventSearch','EventLogSelector','Log','EventLevelFilter','Level'),@('ImeGrid','ImeSearch','ImeThemeFilter','Theme','ImeLevelFilter','Severity'),@('ZipFilesGrid','ZipSearch','ZipCategoryFilter','Category','',''))){
        $query=$c[$spec[1]].Text;$choice=$c[$spec[2]].SelectedItem
        if($choice -is [Windows.Controls.ComboBoxItem]){$choice=$choice.Content}
        $level=if($spec[4]){$c[$spec[4]].SelectedItem}else{'All'}
        if($level -is [Windows.Controls.ComboBoxItem]){$level=$level.Content}
        $c[$spec[0]].ItemsSource=@($s.Rows[$spec[0]]|Where-Object {($_.PSObject.Properties.Value -join ' ') -like ('*'+$query+'*') -and (-not $choice -or $choice -eq 'All' -or $_.($spec[3]) -eq $choice) -and (-not $level -or $level -eq 'All' -or $_.($spec[5]) -eq $level)})
    }
}
function script:Connect-SmartM365DemoProfileEvents {
    $s=$script:SmartM365Demo
    foreach($pair in @(@('AppsSearch','AppsGrid'),@('EventSearch','EventLogGrid'),@('ImeSearch','ImeGrid'),@('MdmPolicyFilter','MdmPoliciesGrid'),@('ZipSearch','ZipFilesGrid'),@('SearchFilter','SearchGrid'))){
        $s.Controls[$pair[0]].Tag=$pair[1]
        $s.Controls[$pair[0]].Add_TextChanged({param($box)
            $s=$script:SmartM365Demo;$grid=[string]$box.Tag;$query=$box.Text
            $s.Controls[$grid].ItemsSource=@($s.Rows[$grid]|Where-Object {($_.PSObject.Properties.Value -join ' ') -like ('*'+$query+'*')})
        })
    }
    foreach($name in @('EventLogSelector','EventLevelFilter','ImeThemeFilter','ImeLevelFilter','ZipCategoryFilter')){$s.Controls[$name].Add_SelectionChanged({Update-AnalyzerDemoFilters})}
    foreach($name in @('EventSearch','ImeSearch','ZipSearch')){$s.Controls[$name].Add_TextChanged({Update-AnalyzerDemoFilters})}
    foreach($pair in @(@('EventLogGrid','EventDetailText'),@('ImeGrid','ImeText'),@('InsightsGrid','InsightsText'),@('ComplianceGrid','ComplianceText'),@('EnrollmentsGrid','EnrollmentsText'),@('ZipFilesGrid','ZipDetailText'),@('SearchGrid','SearchDetailText'))){$grid=$s.Controls[$pair[0]];$grid.Tag=$pair[1];$grid.Add_SelectionChanged({param($grid)if($grid.SelectedItem){$script:SmartM365Demo.Controls[[string]$grid.Tag].Text='DEMO - Sample data'+[Environment]::NewLine+(($grid.SelectedItem.PSObject.Properties|ForEach-Object {$_.Name+': '+$_.Value}) -join [Environment]::NewLine)}})}
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    switch($Name){
        'BtnExportHtml' {$paths=@(Export-SmartM365DemoProfileOutputs);Show-SmartM365DemoFile $paths[0];return $true}
        'BtnRunAi' {Write-SmartM365DemoActivity 'Fictional AI explanation displayed. No API, Ollama or credential was used.';return $true}
        'BtnAnalyzeLocal' {Initialize-SmartM365DemoProfile;Write-SmartM365DemoActivity 'Synthetic diagnostic loaded; local collection and elevation were skipped.';return $true}
        'BtnAnalyzeZip' {Initialize-SmartM365DemoProfile;Write-SmartM365DemoActivity 'Synthetic ZIP analysis loaded. No user ZIP was opened.';return $true}
        'BtnExtractCab' {Write-SmartM365DemoActivity 'Synthetic CAB contents are already loaded. expand.exe was not started.';return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo
    return @(Export-SmartM365DemoNativeReport -Kind Analyzer -SourcePath $s.GuiPath -Name 'EndpointDiagnostics-DEMO.html' -Data $s.Data.Analysis)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATm1zpu+OaCJR/
# x+k7ZfOh2qBSW/HdgKevyirLXdqLKqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINvFdWqOnmGZbKW+TF2dJ9QCShhtnDzzHNpFjg5tQigXMA0GCSqG
# SIb3DQEBAQUABIIBgJTOSiO8Z9SCVFzA906Yhf4GoWPoUrjzIyVatHszUkY1KOHT
# LrnvCDVszI84KLFu5ge/o5ATW7dgg0jy7BEApcBac1u7cyh0/9AdLfSntFQaIOB9
# JYANsml9A9MTbNEcSBRQbfS7kLeO2KD3F6dUdntJdvtxhKGtqkg0xts1OCs3EF3G
# WMTKk1kA4Y2uNJttRHx1HWII+4apuwobKBiq+vgeeOSJOKjy516+GKl3pe5uCm/X
# /bmsc0iLQFIfshx2eeCucNTyYnPU2KJX1zPls1ATz5s6KNZ30dJM5sDpKvN3BFog
# 7SBHtoA2WRUl2PovsHYwTfZrj25zJdic6bzDHs6gSFcJ3PoLllMtn6uVMooxnIDK
# DvgccnSKZlBU1I9PCh9KpMb9t79QSfBW7dr5c8kh5SRRVFLpEoaZVWRw6sv6INYN
# /srh7TgmkvKj//XrrvZMPKyf/rNoYBZcVsGMZdCbkHUoJx8wZh5uo7AxVJKyT+Qy
# 0zGCQyQSvLFaFhgBaqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMDQ2
# MzlaMC8GCSqGSIb3DQEJBDEiBCBW/O7Faxjaz/Yw8xNDlFgsIeiJGHoZs1/tDDtS
# ycdptzANBgkqhkiG9w0BAQEFAASCAgB3QleMs2uMe4dRKVawli3+Vl59DckzwPw7
# vaYPtVw9qC61tReeOBe9X+knAib1snw3JIZ1JYSlATQ1TMRkE7rwhttfCRx1HBjI
# DflkgUl4eUYMGFYdbpii0V/x/jw65RmKhnZCVcj5rWhZtruH4eHZ1VRSIIyXU39P
# PwWUOXxBeZojPJo5F/HUfaXTRNpa2XfaHwUgEoW5ZalqQyui4U/uL2kknUZKiyOy
# +TmH1SUPQi6Hy0X9J+6jmZLE+tY/G3ZkYbbiCctLWiTr8aw9xHTKOCj1TO3GtKfY
# VDJDMXZakElSLfF/8iazg8SadE2QIslY/PcCJHQcfXN0OE5V3XleRXNlIVLtnN8b
# e8Z25CfhfV8UwWimiZB6BrIGYH6WJmzQaIAZwSfr6aGqYCIeKh/eEifuuaPq2KSO
# 3+tNA9Ul0qav4cqkycjIYuSimsYbb7RDN4I3Tq9deHFawdSGDAWFfOrxFktjGzfE
# 6IeSXZo1agCPlhGwi2iMA1cVmeMAXrTh2TG7w39YzLG/L+cG9+1UP3cVsTDUS194
# YTc5k9jJqy+J1UP8dzS8KFDszJ2IOGHBrXArRPf7PFVqiH7i/bf8LrIQ5QYV3cm4
# Jm2CaEyGmxxEsWVOa1ZwkROgfa+WhevXXFY2r4cZ/bxNKeBpVdl1Xm6oqxUar/LT
# UbnuofbXsg==
# SIG # End signature block
