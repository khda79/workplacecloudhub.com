#Requires -Version 5.1
<#
.SYNOPSIS
Fictional device registration scenarios for the original user/admin interface.
.VERSION
1.4.1
#>
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $issue=$s.Scenario -eq 'Issues'
    $texts=@{
        EyebrowText='SMARTM365 | DEMO';TitleText='Smart Device Registration Tool';SubtitleText='Device-side registration and enrollment evidence - Contoso sample'
        DeviceNameText='CTSO-LAP-0142 | CONTOSO | user01@contoso.example';ActionTitleText='Diagnostics and guarded repair'
        EnrollmentBannerText=$(if($issue){'DEMO - Intune enrollment needs attention'}else{'DEMO - Device enrolled in Intune'})
        MdmPolicyBannerText='MDM auto-enrollment policy: configured (sample)'
        IssuesTitleText='Issues and next actions';InformationTitleText='Device information';DiagnosticOutputTitleText='Diagnostic output';DsregTitleText='Device registration status'
        ModeLabelText='Mode';RuntimeSettingsText='Demo state is kept in memory. No registry, task, enrollment or profile is changed.'
        ConfigurationTitleText='Effective demonstration settings';ConfigSummaryText='Organization: Contoso | Authentication: none | Device profile: Hybrid | Output: separate demo session'
        FooterText='DEMO - Sample data | Diagnostic and repair commands are simulated'
    }
    foreach($key in $texts.Keys){$c[$key].Text=$texts[$key]}
    $c.IssuesBox.Text=if($issue){"Device authentication: FAILED (sample)`r`nStale enrollment trace detected (sample).`r`nNext action: review identity and use guarded repair."}else{"No blocking issue in this fictional scenario.`r`nNo device assessment was performed."}
    $c.InfoBox.Text="DEMO - Read-only user diagnostics 1.3`r`nRegistration: Healthy | User SSO: Present | Local Intune: $(if($issue){'Unknown'}else{'Detected'})`r`nCoverage: $(if($issue){'Partial - inaccessible evidence remains unknown'}else{'Complete for this fictional sample'})`r`nPRT dates and WAM/Hello/CloudTgt/OnPremTgt: available (sample)`r`nPrior failure for another account: informational, no refresh`r`nDNS/TCP443: reachable (sample; TLS and SYSTEM not tested)`r`nComputer: CTSO-LAP-0142 | User: user01@contoso.example`r`nNo native device or tenant commands executed."
    $c.DsregOutputBox.Text="DEMO - Synthetic dsregcmd evidence`r`nAzureAdJoined : YES`r`nDomainJoined : YES`r`nDomainName : CONTOSO`r`nDeviceName : CTSO-LAP-0142`r`nDeviceId : 00000000-0000-4000-8000-000000000142`r`nTenantName : Contoso (DEMO)`r`nTenantId : 00000000-0000-4000-8000-000000000001`r`nAzureAdPrt : YES`r`nDeviceAuthStatus : $(if($issue){'FAILED'}else{'SUCCESS'})"
    $c.OutputBox.Text="DEMO - Fictional user diagnostics 1.3`r`n[User PRT / SSO]`r`nAzureAdPrt: YES | Update: 2026-11-10T08:00:00Z | Expiry: 2026-11-24T08:00:00Z`r`nWamDefaultSet: YES | NgcSet: NO | CloudTgt: YES | OnPremTgt: YES`r`nAcquirePrtDiagnostics: other@contoso.example; RelevantFailure=False`r`n[Connectivity: current user, DNS and TCP 443 only]`r`nlogin.microsoftonline.com: DNS=Resolved; TCP443=Reachable (sample)`r`n[Local MDM classification]`r`nMS DM Server: Intune (sample) | WMI_Bridge_SCCM_Server: ManagementBridge (sample)`r`n[Recent accessible events]`r`nNo matching fictional events. Inaccessible logs would remain Unknown.`r`n[Next steps - guidance only]`r`nNo repair, enrollment or PRT refresh requested. No native commands executed."
    Set-SmartM365DemoCombo ModeCombo @('User','Admin') 0
    if (-not (Get-Command Get-DrtUserViewModel -ErrorAction SilentlyContinue)) {
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($s.GuiPath,[ref]$tokens,[ref]$errors)
        foreach($fn in $ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -in @('Get-DrtUserViewStrings','Get-DrtUserViewModel')}) { Invoke-Expression $fn.Extent.Text }
    }
    $sample=[pscustomobject]@{Status='HEALTHY';UserSsoState='Present';RegistrationHealth='Healthy';LocalMdmState='Detected';DiagnosticCoverage='Complete';MdmNonIntuneEnrollmentDetected=$issue}
    $findings=if($issue){@('[WARNING] Additional management trace (fictional)')}else{@()}
    $model=Get-DrtUserViewModel -Result $sample -Findings $findings
    $c.UserResultTitle.Text='DEMO - '+$model.Title;$c.UserResultDetail.Text=$model.Detail;$c.UserNextStep.Text=$model.Next
    $c.UserAccountValue.Text=$model.Account;$c.UserDeviceValue.Text=$model.Device;$c.UserManagementValue.Text=$model.Management
    $t=Get-DrtUserViewStrings
    $c.UserNoChangesText.Text='DEMO - Sample data. No device assessment was performed.';$c.UserLocalEvidenceText.Text=$t.LocalEvidence
    $c.UserResultBanner.Background=if($issue){'#FFFAEB'}else{'#ECFDF3'}
    $c.UserCopyButton.Content=$t.Copy;$c.UserCopyButton.IsEnabled=$true;$c.RunButton.Content=$t.Check
    $c.LicenseNotice.Text='DEMO - Optional own-account Graph check. No Microsoft sign-in or network call in demos.';$c.LicenseStatus.Text='DEMO - Not checked.'
    $c.LicenseAccountBox.Text='user01@contoso.example';$c.LicenseAccountLabel.Text='Microsoft work account to check'
    $c.UserPrtStatus.Text='';$c.UserRenewPrtButton.Content='Renew my work connection';$c.UserRenewPrtButton.IsEnabled=$true
    Set-SmartM365DemoUserLayout
    foreach($name in @('TriggerJoinCheck','TriggerIntuneAutoEnrollCheck','AuditOnlyCheck')){$c[$name].IsChecked=$true}
    $c.InfoBox.Text+="`r`n[INFO] Microsoft linked enrollment (WinDC/MMP-C) recognized (sample). Presence only; synchronization and health not checked."
    $c.OutputBox.Text+="`r`n[Local MDM classification - additional Microsoft entry]`r`nMicrosoftLinkedEnrollment; Provider=Microsoft Device Management; EnrollmentType=26 (fictional). Presence only; no enrollment change."
    $s.Data.SupportSummary=$c.InfoBox.Text+"`r`n`r`n"+$c.IssuesBox.Text
}
function script:Connect-SmartM365DemoProfileEvents {
    $script:SmartM365Demo.Controls.UserRenewPrtButton.Content='Renew my work connection'
    $script:SmartM365Demo.Controls.UserRenewPrtButton.IsEnabled=$true
    $script:SmartM365Demo.Controls.UserRenewPrtButton.Add_Click({$script:SmartM365Demo.Controls.UserPrtStatus.Text='DEMO - Request sent, renewal not confirmed (simulated). No native command executed.'})
    $script:SmartM365Demo.Controls.LicenseButton.Add_Click({$script:SmartM365Demo.Controls.LicenseStatus.Text='DEMO - Unverifiable (simulated denied access). No sign-in or network call.'})
    $script:SmartM365Demo.Controls.UserDetailsExpander.Add_Expanded({$script:SmartM365Demo.Controls.UserSimplePanel.Visibility='Collapsed'})
    $script:SmartM365Demo.Controls.UserDetailsExpander.Add_Collapsed({if($script:SmartM365Demo.Controls.ModeCombo.SelectedItem -eq 'User'){$script:SmartM365Demo.Controls.UserSimplePanel.Visibility='Visible'}})
    $script:SmartM365Demo.Controls.MainTabs.Add_SelectionChanged({
        $c=$script:SmartM365Demo.Controls
        if ($c.MainTabs.SelectedItem -ne $c.SummaryTab -and $c.UserDetailsExpander.Content -eq $c.MainTabs) { $c.UserDetailsExpander.IsExpanded=$true }
    })
    $script:SmartM365Demo.Controls.ModeCombo.Add_SelectionChanged({
        $c=$script:SmartM365Demo.Controls;$admin=$c.ModeCombo.SelectedItem -eq 'Admin'
        $c.RepairButton.Visibility=if($admin){'Visible'}else{'Collapsed'}
        foreach($name in @('RepairCheck','TriggerJoinCheck','TriggerIntuneAutoEnrollCheck','RemoveStaleIntuneCheck','RemoveNonIntuneMdmCheck','AllowIntuneCheck')){$c[$name].IsEnabled=$admin}
        $c.LicenseNotice.Text='DEMO - Optional own-account Graph check. No Microsoft sign-in or network call in demos.';$c.LicenseStatus.Text='DEMO - Not checked.'
    Set-SmartM365DemoUserLayout
    })
}
function script:Set-SmartM365DemoUserLayout {
    $c=$script:SmartM365Demo.Controls;$admin=$c.ModeCombo.SelectedItem -eq 'Admin'
    if($admin){
        if($c.UserDetailsExpander.Content -eq $c.MainTabs){$c.UserDetailsExpander.Content=$null;[void]$c.ContentGrid.Children.Add($c.MainTabs)}
        $c.UserSimplePanel.Visibility='Collapsed';$c.UserDetailsExpander.Visibility='Collapsed';$c.UserCopyButton.Visibility='Collapsed';$c.OptionsTab.Visibility='Visible';$c.RefreshPrtButton.Visibility='Visible';$c.RepairButton.Visibility='Visible'
    }else{
        if($c.UserDetailsExpander.Content -ne $c.MainTabs){[void]$c.ContentGrid.Children.Remove($c.MainTabs);$c.UserDetailsExpander.Content=$c.MainTabs}
        $c.UserSimplePanel.Visibility='Visible';$c.UserDetailsExpander.Visibility='Visible';$c.UserCopyButton.Visibility='Visible';$c.OptionsTab.Visibility='Collapsed';$c.RefreshPrtButton.Visibility='Collapsed';$c.RepairButton.Visibility='Collapsed'
    }
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo
    switch($Name){
        'UserCopyButton' {Write-SmartM365DemoActivity 'Report copy simulated locally. No email or Outlook action.';return $true}
        'RepairButton' {$s.Scenario='Completed';Initialize-SmartM365DemoProfile;Write-SmartM365DemoActivity 'Guarded repair completed in memory. No dsregcmd, registry, task, enrollment or reboot operation ran.';return $true}
        'RunButton' {Initialize-SmartM365DemoProfile;Write-SmartM365DemoActivity 'Synthetic diagnostic completed.';return $true}
        'RefreshPrtButton' {Write-SmartM365DemoActivity 'PRT renewal simulated successfully.';return $true}
        'CopySummaryButton' {[void](Show-SmartM365DemoDetail -Title 'Support summary and email preview' -Text ($s.Data.SupportSummary+"`r`n`r`nTo: demo.helpdesk@contoso.example`r`nSubject: DEMO - Device registration support`r`nSending is simulated; no email is sent."));return $true}
        'CreateBundleButton' {[void](Export-SmartM365DemoProfileOutputs);[void](Show-SmartM365DemoDetail -Title 'Support bundle email preview' -Text ($s.Data.SupportSummary+"`r`n`r`nTo: demo.helpdesk@contoso.example`r`nSubject: DEMO - Device registration support bundle`r`nAttachment: DeviceRegistration-SupportBundle-DEMO.zip`r`nOnly synthetic files were added. No Outlook session was opened and no email was sent."));return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo
    $path=Get-SmartM365DemoOutputPath 'DeviceRegistration-SupportSummary-DEMO.txt'
    [IO.File]::WriteAllText($path,('DEMO - Sample data'+[Environment]::NewLine+$s.Data.SupportSummary))
    $s.Reports.SupportSummary=$path
    $eml=Get-SmartM365DemoOutputPath 'DeviceRegistration-SupportEmail-DEMO.eml'
    $message="From: demo.operator@contoso.example`r`nTo: demo.helpdesk@contoso.example`r`nSubject: DEMO - Device registration support`r`nMIME-Version: 1.0`r`nContent-Type: text/plain; charset=utf-8`r`n`r`nDEMO - Sample data. No email was sent.`r`n"+$s.Data.SupportSummary
    [IO.File]::WriteAllText($eml,$message,[Text.UTF8Encoding]::new($false))
    Add-Type -AssemblyName System.IO.Compression
    $zip=Get-SmartM365DemoOutputPath 'DeviceRegistration-SupportBundle-DEMO.zip'
    $stream=[IO.File]::Open($zip,[IO.FileMode]::Create)
    $archive=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
    try{
        foreach($file in @($path,$eml)){
            $entry=$archive.CreateEntry([IO.Path]::GetFileName($file));$writer=[IO.StreamWriter]::new($entry.Open())
            try{$writer.Write([IO.File]::ReadAllText($file))}finally{$writer.Dispose()}
        }
    }finally{$archive.Dispose();$stream.Dispose()}
    return @($path,$eml,$zip)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDeLjHJcDnluGu5
# v+kReErXc2R1EYaj0LZFkAtqvzgR0KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMYsz9cefMe5oWDVQ61z25GcVM4OlwHYVTHj5k1iUVH0MA0GCSqG
# SIb3DQEBAQUABIIBgHM3gUbRV34St9BEea7qLcZLKOmwhAcceCNMdJt4dbuszo6j
# ese+Ml1fqSFf4Tx01RK9NQzNjq145T7kJVYDE1foFo0oYEAZTFi2noXLsHq6uaHS
# hc6ZMfMB/eEw27P4BEiyvGO9oj1iRkXhFRBnSRlVNFltdwg/XRTPoNFSUh8vvtxk
# +GgG0hImkDoawFWBovqoZ/DO7W15dw/kAK/OO0r/JJUkFs3Xp9k0tiBDRizD2XQl
# 28z+REnhJL1AIQycOyoHvFqcpyTSeSP75TRoVFGIqj5JJHDPI19jvHRVblOPaK/e
# yI4naa/cDQTpJ2cJJnbVjdOI8140zPLI1fIBbIOjCZBxd9oRHQwhRzUT6Mbu76rb
# shggGXoo3mUW0YsGDR0/dzbyhrjAdZkbCP2jcz2VS7u2tOo/v6RS78ZrP7mgqmJU
# ugEchvmvM+CVBXlg+k/nWJ2FkDYxzbO/2wDA4DvOl9knCFmMmzm5Z2AYcphEqwYw
# A4s23tQet0+7fUpslKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIyMjM4
# NDRaMC8GCSqGSIb3DQEJBDEiBCDUMTV055S/JhyD5baLEltff+q8WAiJ2o7ceP+p
# WxfxdjANBgkqhkiG9w0BAQEFAASCAgCzcmn1S8NP0kYpqfLbMbGt0vUqwC8OFxCE
# VRWbMsHTNQiuRKl/RPlWXKAha31j5yS+wbDTUiIy9ZRZ4HIh40/96MXP6Dxj73ID
# FgLd8xPioqvlWUbx0d2v3l2mxByy2HbAkWb9PobqnmYlcJXL+abAoxD5ZkJiuvI6
# qgIl5bIueklhUeR6iFBdXqA3+eUc+g+QNsNXM1HDr9vOPXvjayqPZ4ZuhldPxU5H
# 2MHjD+lLA3OyEzZGAessgzLCwiTBGGep6vlggK4v5NkZAlBJRB2dD0RlPUwisCXG
# ko5FgEZxjHgfty0zRITxIJ9+Zsy01iIQnDAdfffYcb9mqt7pHjAcbTo4B3mI3iwI
# JIoP3HZCUs7wA+vT2bAlAxu9hCBiEDEIDZQ22o3MQU9r6Qz0A/v0IXmcZF/T1Gwo
# ahCElsoEKPhjjvXMq3YQMDQDrE7NN6bRwxTYKzl6ggMTPB/AQxMUbP1MhH3VHqy6
# vyucPKGZGXsHKh7xlBnlxTS4E8n7v30XiOndEUrdahGrypRnOmNGMTczuD83t2yO
# 7Ha8KsEY2/ejF/YcKk6izryzESdu8A9DFcCJZQkDWydpKAgXNeF0G7e98Q8nhPaW
# pWtDZEvUeCOPYhE4yD4TvsnE3i11dF4TtjriT2dnmxRIN1cXjkeGq0RnmxnXjxgv
# vZoJmANz1w==
# SIG # End signature block
