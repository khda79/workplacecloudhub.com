#Requires -Version 7.0
<#
.SYNOPSIS
Complete assessment options and fictional mailbox evidence for the readiness GUI.
.VERSION
1.0.0
#>
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $modulePath=Join-Path $s.AppRoot 'SmartM365.ExchangeMigrationReadiness.psm1'
    . ([scriptblock]::Create((Get-SmartM365DemoFunctionText $modulePath @('Get-SemrCheckCatalog')) -join [Environment]::NewLine))
    $options=@(Get-SemrCheckCatalog | ForEach-Object {[pscustomobject]@{Enabled=$true;CanDisable=$true;Name=$_.Name;Category=$_.Category;Description=$_.Description;CheckId=$_.CheckId}})
    $s.Data.Options=$options;$c.OptionsItemsControl.ItemsSource=$options
    $c.OptionsSummaryText.Text=('{0} assessment checks available - demo switches affect memory only.' -f $options.Count)
    $summary=@(for($i=0;$i -lt 12;$i++){
        $decision=if($s.Scenario -eq 'Completed'){'GO'}else{@('GO','GO','GO-WARNING','NO-GO','UNKNOWN','GO')[$i%6]}
        New-SmartM365DemoRow -Columns @('EmailAddress','UserPrincipalName','MailboxSizeGb','TargetLicense','AssignedLicenses','Decision','BlockingCount','WarningCount','UnknownCount','DataCoverage','BlockingCodes','RecommendedAction') -Index $i -Values @{
            EmailAddress=('user{0:00}@contoso.example' -f ($i+1));UserPrincipalName=('user{0:00}@contoso.example' -f ($i+1));MailboxSizeGb=(12+$i*2)
            TargetLicense='Microsoft 365 E3';AssignedLicenses='Microsoft 365 E3';Decision=$decision
            BlockingCount=[int]($decision -eq 'NO-GO');WarningCount=[int]($decision -eq 'GO-WARNING');UnknownCount=[int]($decision -eq 'UNKNOWN')
            DataCoverage='DEMO';BlockingCodes=$(if($decision -eq 'NO-GO'){'DEMO_ArchiveQuota'}else{''})
            RecommendedAction=$(if($decision -eq 'GO'){'Ready in this sample scenario.'}else{'Review the fictional mailbox finding before migration.'})
        }
    })
    $globals=@(New-SmartM365DemoRow -Columns @('CheckId','Category','Result','Message','RecommendedAction') -Values @{CheckId='DEMO_HybridEndpoint';Category='Hybrid';Result='PASS';Message='Contoso hybrid endpoint is available (synthetic).';RecommendedAction='No action in demo.'})
    $findings=@(foreach($row in $summary){[pscustomobject]@{EmailAddress=$row.EmailAddress;CheckId='DEMO_MailboxReadiness';Category='Mailbox';Result=$(if($row.Decision -eq 'NO-GO'){'FAIL'}elseif($row.Decision -eq 'UNKNOWN'){'UNKNOWN'}elseif($row.Decision -eq 'GO-WARNING'){'WARN'}else{'PASS'});Message=$row.RecommendedAction;RecommendedAction=$row.RecommendedAction;Source='Synthetic Contoso evidence'}})
    $sources=@(foreach($name in @('Active Directory','Exchange on-premises','Exchange Online','Microsoft Graph','Hybrid endpoint')){[pscustomobject]@{Source=$name;Available=$true;Status='DEMO';Required=$true;Details='Fictional source snapshot; no connection was made.';CheckedAt='2026-10-18 09:30'}})
    $s.Data.Assessment=[pscustomobject]@{RunId='DEMO-20261018-093000';AssessmentStatus='DEMO';AssessmentPhase='PreCreate';Summary=$summary;GlobalFindings=$globals;Findings=$findings;LiveSources=$sources;StartedAt=[datetime]'2026-10-18T09:28:00';CompletedAt=[datetime]'2026-10-18T09:30:00';Hybrid=[pscustomobject]@{EndpointName='Contoso-Hybrid-DEMO'};EntraConnect=[pscustomobject]@{Message='Synchronization healthy (synthetic)'}}
    $c.CsvPathBox.Text='Demo/Contoso-MigrationBatch.csv';$c.CsvMetadataText.Text='12 fictional mailboxes | Contoso | DEMO'
    Set-SmartM365DemoCombo MigrationEndpointCombo @('Contoso-Hybrid-DEMO')
    Set-SmartM365DemoGrid BatchGrid $summary @('EmailAddress','UserPrincipalName','TargetLicense')
    Set-SmartM365DemoGrid SummaryGrid $summary @('EmailAddress','Decision','MailboxSizeGb','TargetLicense','BlockingCount','WarningCount','UnknownCount','RecommendedAction')
    Set-SmartM365DemoGrid GlobalFindingsGrid $globals @('CheckId','Category','Result','Message','RecommendedAction')
    Set-SmartM365DemoGrid FindingsGrid $findings @('EmailAddress','CheckId','Result','Message','RecommendedAction')
    Set-SmartM365DemoGrid PermissionsGrid @([pscustomobject]@{Mailbox='user01@contoso.example';Trustee='team.support@contoso.example';Permission='FullAccess';Source='On-prem (sample)';Target='EXO (sample)';Result='Matched'}) @('Mailbox','Trustee','Permission','Source','Target','Result')
    foreach($pair in @(@('GoCountText','GO'),@('WarningCountText','GO-WARNING'),@('NoGoCountText','NO-GO'),@('UnknownCountText','UNKNOWN'))){$c[$pair[0]].Text=[string]@($summary|Where-Object Decision -EQ $pair[1]).Count}
    foreach($name in @('AdStateText','OnPremStateText','ExoStateText','GraphStateText','HybridStateText')){$c[$name].Text='DEMO - Fictional source available. No live connection.'}
    $c.FooterText.Text='DEMO - Sample data | Readiness results are simulated';$c.ResultsTab.IsEnabled=$true
    $c.VersionText.Text='DEMO'
}
function script:Connect-SmartM365DemoProfileEvents {
    $script:SmartM365Demo.Controls.MailboxFilterBox.Add_TextChanged({
        $s=$script:SmartM365Demo;$query=$s.Controls.MailboxFilterBox.Text
        $s.Controls.FindingsGrid.ItemsSource=@($s.Rows.FindingsGrid|Where-Object {$_.EmailAddress -like ('*'+$query+'*')})
    })
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    switch($Name){
        'RunButton' {Initialize-SmartM365DemoProfile;$c.MainTabs.SelectedItem=$c.ResultsTab;$c.RunProgress.Value=100;[void](Export-SmartM365DemoProfileOutputs);Write-SmartM365DemoActivity 'Assessment simulated; native HTML report generated from fictional mailboxes.';return $true}
        'OpenOutputButton' {$paths=@(Export-SmartM365DemoProfileOutputs);Show-SmartM365DemoFile $paths[0];return $true}
        'EnableAllOptionsButton' {foreach($option in $s.Data.Options){$option.Enabled=$true};$c.OptionsItemsControl.Items.Refresh();return $true}
        'ResetOptionsButton' {Initialize-SmartM365DemoProfile;return $true}
        'ClearFilterButton' {$c.MailboxFilterBox.Text='';return $true}
        'ComparePermissionsButton' {Write-SmartM365DemoActivity 'Permissions comparison matched the fictional source and target evidence.';return $true}
        'CancelButton' {Write-SmartM365DemoActivity 'Demo assessment cancelled. No background process was started.';return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo
    $path=Export-SmartM365DemoNativeReport -Kind Readiness -SourcePath (Join-Path $s.AppRoot 'SmartM365.ExchangeMigrationReadiness.psm1') -Name 'ExchangeMigrationReadiness-DEMO.html' -Data $s.Data.Assessment
    $csv=Get-SmartM365DemoOutputPath 'ExchangeMigrationReadiness-DEMO.csv'
    $s.Data.Assessment.Summary|Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
    return @($path,$csv)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2xMMbzWSiscFs
# Y7qh77PFpYVceijbnh8H2uQ5MJ98/KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHyAnTt/AuScXaTO9T3Nde41nSJziA/40W4KZbfvBhjnMA0GCSqG
# SIb3DQEBAQUABIIBgDU7VRL5NX1Dx49djBfYeh/I60hP5g63902hg+PPm9zC1WXB
# KADzP87uK6fxP27cZqowfAUDK8ArwN7dX890p3KL0P2rZdBnX/jOA3eZ8wdCFoZN
# 9rDdEadua3M/A2I9myBBw1zk43NP3m2C42tVbIjZxxwewHYdxx9Ix76nYjn/0jjD
# 7hjCIWeiHCqACRqg+S/tuTmz3orEWiH8ZIy5CkOEGT3rs1uGjBy+Z+h/AMrTS985
# wOKmcn0OS5TPEeHLgHWc+NvQVOoYqBfVGpWgQYlC0fc0F29D1JbOOdya8Tsitbzj
# PiIu6L+rCmUAuI69wNUHpYXe42bB97UYnwP+F4VV/D7y1KkOJg6U7CnwZVPS1rGH
# RvG6SaGNVrQW6tIfPM2v0PynNAiV4VU1thVqew0jo7QF90JkiX4iwFmejM9q9h8Z
# Kor5Vq3gH+NMAFcRGaqZ/1QgyNO2NW6rmgws9tK/3WE8m/iqrVRFmQugfyANXIxP
# fXhHwzXtqLcTVcFByaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMDQ2
# MzJaMC8GCSqGSIb3DQEJBDEiBCBqCqelQ0RLvnXcu3+5bCEgjirSL0QIHrcb6ZN9
# OAXOmzANBgkqhkiG9w0BAQEFAASCAgAzzhJb7bs4cekfKfXXzkql6DkuaCguyidO
# rnOARrihywkxsHMZs1bTwdNJhlgCWNAeiztbdiZqxrU9Q0h6DrJQSr25yCh2YOR7
# iCHqYUlzmpcQiBZ/r4MagKAJEOwXLMBnr7gCJfmH0/EAPf4p1vFyrYGh6OA/mM9p
# kHEUc6+QPSqQXIXV/KcmZ3pz9q5fzQCjBy1Wi/9xgop3SMmE0o5l7XBnaZM6ecuP
# y5uvYDEWw7EcEcgJObWDrLqg4h0kTeQ5w/+XrfFygIRIxQWx8+h0mMjSH8wACvxr
# yJ2UoowvXsG4a6Yhclu9/z/HxvvJBAVfsThZ4uLkEQcLBbhnccTjKplIkDN7o2bh
# Tj7D45bdi8Of4jI0KZcuqkwr1xeOEDD5OpZwMHetE3aPXg2lql4kJXAa1RQe6Ckv
# D0P4LT7tczJgA97dCKZNjiKMZ18iV1vv4PcfEHLMGhB7BYle5c3cs4qogIlPXQu9
# G2CJIKS+gUfTGeQZmQdJlqSQ/4K0mT7ZyqgOqtrdW0M493nj5hS4RueFQfKFAwvq
# 0facJ4OEdBwvk4PbCFJP55HD5nlwOgHbu+Nwl8sPlWDY5+7ZISmJ3dP1LrFW9qNz
# Tr76y8cLVy4dPz7GN/YXF9Qqv/PYkZsJsdBx4h7elZpO3iPjPtqiDywT9u4x4GYa
# ROkflkJHpw==
# SIG # End signature block
