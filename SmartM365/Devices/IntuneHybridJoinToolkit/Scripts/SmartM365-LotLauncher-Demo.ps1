#Requires -Version 5.1
<#
.SYNOPSIS
Fictional manual/automatic LOTs, single-PC runs and native HTML reports.
.VERSION
1.0.0
#>
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $windows=$s.GuiPath -like '*Windows11*'
    $s.Data.Kind=if($windows){'Windows11'}else{'HybridJoin'}
    if(-not $s.Data.ContainsKey('Lots')){$s.Data.Lots=@('LOT-CONTOSO-PARIS','LOT-CONTOSO-LYON','LOT-CONTOSO-PILOT')}
    Set-SmartM365DemoCombo LotCombo $s.Data.Lots
    Set-SmartM365DemoCombo LotModeCombo @('Loop','Once','LoopIgnoreRunGuard','OnceIgnoreRunGuard')
    Set-SmartM365DemoCombo SingleModeCombo @('Once','OnceIgnoreRunGuard','Loop','LoopIgnoreRunGuard')
    $c.LotDeviceCountText.Text='12 sample devices';$c.LotWrappersText.Text='Demo launchers ready - no scripts will be executed'
    if($windows){$c.LotScopeText.Text='CONTOSO | Windows 10 to Windows 11';Set-SmartM365DemoCombo AutomaticSourceCombo @('AD + Intune','AD','Intune')}
    else{$c.LotAdDomainText.Text='ad.contoso.example'}
    $c.SingleComputerText.Text='CTSO-LAP-0142';$c.SingleRunFolderText.Text='Demo/Runs/CTSO-LAP-0142'
    $c.NewLotNameText.Text='LOT-CONTOSO-DEMO';$c.NewLotComputersPathText.Text='Demo/Computers.txt'
    $c.AutomaticLotNameText.Text='LOT-CONTOSO-AUTO';$c.AutomaticNamePrefixText.Text='CTSO-LAP-';$c.AutomaticNameContainsText.Text=''
    $c.AutomaticLastLogonDaysText.Text='90';$c.AutomaticExcludeStaleAdCheck.IsChecked=$true
    $c.AutomaticEvidencePathText.Text='Demo/AutomaticSelection/20261018-093000'
    $c.AutomaticSummaryText.Text="DEMO - Contoso automatic selection`r`nSource devices: 18`r`nEligible: 12`r`nExcluded: 6 (2 stale AD, 2 already managed, 2 incompatible or ambiguous).`r`nSelected: CTSO-LAP-0001 to CTSO-LAP-0012`r`nNo AD, Intune, endpoint or inventory cache was queried."
    $defaults=@{GlobalLimitText=15;GlobalLimitOptionText=15;ThrottleText=5;CycleDelayText=30;MaxCyclesText=3;PsExecTimeoutText=30;GlobalLeaseText=120;IntuneRetrySleepText=5;IntuneRetryMaxText=3;StaleCleanupDelayText=10;RebootDelayText=60;NightPauseStartText=22;NightPauseEndText=6;ComputerDelayText=2;SetupCandidateLimitText=3;SetupCopyIpgText=10;SetupCopyJitterText=30;SetupSubnetLimitText=2;SetupSubnetLeaseText=120;ForceRequiredRebootDaysText=7}
    foreach($key in $defaults.Keys){if($c.ContainsKey($key)){$c[$key].Text=[string]$defaults[$key]}}
    foreach($name in @('DryRunCheck','AuditOnlyCheck','KeepCentralHistoryCheck')){$c[$name].IsChecked=$true}
    if($windows){
        Set-SmartM365DemoCombo SetupModeCombo @('LocalCache','Share','Auto')
        Set-SmartM365DemoCombo SetupLanguageCombo @('MatchSystem','Any','fr-FR','en-GB','en-US','de-DE','es-ES','it-IT','nl-NL','pt-PT','pl-PL')
        Set-SmartM365DemoCombo SetupDynamicUpdateCombo @('Disable','Enable','NoDrivers','NoLCU','NoDriversNoLCU')
        Set-SmartM365DemoCombo SetupSubnetPrefixCombo @('Auto','20','21','22','23','24','25','26','27','28','29','30','31','32')
        $c.SetupSourceText.Text='Demo/Windows11-24H2';$c.SetupSourceMapText.Text='Demo/SetupSourceMap.csv';$c.SetupMediaIdText.Text='CTSO-W11-24H2-DEMO'
        $c.SetupSubnetGateRootText.Text='Demo/SubnetGate'
    }
    $c.ActivityText.Text="[2026-10-18 09:30:00] DEMO - Three fictional LOTs loaded.`r`n[2026-10-18 09:30:00] All options remain visible. Actions, repairs, reboots and workers are simulated."
}
function script:Connect-SmartM365DemoProfileEvents {
    $script:SmartM365Demo.Controls.LotCombo.Add_SelectionChanged({
        $c=$script:SmartM365Demo.Controls
        $c.LotDeviceCountText.Text=if($c.LotCombo.SelectedIndex -eq 2){'4 sample devices'}else{'12 sample devices'}
    })
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    switch -Regex($Name){
        '^(LaunchLotButton|LaunchAllButton|LaunchSingleButton|OpenLotReportsButton)$' {
            $paths=@(Export-SmartM365DemoProfileOutputs)
            Write-SmartM365DemoActivity ($Name+' completed in demo memory. No PsExec, worker, scheduled task, endpoint action or reboot was started.')
            Show-SmartM365DemoFile $paths[0]
            return $true
        }
        '^(CreateLotButton|AutomaticCreateButton)$' {
            $lot=if($Name -eq 'CreateLotButton'){$c.NewLotNameText.Text}else{$c.AutomaticLotNameText.Text}
            if([string]::IsNullOrWhiteSpace($lot)){$lot='LOT-CONTOSO-DEMO'}
            if($lot -notin $s.Data.Lots){$s.Data.Lots+=@($lot)}
            Set-SmartM365DemoCombo LotCombo $s.Data.Lots ($s.Data.Lots.Count-1)
            Write-SmartM365DemoActivity ('LOT created in memory: '+$lot)
            return $true
        }
        '^AutomaticPreviewButton$' {Initialize-SmartM365DemoProfile;[void](Show-SmartM365DemoDetail -Title 'Automatic LOT selection preview' -Text $c.AutomaticSummaryText.Text);Write-SmartM365DemoActivity 'Automatic preview ready - fictional source evidence only.';return $true}
        '^StopAllButton$' {Write-SmartM365DemoActivity 'All simulated LOTs stopped. No real process was contacted.';return $true}
        '^(OpenLotComputersButton|OpenNewLotComputersButton)$' {[void](Show-SmartM365DemoDetail -Title 'Computers.txt' -Text ((1..12|ForEach-Object {'CTSO-LAP-{0:0000}' -f $_}) -join [Environment]::NewLine));return $true}
        '^(OpenLotAdDomainButton|AutomaticOpenEvidenceButton)$' {[void](Show-SmartM365DemoDetail -Title 'Fictional LOT evidence' -Text $c.AutomaticSummaryText.Text);return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $s=$script:SmartM365Demo
    $source=if($s.Data.Kind -eq 'Windows11'){'SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'}else{'SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1'}
    $path=Export-SmartM365DemoNativeReport -Kind $s.Data.Kind -SourcePath (Join-Path $s.AppRoot $source) -Name ($s.Data.Kind+'-LOT-Summary-DEMO.html')
    $csv=Get-SmartM365DemoOutputPath ($s.Data.Kind+'-LOT-Summary-DEMO.csv')
    @(1..12|ForEach-Object {[pscustomobject]@{Computer=('CTSO-LAP-{0:0000}' -f $_);Status=@('SUCCESS','SUCCESS','PENDING_REBOOT','ADMIN_SHARE_UNREACHABLE')[($_-1)%4];NextAction='Review sample result';Mode='DEMO'}})|Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
    return @($path,$csv)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASo68MnJpqRJRp
# 10sDd8Uz/F+K9hSlF9CXGfl6vglNsqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBXZrlcSY/8xIm1hhTTisyxpm+V0M2i+vDIoCBVOgR2sMA0GCSqG
# SIb3DQEBAQUABIIBgDAGaJV5mFHXQhczTdyw/yD6ZSE3G90llRn2xOL7tZSmOBUY
# s5dqhdLLmDnqmT5wjEVDZu8gH7Dwim9fnbmzl3t+eNOdVqqjygGK8T4ZJCWDMDOT
# NmE1SPaC0Ja0McUX4FTxgwcNc6K3foB6+aDoAFO9AvQuTUT93Ze6Zbo+czs7YBGm
# IklPyqUxJNQuqmkwArPjaeEheUrj+9LtoCXn7CWIF3NlkpZtvsFM5cgnZ3ejribl
# QlVQeiiDZEZST/ZtokgrJ1iHRuTJjVyGNCgn/dso58FY8qygvZ8E7zthSGW3n5fk
# MiSgsqgD7qljq90JumjAFt6+iKVNNRo55JzewjYp6DGq43kKdM2AQNUi0FQwOEqZ
# xLvJ8tUp/jj5GvT0e22sPjBxjar9W8VcmyUuaNicQFnAm5BmZLvBl4HfjTdHSoeG
# ZuKUxb8Am2G8E3e3T8j2W4t/7r3NSmkwjuM2ICpuTEwPbJfN37IzS4yel5qG6Gbh
# tWEKHf90ps8S+R9NQaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMDQ2
# MzVaMC8GCSqGSIb3DQEJBDEiBCDCIxA+QJz51fOAE/ByaBZauV9aeIfQtEOyUZc6
# tc0mtjANBgkqhkiG9w0BAQEFAASCAgAR3y9ayQ7JDK4KgukcpQtkn5uHDQ/spUGo
# ih1BI5TuhXNuhoNn2J3zDvYgl3PUckcrieUSgMzGbWIt7cvq2IwBy8nQ94tCSqVv
# zm7zeBTx+M/KXLp5f2Kne2bLNiePFB/9w7bgceWRoFvipFQrjTz581C9oCRKe6UO
# XcjfyQWTxGi/hokYSz8sDz2kpa8rUKWIZc9ZjbnSVOsXY80D3ISVo3v4fdSZF5xW
# ymcauhCRdhLb91ivENM4hlMeOr1+toAMPjKBQWg3SjaFnDFawqOhPtUvojNu/neM
# sMLC0wVJU+UNUTw6G8ZOZw85hbvyiD6owBM+GsOm2pPvaXOXrY+W7S/uoxnz2adZ
# Nf1OunTEv7V6ztg1lLVI4bS7ISEEBiO3erZc3RMcmMRz6Xo8Iwx2Msr7kfsLpa19
# xKXGQBkMDr6uNbXiTOqmSZGnVVh78E4/mm3hqGCSv8i2L0jYWQxtLUs3vX3Ui+5V
# XDyh85kOUBoWeP4BNrkAXzTYwiF9lmM54k3yetNlkKzXOPuhWzYb5MtU0SGeUc2U
# 4QgVb9Bneb2k1TXcLhMmdb/KI0BC9ny/HUj2DwEbEMfdEgVn62t+VPNxEFpi1fIS
# Go0/RMXRAlV6Q2LfcPTw+HycmRNsVg0O8vHpL0PsGyUCwdwnyZD6LWmaqdt7obPx
# 4+RVZR/xQA==
# SIG # End signature block
