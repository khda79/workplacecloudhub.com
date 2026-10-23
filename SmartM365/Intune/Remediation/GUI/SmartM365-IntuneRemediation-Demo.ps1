#Requires -Version 7.0
<#
.SYNOPSIS
Fictional local/cloud packages, editors, execution results and safe lifecycle actions.
.VERSION
1.1.0
#>
function script:Initialize-SmartM365DemoProfile {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $names=@('DeliveryOptimization ContentEngine Health','Disk Space Cleanup','IntuneManagementExtension Health','MDM Enrollment Repair','BitLocker OSDrive Protection','Upgrade Diagnostics','Upgrade Staging Health','Recovery Check','Time Zone','WinHTTP Proxy Configuration','Repair Windows Update and DISM','Force Allow OS Upgrade','GroupPolicy Stale','WindowsUpdateLog Errors','WUfB Identity Drift','UWP Store AppRepository WU Health','Windows Update Reset','Windows Update Service and Scan Health','Windows Update Policy Blockers','Windows Update Cache Health','Reset Windows Hello','W11 Upgrade Compatibility Indicators (detection only)','Windows LAPS Health','Windows Time Sync','Defender Signature Freshness','Network Certificate Expiry','BITS Stale Jobs','Print Spooler Health','Pending Reboot Age','Pending Reboot Age - Device Reboot Manager')
    $rows=@(for($i=0;$i -lt $names.Count;$i++){
        [pscustomobject]@{Id=('00000000-0000-4000-8000-'+($i+1).ToString('000000000000'));DisplayName=$names[$i];Folder=('Demo/Packages/'+($names[$i] -replace ' ','-'));IntuneStatus=$(if($i%3 -eq 0){'Update available (demo)'}else{'In sync (demo)'});DetectionOnly=($names[$i] -like '*(detection only)*');PSScriptAnalyzerStatus='Passed (sample)';PortalStatus='Active (demo)';NoIssueDetectedDeviceCount=(30+$i);IssueDetectedDeviceCount=5;IssueRemediatedDeviceCount=4;IssueReoccurredDeviceCount=1;IssueRemediatedCumulativeDeviceCount=12}
    })
    $s.Data.Packages=$rows;$s.Data.CloudPackages=@($rows|ForEach-Object {$_.PSObject.Copy()})
    Set-SmartM365DemoGrid LocalGrid $rows
    Set-SmartM365DemoGrid CloudActiveGrid $s.Data.CloudPackages
    $archive=$rows[0].PSObject.Copy();$archive.DisplayName='Contoso retired package';$archive.PortalStatus='Archived (demo)'
    Set-SmartM365DemoGrid CloudOtherGrid @($archive)
    $s.Data.Detection="# DEMO - Synthetic detection script; displayed and exported as text only.`r`nWrite-Output 'Contoso demonstration: no issue detected.'`r`nexit 0"
    $s.Data.Remediation="# DEMO - Synthetic remediation script; displayed and exported as text only.`r`nWrite-Output 'Contoso demonstration: remediation simulated.'`r`nexit 0"
    foreach($name in @('LocalDetectionEditor','CloudDetectionEditor')){$c[$name].Text=$s.Data.Detection}
    foreach($name in @('LocalRemediationEditor','CloudRemediationEditor')){$c[$name].Text=$s.Data.Remediation}
    $c.PackageDetails.Text="DEMO - Local package`r`nAssignments: Contoso Pilot Devices (fictional)`r`nSchedule: Daily 09:00`r`nAll scripts and saves remain in the demo session."
    $c.CloudDetails.Text="DEMO - Intune cloud package`r`nAuthentication: not used`r`nAssignments: Contoso Pilot Devices`r`nLast execution: 2026-10-23 09:30 (synthetic)"
    $c.LogTextBox.Text='[2026-10-23 09:30:00] DEMO - Local and cloud package fixtures loaded; no package files or Graph data were read.'
    $c.BusyOverlay.Visibility='Collapsed'
}
function script:Connect-SmartM365DemoProfileEvents {
    $s=$script:SmartM365Demo;$c=$s.Controls
    $c.LocalGrid.Add_SelectionChanged({$s=$script:SmartM365Demo;$row=$s.Controls.LocalGrid.SelectedItem;if($row){$s.Controls.PackageDetails.Text='DEMO - '+$row.DisplayName+"`r`nFolder: "+$row.Folder+"`r`nIntune: "+$row.IntuneStatus}})
    foreach($name in @('CloudActiveGrid','CloudOtherGrid')){$c[$name].Add_SelectionChanged({param($grid)$row=$grid.SelectedItem;if($row){$script:SmartM365Demo.Controls.CloudDetails.Text='DEMO - '+$row.DisplayName+"`r`nAssignments: Contoso Pilot Devices`r`nNo Graph request was made."}})}
}
function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s=$script:SmartM365Demo;$c=$s.Controls
    switch -Regex($Name){
        '^ConnectButton$' {Write-SmartM365DemoActivity 'Graph connection simulated. No token or sign-in was requested.';return $true}
        '^Analyze' {foreach($row in $s.Data.Packages){$row.PSScriptAnalyzerStatus='Passed (sample)'};$c.LocalGrid.Items.Refresh();Write-SmartM365DemoActivity 'Sample script analysis completed; no module was installed or script executed.';return $true}
        '^Save' {
            foreach($field in @('LocalDetectionEditor','LocalRemediationEditor','CloudDetectionEditor','CloudRemediationEditor')){
                $path=Get-SmartM365DemoOutputPath ($field+'-DEMO.ps1')
                [IO.File]::WriteAllText($path,$c[$field].Text,[Text.UTF8Encoding]::new($false))
            }
            Write-SmartM365DemoActivity 'Editor contents saved to the separate demo session only.';return $true
        }
        '^Publish' {
            $local=$c.LocalGrid.SelectedItem
            if($local){$copy=$local.PSObject.Copy();$copy.DisplayName+=' - published demo';$s.Data.CloudPackages+=@($copy);Set-SmartM365DemoGrid CloudActiveGrid $s.Data.CloudPackages}
            Write-SmartM365DemoActivity 'Package publication simulated in memory. Intune was not contacted.';return $true
        }
        '^DeleteCloudButton$' {$row=$c.CloudActiveGrid.SelectedItem;if($row){$s.Data.CloudPackages=@($s.Data.CloudPackages|Where-Object Id -NE $row.Id);Set-SmartM365DemoGrid CloudActiveGrid $s.Data.CloudPackages};Write-SmartM365DemoActivity 'Selected fictional cloud package deleted in memory.';return $true}
        '^ResetHistoryButton$' {$row=$c.CloudActiveGrid.SelectedItem;if($row){$row.NoIssueDetectedDeviceCount=0;$row.IssueDetectedDeviceCount=0;$row.IssueRemediatedDeviceCount=0;$row.IssueReoccurredDeviceCount=0;$row.IssueRemediatedCumulativeDeviceCount=0;$c.CloudActiveGrid.Items.Refresh()};Write-SmartM365DemoActivity 'Execution history reset in demo memory only.';return $true}
        '^CompareLocalCloudButton$' {[void](Show-SmartM365DemoDetail -Title 'Local / cloud comparison' -Text "DEMO - Sample scripts match.`r`nDetection: identical`r`nRemediation: identical`r`nAssignments: Contoso Pilot Devices (fictional)`r`nNo local package or cloud object was read or modified.");return $true}
        '^ExportReportButton$' {[void](Export-SmartM365DemoProfileOutputs);Write-SmartM365DemoActivity 'Fictional execution CSV exported.';return $true}
    }
    return $false
}
function script:Export-SmartM365DemoProfileOutputs {
    $path=Get-SmartM365DemoOutputPath 'IntuneRemediation-Execution-DEMO.csv'
    @(1..12|ForEach-Object {[pscustomobject]@{Device=('CTSO-LAP-{0:0000}' -f $_);User=('user{0:00}@contoso.example' -f $_);Detection='Issue detected (sample)';Remediation='Remediated (sample)';LastRun='2026-10-23 09:30';Mode='DEMO'}})|Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8
    $script:SmartM365Demo.Reports.Execution=$path
    return @($path)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC3Sw/9BPdQBDmR
# 0AWJE33VOqPcFd5GrukmLx11XRAJBqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB+f0QM3tSHLNm4j91kMvj03+dSAnNeXu72JDX96cBD5MA0GCSqG
# SIb3DQEBAQUABIIBgCB/NzFVjHHNafvk60Tyfaa38ydOBqZ20FAbmvIzLCxcslCW
# dM7SjIlmfeFZzyiK39O/dRckuXsB4RIEuP7T0bo0lp7EABAjkwys8N+ucTIhg/Df
# lZGHW8ibr4072WkNsf2CHkL6WfTKZK40qdVM52YF79Tq8uHYzkJUlBnGN7wUACTA
# YFeUAcQEpVs6x7cKRmNBuyWS7cMbVxgvPd/lAPGdBPNguJLpciIUVDWIIO1+NQIq
# 7THUOZqmdtuR+xsTdNBuYcrUpQw0bU6GYma5qDVq/sozPRzCJZLCZjyOyC1lDUXF
# JccSkx6AovjArkJ9EVKLeTdzdEYRBkFu0pjfaP5hOegM3XqjsntEuLwIWIpxfGZR
# A+wpFlgGMmbii8kmME/+Hkrw67g99Aoof+RnCiLsjAExsLO4I/MeXFLyxP6Hinv+
# CCiax0hpEfForXBZAsInqgNsqUsE5uQPOsI7snmn5wMkVVWIraAtL7gFdtzqYZYC
# 5KbiTuF2IHzQWgcG36GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxOTEy
# MTBaMC8GCSqGSIb3DQEJBDEiBCBKUr3uGkhNH3AhgyyYDTEYKz3tyDIjFGxo8cQ2
# 0K3kGzANBgkqhkiG9w0BAQEFAASCAgBatu5kc0Ti06/sVfpHwSLdyvq+MqGKWWQ9
# AdsjUnwh04L7U6HnRhOzFoZCGRkrNtgkN0nizuo66C9dZymV9Dj9aG1y3JgovglV
# jxQPg3C8QZChIssZNfvKonEjdpy50R834YiSReZlimuI6wRx5l4O+sUW382wNnFo
# IqAwWdAAivV7ahs+K0gydSwHBWxYgty+uKRvxy78v/9umFcgacTYPleb1/MerDs9
# 9dRxHGmO/NSri8SK9q8fZHun65B7B/J61SHgr3Cnpt+2j+wOiZvnGKU66kaxQPaJ
# 5Q0jkCszZ5rEr7orAOxewyz6S2kGXCQZIO/HMyHRkzplG3wSjYFSIQVuArCqRI0I
# qunVKCisSgp3oJPmU4DpI2bKowvU1xWaTDfhkI8rMnkUi5ThdnG9FzyvW13zxZa7
# tUTgaNUQ0sKrnSSu89+5KFdkdO58EqtafQOfP06LXx/qcaR0YIKDUwjghrtMOsWl
# YMjyjwWY3HzdpxFz4MMQo7m/bTv0MiWjH3jr9HFXQJArx4hiMhRtjGRUwSDnEfdm
# xesY2pKxOah23N0c4UIvGpSqbFyQPdTChaWp3r1EwVCgv0lUDNEPKzFfyNXmcXQW
# 9FSsnL0S7aqiC5dTBxhELDTmabKytombJ6MKhQHkD2o56/f4zF6/6S3J7Mz6Q+AC
# qhiuAvgg0A==
# SIG # End signature block
