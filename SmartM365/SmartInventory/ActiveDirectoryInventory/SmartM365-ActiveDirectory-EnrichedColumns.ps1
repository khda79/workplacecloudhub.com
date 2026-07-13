<#
.SYNOPSIS
    Returns AD_Computers_AllDomains_Enriched.csv column names required by SmartWorkplace Power BI.
#>
@(
    'ObjectGUID_Norm'
    'BuildNumber'
    'IsWindows10LTSC_OS'
    'IsWindows10LTSC_BuildFallback'
    'IsWindows10LTSC'
    'OSMinToUpdateW11'
    'OSMinToUpdateW11_Num'
    'OrganizationalUnit'
    'LastLogonDateConverted'
    'DomainAndOrganizationalUnitFromADManage'
    'OperatingSystem_Major_Current'
    'AccountToDeleteFromAD'
    'FriendlyOSVersionName'
    'IsInDisabledObjectsOU'
    'LabelFromEntity'
    'Manufacturer'
    'NeedToBeUpgrade'
    'ShortOSName'
    'Subnet'
    'TypeEtablissement'
    'W11Eligibilty'
    'WIN11_GLOBALRESULT'
    'Win11_Manufacturer'
    'Is_In_OrganizationOU'
    'LastOUName'
    'ExistsInIntune'
    'LicenseFromM365_Users_From_AD'
    'OU_Path_Users_From_AD'
    'LastOUName_Users_From_AD'
    'DiskFreeStorage_Go_From_M365'
    'DiskTotalStorage_Go_From_M365'
    'Enrollment_date_From_M365'
    'IsFreeStorageNotEnoughForWin11Update'
    'Last_Logged_UserDomain'
    'PrimarySMTPaddress_From_AD_Or_M365'
    'Primary_user_email_address_From_M365'
    'DiskTotalFreeSpace'
    'DiskTotalSize'
    'DiskFreePercent'
    'Mig_Potential_Issues_Score'
    'IsWindows11'
    'IsWindows11-Bool'
    'PrimarySMTPaddress_From_S1_Or_M365'
    'OS_version_From_M365'
    'LastcheckIn_date_From_M365'
    'UpgradeEligibility_From_M365'
    'UpgradeEligibility_Label_From_M365'
    'Model_From_M365'
    'Ownership_From_M365'
    'Compliance_From_M365'
    'Memory'
    'Last Reboot Date'
    'Memory_GB_Number'
    'IsTargetOU_HQ'
    'PhysicalMemoryGB_From_M365'
    'Intune_registered_From_M365'
    'IsNotTargetOU_HQ'
    'IsWindows11AndIntune'
    'PrimarySMTPaddressUser'
    'Mig-FlagMatchExistsW1124H2'
    'Mig-FlagMatchExistsW1124H2-Bool'
    'AzureEntra_RegistrationDateTime'
    'AzureEntra_ApproximateLastSignInDateTime'
    'AzureEntra_DaysSinceLastSignIn'
    'AzureEntra_HardwareId_DeviceCount'
    'HasDifferentDeviceId'
    'AzureEntra_CorrelationStatus_Robust'
    'EntraNameMatchesCountByName'
    'EntraNameMatchesByGUIDList'
    'EntraMatchesCountByGUID'
    'EntraExistsByGUID'
    'EntraRegisteredPending'
    'Mig-CompletedW1124H2Devices%'
    'OrganizationalUnit_User_From_AD'
    'Mig_Potential_Issues_Devices_Level2'
    'LicenseFromM365_Users_From_LastUserS1'
    'Primary_username_address_From_M365'
    'LicenseFromM365_Users_From_PrimaryIntuneUser'
    'IsMailBox_From_Last_Logged_UserDomain'
    'IsMailBox_From_PrimaryIntuneUser'
    'IsMailBoxSize_From_Last_Logged_UserDomain'
    'IsMailBoxSize_From_PrimaryIntuneUser'
    'OrganizationalUnit_LastUserS1_From_AD'
    'OrganizationalUnit_PrimaryUser_From_AD'
    'DistinguishedName_Last_Logged_UserDomain'
    'Model'
    'Enrollment_date_only_From_M365'
    'Mig-MigrationPlanW11-StartDate'
    'TypeEntity'
    'Mig-MigrationPlanW11-Phase'
    'Last_Logged_User'
    'Last_Logged_User_FromM365'
    'Last_Logged_User_FromSentinel'
    'WIN11_GLOBALRESULT_OLD'
    'WU_AggregateState_loc_Policy_Windows_11_24H2'
    'DeviceID_From_M365'
    'WU_AggregateState_loc_Policy_Windows_11_25H2'
    'WU_AggregateState_loc_Policy_Autopatch_FeatureUpdate_Anchor'
    'WU_CurrentDeviceUpdateStatus_loc_Policy_Windows_11_24H2'
    'WU_CurrentDeviceUpdateStatus_loc_Policy_Windows_11_25H2'
    'WU_CurrentDeviceUpdateStatus_loc_Policy_Autopatch_FeatureUpdate_Anchor'
    'WU_LatestAlertMessage_loc_Policy_Windows_11_25H2'
    'WU_LatestAlertMessage_loc_Policy_Autopatch_FeatureUpdate_Anchor'
    'WU_LatestAlertMessage_loc_Policy_Windows_11_24H2'
    'OperatingSystem_20251217_111133'
    'OperatingSystem_Changed'
    'OperatingSystem_Changed_Text'
    'OperatingSystem_Major_20251217'
    'OperatingSystem_10to11_Changed_Text'
    'OperatingSystem_10to11_Changed'
    'OperatingSystem_20250918_170952'
    'OperatingSystem_Snapshot_Priority'
    'OperatingSystem_Major_Snapshot_Priority'
    'OperatingSystem_SnapshotUsed'
    'entraDeviceId_norm'
    'entraDeviceId_match_count'
    'Last_Logged_User_Source'
    'Mig-MigrationPlanW11-Phase-Reason'
    'AzureEntra_LastSignIn_Status'
    'AzureEntra_LastSignIn_OlderThan3M'
    'AzureEntra_LastSignIn_Recent_3M'
    'LastSignIn_Merged'
    'LastSignIn_Merged_Active_Last90Days'
    'FriendlyOSVersionName_Snapshot'
    'operatingSystemVersion_20250918_170952'
    'operatingSystemVersion_20251217_111133'
    'OperatingSystemVersion_Snapshot_Priority'
    'WIN11_GLOBALRESULT_INTUNE_ONLY'
    'WIN11_GLOBALRESULT_M365_ONLY'
    'WIN11_GLOBALRESULT_SOURCE'
    'AzureEntra_ObjectId'
    'IsWindows1125H2-Bool'
    'IsWindows1124H2-Bool'
    'LastSignIn_Merged_Active_Last30Days'
    'LastSignIn_Merged_Active_Last45Days'
    'MailboxSize_From_Last_Logged_UserDomain'
    'BIOS_Model'
    'BIOS_Version'
    'BIOS_Key'
    'SecureBootStatus'
    'Enabled_From_Last_Logged_UserDomain'
    'LastLogonDate_Category'
    'AzureEntraLastSignInDate_Category'
    'LastRebootDate_Category'
    'Last Active Date'
    'LastActiveDate_Category'
    'LastRebootDate_Category2'
    'LastActiveDate_Category2'
    'AzureEntraLastSignInDate_Category2'
    'AzureEntraLastSignInDate_Category30Days'
    'LastRebootDate_Category30Days'
    'LastActiveDate_Category30Days'
    'LastLogonDate_Category30Days'
    'WU_BlockingReason_loc_Policy_Autopatch_FeatureUpdate_Anchor'
    'WU_BlockingReason_loc_Policy_Windows_11_24H2'
    'WU_BlockingReason_loc_Policy_Windows_11_25H2'
    'WU_RiskBucket_loc_Policy_Windows_11_24H2'
    'WU_RiskBucket_loc_Policy_Windows_11_25H2'
    'WU_RiskBucket_loc_Policy_Autopatch_FeatureUpdate_Anchor'
    'Mig-MigrationPlanW11-Action'
    'IsVirtualMachine'
    'ModelReleaseDate'
)


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCiUXw8Koc1nH09
# ofWiw4oH9ugiZzZ2Fixj17aMUyNgiaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIDQ6sWNc08m4hZAhDfriJDMjdDu7fmBCXYcgbNwAaQMsMA0GCSqG
# SIb3DQEBAQUABIIBgG47oZjak52jZLaYcw1+padqNm64Trolu5hW1wtWM+j5JVSU
# SDwiBttbgfeqcUUbfQv/eeb9AgI73fFjiqodbCYTrn1WRV4iK30xu7tjNQGJfBJN
# JzHOAN6IolYeOoR9hl6YuIrc36HCI+xxoWYNToFu08WAPrulAUnihW6bfGFoBUBt
# wRhuLp6OM+De3nAycgecpv21wJEv/nin9h7Ooubcfq76k6mEc8fEl+/g2wkwL6X/
# gblgGI2wdANb/8Hu2UeZRq6nX8GwmxJZGNApjnH3u+iEKBQybQlJMfxiN64chaDc
# 5gIBsDZmvOJsonCZN1hr8onCi6bixC2q8NOgXWDARv12XFdpiJStD0A2cdQrbuvb
# +YuDVupgCfhiov2nUIhF3poEx550rRrPRiBcVMFxxUDN3o9ewnMDekZj2Hxi54Xe
# uf82ORw5C25BBVbvP41WAhrfa+PSlJP3CcLB8+trwAIAbhVRBV5nSv2GiYfz6x3O
# i38qbj9aYP81yNeERqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjFaMC8GCSqGSIb3DQEJBDEiBCA32Tg2u5LMzA77G+StYYLD6ynx6mrBmXafMjW9
# q6phiTANBgkqhkiG9w0BAQEFAASCAgDPTuvHX4DVBJQR5vqI333YM7WcqjCkzVdx
# GYat0R9JgmsOOxFOBe6gc8oPbKFyusrhVRsxrd2Cla/sWWYu4uueYaV0T2h007Wo
# +tp6EIOqEAStUrk1rWfLkUrAq3M/L05ts3EIrnr7BjjQZnlRp3z4w2zyVQ374jQF
# 069GsAqiCjCFnNAP2xJGotGKEvuPm58XzF2pF50SLBM9hK7pKY0p4Q3y7lENJXhy
# bh7N4XWN90gTxefY6S5iv/Lz2CpqhD+E6M9oL/z5cyJh43RO4DxcvR2FZZjA3BRP
# pR4cDQoWwIrRWOqS6dVnTs7CuwvQTzowIbSKxVQ0Ur5MAER8GodJDyov+yj/5PT+
# Vm2NEoTsfngQtw+sv7QROBOYqCr4ggOVk1iGUrlx1U2B84ngWbxhzn9ylZVG7BQj
# ZDsexVbNh7TI6aUuRjD9GBlQd57P+WanipPYSnIWlDhVq9U38UDNQTROL/vua2hi
# 82zvWqgk0l/0QtUJjel0E4vThwDG+8HUztL/3VXzAXdwaLZ76eu9RSs/ejNwG1bB
# TFOVFelq5/5XlncejU71XOVK3zjk8GZgvL33knGnUiNk4DQxvFWW7VkesODRmpvE
# 8IRQVlKd+theZE87gyJIO60Iu253VL8sTOUUvGUJdURS8c+Mli4jfqXh5P8IE563
# wVXul4ZB3g==
# SIG # End signature block
