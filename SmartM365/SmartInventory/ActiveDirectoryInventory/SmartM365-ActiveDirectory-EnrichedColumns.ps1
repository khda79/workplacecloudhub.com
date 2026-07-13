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
    'DomainAndOrganizationalUnit'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCB1Pp6wuQh710I
# NZHNsaoWZ2/Jb73LT2AzywubD4gNVaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBsyIIhqgbd0hKUZNrTXyrA
# kRxW3b7C2hIqCzzQBt8azTANBgkqhkiG9w0BAQEFAASCAYA0a/2xDDmq7gMswbCB
# VGm0PABA30vOgW5RRJy+u4N1awicJQJXFOi/ZVtO1Qk96udsN4P6thfPVi9kuih8
# AAuqBWG2aNrtFNX8rxkGjurZiCXkizfgW3f3Sl/vwZiBaL82b/g1ynVjnisfbAjV
# 13JXX6LNF22f/Bv+5wgSCo0vrC3oPq6TKSignn4vHdZIrtOIarMhrYeXTNZN4+HC
# bpUo/Y9hbuHPe8bGR3oOupO33/dq3Cy8jIsrHap34uulOjrcezFb1nNdI7F4kzfB
# 45tBvQxJDXQFFVK2AIwM/e6rrGnGU6B87xPGVk6+mIs9QaXyKlBEw3rOtB06Ahqb
# +xoJlFwvk3jFHVoBqL7upS4HFP6oY55wqj1TIfDfH2eB+wMP0H2oLnwmMCdutgFW
# Q20sbLN1UjvAgPLQ0u+96q8X6wtn4C5HapdObxrH4VgOO8Xu9CI2bZ4fN70neYsr
# 0Etf7WdmP226m35/u5yRFwh/0QCNV+6JD+AnR+8qM+9aVhg=
# SIG # End signature block
