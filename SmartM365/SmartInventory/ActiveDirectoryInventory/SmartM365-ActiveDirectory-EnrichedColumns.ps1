<#
.SYNOPSIS
    Returns AD_Computers_AllDomains_Enriched.csv column names required by SmartWorkplace Power BI.
#>
@(
    'ObjectGUID_Norm'
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA68RF4kZUc2NXK
# t+t1b8dK2ewSfOhtPFn3JVog1w9bbKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAx6egB6SnvU/MOaS3eKQe8DdEblY7AD7zlL+p4rwseSDANBgkqhkiG9w0B
# AQEFAASCAYAxy1hug5NKbp3IWwDhpV1EWUyO+Whj6lg9w/BzsSIyQ2dgtz+w8sqd
# URE9oJs0AssD8bIs/MhEsEI0kVApM4AA7K1wFZ069R0djKyq3biZsWNiCi7vfw6n
# wZeBGNtgq8dhCP1XoDuJd4YmT07WX4u5zY3YfFlScnAcPxZPtq1laE9QZGYbvHLS
# Z3XWetpTykzTVqVHVrzRiTUy7oAUPwqTcneY3JfR68tAhdv123n2yqciX0v4LhcA
# DyByqWdsSdw0Upki2DaDsWvxyVQQq3fW1mFEsqo/aJtns8E4ctZRfZXTgWa7I2kI
# +wkPGz9uoTG/Zg1knTURYmuHZxPVFIyuV4Chx4AqySd0RnlzonBdMYLS/G9iv1xu
# HTwN5O3pcNWm07+oP94MxtypcYc82godHsTSCHYnyOaQ5ulrr8UcTABattNJBivZ
# j45+AHjhQ12ByQRgvnHjvjJmJSftMAjoEGj/q2KJPkjo3OAsumKSb9xpLd5+sxEj
# uzF81XM7Rtc=
# SIG # End signature block
