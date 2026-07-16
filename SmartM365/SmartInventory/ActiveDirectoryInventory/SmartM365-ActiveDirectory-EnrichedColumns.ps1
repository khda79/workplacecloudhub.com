<#
.SYNOPSIS
    Returns calculated AD_Computers_AllDomains.csv column names.
.VERSION
1.2
#>
@(
    'ObjectGuidNormalized'
    'OperatingSystemBuildNumber'
    'IsWindows10LtscByName'
    'IsWindows10LtscByBuild'
    'IsWindows10Ltsc'
    'MeetsWindows11MinimumOsBuild'
    'OrganizationalUnitCode'
    'DomainOrganizationalUnitKey'
    'OperatingSystemDisplayName'
    'IsInDisabledObjectsOU'
    'EntityLabel'
    'Manufacturer'
    'NeedsWindows11Upgrade'
    'OperatingSystemShortName'
    'Subnet'
    'EntitySiteType'
    'Windows11UpgradeEligibility'
    'IsInProductionOU'
    'OrganizationalUnitLeafName'
    'ExistsInIntune'
    'M365LicenseForAdUser'
    'IntunePrimaryUserOrganizationalUnitPath'
    'IntunePrimaryUserOrganizationalUnitName'
    'IntuneDiskFreeGB'
    'IntuneDiskTotalGB'
    'IntuneEnrollmentDateTime'
    'HasInsufficientStorageForWindows11Upgrade'
    'ResolvedPrimarySmtpAddress'
    'IntunePrimaryUserPrincipalName'
    'DiskFreePercent'
    'MigrationIssueScore'
    'IsWindows11'
    'IntuneOperatingSystemVersion'
    'IntuneLastSyncDateTime'
    'IntuneWindows11EligibilityCode'
    'IntuneWindows11EligibilityLabel'
    'IntuneDeviceModel'
    'IntuneOwnership'
    'IntuneComplianceState'
    'LastRebootDateTime'
    'IsInTargetHeadquartersOU'
    'PhysicalMemoryGB'
    'IsRegisteredInIntune'
    'IsOutsideTargetHeadquartersOU'
    'IsWindows11AndManagedByIntune'
    'EntraRegistrationDateTime'
    'EntraApproximateLastSignInDateTime'
    'EntraDaysSinceLastSignIn'
    'EntraHardwareIdDeviceCount'
    'HasDifferentDeviceId'
    'EntraCorrelationStatus'
    'EntraNameMatchCount'
    'EntraNameMatchesByObjectGuidList'
    'EntraMatchCountByObjectGuid'
    'ExistsInEntraByObjectGuid'
    'IsEntraRegistrationPending'
    'DeviceMigrationIssueSeverity'
    'M365LicenseForIntunePrimaryUser'
    'HasMailboxForIntunePrimaryUser'
    'IntunePrimaryUserMailboxSizeGB'
    'IntuneEnrollmentDate'
    'EntityIntegrationType'
    'Windows11MigrationPhase'
    'IntuneLastLoggedOnUser'
    'WindowsUpdateWindows1124H2AggregateState'
    'IntuneDeviceId'
    'WindowsUpdateWindows1125H2AggregateState'
    'WindowsUpdateAutopatchFeatureUpdateAnchorAggregateState'
    'WindowsUpdateWindows1124H2CurrentDeviceUpdateStatus'
    'WindowsUpdateWindows1125H2CurrentDeviceUpdateStatus'
    'WindowsUpdateAutopatchFeatureUpdateAnchorCurrentDeviceUpdateStatus'
    'WindowsUpdateWindows1125H2LatestAlertMessage'
    'WindowsUpdateAutopatchFeatureUpdateAnchorLatestAlertMessage'
    'WindowsUpdateWindows1124H2LatestAlertMessage'
    'EntraDeviceIdNormalized'
    'EntraDeviceIdMatchCount'
    'Windows11MigrationPhaseReason'
    'EntraLastSignInStatus'
    'IsEntraLastSignInOlderThan90Days'
    'IsEntraLastSignInWithin90Days'
    'ResolvedLastSignInDateTime'
    'IsActiveInLast90Days'
    'Windows11UpgradeEligibilitySource'
    'EntraObjectId'
    'IsWindows1125H2'
    'IsWindows1124H2'
    'IsActiveInLast30Days'
    'IsActiveInLast45Days'
    'BiosModel'
    'BiosVersion'
    'BiosKey'
    'SecureBootStatus'
    'AdLastLogonCategory'
    'EntraLastSignInCategory'
    'LastRebootCategory'
    'LastActivityDateTime'
    'LastActivityCategory'
    'EntraLastSignIn30DayCategory'
    'LastReboot30DayCategory'
    'LastActivity30DayCategory'
    'AdLastLogon30DayCategory'
    'WindowsUpdateAutopatchFeatureUpdateAnchorBlockingReason'
    'WindowsUpdateWindows1124H2BlockingReason'
    'WindowsUpdateWindows1125H2BlockingReason'
    'WindowsUpdateWindows1124H2RiskBucket'
    'WindowsUpdateWindows1125H2RiskBucket'
    'WindowsUpdateAutopatchFeatureUpdateAnchorRiskBucket'
    'Windows11MigrationAction'
    'IsVirtualMachine'
    'MatchedConfiguredGroups'
    'IsMemberOfConfiguredGroup01'
    'IsMemberOfConfiguredGroup02'
    'IsMemberOfConfiguredGroup03'
    'IsMemberOfConfiguredGroup04'
    'IsMemberOfConfiguredGroup05'
    'IsMemberOfConfiguredGroup06'
    'IsMemberOfConfiguredGroup07'
    'IsMemberOfConfiguredGroup08'
    'IsMemberOfConfiguredGroup09'
    'IsMemberOfConfiguredGroup10'
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
