<#
.SYNOPSIS
Initializes and builds the SmartWorkplaceCMDB normalized output model.

.DESCRIPTION
This first build scaffold creates the tenant output folders, initializes empty CMDB and Power BI-ready tables, and writes a build manifest. Source collectors will populate the raw and normalized tables in later phases.

.VERSION
0.1.2
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'Default',
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [switch]$ValidateOnly
)

$ScriptVersion = '0.1.2'
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
$modulePath = Join-Path -Path $projectRoot -ChildPath 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'

Import-Module $modulePath -Force

$paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant $Tenant -DataRootPath $DataRootPath -DataAllRootPath $DataAllRootPath -LatestOutputRootPath $LatestOutputRootPath -LogRootPath $LogRootPath

if ($ValidateOnly) {
    [pscustomobject]@{
        Status               = 'Valid'
        ScriptVersion        = $ScriptVersion
        TenantKey            = $paths.TenantKey
        ProjectRootPath      = $paths.ProjectRootPath
        LatestOutputRootPath = $paths.LatestOutputRootPath
        PowerBILatestPath    = $paths.PowerBILatestPath
    } | Format-List
    return
}

Initialize-SmartWorkplaceCMDBTenantFolder -Paths $paths

$buildTimestamp = (Get-Date).ToString('o')

$tableDefinitions = @(
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Users.csv'; Columns = @('TenantKey','CmdbUserId','SourceSystem','SourceUserId','UserPrincipalName','DisplayName','AccountEnabled','UserType','Department','JobTitle','ManagerUserId','CreatedDateTime','LastSignInDateTime','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Devices.csv'; Columns = @('TenantKey','CmdbDeviceId','SourceSystem','SourceDeviceId','DeviceName','OperatingSystem','OperatingSystemVersion','Ownership','ComplianceState','ManagementState','PrimaryUserId','LastSyncDateTime','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Groups.csv'; Columns = @('TenantKey','CmdbGroupId','SourceSystem','SourceGroupId','DisplayName','MailEnabled','SecurityEnabled','GroupTypes','MemberCount','OwnerCount','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Licenses.csv'; Columns = @('TenantKey','CmdbLicenseId','SourceSystem','SkuId','SkuPartNumber','ConsumedUnits','EnabledUnits','SuspendedUnits','WarningUnits','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Mailboxes.csv'; Columns = @('TenantKey','CmdbMailboxId','SourceSystem','ExternalDirectoryObjectId','UserPrincipalName','DisplayName','RecipientTypeDetails','PrimarySmtpAddress','MailboxPlan','ArchiveStatus','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_UserDeviceRelationships.csv'; Columns = @('TenantKey','CmdbRelationshipId','CmdbUserId','CmdbDeviceId','RelationshipType','SourceSystem','ConfidenceScore','Evidence','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Relationships.csv'; Columns = @('TenantKey','CmdbRelationshipId','FromEntityType','FromEntityId','ToEntityType','ToEntityId','RelationshipType','SourceSystem','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_DataQuality.csv'; Columns = @('TenantKey','FindingId','Severity','EntityType','EntityId','FindingType','Description','SourceSystem','DetectedDateTime','RecommendedAction') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimTenant.csv'; Columns = @('TenantKey','TenantDisplayName','Environment','LastRefreshDateTime') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimUser.csv'; Columns = @('TenantKey','TenantUserKey','CmdbUserId','UserPrincipalName','DisplayName','AccountEnabled','UserType','Department','JobTitle','ConfidenceScore') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimDevice.csv'; Columns = @('TenantKey','TenantDeviceKey','CmdbDeviceId','DeviceName','OperatingSystem','OperatingSystemVersion','Ownership','ComplianceState','ManagementState','ConfidenceScore') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimGroup.csv'; Columns = @('TenantKey','TenantGroupKey','CmdbGroupId','DisplayName','MailEnabled','SecurityEnabled','GroupTypes') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimLicenseSku.csv'; Columns = @('TenantKey','TenantSkuKey','SkuId','SkuPartNumber','ConsumedUnits','EnabledUnits') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimDate.csv'; Columns = @('Date','Year','Quarter','Month','MonthName','Day') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactUserLicense.csv'; Columns = @('TenantKey','TenantUserKey','TenantSkuKey','CmdbUserId','SkuId','AssignmentState','AssignedDateTime','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactDeviceCompliance.csv'; Columns = @('TenantKey','TenantDeviceKey','CmdbDeviceId','ComplianceState','LastSyncDateTime','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactUserDeviceRelationship.csv'; Columns = @('TenantKey','TenantRelationshipKey','TenantUserKey','TenantDeviceKey','CmdbRelationshipId','CmdbUserId','CmdbDeviceId','RelationshipType','ConfidenceScore','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactMailbox.csv'; Columns = @('TenantKey','TenantMailboxKey','TenantUserKey','CmdbMailboxId','CmdbUserId','RecipientTypeDetails','ArchiveStatus','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactDataQuality.csv'; Columns = @('TenantKey','TenantFindingKey','FindingId','Severity','EntityType','FindingType','DetectedDateTime','SourceSystem') }
)

foreach ($table in $tableDefinitions) {
    Export-SmartWorkplaceCMDBCsv -InputObject @() -Path (Join-Path -Path $table.Folder -ChildPath $table.Name) -Columns $table.Columns
}

$manifest = [pscustomobject]@{
    TenantKey            = $paths.TenantKey
    ScriptVersion        = $ScriptVersion
    BuildDateTime        = $buildTimestamp
    CmdbTableCount       = ($tableDefinitions | Where-Object { $_.Folder -eq $paths.CmdbLatestPath }).Count
    PowerBITableCount    = ($tableDefinitions | Where-Object { $_.Folder -eq $paths.PowerBILatestPath }).Count
    LatestOutputRootPath = $paths.LatestOutputRootPath
}

Export-SmartWorkplaceCMDBCsv -InputObject @($manifest) -Path (Join-Path -Path $paths.CmdbLatestPath -ChildPath 'CMDB_BuildManifest.csv')

Write-Information ("SmartWorkplaceCMDB build scaffold completed for tenant '{0}'." -f $paths.TenantKey) -InformationAction Continue
Write-Information ("Latest output: {0}" -f $paths.LatestOutputRootPath) -InformationAction Continue



# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBA4D8USKqGEf8i
# hcCGVH5yz/M7ubLAIQ7LNk9Ln54HhKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJKURTOVbniITC5Aqd5UHazzCzDQCTSuDdAYPsdJRUVGMA0GCSqG
# SIb3DQEBAQUABIIBgJQBCB1pA1PQLR5u7T3abhuLgSzqCUkBTXO9v70r8iNfD4GT
# ZL93mDOIXzXzrsdT89alU6p73JBsCOw2qwat/ae7aF5PXaGT0zDp2IlMIMufxIo7
# /yV4KyamVjogr5cu/a4JfreF5UZV5CM/tpty3r2JlxytGKIFlbmmdGkscyT4VqeO
# wGXC8XH5sl7y2NP24ZAXpOSzsa/4V9JHlZxmhR5IAvHEDuv9Nb/CEWSxXjzar5mO
# HkBf6ExM4hFXh3uybhcH2JV8QARImdD1WzAyz9wBGtqR3oz12pcgdzVVt7oOYIM2
# NzLMCeax7C5mfeENjvwj2aYyQIecgX7lzDfziCotfg5mWI9QeVVmyBAhW7Eckpuq
# Uzb28sIIbiStQIqfKSt/4YA3sTVo4qfbNPBknmUxEl6nYiHJfJZk3CBBjHlOojO0
# m6kNRqlgf60IW6E8a8ZGWnFPOCUBhzuTLtm0DW43bUp/84BXIh6P+yfLuwQ6pyae
# Zzhb1kdVHuDJVtY0kKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMxMDUy
# MjVaMC8GCSqGSIb3DQEJBDEiBCAmy4+NP9TBk13q5cdAgpfEI7IPC3lcX2BLd089
# +SAR+jANBgkqhkiG9w0BAQEFAASCAgAbt69eXdhvqxpC44YZXBU5KmlwUv9CZpXM
# wdCsN61J8DkzXGWI+GWdIi1TF3tbPjjaVBHrLzcv72ce6K6T2Ucl2WKWqxkqUmCv
# uchQzG7Q4uXjVFk/Hip6wXtaTMpJWWLO9amtgzvJwMflMzMe6gn0qUGnQfALo0Uf
# XdR3tffIuYp8TQ+S8DV3I+HmFBxYp497vJ7d9mlLv/hK/lPmZ6Kw2argL6k38UVm
# 99fY5APo9YlWNNQV0+FPPmVrq87v8UEjvMco3rVCwp5GecNfEOEIQA1hMUFHxEa7
# 09ozcvVHjjOKzm26aJFN2o6aysJk3P4NBsmpNaWZGX5DNDwGpYv7OIeVUToSOWjz
# Wq2N4MjW1J3LZfoYqX1BGnz4yon22e8yS0+TpBLdFjHTFIJIv4UsVqQMaKVlFafK
# +Li11MS3DWSec+qLn0mWBcsVb41kxDEYwW2bLHRGTP5f78rFLPXAB2+nscAGcyTi
# S4Tswmt3HAC4mYa4aM2klqBsBjDcEennZi+GwNtFMUsB1h2ioNc7ocKGGBIeVprO
# ByCpDKyC7FYNU9y0dQxOdrCJu9gwkBw2Xk3jGRs4XYsrd7yvCtmkpIdqxqQQrEQ0
# WHNaOe57opIQzy/PLJiQI5RGrkOpGDYqaQ+j4pE7lngxg69HtQlu9pRmTevpI8IU
# EZQTDTEPew==
# SIG # End signature block
