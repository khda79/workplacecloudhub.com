<#
.SYNOPSIS
Initializes and builds the SmartWorkplaceCMDB normalized output model.

.DESCRIPTION
This first build scaffold creates the tenant output folders, initializes empty CMDB and Power BI-ready tables, and writes a build manifest. Source collectors will populate the raw and normalized tables in later phases.

.VERSION
0.1.1
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

$ScriptVersion = '0.1.1'
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
        Tenant               = $paths.TenantKey
        ProjectRootPath      = $paths.ProjectRootPath
        LatestOutputRootPath = $paths.LatestOutputRootPath
        PowerBILatestPath    = $paths.PowerBILatestPath
    } | Format-List
    return
}

Initialize-SmartWorkplaceCMDBTenantFolder -Paths $paths

$buildTimestamp = (Get-Date).ToString('o')

$tableDefinitions = @(
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Users.csv'; Columns = @('CmdbUserId','SourceSystem','SourceUserId','UserPrincipalName','DisplayName','AccountEnabled','UserType','Department','JobTitle','ManagerUserId','CreatedDateTime','LastSignInDateTime','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Devices.csv'; Columns = @('CmdbDeviceId','SourceSystem','SourceDeviceId','DeviceName','OperatingSystem','OperatingSystemVersion','Ownership','ComplianceState','ManagementState','PrimaryUserId','LastSyncDateTime','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Groups.csv'; Columns = @('CmdbGroupId','SourceSystem','SourceGroupId','DisplayName','MailEnabled','SecurityEnabled','GroupTypes','MemberCount','OwnerCount','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Licenses.csv'; Columns = @('CmdbLicenseId','SourceSystem','SkuId','SkuPartNumber','ConsumedUnits','EnabledUnits','SuspendedUnits','WarningUnits','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Mailboxes.csv'; Columns = @('CmdbMailboxId','SourceSystem','ExternalDirectoryObjectId','UserPrincipalName','DisplayName','RecipientTypeDetails','PrimarySmtpAddress','MailboxPlan','ArchiveStatus','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_UserDeviceRelationships.csv'; Columns = @('CmdbRelationshipId','CmdbUserId','CmdbDeviceId','RelationshipType','SourceSystem','ConfidenceScore','Evidence','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_Relationships.csv'; Columns = @('CmdbRelationshipId','FromEntityType','FromEntityId','ToEntityType','ToEntityId','RelationshipType','SourceSystem','ConfidenceScore','SourceCollectedDateTime') },
    [pscustomobject]@{ Folder = $paths.CmdbLatestPath; Name = 'CMDB_DataQuality.csv'; Columns = @('FindingId','Severity','EntityType','EntityId','FindingType','Description','SourceSystem','DetectedDateTime','RecommendedAction') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimTenant.csv'; Columns = @('TenantKey','TenantDisplayName','Environment','LastRefreshDateTime') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimUser.csv'; Columns = @('CmdbUserId','UserPrincipalName','DisplayName','AccountEnabled','UserType','Department','JobTitle','ConfidenceScore') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimDevice.csv'; Columns = @('CmdbDeviceId','DeviceName','OperatingSystem','OperatingSystemVersion','Ownership','ComplianceState','ManagementState','ConfidenceScore') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimGroup.csv'; Columns = @('CmdbGroupId','DisplayName','MailEnabled','SecurityEnabled','GroupTypes') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimLicenseSku.csv'; Columns = @('SkuId','SkuPartNumber','ConsumedUnits','EnabledUnits') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'DimDate.csv'; Columns = @('Date','Year','Quarter','Month','MonthName','Day') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactUserLicense.csv'; Columns = @('CmdbUserId','SkuId','AssignmentState','AssignedDateTime','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactDeviceCompliance.csv'; Columns = @('CmdbDeviceId','ComplianceState','LastSyncDateTime','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactUserDeviceRelationship.csv'; Columns = @('CmdbRelationshipId','CmdbUserId','CmdbDeviceId','RelationshipType','ConfidenceScore','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactMailbox.csv'; Columns = @('CmdbMailboxId','CmdbUserId','RecipientTypeDetails','ArchiveStatus','SourceSystem') },
    [pscustomobject]@{ Folder = $paths.PowerBILatestPath; Name = 'FactDataQuality.csv'; Columns = @('FindingId','Severity','EntityType','FindingType','DetectedDateTime','SourceSystem') }
)

foreach ($table in $tableDefinitions) {
    Export-SmartWorkplaceCMDBCsv -InputObject @() -Path (Join-Path -Path $table.Folder -ChildPath $table.Name) -Columns $table.Columns
}

$manifest = [pscustomobject]@{
    Tenant               = $paths.TenantKey
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCuhmeKQcvvE4uB
# FzdmNYKvNEQuvfKwztIx/DGVi37y9KCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDd7ZYrIIvy3oghf+WpoXlpQeAhinpl1IPyHq9+2iCGUjANBgkqhkiG9w0B
# AQEFAASCAYBgeCXTcUJHnlo+0jBta6aEYDty/6IKDuyWt16CHUs3ukBeuKOvkaqT
# peCr293V7kx9QGb5mmrPM8uPA+jxrC4ZvQQESwp0+1zLl7/406+53O94zqLNLCp6
# bnqmG0302V3c2FiBTMZ5Fdvo2CrUV0HtaFIwCnQ9IUwdGip3fu58wF1vpYSswX5t
# 8rM2wNwyxVNqkWsnxF1lMxFD4hmz3Gwx4VWIjLn8De6M97pW/Xk4nGVUg2XIdgWI
# 9I0o+TIqCZQuvevHa6mdHBZuEoKMSF5FxEddjS/VrewOgyG2k/b5h6lQYqISPqyL
# F9Xd0ozCgXmfESF/+i2wcA/iQFtuGjYnGe639dfYMXrwX9zcINESFSlk/zDicSv9
# FW6RZN1+LqZZvp8qY6RgqILBezsRTX7z3kc8VIr8/Z7AwAb6EvQE8GQzHITag82Z
# ionw16yKvfopoldl6jITobTPTEHwul2TUa0Ox7q0mWpjKNfcgi7bikOoTMk7aDSf
# CJ2Pra7DgCo=
# SIG # End signature block
