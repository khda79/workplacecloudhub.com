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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCuhmeKQcvvE4uB
# FzdmNYKvNEQuvfKwztIx/DGVi37y9KCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDd7ZYrIIvy3oghf+Wp
# oXlpQeAhinpl1IPyHq9+2iCGUjANBgkqhkiG9w0BAQEFAASCAYAisiFgdcOT1ZXu
# FWXs1fHE/NQdiS8evrKsRAoSs3HXIPDCc5TX6dUpB1Vvdake/PJ+PNrskO/NEiGV
# E4GAgGHFuiLTXaAvzYYnRksvd6eWeTdbcRUNYpju815hyw7jrsRgWy+BOS95FA1y
# /WipR6ANpncN6mPZDIO28XXKCYtWHAzk1DwZyVN5wdJRYAWvHhppt8X0ZXxOdvMq
# h4Yxd2cn3zGOW/nEbLNVqYU3gqYr+j4QBXKpOE1igRYOtkXBtdNn/BjlIzPFpsNs
# O14ibKcIWJYzHfkZEiN5t43sa4tu2Xa7ob1UwJHNFtB9rkS4euTju1dy/AyZd7PM
# OxRSb5P5i0IsjgeMXk67EcokoE7KCy0JOVzNKne4qjbObLg4+H3mFTPhEDBlLZ8j
# ixYF4lBbHUs3NcWguOw2pxdLj3Vici4gDa7LbEZTP1aAKuIdPNOiqaDkE+GGmQdk
# oHTe7L9usU9BeAlPQw121IqOqC+5aD/9JLxflH9c3I+0PdVUQ0s=
# SIG # End signature block
