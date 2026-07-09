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


