@{
    Name = 'NewMigration'

    Source = @{
        # Supported values: SP2016, SP2019, SPO
        Type = 'SP2019'
        WebApplicationUrl = 'https://source-sharepoint.example.com'
        PermissionRootPath = '/SOURCE'
    }

    Target = @{
        # Supported values: SP2016, SP2019, SPO
        Type = 'SPO'
        SiteUrl = 'https://yourtenant.sharepoint.com/sites/NewMigration'
        TenantAdminUrl = 'https://yourtenant-admin.sharepoint.com'
        PrefixToRemove = '/sites/NewMigration'
        PermissionRootPath = '/sites/NewMigration'
    }

    Comparison = @{
        MaxScanAgeDifferenceHours = 12
        PermissionMaxScanAgeDifferenceHours = 24
        SizeToleranceBytes = 10240
        ModifiedDateToleranceMinutes = 0
        PathMappingsFile = 'migration.mapping.txt'
        SourceModifiedTimeZone = 'Auto'
        TargetModifiedTimeZone = 'Auto'
        ShareGateReplacementCharacter = '_'
        AllowDuplicateKeysForDeleteScript = $true
    }

    Permissions = @{
        SourceDocumentLibrariesOnly = $false
        TargetDocumentLibrariesOnly = $false
        IncludeItemPermissions = $true
        ItemProgressInterval = 500
    }

    Output = @{
        SourceFileScans = 'scans\source\files'
        SourcePermissionScans = 'scans\source\permissions'
        TargetFileScans = 'scans\target\files'
        TargetPermissionScans = 'scans\target\permissions'
        FileComparisons = 'comparisons\files'
        PermissionComparisons = 'comparisons\permissions'
        SourceHistoryComparisons = 'comparisons\source-history'
        GeneratedOperations = 'operations\generated'
        Logs = 'logs'
    }
}
