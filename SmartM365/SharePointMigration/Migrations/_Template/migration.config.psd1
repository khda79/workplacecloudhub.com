@{
    Name = 'NewMigration'

    Source = @{
        Type = 'SP2019'
        WebApplicationUrl = 'https://source-sharepoint.example.com'
        UrlsFile = 'migration.config.source.txt'
        PermissionRootPath = '/SOURCE'
    }

    Target = @{
        Type = 'SPO'
        SiteUrl = 'https://yourtenant.sharepoint.com/sites/NewMigration'
        UrlsFile = 'migration.config.target.txt'
        PrefixToRemove = '/sites/NewMigration'
        PermissionRootPath = '/sites/NewMigration'
    }

    Comparison = @{
        MaxScanAgeDifferenceHours = 12
        PermissionMaxScanAgeDifferenceHours = 24
        SizeToleranceBytes = 10240
        ModifiedDateToleranceMinutes = 0
        SourceModifiedTimeZone = 'Local'
        TargetModifiedTimeZone = 'UTC'
        ShareGateReplacementCharacter = '_'
        AllowDuplicateKeysForDeleteScript = $true
    }

    Permissions = @{
        SourceDocumentLibrariesOnly = $false
        TargetDocumentLibrariesOnly = $true
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
