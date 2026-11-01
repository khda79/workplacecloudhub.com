@{
    RootModule        = 'SmartEndpointMigration.Campaign.Prerequisites.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'ac605bed-cf02-4ec1-913f-e79aa0b2cafe'
    Author            = 'https://github.com/khda79/workplacecloudhub.com'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = 'WorkplaceCloudHub'
    Description       = 'Prerequisites campaign rules of the Smart Endpoint Migration Toolkit (scheduled tasks used by the HybridJoin and Windows11Upgrade campaigns).'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-SemtPrerequisitesProfile')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
