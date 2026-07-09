@{
    RootModule        = 'SmartWorkplaceCMDB.Core.psm1'
    ModuleVersion     = '0.1.1'
    GUID              = 'fe81d6e3-5d5b-4ec0-9c8b-02f82d9bc001'
    Author            = 'WorkplaceCloudHub'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Core helpers for SmartWorkplaceCMDB.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-SmartWorkplaceCMDBProjectRoot',
        'Read-SmartWorkplaceCMDBJsonFile',
        'Resolve-SmartWorkplaceCMDBTenantPath',
        'Initialize-SmartWorkplaceCMDBTenantFolder',
        'Export-SmartWorkplaceCMDBCsv'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}


