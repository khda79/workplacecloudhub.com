@{
    RootModule        = 'SmartAzure.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a51c769d-c5b8-4e55-915d-8c5af2f4cb68'
    Author            = 'Internal automation team'
    CompanyName       = 'Internal'
    Copyright         = '(c) Internal automation team. All rights reserved.'
    Description       = 'Core SmartAzure automation tools: logging, module preflight, Azure session, CSV export, and inventory helpers.'
    FunctionsToExport = @(
        'Set-SmartAzureCoreContext',
        'Write-SmartAzureLog',
        'Import-RequiredModule',
        'Test-SmartAzureFileLocked',
        'Remove-SmartAzureOldFiles',
        'ConvertTo-CompactJson',
        'Get-ObjectPropertyValue',
        'Get-SmartAzureNestedPropertyValue',
        'Invoke-SmartAzureArmGetPaged',
        'Invoke-SafeInventoryBlock',
        'Export-SmartAzureCsv',
        'Export-SmartAzureCsvFromConvert',
        'Invoke-SmartAzurePreflight',
        'Connect-SmartAzureCloudSession'
    )
    CmdletsToExport   = @()
    VariablesToExport = '*'
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('SmartAzure', 'Inventory', 'Azure')
            ProjectUri = 'https://github.com/khda79/workplacecloudhub.com'
        }
    }
}
