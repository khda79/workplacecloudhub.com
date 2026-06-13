@{
    RootModule        = 'SmartAzureVirtualDesktop.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '53b1de8d-6c19-4e28-8cd1-7f33a4ec39b7'
    Author            = 'https://github.com/khda79/workplacecloudhub.com'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Shared helpers for SmartAzureVirtualDesktop inventory scripts.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Set-SmartAvdCoreContext',
        'Write-SmartAvdLog',
        'Import-SmartAvdRequiredModule',
        'ConvertTo-SmartAvdCompactJson',
        'Get-SmartAvdObjectPropertyValue',
        'Get-SmartAvdNestedPropertyValue',
        'Invoke-SmartAvdArmGetPaged',
        'Invoke-SmartAvdSafeInventoryBlock',
        'Export-SmartAvdCsv',
        'Connect-SmartAvdCloudSession',
        'Get-SmartAvdResourceGroupNameFromId',
        'Get-SmartAvdResourceNameFromId',
        'Get-SmartAvdSubscriptionResources',
        'Get-SmartAvdResourceById',
        'Get-SmartAvdChildResources',
        'Get-SmartAvdAgeInDays',
        'ConvertTo-SmartAvdBoolString'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
