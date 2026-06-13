@{
    RootModule        = 'SmartCitrix.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '4efa0169-5db6-4ccf-a36c-5d7c415c794a'
    Author            = 'https://github.com/khda79/workplacecloudhub.com'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Shared helpers for SmartCitrix inventory scripts.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Set-SmartCitrixCoreContext',
        'Write-SmartCitrixLog',
        'Import-SmartCitrixPowerShellComponent',
        'ConvertTo-SmartCitrixCompactJson',
        'Get-SmartCitrixObjectPropertyValue',
        'Invoke-SmartCitrixSafeInventoryBlock',
        'Invoke-SmartCitrixSdkCommand',
        'ConvertTo-SmartCitrixFlatRow',
        'ConvertTo-SmartCitrixFlatRows',
        'Export-SmartCitrixCsv',
        'New-SmartCitrixSummaryRow'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
