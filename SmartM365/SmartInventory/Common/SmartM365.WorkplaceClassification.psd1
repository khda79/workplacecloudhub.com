@{
    RootModule='SmartM365.WorkplaceClassification.psm1'
    ModuleVersion='1.0.0'
    GUID='9b6d2f41-3c7e-4a58-8e1d-5f0a2c4b7e93'
    Author='WorkplaceCloudHub'
    Description='Reads the private SmartWorkplaceIntelligence site and persona classification workbooks and applies the same site and persona rules as the prepared workforce evidence.'
    PowerShellVersion='7.0'
    FunctionsToExport=@(
        'Get-SmartM365WorkplaceClassificationWorkbookName','Get-SmartM365WorkplacePersonaLabel','Test-SmartM365WorkplaceClassificationWorkbook',
        'Receive-SmartM365WorkplaceClassificationWorkbook','ConvertTo-SmartM365ClassificationSearchText','Read-SmartM365WorkplaceSiteClassification',
        'Get-SmartM365WorkplaceSite','Read-SmartM365WorkplacePersonaClassification','Get-SmartM365WorkplacePersona'
    )
    CmdletsToExport=@()
    VariablesToExport=@()
    AliasesToExport=@()
}
