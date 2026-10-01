@{
    RootModule='SmartM365.AccountClassification.psm1'
    ModuleVersion='1.0.0'
    GUID='4c3f6b0e-8d2a-4f17-9b55-2e7a1c9d6f30'
    Author='WorkplaceCloudHub'
    Description='Reads the private account-classification rules (AccountClassification.local.json) shared by the AD enrichment and the prepared workforce evidence.'
    PowerShellVersion='7.0'
    FunctionsToExport=@(
        'Get-SmartM365AccountClassificationDefaultPath','Resolve-SmartM365AccountClassificationPath','Read-SmartM365AccountClassification'
    )
    CmdletsToExport=@()
    VariablesToExport=@()
    AliasesToExport=@()
}
