#
# Module manifest for SmartM365.SharePoint
# SharePoint Online file upload via Microsoft Graph REST API
#
# Version : 1.0.0
# Author  : Internal IT
#

@{
    # Module identity
    ModuleVersion     = '1.0.2'
    GUID              = '4C5661A1-8D39-4A69-ADE4-E8B6FB16209E'
    Author            = 'Internal IT'
    CompanyName       = 'Internal'
    Copyright         = '(c) Internal IT. All rights reserved.'
    Description       = 'SharePoint Online file upload module via Microsoft Graph REST API. Uses Invoke-RestMethod with a pre-acquired Graph Bearer token. No Microsoft.Graph SDK dependency.'

    # Requirements
    PowerShellVersion = '7.0'

    # Root module
    RootModule        = 'SmartM365.SharePoint.psm1'

    # Exports
    FunctionsToExport = @(
        'Resolve-SmartM365SpDriveId',
        'Invoke-SmartM365SpFileUpload'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Metadata
    PrivateData = @{
        PSData = @{
            Tags        = @('SharePoint', 'Graph', 'SmartM365', 'Upload', 'M365')
            ProjectUri  = ''
            ReleaseNotes = @'
1.0.2
  - Invoke-SmartM365SpFileUpload: ajout parametre ForceUpload (bool, defaut $true).
    La copie est toujours forcee par defaut; passer ForceUpload=$false pour
    retrouver le comportement skip-if-up-to-date.

1.0.1
  - Fix: Get-SpRemoteLastModified — locale-safe date parsing.
    Invoke-RestMethod auto-deserializes ISO 8601 dates to [datetime]; calling .ToString()
    without a format on fr-FR culture swaps day/month (e.g. May 13 -> Dec 5).
    Fix: if $raw is already a [datetime], use it directly instead of round-tripping via string.

1.0.0
  - Initial release.
  - Resolve-SmartM365SpDriveId: resolves SharePoint site + document library drive ID.
  - Invoke-SmartM365SpFileUpload: single-file upload with ForceUpload param,
    small/large dispatch (chunked resumable), timestamp preservation, Logger bridge.
  - No Microsoft.Graph SDK dependency (pure Invoke-RestMethod).
  - Compatible with any script holding a Graph Bearer token.
'@
        }
    }
}
