# Private classification workbooks

For the unified prepared contract, native CMDB coverage and staged server cutover, see [CMDB consolidation](../docs/CMDB-CONSOLIDATION.md). Deployment now requires Python 3.10+ in addition to PowerShell 7. No collector or preparation is run by the offline tests.

The two `*-template.xlsx` files are generic examples. They contain no customer site codes, job titles, or persona mappings. Copy and adapt them outside source control before loading a customer report.

By default, `New-WorkforceIdentityEvidence.ps1` reads `SmartWorkplaceIntelligence-PersonaClassification.xlsx` and `SmartWorkplaceIntelligence-SiteClassification.xlsx` from its `-DataRoot` directory. The paths can also be overridden with `-PersonaClassificationPath` and `-SiteClassificationPath`.

In the site workbook, `Settings > Directory Site Attribute` names the AD user field containing `Directory Site Code` (initially `extensionAttribute13`). Each enabled site code must be unique and have exactly one `Site Type`. Source `HQ=TRUE` maps to `HQ` regardless of its directory path. Other site types must be reviewed against the customer directory before production use.

Persona precedence is: explicit non-human exclusion → review; otherwise `HQ` site → `Headquarters staff`; otherwise a unique job-title rule; otherwise AD `Description` as fallback. Conflicting or unsupported evidence remains for review. The AD description itself is not exported to the report.
