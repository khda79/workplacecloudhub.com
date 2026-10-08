# Smart Workplace CMDB

SmartWorkplaceCMDB is a workplace inventory and decision-support solution.
It brings identities, devices, applications, Microsoft 365 licenses and
messaging into one Power BI report, with source lineage and qualification
visible alongside the business indicators.

It helps workplace, identity and operations teams answer practical questions:
which devices are managed, where compliance or lifecycle gaps exist, how AD
and cloud identities relate, how licenses are assigned, and which candidates
deserve a license recovery review. Findings support investigation; they do not
automatically authorize deleting an object or removing a license.

## How it works

1. **SmartM365 SmartInventory collects the evidence.** This product no longer
   runs a second set of AD, Graph or Exchange collectors.
2. **SmartM365 prepares a CMDB batch.** Its preparation contract defines the
   tables, tenant identity, lineage, integrity and source-specific freshness.
3. **The CMDB reader validates and pins that batch.** Power BI receives the
   same retained bytes throughout the read session, even if synchronized
   source files change during the refresh.
4. **Power BI Desktop refreshes the model natively.** The foreground launcher
   keeps the private loopback reader alive; it does not process the model
   through XMLA, collect new evidence, upload data or save the report.

The license recovery pages additionally read the four published SmartM365
license report CSVs and their snapshot descriptor. They represent the
published license report, not a recalculation from an unrelated CMDB batch.

## Report coverage

| Pages | Purpose |
|---|---|
| 01 Executive Overview | Shared workplace, identity, messaging and licensing context |
| 02 Workplace Health | Compliance, data quality and investigation priorities |
| 03 Transformation & Lifecycle | OS adoption, upgrade readiness and Autopilot |
| 04 Licensing & Assignments | License assignments and their observed paths |
| 04A License Overview | Recovery indicators first, followed by license capacity |
| 04B License Recovery | Candidates to review, not automatic removal instructions |
| 04C License Evidence | Identity/mailbox gaps, source freshness and qualification |
| 05 Fleet, Hardware & Apps | Device inventory, hardware and observed applications |
| 06 People & Messaging | Users, account context and mailbox hosting |
| 07 Services & Impact | Observed relationships and available service context |
| 08 Device 360 / 09 User 360 / 10 Group 360 | Focused entity investigation and drillthrough |

## Repository layout

```text
SmartWorkplaceCMDB/
  README.md
  RELEASE.json
  PowerBI/
    Project/       Generic PBIP, PBIR resources and BIM semantic model
    Queries/       Reusable prepared reader and type-conversion functions
  Scripts/         Batch validation, ephemeral reader and refresh guidance
  Launchers/       Signed PowerShell foreground refresh entry point
  Config/          Public configuration template only
  Tests/           Synthetic reader, protocol, date and public-project checks
  Docs/            Current architecture and refresh/deployment instructions
  Release/         Public file allowlist, packager and concise release notes
  .local/          Ignored private Power BI working project and configuration
```

The public project has neutral parameters and **no imported data or model
cache**. Do not refresh it against production directly. Create a private
working copy under `.local/PowerBI`, configure its expected identity and
license report source, and use a validated private reader session.

The private project remains outside Git. Updating the generic project does
not overwrite an existing tenant working copy or its unsaved changes.

## Getting started

- Use this product within the WorkplaceCloudHub repository: its reader needs
  the matching SmartM365 preparation contract, source registry and freshness
  implementation. A standalone product ZIP is not a self-contained collector.
- Copy `Config/refresh.local.json.template` to
  `.local/Config/refresh.local.json`. Configure the prepared-batch directory
  and independently known tenant identity; never infer identity from the
  files being validated.
- Run `Launchers/Start-SmartWorkplaceCMDB-Refresh.ps1 -ValidateOnly` with a
  Python 3.11+ executable. This verifies the batch, not a Power BI refresh.
- Follow [Refresh and deployment](Docs/REFRESH.md) to create a private project,
  configure Power Query and refresh/save in Desktop.

See [Architecture and boundaries](Docs/ARCHITECTURE.md) for the source flows,
weekly application evidence policy and qualification limitations.

## Safety and release status

- This is an analytical inventory, not a transactional CMDB master or an
  instantaneous, perfectly synchronized view of every source.
- Application evidence is weekly. Unresolved relations are retained without
  inventing missing parents; available device context can differ in age.
- Missing or unqualified evidence must not be presented as a verified zero.
- Native refresh, structural tests, Git publication and tenant operational
  acceptance are different checks. The release metadata does not claim
  universal production qualification.
- There is no automatic license removal, device deletion or Fabric deployment.
