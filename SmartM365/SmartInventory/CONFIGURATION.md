# Shared configuration loading

All 51 publication collectors, the inventory orchestrator, WorkplaceEvidence
Prepare and `Invoke-PreparedSharePointRetention.ps1` use the helpers in
`SmartM365/Config/SmartM365-TenantContext.ps1` (1.0.16). Core 1.0.121,
Windows PowerShell 5 compatibility 1.0.70 and EvidenceCollector.Common 1.0.23
use the same helpers. Deploy the complete reviewed runtime together.

## Precedence and locations

Script-local values override tenant values, which override global values. Script
defaults apply last. Existing explicit command-line overrides remain in the entry
scripts. Tenant initialization continues to own identity validation, profile mail
settings and output roots.

Global and tenant files remain in `SmartM365/Config`. `SmartInventoryConfigRootPath`
relocates private inventory files while preserving their relative directory.
Templates stay next to the code. The common reader uses the existing JSON transport
for `.local.json.txt` preference, legacy `.local.json` handling, conflicting names,
UNC qualification, locks, retries and atomic template enrichment. LocalCache does
not introduce another configuration source. An unreadable or malformed private
file cannot silently select a stale adjacent copy or a template.

`false`, `0`, empty arrays and structured objects are explicit values. The markers
`__USE_GLOBAL__` and `USE_GLOBAL` request inheritance. The four upload keys also
inherit for null and blank/whitespace strings. Other settings retain their existing
caller-specific null/blank semantics through explicit compatibility modes.
Local structured settings replace the corresponding effective setting; this is
not a recursive merge of local and tenant business objects. Recursive enrichment
from a template only adds missing keys and preserves existing values.

## API and compatibility adapters

- `Read-SmartM365ScriptConfiguration` accepts an explicit script or configuration
  path. `Create` initializes missing private files from their templates; `Optional`
  leaves an absent file absent; `Template` reads the template without creating a
  private file; `CreateIfTemplate` retains HealthCheck's legacy optional behavior.
  Existing private files receive required template enrichment in all modes.
  The default result preserves nested JSON objects and arrays; `AsHashtable`
  retains the dictionary contract of Prepare and SharePoint farm collectors.
- `Get-SmartM365ConfigurationValue` resolves local/effective/default values and
  tokens. Explicit fallback modes preserve existing non-upload blank/null rules;
  defaults remain literal unless `ResolveDefault` is requested.
- `Merge-SmartM365Configuration` handles premerged configurations. `Overlay`
  preserves Exchange null/blank behavior; `Evidence` preserves the evidence
  runtime's null handling. Upload inheritance always uses the common rule.
- `Resolve-SmartM365ConfigurationTokens` resolves bounded token chains. Strict
  callers reject unresolved tokens. Cycles fail explicitly instead of recursing
  indefinitely. Arrays and objects are not converted to strings.
- `Get-SmartM365ScriptConfigurationContext` provides local/effective settings,
  resolved paths and source metadata. It may initialize/enrich private files
  according to `MissingPolicy`; it is not a read-only audit command. Do not print
  its complete contents: they may contain credentials. It does not initialize
  collector globals or start collection.

Thin functions with historical names remain where callers and offline harnesses
depend on their signatures. They delegate to the common helpers and contain no
independent JSON reader, merge algorithm or token loop. Core's indirect caller
lookup retains the no-file skip before UNC activation. Existing malformed caller
configuration now fails rather than being treated as empty.

Two inconsistent historical failure paths are intentionally tightened: invalid
templates block required enrichment of an existing evidence configuration, and
cyclic token references fail explicitly. No live reload is introduced; the
orchestrator retains its existing configuration lifecycle.

## Independent SharePoint uploads

The common loader supplies the existing `SmartM365-SharePointUploadPolicy.ps1`
resolver; it does not implement another upload policy.

| Setting | Scope |
| --- | --- |
| `EnableSharePointUploadDataAll` | DATA-ALL, WeeklyHistory, associated remote retention |
| `EnableSharePointUploadDataLast` | DATA-LAST CSVs and source receipts |
| `EnableSharePointUploadDataPowerBI` | Automatic prepared batch publication |
| `EnableSharePointUploadLogAll` | Logs and transcripts |

Explicit category values override the legacy `EnableSharePointUpload` fallback.
Existing run guards remain authoritative. `TransferOnly` remains an explicit
prepared transfer even when automatic DATA-POWERBI publication is disabled.
Preparation and maintenance retain their existing retention modes and protections.

## Offline validation

`Test-SmartM365ConfigurationLoaderOffline.ps1` checks the three-tier matrix,
structured values, token chains/cycles, reader policies, enrichment, provenance,
tenant isolation and all 51 collector routes plus infrastructure readers under
PowerShell 5.1 and 7. Existing SharePoint configuration, policy, gate, receipts,
lifecycle, weekly history, relocated-config, JSON retry and prepared-transfer/
retention suites exercise their contracts with synthetic fixtures. These results
do not qualify deployment, production SMB access or actual uploads.
