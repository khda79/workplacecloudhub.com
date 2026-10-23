# On-Premises Mailboxes

Exchange 2016 local mailbox inventory script with integrated reporting.

## Scripts

- `SmartM365-Exchange-Local-Mailboxes-Inventory.ps1`: detailed mailbox, permissions, statistics, archive, retention, mobile device inventory, and integrated mailbox reporting (`-GenerateReport` / `-ReportOnly`). Use `-DryRun` to validate paths, report CSV schema, and Exchange prerequisites without collecting inventory. Remote mailbox delegation is intentionally disabled by default because it performs two expensive permission queries per remote mailbox; enable it explicitly with `-IncludeRemoteMailboxDelegation`.
- Version 1.41 preserves the 1.40 schema/readiness fields and prevents the standard run from performing remote-mailbox delegation queries unless explicitly requested.
- Version 1.43 preserves the 1.42 schema and publishes current copies, daily
  histories and backup copies atomically through SmartM365.Core compatibility
  1.0.37. A failed current promotion is terminating and leaves the preceding
  `DATA-LAST` file intact.
- Version 1.40 ajoute `InventorySchemaVersion`, conserve les propriétés de readiness de la v1.39 et expose l’état de collecte des gros éléments.
- La collecte des gros éléments est désactivée par défaut car elle est coûteuse. Utiliser `-CollectLargeItemStatistics` pour exécuter une estimation `Search-Mailbox` à 35 Mo et alimenter `LargeItemCount-Over-35MB`, `LargeItemCollectionStatus` et `LargeItemThresholdMB`.
- The existing `EmailAddresses` column remains unchanged for compatibility; the complete prefixed collection, including X500 values, is exported in `EmailAddressesAll`.
- The once-per-day mailbox summary email includes the ProxyAddresses audit from its same-day completed source receipt and matching `DATA-LAST` check/summary CSVs. It includes the proxy metrics, missing-address and duplicate-alias previews, and output paths. When current evidence is unavailable, the email labels the proxy section unavailable without treating an older CSV as today's audit. The collectors remain separate.
- The mailbox `Summary` section displays the same seven inventory totals as KPI cards in two rows (four cards, then three); the domain table and source CSVs retain their existing values.
- Version 1.61 adds `Diff against previous scans` beneath those cards for full, qualified mailbox runs. `Previous` uses the latest earlier qualified day; J-7 and J-30 require the exact date. A new `Exchange_OnPrem_Mailboxes_DailySummary.csv` stores one qualified snapshot per day and replaces that day's row after a complete rerun. Older `DailyStats` rows are not backfilled because multiple runs can share a day. Missing or incompatible snapshots display `n/a`; local size deltas also require complete mailbox statistics in both scans. `ReportOnly`, bounded, and partial runs do not add snapshots.
- `ShowMailLinks` in the script-local JSON configuration controls clickable SharePoint and other file links in the combined email. It defaults to `false`; file paths remain visible as plain text. Existing local JSON files receive the missing key from the template on the next run without replacing their other values. Set it to `true` only when recipients should receive clickable links.

## Launchers

- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory.cmd`: standard production mailbox inventory.
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory-OnlyADPermissions.cmd`: AD permission export only (`-OnlyADPermission`).
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions.cmd`: standard inventory with AD permissions (`-IncludeADPermission`).
## Configuration

### Exchange connection mode (1.70)

`ExchangeConnectionMode` in this collector's local JSON configuration accepts
`Remote` (default) or `LocalSnapin`. It applies to Inventory,
WithADPermissions and OnlyADPermissions. Existing local files receive the missing
key from the template without replacing an explicit value.

Remote opens one dedicated Microsoft.Exchange PowerShell session per run, using
Kerberos and the current Windows identity against the local server's resolved
DNS name. There is no stored credential and no automatic fallback to LocalSnapin.
A connection/import/scope failure stops the run with an explicit error. The
session is closed on completion, error or interruption; its elapsed duration is
logged. Calls remain sequential and no additional sessions or parallel workers
are created per mailbox. ReportOnly opens no Exchange session.

The Exchange /PowerShell endpoint may use NoLanguage. Only imported Exchange
cmdlets and their parameters are sent to it; all filtering, projections and
conversions run on the client. The shared session/value helper is under
Modules/SmartM365.Core, which all three launchers already copy recursively.
Deploy the collector, helper and template together.

Remote size conversion uses the exact byte count in `X GB (n bytes)` rather than
the rounded display size. Unknown values remain `N/A`, zero remains zero,
`Unlimited` remains `Unlimited`, and explicit/conversion errors remain `Error`.
GUID text is parsed on the client. Date display text is retained; an ISO companion
is emitted only for a typed date or unambiguous ISO text with an explicit offset.
No timezone is inferred from culture-dependent or offsetless text. CSV column
names and order are unchanged.

Rollback is per collector: set `"ExchangeConnectionMode": "LocalSnapin"` in
the source collector's runtime `.local.json.txt` (or legacy `.local.json` in
Readers mode), not only in its disposable LauncherCache copy. LocalSnapin keeps
the preceding snap-in loading and native-value processing. It does not fix the
direct MAPI access denial reproduced with a limited service account.

Offline checks mock session creation and Exchange cmdlets, use PSSerializer
round-trips, and never run a collector or connect to an Exchange server. Manual
server qualification must separately check the scheduled account, forest/domain
coverage, sizes/quotas/archives/dates, FullAccess and SendAs, CSV receipts and
session cleanup. A passing offline suite is not deployment or live collection
proof. Do not change the Calendar or Infrastructure collectors as part of this lot.

SharePoint upload is inherited from the tenant/global configuration by default. If `SmartM365-Exchange-Local-Mailboxes-Inventory.local.json` already exists on an Exchange server from an older template, make sure it contains:

```json
"EnableSharePointUpload": "__USE_GLOBAL__"
```

or set it explicitly to `true` for that script.
