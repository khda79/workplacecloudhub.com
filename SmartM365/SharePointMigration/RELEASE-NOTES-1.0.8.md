# Smart SharePoint Migration Toolkit 1.0.8

Release 2026-09-08. GUI 1.0.8; generic launcher 1.0.18.

- GUI actions now use the discovered migration directory even when the report
  label in `migration.config.psd1` differs from the folder name.
- Permission comparisons enforce `PermissionMaxScanAgeDifferenceHours` before
  attempting an Entra cache refresh. Existing interactive YES, warned Force and
  noninteractive rejection behavior is retained.
- PowerShell 7 launchers dispatch SP2016/SP2019 inventory actions to Windows
  PowerShell 5.1, preserving literal paths/switches and child failure codes.
  SPO scans remain in the current host.
- Both Python comparison engines and the PowerShell mapping reader reject extra
  fields and empty mapping files. Invalid lines no longer silently discard data.
  The PowerShell reader releases the file before validation can throw.
- Documentation covers GUI startup, source-only public distribution at audit
  time, manual updates, per-operation access, evidence limits and cleanup review.

Offline validation covers migration-folder selection, mapping syntax, permission
scan age thresholds and overrides, actual child-process host routing through
synthetic inventory scripts, file versions, older target dates, duplicate keys,
scope filtering and guarded cleanup generation. GUI resources load under Windows
PowerShell 5.1 and PowerShell 7 with `-ValidateOnly`.

No tenant authentication, real inventory, migration, deletion or site
administration was performed during validation. SP2016/SP2019 farm access, real PnP/Graph permission
coverage, throttling and a complete interactive operator workflow still require
an authorized environment pilot. Scan output existence is not proof of complete
coverage. This is the first dedicated release; no previously published asset is replaced.
