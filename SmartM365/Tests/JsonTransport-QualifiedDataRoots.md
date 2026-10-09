# Qualified local data roots

JSON transport 1.0.6 adds an opt-in writer qualification in the private
`Config/SmartM365-JsonTransport.policy.local.json.txt` configuration:

```json
{
  "QualifiedUncRoots": [],
  "QualifiedSharePointDrives": [],
  "QualifiedLocalDataRoots": [
    { "Root": "C:\\ExampleDataAlias", "Target": "C:\\ExampleSync\\Data" }
  ]
}
```

The configuration contains deployment values, never committed customer paths.
Without this entry, the existing strict reparse-point policy remains in force.
The qualification covers the declared directory's complete subtree, not only
`DATA-POWERBI`. Both alias and physical-target paths are checked against the same
qualification. The alias must be the exact NTFS junction to the declared target;
ordinary/cloud roots can instead declare identical Root and Target values.

The writer inspects native reparse tags, accepting only `IO_REPARSE_TAG_CLOUD`
and `CLOUD_1` through `CLOUD_F` along the qualified target's path and subtree.
Internal junctions, symlinks, unknown tags, overlapping matching roots and a
changed/missing alias target fail closed. Cloud ancestor directories are allowed
only along the explicitly qualified target path. The exact alias's own ancestors
must not contain another redirect. Checks are repeated after acquiring the lock
and before atomic replacement. They are not a security boundary against an
administrator racing filesystem changes: protect policy/root ACLs appropriately.

Hashes, validation, exclusive locks, atomic replacement, retry limits and
last-good-state protection remain in place. Legacy migration, retirement and
consumption deletion retain the strict unlinked-path checks. Prepared-feed batch
retention is not modified. This does not migrate legacy metadata automatically.

Run `Test-SmartM365-JsonQualifiedDataRoot.ps1` in PS7 and Windows PowerShell 5.1,
plus the existing JSON transport and prepared publisher regression suites.
Fixtures are new temporary directories and are deliberately retained. Test
qualification is injected only into the test process's module, not into files.
Real cloud-root inspection must be read-only until publication is separately
approved. Synthetic tests do not prove OneDrive/SharePoint synchronization or
Power BI rendering/refresh readiness.

Native tag references:

- [Microsoft reparse tags](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-fscc/c8e77b37-3909-4fe6-a4ea-2b9d423b1ee4)
- [WIN32_FIND_DATAW](https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-win32_find_dataw)
