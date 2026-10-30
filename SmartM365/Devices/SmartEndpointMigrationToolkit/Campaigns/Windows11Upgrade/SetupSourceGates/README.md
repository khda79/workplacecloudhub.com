# Smart Endpoint Migration Toolkit - Windows 11 setup source gates

This folder is a template for the setup source copy gate root of the Windows11Upgrade campaign.

For production Wave/PsExec runs, do not use this repository folder as the active gate root.
Create a dedicated UNC share that target computer accounts can write to, then set the options
(environment variable names shown; the same keys exist in `Wave.local.json.txt` and
`Config\SmartEndpointMigrationToolkit.local.json.txt`):

```text
SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_LIMIT=1
SEMT_WINDOWS11_SETUP_SUBNET_PREFIX_LENGTH=Auto
SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=90
SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_GATE_ROOT=\\server\share\SetupSourceGates
```

The per-source gate uses `SEMT_WINDOWS11_SETUP_SOURCE_CONCURRENCY_LIMIT` and
`SEMT_WINDOWS11_SETUP_SOURCE_CONCURRENCY_GATE_ROOT` in the same way.

The endpoint script uses this setup source gate only around the `robocopy` setup media copy
phase. Windows Setup itself runs outside this gate.

With `SEMT_WINDOWS11_SETUP_SUBNET_PREFIX_LENGTH=Auto`, the target detects the prefix length from
the local interface used to reach the setup source. If detection fails, it falls back to `/24`;
set a numeric value from `1` to `32` to force a specific prefix.

Required permissions on the production share:

- Share permission: Change or Full Control for the target computer accounts, or a group such as Domain Computers.
- NTFS permission: target computer accounts (or that group) need Modify on the UNC gate root,
  because each target creates, refreshes and removes its own slot folders there.
- The technician account does not coordinate the gate; target devices running as SYSTEM do.

Lease behavior:

- One slot directory is created per allowed copy in `SetupSubnetCopy/<subnet-key>/slot-NNN`.
- Each slot contains `lease.json.txt` with computer, run, source, subnet, creation time, and heartbeat time.
- The heartbeat is refreshed during `robocopy`.
- If a device crashes or disappears, another device can remove the stale slot after
  `SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES`.

Intune Win32 packages do not use these gates: the media is delivered by Intune.

Keep this folder empty except for this README. Runtime lease files belong on the production
gate UNC share, not in Git.
