# Windows 11 Upgrade Toolkit Setup Source Gates

This folder is a template for the setup source copy gate root.

For production LOT/PsExec runs, do not use this repository folder as the active gate root. Create a dedicated UNC share that target computer accounts can write to, then set:

```text
W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT=1
W11UT_SETUP_SUBNET_PREFIX_LENGTH=Auto
W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=90
W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT=\\server\share\SetupSourceGates
```

The endpoint script uses this setup source gate only around the `robocopy` setup media copy phase. Windows Setup itself runs outside this gate.

With `W11UT_SETUP_SUBNET_PREFIX_LENGTH=Auto`, the target detects the prefix length from the local interface used to reach the setup source. If detection fails, it falls back to `/24`; set a numeric value from `1` to `32` to force a specific prefix.

Required permissions on the production share:

- Share permission: Change or Full Control for the target computer accounts, or a group such as Domain Computers.
- NTFS permission: Modify for the same target computer accounts or group.
- The technician account does not coordinate the gate; target devices running as SYSTEM do.

Lease behavior:

- One slot directory is created per allowed copy in `SetupSubnetCopy/<subnet-key>/slot-NNN`.
- Each slot contains `lease.json` with computer, run, source, subnet, creation time, and heartbeat time.
- The heartbeat is refreshed during `robocopy`.
- If a device crashes or disappears, another device can remove the stale slot after `W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES`.

Keep this folder empty except for this README. Runtime lease files belong on the production SetupSourceGates UNC share, not in Git.
