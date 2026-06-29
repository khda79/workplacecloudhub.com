# Windows 11 Upgrade Toolkit Gate Root

This folder is a template for the setup media copy gate root.

For production LOT/PsExec runs, do not use this repository folder as the active gate root. Create a dedicated UNC share that target computer accounts can write to, then set:

```text
W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT=1
W11UT_SETUP_SUBNET_PREFIX_LENGTH=24
W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES=60
W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT=\\server\share\W11UT-Gates
```

The endpoint script uses this gate only around the `robocopy` setup media copy phase. Windows Setup itself runs outside this gate.

Required permissions on the production share:

- Share permission: Change or Full Control for the target computer accounts, or a group such as Domain Computers.
- NTFS permission: Modify for the same target computer accounts or group.
- The technician account does not coordinate the gate; target devices running as SYSTEM do.

Lease behavior:

- One slot directory is created per allowed copy in `SetupSubnetCopy/<subnet-key>/slot-NNN`.
- Each slot contains `lease.json` with computer, run, source, subnet, creation time, and heartbeat time.
- The heartbeat is refreshed during `robocopy`.
- If a device crashes or disappears, another device can remove the stale slot after `W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES`.

Keep this folder empty except for this README. Runtime lease files belong on the production UNC share, not in Git.
