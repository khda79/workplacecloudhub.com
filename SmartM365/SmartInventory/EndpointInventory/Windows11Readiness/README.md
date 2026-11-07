# Windows 11 hardware readiness inventory

This independent, opt-in inventory runs Windows PowerShell 5.1 hardware checks on explicitly named Windows endpoints. It does not alter SWI or Intune eligibility. The target list is the observed scope, not evidence of the entire tenant population.

## Run

Prepare a UTF-8 text file with one authorized DNS computer name per line. Lines beginning with `#` are comments. Run from PowerShell 7 on an authorized inventory host, or from the centralized launcher:

```powershell
.\Launchers\OnPremises\Start-SmartM365-Windows11-HardwareReadiness-Inventory.cmd -Tenant test -ComputerListPath C:\Path\targets.txt
```

Run the launcher from the SmartInventory folder. The effective tenant profile and its paths must exist. The script creates its local configuration from the template on first run. Explicitly choose the correct tenant. `localhost` or the current computer name uses the current local identity. Other targets use authenticated PowerShell remoting and require an existing WinRM setup and sufficient rights. `-Credential` can supply an existing PowerShell credential interactively; credentials are not saved. Remote processing is sequential. The default timeout is 120 seconds per target; `-TimeoutSeconds` changes it. Local execution has no timeout. A remote timeout does not establish that the remote process has stopped.

The official-style hardware test runs in Windows PowerShell 5.1 on the endpoint. The current account must have the permissions needed for TPM and Secure Boot checks, or those checks can be `UNDETERMINED`. The source file is the existing repository copy from EndpointDiagnosticsAnalyzer, redistributed with its Microsoft MIT notice intact. Its exact bytes are pinned by SHA-256 in the inventory entry. The local repository copy has not independently been proven byte-identical to a current Microsoft download. Review provenance and a replacement before updating the pin; do not disable endpoint execution policy or security controls.

## Outputs and qualification

The inventory uses tenant-scoped SmartM365 paths. `Endpoint_Windows11_HardwareReadiness.csv` in DATA-LAST and weekly history are promoted only when every explicitly requested target returned a result. Each run also produces timestamped results and errors in DATA-ALL. If any target fails, timestamped partial evidence and errors remain, the run fails, and DATA-LAST and history remain unchanged. An inaccessible PC is not reported as incompatible. A returned `UNDETERMINED` is valid evidence of an incomplete assessment and yields a warning, not a failure classification.

Rows carry the requested target, observed computer name and domain, scope hash/count, run ID, Microsoft result, reason, raw diagnostic log and source hash. They also retain the historical `pc.*` fields; `pc.CPU` and `pc.CPU_Check` are empty because their historical mapping is unknown. `SourceFile` is the requested target here. `ScopeHash` distinguishes different explicit target lists. Do not join on the short computer name alone. Secure Boot in the Microsoft result describes its compatibility check, not proof that Secure Boot is enabled.

No job is scheduled or registered in the central job catalog. The centralized launcher calls the entry directly and requires the whole SmartM365 tree, not just the launcher. Publication, deployment, permissions, scheduling and remote connectivity must be qualified independently. Test locally first with an authorized administrator and then with a small approved target list.

## Offline test

```powershell
pwsh -NoProfile -File .\EndpointInventory\Windows11Readiness\Tests\Test-SmartM365Windows11ReadinessOffline.ps1
```

This uses fake endpoint results and a fake unreachable target in an isolated temporary tree. It checks promotion and preservation behavior without contacting a PC or cloud service.
