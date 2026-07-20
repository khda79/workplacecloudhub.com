# SmartWorkplaceCMDB Scheduled Orchestrator

`Start-SmartWorkplaceCMDB-OrchestratorScheduledTask-Installer.cmd` launches the
guided installer for unattended daily CMDB collection.

The installer creates:

```text
\WCH\SmartWorkplaceCMDB Orchestrator - <Tenant>
```

The task runs PowerShell 7 with:

```text
-File SmartWorkplaceCMDB-Orchestrator.ps1 -Tenant <Tenant> -Collect
```

The default daily start time is `02:00`. The guided workflow can select another
time in `HH:mm` format.

## Safety model

- The installer is preview-only unless `-Execute` is explicitly confirmed.
- `-StartNow` is disabled by default because it launches a live full collection
  and may publish the resulting CSV files to SharePoint.
- `SYSTEM` and `LocalSystem` are refused.
- The task must run under the dedicated Windows account that owns the
  application certificate in `Cert:\CurrentUser\My`.
- The account also needs Active Directory read access, write access to the
  configured UNC data root, SharePoint access, and read access to this project.
- The password is requested through `Get-Credential`; it cannot be provided as
  a command-line parameter or stored by SmartWorkplaceCMDB.

## Guided installation

Run:

```powershell
.\SmartWorkplaceCMDB\Launchers\Orchestrator\Start-SmartWorkplaceCMDB-OrchestratorScheduledTask-Installer.cmd
```

Review the displayed task definition, then confirm execution when ready.

## Non-interactive preview and installation

Preview:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass `
  -File .\SmartWorkplaceCMDB\Orchestration\Install-SmartWorkplaceCMDB-OrchestratorScheduledTask.ps1 `
  -Tenant prod `
  -ServiceAccount "DOMAIN\svc-cmdb" `
  -DailyAt "02:00"
```

Apply the reviewed plan by adding `-Execute`. Add `-StartNow` only when an
immediate production collection is intended.

Preview removal with `-Uninstall`. Apply removal with `-Uninstall -Execute`.
