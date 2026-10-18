# SmartM365 GUI demos

Use these local launchers for website screenshots and presentations. Every demo uses fictional Contoso data, displays a demo label and starts without operational authentication or collection. Keep each application folder and its local helpers/assets together.

| Application | Demo launcher | Demo outputs |
| --- | --- | --- |
| Device Reboot Manager | `Devices/DeviceRebootManager/Start-SmartM365-DeviceRebootManager-Demo.cmd` | Restart notification scenarios |
| Inventory Orchestrator | `SmartInventory/Orchestrator/Start-SmartM365-Inventory-Orchestrator-Demo.cmd` | HTML reports and email previews; full shipped planning catalogue |
| Device Registration Tool | `Devices/DeviceRegistrationTool/Start-SmartM365-DeviceRegistration-Demo.cmd` | User/Admin views, support summary, EML preview and synthetic ZIP |
| Endpoint Diagnostics Analyzer | `Devices/EndpointDiagnosticsAnalyzer/Start-SmartM365-EndpointDiagnosticsAnalyzer-Demo.cmd` | All diagnostic tabs, native HTML report and synthetic exports |
| Intune Hybrid Join Toolkit | `Devices/IntuneHybridJoinToolkit/Start-SmartM365-IntuneHybridJoinRepair-Demo.cmd` | Existing/single/new/automatic LOT, options, native HTML and CSV |
| Windows 11 Upgrade Toolkit | `Devices/Windows11UpgradeToolkit/Start-SmartM365-Windows11Upgrade-Demo.cmd` | LOT views, Setup options, native HTML and CSV |
| Exchange Migration Readiness | `Exchange/MigrationReadiness/Start-SmartM365-ExchangeMigrationReadiness-Demo.cmd` | Complete assessment check catalogue, mailbox views, native HTML and CSV |
| Exchange User Notifications | `Exchange/ExchangeUserNotifications/Start-SmartM365-ExchangeUserNotifications-Demo.cmd` | All three campaigns and every shipped language template; summary email preview |
| Intune Remediation | `Intune/Remediation/Start-SmartM365-IntuneRemediation-Demo.cmd` | Local/cloud packages, script editors, simulated lifecycle and execution CSV |
| SharePoint Migration | `SharePointMigration/Start-SmartM365-SharePointMigration-Demo.cmd` | All eight panels, batch view, new migration wizard, farm window and eight HTML views |

Six report-producing applications are covered: Orchestrator, Endpoint Diagnostics Analyzer, the two LOT launchers, Exchange Migration Readiness and SharePoint Migration. Notifications also exports its email templates as HTML previews.

## Capture workflow

1. Open a demo launcher. The normal local WorkplaceCloudHub splash is retained.
2. Choose a scene and navigate the original tabs/options. Business buttons simulate their outcomes; they do not execute operational scripts.
3. Use **Export demo outputs**, then **Open demo outputs** to inspect the generated reports/previews. Each session owns a separate `%TEMP%\SmartM365-Demo\` folder.
4. Capture the desired window or open an HTML preview in a browser. Keep the demo label visible in published examples.

For the eight extended demos, optional arguments include:

```powershell
.\Start-SmartM365-ExchangeMigrationReadiness-Demo.cmd -DemoScenario Issues -DemoView Results -DemoWindowWidth 1800 -DemoWindowHeight 1000
```

Scenes are `Overview`, `Issues` and `Completed`; the reboot pilot uses its own `Recommended`, `Required` and `Postponed` scenes. Window dimensions are adjustable. SharePoint's dashboard defaults to 1800 pixels wide to accommodate its portfolio columns.

Demo edits remain in memory. Generated scripts and mail drafts are sample text; no script, repair, reboot, migration, cloud publication or email delivery is performed. Opening a generated file/folder and clicking the public branding link are explicit local UI actions.

## Local capture storage and refresh

Each application keeps its curated demo material in a Git-ignored `DemoAssets/` directory at its application root (under `Intune/Remediation/GUI/` for Intune Remediation). Open `DemoAssets/index.html` for the local gallery. The folder contains applicable GUI screenshots, report and email previews, standalone HTML files, associated synthetic exports, and website/presentation image variants. `manifest.json.txt` records the capture date, source version/hash and asset hashes.

After every GUI modification, refresh its demo screenshots and the related reports, email previews and website/presentation variants. Shared UI, branding, report-renderer or email-template changes also require refreshing the affected applications. Keep every visible feature represented, use fictional data only, retain the DEMO label and update the manifest. Never present stale captures as current. Do not stage or force-add `DemoAssets/`.

Runtime demo exports still use their isolated temporary session folders. Copy the reviewed outputs into the owning application's `DemoAssets/`; this is the persistent local collection for website and presentation preparation.

## Offline validation

Use an STA PowerShell process and the application's required host (Windows PowerShell 5.1 for the four device demos, PowerShell 7 for the others; SharePoint requires 7.4+).

```powershell
.\Tests\Test-SmartM365GuiDemoApplicationsOffline.ps1 -Application SharePoint -PreviewFolder "$env:TEMP\SmartM365-Demo-QA"
.\Tests\Test-SmartM365GuiDemoApplicationsOffline.ps1 -Application Notifications
.\Tests\Test-SmartM365GuiDemoOffline.ps1 -Application Orchestrator
```

The extended suite covers original named controls, all tabs, scene selection, real WPF button events, native/synthetic exports and email token replacement. It blocks operational commands. `-InteractiveSmoke` briefly displays the demo and closes only that demo window. These are offline/synthetic checks, not tenant or endpoint qualification.

The maintained common host/report helper files live in `Modules/SmartM365.UI/` and are copied identically into the eight application script directories for standalone portability. The tests verify that these copies match. Each local profile reads only shipped neutral view/report/template sources; normal application launchers and operational startup paths remain separate.
