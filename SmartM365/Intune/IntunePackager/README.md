# Smart IntunePackager

PowerShell 7/WPF tool that builds Intune Win32 packages (`.intunewin`) with the official Microsoft Win32 Content Prep Tool and publishes them to Microsoft Intune through Microsoft Graph (beta, `win32LobApp`).

The application is standalone: copy the `IntunePackager` folder anywhere and it runs without any other file from this repository. It never assigns apps: assignments are made by the administrator in the Intune admin center.

## Features

- **Packaging**: source folder, setup file and output folder; official `IntuneWinAppUtil.exe` pinned by version and SHA256, downloaded on demand from the official Microsoft GitHub repository, verified for a valid Microsoft Authenticode signature, stored outside the repository. Optional short-path staging for deep source trees. The packaging engine is replaceable (engine registry in `SmartM365.IntunePackager.Packaging.psm1`).
- **Metadata detection**: MSI `ProductCode`, `ProductVersion`, `Manufacturer` and `ProductName` prefill the form and the MSI detection rule. PSAppDeployToolkit packages (`Invoke-AppDeployToolkit.exe` v4, `Deploy-Application.exe` v3) prefill the install and uninstall commands.
- **App definition**: name, description, publisher, version, notes, information and privacy URLs, developer, owner, logo (png or jpg, sent as base64), install and uninstall commands, install behavior (system or user), restart behavior, maximum run time, editable return codes (Intune defaults), architectures (`allowedArchitectures`: x64, arm64, x86, or not restricted), minimum Windows release, minimum free disk space, detection rules (MSI product code, file or folder, registry, custom PowerShell script).
- **Profiles**: the whole definition is saved and loaded as JSON to rebuild and republish a package reproducibly.
- **Publication**: create a Win32 app, or add a new content version to an existing app. Full flow: app, content version, `Detection.xml` encryption data, content file, SAS URI, Azure Storage Put Block / Put Block List with retries and SAS renewal, commit with `fileEncryptionInfo`, commit state, `committedContentVersion`. Graph throttling (429, 503, 504) honors `Retry-After`. Each failure names its step.
- **Update an existing app**: differential update and import from Intune, see [Updating an existing app](#updating-an-existing-app).
- **Result**: app id, direct link to the app page in the Intune admin center and a reminder that the app is not assigned.
- **Logging**: every line starts with `yyyy-MM-dd HH:mm:ss`, in the UI log panel (exportable) and in the log file.

## Requirements

- Windows with PowerShell 7 (the CMD launchers use `%ProgramFiles%\PowerShell\7\pwsh.exe`).
- `Microsoft.Graph.Authentication` 2.35.1 or later: `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`.
- Internet access to `raw.githubusercontent.com` the first time a Content Prep Tool version is used, and to Microsoft Graph and Azure Storage for publication.
- An Intune role that can create and update apps (for example Intune Administrator or Application Manager) for delegated sign-in.
- Pester 5 for the tests only.

## Permissions

Smart IntunePackager uses a single Microsoft Graph permission: **`DeviceManagementApps.ReadWrite.All`**. The read-only CLI import (`-Action Import`) and the diagnostic tool request `DeviceManagementApps.Read.All`; an existing `DeviceManagementApps.ReadWrite.All` grant is accepted instead.

| Sign-in | Permission type | App registration |
| --- | --- | --- |
| Delegated, WAM (default client) | Delegated, consented for Microsoft Graph Command Line Tools | none |
| Delegated, browser | Delegated, admin consent | your own public client app with the `http://localhost` redirect URI (Mobile and desktop applications) |
| Delegated, device code (CLI only) | Delegated | default client or your own app |
| Certificate | Application, admin consent | your own app with a certificate credential; the certificate is in the current user store |

### Recommended GUI sign-in: browser with a dedicated app registration

For the GUI, use `DelegatedSignIn=Browser` with a dedicated app registration. It avoids the Windows account broker (WAM):

- with the default Microsoft Graph PowerShell client, the SDK always signs in through WAM;
- WAM uses the PowerShell console as parent window, and the console is hidden in the GUI, so the sign-in window can open behind the application;
- the native broker can also crash PowerShell at process exit (see Known limitations).

Create the app registration once per tenant (Microsoft Entra admin center, an account allowed to register apps and grant admin consent):

1. **App registrations > New registration**. Name: for example `Smart IntunePackager`. Supported account types: **Accounts in this organizational directory only**. Leave the redirect URI empty and select **Register**.
2. On **Overview**, copy the **Application (client) ID** and the **Directory (tenant) ID**.
3. **Authentication > Add a platform > Mobile and desktop applications**, enter the custom redirect URI `http://localhost`, then **Configure**.
4. Still in **Authentication**, set **Allow public client flows** to **Yes**, then **Save**. No client secret and no certificate are needed.
5. **API permissions**: remove the default `User.Read` permission if your policy prefers no extra permission, then **Add a permission > Microsoft Graph > Delegated permissions > DeviceManagementApps.ReadWrite.All > Add permissions**. Do not add any other permission.
6. **Grant admin consent for <tenant>** and confirm. The status column must show *Granted*.
7. Optional: **Enterprise applications > Smart IntunePackager > Properties > Assignment required = Yes**, then assign the administrators allowed to use the tool.
8. In `SmartM365-IntunePackager-GUI.config.json`, set the tenant profile:

```json
"prod": { "TenantId": "<directory (tenant) id>", "AuthMode": "Delegated", "DelegatedSignIn": "Browser", "ClientId": "<application (client) id>", "CertificateThumbprint": "" }
```

The signed-in administrator still needs an Intune role that can create and update apps; the delegated permission does not replace it. At sign-in the default browser opens on `http://localhost`; complete the sign-in there and return to the application.

## Files

| File | Purpose |
| --- | --- |
| `SmartM365-IntunePackager-GUI.ps1` | Entry point: GUI, CLI and demo |
| `SmartM365.IntunePackager.Core.psm1` | Configuration, logging, profiles, MSI and PSADT discovery, `Detection.xml`, Graph payloads |
| `SmartM365.IntunePackager.Packaging.psm1` | Packaging engines and the verified Win32 Content Prep Tool |
| `SmartM365.IntunePackager.Graph.psm1` | Graph connection, REST requests, upload and publication steps |
| `SmartM365-IntunePackager-GUI.config.json.template` | Committed configuration template |
| `SmartM365-IntunePackager-Demo.ps1` | Fictional data for the offline demo |
| `SmartM365.GuiDemo.ps1`, `SmartM365.GuiDemo.Reports.ps1`, `SmartM365.GuiSplash.ps1`, `WorkplaceCloudHub.ico`, `WorkplaceCloudHub-lockup-WPF.png` | Local copies of the shared SmartM365 UI helpers and branding |
| `Start-SmartM365-IntunePackager-GUI.cmd`, `Start-SmartM365-IntunePackager-Demo.cmd` | Launchers |
| `Examples/` | Fictional sample profiles and the sample MSI generator |
| `Tests/` | Pester 5 tests |
| `BACKLOG.md` | Findings in other components, not fixed by this tool |

## Configuration

The local configuration `SmartM365-IntunePackager-GUI.config.json` is created from the template on first run and is ignored by Git. On every load, keys added to the template are merged into it without overwriting local values, and the added keys are logged.

```json
{
  "Tenants": {
    "test": { "TenantId": "00000000-0000-0000-0000-000000000000", "AuthMode": "Delegated", "DelegatedSignIn": "Auto", "ClientId": "", "CertificateThumbprint": "" }
  }
}
```

The value above is a format example only: replace it with your tenant GUID or tenant domain (`tenant.onmicrosoft.com`). Placeholders, the empty GUID and documentation domains are rejected by validation.

- `-Tenant` selects the profile (default `test`). An empty `TenantId` lets the signed-in account choose the tenant and is logged as a warning.
- `AuthMode`: `Delegated` or `Certificate` (`TenantId`, `ClientId` and `CertificateThumbprint` required).
- `DelegatedSignIn`: `Auto` (browser when `ClientId` is set, otherwise WAM), `Wam`, `Browser` (requires `ClientId`, recommended for the GUI) or `DeviceCode` (CLI only: the code is written to the console, which is hidden in the GUI).
- `Gui.LastTenant`: tenant profile preselected at the next GUI start (saved after a successful connection). `-Tenant` passed explicitly takes precedence. The GUI never connects automatically.
- `ContentPrepTool`: pinned `Version`, official `DownloadUrl`, expected `Sha256` and `StoragePath` (default `%LOCALAPPDATA%\SmartM365\IntunePackager\Tools`). The download URL must point to the official Microsoft repository.
- `Upload`: `BlockSizeMB` (1 to 100, default 16), `MaxRetries`, `RenewSasBeforeExpiryMinutes`, `HttpTimeoutMinutes`.
- `Graph`: `BaseUri`, `MaxRetries`, `PollSeconds`, `PollTimeoutMinutes`. `Paths`: `LogDirectory` (default `Logs`), `DefaultOutputFolder`, `ProfilesFolder`.

Never commit the local configuration, logs, packages or dry-run outputs.

## Usage

Run the commands from the application folder.

GUI:

```powershell
.\Start-SmartM365-IntunePackager-GUI.cmd
```

The GUI opens with a WorkplaceCloudHub splash. Select the tenant, fill or open a profile, detect MSI/PSADT metadata, build, then publish (or select **Dry run**). Packaging, sign-in and publication run in a background runspace; the progress bar and the log panel follow the operation, and **Cancel operation** stops an upload between blocks.

CLI (no window, no splash):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -ProfilePath .\Examples\Contoso-SampleApp.profile.json.template -ValidateOnly
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -ProfilePath .\Examples\Contoso-SampleApp.profile.json.template -Action BuildAndPublish -DryRun
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -ProfilePath .\Examples\Contoso-SampleApp.profile.json.template -Action BuildAndPublish
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -ProfilePath .\Examples\Contoso-SampleApp.profile.json.template -Action Publish -PublishMode UpdateExisting -ExistingAppId <app-id>
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-IntunePackager-GUI.ps1 -Tenant test -Action Import -ExistingAppId <app-id> -ProfilePath .\Profiles\Contoso-SampleApp.profile.json
```

- `-ValidateOnly` checks the configuration, the tenant profile and the app profile. It builds nothing and does not connect. In GUI mode it builds the window without showing it.
- `-DryRun` builds the package and writes every planned Graph request with its payload (secrets redacted, long base64 values truncated) to `<output>\DryRun\<timestamp>\`. Graph is not called. `-WhatIf` has the same effect on publication.
- A real publication asks for confirmation. `-UseDeviceCode` forces a device code sign-in.
- `-Action Import` reads an existing Win32 app (GET only) and writes its definition to `-ProfilePath`, with the detection script and the logo next to it. It signs in with `DeviceManagementApps.Read.All` (`-UsePublishScope` reuses `DeviceManagementApps.ReadWrite.All`). An existing profile is not overwritten without `-Force`.
- `-Force` on an UpdateExisting publication publishes even when the app was modified in Intune after the import.
- `-TestSasRenewal` is a validation option: it requests one SAS renewal after upload block 2, before block 3, blocks being counted from 1 (use a package of at least three blocks).

## Updating an existing app

An UpdateExisting publication adds a new content version and sends a **differential PATCH**:

- always: `committedContentVersion` and the package properties (`fileName`, `setupFilePath`, `msiInformation`);
- only when they differ from the app in Intune: the properties edited by this tool (name, description, publisher, version, notes, URLs, developer, owner, commands, minimum Windows release, minimum free disk space, architectures, return codes, logo). `installExperience` is merged into the current object, so `inUseBehavior` is kept;
- never: what the tool does not edit (`isFeatured`, `allowAvailableUninstall`, scope tags, memory, processor and CPU requirements, `minimumSupportedOperatingSystem`, install scripts, requirement rules).

A changed detection is sent through the `rules` collection together with the current requirement rules, never through `detectionRules` alone. Detection scripts are compared on normalized content (encoding, BOM, line endings). Detection type `Preserved` keeps the detection of the app as it is (used when an imported detection cannot be edited by the form, for example mixed rule types).

After the update, a **post-publication check** reads the app again and logs `[Verify]` warnings when the committed content version, the requirement rules, a preserved property or the detection differ from what was expected.

**Import definition from Intune** (GUI: Publish tab, **Import definition from Intune**; CLI: `-Action Import`) fills the form from the selected app in read-only mode: metadata, commands, install behavior, return codes, architectures, Windows release, disk space and detection. The detection script and logo are saved under `%LOCALAPPDATA%\SmartM365\IntunePackager\Imported\<app id>`. The package source and output folders stay empty. The properties kept as they are in Intune are listed in the form.

**Concurrency**: the import records the app's `lastModifiedDateTime`. If the app was modified in Intune since then, the publication is blocked before any confirmation is asked: the GUI asks to re-import the definition, force the publication or cancel; the CLI stops with exit code 1 unless `-Force` is passed, logs the block once and prints the exact re-import (`-Action Import ... -Force`) and republish (`... -Force`) command lines. After a successful publication, `ImportedFrom` (`lastModifiedDateTime`, content version, display name, `SyncedAt`) is updated with the state read by the post-publication check, in memory and in the profile file when it was imported from the same app (only that section of the file is rewritten), so the next publication with the same profile is not blocked by its own previous update.

The confirmation lists the changes sent to Intune and warns when the new MSI ProductCode differs from the ProductCode detection rule, or when the new package is not an MSI while the app was.

## Sample packages

`Examples/SmartM365-IntunePackager-New-SampleMsi.ps1` generates the fictional "Contoso Sample App" MSI used by the sample profiles. It only writes `HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\IntunePackagerSample\Version`. `-PayloadSizeMB 150` adds an incompressible file for multi-block upload tests; use it with `Contoso-SampleApp-Large.profile.json.template`.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Examples\SmartM365-IntunePackager-New-SampleMsi.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Examples\SmartM365-IntunePackager-New-SampleMsi.ps1 -OutputFolder "$env:LOCALAPPDATA\SmartM365\IntunePackager\Samples\Contoso-SampleApp-Large\Source" -PayloadSizeMB 150
```

## Offline demo mode

```powershell
.\Start-SmartM365-IntunePackager-Demo.cmd
```

The demo shows the original window with fictional Contoso data, the DEMO label and the usual splash. Scenes: `Overview`, `Issues` (validation and throttling messages) and `Completed` (published app id and portal link). Every action is simulated; nothing is built, no sign-in happens and Graph is not called. **Export demo outputs** writes a sample profile, a dry-run payload and a log under `%TEMP%\SmartM365-Demo\`. Offline validation: `SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1 -Application IntunePackager`.

## Screenshots

<!-- Screenshot placeholders: Package tab, App information, Detection, Publish result, Demo overview. -->

## Tests

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\SmartM365-IntunePackager-RunTests.ps1
```

The unit tests are offline (Graph and Azure Storage are mocked). `-IncludeIntegration` also runs the real Content Prep Tool when `IP_TEST_CONTENT_PREP_TOOL` points to a verified `IntuneWinAppUtil.exe`.

## Out of scope

- **Assignments are intentionally excluded**: create them in the Intune admin center. The Graph engine keeps a reserved region for a future shared implementation.
- Planned evolutions: supersedence, dependencies between apps, editing requirement rules (memory, processors, script) and `allowAvailableUninstall`. Existing values are kept by updates.

## Known limitations

- **WAM sign-in window behind the GUI**: `Connect-MgGraph` (Microsoft.Graph.Authentication 2.38.1) has no parameter for a parent window; the SDK uses the PowerShell console window, which the GUI launcher hides. While the sign-in is pending, the GUI shows a banner and flashes in the taskbar; use Alt+Tab or the taskbar to reach the sign-in window, or use `DelegatedSignIn=Browser`.
- **Process exit crash with WAM**: with the default Microsoft Graph PowerShell client, the SDK signs in through the Web Account Manager broker (`msalruntime.dll`). PowerShell can then end with `Fatal error. Internal CLR error. (0x80131506)` during process shutdown, after the work is done. The log records the final result and warns when the broker is loaded. Use `DelegatedSignIn=Browser` with your own app registration, `DeviceCode` or certificate authentication to avoid loading the broker.
- `Browser` sign-in temporarily sets the persisted SDK option `DisableLoginByWAM` when it is not already set, and restores it on disconnect.
- Packaging itself cannot be canceled; uploads and state waits can.
- Minimum Windows release values: Windows 10 up to 21H1 use the release name (`1607` ... `2004`, `2H20` for 20H2, `21H1`); Windows 10 21H2 and 22H2 use `Windows10_21H2` and `Windows10_22H2` (`Windows10_22H2` observed on a tenant, `Windows10_21H2` inferred); Windows 11 uses `Windows11_<release>`. Profiles saved with `21H2` or `22H2` are converted when loaded.
- One operation runs at a time.

## Deviations from SmartM365/RULES.md

These deviations keep the application standalone when copied out of the repository:

| Rule | Deviation | Reason |
| --- | --- | --- |
| Tenant values in `Config/Tenants/*.local.json` through `SmartM365-TenantContext.ps1` | Tenant profiles live in the local app configuration, selected with `-Tenant` (default `test`) | No dependency on `SmartM365/Config` |
| Prefer Graph SDK cmdlets | Only `Microsoft.Graph.Authentication`; every call uses `Invoke-MgGraphRequest` (REST) | One small dependency; same approach as the existing Win32 publishers |
| Logs under `LogAllRootPath` | Logs in the app `Logs` folder (configurable) | No dependency on the shared path layout |
| Teams/e-mail notifications, startup and completion banners | Not implemented | They target inventory scripts and need the shared modules |
| `SmartM365-` file prefix | Modules use `SmartM365.IntunePackager.*.psm1` | Same naming as the local `SmartM365.GuiDemo.ps1` helpers |

## Security

- No secret is stored. Certificate authentication uses a thumbprint from the current user store.
- The Content Prep Tool is used only when its SHA256 matches the pinned value and its Microsoft signature is valid.
- Dry-run payloads redact `fileEncryptionInfo` values and SAS URIs.
- Repository PowerShell files are Authenticode-signed by the repository pre-push hook.
