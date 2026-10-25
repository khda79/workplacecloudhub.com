# Smart Endpoint Migration Toolkit - Windows 11 setup source

This folder is the local setup media staging area of the Windows11Upgrade campaign. Use it
for local/direct tests, to build Intune Win32 packages, or to prepare media before
publishing it to a network share.

For Wave/PsExec runs, configure the `Windows11Upgrade.SetupSourcePath` option
(`SEMT_WINDOWS11_SETUP_SOURCE`) as a UNC path reachable by target computers, for example:

```text
\\server\share\Windows11
```

Copy the full contents of a Windows 11 ISO or extracted installation media into
this folder only when you are testing locally, building Intune packages or preparing the
share content.

Copy the full extracted contents of one ISO directly into this folder, or create one
subfolder per language. Copying the `.iso` file itself is not a usable setup source. For example:

```text
SetupSource\FR-fr
SetupSource\EN-gb
```

When `Windows11Upgrade.SetupLanguage` (`SEMT_WINDOWS11_SETUP_LANGUAGE`) is `MatchSystem`,
the endpoint can select a language subfolder automatically by reading each subfolder's
`sources\lang.ini`.

Required files:

- `setup.exe`
- `sources\install.wim` or `sources\install.esd`
- `sources\lang.ini`

Language handling:

- Default validation is `MatchSystem`: resolve the Windows installation language from
  registry `InstallLanguage`, falling back to `InstalledUICulture`. That language must be
  listed in `sources\lang.ini` under `[Available UI Languages]`; `SystemLocale` is diagnostic only.
- For a French Windows estate, use a French Windows 11 source such as `fr-FR`.
- For an English UK Windows estate, use an English Windows 11 source such as
  `en-GB`.
- Mixed-language estates should use one Wave per language with a dedicated setup
  source and `SEMT_WINDOWS11_SETUP_LANGUAGE` value, or a shared `SetupSource` root with
  one valid subfolder per language.
- `SEMT_WINDOWS11_SETUP_LANGUAGE=Any` disables language matching and should only be used
  intentionally.

Wave/PsExec copy handling:

- The target computer copies the Windows 11 media itself into
  `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\SetupMedia\<MediaId>-<Language>`
  (for example `SetupMedia\Win11-fr-FR`).
- Media left by the previous Windows11UpgradeToolkit under
  `C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia` is reused by the endpoint before
  any network copy: idle unmanaged media is moved on the same volume after validation, and media
  still owned by a legacy Intune Win32 package is reused in place (never moved or cleaned).
- For remote Wave runs, set `SEMT_WINDOWS11_SETUP_SOURCE` to a path reachable by the
  target SYSTEM context, preferably a site-local UNC share. A setup source map
  (`Templates\SetupSourceMap.csv.template`) can route subnets, IP prefixes or computer name
  prefixes to different shares.
- A local repository path works for local/direct testing, but remote targets
  cannot use the technician workstation path unless it is explicitly shared and
  accessible to them.

Recommended preparation:

1. Download the Windows 11 ISO from the official Microsoft download page or from
   your enterprise software distribution source.
2. Mount the ISO on the operator workstation.
3. Copy all files and folders from the mounted ISO root into this `SetupSource`
   folder or into a language subfolder such as `SetupSource\FR-fr`.
4. Confirm that `setup.exe` exists at the media root.
5. Confirm that `sources\install.wim` or `sources\install.esd` exists.
6. Confirm that `sources\lang.ini` contains the expected language.
7. Generate the SHA256 manifest from the product root:

   ```powershell
   .\Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1 -MediaRoot .\Campaigns\Windows11Upgrade\SetupSource\FR-fr -Force
   ```

   Or generate manifests for every direct media folder under `SetupSource` (this is what
   `New-SetupMediaManifest.cmd` at the product root does without arguments):

   ```powershell
   .\Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1 -SetupSourceRoot .\Campaigns\Windows11Upgrade\SetupSource -Force
   ```

The manifest is per media folder. Mixed-language sources should have one
`SmartM365-SetupMediaManifest.sha256.csv` inside each language folder. The file name is
unchanged from the previous toolkit because the endpoint and the Intune package validate it.

Do not commit Windows installation media to Git. The toolkit expects this folder
to be populated locally by the operator.
