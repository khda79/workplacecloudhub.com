# Windows 11 setup source

This folder is a local setup media staging area for local/direct tests or for
preparing media before publishing it to a network share.

For LOT/PsExec runs, configure `Setup source` / `W11UT_SETUP_SOURCE` as a UNC
path reachable by target computers, for example:

```text
\\server\share\Windows11
```

Copy the full contents of a Windows 11 ISO or extracted installation media into
this folder only when you are testing locally or preparing the share content.

You can either copy one ISO directly into this folder, or create one subfolder
per language, for example:

```text
SetupSource\fr-FR
SetupSource\en-gb
```

When `W11UT_SETUP_LANGUAGE=MatchSystem`, the toolkit can select a language
subfolder automatically by reading each subfolder's `sources\lang.ini`.

Required files:

- `setup.exe`
- `sources\install.wim` or `sources\install.esd`
- `sources\lang.ini`

Language handling:

- Default validation is `MatchSystem`: the target computer language must be
  listed in `sources\lang.ini` under `[Available UI Languages]`.
- For a French Windows estate, use a French Windows 11 source such as `fr-FR`.
- For an English UK Windows estate, use an English Windows 11 source such as
  `en-GB`.
- Mixed-language estates should use one LOT per language with a dedicated setup
  source and `W11UT_SETUP_LANGUAGE` value, or a shared `SetupSource` root with
  one valid subfolder per language.
- `W11UT_SETUP_LANGUAGE=Any` disables language matching and should only be used
  intentionally.

LOT/PsExec copy handling:

- The target computer copies the Windows 11 media itself into
  `C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia`.
- For remote LOT runs, set `W11UT_SETUP_SOURCE` to a path reachable by the
  target SYSTEM context, preferably a site-local UNC share.
- The GUI shows `\\server\share\Windows11` as a placeholder and does not send
  that placeholder unless you replace it with a real path.
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
7. Generate the optional SHA256 manifest from the toolkit root:

   ```powershell
   .\Scripts\New-SmartM365SetupMediaManifest.ps1 -MediaRoot .\SetupSource\FR-fr -Force
   ```

   Or generate manifests for every direct media folder under `SetupSource`:

   ```powershell
   .\Scripts\New-SmartM365SetupMediaManifest.ps1 -SetupSourceRoot .\SetupSource -Force
   ```

The manifest is per media folder. Mixed-language sources should have one
`SmartM365-SetupMediaManifest.sha256.csv` inside each language folder.

Do not commit Windows installation media to Git. The toolkit expects this folder
to be populated locally by the operator.
