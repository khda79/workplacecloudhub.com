# Windows 11 setup source

This folder is the default `Setup source` used by the Windows 11 Upgrade LOT
Launcher.

Copy the full contents of a Windows 11 ISO or extracted installation media into
this folder before enabling `Allow setup upgrade`.

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

Recommended preparation:

1. Download the Windows 11 ISO from the official Microsoft download page or from
   your enterprise software distribution source.
2. Mount the ISO on the operator workstation.
3. Copy all files and folders from the mounted ISO root into this `SetupSource`
   folder.
4. Confirm that `setup.exe` exists at the root of this folder.
5. Confirm that `sources\install.wim` or `sources\install.esd` exists.
6. Confirm that `sources\lang.ini` contains the expected language.

Do not commit Windows installation media to Git. The toolkit expects this folder
to be populated locally by the operator.
