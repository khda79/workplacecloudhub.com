# Module Rules

Read this file together with the repository root `RULES.md`.

## Scope

These rules apply to shared runtime modules under `Modules`.

## Structure

- Runtime helpers live in `Modules/SmartM365.Core`.
- SharePoint-specific helpers live in `Modules/SmartM365.SharePoint`.
- Windows PowerShell 5.1 compatibility helpers live under `Modules/SmartM365.Core/Compatibility/WindowsPowerShell5`.
- Do not recreate `Modules.PS5` or `Modules.PS7`.

## Module Design

- Put shared behavior in `SmartM365.Core` when it is used by several script families.
- Keep SharePoint upload behavior centralized.
- Keep Graph connection for SharePoint upload centralized so scripts that do not otherwise use Graph can still upload CSV files.
- Do not put customer-specific values, tenant IDs, app IDs, certificate thumbprints, credentials, internal hostnames, or share paths in module source.
- Module local config files may use `__USE_GLOBAL__` and must remain ignored by Git.

## Compatibility

- Keep PowerShell 5.1 compatibility code isolated in the compatibility module.
- Do not add new PowerShell 7-only syntax to the Windows PowerShell 5.1 compatibility module.
