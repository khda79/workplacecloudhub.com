# Intune Detection And Remediation Rules

Read this file together with the repository root `RULES.md`.

These rules apply to Intune detection and remediation scripts under `IntuneRemediation`.

## Packaging

- Intune remediation packages are uploaded as a detection script, an optional remediation script, and package metadata.
- Do not rely on `local.json`, `SmartM365.global.local.json`, modules, or external companion files for these scripts.
- Any configurable value required by Intune must be a safe, generic constant inside the script, with no tenant, customer, credential, certificate, or personal data.
- Active detection/remediation PowerShell script filenames must be prefixed with `SmartM365-`.
- Use explicit package role suffixes such as `SmartM365-<ScenarioName>-Detection.ps1` and `SmartM365-<ScenarioName>-Remediation.ps1`.
- Do not include version numbers in active script file names. Use the script header for the active version and rely on Git history for future changes.

## Privacy And Secrets

- Scripts must not contain customer names, tenant names, domains, app IDs, tenant IDs, certificate thumbprints, passwords, secrets, internal URLs, internal share paths, or personal data.
- Scripts must not collect personal data from devices.
- Script output shown in Intune must stay short and technical. Keep it under the Intune remediation output limit of 2,048 characters.

## Local Output Layout

All logs and generated troubleshooting output must stay under:

```text
C:\ProgramData\SmartM365\IntuneRemediation
```

Use this layout:

```text
C:\ProgramData\SmartM365\IntuneRemediation\Logs\<Scenario>\<ScriptName>.log
C:\ProgramData\SmartM365\IntuneRemediation\Output\<Scenario>\<GeneratedFile>
C:\ProgramData\SmartM365\IntuneRemediation\Temp\<Scenario>\<TemporaryFile>
C:\ProgramData\SmartM365\IntuneRemediation\Tools\<ToolName>\<ToolFile>
```

Do not write SmartM365 remediation logs to:

- `C:\Windows\Temp`
- `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs`
- `C:\ProgramData\Intune\Logs`
- `C:\ProgramData\DeviceOps`

It is acceptable to inspect or clean Windows system paths such as `C:\Windows\Temp` when that is the purpose of the remediation. The rule above only forbids using those paths for SmartM365-generated logs or output.
