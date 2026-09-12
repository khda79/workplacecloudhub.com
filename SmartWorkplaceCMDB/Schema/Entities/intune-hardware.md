# Intune hardware collection — V1

The list-only Intune hardware collector uses Microsoft Graph v1.0
`managedDevices` with the existing app-only certificate configuration and
`DeviceManagementManagedDevices.Read.All`. Live execution requires explicit
`-Collect`; validation and fixture modes never authenticate.

## Supported fields

- `ManagedDeviceId` and `AzureAdDeviceId` for existing identity correlation.
- Source-reported serial number, manufacturer and model.
- Source-reported total storage bytes as a nonnegative Int64 value.
- Retrieval time plus tenant, mode, row-count and SHA-256 sidecar evidence.

RAM, IMEI, phone, IP address, asset tag and warranty are excluded. Zero storage
is retained as `ZeroReported`, not interpreted as verified physical capacity.
Missing values remain explicit. Duplicate managed-device IDs fail; repeated
serial numbers are retained and never merged.

## Operation

```powershell
# Offline contract/configuration validation
pwsh -NoProfile -File .\SmartWorkplaceCMDB\Collectors\Intune\SmartWorkplaceCMDB-IntuneHardware-Collect.ps1 -Tenant example -ValidateOnly

# Explicit live collection after tenant and output-scope authorization
pwsh -NoProfile -File .\SmartWorkplaceCMDB\Collectors\Intune\SmartWorkplaceCMDB-IntuneHardware-Collect.ps1 -Tenant example -Collect
```

The orchestrator includes `IntuneHardwareCollect` in `Full` and
`IntuneDevices`; it invokes Graph only when the overall run has explicit
`-Collect`. Output is written to
`DATA-LAST/Raw/Intune/Intune_DeviceHardware.csv`, with history under
`DATA-ALL/Intune/DeviceHardware`.

Consumers must stop on failed/in-progress state, hash/count/tenant mismatch,
duplicate native ID, invalid capacity or unresolved native mapping. Hardware
retrieval time is separate from inventory dates and is not a per-attribute
freshness guarantee.

Synthetic PowerShell 7 and Windows PowerShell 5.1 tests cover projection,
missing/zero values, evidence, isolation, failure recovery, tenant checks and CI
mapping. A prior authorized private snapshot supplied aggregate populated
evidence; the V1 finalization did not rerun the live connector and does not
claim fresh tenant or production qualification.
