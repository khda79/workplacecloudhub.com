# Hardware reporting — V1

The canonical report presents fleet composition and hardware coverage on the
combined **Fleet & Hardware** page. Source-record equipment and collection-date
evidence are integrated into the hidden **Device 360** drill-through page.

## Source and grain

`Intune_DeviceHardware.csv` contains one source-reported hardware observation
per Intune managed-device identity. `CMDB_CIDeviceHardware.csv` retains the
existing CMDB device key, native source identity, retrieval timestamp and
collection evidence. Serial number, manufacturer, model and total storage bytes
are supported. RAM, asset tag, warranty and inferred physical capacity are not.

## Definitions

- Hardware coverage = distinct filtered devices with a validated hardware
  record / all filtered devices.
- Attribute checks use hardware-record scope. Missing, zero and unknown values
  are distinct states.
- Repeated serial numbers are observations and never automatic merge evidence.
- Devices without a matching record remain visible as uncovered; their missing
  attributes are not fabricated.

Device 360 displays hardware detail only when exactly one device is selected.
Multiple or empty selections keep the gated detail blank.

## Validation boundary

Synthetic collector, contract, CI-adapter and report tests are part of the V1
release. A prior private read-only snapshot supplied aggregate populated
hardware evidence. The V1 finalization did not rerun the live Intune connector,
so that evidence is not represented as fresh tenant or production
qualification. The report and all source data remain private and outside the
release package.
