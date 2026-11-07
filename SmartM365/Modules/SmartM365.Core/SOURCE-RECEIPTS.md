# Current export proof and latest attempt

SmartInventory keeps two small, current-only metadata files per registered producer in DATA-LAST. No CSV history or duplicate snapshot is introduced.

- `SmartInventory_<producer>.current.json.txt` is the last qualified export proof: original acquisition interval, scope, tenant identity, file hashes and logical row counts. Its receipt contract remains 1.2; new proofs declare `PublicationProtocol: 1`.
- `SmartInventory_<producer>.run.json.txt` is the latest attempt, owned by `SmartInventory-SourceRun`, contract 1.0. Its status is Collecting, Publishing, Completed or Failed. It is not a completion proof.

A new attempt does not overwrite the current proof. Consumers may use the previous export during Collecting, or after an error that occurred before canonical publication, only when their existing identity, scope, hash, row-count and freshness checks still pass. Acquisition age is never restarted by a retry or preparation.

The shared atomic publishers set Publishing before the first canonical CSV replacement. `UnqualifiedPublication` remains true until all required current-run outputs have been verified and a new completion proof has been written. A failed or interrupted replacement is therefore rejected; starting another attempt does not clear this fence. A fully qualified replacement clears it. The producer's exclusive collection lock still prevents overlapping runs of the same producer.

Consumers inspect the run state when it exists, including alongside a legacy receipt. A protocol-1 proof requires valid run state. CMDB rechecks publication state, receipt bytes and CSV hashes throughout preparation. Licenses and the on-premises proxy-address email reader also recheck their CSVs and proof after reading. A collecting status is not a claim that the process is still alive.

Existing completed legacy proofs remain readable. Existing Running/Failed proofs cannot be reconstructed into a previous valid proof: the producer must finish a qualified export. First acquisition creates run state only, never fabricated completion evidence.

Failed attempts publish their run state without re-uploading an older proof as new evidence. Successful attempts publish both metadata files using the configured shared upload behavior. Synchronization can deliver files out of order; hashes and proof validation remain mandatory. A local run-state file or local file lock alone is not a distributed lock across SharePoint-synchronized copies.

## Staged source publication (Core 1.0.93, helper 1.4.0; artifacts Core 1.0.95, helper 1.4.1)

A producer that calls `Enable-SmartM365StagedSourcePublication` right after starting its receipt never changes DATA-LAST before it completes. Its DATA-LAST CSVs are written to `DATA-LAST\.staging\<producer>\<RunId>\` (same volume, so each promotion is an atomic replace). The run state declares `StagedPublication: true` and stays Collecting during the whole collection.

At completion, a qualified run is verified on its staged copies, then promoted: the run state is set to Publishing with `UnqualifiedPublication`, every CSV replaces its canonical file and is re-hashed, and only then is the new proof written, last. A consumer comparing hashes with the proof during promotion sees a mismatch or the Publishing fence, never a mixed set accepted as valid. If promotion fails part-way, no proof is written and the publication stays unqualified until the next complete run. A failed, cancelled or interrupted run leaves DATA-LAST, the previous proof and its qualification untouched.

A producer can declare a non-CSV DATA-LAST file with `Resolve-SmartM365SourcePublicationPath -Artifact` (for example the license report snapshot). It is staged and promoted after the CSVs and before the proof; an artifact the run did not write keeps its previous copy.

The SharePoint copy and WeeklyHistory of the promoted CSVs, and producer callbacks such as cache manifests, run after the proof; a failure there is a warning. `.staging` is never uploaded. The staging folder is removed at completion; a folder left by a crash is removed at the next start of the same producer, under its exclusive lock. A direct DATA-LAST write during a staged run removes the declaration for the rest of that run. Maintenance scripts that walk DATA-LAST ignore `.staging`.

Opted-in producers: IntuneConfigurations, IntuneManagedApps, IntuneEnrollmentFailures, Intune DiscoveredApps and EXO mailboxes (full and fast runs; permissions-only runs hold no receipt and Top100/DebugUPN diagnostics keep direct suffixed copies) and Licences (licensing CSVs, report CSVs, SKU decisions and report snapshot; the report reads the staged copies of its own run and is uploaded after the promotion) Active Directory (combined, duplicate-identity and daily copies; domain workers and report-only modes hold no receipt) and Exchange hybrid identity issues (issue and summary CSVs; their DATA-LAST SharePoint copies follow the promotion), Microsoft Teams (teams, members, channels and guests CSVs) and SharePoint Online (tenant, sites, lists, permissions and external sharing CSVs, empty exports included). For Teams and SharePoint Online, WeeklyHistory reads the run copy, never the previous DATA-LAST file, and the DATA-LAST SharePoint copies follow the promotion. Intune devices (managed devices, hardware and device inventory CSVs), Entra devices (full, main and HardwareId reports, header-only reports included) and active users (empty export included) publish only through Core, so their DATA-LAST copies, uploads and history follow the promotion. Other producers keep direct publication. The orchestrator reads the declaration: under `FreshSuccess`, a long staged run does not hold its consumers back, while an unstaged one does.

Orchestrator dependency success history is unchanged. CMDB freshness stays 48 hours for normal sources and 240 hours for weekly application evidence, with its existing 168-hour warning. Intelligence source capture and stable history keys are unchanged.

Offline regression commands (no collectors, tenant calls, uploads or live data):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptsOffline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptsOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365CmdbProducerReceipts.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptReadersOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365StagedSourcePublicationOffline.ps1
python -B -m unittest discover -s SmartM365/Tests -p 'test_cmdb*.py'
```
