# Current export proof and latest attempt

SmartInventory keeps two small, current-only metadata files per registered producer in DATA-LAST. No CSV history or duplicate snapshot is introduced.

- `SmartInventory_<producer>.current.json.txt` is the last qualified export proof: original acquisition interval, scope, tenant identity, file hashes and logical row counts. Its receipt contract remains 1.2; new proofs declare `PublicationProtocol: 1`.
- `SmartInventory_<producer>.run.json.txt` is the latest attempt, owned by `SmartInventory-SourceRun`, contract 1.0. Its status is Collecting, Publishing, Completed or Failed. It is not a completion proof.

A new attempt does not overwrite the current proof. Consumers may use the previous export during Collecting, or after an error that occurred before canonical publication, only when their existing identity, scope, hash, row-count and freshness checks still pass. Acquisition age is never restarted by a retry or preparation.

The shared atomic publishers set Publishing before the first canonical CSV replacement. `UnqualifiedPublication` remains true until all required current-run outputs have been verified and a new completion proof has been written. A failed or interrupted replacement is therefore rejected; starting another attempt does not clear this fence. A fully qualified replacement clears it. The producer's exclusive collection lock still prevents overlapping runs of the same producer.

Consumers inspect the run state when it exists, including alongside a legacy receipt. A protocol-1 proof requires valid run state. CMDB rechecks publication state, receipt bytes and CSV hashes throughout preparation. Licenses and the on-premises proxy-address email reader also recheck their CSVs and proof after reading. A collecting status is not a claim that the process is still alive.

Existing completed legacy proofs remain readable. Existing Running/Failed proofs cannot be reconstructed into a previous valid proof: the producer must finish a qualified export. First acquisition creates run state only, never fabricated completion evidence.

Failed attempts publish their run state without re-uploading an older proof as new evidence. Successful attempts publish both metadata files using the configured shared upload behavior. Synchronization can deliver files out of order; hashes and proof validation remain mandatory. A local run-state file or local file lock alone is not a distributed lock across SharePoint-synchronized copies.

Orchestrator scheduling and dependency success history are unchanged. CMDB freshness stays 48 hours for normal sources and 240 hours for weekly application evidence, with its existing 168-hour warning. Intelligence source capture and stable history keys are unchanged.

Offline regression commands (no collectors, tenant calls, uploads or live data):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptsOffline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptsOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365CmdbProducerReceipts.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365SourceReceiptReadersOffline.ps1
python -B -m unittest discover -s SmartM365/Tests -p 'test_cmdb*.py'
```
