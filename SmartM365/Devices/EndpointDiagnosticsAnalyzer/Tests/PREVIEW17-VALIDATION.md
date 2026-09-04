# Preview17 local validation — 2026-09-04

Scope: corrections following the preview1–preview16 retrospective audit. No Gallery/Intune publication, tenant operation, AI request, Git commit/push or new live-device collection was performed.

## Verified

- 28 focused correctness assertions pass on Windows PowerShell 5.1.19041.6456 and PowerShell 7.6.5. Cases cover malformed/missing evidence, mixed WU success/failure, IME identities and explicit later success, full/numeric EVTX records, neutral upgrade values, collection filename collisions, multi-source MDM parsing, and restricted redacted-text export.
- Existing analysis, performance/cache/coverage and process-timeout regression suites pass on both engines. GUI `-ValidateOnly` passes on both engines.
- Packaging/deployment tests pass on both engines: positive install/detection, corrupt payload rejection, unchanged existing runtime after rejected package, rollback after simulated activation failure, Intune ownership rejection before Gallery update, and owned-fixture uninstall. Application/user data was not removed.
- 18 PowerShell files parse in packaging validation. Changed PowerShell files are signed with the pinned WorkplaceCloudHub certificate and a verified timestamp. Bundle SHA-256 checksums and source/bundle/installed GUI hashes match.

## Real archive replays

The two existing diagnostic archives were replayed on PowerShell 7; input hashes remained unchanged. These are application checks, not a certification of either device's health.

| Evidence | Local archive | Intune archive |
| --- | ---: | ---: |
| Outer ZIP files inventoried/extracted | 215 / 215 | 1,154 / 1,154 |
| EVTX files completely read | 8 / 8 | 18 / 18 |
| Native EVTX records, independently counted | 69,674 | 100,779 |
| Retained logical events | 67,507 | 100,779 |
| Identical duplicate copies, source paths retained | 2,167 | 0 |
| Non-numeric event IDs | 0 | 0 |
| Recognized CMTrace records read | 239,456 | 345,760 |
| Full timeline records | 68,277 | 101,905 |
| Successful replay analysis duration, warm IME/WU caches | 202 s | 252 s |

Durations exclude final CLIXML serialization and GUI import, and are not a promised runtime. The first cold attempts uncovered a generic-list coverage conversion defect; after correction the complete replays passed. The native event count equals retained logical events plus duplicate copies. The latest-300 timeline matches the globally ordered full timeline. Cached IME paths resolve to the current extraction. HTML checks confirm coverage, summary limits and the recorded score; final health presentation was reevaluated with current code.

## Delivery and remaining validation boundaries

Preview17 is installed locally in CurrentUser scope and its GUI was started for the user's manual test. Preview16 is retained in a sibling `.backup-<id>` directory. Automatic updates remain disabled. No scheduled task is created.

- These tests ran on Windows 10; the final preview17 has not been exercised interactively on Windows 11 or deployed through Intune.
- The local collector's copy/timeout contracts are tested with fixtures; a complete new live collection is left to the user.
- Coverage describes supported diagnostic areas. Extracting/indexing a file does not mean every format or every possible fault is interpreted. Internal upgrade indicators are unverified context; vendor hardware-readiness logic is not independently certified across all hardware.
- HTML is a labelled summary. Raw source logs and complete model records remain available; informational IME records without a recognized operation are not all materialized as findings. Unknown timestamps and ambiguous success/failure correlations remain unresolved rather than inferred.
- Redacted-text ZIP export intentionally blocks EVTX/CAB and other unsupported/undecodable content. It is not complete anonymization; outputs require manual privacy review. Nested CAB resource-exhaustion defenses and an exhaustive adversarial parser/security matrix remain further hardening work, not claims of this validation.

The preview remains a test release. Passing these cases is evidence for the corrected contracts, not proof that every diagnostic rule is universally exact.
