# Preview18 validation - 2026-09-04

Scope: collection completeness, categorical collection outcomes, and measured finalization. No Gallery/Intune publication, tenant change, AI request, Git operation or live-device collection was performed by the agent.

## Checks

- Fifteen new synthetic checks pass on Windows PowerShell 5.1.19041.6456 and PowerShell 7.6.5. These retain all 201 eligible logs older than 90 days; distinguish extension exclusions, explicit count/age limits, missing sources and failed copying; preserve errors for access denial, timeout and malformed metadata; exercise the actual registry loop with absent/denied fixtures; assert no production folder-copy count/age arguments; and execute the real complete-result serialization branch.
- The existing 28 correctness assertions and analysis/cache/progress/timeout suites pass on both engines. GUI ValidateOnly passes on both. The WPF render-dispatch call was tested without opening a window.
- Packaging passes on both engines, including nineteen PowerShell files, signature/integrity preflight, temporary installation, detection and guarded uninstall/rollback cases. Automatic updates remain disabled and no scheduled task is created.
- The source retains the same existing non-ASCII content (34 characters); no non-ASCII text was added by these corrections. Static analysis still reports the two pre-existing plaintext-to-SecureString conversion findings at the DPAPI input boundary. The adjacent collection loop no longer assigns to the automatic Error variable. This is not a claim of a warning-free static scan or a full security audit.

## Boundaries

- No record/message limit was introduced for result transfer. Stage 1 exports the complete model; stages 2 and 3 load and render it. Loading/rendering remain synchronous and may pause the GUI; their stage is painted first and duration is logged. Interactive finalization timing awaits the user's next test.
- File-type scope remains explicit. Unavailable sources and partial operations mark incomplete coverage without claiming a device fault. Native command timeouts and archive security limits remain enabled.
- Legacy results.xml remains conservatively interpreted; an older generic failure is not retroactively downgraded by guessing its cause. A new local capture is needed to obtain the new structured collection statuses and files excluded by earlier capture limits.
- These checks run on Windows 10. Interactive Windows 11 and deployed Intune validation are not claimed.
