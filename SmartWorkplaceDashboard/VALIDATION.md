# Validation — BETA 3.8.0-beta.2, local preparation

Performed 2026-09-10 on the canonical project, with synthetic inputs only. No commit, push, release replacement or site upload was performed.

| Check | Actual result |
| --- | --- |
| Canonical Node build integration | 12 passed, including all 161 cards and 204 textbox definitions, seven DAX variable regressions, deterministic regeneration, exclusion of private/cache files, and strict CSV checks for generator/Desktop metadata |
| Generated-code regression contracts | 23 passed, including absent-input guards for migration and Exchange storage rates |
| Microsoft PQTest 2.155.2 in-memory M execution | 49 passed, zero failures |
| Signed PowerShell validator with AllSigned and synthetic headers | Passed: 90 source tables, two derived tables, 189 measures, 19 pages, 501 visuals |
| Specialized PBIR validator | Zero errors and zero warnings after explicit textbox-padding correction |
| Regression evidence | The new card test fails against beta.1; the PowerShell validator's obsolete property-name assertions were also corrected |
| Desktop opening | Canonical PBIP opened with 19 pages in Desktop 2.157.1354.0 after isolating its incompatible old private cache |
| Desktop local synthetic refresh | XMLA Full refresh followed by successful DAX row counts: six unified users and three unified devices |
| Desktop DAX | All 189 measures Ready and executed in one query; 12 expected KPI values and 10 values across France/Germany contexts matched |
| Desktop render | All 19 pages captured and inspected; labels including the 90d inactivity threshold are visible after correction; empty migration rate displays blank |
| Ratio regression in Desktop | Seven query-scoped synthetic scenarios, 14 expected values; absent/zero/negative denominators, absent free space and valid rates 0/0.75/1 checked |

An earlier opening attempt remained unready (-32502), and native capture failed with E_NOINTERFACE. The later recovery below supersedes that blocked status. The Desktop bridge established the full canonical file path for PID 6196; its Analysis Services child PID 9084 / port 53275 was verified before MCP operations. No other report instance was changed.

The public model has its neutral parameters restored (LocalFolder, C:\SmartM365\DATA, blank ExpectedTenantKey, HistoryMonths=24). No synthetic data, cached model or local data path is distributed. Test preparation now generates only the synthetic fixture and never another PBIP entry point. The experimental local-function wrappers were not integrated.

PowerShell signing is local, self-signed and untimestamped. The beta.2 package/release is not published. Full visual/interaction coverage, CSV/Excel/PDF exports, gateways, SharePoint, permissions and tenant-scale behavior remain unqualified. Neither in-memory tests nor local synthetic Desktop tests qualify a real environment.

## Desktop opening recovery — 2026-09-10

The user supplied a DataModelLoadFailed / PFE_XM_DATATYPE_CONVERSION_FAILED error for AD_Computers_AllDomains.IntuneDiskFreeGB, PhysicalMemoryGB and IntunePrimaryUserMailboxSizeGB. A pre-existing local cache.abf dated 2026-07-23 was present. It was moved intact to a Git-ignored private backup; no cached rows were inspected, exported or refreshed.

Reopening the SAME canonical PBIP without that old cache succeeded on Power BI Desktop 2.157.1354.0. The Desktop API confirmed the exact canonical path and 19 pages, PID 6196. Its local Analysis Services process 9084 was verified as a child of 6196 before MCP connection. MCP reports all three numeric columns as Double / Ready with empty error messages. No column was changed to resolve this opening problem.

This proves recovery from the local opening error without changing those three column types. See Microsoft's documented cache loading behavior: https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-dataset#pbicacheabf.

The subsequent local XMLA Full refresh used 90 synthetic CSVs: 14 fabricated rows across four source tables and 86 header-only sources, with HistoryMonths=0 and an explicit synthetic tenant key. DeviceDetail loaded successfully; the earlier March 2026 Formula.Firewall error was not reproduced on this August engine/path. This does not establish compatibility with the March engine, UI refresh, gateways or SharePoint.

DAX initially failed with SYNTAXERROR: seven primary license/capacity measures used VAR Rows, producing seven SemanticError and eight DependencyError states. Renaming only that variable to _LicenseRows through MCP removed all 15 errors. The same seven expressions were ported to the generator and covered by a build regression. All 189 measures then executed successfully. Expected checks: enabled=4, disabled=1, guests=1, inactive90d=1; E3 purchased=10, consumed=3, available=7, utilization=0.3; invalid E5 purchased/available=null; unified users=6 and devices=3. France/Germany each contain three users and retain global E3 capacity=10. Unknown enabled status is not counted as disabled.

The live model's neutral parameters were restored and read back before serialization to the canonical model.bim. Desktop exported compatibility 1606 and internal rowNumber columns. The validator now supports the observed 1606 and generated 1600 formats, excluding only typed internal rowNumber columns from CSV comparisons. A regression confirms that an ordinary source field with a similar name still fails schema validation. No synthetic path or row values were found in the serialized source model. The generator remains deterministic and emits 1600; a regenerated definition has not been separately reopened in Desktop after this final change.

After explicit zero padding was applied to all 204 textboxes, the canonical report was reloaded through the verified Desktop PID, using reloadModelDefinition=false. All 19 pages were captured and reviewed, with zero PBIR warnings. Textbox scrollbars and the truncated 90d inactivity label were resolved. Horizontal scrolling in wide tables and abbreviated automatic axis titles remain; populated dense visuals and accessibility are not fully qualified.

The visual review demonstrated an invented 100% Migration Success Rate on an empty source. Exchange Storage Used Rate had the same complement-of-blank issue. Both measures were corrected through live MCP, then serialized and ported to the canonical generator. Seven Desktop DAX scenarios passed (14 values): empty, zero, negative denominators; valid rates 0%, 75%, 100%; and absent free space. Query-scoped definitions use the exact expressions read from the live model and synthetic dependency values. Initial dependency-only query overrides did not propagate into precompiled model measures; those unsuccessful harness attempts are not counted as test passes. No source measures were persistently replaced by mocks.

Native exports were not executed: the installed Desktop bridge manifest exposes application.state.get/v1, report.snapshot.capture/v1 and file.reload/v1, with no CSV, Excel, PDF or interactive filter operation. Native computer control is unavailable in this session. No custom CSV/PDF output is substituted as proof. These tests remain an explicit operator qualification item; no tenant access is required to complete them using the existing synthetic fixture.

## Historical beta.1 evidence

# Offline validation — BETA 3.8.0-beta.1

Performed 2026-09-09, only on isolated copies and synthetic inputs.

| Check | Actual result |
| --- | --- |
| Existing PowerShell validator against baseline + 90 synthetic headers | Passed: 90 source + 2 derived tables, 189 measures, 19 pages, 501 visuals |
| Added generated-code regression contracts | 22 failures on baseline; 22 passes after corrections |
| Microsoft PQTest 2.155.2 in-memory M execution | 49 scenarios passed, including all 90 source partition schemas |
| Before/after behavioral comparison | First 42 comparable M scenarios: 25 baseline failures, 17 passes; all 42 pass after fixes |
| Node build integration | 7 passed: clean standalone build, deterministic rebuild, stable version/path traversal/duplicate header/missing-root rejection, preservation after failed build |
| Candidate PowerShell validator | Passed with PowerShell 7.6.5; Node 24.17.0 |
| Source and visual references | Validator checked field/measure references, canvas bounds, 161 KPI labels, country slicers and 16 relationship declarations |

Results are in validation/offline-results.json. The M suite tests conversions, missing/malformed/semicolon CSVs, tenant mismatch, partial flags, timestamps/history, duplicate/missing identity keys, Base64 case, conflicting user/device IDs, unknown activity and mocked local/SharePoint file selection. All source operations are in-memory stubs. These are real M-engine executions, but they do not execute Desktop DAX or render report values.

The Microsoft PQTest test harness was obtained from the official Microsoft.PowerQuery.SdkTools 2.155.2 NuGet package, kept outside the distributable. No credential was entered and no tenant connection was made. To reproduce, install that test dependency in an isolated development location, create a ZIP containing tests/Offline.pq, rename it Offline.mez, generate a query and run PQTest with a separate empty credential-cache path:

```powershell
node ./scripts/New-SyntheticQuery.cjs . ./synthetic.query.pq
PQTest.exe run-test -e ./Offline.mez -q ./synthetic.query.pq -cfp ./empty-credentials.json
```

Check both PQTest Status and every Output[].Passed value; process exit code alone is insufficient. Generated queries are test artifacts and must not replace the public model. PowerShell 5.1 is not a supported validator runtime; #Requires -Version 7.0 makes that prerequisite explicit.

The prepared ZIP is independently extracted and checked against its SHA-256 file manifest. The validator's Authenticode signature is verified in the source and extracted copy, and executed with PowerShell 7 AllSigned. Signing is local, without a timestamp. See the external package manifest for final bytes and checksum; no release/push/publication is performed by these tests.

Unqualified: Power BI Desktop end-to-end refresh, DAX results/filter semantics, interactive charts and table values, CSV/Excel/PDF export, gateway/service/SharePoint connectivity, permissions and tenant-scale performance. No offline result is a real-environment qualification.

Reference semantics: [DAX comparison operators](https://learn.microsoft.com/en-us/dax/dax-operator-reference), [Power Query duplicate handling](https://learn.microsoft.com/en-us/powerquery-m/table-distinct), [PQTest commands](https://learn.microsoft.com/en-us/power-query/sdk-tools/pqtest-commands-options), [Power BI project format](https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-overview).
