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
