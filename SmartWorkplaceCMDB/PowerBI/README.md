# Smart Workplace CMDB — prepared snapshot refresh

## One-time setup

1. Use the migrated private SmartWorkplaceCMDB.pbip project. This repository does
   not redistribute that tenant model or its cache.
2. Copy Config/refresh.local.json.template to Config/refresh.local.json. Set the
   absolute synchronized DATA-POWERBI-CMDB path and the four independently known
   expected identity values. Never derive the expected tenant from the batch
   being validated. The local configuration is ignored by Git.
3. Install Python 3.11+ or supply its executable with -PythonPath. Keep the
   matching SmartM365 prepared contract, registry and cmdb_freshness.py in the
   same repository checkout. No permanent Windows service is installed.

From the repository root, test without starting a listener or editing the model:

~~~powershell
pwsh -NoProfile -ExecutionPolicy AllSigned -File .\SmartWorkplaceCMDB\Launchers\Start-SmartWorkplaceCMDB-Refresh.ps1 -ValidateOnly
~~~

## Every refresh

Run the same command without -ValidateOnly. Initial validation may take several
minutes and substantial RAM for application relations: the 4 GiB default is a
retained-binary budget, not a total process-memory limit. Never truncate a batch.

When ready, the console prints private CMDBReadBaseUrl, CMDBReadToken and
CMDBReadBatch values. In Desktop, use Transform data → Edit parameters to set
those three text parameters. Apply, then perform the native Desktop refresh.
If prompted, use anonymous credentials for the loopback URL. Leave the console
open during the complete refresh, then save after success and press Enter in
the console to stop the reader. Ctrl+C also stops it. Do not record the token in
shared transcripts, screenshots, public logs or Git.

There is no automatic Desktop parameter write, model export or XMLA processing.
The launcher does not claim refresh success or save the PBIP. An old session
stored in the project is intentionally unusable after stopping its reader.
A new refresh requires a new validation and all three new values. Do not click
Refresh after closing the reader; do not substitute File.Contents queries.

The lifetime is at most two hours and can end earlier at the earliest source
expiry. If it expires mid-refresh, reject that refresh and start a fresh session
with a valid prepared batch. A recent preparation does not refresh old inputs.
DiscoveredApps targets a weekly 168-hour age, warning thereafter, with a hard
240-hour limit; core source evidence retains its 48-hour limit.

## Qualification boundaries

The privately migrated model loaded all 46 prepared tables with matching row
counts. Eight opt-in wall-clock date mappings and the one Autopilot contact
sentinel were verified using native M tests and a targeted native Desktop
refresh. Strict UTC parsing is unchanged elsewhere. No source files were edited.

Four License Report tables independently read the published SmartM365 licensing
snapshot from DATA-LAST. Retain that path/configuration; their freshness and
publication identity are separate from the prepared batch. The 13-page report's
final functional page review was waived, not passed. This launcher does not
publish a report to Fabric, guarantee unattended scheduled refresh or restart
Power BI after a crash.

## Offline checks

~~~powershell
python -B -m unittest discover -s SmartWorkplaceCMDB/Tests -p 'test_prepared*.py'
python -B -m unittest discover -s SmartWorkplaceCMDB/Tests -p test_refresh_prepared_report.py
python -B -m unittest discover -s SmartWorkplaceCMDB/Tests -p test_release_package.py
~~~

Tests/prepared_dates.query.pq is a synthetic native-M query against
Queries/prepared-types-candidate.pq. Run with the Microsoft Power Query SDK
PQTest run-compare, an inert test extension, -q for the query and -pa for the
function. This is offline parsing proof, not a real producer or tenant run.

Historical build_report.py, report_360.py, report_cockpit.py, report_hardware.py
and restyle_report.py are developer design utilities using legacy fixtures.
They are not a supported refresh/rebuild route for the migrated private model.
