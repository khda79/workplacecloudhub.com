# Refresh and deployment

## Prerequisites

Use the matching WorkplaceCloudHub monorepo with SmartM365 preparation assets,
Python 3.11+ (standard library only), PowerShell 7 and Desktop with PBIP support.
No collector, upload, scheduler change, XMLA refresh or Fabric publication is
performed by this workflow.

## Private working project

Close Desktop before moving/replacing project files. If .local/PowerBI already
exists, use it; never overwrite it with the generic project.

For a new workspace, from the product root:

```powershell
New-Item -ItemType Directory -Path '.local/Config' -Force | Out-Null
if (Test-Path -LiteralPath '.local/PowerBI') {
    throw 'Private Power BI project already exists; initialization refused.'
}
Copy-Item -LiteralPath 'PowerBI/Project' -Destination '.local/PowerBI' -Recurse
if (-not (Test-Path -LiteralPath '.local/Config/refresh.local.json')) {
    Copy-Item -LiteralPath 'Config/refresh.local.json.template' -Destination '.local/Config/refresh.local.json'
}
```

Configure PreparedRoot and all four ExpectedIdentity values independently in
the local JSON. Set the private model's CMDBExpectedIdentity record to that
same known identity, not an identity inferred from incoming evidence.

Set CMDBLicenseReportRoot to the synchronized SmartM365 DATA-LAST directory
containing the four M365_Licenses_Report CSVs and their
M365_Licenses_ReportSnapshot.json.txt. This is separate from PreparedRoot.
CMDBReportDataRoot is retained for compatibility, not as the prepared source.

## Validate, refresh and save

From the product root, substitute the actual Python interpreter:

```powershell
pwsh -NoProfile -ExecutionPolicy AllSigned -File './Launchers/Start-SmartWorkplaceCMDB-Refresh.ps1' -PythonPath 'C:/Path/To/python.exe' -ValidateOnly
```

After validation, run without ValidateOnly and keep the console open. Open
.local/PowerBI/SmartWorkplaceCMDB.pbip. In Desktop, edit CMDBReadBaseUrl,
CMDBReadToken and CMDBReadBatch using the three private session values printed
by the launcher; use anonymous loopback credentials. Apply, refresh natively,
then save after success. Press Enter in the reader console only afterwards.
Never paste session tokens into Git, logs or shared messages.

ValidateOnly verifies the batch, not Desktop refresh success. If a source or
session expires, acquire a valid batch and create a new session; do not bypass
the guards.

## Updates and packaging

Existing private reports are not automatically replaced by public updates.
Preserve private parameters/settings and reconcile changes explicitly while
Desktop is closed.

From the repository root, run:
`python -B -m unittest discover -s SmartWorkplaceCMDB/Tests -p "test_*.py"`.

Release/Files.json is the explicit public allowlist. The signed
Release/SmartWorkplaceCMDB-Package.ps1 packages only those files into a new
output directory. The product ZIP requires the matching monorepo SmartM365
dependencies; it is not a standalone collector. Git publication does not
publish a model to Fabric or grant access to tenant evidence.
