<#
.SYNOPSIS
Checks the SmartInventory launch catalog and generates the one-line aliases of the legacy launcher names.

.DESCRIPTION
The .cmd launchers were replaced by the catalog (Orchestrator-Jobs.json.template) and the single
entry point SmartInventory.cmd. For one version, every legacy name stays as a one-line alias:
- Launchers\Cloud, Launchers\OnPremises and Launchers\Orchestrator: the LegacyLaunchers of the
  catalog entries, alias "SmartInventory.cmd <Entry>";
- Launchers\Jobs\Start-<Job>.cmd: "SmartInventory.cmd <Job>" (run now on this server);
- LaunchersByOrchestrator\Request-<Job>.cmd: "SmartInventory.cmd <Job> -Request", and
  Request-SmartM365-OrchestratorJob.cmd: "SmartInventory.cmd -Request <Job[,Job...]>".
Job aliases are generated for every non-manual job of the template, as the former launchers were.

With -Check the script writes nothing: it validates the catalog (scripts, parameters by AST,
execution, includes, tokens) and compares the aliases with the folders; it exits 1 on any
difference (used by the offline tests).

.PARAMETER Check
Validates and compares without writing.

.VERSION
1.2.0
#>
#requires -Version 7.2
[CmdletBinding()]
param([switch]$Check)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner
Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Inventory.Launch.psm1') -Force -ErrorAction Stop

$smartInventoryRoot = Split-Path -Path $PSScriptRoot -Parent
$catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $smartInventoryRoot
$differences = [Collections.Generic.List[string]]::new()
foreach ($catalogError in @(Test-SmartM365InventoryCatalog -Catalog $catalog)) { $differences.Add("catalog: $catalogError") }

function Get-AliasContent {
    param([Parameter(Mandatory)][string]$RelativePath, [Parameter(Mandatory)][string]$Arguments)
    $depth = $RelativePath.Split('\').Count - 1
    return ('@call "%~dp0' + ('..\' * $depth) + 'SmartInventory.cmd" ' + $Arguments + ' %*' + "`r`n")
}

$expected = [ordered]@{}
foreach ($entry in @($catalog.Entries)) {
    $legacy = $entry.Definition.PSObject.Properties['LegacyLaunchers']
    if ($null -eq $legacy) { continue }
    foreach ($launcher in @($legacy.Value)) {
        if ($launcher -notmatch '^Launchers\\(Cloud|OnPremises|Orchestrator)\\[A-Za-z0-9._-]+\.cmd$') { $differences.Add("catalog: $($entry.Name): unexpected legacy launcher $launcher"); continue }
        $expected[[string]$launcher] = Get-AliasContent -RelativePath $launcher -Arguments $entry.Name
    }
}
foreach ($entry in @($catalog.Entries | Where-Object { $_.Kind -eq 'Job' } | Sort-Object Name)) {
    $mode = if ($entry.Definition.PSObject.Properties['AssignmentMode'] -and $entry.Definition.AssignmentMode) { [string]$entry.Definition.AssignmentMode } else { 'Legacy' }
    # Same job set as the former generated launchers; the legacy CMDB preparation never had one.
    if ($mode -eq 'Manual' -or $entry.Name -eq 'CmdbEvidence-Prepare') { continue }
    $expected["Launchers\Jobs\Start-$($entry.Name).cmd"] = Get-AliasContent -RelativePath "Launchers\Jobs\Start-$($entry.Name).cmd" -Arguments $entry.Name
    $expected["LaunchersByOrchestrator\Request-$($entry.Name).cmd"] = Get-AliasContent -RelativePath "LaunchersByOrchestrator\Request-$($entry.Name).cmd" -Arguments "$($entry.Name) -Request"
}
$expected['LaunchersByOrchestrator\Request-SmartM365-OrchestratorJob.cmd'] = Get-AliasContent -RelativePath 'LaunchersByOrchestrator\Request-SmartM365-OrchestratorJob.cmd' -Arguments '-Request'

$existing = @()
foreach ($folder in @('Launchers', 'LaunchersByOrchestrator')) {
    $path = Join-Path $smartInventoryRoot $folder
    if (Test-Path -LiteralPath $path) { $existing += @(Get-ChildItem -LiteralPath $path -Filter '*.cmd' -File -Recurse | ForEach-Object { $_.FullName.Substring($smartInventoryRoot.Length + 1) }) }
}
foreach ($relative in $expected.Keys) {
    $path = Join-Path $smartInventoryRoot $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $differences.Add("missing: $relative") }
    elseif ([IO.File]::ReadAllText($path) -cne $expected[$relative]) { $differences.Add("outdated: $relative") }
    if (-not $Check) {
        New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
        [IO.File]::WriteAllText($path, $expected[$relative], [Text.Encoding]::ASCII)
    }
}
foreach ($relative in @($existing | Where-Object { -not $expected.Contains($_) })) {
    $differences.Add("obsolete: $relative")
    if (-not $Check) { Remove-Item -LiteralPath (Join-Path $smartInventoryRoot $relative) -Force }
}

if ($Check) {
    if ($differences.Count -gt 0) {
        $differences | ForEach-Object { Write-Output $_ }
        Write-Output ("SmartInventory catalog or legacy aliases are out of date ({0} difference(s)). Run New-SmartM365-OrchestratorJobLaunchers.ps1." -f $differences.Count)
        exit 1
    }
    Write-Output ("SmartInventory catalog valid; legacy aliases up to date: {0} alias(es)." -f $expected.Count)
    exit 0
}
$catalogErrors = @($differences | Where-Object { $_ -like 'catalog:*' })
if ($catalogErrors.Count -gt 0) { $catalogErrors | ForEach-Object { Write-Output $_ }; exit 1 }
Write-Output ("Legacy aliases generated: {0} alias(es); {1} change(s)." -f $expected.Count, $differences.Count)
