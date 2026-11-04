<#
.SYNOPSIS
    Exercise launcher identity mapping preflight with synthetic offline inventories.
.VERSION
    1.0.2
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
$toolkit = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CrossTenantPreflight-' + [guid]::NewGuid().ToString('N'))
try {
    foreach ($directory in @('Migrations\Synthetic\scans', 'Migrations\Synthetic\target-scans', 'Scripts\Launchers\Generic', 'Scripts\Compare', 'Scripts\Inventory')) {
        New-Item -ItemType Directory -Path (Join-Path $root $directory) -Force | Out-Null
    }
    foreach ($file in @('SmartM365-SharePointMigration-LauncherCommon.ps1', 'SmartM365-SharePointMigration-Storage.ps1')) {
        Copy-Item -LiteralPath (Join-Path $toolkit "Scripts\Launchers\$file") -Destination (Join-Path $root 'Scripts\Launchers')
    }
    $launcher = Join-Path $root 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1') -Destination $launcher
    Get-ChildItem -LiteralPath (Join-Path $toolkit 'Scripts\Compare') -Filter '*.py' -File | Copy-Item -Destination (Join-Path $root 'Scripts\Compare')
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\console_lifecycle.py') -Destination (Join-Path $root 'Scripts')
    $migration = Join-Path $root 'Migrations\Synthetic'
    @'
@{
 Name='Synthetic'
 Source=@{Type='SPO'; SiteUrl='https://source.example.test/sites/a'; PermissionRootPath='/sites/a'; AuthConfigPath='source.local.psd1'}
 Target=@{Type='SPO'; SiteUrl='https://target.example.test/sites/a'; PermissionRootPath='/sites/a'; AuthConfigPath='target.local.psd1'}
 Comparison=@{CrossTenantIdentityMappingPath='identity-map.csv'; PathMappingsFile='mapping.txt'; ShareGateReplacementCharacter='_'; EntraUsersCacheEnabled=$true}
 Permissions=@{SourceDocumentLibrariesOnly=$false; TargetDocumentLibrariesOnly=$false}
 Output=@{SourcePermissionScans='scans'; TargetPermissionScans='target-scans'; PermissionComparisons='comparisons'; Logs='logs'}
}
'@ | Set-Content -LiteralPath (Join-Path $migration 'migration.config.psd1') -Encoding UTF8
    "@{ TenantId='11111111-1111-1111-1111-111111111111' }" | Set-Content -LiteralPath (Join-Path $migration 'source.local.psd1') -Encoding UTF8
    "@{ TenantId='22222222-2222-2222-2222-222222222222' }" | Set-Content -LiteralPath (Join-Path $migration 'target.local.psd1') -Encoding UTF8
    'https://source.example.test/sites/a https://target.example.test/sites/a' | Set-Content -LiteralPath (Join-Path $migration 'mapping.txt') -Encoding UTF8
    @'
param([string]$OutputPath,[string]$TenantHost,[string]$MaxCacheAgeHours,[string]$TenantId,[switch]$Connect,[switch]$InteractiveAuth,[switch]$DeviceLogin)
throw 'Cross-tenant comparison must not query the target directory cache.'
'@ | Set-Content -LiteralPath (Join-Path $root 'Scripts\Inventory\SmartM365-EntraUsers-CacheExport.ps1') -Encoding UTF8
    $inventory = Join-Path $migration 'scans\SPO-PermissionInventory-Synthetic-20261104-100000.csv'
    @'
WebUrl,ObjectScope,ObjectServerRelativeUrl,PermissionLevels,PrincipalType,PrincipalLoginName
https://source.example.test/sites/a,Web,/sites/a,Read,User,one@source.example.test
https://source.example.test/sites/a,Web,/sites/a,Read,User,two@source.example.test
'@ | Set-Content -LiteralPath $inventory -Encoding UTF8
    $mapping = Join-Path $migration 'identity-map.csv'
    @'
SourceWebUrl,SourcePrincipalType,SourcePrincipalLoginName,TargetWebUrl,TargetPrincipalType,TargetPrincipalLoginName,Status
https://source.example.test/sites/a,User,one@source.example.test,https://target.example.test/sites/a,User,one@target.example.test,Confirmed
'@ | Set-Content -LiteralPath $mapping -Encoding UTF8
    $originalHash = (Get-FileHash -LiteralPath $inventory -Algorithm SHA256).Hash
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ValidateCrossTenantIdentityMappings -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0 -or $output -notmatch 'unresolved source rows') { throw "Incomplete mapping was not rejected: $output" }
    $reports = @(Get-ChildItem -LiteralPath (Join-Path $migration 'identity-mapping-reviews') -Filter '*.csv')
    if ($reports.Count -ne 1 -or @(Import-Csv -LiteralPath $reports[0].FullName).Count -ne 1) { throw 'Unresolved evidence was not saved.' }
    Add-Content -LiteralPath $mapping -Encoding UTF8 -Value 'https://source.example.test/sites/a,User,two@source.example.test,https://target.example.test/sites/a,User,two@target.example.test,Confirmed'
    $target = Join-Path $migration 'target.csv'
    @'
WebUrl,ObjectScope,ObjectServerRelativeUrl,PermissionLevels,PrincipalType,PrincipalLoginName
https://target.example.test/sites/a,Web,/sites/a,Read,User,one@target.example.test
'@ | Set-Content -LiteralPath $target -Encoding UTF8
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ValidateCrossTenantIdentityMappings -TargetCsv $target -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0 -or $output -notmatch 'Permissions and group membership were not compared' -or $output -notmatch '"MappedSourceRowsWithTargetPrincipalObserved": 1') { throw "Complete mapping preflight failed: $output" }
    $targetInventory = Join-Path $migration 'target-scans\SPO-PermissionInventory-Synthetic-20261104-100000.csv'
    Copy-Item -LiteralPath $target -Destination $targetInventory
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ComparePermissions -SourceCsv $inventory -TargetCsv $targetInventory -Force -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Cross-tenant permission comparison failed: $output" }
    $summaryPath = Get-ChildItem -LiteralPath (Join-Path $migration 'comparisons') -Recurse -Filter 'Synthetic-Summary.csv' | Where-Object { $_.DirectoryName -notmatch 'DocumentLibraries|OtherScopes' } | Select-Object -First 1
    if (-not $summaryPath) { throw 'Comparison summary was not created.' }
    $summary = Import-Csv -LiteralPath $summaryPath.FullName
    if ($summary.CrossTenantIdentityMode -ne 'ExplicitMapping' -or $summary.MatchedPermissions -ne '1' -or $summary.MissingInSPO -ne '1' -or $summary.ValidationStatus -ne 'ReviewNeeded') { throw 'Cross-tenant summary did not preserve mapped and missing assignments with review status.' }
    $mappingLines = @(Get-Content -LiteralPath $mapping)
    $mappingLines[0..1] | Set-Content -LiteralPath $mapping -Encoding UTF8
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ComparePermissions -SourceCsv $inventory -TargetCsv $targetInventory -Force -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0 -or $output -notmatch 'unresolved source rows') { throw 'Incomplete cross-tenant mapping was not blocked before comparison.' }
    $topSummaries = @(Get-ChildItem -LiteralPath (Join-Path $migration 'comparisons') -Recurse -Filter 'Synthetic-Summary.csv' | Where-Object { $_.DirectoryName -notmatch 'DocumentLibraries|OtherScopes' })
    if ($topSummaries.Count -ne 1) { throw 'Incomplete mapping published a new comparison summary.' }
    if ((Get-FileHash -LiteralPath $inventory -Algorithm SHA256).Hash -ne $originalHash) { throw 'Source inventory was modified.' }
    'PASS: unresolved and complete mapping preflight through the launcher; no tenant access.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($root)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-CrossTenantPreflight-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
