<#
.SYNOPSIS
    Verify SPO source batch planning without network access.
.VERSION
    1.0.0
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SpoSourceBatch-' + [guid]::NewGuid().ToString('N'))
$batch = Join-Path $PSScriptRoot '..\Scripts\Launchers\Generic\SmartM365-SharePointMigration-TargetScanBatch.ps1'
$launcher = Join-Path $PSScriptRoot '..\Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
try {
    $migration = Join-Path $fixture 'Migrations\Synthetic'
    New-Item -ItemType Directory -Path $migration -Force | Out-Null
    "@{ Name='Synthetic'; Source=@{Type='SPO'; SiteUrl='https://source.example.test/sites/source'}; Target=@{Type='SP2019'} }" |
        Set-Content -LiteralPath (Join-Path $migration 'migration.config.psd1')
    $output = (& $batch -ProjectRoot $fixture -LauncherPath $launcher -Side Source -PlanOnly 6>&1 | Out-String)
    if ($output -notmatch 'PLAN ScanSourceFiles Synthetic' -or $output -notmatch 'PLAN ScanSourcePermissions Synthetic') {
        throw "Source plan omitted expected scans: $output"
    }
    if (Test-Path -LiteralPath (Join-Path $fixture 'Migrations\logs')) { throw 'Plan unexpectedly created logs.' }
    'PASS: SPO source batch plans both scans without executing them.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($fixture)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-SpoSourceBatch-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
