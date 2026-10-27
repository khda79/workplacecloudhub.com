<#
.SYNOPSIS
    Offline checks for the configurable SharePointMigration storage roots (Config\storage.json.txt).
.DESCRIPTION
    Uses synthetic toolkit and migration folders under TEMP only. Covers default resolution,
    template self-heal, shared roots, launcher config lookup, GUI activity and batch log roots,
    cross-volume archive move and the toolkit.root.txt pointer read by the template launchers.
.VERSION
    1.0.0
#>
[CmdletBinding()]
# CrossVolumeRoot: optional writable folder on another volume than TEMP for the archive copy path.
param([string]$CrossVolumeRoot = '')
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$launchers = Join-Path $project 'Scripts\Launchers'
. (Join-Path $launchers 'SmartM365-SharePointMigration-Storage.ps1')
. (Join-Path $launchers 'SmartM365-SharePointMigration-LauncherCommon.ps1')
. (Join-Path $launchers 'Generic\SmartM365-SharePointMigration-GuiActivity.ps1')
. (Join-Path $launchers 'Generic\SmartM365-SharePointMigration-BatchGui.ps1')

$script:Checks = 0
function Assert-Storage {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}
function Write-StorageJson {
    param([string]$Toolkit, [hashtable]$Values)
    $json = [pscustomobject]$Values | ConvertTo-Json
    [IO.File]::WriteAllText((Join-Path $Toolkit 'Config\storage.json.txt'), $json, [Text.UTF8Encoding]::new($false))
}

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
$root = Join-Path $tempRoot ('SharePointMigration-storage-test-' + [guid]::NewGuid().ToString('N'))
try {
    $toolkit = Join-Path $root 'Toolkit'
    $shared = Join-Path $root 'Shared'
    [void](New-Item -ItemType Directory -Path (Join-Path $toolkit 'Config') -Force)
    Copy-Item -LiteralPath (Join-Path $project 'Config\storage.json.template') -Destination (Join-Path $toolkit 'Config')

    # 1. Default layout and runtime creation from the template.
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $toolkit 'Migrations')) 'default Migrations root'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $toolkit 'ArchivedMigrations')) 'default ArchivedMigrations root'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $toolkit 'Migrations\logs')) 'default logs root'
    Assert-Storage $storage.IsDefault 'default flag'
    Assert-Storage (Test-Path -LiteralPath (Join-Path $toolkit 'Config\storage.json.txt')) 'runtime created from template'

    # 2. Self-heal adds missing template keys without touching existing values.
    [IO.File]::WriteAllText((Join-Path $toolkit 'Config\storage.json.txt'), '{ "MigrationsRoot": "" }')
    [void](Get-SmartM365MigrationStorage -ProjectRoot $toolkit)
    $healed = Get-Content -LiteralPath (Join-Path $toolkit 'Config\storage.json.txt') -Raw | ConvertFrom-Json
    Assert-Storage ($healed.PSObject.Properties['ArchivedMigrationsRoot'] -and $healed.PSObject.Properties['LogsRoot']) 'self-heal added keys'

    # 3. Shared roots; logs follow Migrations when LogsRoot is empty.
    Start-Sleep -Milliseconds 20
    Write-StorageJson $toolkit @{ MigrationsRoot = (Join-Path $shared 'Migrations'); ArchivedMigrationsRoot = (Join-Path $shared 'Archive'); LogsRoot = '' }
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $shared 'Migrations')) 'shared Migrations root'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $shared 'Archive')) 'shared archive root'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $shared 'Migrations\logs')) 'logs follow shared Migrations'
    Assert-Storage (-not $storage.IsDefault) 'configured flag'

    $env:SPMIG_STORAGE_TEST = $shared
    Start-Sleep -Milliseconds 20
    Write-StorageJson $toolkit @{ MigrationsRoot = '%SPMIG_STORAGE_TEST%\Migrations\'; ArchivedMigrationsRoot = 'Old'; LogsRoot = (Join-Path $shared 'Logs') }
    $storage = Get-SmartM365MigrationStorage -ProjectRoot $toolkit
    Assert-Storage ($storage.MigrationsRoot -eq (Join-Path $shared 'Migrations')) 'environment variable and trailing slash'
    Assert-Storage ($storage.ArchivedMigrationsRoot -eq (Join-Path $toolkit 'Old')) 'relative path resolved from the toolkit'
    Assert-Storage ($storage.LogsRoot -eq (Join-Path $shared 'Logs')) 'dedicated logs root'
    Assert-Storage ((Get-SmartM365MigrationLogsRoot -ProjectRoot $toolkit -ChildPath 'gui-activity') -eq (Join-Path $shared 'Logs\gui-activity')) 'logs child path'

    # 4. Consumers: launcher config, GUI activity, batch results.
    $migration = Join-Path $shared 'Migrations\Alpha'
    [void](New-Item -ItemType Directory -Path $migration -Force)
    [IO.File]::WriteAllText((Join-Path $migration 'migration.config.psd1'), "@{ Name='Alpha'; Source=@{Type='SP2019'}; Target=@{Type='SPO'} }")
    $loaded = Import-LauncherMigrationConfig -ProjectRoot $toolkit -MigrationName 'Alpha'
    Assert-Storage ($loaded.Root -eq $migration -and $loaded.Data.Name -eq 'Alpha') 'launcher reads the shared migration'
    $activity = New-SmartM365GuiActivity -ProjectRoot $toolkit -Migration 'Alpha' -Action 'Test'
    Assert-Storage ($activity.StartsWith((Join-Path $shared 'Logs\gui-activity\'), [StringComparison]::OrdinalIgnoreCase)) 'GUI activity under the shared logs root'
    $batchId = '20260101-010101-abcdef12'
    $batchDir = New-Item -ItemType Directory -Path (Join-Path $shared "Logs\comparison-batches\$batchId") -Force
    [IO.File]::WriteAllText((Join-Path $batchDir.FullName 'batch.log'), 'Batch finished: 1 successful')
    $result = Get-SmartM365BatchResult -Root $toolkit -Kind Comparison
    Assert-Storage ($result.Directory -eq $batchDir.FullName) 'batch results under the shared logs root'

    # 5. Invalid configuration fails explicitly instead of silently using the toolkit folders.
    Start-Sleep -Milliseconds 20
    [IO.File]::WriteAllText((Join-Path $toolkit 'Config\storage.json.txt'), '{ invalid')
    $failed = $false
    try { [void](Get-SmartM365MigrationStorage -ProjectRoot $toolkit) } catch { $failed = $_.Exception.Message -like '*Invalid SharePointMigration storage configuration*' }
    Assert-Storage $failed 'invalid storage configuration rejected'

    # 6. Archive move: same volume and forced copy path (verification, source removal).
    $src = New-Item -ItemType Directory -Path (Join-Path $root 'MoveA\Beta\scans') -Force
    [IO.File]::WriteAllText((Join-Path $src.FullName 'a.csv'), 'x,y')
    Move-SmartM365MigrationDirectory -Source (Join-Path $root 'MoveA\Beta') -Destination (Join-Path $root 'MoveB\Beta' | ForEach-Object { [void](New-Item -ItemType Directory -Path (Split-Path $_) -Force); $_ })
    Assert-Storage ((Test-Path -LiteralPath (Join-Path $root 'MoveB\Beta\scans\a.csv')) -and -not (Test-Path -LiteralPath (Join-Path $root 'MoveA\Beta'))) 'same-volume move'
    $crossBase = if ($CrossVolumeRoot) { $CrossVolumeRoot } else { $PSScriptRoot }
    if (-not [string]::Equals([IO.Path]::GetPathRoot($crossBase), [IO.Path]::GetPathRoot($root), [StringComparison]::OrdinalIgnoreCase)) {
        $crossRoot = Join-Path $crossBase ('.storage-move-test-' + [guid]::NewGuid().ToString('N'))
        try {
            [void](New-Item -ItemType Directory -Path $crossRoot)
            Move-SmartM365MigrationDirectory -Source (Join-Path $root 'MoveB\Beta') -Destination (Join-Path $crossRoot 'Beta')
            Assert-Storage ((Test-Path -LiteralPath (Join-Path $crossRoot 'Beta\scans\a.csv')) -and -not (Test-Path -LiteralPath (Join-Path $root 'MoveB\Beta'))) 'cross-volume move'
            Assert-Storage (@(Get-ChildItem -LiteralPath $crossRoot -Force -Filter '.moving-*').Count -eq 0) 'no staging folder left'
        }
        finally { if (Test-Path -LiteralPath $crossRoot) { Remove-Item -LiteralPath $crossRoot -Recurse -Force } }
    }
    else { Write-Host 'SKIP cross-volume move: no folder on another volume (use -CrossVolumeRoot).' }
    $refused = $false
    try { Move-SmartM365MigrationDirectory -Source $migration -Destination (Join-Path $shared 'Migrations') } catch { $refused = $true }
    Assert-Storage $refused 'move refuses an existing destination'

    # 7. Template launcher reads toolkit.root.txt when Migrations lives outside the toolkit.
    $stub = Join-Path $toolkit 'Scripts\Launchers\Generic'
    [void](New-Item -ItemType Directory -Path $stub -Force)
    [IO.File]::WriteAllText((Join-Path $stub 'SmartM365-SharePointMigration-Launcher.ps1'),
        "param([string]`$MigrationName, [string]`$Action) [IO.File]::WriteAllText('$root\launcher-called.txt', `"`$MigrationName|`$Action|`$PSCommandPath`")")
    $pointer = Set-SmartM365MigrationToolkitPointer -MigrationsRoot (Join-Path $shared 'Migrations') -ToolkitRoot $toolkit
    Assert-Storage ((Get-Content -LiteralPath $pointer -Raw).Trim() -eq $toolkit) 'pointer content'
    $launcherDir = New-Item -ItemType Directory -Path (Join-Path $migration 'launchers\scheduled-tasks\files') -Force
    Copy-Item -LiteralPath (Join-Path $project 'Migrations_Template\launchers\scheduled-tasks\files\03-Compare-Files-Scheduled.cmd') -Destination $launcherDir.FullName
    & cmd.exe /d /c ('"{0}"' -f (Join-Path $launcherDir.FullName '03-Compare-Files-Scheduled.cmd')) | Out-Null
    $called = if (Test-Path -LiteralPath "$root\launcher-called.txt") { (Get-Content -LiteralPath "$root\launcher-called.txt" -Raw).Split('|') } else { @() }
    Assert-Storage ($called.Count -eq 3 -and $called[0] -eq 'Alpha' -and $called[1] -eq 'CompareFiles' -and
        $called[2].StartsWith($toolkit, [StringComparison]::OrdinalIgnoreCase)) 'template launcher resolved the toolkit through toolkit.root.txt'

    Write-Host ("PASS Test-SmartM365MigrationStorageOffline: {0} checks" -f $script:Checks)
}
finally {
    Remove-Item Env:SPMIG_STORAGE_TEST -ErrorAction SilentlyContinue
    if ($root.StartsWith($tempRoot + '\', [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $root)) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}
