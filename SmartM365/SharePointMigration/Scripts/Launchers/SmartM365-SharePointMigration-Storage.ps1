<#
.SYNOPSIS
    Resolves the SharePointMigration data roots (Migrations, ArchivedMigrations, shared logs).
.DESCRIPTION
    Config/storage.json.txt may point the data roots to a shared folder. Empty values keep
    the historical layout inside the toolkit: Migrations, ArchivedMigrations and
    Migrations\logs. The runtime file is created from, and completed with, its template.
    Windows PowerShell 5.1 compatible; no dependency on other toolkit helpers.
.VERSION
    1.0.0
#>

function Initialize-SmartM365MigrationStorageConfig {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $configRoot = Join-Path -Path $ProjectRoot -ChildPath 'Config'
    $template = Join-Path -Path $configRoot -ChildPath 'storage.json.template'
    $runtime = Join-Path -Path $configRoot -ChildPath 'storage.json.txt'
    if (-not (Test-Path -LiteralPath $template -PathType Leaf)) { return $runtime }
    try {
        if (-not (Test-Path -LiteralPath $runtime -PathType Leaf)) {
            Copy-Item -LiteralPath $template -Destination $runtime -ErrorAction Stop
            return $runtime
        }
        $templateData = Get-Content -LiteralPath $template -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $runtimeData = Get-Content -LiteralPath $runtime -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $added = @()
        foreach ($property in $templateData.PSObject.Properties) {
            if (-not $runtimeData.PSObject.Properties[$property.Name]) {
                $runtimeData | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
                $added += $property.Name
            }
        }
        if ($added.Count) {
            $json = $runtimeData | ConvertTo-Json -Depth 4
            [System.IO.File]::WriteAllText($runtime, $json + "`r`n", [System.Text.UTF8Encoding]::new($false))
            Write-Verbose ("storage.json.txt: added keys {0}" -f ($added -join ', '))
        }
    }
    catch {
        # A read-only toolkit copy must still run; an unreadable runtime file fails in the resolver.
        Write-Verbose ("storage.json.txt self-heal skipped: {0}" -f $_.Exception.Message)
    }
    return $runtime
}

function Resolve-SmartM365MigrationStoragePath {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string]$Default
    )

    $text = if ($null -eq $Value) { '' } else { $Value.Trim() }
    if (-not $text) { $text = $Default }
    $text = [Environment]::ExpandEnvironmentVariables($text)
    if (-not [System.IO.Path]::IsPathRooted($text)) { $text = Join-Path -Path $ProjectRoot -ChildPath $text }
    $full = [System.IO.Path]::GetFullPath($text)
    if ($full.Length -gt 3 -and $full.EndsWith('\')) { $full = $full.TrimEnd('\') }
    return $full
}

function Get-SmartM365MigrationStorage {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $root = [System.IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\')
    $configPath = Initialize-SmartM365MigrationStorageConfig -ProjectRoot $root
    $stamp = ''
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        $stamp = (Get-Item -LiteralPath $configPath).LastWriteTimeUtc.Ticks.ToString()
    }
    $cacheKey = $root + '|' + $stamp
    $cache = Get-Variable -Name SmartM365MigrationStorageCache -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($cache -and $cache.Key -eq $cacheKey) { return $cache.Value }

    $data = $null
    if ($stamp) {
        try { $data = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
        catch { throw "Invalid SharePointMigration storage configuration '$configPath': $($_.Exception.Message)" }
    }
    $values = @{}
    foreach ($name in @('MigrationsRoot', 'ArchivedMigrationsRoot', 'LogsRoot')) {
        $values[$name] = ''
        if ($data -and $data.PSObject.Properties[$name] -and $null -ne $data.$name) { $values[$name] = [string]$data.$name }
    }
    $migrationsRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.MigrationsRoot -Default 'Migrations'
    $archivedRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.ArchivedMigrationsRoot -Default 'ArchivedMigrations'
    $logsRoot = Resolve-SmartM365MigrationStoragePath -ProjectRoot $root -Value $values.LogsRoot -Default (Join-Path $migrationsRoot 'logs')
    $storage = [pscustomobject]@{
        ProjectRoot            = $root
        ConfigPath             = $configPath
        MigrationsRoot         = $migrationsRoot
        ArchivedMigrationsRoot = $archivedRoot
        LogsRoot               = $logsRoot
        IsDefault              = -not ($values.MigrationsRoot.Trim() -or $values.ArchivedMigrationsRoot.Trim() -or $values.LogsRoot.Trim())
    }
    $script:SmartM365MigrationStorageCache = @{ Key = $cacheKey; Value = $storage }
    return $storage
}

function Get-SmartM365MigrationsRoot {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)
    return (Get-SmartM365MigrationStorage -ProjectRoot $ProjectRoot).MigrationsRoot
}

function Get-SmartM365MigrationLogsRoot {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot, [string]$ChildPath = '')
    $logsRoot = (Get-SmartM365MigrationStorage -ProjectRoot $ProjectRoot).LogsRoot
    if ($ChildPath) { return (Join-Path -Path $logsRoot -ChildPath $ChildPath) }
    return $logsRoot
}

function Get-SmartM365DirectoryFingerprint {
    param([Parameter(Mandatory = $true)][string]$Path)
    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction Stop)
    $bytes = [long]0
    foreach ($file in $files) { $bytes += $file.Length }
    [pscustomobject]@{ Files = $files.Count; Bytes = $bytes }
}

# Directory.Move only works inside one volume; shared roots can sit on different shares.
function Move-SmartM365MigrationDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    if (Test-Path -LiteralPath $Destination) { throw "Destination already exists: $Destination" }
    $sameVolume = [string]::Equals([System.IO.Path]::GetPathRoot($Source), [System.IO.Path]::GetPathRoot($Destination),
        [System.StringComparison]::OrdinalIgnoreCase)
    if ($sameVolume) {
        try { [System.IO.Directory]::Move($Source, $Destination); return }
        catch [System.IO.IOException] {
            # Mapped drives or DFS links can share a root text and still be distinct volumes.
            if (Test-Path -LiteralPath $Destination) { throw }
        }
    }
    $expected = Get-SmartM365DirectoryFingerprint -Path $Source
    $staging = Join-Path -Path (Split-Path -Path $Destination -Parent) -ChildPath ('.moving-{0}-{1}' -f
        (Split-Path -Path $Destination -Leaf), [guid]::NewGuid().ToString('N'))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -Recurse -Force -ErrorAction Stop
        $copied = Get-SmartM365DirectoryFingerprint -Path $staging
        if ($copied.Files -ne $expected.Files -or $copied.Bytes -ne $expected.Bytes) {
            throw ("Copy verification failed: {0} files / {1} bytes expected, {2} / {3} copied." -f
                $expected.Files, $expected.Bytes, $copied.Files, $copied.Bytes)
        }
        Rename-Item -LiteralPath $staging -NewName (Split-Path -Path $Destination -Leaf) -ErrorAction Stop
    }
    catch {
        if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
        throw
    }
    try { Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop }
    catch { throw "The folder was copied to '$Destination' but the source could not be removed: $($_.Exception.Message)" }
}

# Per-migration .cmd launchers read this pointer when Migrations lives outside the toolkit.
function Set-SmartM365MigrationToolkitPointer {
    param(
        [Parameter(Mandatory = $true)][string]$MigrationsRoot,
        [Parameter(Mandatory = $true)][string]$ToolkitRoot
    )
    if ($ToolkitRoot -match '[^\x20-\x7E]') { throw "The toolkit path must be ASCII for cmd launchers: $ToolkitRoot" }
    $path = Join-Path -Path $MigrationsRoot -ChildPath 'toolkit.root.txt'
    [System.IO.File]::WriteAllText($path, $ToolkitRoot.TrimEnd('\') + "`r`n", [System.Text.Encoding]::ASCII)
    return $path
}
