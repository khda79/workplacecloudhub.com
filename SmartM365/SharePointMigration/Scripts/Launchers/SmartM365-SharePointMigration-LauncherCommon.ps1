function Get-LauncherProjectRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LauncherScriptRoot
    )

    return (Resolve-Path -LiteralPath (Join-Path -Path $LauncherScriptRoot -ChildPath '..\..\..')).ProviderPath
}

function Import-LauncherMigrationConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$MigrationName,

        [string]$ConfigFileName = 'migration.config.psd1'
    )

    $migrationRoot = Join-Path -Path $ProjectRoot -ChildPath ("Migrations\{0}" -f $MigrationName)
    $configPath = Join-Path -Path $migrationRoot -ChildPath $ConfigFileName
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "Migration config not found: $configPath"
    }

    [pscustomobject]@{
        Root = $migrationRoot
        Path = $configPath
        Data = Import-PowerShellDataFile -LiteralPath $configPath
    }
}

function Resolve-LauncherMigrationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path -Path $MigrationRoot -ChildPath $Path
}

function Use-DefaultParameterValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        $Value
    )

    if ($BoundParameters.ContainsKey($Name)) {
        return $null
    }

    return $Value
}
