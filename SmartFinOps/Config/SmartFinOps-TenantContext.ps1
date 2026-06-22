function ConvertTo-SmartFinOpsHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = [ordered]@{}
    if ($null -eq $InputObject) { return $hash }
    foreach ($property in $InputObject.PSObject.Properties) { $hash[$property.Name] = $property.Value }
    return $hash
}

function Read-SmartFinOpsJsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) { throw "Configuration file not found: $Path" }
        return [ordered]@{}
    }

    try { return ConvertTo-SmartFinOpsHashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) }
    catch { throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message) }
}

function Find-SmartFinOpsRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    $searchRoot = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).Path } else { $StartPath }
    $resolvedSearchRoot = Resolve-Path -LiteralPath $searchRoot -ErrorAction SilentlyContinue
    if ($null -ne $resolvedSearchRoot) { $searchRoot = $resolvedSearchRoot.Path }

    while ($searchRoot) {
        if ((Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartFinOps.global.local.json')) -or
            (Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartFinOps.global.local.json.template'))) {
            return $searchRoot
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw "SmartFinOps root not found from '$StartPath'."
}

function Test-SmartFinOpsWritableDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        $probePath = Join-Path -Path $Path -ChildPath ('.smartfinops-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $probePath -Value 'test' -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch { return $false }
}
function Resolve-SmartFinOpsExternalTokenValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [AllowNull()]$Config
    )

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($null -eq $Config) { return $Value }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $Config.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }
            $tokenValue = Resolve-SmartFinOpsExternalTokenValue -Value $tokenProperty.Value -Config $Config
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartFinOpsSmartM365LatestCsvFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepositoryRootPath,
        [Parameter(Mandatory)][string]$TenantKey
    )

    $smartM365RootPath = Join-Path -Path $RepositoryRootPath -ChildPath 'SmartM365'
    $smartM365TenantContextPath = Join-Path -Path $smartM365RootPath -ChildPath 'Config\SmartM365-TenantContext.ps1'
    if (-not (Test-Path -LiteralPath $smartM365TenantContextPath)) { return '' }

    try {
        . $smartM365TenantContextPath
        $smartM365Config = Get-SmartM365EffectiveGlobalConfig -StartPath $smartM365RootPath -TenantKey $TenantKey
        return [string](Resolve-SmartFinOpsExternalTokenValue -Value $smartM365Config.LatestCsvFolderPath -Config $smartM365Config)
    }
    catch {
        return ''
    }
}


function Get-SmartFinOpsEffectiveGlobalConfig {
    [CmdletBinding()]
    param(
        [string]$StartPath,
        [string]$TenantKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($TenantKey)) { $TenantKey = 'test' }
    $rootPath = Find-SmartFinOpsRoot -StartPath $StartPath
    $repositoryRootPath = Split-Path -Path $rootPath -Parent
    $scriptStartPath = if ([string]::IsNullOrWhiteSpace($StartPath)) { $rootPath } else { $StartPath }
    $scriptOutputRootPath = Join-Path -Path $scriptStartPath -ChildPath 'Output'
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'SmartFinOps.global.local.json'
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $TenantKey)

    $globalConfig = Read-SmartFinOpsJsonConfig -Path $globalConfigPath
    $tenantConfig = Read-SmartFinOpsJsonConfig -Path $tenantConfigPath

    if ($tenantConfig.Contains('TenantKey') -and -not [string]::IsNullOrWhiteSpace([string]$tenantConfig['TenantKey']) -and [string]$tenantConfig['TenantKey'] -ne $TenantKey) {
        throw "Tenant profile key mismatch. File '$tenantConfigPath' contains TenantKey '$($tenantConfig['TenantKey'])' but requested '$TenantKey'."
    }
    foreach ($key in $tenantConfig.Keys) { $globalConfig[$key] = $tenantConfig[$key] }

    $globalConfig['TenantKey'] = $TenantKey
    $globalConfig['SmartFinOpsRootPath'] = $rootPath
    $globalConfig['RepositoryRootPath'] = $repositoryRootPath
    $globalConfig['ScriptOutputRootPath'] = $scriptOutputRootPath

    $defaultWorkspaceRootPath = $rootPath
    $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
    $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
    $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
    $useScriptOutputFallback = $false

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or [string]$globalConfig['WorkspaceRootPath'] -in @('.', '{{SmartFinOpsRootPath}}')) {
        $candidateDataRoot = Join-Path -Path $rootPath -ChildPath 'Data'
        if (Test-SmartFinOpsWritableDirectory -Path $candidateDataRoot) {
            $defaultWorkspaceRootPath = $rootPath
        }
        else {
            $defaultWorkspaceRootPath = $scriptOutputRootPath
            $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\DATA-ALL'
            $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\DATA-LAST'
            $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\LOG-ALL'
            $useScriptOutputFallback = $true
        }
    }

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath'])) { $globalConfig['WorkspaceRootPath'] = $defaultWorkspaceRootPath }
    if (-not $globalConfig.Contains('DataAllRootPath') -or ($useScriptOutputFallback -and [string]$globalConfig['DataAllRootPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL')) { $globalConfig['DataAllRootPath'] = $defaultDataAllRootPath }
    if (-not $globalConfig.Contains('LatestCsvFolderPath') -or ($useScriptOutputFallback -and [string]$globalConfig['LatestCsvFolderPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST')) { $globalConfig['LatestCsvFolderPath'] = $defaultLatestCsvFolderPath }
    if (-not $globalConfig.Contains('LogAllRootPath') -or ($useScriptOutputFallback -and [string]$globalConfig['LogAllRootPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL')) { $globalConfig['LogAllRootPath'] = $defaultLogAllRootPath }
    $smartM365LatestCsvDefault = Get-SmartFinOpsSmartM365LatestCsvFolderPath -RepositoryRootPath $repositoryRootPath -TenantKey $TenantKey
    if ([string]::IsNullOrWhiteSpace($smartM365LatestCsvDefault)) { $smartM365LatestCsvDefault = '{{RepositoryRootPath}}\SmartM365\Data\Tenants\{{TenantKey}}\DATA-LAST' }
    if (-not $globalConfig.Contains('SmartM365LatestCsvFolderPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['SmartM365LatestCsvFolderPath']) -or
        [string]$globalConfig['SmartM365LatestCsvFolderPath'] -in @('__USE_SMARTM365__', 'USE_SMARTM365')) {
        $globalConfig['SmartM365LatestCsvFolderPath'] = $smartM365LatestCsvDefault
    }
    if (-not $globalConfig.Contains('RetentionMaxCSV')) { $globalConfig['RetentionMaxCSV'] = 30 }
    if (-not $globalConfig.Contains('RetentionMaxLogs')) { $globalConfig['RetentionMaxLogs'] = 30 }
    if (-not $globalConfig.Contains('StaleUserDays')) { $globalConfig['StaleUserDays'] = 90 }
    if (-not $globalConfig.Contains('StaleDeviceDays')) { $globalConfig['StaleDeviceDays'] = 60 }
    if (-not $globalConfig.Contains('ReportTitle')) { $globalConfig['ReportTitle'] = 'SmartFinOps Workplace Report' }

    return [pscustomobject]$globalConfig
}

function Resolve-SmartFinOpsConfigTokenValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($null -eq $script:SmartFinOpsGlobalConfig) { throw 'SmartFinOps tenant context has not been initialized.' }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartFinOpsGlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }
            $tokenValue = Resolve-SmartFinOpsConfigTokenValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartFinOpsConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $script:SmartFinOpsGlobalConfig) { throw 'SmartFinOps tenant context has not been initialized.' }
    $property = $script:SmartFinOpsGlobalConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    if ($property.Value -is [string]) {
        $value = $property.Value.Trim()
        if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $DefaultValue }
    }
    return Resolve-SmartFinOpsConfigTokenValue -Value $property.Value
}

function Get-SmartFinOpsScriptLocalConfig {
    [CmdletBinding()]
    param([string]$ScriptPath)

    $effectiveScriptPath = if ([string]::IsNullOrWhiteSpace($ScriptPath)) { $PSCommandPath } else { $ScriptPath }
    if ([string]::IsNullOrWhiteSpace($effectiveScriptPath)) { return [pscustomobject]@{} }
    $scriptFolder = Split-Path -Path $effectiveScriptPath -Parent
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($effectiveScriptPath)
    $configPath = Join-Path -Path $scriptFolder -ChildPath ("{0}.local.json" -f $scriptBaseName)
    if (-not (Test-Path -LiteralPath $configPath)) { return [pscustomobject]@{} }
    try { return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message) }
}

function Get-SmartFinOpsScriptConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue = $null
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartFinOpsConfigTokenValue -Value $property.Value }
        }
        else { return Resolve-SmartFinOpsConfigTokenValue -Value $property.Value }
    }
    return Get-SmartFinOpsConfigValue -Name $Name -DefaultValue $DefaultValue
}

function Resolve-SmartFinOpsOutputRoots {
    [CmdletBinding()]
    param(
        [string]$OutputRoot,
        [string]$LatestOutputRoot,
        [Parameter(Mandatory)][string]$AreaPath
    )

    if ($null -eq $script:SmartFinOpsGlobalConfig) { throw 'SmartFinOps tenant context has not been initialized.' }
    $resolvedOutputRoot = if ([string]::IsNullOrWhiteSpace($OutputRoot)) { Join-Path -Path (Resolve-SmartFinOpsConfigTokenValue -Value '{{DataAllRootPath}}') -ChildPath $AreaPath } else { Resolve-SmartFinOpsConfigTokenValue -Value $OutputRoot }
    $resolvedLatestOutputRoot = if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { Resolve-SmartFinOpsConfigTokenValue -Value '{{LatestCsvFolderPath}}' } else { Resolve-SmartFinOpsConfigTokenValue -Value $LatestOutputRoot }
    [pscustomobject]@{ OutputRoot = $resolvedOutputRoot; LatestOutputRoot = $resolvedLatestOutputRoot }
}

function Initialize-SmartFinOpsTenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }
    $global:SmartFinOpsTenant = $Tenant
    $script:SmartFinOpsGlobalConfig = Get-SmartFinOpsEffectiveGlobalConfig -StartPath $StartPath -TenantKey $Tenant
    return $script:SmartFinOpsGlobalConfig
}
