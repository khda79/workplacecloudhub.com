<#
.SYNOPSIS
  SmartFinOps tenant context: global and tenant local JSON configuration, path tokens and console banners.
.DESCRIPTION
  Dot-sourced by the SmartFinOps V2 scripts. -Tenant is the profile key (Config\Tenants\<profile>.local.json with its
  ProfileKey; a legacy TenantKey entry is renamed to ProfileKey before template synchronization). SmartM365LatestCsvFolderPath,
  SmartM365DataAllRootPath, SmartM365LogAllRootPath, SmartM365TenantKey (the organization TenantKey carried by the source
  receipts) and EnableSharePointUpload set to __USE_SMARTM365__ (or left empty) are resolved through the SmartM365 tenant
  context of that profile only when a script asks for them; a failed resolution stops the script with the reason and
  never falls back to a default. A resolved path that still holds a {{token}} is refused. Get-SmartFinOpsDataLayout
  returns the SmartM365 DATA-LAST, DATA-ALL and LOG-ALL folders of the profile, which must be existing siblings. Before any
  SharePoint upload, Initialize-SmartFinOpsSmartM365Context initializes the SmartM365 context of the same profile (Graph
  identity and SharePoint target of that profile); a SmartM365 context of another profile is refused, and SmartFinOps never
  creates a SmartM365 runtime JSON.
.VERSION
1.4.1
#>
$script:SmartFinOpsTenantContextVersion = '1.4.1'
$script:SmartFinOpsSmartM365Config = $null
$script:SmartFinOpsUseSmartM365Markers = @('__USE_SMARTM365__', 'USE_SMARTM365')
# SmartFinOps path setting -> SmartM365 key and the leaf name the resolved folder must have.
$script:SmartFinOpsSmartM365Paths = [ordered]@{
    SmartM365LatestCsvFolderPath = @('LatestCsvFolderPath', 'DATA-LAST')
    SmartM365DataAllRootPath = @('DataAllRootPath', 'DATA-ALL')
    SmartM365LogAllRootPath = @('LogAllRootPath', 'LOG-ALL')
}
# SmartM365 SharePoint upload settings used by the SmartFinOps uploads (values come from the SmartM365 context only).
$script:SmartFinOpsSharePointKeys = @('SharePointSiteHostname', 'SharePointSitePath', 'SharePointLibraryDisplayName', 'SharePointTargetFolderPath', 'AppId', 'TenantId', 'Thumb')

function Write-SmartFinOpsStartupBanner {
    [CmdletBinding()]
    param()

    $bannerAlreadyShown = Get-Variable -Name SmartFinOpsBrandBannerConsoleShown -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ([bool]$bannerAlreadyShown) { return }

    Set-Variable -Name SmartFinOpsBrandBannerConsoleShown -Scope Global -Value $true
    Microsoft.PowerShell.Utility\Write-Host '================================================================================' -ForegroundColor DarkCyan
    Microsoft.PowerShell.Utility\Write-Host ' SmartFinOps by WorkplaceCloudHub' -ForegroundColor Cyan
    Microsoft.PowerShell.Utility\Write-Host ' Website : https://workplacecloudhub.com' -ForegroundColor Yellow
    Microsoft.PowerShell.Utility\Write-Host ' GitHub  : https://github.com/khda79/workplacecloudhub.com' -ForegroundColor Yellow
    Microsoft.PowerShell.Utility\Write-Host '================================================================================' -ForegroundColor DarkCyan
}

function Write-SmartFinOpsCompletionBanner {
    [CmdletBinding()]
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = 'Auto',
        [string]$ScriptName = '',
        [AllowNull()][Nullable[datetime]]$StartedAt = $null,
        [AllowNull()][Nullable[datetime]]$EndedAt = $null,
        [int]$WarningCount = 0,
        [int]$ErrorCount = 0,
        [int]$GeneratedCsvFiles = 0,
        [string]$LogPath = ''
    )
    # Same format as the SmartM365 completion banner; local copy so SmartFinOps does not depend on SmartM365 files.
    if ($null -eq $EndedAt -or $EndedAt -eq [datetime]::MinValue) { $EndedAt = Get-Date }
    if ($null -eq $StartedAt -or $StartedAt -eq [datetime]::MinValue) { $StartedAt = $EndedAt }
    if ([string]::IsNullOrWhiteSpace($ScriptName)) { $ScriptName = 'SmartFinOps' }

    if ($Status -eq 'Auto') {
        $Status = if ($ErrorCount -gt 0) { 'Failed' } elseif ($WarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    }

    $statusLabel = switch ($Status) {
        'Failed' { 'FAILED' }
        'CompletedWithWarnings' { 'COMPLETED WITH WARNINGS' }
        default { 'SUCCESS' }
    }
    $title = switch ($Status) {
        'Failed' { 'Execution failed' }
        'CompletedWithWarnings' { 'Execution completed with warnings' }
        default { 'Execution completed' }
    }
    $statusColor = switch ($Status) {
        'Failed' { 'Red' }
        'CompletedWithWarnings' { 'Yellow' }
        default { 'Green' }
    }
    $duration = $EndedAt - $StartedAt
    $durationText = '{0:00}:{1:00}:{2:00}' -f ([int][Math]::Floor($duration.TotalHours)), $duration.Minutes, $duration.Seconds
    $effectiveLogPath = if ([string]::IsNullOrWhiteSpace($LogPath)) { '' } else { [string]$LogPath }
    $runKey = '{0}|{1}|{2}' -f $ScriptName, $StartedAt.Ticks, $effectiveLogPath
    $existingRunKey = [string](Get-Variable -Name SmartFinOpsCompletionBannerRunKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if ($existingRunKey -eq $runKey) { return }

    $lines = @(
        '================================================================================',
        (' SmartFinOps by WorkplaceCloudHub - {0}' -f $title),
        (' Script    : {0}' -f $ScriptName),
        (' Status    : {0}' -f $statusLabel),
        (' Duration  : {0}' -f $durationText),
        (' Warnings  : {0}' -f $WarningCount),
        (' Errors    : {0}' -f $ErrorCount),
        (' CSV files : {0}' -f $GeneratedCsvFiles)
    )
    if (-not [string]::IsNullOrWhiteSpace($effectiveLogPath)) {
        $lines += (' Log       : {0}' -f $effectiveLogPath)
    }
    $lines += '================================================================================'

    Set-Variable -Name SmartFinOpsCompletionBannerRunKey -Scope Global -Value $runKey
    Microsoft.PowerShell.Utility\Write-Host $lines[0] -ForegroundColor DarkCyan
    Microsoft.PowerShell.Utility\Write-Host $lines[1] -ForegroundColor Cyan
    for ($i = 2; $i -lt ($lines.Count - 1); $i++) {
        $color = if ($lines[$i] -like ' Status*') { $statusColor } else { 'Gray' }
        Microsoft.PowerShell.Utility\Write-Host $lines[$i] -ForegroundColor $color
    }
    Microsoft.PowerShell.Utility\Write-Host $lines[-1] -ForegroundColor DarkCyan

    if (-not [string]::IsNullOrWhiteSpace($effectiveLogPath)) {
        try {
            $logFolder = Split-Path -Path $effectiveLogPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($logFolder) -and -not (Test-Path -LiteralPath $logFolder)) {
                New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
            }
            Add-Content -LiteralPath $effectiveLogPath -Value $lines -Encoding UTF8
        }
        catch {
            Microsoft.PowerShell.Utility\Write-Warning ("SmartFinOps completion banner could not be written to '{0}': {1}" -f $effectiveLogPath, $_.Exception.Message)
        }
    }
}

function ConvertTo-SmartFinOpsFileSystemPath {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Path)

    # .NET reads a provider-qualified path (Microsoft.PowerShell.Core\FileSystem::\\server\share\...) or a relative
    # path against the process directory: always hand it the plain absolute file-system path.
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Assert-SmartFinOpsResolvedPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][AllowEmptyString()][string]$Value
    )

    $token = [regex]::Match([string]$Value, '\{\{[^{}]*\}\}')
    if ($token.Success) { throw ("{0} still contains an unresolved token {1} after resolution: '{2}'. Fix the SmartFinOps or SmartM365 configuration, or pass the path explicitly." -f $Name, $token.Value, $Value) }
}

function ConvertTo-SmartFinOpsHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = [ordered]@{}
    if ($null -eq $InputObject) { return $hash }
    foreach ($property in $InputObject.PSObject.Properties) { $hash[$property.Name] = $property.Value }
    return $hash
}

function Get-SmartFinOpsJsonTemplatePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if ($Path -like '*Config\Tenants\*.local.json') {
        return (Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template')
    }

    if ($Path -like '*.local.json') {
        return ('{0}.template' -f $Path)
    }

    return ''
}

function Add-SmartFinOpsMissingJsonTemplateProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)]$Template,
        [string]$PrefixPath = ''
    )

    $added = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Template) { return @() }

    foreach ($property in $Template.PSObject.Properties) {
        $name = $property.Name
        $propertyPath = if ([string]::IsNullOrWhiteSpace($PrefixPath)) { $name } else { '{0}.{1}' -f $PrefixPath, $name }
        $exists = $false
        $currentValue = $null

        if ($Target -is [System.Collections.IDictionary]) {
            $exists = $Target.Contains($name)
            if ($exists) { $currentValue = $Target[$name] }
        }
        else {
            $currentProperty = $Target.PSObject.Properties[$name]
            $exists = ($null -ne $currentProperty)
            if ($exists) { $currentValue = $currentProperty.Value }
        }

        if (-not $exists) {
            if ($Target -is [System.Collections.IDictionary]) {
                $Target[$name] = $property.Value
            }
            else {
                Add-Member -InputObject $Target -MemberType NoteProperty -Name $name -Value $property.Value -Force
            }
            $added.Add($propertyPath) | Out-Null
            continue
        }

        if ($null -ne $currentValue -and $null -ne $property.Value -and $currentValue -is [pscustomobject] -and $property.Value -is [pscustomobject]) {
            foreach ($nestedPath in (Add-SmartFinOpsMissingJsonTemplateProperties -Target $currentValue -Template $property.Value -PrefixPath $propertyPath)) {
                $added.Add($nestedPath) | Out-Null
            }
        }
    }

    return @($added)
}

function Sync-SmartFinOpsJsonConfigWithTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartFinOpsJsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $addedKeys = @(Add-SmartFinOpsMissingJsonTemplateProperties -Target $Config -Template $template)
        if ($addedKeys.Count -gt 0) {
            $tempPath = '{0}.tmp' -f $Path
            $Config | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $tempPath -Encoding UTF8 -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
            Write-Host ("Updated local JSON from template: {0}; added keys: {1}" -f $Path, ($addedKeys -join ', ')) -ForegroundColor Yellow
        }
        return $Config
    }
    catch {
        $syncErrorMessage = [string]$PSItem.Exception.Message
        throw ("Failed to synchronize local JSON '{0}' with template '{1}': {2}" -f $Path, $TemplatePath, $syncErrorMessage)
    }
}
function Convert-SmartFinOpsLegacyProfileKey {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    # Profile files written before 1.3.0 name the profile key TenantKey; the template now holds ProfileKey.
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    try { $document = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop }
    catch { throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message) }
    if ($null -eq $document -or -not $document.Contains('TenantKey')) { return }
    $legacyValue = [string]$document['TenantKey']
    if ($document.Contains('ProfileKey') -and -not [string]::IsNullOrWhiteSpace([string]$document['ProfileKey']) -and
        -not [string]::IsNullOrWhiteSpace($legacyValue) -and [string]$document['ProfileKey'] -cne $legacyValue) {
        throw ("Profile file '{0}' contains ProfileKey '{1}' and legacy TenantKey '{2}': keep a single ProfileKey." -f $Path, $document['ProfileKey'], $legacyValue)
    }
    $migrated = [ordered]@{}
    foreach ($key in $document.Keys) {
        if ($key -ceq 'TenantKey') {
            if (-not $migrated.Contains('ProfileKey')) { $migrated['ProfileKey'] = $(if ($document.Contains('ProfileKey') -and -not [string]::IsNullOrWhiteSpace([string]$document['ProfileKey'])) { $document['ProfileKey'] } else { $legacyValue }) }
            continue
        }
        if ($key -ceq 'ProfileKey' -and $migrated.Contains('ProfileKey')) { continue }
        $migrated[$key] = $document[$key]
    }
    $tempPath = '{0}.tmp' -f $Path
    $migrated | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $tempPath -Encoding UTF8 -ErrorAction Stop
    Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
    Write-Host ("Updated local JSON: {0}; renamed legacy key TenantKey to ProfileKey ('{1}')" -f $Path, $migrated['ProfileKey']) -ForegroundColor Yellow
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

    try { $config = ConvertTo-SmartFinOpsHashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        return (Sync-SmartFinOpsJsonConfigWithTemplate -Config $config -Path $Path) }
    catch { throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message) }
}

function Find-SmartFinOpsRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    # ProviderPath, never Path: on a UNC share Path is provider-qualified and breaks every derived path.
    $searchRoot = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).ProviderPath } else { $StartPath }
    $resolvedSearchRoot = Resolve-Path -LiteralPath $searchRoot -ErrorAction SilentlyContinue
    if ($null -ne $resolvedSearchRoot) { $searchRoot = $resolvedSearchRoot.ProviderPath }

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

function Assert-SmartFinOpsSmartM365Profile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ProfileKey)

    # A SmartM365 context of another profile in this session would give its identity and SharePoint target to SmartFinOps.
    $loaded = [string](Get-Variable -Name SmartM365ProfileKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    $loadedConfig = Get-Variable -Name SmartM365GlobalConfig -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    $configProfile = ''
    if ($null -ne $loadedConfig -and $null -ne $loadedConfig.PSObject.Properties['ProfileKey']) { $configProfile = [string]$loadedConfig.PSObject.Properties['ProfileKey'].Value }
    foreach ($other in @($loaded, $configProfile)) {
        if (-not [string]::IsNullOrWhiteSpace($other) -and $other.Trim() -ine $ProfileKey.Trim()) {
            throw ("A SmartM365 context of profile '{0}' is already loaded in this session; SmartFinOps profile '{1}' refuses to use it." -f $other.Trim(), $ProfileKey)
        }
    }
}

function Get-SmartFinOpsSmartM365Config {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepositoryRootPath,
        [Parameter(Mandatory)][string]$ProfileKey,
        [Parameter(Mandatory)][string]$Setting
    )

    # The SmartM365 effective configuration of the profile, read once per run; every failure is explicit.
    if ($null -ne $script:SmartFinOpsSmartM365Config -and $script:SmartFinOpsSmartM365Config.ProfileKeyRequested -ceq $ProfileKey) { return $script:SmartFinOpsSmartM365Config.Config }
    $smartM365RootPath = Join-Path -Path $RepositoryRootPath -ChildPath 'SmartM365'
    $smartM365TenantContextPath = Join-Path -Path $smartM365RootPath -ChildPath 'Config\SmartM365-TenantContext.ps1'
    $prefix = "$Setting is __USE_SMARTM365__ but its resolution through SmartM365 failed for profile '$ProfileKey'"
    if (-not (Test-Path -LiteralPath $smartM365TenantContextPath -PathType Leaf)) {
        throw ("{0}: SmartM365 tenant context not found: {1}. Set {2} in the SmartFinOps configuration or pass -{2}." -f $prefix, $smartM365TenantContextPath, $Setting)
    }
    Assert-SmartFinOpsSmartM365Profile -ProfileKey $ProfileKey
    # The SmartM365 context creates a missing runtime JSON from its template: SmartFinOps never lets it do so.
    foreach ($runtimePath in @((Join-Path $smartM365RootPath 'Config\SmartM365.global.local.json'), (Join-Path $smartM365RootPath ("Config\Tenants\{0}.local.json" -f $ProfileKey)))) {
        if (-not ((Test-Path -LiteralPath $runtimePath -PathType Leaf) -or (Test-Path -LiteralPath ($runtimePath + '.txt') -PathType Leaf))) {
            throw ("{0}: SmartM365 runtime configuration not found: {1} (or .txt). SmartFinOps never creates SmartM365 runtime JSON; configure the SmartM365 profile first." -f $prefix, $runtimePath)
        }
    }
    try {
        . $smartM365TenantContextPath
        $smartM365Config = Get-SmartM365EffectiveGlobalConfig -StartPath $smartM365RootPath -TenantKey $ProfileKey
    }
    catch {
        $invocation = $_.InvocationInfo
        $position = if ($invocation -and $invocation.ScriptName) { '{0}:{1}' -f $invocation.ScriptName, $invocation.ScriptLineNumber } elseif ($invocation) { 'line {0}' -f $invocation.ScriptLineNumber } else { 'unknown position' }
        throw ("{0}: {1} (at {2}; context {3})" -f $prefix, $_.Exception.Message, $position, $smartM365TenantContextPath)
    }
    $script:SmartFinOpsSmartM365Config = [pscustomobject]@{ ProfileKeyRequested=$ProfileKey; Config=$smartM365Config; ContextPath=$smartM365TenantContextPath }
    return $smartM365Config
}

function Get-SmartFinOpsSmartM365PathValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepositoryRootPath,
        [Parameter(Mandatory)][string]$ProfileKey,
        [Parameter(Mandatory)][string]$Setting
    )

    # Called only for a SmartM365 path setting = __USE_SMARTM365__: never a default path.
    $smartM365Key = $script:SmartFinOpsSmartM365Paths[$Setting][0]
    $smartM365Config = Get-SmartFinOpsSmartM365Config -RepositoryRootPath $RepositoryRootPath -ProfileKey $ProfileKey -Setting $Setting
    $prefix = "$Setting is __USE_SMARTM365__ but its resolution through SmartM365 failed for profile '$ProfileKey'"
    $property = $smartM365Config.PSObject.Properties[$smartM365Key]
    $resolved = if ($null -ne $property) { [string](Resolve-SmartFinOpsExternalTokenValue -Value $property.Value -Config $smartM365Config) } else { '' }
    if ([string]::IsNullOrWhiteSpace($resolved)) { throw ("{0}: the SmartM365 context returned an empty {1} ({2})." -f $prefix, $smartM365Key, $script:SmartFinOpsSmartM365Config.ContextPath) }
    Assert-SmartFinOpsResolvedPath -Name ("{0} (resolved through SmartM365)" -f $Setting) -Value $resolved
    $resolved = ConvertTo-SmartFinOpsFileSystemPath $resolved
    # SmartM365 falls back to an Output folder next to the script when its Data root is not writable: never write there.
    $outputProperty = $smartM365Config.PSObject.Properties['ScriptOutputRootPath']
    if ($null -ne $outputProperty -and -not [string]::IsNullOrWhiteSpace([string]$outputProperty.Value)) {
        $outputRoot = ([IO.Path]::GetFullPath((ConvertTo-SmartFinOpsFileSystemPath ([string]$outputProperty.Value)))).TrimEnd('\') + '\'
        if (([IO.Path]::GetFullPath($resolved).TrimEnd('\') + '\').StartsWith($outputRoot, [StringComparison]::OrdinalIgnoreCase)) {
            throw ("{0}: {1} resolves to the SmartM365 Output fallback ({2}); configure the SmartM365 data roots of the profile." -f $prefix, $smartM365Key, $resolved)
        }
    }
    return $resolved
}

function Get-SmartFinOpsSmartM365TenantKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepositoryRootPath,
        [Parameter(Mandatory)][string]$ProfileKey
    )

    # Called only for SmartM365TenantKey = __USE_SMARTM365__: the TenantKey (OrganizationKey-EnvironmentKey) of the SmartM365 profile.
    $smartM365Config = Get-SmartFinOpsSmartM365Config -RepositoryRootPath $RepositoryRootPath -ProfileKey $ProfileKey -Setting 'SmartM365TenantKey'
    $tenantKeyProperty = $smartM365Config.PSObject.Properties['TenantKey']
    $tenantKey = if ($null -ne $tenantKeyProperty) { ([string]$tenantKeyProperty.Value).Trim() } else { '' }
    if (-not $tenantKey) { throw ("SmartM365TenantKey is __USE_SMARTM365__ but the SmartM365 context of profile '{0}' returned no TenantKey ({1}). Set SmartM365TenantKey in the SmartFinOps configuration or pass -SmartM365TenantKey." -f $ProfileKey, $script:SmartFinOpsSmartM365Config.ContextPath) }
    Assert-SmartFinOpsResolvedPath -Name 'SmartM365TenantKey (resolved through SmartM365)' -Value $tenantKey
    return $tenantKey
}


function Get-SmartFinOpsEffectiveGlobalConfig {
    [CmdletBinding()]
    param(
        [string]$StartPath,
        [string]$ProfileKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($ProfileKey)) { $ProfileKey = 'test' }
    $rootPath = Find-SmartFinOpsRoot -StartPath $StartPath
    $repositoryRootPath = Split-Path -Path $rootPath -Parent
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'SmartFinOps.global.local.json'
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $ProfileKey)

    $globalConfig = Read-SmartFinOpsJsonConfig -Path $globalConfigPath
    Convert-SmartFinOpsLegacyProfileKey -Path $tenantConfigPath
    $tenantConfig = Read-SmartFinOpsJsonConfig -Path $tenantConfigPath

    if ($tenantConfig.Contains('ProfileKey') -and -not [string]::IsNullOrWhiteSpace([string]$tenantConfig['ProfileKey']) -and [string]$tenantConfig['ProfileKey'] -ne $ProfileKey) {
        throw "Profile key mismatch. File '$tenantConfigPath' contains ProfileKey '$($tenantConfig['ProfileKey'])' but requested '$ProfileKey'."
    }
    foreach ($key in $tenantConfig.Keys) {
        $tenantValue = $tenantConfig[$key]
        if ($tenantValue -is [string] -and $tenantValue.Trim() -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { continue }
        $globalConfig[$key] = $tenantValue
    }

    # {{ProfileKey}} names the profile; {{TenantKey}} stays a read alias for configurations written before 1.3.0.
    $globalConfig['ProfileKey'] = $ProfileKey
    $globalConfig['TenantKey'] = $ProfileKey
    $globalConfig['SmartFinOpsRootPath'] = $rootPath
    $globalConfig['RepositoryRootPath'] = $repositoryRootPath

    # Empty or __USE_SMARTM365__: resolved through SmartM365 only when a script asks for it (Get-SmartFinOpsConfigValue).
    foreach ($key in @(@($script:SmartFinOpsSmartM365Paths.Keys) + @('SmartM365TenantKey', 'EnableSharePointUpload'))) {
        if (-not $globalConfig.Contains($key) -or ($globalConfig[$key] -is [string] -and
            ([string]::IsNullOrWhiteSpace([string]$globalConfig[$key]) -or [string]$globalConfig[$key] -in $script:SmartFinOpsUseSmartM365Markers))) {
            $globalConfig[$key] = '__USE_SMARTM365__'
        }
    }

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
        if ($script:SmartFinOpsSmartM365Paths.Contains($Name)) {
            if ($value -in $script:SmartFinOpsUseSmartM365Markers) { return Get-SmartFinOpsSmartM365PathValue -RepositoryRootPath ([string]$script:SmartFinOpsGlobalConfig.RepositoryRootPath) -ProfileKey ([string]$script:SmartFinOpsGlobalConfig.ProfileKey) -Setting $Name }
            $resolvedPath = [string](Resolve-SmartFinOpsConfigTokenValue -Value $value)
            Assert-SmartFinOpsResolvedPath -Name $Name -Value $resolvedPath
            return (ConvertTo-SmartFinOpsFileSystemPath $resolvedPath)
        }
        if ($Name -eq 'SmartM365TenantKey') {
            if ($value -in $script:SmartFinOpsUseSmartM365Markers) { return Get-SmartFinOpsSmartM365TenantKey -RepositoryRootPath ([string]$script:SmartFinOpsGlobalConfig.RepositoryRootPath) -ProfileKey ([string]$script:SmartFinOpsGlobalConfig.ProfileKey) }
            Assert-SmartFinOpsResolvedPath -Name 'SmartM365TenantKey' -Value $value
            return $value
        }
        if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $DefaultValue }
    }
    return Resolve-SmartFinOpsConfigTokenValue -Value $property.Value
}

function Get-SmartFinOpsDataLayout {
    [CmdletBinding()]
    param(
        [string]$DataLast,
        [string]$DataAll,
        [string]$LogAll
    )

    # SmartM365 data tree of the profile: DATA-LAST (current files), DATA-ALL (history), LOG-ALL (logs).
    if ($null -eq $script:SmartFinOpsGlobalConfig) { throw 'SmartFinOps tenant context has not been initialized.' }
    $overrides = @{ SmartM365LatestCsvFolderPath = $DataLast; SmartM365DataAllRootPath = $DataAll; SmartM365LogAllRootPath = $LogAll }
    $folders = [ordered]@{}
    foreach ($setting in $script:SmartFinOpsSmartM365Paths.Keys) {
        $value = if ($overrides[$setting]) { [string]$overrides[$setting] } else { [string](Get-SmartFinOpsConfigValue -Name $setting) }
        Assert-SmartFinOpsResolvedPath -Name $setting -Value $value
        if ([string]::IsNullOrWhiteSpace($value)) { throw "$setting is empty." }
        $path = [IO.Path]::GetFullPath((ConvertTo-SmartFinOpsFileSystemPath $value)).TrimEnd('\')
        $leaf = $script:SmartFinOpsSmartM365Paths[$setting][1]
        if ((Split-Path -Path $path -Leaf) -ne $leaf) { throw ("{0} must name a folder called {1}: {2}" -f $setting, $leaf, $path) }
        if (-not (Test-Path -LiteralPath $path -PathType Container)) { throw ("{0} folder not found: {1}. SmartFinOps never creates the SmartM365 data roots." -f $setting, $path) }
        $folders[$leaf] = $path
    }
    $parents = @($folders.Values | ForEach-Object { (Split-Path -Path $_ -Parent).TrimEnd('\') } | Sort-Object -Unique)
    if ($parents.Count -ne 1) { throw ("DATA-LAST, DATA-ALL and LOG-ALL must be sibling folders of the same tenant data root: {0}" -f (@($folders.Values) -join ' | ')) }
    [pscustomobject]@{
        TenantDataRoot = $parents[0]
        DataLast = $folders['DATA-LAST']
        DataAll = $folders['DATA-ALL']
        LogAll = $folders['LOG-ALL']
        HistoryRoot = Join-Path -Path $folders['DATA-ALL'] -ChildPath 'FinOps\Workplace'
    }
}

function Initialize-SmartFinOpsSmartM365Context {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProfileKey,
        [string]$TenantKey
    )

    # Same initialization as a SmartM365 script of this profile, after the runtime-file and profile checks.
    $repositoryRootPath = [string]$script:SmartFinOpsGlobalConfig.RepositoryRootPath
    $null = Get-SmartFinOpsSmartM365Config -RepositoryRootPath $repositoryRootPath -ProfileKey $ProfileKey -Setting 'EnableSharePointUpload'
    $smartM365RootPath = Join-Path -Path $repositoryRootPath -ChildPath 'SmartM365'
    $contextPath = Join-Path -Path $smartM365RootPath -ChildPath 'Config\SmartM365-TenantContext.ps1'
    try {
        . $contextPath
        # SmartFinOps already printed its own banner.
        Set-Variable -Name SmartM365BrandBannerConsoleShown -Scope Global -Value $true
        $effective = Initialize-SmartM365TenantContext -Tenant $ProfileKey -StartPath $smartM365RootPath
    }
    catch { throw ("The SmartM365 context of profile '{0}' could not be initialized: {1}" -f $ProfileKey, $_.Exception.Message) }
    $loadedProfile = [string](Get-Variable -Name SmartM365ProfileKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if ($loadedProfile -ine $ProfileKey) { throw ("The SmartM365 context loaded profile '{0}' instead of '{1}'." -f $loadedProfile, $ProfileKey) }
    $loadedTenantKey = [string](Get-Variable -Name SmartM365TenantKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if (-not [string]::IsNullOrWhiteSpace($TenantKey) -and $loadedTenantKey -cne $TenantKey) {
        throw ("The SmartM365 context of profile '{0}' has TenantKey '{1}', SmartFinOps expects '{2}'." -f $ProfileKey, $loadedTenantKey, $TenantKey)
    }
    return $effective
}

function Test-SmartFinOpsConfiguredValue {
    param([AllowNull()][string]$Value)
    # Same rule as SmartM365: empty, __USE_GLOBAL__ and all-zero placeholders are not configured; tokens are not resolved values.
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $text = $Value.Trim()
    return -not ($text -in @('__USE_GLOBAL__', 'USE_GLOBAL') -or $text -match '\{\{' -or $text -match '^0{8}-0{4}-0{4}-0{4}-0{12}$' -or $text -match '^0{40}$')
}

function Get-SmartFinOpsSharePointConfiguration {
    [CmdletBinding()]
    param([string]$TenantKey)

    # EnableSharePointUpload: true/false, or __USE_SMARTM365__ (default) for the SmartM365 setting of the profile.
    # Site, library, target folder and app-only identity always come from the SmartM365 context of the profile.
    if ($null -eq $script:SmartFinOpsGlobalConfig) { throw 'SmartFinOps tenant context has not been initialized.' }
    $configured = $script:SmartFinOpsGlobalConfig.PSObject.Properties['EnableSharePointUpload']
    $value = if ($null -ne $configured) { $configured.Value } else { '__USE_SMARTM365__' }
    $smartM365Config = $null
    if ($value -is [string] -and $value.Trim() -in $script:SmartFinOpsUseSmartM365Markers) {
        $smartM365Config = Get-SmartFinOpsSmartM365Config -RepositoryRootPath ([string]$script:SmartFinOpsGlobalConfig.RepositoryRootPath) -ProfileKey ([string]$script:SmartFinOpsGlobalConfig.ProfileKey) -Setting 'EnableSharePointUpload'
        $property = $smartM365Config.PSObject.Properties['EnableSharePointUpload']
        $value = if ($null -ne $property) { $property.Value } else { $false }
    }
    $enabled = $false
    if ($value -is [bool]) { $enabled = $value }
    elseif (-not [bool]::TryParse([string]$value, [ref]$enabled)) { throw ("EnableSharePointUpload must be true, false or __USE_SMARTM365__ (got '{0}')." -f $value) }
    $settings = [ordered]@{}
    $missing = [System.Collections.Generic.List[string]]::new()
    if ($enabled) {
        # Upload identity and target are those of the SmartM365 context of this profile, initialized as for a SmartM365 script.
        $profileKey = [string]$script:SmartFinOpsGlobalConfig.ProfileKey
        $effective = Initialize-SmartFinOpsSmartM365Context -ProfileKey $profileKey -TenantKey $TenantKey
        foreach ($key in $script:SmartFinOpsSharePointKeys) {
            $candidates = if ($key -eq 'Thumb') { @('Thumb', 'Thumbprint') } else { @($key) }
            $resolved = ''
            foreach ($candidate in $candidates) {
                $property = $effective.PSObject.Properties[$candidate]
                $value = if ($null -ne $property) { [string](Resolve-SmartFinOpsExternalTokenValue -Value $property.Value -Config $effective) } else { '' }
                if (Test-SmartFinOpsConfiguredValue $value) { $resolved = $value.Trim(); break }
            }
            if (-not $resolved) { $missing.Add($key); continue }
            $settings[$key] = $resolved
        }
        if ($missing.Count -eq 0) {
            # The SmartM365.Core SharePoint helpers read this identity from the session, like any SmartM365 script.
            $sessionValues = [ordered]@{
                AppId = $settings.AppId; TenantId = $settings.TenantId; Thumb = $settings.Thumb; Thumbprint = $settings.Thumb; EnableSharePointUpload = $true
                SharePointSiteHostname = $settings.SharePointSiteHostname; SharePointSitePath = $settings.SharePointSitePath
                SharePointLibraryDisplayName = $settings.SharePointLibraryDisplayName; SharePointTargetFolderPath = $settings.SharePointTargetFolderPath
            }
            foreach ($name in $sessionValues.Keys) { Set-Variable -Name $name -Scope Global -Value $sessionValues[$name] }
        }
    }
    [pscustomobject]@{
        Enabled = $enabled
        ProfileKey = [string]$script:SmartFinOpsGlobalConfig.ProfileKey
        Missing = @($missing)
        Parameters = $(if ($enabled -and $missing.Count -eq 0) {
            @{ Enabled = $true; SiteHostname = $settings.SharePointSiteHostname; SitePath = $settings.SharePointSitePath; LibraryDisplayName = $settings.SharePointLibraryDisplayName
               TargetFolderPath = $settings.SharePointTargetFolderPath; AppId = $settings.AppId; TenantId = $settings.TenantId; Thumbprint = $settings.Thumb }
        } else { $null })
    }
}

function Initialize-SmartFinOpsTenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }
    $global:SmartFinOpsTenant = $Tenant
    $script:SmartFinOpsGlobalConfig = Get-SmartFinOpsEffectiveGlobalConfig -StartPath $StartPath -ProfileKey $Tenant
    return $script:SmartFinOpsGlobalConfig
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBPA+nj+UwVAzR0
# qS5GdzM+aMyiCfMEHgx3bFQPvhajJaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIAbppkPBPiYVFqoiR/SFLZPuqIIqRiMTmevP72QEQukrMA0GCSqG
# SIb3DQEBAQUABIIBgI9LviDHiZB4TXXT6dXeKfploWCDTfvs15iJMQo++S9Mw207
# O5iK1lVWvsb7Hv7cwEcmdo9ouKsvN4222W+jcgmiTG8/pMbP5EPEqn6pUOlIB5We
# /iNQLzOIO9kYnbClYipHSsXAgo3uO5mNf0AbzgYS4v5VlN28lByN3Seg+H6SpXHO
# LhYLjnLzh0gACpmRx20+EgP2YzRXTirP+kIQj86D0zavBE67KP+glufjPOO/cdTH
# Z+qPJ+I0bdKwoEglgCH6dnld9eL9r+YUdRbleD5zs5ofenzH7dvMzbdDxA0eGMvA
# OYtDH/AsK/9WSxoBAkX+9staBDldGCVUIEYTOj9TOfY+sodoj+n3pGv3TH05l+sf
# BD6qxVaQALB8R5EzjTMcbV9zN7ib7s2Anqmc9mCEY0Z25M7+bIVskPo9Qeo/Qfhz
# Fm0Yt2he4WlYEdruYd70fXRMbW/tIWLiiPVFIr4Xw6NNxGROnefgHmJNtAqMsqlS
# FWU6dPtPq2zzwQPYhqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkwMDUw
# NTJaMC8GCSqGSIb3DQEJBDEiBCAS6EibBJ41HrC3XnHDuM6gdH5uHtf3bGdWnRRg
# cnKFTTANBgkqhkiG9w0BAQEFAASCAgAHup1RPfZCBprEzXeBxj4vC6oj0qgeyx78
# o5l+K3eaEINbn42dzC51p4v52sSEPBtL8vq6J5wKZw6MhWuacskt0GvR8VlrqQCQ
# SypqoCzGUDBy43N0z6QkGe+EXwo/nhFJXxjqjylsS3pXzMV7cN0k8llLkRH88wbT
# 53WgJIvaiLc9vAWxlekgLPVyCNRN52cbWBaX0fmVdSHx92LwABLRyuipLvAwIGJV
# C2OhihU4ol7XqKC3XgheOXDKgK5Lkdlw3GOje0R1etRXgl0qqEETObWQTObMXZfy
# 8EhcOkGOpEW9geI8UwHCnekkl5iDt/pRyg4o0ysP41G/wq2EOW8zXcCo9w4qwzXY
# GOpgPQxA1NYG1NyBxPbGsJsHaSCliGyG7z35160OHreu0/bfHNPfAc/Nt6rXOqYh
# hSpJSn1HOpUgmA3hd3KjQ0dC3e4murmlsRW+0HyKcpGQcDwsxkEF9PY4Zqhv2nuO
# Zp8aBTqukoW2BeVVNyiZ+6/8Qnjxx6a0EMX+7PRGXYczZdU7W76+u5hto+lGPn1H
# JixAlcxPZUdQm7qmqSbfhjgl5KAa2hxUiY8RSPPzuv8wal583NYw+SjX8ltCTC8I
# zatPOy7+w2uPmjA7Lu3z/sa3MsyAjK+/WuEuFGkwqTijh9eeFsi0N1bcsGpbALBL
# pjDlrkiomg==
# SIG # End signature block
