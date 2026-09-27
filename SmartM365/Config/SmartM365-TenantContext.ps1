Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
<#
.SYNOPSIS
Loads SmartM365 global and tenant configuration context.

.DESCRIPTION
Merges global and tenant local JSON configuration, resolves workspace paths, and exposes the effective tenant context used by SmartM365 scripts.

.VERSION
1.0.8
#>
function Write-SmartM365StartupBanner {
    [CmdletBinding()]
    param()

    $bannerAlreadyShown = Get-Variable -Name SmartM365BrandBannerConsoleShown -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ([bool]$bannerAlreadyShown) { return }

    Set-Variable -Name SmartM365BrandBannerConsoleShown -Scope Global -Value $true
    Microsoft.PowerShell.Utility\Write-Host '================================================================================' -ForegroundColor DarkCyan
    Microsoft.PowerShell.Utility\Write-Host ' SmartM365 by WorkplaceCloudHub' -ForegroundColor Cyan
    Microsoft.PowerShell.Utility\Write-Host ' Website : https://workplacecloudhub.com' -ForegroundColor Yellow
    Microsoft.PowerShell.Utility\Write-Host ' GitHub  : https://github.com/khda79/workplacecloudhub.com' -ForegroundColor Yellow
    Microsoft.PowerShell.Utility\Write-Host '================================================================================' -ForegroundColor DarkCyan
}

function Write-SmartM365CompletionBanner {
    [CmdletBinding()]
    param(
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings')]
        [string]$Status = 'Auto',
        [string]$ScriptName = '',
        [AllowNull()][Nullable[datetime]]$StartedAt = $null,
        [AllowNull()][Nullable[datetime]]$EndedAt = $null,
        [int]$WarningCount = -1,
        [int]$ErrorCount = -1,
        [int]$GeneratedCsvFiles = -1,
        [string]$LogPath = ''
    )
    $scriptNameVariable = Get-Variable -Name SmartM365ScriptName -Scope Global -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($ScriptName) -and $scriptNameVariable) { $ScriptName = [string]$scriptNameVariable.Value }
    $startedAtVariable = Get-Variable -Name SmartM365ExecutionStartTime -Scope Global -ErrorAction SilentlyContinue
    if (($null -eq $StartedAt -or $StartedAt -eq [datetime]::MinValue) -and $startedAtVariable) { $StartedAt = [datetime]$startedAtVariable.Value }
    $warningVariable = Get-Variable -Name SmartM365WarningCount -Scope Global -ErrorAction SilentlyContinue
    if ($WarningCount -lt 0) { $WarningCount = if ($warningVariable) { [int]$warningVariable.Value } else { 0 } }
    $errorVariable = Get-Variable -Name SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
    if ($ErrorCount -lt 0) { $ErrorCount = if ($errorVariable) { [int]$errorVariable.Value } else { 0 } }
    $csvVariable = Get-Variable -Name csvGeneratedPaths -Scope Global -ErrorAction SilentlyContinue
    if ($GeneratedCsvFiles -lt 0) { $GeneratedCsvFiles = if ($csvVariable -and $csvVariable.Value) { @($csvVariable.Value).Count } else { 0 } }
    $logVariable = Get-Variable -Name LogTextFile -Scope Global -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($LogPath) -and $logVariable) { $LogPath = [string]$logVariable.Value }


    if ($null -eq $EndedAt -or $EndedAt -eq [datetime]::MinValue) { $EndedAt = Get-Date }
    if ($null -eq $StartedAt -or $StartedAt -eq [datetime]::MinValue) { $StartedAt = $EndedAt }
    if ([string]::IsNullOrWhiteSpace($ScriptName)) { $ScriptName = 'SmartM365' }

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
    $existingRunKey = [string](Get-Variable -Name SmartM365CompletionBannerRunKey -Scope Global -ValueOnly -ErrorAction SilentlyContinue)
    if ($existingRunKey -eq $runKey) { return }

    $lines = @(
        '================================================================================',
        (' SmartM365 by WorkplaceCloudHub - {0}' -f $title),
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

    Set-Variable -Name SmartM365CompletionBannerRunKey -Scope Global -Value $runKey
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
            Microsoft.PowerShell.Utility\Write-Warning ("SmartM365 completion banner could not be written to '{0}': {1}" -f $effectiveLogPath, $_.Exception.Message)
        }
    }
}

function ConvertTo-SmartM365Hashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = [ordered]@{}
    if ($null -eq $InputObject) { return $hash }

    foreach ($property in $InputObject.PSObject.Properties) {
        $hash[$property.Name] = $property.Value
    }

    return $hash
}

function Initialize-SmartM365LocalJsonFromTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath,
        [string]$ConfigDescription = 'local configuration'
    )

    $Path = Resolve-SmartM365JsonConfigurationPath -Path $Path
    if (Test-Path -LiteralPath $Path) { return $false }

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = if ((Get-SmartM365JsonNames $Path).Legacy -like '*Config\Tenants\*.local.json') {
            Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template'
        }
        elseif ((Get-SmartM365JsonNames $Path).Legacy -like '*.local.json') {
            Get-SmartM365JsonTemplateName -Path $Path
        }
        else {
            ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        $message = @(
            "Local JSON not found: $Path",
            "Template to copy is missing: $TemplatePath",
            'Create the missing local JSON from the matching template, then run the script again.'
        ) -join [Environment]::NewLine
        throw $message
    }

    try {
        Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([IO.File]::ReadAllBytes($TemplatePath)) -ExpectedSHA256 'ABSENT' -Validate {
            param($document)
            if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration template must be an object.' }
        } | Out-Null
    }
    catch {
        throw ("Failed to create {0} '{1}' from template '{2}': {3}" -f $ConfigDescription, $Path, $TemplatePath, $_.Exception.Message)
    }

    $message = @(
        "Created $ConfigDescription from template.",
        "Local JSON: $Path",
        "Template: $TemplatePath",
        'Review the generated local JSON values; continuing with default template values unless edited before next run.'
    ) -join [Environment]::NewLine

    Write-Host $message -ForegroundColor Yellow

    return $true
}

function Get-SmartM365JsonTemplatePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $Path = (Get-SmartM365JsonNames $Path).Legacy
    if ($Path -like '*Config\Tenants\*.local.json') {
        return (Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template')
    }

    if ($Path -like '*.local.json') {
        return ('{0}.template' -f $Path)
    }

    return ''
}

function Add-SmartM365MissingJsonTemplateProperties {
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
            foreach ($nestedPath in (Add-SmartM365MissingJsonTemplateProperties -Target $currentValue -Template $property.Value -PrefixPath $propertyPath)) {
                $added.Add($nestedPath) | Out-Null
            }
        }
    }

    return @($added)
}

function Sync-SmartM365JsonConfigWithTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath,
        [string]$ExpectedSHA256 = ''
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartM365JsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (-not $ExpectedSHA256) {
            $currentRead = Read-SmartM365JsonDocument -Path $Path
            if (($Config | ConvertTo-Json -Depth 100 -Compress) -cne ($currentRead.Document | ConvertTo-Json -Depth 100 -Compress)) {
                throw 'Configuration changed before template enrichment; refusing a lost update.'
            }
            $ExpectedSHA256 = $currentRead.SHA256
            $Path = $currentRead.Path
        }
        $addedKeys = @(Add-SmartM365MissingJsonTemplateProperties -Target $Config -Template $template)
        if ($addedKeys.Count -gt 0) {
            $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($Config | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
            if (-not $ExpectedSHA256) { throw 'Template update requires the hash of the configuration that was read.' }
            Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -ExpectedSHA256 $ExpectedSHA256 -Validate {
                param($document)
                if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration must remain an object.' }
            } | Out-Null
            Write-Host ("Updated local JSON from template: {0}; added keys: {1}" -f $Path, ($addedKeys -join ', ')) -ForegroundColor Yellow
        }
        return $Config
    }
    catch {
        $syncErrorMessage = [string]$PSItem.Exception.Message
        throw ("Failed to synchronize local JSON '{0}' with template '{1}': {2}" -f $Path, $TemplatePath, $syncErrorMessage)
    }
}
function Read-SmartM365JsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    $Path = Resolve-SmartM365JsonConfigurationPath -Path $Path
    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) {
            $templatePath = Get-SmartM365JsonTemplatePath -Path $Path
            Initialize-SmartM365LocalJsonFromTemplate -Path $Path -TemplatePath $templatePath -ConfigDescription 'required local configuration' | Out-Null
        }
        else {
            return [ordered]@{}
        }
    }

    try {
        $read = Read-SmartM365JsonDocument -Path $Path
        $config = ConvertTo-SmartM365Hashtable -InputObject $read.Document
        return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $read.Path -ExpectedSHA256 $read.SHA256)
    }
    catch {
        $message = @(
            ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message),
            'The file is not valid JSON. Check quotes, commas, and Windows paths.',
            'Windows paths in JSON must escape backslashes, for example "Z:\\GIT\\SmartM365", or use forward slashes, for example "Z:/GIT/SmartM365".',
            'Do not write paths with single backslashes such as "Z:\GIT\SmartM365" because JSON treats sequences like \e as invalid escapes.'
        )
        throw ($message -join [Environment]::NewLine)
    }
}

function Find-SmartM365Root {
    [CmdletBinding()]
    param([string]$StartPath)

    $searchRoot = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).Path } else { $StartPath }
    while ($searchRoot) {
        if ((Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json.txt')) -or
            (Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json')) -or
            (Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json.template'))) {
            return $searchRoot
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw "SmartM365 root not found from '$StartPath'."
}

function Test-SmartM365WritableDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        $probePath = Join-Path -Path $Path -ChildPath ('.smartm365-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $probePath -Value 'test' -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Get-SmartM365EffectiveGlobalConfig {
    [CmdletBinding()]
    param(
        [string]$StartPath,
        [Alias('TenantKey')][string]$ProfileKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($ProfileKey)) { $ProfileKey = 'test' }
    $ProfileKey = $ProfileKey.Trim().ToLowerInvariant()

    $rootPath = Find-SmartM365Root -StartPath $StartPath
    $scriptStartPath = if ([string]::IsNullOrWhiteSpace($StartPath)) { $rootPath } else { $StartPath }
    $scriptOutputRootPath = Join-Path -Path $scriptStartPath -ChildPath 'Output'
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'Config\SmartM365.global.local.json'
    $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $ProfileKey)
    $tenantConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $tenantConfigPath

    $globalConfig = Read-SmartM365JsonConfig -Path $globalConfigPath -Required
    $tenantConfig = Read-SmartM365JsonConfig -Path $tenantConfigPath -Required

    $configuredProfileKey = if ($tenantConfig.Contains('ProfileKey')) { [string]$tenantConfig['ProfileKey'] } else { '' }
    $organizationKey = if ($tenantConfig.Contains('OrganizationKey')) { [string]$tenantConfig['OrganizationKey'] } else { '' }
    $environmentKey = if ($tenantConfig.Contains('EnvironmentKey')) { [string]$tenantConfig['EnvironmentKey'] } else { '' }
    $configuredTenantKey = if ($tenantConfig.Contains('TenantKey')) { [string]$tenantConfig['TenantKey'] } else { '' }

    $configuredProfileKey = $configuredProfileKey.Trim().ToLowerInvariant()
    $organizationKey = $organizationKey.Trim().ToLowerInvariant()
    $environmentKey = $environmentKey.Trim().ToLowerInvariant()
    $configuredTenantKey = $configuredTenantKey.Trim().ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($configuredProfileKey) -or $configuredProfileKey -ne $ProfileKey) {
        throw "Tenant profile mismatch. File '$tenantConfigPath' must contain ProfileKey '$ProfileKey'."
    }
    foreach ($identityValue in @(
            [pscustomobject]@{ Name = 'OrganizationKey'; Value = $organizationKey },
            [pscustomobject]@{ Name = 'EnvironmentKey'; Value = $environmentKey }
        )) {
        if ([string]::IsNullOrWhiteSpace($identityValue.Value) -or $identityValue.Value -notmatch '^[a-z0-9][a-z0-9-]*$') {
            throw "Tenant profile '$tenantConfigPath' requires $($identityValue.Name) using lowercase letters, digits, and hyphens."
        }
    }

    $expectedTenantKey = ('{0}-{1}' -f $organizationKey, $environmentKey)
    if ($configuredTenantKey -ne $expectedTenantKey) {
        throw "Tenant profile '$tenantConfigPath' must contain TenantKey '$expectedTenantKey' (OrganizationKey-EnvironmentKey), found '$configuredTenantKey'."
    }

    foreach ($key in $tenantConfig.Keys) {
        $tenantValue = $tenantConfig[$key]
        if ($tenantValue -is [string]) {
            $tenantText = $tenantValue.Trim()
            if ([string]::IsNullOrWhiteSpace($tenantText) -or $tenantText -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { continue }
        }
        $globalConfig[$key] = $tenantValue
    }

    $globalConfig['ProfileKey'] = $ProfileKey
    $globalConfig['OrganizationKey'] = $organizationKey
    $globalConfig['EnvironmentKey'] = $environmentKey
    $globalConfig['TenantKey'] = $expectedTenantKey
    $globalConfig['SmartM365RootPath'] = $rootPath
    $globalConfig['ScriptOutputRootPath'] = $scriptOutputRootPath

    $defaultWorkspaceRootPath = $rootPath
    $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'
    $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-LAST'
    $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\LOG-ALL'
    $useScriptOutputFallback = $false

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or [string]$globalConfig['WorkspaceRootPath'] -eq '{{SmartM365RootPath}}') {
        $candidateDataRoot = Join-Path -Path $rootPath -ChildPath 'Data'
        if (Test-SmartM365WritableDirectory -Path $candidateDataRoot) {
            $defaultWorkspaceRootPath = $rootPath
        }
        else {
            $defaultWorkspaceRootPath = $scriptOutputRootPath
            $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-ALL'
            $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-LAST'
            $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\LOG-ALL'
            $useScriptOutputFallback = $true
        }
    }

    $legacyDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
    $legacyLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
    $legacyLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
    if (-not $globalConfig.Contains('DataAllRootPath') -or ($useScriptOutputFallback -and [string]$globalConfig['DataAllRootPath'] -in @($legacyDataAllRootPath, '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'))) { $globalConfig['DataAllRootPath'] = $defaultDataAllRootPath }
    if (-not $globalConfig.Contains('LatestCsvFolderPath') -or ($useScriptOutputFallback -and [string]$globalConfig['LatestCsvFolderPath'] -in @($legacyLatestCsvFolderPath, '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-LAST'))) { $globalConfig['LatestCsvFolderPath'] = $defaultLatestCsvFolderPath }
    if (-not $globalConfig.Contains('LogAllRootPath') -or ($useScriptOutputFallback -and [string]$globalConfig['LogAllRootPath'] -in @($legacyLogAllRootPath, '{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\LOG-ALL'))) { $globalConfig['LogAllRootPath'] = $defaultLogAllRootPath }
    if (-not $globalConfig.Contains('WorkspaceRootPath') -or [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or [string]$globalConfig['WorkspaceRootPath'] -eq '{{SmartM365RootPath}}') { $globalConfig['WorkspaceRootPath'] = $defaultWorkspaceRootPath }

    foreach ($pathKey in @('DataAllRootPath', 'LatestCsvFolderPath', 'LogAllRootPath')) {
        if ($globalConfig.Contains($pathKey) -and $globalConfig[$pathKey] -is [string]) {
            $globalConfig[$pathKey] = ([string]$globalConfig[$pathKey]).Replace('{{TenantKey}}', '{{ProfileKey}}')
        }
    }

    return [pscustomobject]$globalConfig
}
function Initialize-SmartM365TenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    Write-SmartM365StartupBanner

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }
    $profileKey = $Tenant.Trim().ToLowerInvariant()
    $effectiveConfig = Get-SmartM365EffectiveGlobalConfig -StartPath $StartPath -ProfileKey $profileKey

    $global:SmartM365Tenant = $profileKey
    $global:SmartM365ProfileKey = [string]$effectiveConfig.ProfileKey
    $global:SmartM365OrganizationKey = [string]$effectiveConfig.OrganizationKey
    $global:SmartM365EnvironmentKey = [string]$effectiveConfig.EnvironmentKey
    $global:SmartM365TenantKey = [string]$effectiveConfig.TenantKey
    $global:SmartM365TenantId = [string]$effectiveConfig.TenantId
    $configuredMailTenantName = ''
    if ($effectiveConfig -is [System.Collections.IDictionary]) {
        if ($effectiveConfig.Contains('MailClientName')) {
            $configuredMailTenantName = [string]$effectiveConfig['MailClientName']
        }
    }
    elseif ($null -ne $effectiveConfig) {
        $mailClientNameProperty = $effectiveConfig.PSObject.Properties['MailClientName']
        if ($null -ne $mailClientNameProperty) {
            $configuredMailTenantName = [string]$mailClientNameProperty.Value
        }
    }
    $global:SmartM365MailTenantName = if (-not [string]::IsNullOrWhiteSpace($configuredMailTenantName)) {
        $configuredMailTenantName.Trim()
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$effectiveConfig.OrganizationKey)) {
        ([string]$effectiveConfig.OrganizationKey).Trim().ToUpperInvariant()
    }
    else {
        [string]$effectiveConfig.TenantKey
    }
    $global:SmartM365GlobalConfig = $effectiveConfig
    $script:SmartM365GlobalConfig = $effectiveConfig
    return $effectiveConfig
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAd6Eiu28sBIo4K
# zDc3cmsQRYLwuSeVrnPhzM8JzXGus6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPKGjJaLHJV8MnWnQKtAuG2AXxUWNBYQVX9gvbih4ffbMA0GCSqG
# SIb3DQEBAQUABIIBgKLA7ELYPTcZVCJVR/kKkVaPgB+bglXUsr3/xGkwBCd95mKh
# +3hguoCX6idbTOjOrgAm/SNwRq1Gv71jCS3t7JOI/aHNy7Fbj0j8CtezBnEBayT+
# bKV6/xc240E8F3poG4dGYPx+od1m0nAABn43thi4/YONwNf2mOfPAEXf7EXY7y1C
# Mi6F8Fe1Z8GJ2uYS7zidT/Ze0QFmIDcHsg9d0CIAAgSnFzxUH4TEr0k8QN8xzV9s
# pui3VEIXFa5/7noPc/npyCMSEf7hRoTIW+Zhz7aDAx4ICloyBfhCcuAgwKpbQHCG
# UF73GQkvkL6wT1Vc9/ubxXYw/Y2RMDUUMA8HwqaMJX061EJx9hdf5A1hCaFgSc+d
# G3znXl2kog0xVAnHiRLkoQHsfcKGNKYVEuwzWbtPptYG17stZh+dY0SENKNFZgFi
# M9yzQbWkhmD9EiNKqad/qic+JYyXDdqiXZZeDYezgXvViTmrr56vETEkhvAgtf6L
# PYBqEzNeorJiwp0eoqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# MzVaMC8GCSqGSIb3DQEJBDEiBCD4+ZPD3gvouuvDcqtBUztEz5pOTGXvl+VVRCGy
# OjVwJDANBgkqhkiG9w0BAQEFAASCAgCxrCpQYKYiwJ9OxadXwxqX7FIPnxdl9gLm
# ti8kNvzOd2ERP2vqxtJ/bkTaxDShkkH6GRKOUZeT1oHLEerux0Vx4IUTyN2FSFOE
# wTX9IocFyDQgsdX/54MJFVhNFgBZ8fquIbVaghDgI3qY6ZkuKz4O5oPWHBr3IslH
# NyuN/yBpMhtTBaeZUuZQrjPhMlbmdo4BuK9ESJrTn/AQXJmOOZBPRhWDR9z3vhng
# HO1/6fGeyQS18AwCXHkX7eQq4KHg/fVACQkSCcYGhrqTEnkEwcDtrQOpQFR9iLdV
# 5Sl8LVcitbzHP2HpLENnZrk5hXEK9UU+U/QZ+pJr4LdR5eK8Wd2VnxexAmLJMvMz
# 2dN4Jdpn8HOHdb3yCv4fkPZI+eKGJzwIR3MsxY0c8TCLtPVkPjyM5mVbt0ow+0au
# /KF9PMkhjqDD1SxG5whSSfs55IjlOfhrhlmLq8R+oyec7GP3ZZicLReSRWAEQncc
# aQLP7YvDcPRjzdfxjM5E6QlizAOrx65TRxU3eUIeXjfBPYMIXb2hxfRtooxtTil5
# llzcBQBjIm+NIvnap0AlypgNJrYYqS989uLQzBAALRH5zxFaOGqqZsd2bd1KjA1n
# znbPeIq6z7lHLnyTSkixQpQotAL0uhJdGNA8cGZ0LE8Qk/7dBvXuXTnIt4HqdVWU
# vyoxLc4vgA==
# SIG # End signature block
