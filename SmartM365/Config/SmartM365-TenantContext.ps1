Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.8' -Global -ErrorAction Stop
<#
.SYNOPSIS
Loads SmartM365 global and tenant configuration context.

.DESCRIPTION
Merges global and tenant local JSON configuration, resolves workspace paths, and exposes the effective tenant context used by SmartM365 scripts.

.VERSION
1.0.12
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
        [ValidateSet('Auto', 'Success', 'Failed', 'CompletedWithWarnings', 'Cancelled')]
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

    # In the cleanup of a SmartInventory run, the shared lifecycle decides like the Core banner does: a run
    # interrupted before its end (Ctrl+C) is Cancelled, never Success. Completed runs keep their status.
    $lifecycle = Get-Variable -Name SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($lifecycle -and $lifecycle.InCleanup -and (Get-Command -Name Get-SmartM365InventoryCompletionStatus -ErrorAction SilentlyContinue)) {
        $Status = Get-SmartM365InventoryCompletionStatus -Status $Status -ErrorRecord $null
    }

    if ($Status -eq 'Auto') {
        $Status = if ($ErrorCount -gt 0) { 'Failed' } elseif ($WarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    }

    $statusLabel = switch ($Status) {
        'Failed' { 'FAILED' }
        'CompletedWithWarnings' { 'COMPLETED WITH WARNINGS' }
        'Cancelled' { 'CANCELLED' }
        default { 'SUCCESS' }
    }
    $title = switch ($Status) {
        'Failed' { 'Execution failed' }
        'CompletedWithWarnings' { 'Execution completed with warnings' }
        'Cancelled' { 'Execution cancelled' }
        default { 'Execution completed' }
    }
    $statusColor = switch ($Status) {
        'Failed' { 'Red' }
        'CompletedWithWarnings' { 'Yellow' }
        'Cancelled' { 'Yellow' }
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
        $TemplatePath = Get-SmartM365JsonTemplatePath -Path $Path
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
    $legacyTemplate = if ($Path -like '*Config\Tenants\*.local.json') {
        Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template'
    }
    elseif ($Path -like '*.local.json') { '{0}.template' -f $Path }
    else { return '' }
    $preferredTemplate = $legacyTemplate -replace '\.json\.template$', '.json.txt.template'
    # Presence wins, even for an invalid file/directory: readers must fail, not fall back.
    if (Test-Path -LiteralPath $preferredTemplate) { return $preferredTemplate }
    return $legacyTemplate
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
        [string]$ExpectedSHA256 = '',
        [ValidateRange(0, 3600)][int]$ConfigurationWriteTimeoutSeconds = 60
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartM365JsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($template -isnot [System.Management.Automation.PSCustomObject]) {
            throw 'Configuration template must be an object.'
        }
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
            Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -ExpectedSHA256 $ExpectedSHA256 -ReplacementTimeoutSeconds $ConfigurationWriteTimeoutSeconds -Validate {
                param($document)
                if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration must remain an object.' }
            } | Out-Null
            Write-Host ("Updated local JSON from template: {0}; added keys: {1}" -f $Path, ($addedKeys -join ', ')) -ForegroundColor Yellow
        }
        return $Config
    }
    catch {
        $syncErrorMessage = [string]$PSItem.Exception.Message
        throw [InvalidOperationException]::new(("Failed to synchronize local JSON '{0}' with template '{1}': {2}" -f $Path, $TemplatePath, $syncErrorMessage), $_.Exception)
    }
}
function Read-SmartM365JsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required,
        [ValidateRange(0, 3600)][int]$ConfigurationWriteTimeoutSeconds = 60
    )

    $phase = 'read'
    try {
        $Path = Resolve-SmartM365JsonConfigurationPath -Path $Path
        if (-not (Test-Path -LiteralPath $Path)) {
            if ($Required) {
                $phase = 'initialization'
                $templatePath = Get-SmartM365JsonTemplatePath -Path $Path
                Initialize-SmartM365LocalJsonFromTemplate -Path $Path -TemplatePath $templatePath -ConfigDescription 'required local configuration' | Out-Null
            }
            else {
                return [ordered]@{}
            }
        }
        $read = Read-SmartM365JsonDocument -Path $Path
        $config = ConvertTo-SmartM365Hashtable -InputObject $read.Document
        $phase = 'template update'
        return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $read.Path -ExpectedSHA256 $read.SHA256 -ConfigurationWriteTimeoutSeconds $ConfigurationWriteTimeoutSeconds)
    }
    catch {
        $invalidSyntax = $false
        for ($cause = $_.Exception; $null -ne $cause; $cause = $cause.InnerException) {
            if ($cause.Data['SmartM365JsonInvalidSyntax']) { $invalidSyntax = $true }
        }
        $message = @(("Failed to load configuration file '{0}' during {1}: {2}" -f $Path, $phase, $_.Exception.Message))
        if ($invalidSyntax) {
            $message += 'Invalid JSON syntax. Check quotes, commas, and escaped backslashes in Windows paths.'
        }
        throw [InvalidOperationException]::new(($message -join [Environment]::NewLine), $_.Exception)
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
        $probePath = Get-SmartM365SiblingTemporaryPath -Path (Join-Path -Path $Path -ChildPath 'probe') -Extension '.tmp'
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
