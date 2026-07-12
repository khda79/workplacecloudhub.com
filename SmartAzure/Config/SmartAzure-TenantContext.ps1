function ConvertTo-SmartAzureHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = [ordered]@{}
    if ($null -eq $InputObject) { return $hash }

    foreach ($property in $InputObject.PSObject.Properties) {
        $hash[$property.Name] = $property.Value
    }

    return $hash
}

function Get-SmartAzureJsonTemplatePath {
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

function Add-SmartAzureMissingJsonTemplateProperties {
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
            foreach ($nestedPath in (Add-SmartAzureMissingJsonTemplateProperties -Target $currentValue -Template $property.Value -PrefixPath $propertyPath)) {
                $added.Add($nestedPath) | Out-Null
            }
        }
    }

    return @($added)
}

function Sync-SmartAzureJsonConfigWithTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartAzureJsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $addedKeys = @(Add-SmartAzureMissingJsonTemplateProperties -Target $Config -Template $template)
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
function Read-SmartAzureJsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) { throw "Configuration file not found: $Path" }
        return [ordered]@{}
    }

    try {
        $config = ConvertTo-SmartAzureHashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        return (Sync-SmartAzureJsonConfigWithTemplate -Config $config -Path $Path)
    }
    catch {
        throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Find-SmartAzureRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    $searchRoot = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).Path } else { $StartPath }
    $resolvedSearchRoot = Resolve-Path -LiteralPath $searchRoot -ErrorAction SilentlyContinue
    if ($null -ne $resolvedSearchRoot) {
        $searchRoot = $resolvedSearchRoot.Path
    }
    while ($searchRoot) {
        if ((Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartAzure.global.local.json')) -or
            (Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartAzure.global.local.json.template'))) {
            return $searchRoot
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw "SmartAzure root not found from '$StartPath'."
}

function Import-SmartAzureCoreModule {
    [CmdletBinding()]
    param([string]$StartPath)

    if (Get-Command -Name Set-SmartAzureCoreContext -ErrorAction SilentlyContinue) {
        return
    }

    $rootPath = Find-SmartAzureRoot -StartPath $StartPath
    $modulePath = Join-Path -Path $rootPath -ChildPath 'Modules\SmartAzure.Core\SmartAzure.Core.psd1'
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "SmartAzure.Core module not found: $modulePath"
    }

    Import-Module $modulePath -Global -Force -ErrorAction Stop
}

function Test-SmartAzureWritableDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        $probePath = Join-Path -Path $Path -ChildPath ('.smartazure-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $probePath -Value 'test' -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Get-SmartAzureEffectiveGlobalConfig {
    [CmdletBinding()]
    param(
        [string]$StartPath,
        [string]$TenantKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($TenantKey)) { $TenantKey = 'test' }

    $rootPath = Find-SmartAzureRoot -StartPath $StartPath
    $scriptStartPath = if ([string]::IsNullOrWhiteSpace($StartPath)) { $rootPath } else { $StartPath }
    $scriptOutputRootPath = Join-Path -Path $scriptStartPath -ChildPath 'Output'
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'SmartAzure.global.local.json'
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $TenantKey)

    $globalConfig = Read-SmartAzureJsonConfig -Path $globalConfigPath
    $tenantConfig = Read-SmartAzureJsonConfig -Path $tenantConfigPath

    if ($tenantConfig.Contains('TenantKey') -and
        -not [string]::IsNullOrWhiteSpace([string]$tenantConfig['TenantKey']) -and
        [string]$tenantConfig['TenantKey'] -ne $TenantKey) {
        throw "Tenant profile key mismatch. File '$tenantConfigPath' contains TenantKey '$($tenantConfig['TenantKey'])' but requested '$TenantKey'."
    }

    foreach ($key in $tenantConfig.Keys) {
        $globalConfig[$key] = $tenantConfig[$key]
    }

    $globalConfig['TenantKey'] = $TenantKey
    $globalConfig['SmartAzureRootPath'] = $rootPath
    $globalConfig['ScriptOutputRootPath'] = $scriptOutputRootPath

    $defaultWorkspaceRootPath = $rootPath
    $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
    $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
    $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
    $useScriptOutputFallback = $false

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or
        [string]$globalConfig['WorkspaceRootPath'] -in @('.', '{{SmartAzureRootPath}}')) {
        $candidateDataRoot = Join-Path -Path $rootPath -ChildPath 'Data'
        if (Test-SmartAzureWritableDirectory -Path $candidateDataRoot) {
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

    if (-not $globalConfig.Contains('DataAllRootPath') -or
        ($useScriptOutputFallback -and [string]$globalConfig['DataAllRootPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL')) {
        $globalConfig['DataAllRootPath'] = $defaultDataAllRootPath
    }
    if (-not $globalConfig.Contains('LatestCsvFolderPath') -or
        ($useScriptOutputFallback -and [string]$globalConfig['LatestCsvFolderPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST')) {
        $globalConfig['LatestCsvFolderPath'] = $defaultLatestCsvFolderPath
    }
    if (-not $globalConfig.Contains('LogAllRootPath') -or
        ($useScriptOutputFallback -and [string]$globalConfig['LogAllRootPath'] -eq '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL')) {
        $globalConfig['LogAllRootPath'] = $defaultLogAllRootPath
    }
    if (-not $globalConfig.Contains('WorkspaceRootPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or
        [string]$globalConfig['WorkspaceRootPath'] -in @('.', '{{SmartAzureRootPath}}')) {
        $globalConfig['WorkspaceRootPath'] = $defaultWorkspaceRootPath
    }

    return [pscustomobject]$globalConfig
}

function Resolve-SmartAzureConfigTokenValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($null -eq $script:SmartAzureGlobalConfig) {
        throw 'SmartAzure tenant context has not been initialized.'
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartAzureGlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartAzureConfigTokenValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}

function Resolve-SmartAzureOutputRoots {
    [CmdletBinding()]
    param(
        [string]$OutputRoot,
        [string]$LatestOutputRoot,
        [Parameter(Mandatory)][string]$AreaPath
    )

    if ($null -eq $script:SmartAzureGlobalConfig) {
        throw 'SmartAzure tenant context has not been initialized.'
    }

    $resolvedOutputRoot = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        Join-Path -Path (Resolve-SmartAzureConfigTokenValue -Value '{{DataAllRootPath}}') -ChildPath $AreaPath
    }
    else {
        Resolve-SmartAzureConfigTokenValue -Value $OutputRoot
    }

    $resolvedLatestOutputRoot = if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) {
        Resolve-SmartAzureConfigTokenValue -Value '{{LatestCsvFolderPath}}'
    }
    else {
        Resolve-SmartAzureConfigTokenValue -Value $LatestOutputRoot
    }

    [pscustomobject]@{
        OutputRoot       = $resolvedOutputRoot
        LatestOutputRoot = $resolvedLatestOutputRoot
    }
}

function Get-SmartAzureConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $script:SmartAzureGlobalConfig) {
        throw 'SmartAzure tenant context has not been initialized.'
    }

    $property = $script:SmartAzureGlobalConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }

    if ($property.Value -is [string]) {
        $value = $property.Value.Trim()
        if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) {
            return $DefaultValue
        }
    }

    return Resolve-SmartAzureConfigTokenValue -Value $property.Value
}

function Get-SmartAzureScriptLocalConfig {
    [CmdletBinding()]
    param([string]$ScriptPath)

    $effectiveScriptPath = if ([string]::IsNullOrWhiteSpace($ScriptPath)) { $PSCommandPath } else { $ScriptPath }
    if ([string]::IsNullOrWhiteSpace($effectiveScriptPath)) {
        return [pscustomobject]@{}
    }

    $scriptFolder = Split-Path -Path $effectiveScriptPath -Parent
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($effectiveScriptPath)
    $configPath = Join-Path -Path $scriptFolder -ChildPath ("{0}.local.json" -f $scriptBaseName)

    if (-not (Test-Path -LiteralPath $configPath)) {
        return [pscustomobject]@{}
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return (Sync-SmartAzureJsonConfigWithTemplate -Config $config -Path $configPath)
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Get-SmartAzureScriptConfigValue {
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
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return Resolve-SmartAzureConfigTokenValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartAzureConfigTokenValue -Value $property.Value
        }
    }

    return Get-SmartAzureConfigValue -Name $Name -DefaultValue $DefaultValue
}

function Test-SmartAzureConfiguredValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $false }
    if ($Value -isnot [string]) { return $true }

    $trimmed = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $false }
    if ($trimmed -match '^0{8}-0{4}-0{4}-0{4}-0{12}$') { return $false }
    if ($trimmed -match '^0{40}$') { return $false }
    if ($trimmed -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $false }
    return $true
}

function ConvertTo-SmartAzureRecipientArray {
    [CmdletBinding()]
    param([string]$Recipients)

    if ([string]::IsNullOrWhiteSpace($Recipients)) { return @() }
    return @($Recipients -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function ConvertTo-SmartAzureGraphRecipient {
    [CmdletBinding()]
    param([string[]]$Recipients)

    foreach ($recipient in @($Recipients | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        @{
            emailAddress = @{
                address = $recipient
            }
        }
    }
}

function ConvertTo-SmartAzureGraphFileAttachment {
    [CmdletBinding()]
    param([string[]]$Attachments)

    $graphAttachments = @()
    foreach ($attachmentPath in @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Test-Path -LiteralPath $attachmentPath -PathType Leaf)) { continue }

        $file = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop
        if ($file.Length -gt 3MB) {
            throw ("Graph mail attachment '{0}' is larger than 3 MB. Large attachments require an upload session." -f $file.FullName)
        }

        $graphAttachments += @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = $file.Name
            contentType   = 'application/octet-stream'
            contentBytes  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
        }
    }

    return $graphAttachments
}

function Get-SmartAzureExceptionDetails {
    [CmdletBinding()]
    param([AllowNull()]$Exception)

    if ($null -eq $Exception) { return '' }

    $messages = New-Object System.Collections.Generic.List[string]
    $current = $Exception
    while ($null -ne $current) {
        if (-not [string]::IsNullOrWhiteSpace($current.Message)) {
            $messages.Add($current.Message) | Out-Null
        }
        $current = $current.InnerException
    }

    if ($messages.Count -eq 0) { return [string]$Exception }
    return ($messages | Select-Object -Unique) -join ' | '
}

function Connect-SmartAzureGraphAppOnly {
    [CmdletBinding()]
    param(
        [string]$AppId = $(Get-SmartAzureConfigValue -Name 'AppId' -DefaultValue ''),
        [string]$TenantId = $(Get-SmartAzureConfigValue -Name 'TenantId' -DefaultValue ''),
        [string]$Thumbprint = $(Get-SmartAzureConfigValue -Name 'Thumbprint' -DefaultValue ''),
        [string]$Purpose = 'Microsoft Graph'
    )

    try {
        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        }

        if (-not (Test-SmartAzureConfiguredValue -Value $Thumbprint)) {
            $Thumbprint = Get-SmartAzureConfigValue -Name 'Thumb' -DefaultValue ''
        }

        if (-not (Test-SmartAzureConfiguredValue -Value $AppId) -or
            -not (Test-SmartAzureConfiguredValue -Value $TenantId) -or
            -not (Test-SmartAzureConfiguredValue -Value $Thumbprint)) {
            Write-Host ("{0} skipped: Graph app-only connection values are missing (AppId, TenantId, Thumb/Thumbprint)." -f $Purpose)
            return $false
        }

        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null

        $context = Get-MgContext -ErrorAction SilentlyContinue
        return ($null -ne $context)
    }
    catch {
        Write-Host ("{0} skipped: failed to connect Microsoft Graph: {1}" -f $Purpose, (Get-SmartAzureExceptionDetails -Exception $_.Exception))
        return $false
    }
}

function Write-SmartAzureSharePointUploadLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    if (Get-Command -Name Write-SmartAzureLog -ErrorAction SilentlyContinue) {
        Write-SmartAzureLog -Level $Level -Message $Message
        return
    }

    Write-Host ("{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
}

function ConvertTo-SmartAzureGraphDrivePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    (($Path -replace '\\', '/') -replace '^/+', '') -replace ' ', '%20'
}

function Invoke-SmartAzureSharePointCsvUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [bool]$Enabled = [bool](Get-SmartAzureConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false),
        [string]$SiteHostname = $(Get-SmartAzureConfigValue -Name 'SharePointSiteHostname' -DefaultValue ''),
        [string]$SitePath = $(Get-SmartAzureConfigValue -Name 'SharePointSitePath' -DefaultValue ''),
        [string]$LibraryDisplayName = $(Get-SmartAzureConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'),
        [string]$TargetFolderPath = $(Get-SmartAzureConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue ''),
        [string]$AppId = $(Get-SmartAzureConfigValue -Name 'AppId' -DefaultValue ''),
        [string]$TenantId = $(Get-SmartAzureConfigValue -Name 'TenantId' -DefaultValue ''),
        [string]$Thumbprint = $(Get-SmartAzureConfigValue -Name 'Thumbprint' -DefaultValue '')
    )

    if (-not $Enabled) { return }
    if (-not (Test-Path -LiteralPath $LocalFilePath -PathType Leaf)) {
        Write-SmartAzureSharePointUploadLog -Level WARN -Message "SharePoint upload skipped: local CSV not found: $LocalFilePath"
        return
    }
    if ([string]::IsNullOrWhiteSpace($SiteHostname) -or [string]::IsNullOrWhiteSpace($SitePath) -or [string]::IsNullOrWhiteSpace($LibraryDisplayName) -or [string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        Write-SmartAzureSharePointUploadLog -Level WARN -Message 'SharePoint upload skipped: SharePointSiteHostname, SharePointSitePath, SharePointLibraryDisplayName or SharePointTargetFolderPath is missing.'
        return
    }
    if (-not (Test-SmartAzureConfiguredValue -Value $Thumbprint)) {
        $Thumbprint = Get-SmartAzureConfigValue -Name 'Thumb' -DefaultValue ''
    }
    if (-not (Connect-SmartAzureGraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'SharePoint upload')) {
        return
    }

    try {
        $site = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHostname, $SitePath)
        $drives = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives" -f $site.id)
        $driveList = @($drives.value)
        $normalize = { param($Text) if ($null -eq $Text) { '' } else { ([string]$Text).Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{M}', '' } }
        $drive = @($driveList | Where-Object { $_.name -ieq $LibraryDisplayName } | Select-Object -First 1)[0]
        if (-not $drive) {
            $targetNorm = & $normalize $LibraryDisplayName
            $drive = @($driveList | Where-Object { (& $normalize $_.name) -ieq $targetNorm } | Select-Object -First 1)[0]
        }
        if (-not $drive) {
            $available = ($driveList | ForEach-Object { $_.name }) -join ' | '
            Write-SmartAzureSharePointUploadLog -Level WARN -Message "SharePoint upload skipped: document library '$LibraryDisplayName' not found. Available drives: $available"
            return
        }

        $fileName = [System.IO.Path]::GetFileName($LocalFilePath)
        $targetPath = ConvertTo-SmartAzureGraphDrivePath (Join-Path -Path $TargetFolderPath -ChildPath $fileName)
        $bytes = [System.IO.File]::ReadAllBytes($LocalFilePath)
        $uri = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/content" -f $drive.id, $targetPath
        Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $bytes -ContentType 'application/octet-stream' | Out-Null
        Write-SmartAzureSharePointUploadLog -Level SUCCESS -Message "SharePoint CSV uploaded: $TargetFolderPath/$fileName"
    }
    catch {
        Write-SmartAzureSharePointUploadLog -Level WARN -Message ("SharePoint upload failed but script continues: {0}" -f $_.Exception.Message)
    }
}

function Send-SmartAzureGraphMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string]$To,
        [string]$Cc = '',
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$BodyHtml,
        [string[]]$Attachments,
        [string]$AppId = $(Get-SmartAzureConfigValue -Name 'AppId' -DefaultValue ''),
        [string]$TenantId = $(Get-SmartAzureConfigValue -Name 'TenantId' -DefaultValue ''),
        [string]$Thumbprint = $(Get-SmartAzureConfigValue -Name 'Thumbprint' -DefaultValue '')
    )

    $toArray = ConvertTo-SmartAzureRecipientArray -Recipients $To
    $ccArray = if ($Cc) { ConvertTo-SmartAzureRecipientArray -Recipients $Cc } else { @() }

    if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
        throw 'Send-SmartAzureGraphMail: missing required parameters (From/To).'
    }

    if (-not (Connect-SmartAzureGraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint -Purpose 'Graph mail')) {
        throw 'Send-SmartAzureGraphMail: Microsoft Graph app-only connection failed.'
    }

    $message = @{
        subject      = $Subject
        body         = @{
            contentType = 'HTML'
            content     = $BodyHtml
        }
        toRecipients = @(ConvertTo-SmartAzureGraphRecipient -Recipients $toArray)
    }

    if ($ccArray.Count -gt 0) {
        $message['ccRecipients'] = @(ConvertTo-SmartAzureGraphRecipient -Recipients $ccArray)
    }

    $graphAttachments = @(ConvertTo-SmartAzureGraphFileAttachment -Attachments $Attachments)
    if ($graphAttachments.Count -gt 0) {
        $message['attachments'] = $graphAttachments
    }

    $body = @{
        message         = $message
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 12

    $encodedFrom = [System.Uri]::EscapeDataString($From)
    Invoke-MgGraphRequest -Method POST -Uri ("https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $encodedFrom) -Body $body -ContentType 'application/json' | Out-Null
}

function Send-SmartAzureMail {
    [CmdletBinding()]
    param(
        [string]$SmtpServer = $(Get-SmartAzureConfigValue -Name 'SmtpServer' -DefaultValue ''),
        [int]$SmtpPort = [int](Get-SmartAzureConfigValue -Name 'SmtpPort' -DefaultValue 25),
        [string]$From = $(Get-SmartAzureConfigValue -Name 'From' -DefaultValue ''),
        [string]$To = $(Get-SmartAzureConfigValue -Name 'ErrorMailTo' -DefaultValue ''),
        [string]$Cc = $(Get-SmartAzureConfigValue -Name 'Cc' -DefaultValue ''),
        [string]$Subject = $(Get-SmartAzureConfigValue -Name 'Subject' -DefaultValue 'SmartAzure'),
        [string]$Body,
        [string]$BodyHtml,
        [string[]]$Attachments,
        [string]$AppId = $(Get-SmartAzureConfigValue -Name 'AppId' -DefaultValue ''),
        [string]$TenantId = $(Get-SmartAzureConfigValue -Name 'TenantId' -DefaultValue ''),
        [string]$Thumbprint = $(Get-SmartAzureConfigValue -Name 'Thumbprint' -DefaultValue '')
    )

    $toArray = ConvertTo-SmartAzureRecipientArray -Recipients $To
    $ccArray = if ($Cc) { ConvertTo-SmartAzureRecipientArray -Recipients $Cc } else { @() }
    $htmlBody = if (-not [string]::IsNullOrWhiteSpace($BodyHtml)) { $BodyHtml } else { $Body }

    if (-not $toArray -or [string]::IsNullOrWhiteSpace($From)) {
        throw 'Send-SmartAzureMail: missing required parameters (From/To).'
    }

    if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
        Send-SmartAzureGraphMail -From $From -To ($toArray -join ';') -Cc ($ccArray -join ';') -Subject $Subject -BodyHtml $htmlBody -Attachments $Attachments -AppId $AppId -TenantId $TenantId -Thumbprint $Thumbprint
        return
    }

    $mailParams = @{
        SmtpServer  = $SmtpServer
        Port        = $SmtpPort
        From        = $From
        To          = $toArray
        Subject     = $Subject
        Body        = $htmlBody
        BodyAsHtml  = $true
        ErrorAction = 'Stop'
    }
    if ($ccArray.Count -gt 0) { $mailParams['Cc'] = $ccArray }
    if (@($Attachments).Count -gt 0) { $mailParams['Attachments'] = @($Attachments | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) }) }

    Send-MailMessage @mailParams
}

function Send-SmartAzureTeamsNotification {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$WebhookUrl = '',
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO',
        [ValidateSet('Auto','Alerts','Infos')]
        [string]$Channel = 'Auto',
        [hashtable]$Facts,
        [switch]$ThrowOnError
    )

    try {
        $effectiveChannel = if ($Channel -ne 'Auto') { $Channel } elseif ($Level -eq 'ERROR') { 'Alerts' } else { 'Infos' }
        if (-not $PSBoundParameters.ContainsKey('WebhookUrl') -or [string]::IsNullOrWhiteSpace($WebhookUrl)) {
            $webhookConfigName = if ($effectiveChannel -eq 'Alerts') { 'TeamsAlertsWebhookUrl' } else { 'TeamsInfosWebhookUrl' }
            $WebhookUrl = Get-SmartAzureConfigValue -Name $webhookConfigName -DefaultValue ''
            if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
                $WebhookUrl = Get-SmartAzureConfigValue -Name 'TeamsWebhookUrl' -DefaultValue ''
            }
        }

        if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
            return $false
        }

        $themeColors = @{
            INFO    = '0078D4'
            SUCCESS = '107C10'
            WARNING = 'FFB900'
            ERROR   = 'D13438'
        }

        $factList = New-Object System.Collections.Generic.List[hashtable]
        $factList.Add(@{ name = 'Level'; value = $Level }) | Out-Null
        $factList.Add(@{ name = 'Timestamp'; value = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }) | Out-Null
        $factList.Add(@{ name = 'Computer'; value = $env:COMPUTERNAME }) | Out-Null
        if ($null -ne $Facts) {
            foreach ($key in @($Facts.Keys | Sort-Object)) {
                $value = $Facts[$key]
                if ($null -eq $value) { continue }
                $factList.Add(@{ name = [string]$key; value = [string]$value }) | Out-Null
            }
        }

        $payload = @{
            '@type'    = 'MessageCard'
            '@context' = 'https://schema.org/extensions'
            summary    = $Title
            themeColor = $themeColors[$Level]
            title      = $Title
            text       = $Message
            sections   = @(
                @{
                    markdown = $true
                    facts    = @($factList)
                }
            )
        }

        Invoke-RestMethod -Method POST -Uri $WebhookUrl -ContentType 'application/json; charset=utf-8' -Body ($payload | ConvertTo-Json -Depth 8) -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        if ($ThrowOnError) { throw }
        return $false
    }
}

function Send-SmartAzureScriptFailureNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$RunId = '',
        [string]$LogPath = '',
        $Config = $null,
        [hashtable]$Facts
    )

    $message = Get-SmartAzureExceptionDetails -Exception $ErrorRecord.Exception
    $effectiveFacts = @{}
    if ($null -ne $Facts) {
        foreach ($key in $Facts.Keys) { $effectiveFacts[$key] = $Facts[$key] }
    }
    $effectiveFacts['Script'] = $ScriptName
    if (-not [string]::IsNullOrWhiteSpace($RunId)) { $effectiveFacts['RunId'] = $RunId }
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { $effectiveFacts['LogPath'] = $LogPath }
    $effectiveFacts['Tenant'] = $global:SmartAzureTenant

    $teamsWebhookUrl = if ($null -ne $Config) {
        Get-SmartAzureScriptConfigValue -Config $Config -Name 'TeamsAlertsWebhookUrl' -DefaultValue ''
    }
    else {
        Get-SmartAzureConfigValue -Name 'TeamsAlertsWebhookUrl' -DefaultValue ''
    }
    if ([string]::IsNullOrWhiteSpace($teamsWebhookUrl)) {
        $teamsWebhookUrl = if ($null -ne $Config) {
            Get-SmartAzureScriptConfigValue -Config $Config -Name 'TeamsWebhookUrl' -DefaultValue ''
        }
        else {
            Get-SmartAzureConfigValue -Name 'TeamsWebhookUrl' -DefaultValue ''
        }
    }
    [void](Send-SmartAzureTeamsNotification -WebhookUrl $teamsWebhookUrl -Title ("SmartAzure error: {0}" -f $ScriptName) -Message $message -Level 'ERROR' -Channel 'Alerts' -Facts $effectiveFacts)

    $mailTo = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'ErrorMailTo' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'ErrorMailTo' -DefaultValue '' }
    $from = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'From' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'From' -DefaultValue '' }
    if (-not [string]::IsNullOrWhiteSpace($mailTo) -and -not [string]::IsNullOrWhiteSpace($from)) {
        $smtpServer = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'SmtpServer' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'SmtpServer' -DefaultValue '' }
        $smtpPort = [int](if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'SmtpPort' -DefaultValue 25 } else { Get-SmartAzureConfigValue -Name 'SmtpPort' -DefaultValue 25 })
        $cc = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'Cc' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'Cc' -DefaultValue '' }
        $appId = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'AppId' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'AppId' -DefaultValue '' }
        $graphTenantId = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'TenantId' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'TenantId' -DefaultValue '' }
        $thumbprint = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'Thumbprint' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'Thumbprint' -DefaultValue '' }
        if (-not (Test-SmartAzureConfiguredValue -Value $thumbprint)) {
            $thumbprint = if ($null -ne $Config) { Get-SmartAzureScriptConfigValue -Config $Config -Name 'Thumb' -DefaultValue '' } else { Get-SmartAzureConfigValue -Name 'Thumb' -DefaultValue '' }
        }
        $body = @"
<html>
  <body style="font-family:Segoe UI,Arial,sans-serif; font-size:13px; color:#222;">
    <h2 style="margin:0 0 10px 0;">SmartAzure error: $ScriptName</h2>
    <p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p><strong>Tenant:</strong> $global:SmartAzureTenant</p>
    <p><strong>RunId:</strong> $RunId</p>
    <p><strong>Error:</strong> $message</p>
    <p><strong>Log:</strong> $LogPath</p>
  </body>
</html>
"@
        try {
            Send-SmartAzureMail -To $mailTo -From $from -Cc $cc -SmtpServer $smtpServer -SmtpPort $smtpPort -Subject ("SmartAzure error: {0}" -f $ScriptName) -BodyHtml $body -Attachments @($LogPath) -AppId $appId -TenantId $graphTenantId -Thumbprint $thumbprint
        }
        catch {
            Write-Host ("SmartAzure error mail failed: {0}" -f (Get-SmartAzureExceptionDetails -Exception $_.Exception))
        }
    }
}

function Initialize-SmartAzureTenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }

    $global:SmartAzureTenant = $Tenant
    $script:SmartAzureGlobalConfig = Get-SmartAzureEffectiveGlobalConfig -StartPath $StartPath -TenantKey $Tenant
    return $script:SmartAzureGlobalConfig
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBsOekj8QNqxsc
# drhf/T+iZz4x35GYKlermjryfAmIhqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAct1PiQsIOD+321s2bFZeO45uco0WYUq2WG4p5gQBetjANBgkqhkiG9w0B
# AQEFAASCAYBeIydFt44WoI00MMfRy0Et++RRMjm0DKHli7a4bIVKIadkLqP4YUsv
# AQAKs6JYyo/LR8p6sR87QRyIpNlWB2vWG5kMWDNKMTTRuC+JE6MGY2vS02Yo9xTL
# H+eC8ZXtVORkuYyAN8BLRQy+suMXqGGAJWQ71RELcVrSGKSgQYCQl1Ut5VP53zco
# tVrmRvWxwLv/utRkbFLBTEPQtWBdDrQxXEyulgfFz2BWPBl3KCYQ8wcXW+DylFps
# wmcGb2CQHmyTtCWPgXLC8ndh32qHtcbYlC51byH9QtNbWfGvIYIJDjZ+qU8Z6bgf
# SsmHqnEeDgJBAXtbcJGR8qa30FB6nK5YA9YgBI4nuUhscETg9gYB+28M4NhKiW0E
# QWvLJdJnv8PYHalX6JoN+d0naCKmFm6D1Q45WW92oINbJm7Cn+wcBSEWgJ/Bc/+C
# b3BVC7eUcbdPIDL8CwY7lll1+nvM2vIfNqCmMIL2njzqFtJDtVnuAra8/ZoSj9y2
# 1bB8Zrl7wHA=
# SIG # End signature block
