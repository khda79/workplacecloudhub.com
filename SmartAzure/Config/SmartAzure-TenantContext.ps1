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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB0tff3pX81LEiG
# r6fY1a5gBDmQ0KtBm3dH3naBY2Xz8qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEICFS6x7j7YM7CimN+tPLEC2gc12ergoJ/v4YTwwGfpcHMA0GCSqG
# SIb3DQEBAQUABIIBgKmVnpFCf3v1E3exKyTR0uMm48P+MyGWjZ920dgjSw/G5Y3B
# MphSImlwr439dolKWEPHlyv/xEj4EeohxRqNhvhMfSTH3Ek2v3QLoSq7iQ740BB2
# 2opykvDE2YfFjkaOv9PrvI8rEs0JKXttzO4aoqHtKE2XCy/5NddSJ5B/jtw3EkfA
# b4cAIx4OMwAGpMgwIrEYRA22RbETN+zzxH6D3pOnU/RZkgW9QDlSc2Nqvz/7GJcR
# K9sYsvYGN0aoJp7DcGknqKWZKx9J/5EJ6ESzWFpC2PYFmuEa1CiHmvfNXqTdWQ+V
# T6BXKsiYflLEDY7OJ5ISPrzeqMON4ZF4iSqWBMZ0V9s6mf295YN5ZvHSZwut+Hj6
# 2GDlpT19prSIxsGrdWQnV50FQ5wHIbw66zkGcMeOems92O2naNZhZewjou4NeHxm
# U34iwOwBYBTcOa4+WxauzfT8rdZ9WF/G0Atoc0ldYk+fxrfZuyflbnNqMrob4JHn
# PD9Pt5CS+WcXaXxZHKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MDZaMC8GCSqGSIb3DQEJBDEiBCBNgrSYlURysFR/neJdCh5Ktnrwx6/C29nhJqcw
# sC4WbTANBgkqhkiG9w0BAQEFAASCAgARMb+utYBUr2+c5etVcUdbm+f7rrpYjSdI
# zbDGa1/DONZqKWmSNm8ncwqDNS2HwQVsIQ4GjqtLW05m9650ZqpdWw3jzqHemcJZ
# qCLfOAcgFR7SEZjZP1L+qUK8hXvWPD/kEw2J6p4WPR8BGBU+2Z/sDAIFRL1mec4o
# vyhYHh8myLlcBolgKhzGtUn+ABdiV+5Chx0SFHB7Oq2XcowzyapCLgx8ZjYz0jsz
# yqAbGqVKdwkcvok884YgcuRIXGiqS27JZSHvNhV3wLgG0+WAwaFD81OdtPNcSD9K
# LNK2HEce9X2jnVwFlnBGKktI6JLccj00Q1on9V+61r2tYQG55s1TYQzBeX1g9Ljc
# YUv4M1NsAMa22NSRsOxOxRyfND/fJA6ZGng0jDP4AJtEV4/wteiPhTfAjv0ZTGAv
# 1YZsR+CR3LC919ArVQxsB9IjL3N2vlEXf9rPEmqBA8Xcm5emwII3O0hwlQvWqmB6
# GUEzQkPgdkZjx8jInMvsx/AXGatab/StOwMqhM6pc4cXePKXEVTYera9F0HrAQDO
# gypsapLL6kbGcG9vLEjeqwpnh5UnFnrzWLdSXwBc/NdQIcL1+PZOsOWSXG2FR/B0
# 6lXnjOCU9sqXZhcPBKPCgBKlBzOWFXifKvSiIZD6eeF0xnV8bVlXd89kgdACdSDI
# 6LtGcSvubA==
# SIG # End signature block
