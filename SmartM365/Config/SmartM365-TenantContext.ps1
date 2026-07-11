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

    if (Test-Path -LiteralPath $Path) { return $false }

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = if ($Path -like '*Config\Tenants\*.local.json') {
            Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template'
        }
        elseif ($Path -like '*.local.json') {
            '{0}.template' -f $Path
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
        Copy-Item -LiteralPath $TemplatePath -Destination $Path -ErrorAction Stop
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

function Read-SmartM365JsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Required) {
            $templatePath = if ($Path -like '*Config\Tenants\*.local.json') { Join-Path -Path (Split-Path -Path $Path -Parent) -ChildPath 'tenant.local.json.template' } elseif ($Path -like '*.local.json') { '{0}.template' -f $Path } else { '' }
            Initialize-SmartM365LocalJsonFromTemplate -Path $Path -TemplatePath $templatePath -ConfigDescription 'required local configuration' | Out-Null
        }
        else {
            return [ordered]@{}
        }
    }

    try {
        return ConvertTo-SmartM365Hashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
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
        if ((Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json')) -or
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
        [string]$TenantKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($TenantKey)) { $TenantKey = 'test' }

    $rootPath = Find-SmartM365Root -StartPath $StartPath
    $scriptStartPath = if ([string]::IsNullOrWhiteSpace($StartPath)) { $rootPath } else { $StartPath }
    $scriptOutputRootPath = Join-Path -Path $scriptStartPath -ChildPath 'Output'
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'Config\SmartM365.global.local.json'
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $TenantKey)

    $globalConfig = Read-SmartM365JsonConfig -Path $globalConfigPath -Required
    $tenantConfig = Read-SmartM365JsonConfig -Path $tenantConfigPath -Required

    if ($tenantConfig.Contains('TenantKey') -and
        -not [string]::IsNullOrWhiteSpace([string]$tenantConfig['TenantKey']) -and
        [string]$tenantConfig['TenantKey'] -ne $TenantKey) {
        throw "Tenant profile key mismatch. File '$tenantConfigPath' contains TenantKey '$($tenantConfig['TenantKey'])' but requested '$TenantKey'."
    }

    foreach ($key in $tenantConfig.Keys) {
        $globalConfig[$key] = $tenantConfig[$key]
    }

    $globalConfig['TenantKey'] = $TenantKey
    $globalConfig['SmartM365RootPath'] = $rootPath
    $globalConfig['ScriptOutputRootPath'] = $scriptOutputRootPath

    $defaultWorkspaceRootPath = $rootPath
    $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
    $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
    $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
    $useScriptOutputFallback = $false

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or
        [string]$globalConfig['WorkspaceRootPath'] -eq '{{SmartM365RootPath}}') {
        $candidateDataRoot = Join-Path -Path $rootPath -ChildPath 'Data'
        if (Test-SmartM365WritableDirectory -Path $candidateDataRoot) {
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
        [string]$globalConfig['WorkspaceRootPath'] -eq '{{SmartM365RootPath}}') {
        $globalConfig['WorkspaceRootPath'] = $defaultWorkspaceRootPath
    }

    return [pscustomobject]$globalConfig
}

function Initialize-SmartM365TenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }

    $global:SmartM365Tenant = $Tenant
    $script:SmartM365GlobalConfig = Get-SmartM365EffectiveGlobalConfig -StartPath $StartPath -TenantKey $Tenant
    return $script:SmartM365GlobalConfig
}
