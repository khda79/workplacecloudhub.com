function ConvertTo-SmartAvdHashtable {
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    $hash = [ordered]@{}
    if ($null -eq $InputObject) { return $hash }
    foreach ($property in $InputObject.PSObject.Properties) {
        $hash[$property.Name] = $property.Value
    }
    return $hash
}

function Get-SmartAvdJsonTemplatePath {
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

function Add-SmartAvdMissingJsonTemplateProperties {
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
            foreach ($nestedPath in (Add-SmartAvdMissingJsonTemplateProperties -Target $currentValue -Template $property.Value -PrefixPath $propertyPath)) {
                $added.Add($nestedPath) | Out-Null
            }
        }
    }

    return @($added)
}

function Sync-SmartAvdJsonConfigWithTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartAvdJsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $addedKeys = @(Add-SmartAvdMissingJsonTemplateProperties -Target $Config -Template $template)
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
function Read-SmartAvdJsonConfig {
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
        $config = ConvertTo-SmartAvdHashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        return (Sync-SmartAvdJsonConfigWithTemplate -Config $config -Path $Path)
    }
    catch {
        throw ("Failed to read configuration file '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Find-SmartAvdRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    $searchRoot = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).Path } else { $StartPath }
    $resolvedSearchRoot = Resolve-Path -LiteralPath $searchRoot -ErrorAction SilentlyContinue
    if ($null -ne $resolvedSearchRoot) { $searchRoot = $resolvedSearchRoot.Path }

    while ($searchRoot) {
        if ((Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartAzureVirtualDesktop.global.local.json')) -or
            (Test-Path -LiteralPath (Join-Path -Path $searchRoot -ChildPath 'SmartAzureVirtualDesktop.global.local.json.template'))) {
            return $searchRoot
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw "SmartAzureVirtualDesktop root not found from '$StartPath'."
}

function Import-SmartAvdCoreModule {
    [CmdletBinding()]
    param([string]$StartPath)

    if (Get-Command -Name Set-SmartAvdCoreContext -ErrorAction SilentlyContinue) { return }

    $rootPath = Find-SmartAvdRoot -StartPath $StartPath
    $modulePath = Join-Path -Path $rootPath -ChildPath 'Modules\SmartAzureVirtualDesktop.Core\SmartAzureVirtualDesktop.Core.psd1'
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "SmartAzureVirtualDesktop.Core module not found: $modulePath"
    }

    Import-Module $modulePath -Global -Force -ErrorAction Stop
}

function Test-SmartAvdWritableDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        $probePath = Join-Path -Path $Path -ChildPath ('.smartavd-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $probePath -Value 'test' -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Get-SmartAvdEffectiveGlobalConfig {
    [CmdletBinding()]
    param(
        [string]$StartPath,
        [string]$TenantKey = 'test'
    )

    if ([string]::IsNullOrWhiteSpace($TenantKey)) { $TenantKey = 'test' }

    $rootPath = Find-SmartAvdRoot -StartPath $StartPath
    $scriptStartPath = if ([string]::IsNullOrWhiteSpace($StartPath)) { $rootPath } else { $StartPath }
    $scriptOutputRootPath = Join-Path -Path $scriptStartPath -ChildPath 'Output'
    $globalConfigPath = Join-Path -Path $rootPath -ChildPath 'SmartAzureVirtualDesktop.global.local.json'
    $tenantConfigPath = Join-Path -Path $rootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $TenantKey)

    $globalConfig = Read-SmartAvdJsonConfig -Path $globalConfigPath
    $tenantConfig = Read-SmartAvdJsonConfig -Path $tenantConfigPath

    if ($tenantConfig.Contains('TenantKey') -and
        -not [string]::IsNullOrWhiteSpace([string]$tenantConfig['TenantKey']) -and
        [string]$tenantConfig['TenantKey'] -ne $TenantKey) {
        throw "Tenant profile key mismatch. File '$tenantConfigPath' contains TenantKey '$($tenantConfig['TenantKey'])' but requested '$TenantKey'."
    }

    foreach ($key in $tenantConfig.Keys) { $globalConfig[$key] = $tenantConfig[$key] }

    $globalConfig['TenantKey'] = $TenantKey
    $globalConfig['SmartAzureVirtualDesktopRootPath'] = $rootPath
    $globalConfig['ScriptOutputRootPath'] = $scriptOutputRootPath

    $defaultWorkspaceRootPath = $rootPath
    $defaultDataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
    $defaultLatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
    $defaultLogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
    $useScriptOutputFallback = $false

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or
        [string]$globalConfig['WorkspaceRootPath'] -in @('.', '{{SmartAzureVirtualDesktopRootPath}}')) {
        $candidateDataRoot = Join-Path -Path $rootPath -ChildPath 'Data'
        if (Test-SmartAvdWritableDirectory -Path $candidateDataRoot) {
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

    if (-not $globalConfig.Contains('WorkspaceRootPath') -or
        [string]::IsNullOrWhiteSpace([string]$globalConfig['WorkspaceRootPath']) -or
        [string]$globalConfig['WorkspaceRootPath'] -in @('.', '{{SmartAzureVirtualDesktopRootPath}}')) {
        $globalConfig['WorkspaceRootPath'] = $defaultWorkspaceRootPath
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

    return [pscustomobject]$globalConfig
}

function Resolve-SmartAvdConfigTokenValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($null -eq $script:SmartAvdGlobalConfig) { throw 'SmartAzureVirtualDesktop tenant context has not been initialized.' }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartAvdGlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartAvdConfigTokenValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}

function Resolve-SmartAvdOutputRoots {
    [CmdletBinding()]
    param(
        [string]$OutputRoot,
        [string]$LatestOutputRoot,
        [Parameter(Mandatory)][string]$AreaPath
    )

    if ($null -eq $script:SmartAvdGlobalConfig) { throw 'SmartAzureVirtualDesktop tenant context has not been initialized.' }

    $resolvedOutputRoot = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        Join-Path -Path (Resolve-SmartAvdConfigTokenValue -Value '{{DataAllRootPath}}') -ChildPath $AreaPath
    }
    else {
        Resolve-SmartAvdConfigTokenValue -Value $OutputRoot
    }

    $resolvedLatestOutputRoot = if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) {
        Resolve-SmartAvdConfigTokenValue -Value '{{LatestCsvFolderPath}}'
    }
    else {
        Resolve-SmartAvdConfigTokenValue -Value $LatestOutputRoot
    }

    [pscustomobject]@{
        OutputRoot       = $resolvedOutputRoot
        LatestOutputRoot = $resolvedLatestOutputRoot
    }
}

function Get-SmartAvdConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $script:SmartAvdGlobalConfig) { throw 'SmartAzureVirtualDesktop tenant context has not been initialized.' }

    $property = $script:SmartAvdGlobalConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }

    if ($property.Value -is [string]) {
        $value = $property.Value.Trim()
        if ([string]::IsNullOrWhiteSpace($value) -or $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $DefaultValue }
    }

    return Resolve-SmartAvdConfigTokenValue -Value $property.Value
}

function Get-SmartAvdScriptLocalConfig {
    [CmdletBinding()]
    param([string]$ScriptPath)

    $effectiveScriptPath = if ([string]::IsNullOrWhiteSpace($ScriptPath)) { $PSCommandPath } else { $ScriptPath }
    if ([string]::IsNullOrWhiteSpace($effectiveScriptPath)) { return [pscustomobject]@{} }

    $scriptFolder = Split-Path -Path $effectiveScriptPath -Parent
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($effectiveScriptPath)
    $configPath = Join-Path -Path $scriptFolder -ChildPath ("{0}.local.json" -f $scriptBaseName)

    if (-not (Test-Path -LiteralPath $configPath)) { return [pscustomobject]@{} }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return (Sync-SmartAvdJsonConfigWithTemplate -Config $config -Path $configPath)
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Get-SmartAvdScriptConfigValue {
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
                return Resolve-SmartAvdConfigTokenValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartAvdConfigTokenValue -Value $property.Value
        }
    }

    return Get-SmartAvdConfigValue -Name $Name -DefaultValue $DefaultValue
}

function Initialize-SmartAvdTenantContext {
    [CmdletBinding()]
    param(
        [string]$Tenant = 'test',
        [string]$StartPath
    )

    if ([string]::IsNullOrWhiteSpace($Tenant)) { $Tenant = 'test' }

    $global:SmartAzureVirtualDesktopTenant = $Tenant
    $script:SmartAvdGlobalConfig = Get-SmartAvdEffectiveGlobalConfig -StartPath $StartPath -TenantKey $Tenant
    return $script:SmartAvdGlobalConfig
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBV79/Gg4yhDBr1
# pipjIADifM9D2tAtj8QYaWQOq74Dr6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAsNQTcZQA9IdQGTnbzAtcEI7pWfZ18Ne4rx1IuDEMuxzANBgkqhkiG9w0B
# AQEFAASCAYBuHj2aNOeF3Dsrgh7lqaBPWClq0W+bC0KqT00HVJNa6XcchObKfO5O
# Ig+p3q9N/T5i18o6xFfYP1LNeZ8+CXHv5FxrK6+6fRO1hu1s5Xy3WPvjtBzJiCyJ
# 9eRKoo3oUacJkCFslusE9WXbe/tVBWFo3JL0y4OsUb0fHCn4kub7s/MbUOG7J32r
# Nxe4dNuxjAv8MWzn1QlS/m685kE7BOl7RSiBblOA/SrX+c/Zi3d+BHOhKZ42yNbl
# elqI9UtPtGVBI+c839uLb1VOaopYxdsVsAZyACEp31c1K6IFzAKm+BPkUb01j0Ra
# QaJlFZZrp9X/IVKn1m8TnNrvE8UkmNQxG4hDDCTVnbIvtaJgjpNP+Y2XKNDEB2kn
# pk/tq7PAp3bCStC4tOzC3AlnsXrhDg/UOlI9T10SwBHHUdfFaTBsupRHT0khxRUi
# ksfkWRsxC0b7yeiHY8mhAi5FnV5Gdgbis4ySqrtTYJzGsvXiZsuIwcpDQU6LLBX+
# gFcjs64r/lg=
# SIG # End signature block
