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

function Get-SmartM365JsonTemplatePath {
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
        [string]$TemplatePath
    )

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Get-SmartM365JsonTemplatePath -Path $Path
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path -LiteralPath $TemplatePath)) {
        return $Config
    }

    try {
        $template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $addedKeys = @(Add-SmartM365MissingJsonTemplateProperties -Target $Config -Template $template)
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
function Read-SmartM365JsonConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

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
        $config = ConvertTo-SmartM365Hashtable -InputObject (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $Path)
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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhDfLexvcHKtod
# MxDlw6s7yLZW2ynnWbONbxptjVbAlaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBr7xI6cOXmEslAVXv8v8n9ovxpKNNDCfBSyW17J3SMfTANBgkqhkiG9w0B
# AQEFAASCAYCfz8DIkEtKv07B6lJGb8U0obC6vtqLokVEvo69I/vVQkJLoVHPVM/2
# FW6SjXOmImHIYuiXuKYH7wpVvTQHF+IBJfyjP8rgHqfICT1ukDRJAfVc3PdtdHWT
# Rjxct8PaZEkued5Jf2BhsJ6rNFBPgEZpcczySvvP7on5HkyypGsczlWX98RcayZ2
# Go5VIuWLlrYLwgjZ1uS4GIRQ21mB899/1VeRK+bOA7lqXDI2YSckJSrigO+y2Kot
# w16c+jqT/cB6i+2JW1UeRp52n8Pe1s45eZ4OPfuPbAtm5LVv6H1py6ad+ttkGA0X
# EMc1tr/co9pmk9+lUnJHal9DIcu1YP4Id8tJetVTfJi1KqqSbML4nmm0BuG20amq
# TDFsF+CaMYRm0/7ZSjndc7F1KjFtTYgAP8GCrfPYNj5zlNl2VO07hDFiCfDm0xLC
# qK2HfX7+ujLCpjS5kXAOFIvIWF5Pyc6A7xm4XLEBZrEW48rCcTLPA3iKA6aLp+ql
# 1BfK8dBZRNw=
# SIG # End signature block
