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

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB3PtPi26ROIPTX
# WqLNgz3qHOO24jz2H4wAQRnHfnUxq6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBBW18sgfkiANK0s7Vx
# qs2UH/XZC6RLzM4xUJB1qL2myzANBgkqhkiG9w0BAQEFAASCAYAO5tw8sPlZVeEh
# 59qyDVQOe3SmHDbRmCz2bY8K3nvprFluDr1oU5L71P8Ol9B4tn07s9OGIkC98chk
# /DuHZmlkEnp4N1OuCYK2UdMMjk2xSw5n9m+bpIlFOVMBr4CMrfSPHm5I5g8UTDBh
# IGz09GZl4avUI6J4X5x5Py1r17hE4DxPY3tWYvPieJObDC75qmHnQiIWuvUhV1ja
# Bll0OvgQZCb4k+FN4SYWYoP47CtE8ktxJNPzKvRur1lTj0tWs7GJqpQ+V1rMyjF7
# ftmfl0lmXlyF6BSsmLjuL9t1+E7PLKbESS4Gi5Xlp91+vTHn8Mxgk1qC/HicR6tQ
# eSBLQYoYrWTnQxCYrLxPg4d7Ygi7j22PYxgYbde87pkHo7lZsWn1NcRQCqmmRMJv
# hug2ZnKvAQTd0H4aIv8EHlFSGTOVyvuZ1E3XjszeuvATuUCuX3PJq1Th7e2ZLcIF
# Pf1kQk2EniR9bn8Z13X6pJYs5WiaSxr2cqFY9L33UBGYP2zM1hU=
# SIG # End signature block
