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

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD6HrDExyutjOJ9
# DkpVSQTXZqnTtHJ4hsXWGiic9R+iAaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBAEP/GPtr+kD8iztYF
# BdtKIj8O63lkQQqxkptejh6HyDANBgkqhkiG9w0BAQEFAASCAYByAfik9koglwWG
# z/sdJffxv11DsSXKqb9saSZ3eotbN6xDn8TenYo+buy2+mUPwBW1vNWzlIzWk8MM
# 7I06r+HODyvYceu/PDWo8VFQI0w+3fmRNXgCa8SjWYrRsRmY8NJq39wQtbcIecaB
# peUU/oLeDV4TNaByK8G240GkH8YoK2RVMTVJ1qEgBqmrLnmCYQy0kA+qf5pZWpr5
# xTYnyXRO9hiAVsim+3JjNTallUFoEug3gh+GO/2+DhnobIB3A5Vd4X35Tk+fNVvD
# OFvVaj6iZujinsDnxe8otk3GR7yg7lAGkS8rtCIrmk3SlSwpwNioh0bMW3Frnvp8
# 63XDo7jyxNBw+vGbsl1yNSRjRVmxXskfwqPFblI5OnGM0BG7b6/+/i2DARa9LvG4
# TsXOXLmVyahG8bl8wdRlaaBqVf19oZzcyEkp7P8SZx9+n5MDK9FTQ5lqq7YqA0yE
# 68t33Ubpb7s8CdkcmTcPA8LqcawPDg+lkYz1p5LDOTwbgH13A5k=
# SIG # End signature block
