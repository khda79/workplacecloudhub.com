# SmartWorkplaceCMDB.Core
# Version: 0.2.1

$script:SmartWorkplaceCMDBCoreVersion = '0.2.1'

function Get-SmartWorkplaceCMDBProjectRoot {
    [CmdletBinding()]
    param()

    $moduleRoot = Split-Path -Parent $PSScriptRoot
    return (Split-Path -Parent $moduleRoot)
}

function Read-SmartWorkplaceCMDBJsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "JSON file not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "JSON file is empty: $Path"
    }

    return ($raw | ConvertFrom-Json -ErrorAction Stop)
}

function ConvertTo-SmartWorkplaceCMDBHashtable {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $result[[string]$key] = ConvertTo-SmartWorkplaceCMDBHashtable -InputObject $InputObject[$key]
        }
        return $result
    }

    if ($InputObject -is [pscustomobject]) {
        $result = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $result[$property.Name] = ConvertTo-SmartWorkplaceCMDBHashtable -InputObject $property.Value
        }
        return $result
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object {
            ConvertTo-SmartWorkplaceCMDBHashtable -InputObject $_
        })
    }

    return $InputObject
}

function Merge-SmartWorkplaceCMDBHashtable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Base,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Overlay
    )

    $result = [ordered]@{}
    foreach ($key in $Base.Keys) {
        $result[[string]$key] = ConvertTo-SmartWorkplaceCMDBHashtable -InputObject $Base[$key]
    }

    foreach ($key in $Overlay.Keys) {
        $keyName = [string]$key
        if ($result.Contains($keyName) -and
            $result[$keyName] -is [System.Collections.IDictionary] -and
            $Overlay[$key] -is [System.Collections.IDictionary]) {
            $result[$keyName] = Merge-SmartWorkplaceCMDBHashtable -Base $result[$keyName] -Overlay $Overlay[$key]
        }
        else {
            $result[$keyName] = ConvertTo-SmartWorkplaceCMDBHashtable -InputObject $Overlay[$key]
        }
    }

    return $result
}

function Write-SmartWorkplaceCMDBJsonAtomically {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $tempPath = '{0}.tmp.{1}' -f $Path, ([guid]::NewGuid().ToString('N'))
    try {
        $InputObject | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $tempPath -Encoding UTF8 -Force
        Move-Item -LiteralPath $tempPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SmartWorkplaceCMDBConfigDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TemplatePath,

        [Parameter(Mandatory)]
        [string]$RuntimePath,

        [switch]$NoWrite
    )

    $template = ConvertTo-SmartWorkplaceCMDBHashtable -InputObject (Read-SmartWorkplaceCMDBJsonFile -Path $TemplatePath)
    $runtimeExists = Test-Path -LiteralPath $RuntimePath -PathType Leaf
    $runtime = if ($runtimeExists) {
        ConvertTo-SmartWorkplaceCMDBHashtable -InputObject (Read-SmartWorkplaceCMDBJsonFile -Path $RuntimePath)
    }
    else {
        [ordered]@{}
    }

    $effective = Merge-SmartWorkplaceCMDBHashtable -Base $template -Overlay $runtime
    $effectiveJson = $effective | ConvertTo-Json -Depth 100 -Compress
    $runtimeJson = if ($runtimeExists) {
        $runtime | ConvertTo-Json -Depth 100 -Compress
    }
    else {
        ''
    }

    if (-not $NoWrite -and (-not $runtimeExists -or $runtimeJson -cne $effectiveJson)) {
        Write-SmartWorkplaceCMDBJsonAtomically -InputObject $effective -Path $RuntimePath
        $action = if ($runtimeExists) { 'synchronized' } else { 'created' }
        Write-Information ("SmartWorkplaceCMDB runtime configuration {0}: {1}" -f $action, $RuntimePath) -InformationAction Continue
    }

    return $effective
}

function ConvertTo-SmartWorkplaceCMDBKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [ValidateSet('ProfileKey', 'OrganizationKey', 'EnvironmentKey')]
        [string]$Name
    )

    $normalized = $Value.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        throw "$Name is required."
    }
    if ($normalized.Length -gt 64) {
        throw "$Name must not exceed 64 characters."
    }
    if ($normalized -notmatch '^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$') {
        throw "$Name '$Value' is invalid. Use lowercase letters, digits, and internal hyphens only."
    }

    return $normalized
}

function Expand-SmartWorkplaceCMDBPathToken {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Tokens
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $expanded = $Value
    foreach ($key in $Tokens.Keys) {
        $token = '{{' + [string]$key + '}}'
        $expanded = $expanded.Replace($token, [string]$Tokens[$key])
    }
    if ($expanded -match '\{\{[^}]+\}\}') {
        throw "Unresolved path token in '$Value'."
    }

    return [System.IO.Path]::GetFullPath($expanded)
}

function Get-SmartWorkplaceCMDBConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Primary,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Fallback,

        [Parameter(Mandatory)]
        [string]$Name,

        [string]$Placeholder
    )

    $value = if ($Primary.Contains($Name)) { [string]$Primary[$Name] } else { '' }
    if ([string]::IsNullOrWhiteSpace($value) -or
        (-not [string]::IsNullOrWhiteSpace($Placeholder) -and $value -eq $Placeholder)) {
        $value = if ($Fallback.Contains($Name)) { [string]$Fallback[$Name] } else { '' }
    }
    return $value
}

function Resolve-SmartWorkplaceCMDBContext {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary]$BoundParameters = @{},
        [string]$GlobalConfigPath,
        [string]$TenantConfigPath,
        [switch]$NoConfigWrite
    )

    $projectRoot = Get-SmartWorkplaceCMDBProjectRoot
    $globalTemplatePath = Join-Path -Path $projectRoot -ChildPath 'Config\SmartWorkplaceCMDB.global.local.json.template'
    $tenantTemplatePath = Join-Path -Path $projectRoot -ChildPath 'Config\Tenants\tenant.local.json.template'

    if ([string]::IsNullOrWhiteSpace($GlobalConfigPath)) {
        $GlobalConfigPath = Join-Path -Path $projectRoot -ChildPath 'Config\SmartWorkplaceCMDB.global.local.json'
    }

    $globalConfig = Get-SmartWorkplaceCMDBConfigDocument -TemplatePath $globalTemplatePath -RuntimePath $GlobalConfigPath -NoWrite:$NoConfigWrite
    $requestedProfile = if ($BoundParameters.Contains('Tenant')) {
        [string]$BoundParameters['Tenant']
    }
    else {
        [string]$globalConfig['ProfileKey']
    }
    if ([string]::IsNullOrWhiteSpace($requestedProfile)) {
        $requestedProfile = 'default'
    }
    $profileKey = ConvertTo-SmartWorkplaceCMDBKey -Value $requestedProfile -Name 'ProfileKey'

    if ([string]::IsNullOrWhiteSpace($TenantConfigPath)) {
        $TenantConfigPath = Join-Path -Path $projectRoot -ChildPath ("Config\Tenants\{0}.local.json" -f $profileKey)
    }

    $tenantConfig = Get-SmartWorkplaceCMDBConfigDocument -TemplatePath $tenantTemplatePath -RuntimePath $TenantConfigPath -NoWrite:$NoConfigWrite
    $tenantProfile = [string]$tenantConfig['ProfileKey']
    if ($tenantProfile -ne $profileKey) {
        if ([string]::IsNullOrWhiteSpace($tenantProfile) -or $tenantProfile -eq 'default') {
            $tenantConfig['ProfileKey'] = $profileKey
            if (-not $NoConfigWrite) {
                Write-SmartWorkplaceCMDBJsonAtomically -InputObject $tenantConfig -Path $TenantConfigPath
            }
        }
        else {
            throw "Tenant configuration ProfileKey '$tenantProfile' does not match requested profile '$profileKey'."
        }
    }

    $effectiveConfig = Merge-SmartWorkplaceCMDBHashtable -Base $globalConfig -Overlay $tenantConfig
    $effectiveConfig['ProfileKey'] = $profileKey

    $organizationValue = if ($BoundParameters.Contains('OrganizationKey')) {
        [string]$BoundParameters['OrganizationKey']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantConfig -Fallback $globalConfig -Name 'OrganizationKey' -Placeholder 'organization'
    }
    $environmentValue = if ($BoundParameters.Contains('EnvironmentKey')) {
        [string]$BoundParameters['EnvironmentKey']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantConfig -Fallback $globalConfig -Name 'EnvironmentKey' -Placeholder 'default'
    }

    $organizationKey = ConvertTo-SmartWorkplaceCMDBKey -Value $organizationValue -Name 'OrganizationKey'
    $environmentKey = ConvertTo-SmartWorkplaceCMDBKey -Value $environmentValue -Name 'EnvironmentKey'
    $expectedTenantKey = '{0}-{1}' -f $organizationKey, $environmentKey

    $configuredTenantKey = if ($BoundParameters.Contains('TenantKey')) {
        [string]$BoundParameters['TenantKey']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantConfig -Fallback $globalConfig -Name 'TenantKey' -Placeholder 'organization-default'
    }
    if ([string]::IsNullOrWhiteSpace($configuredTenantKey)) {
        $configuredTenantKey = $expectedTenantKey
    }

    $globalGraph = if ($globalConfig.Contains('MicrosoftGraph') -and
        $globalConfig['MicrosoftGraph'] -is [System.Collections.IDictionary]) {
        $globalConfig['MicrosoftGraph']
    }
    else {
        [ordered]@{}
    }
    $tenantGraph = if ($tenantConfig.Contains('MicrosoftGraph') -and
        $tenantConfig['MicrosoftGraph'] -is [System.Collections.IDictionary]) {
        $tenantConfig['MicrosoftGraph']
    }
    else {
        [ordered]@{}
    }
    $tenantId = if ($BoundParameters.Contains('TenantId')) {
        [string]$BoundParameters['TenantId']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantGraph -Fallback $globalGraph -Name 'TenantId'
    }

    $globalOutput = if ($globalConfig.Contains('Output') -and
        $globalConfig['Output'] -is [System.Collections.IDictionary]) {
        $globalConfig['Output']
    }
    else {
        [ordered]@{}
    }
    $tenantOutput = if ($tenantConfig.Contains('Output') -and
        $tenantConfig['Output'] -is [System.Collections.IDictionary]) {
        $tenantConfig['Output']
    }
    else {
        [ordered]@{}
    }

    $dataRootValue = if ($BoundParameters.Contains('DataRootPath')) {
        [string]$BoundParameters['DataRootPath']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantOutput -Fallback $globalOutput -Name 'DataRootPath'
    }
    $baseTokens = [ordered]@{
        ProjectRootPath = $projectRoot
        ProfileKey      = $profileKey
    }
    $dataRootPath = Expand-SmartWorkplaceCMDBPathToken -Value $dataRootValue -Tokens $baseTokens

    $pathTokens = [ordered]@{
        ProjectRootPath = $projectRoot
        ProfileKey      = $profileKey
        DataRootPath    = $dataRootPath
    }
    $dataAllValue = if ($BoundParameters.Contains('DataAllRootPath')) {
        [string]$BoundParameters['DataAllRootPath']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantOutput -Fallback $globalOutput -Name 'DataAllRootPath'
    }
    $latestValue = if ($BoundParameters.Contains('LatestOutputRootPath')) {
        [string]$BoundParameters['LatestOutputRootPath']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantOutput -Fallback $globalOutput -Name 'LatestOutputRootPath'
    }
    $logValue = if ($BoundParameters.Contains('LogRootPath')) {
        [string]$BoundParameters['LogRootPath']
    }
    else {
        Get-SmartWorkplaceCMDBConfigValue -Primary $tenantOutput -Fallback $globalOutput -Name 'LogRootPath'
    }

    $dataAllRootPath = Expand-SmartWorkplaceCMDBPathToken -Value $dataAllValue -Tokens $pathTokens
    $latestOutputRootPath = Expand-SmartWorkplaceCMDBPathToken -Value $latestValue -Tokens $pathTokens
    $logRootPath = Expand-SmartWorkplaceCMDBPathToken -Value $logValue -Tokens $pathTokens

    $paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant $profileKey -OrganizationKey $organizationKey -EnvironmentKey $environmentKey -TenantKey $configuredTenantKey -TenantId $tenantId -DataRootPath $dataRootPath -DataAllRootPath $dataAllRootPath -LatestOutputRootPath $latestOutputRootPath -LogRootPath $logRootPath

    return [pscustomobject]@{
        Configuration    = $effectiveConfig
        Paths            = $paths
        GlobalConfigPath = [System.IO.Path]::GetFullPath($GlobalConfigPath)
        TenantConfigPath = [System.IO.Path]::GetFullPath($TenantConfigPath)
        ContractPath     = Join-Path -Path $projectRoot -ChildPath 'Schema\SmartWorkplaceCMDB.tables.json'
    }
}

function Resolve-SmartWorkplaceCMDBTenantPath {
    [CmdletBinding()]
    param(
        [Alias('ProfileKey')]
        [string]$Tenant = 'default',
        [string]$OrganizationKey = 'organization',
        [string]$EnvironmentKey = 'default',
        [string]$TenantKey,
        [string]$TenantId,
        [string]$DataRootPath,
        [string]$DataAllRootPath,
        [string]$LatestOutputRootPath,
        [string]$LogRootPath
    )

    $projectRoot = Get-SmartWorkplaceCMDBProjectRoot
    $profileKey = ConvertTo-SmartWorkplaceCMDBKey -Value $Tenant -Name 'ProfileKey'
    $organizationKeyValue = ConvertTo-SmartWorkplaceCMDBKey -Value $OrganizationKey -Name 'OrganizationKey'
    $environmentKeyValue = ConvertTo-SmartWorkplaceCMDBKey -Value $EnvironmentKey -Name 'EnvironmentKey'

    $expectedTenantKey = '{0}-{1}' -f $organizationKeyValue, $environmentKeyValue
    if ([string]::IsNullOrWhiteSpace($TenantKey)) {
        $TenantKey = $expectedTenantKey
    }
    elseif ($TenantKey.Trim().ToLowerInvariant() -ne $expectedTenantKey) {
        throw "TenantKey '$TenantKey' must equal OrganizationKey-EnvironmentKey ('$expectedTenantKey')."
    }
    else {
        $TenantKey = $TenantKey.Trim().ToLowerInvariant()
    }

    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $parsedTenantId = [guid]::Empty
        if (-not [guid]::TryParse($TenantId, [ref]$parsedTenantId)) {
            throw "TenantId '$TenantId' is invalid. Supply an Entra tenant GUID or leave it empty."
        }
        $TenantId = $parsedTenantId.ToString()
    }

    if ([string]::IsNullOrWhiteSpace($DataRootPath)) {
        $DataRootPath = Join-Path -Path $projectRoot -ChildPath ("Data\Tenants\{0}" -f $profileKey)
    }
    if ([string]::IsNullOrWhiteSpace($DataAllRootPath)) {
        $DataAllRootPath = Join-Path -Path $DataRootPath -ChildPath 'DATA-ALL'
    }
    if ([string]::IsNullOrWhiteSpace($LatestOutputRootPath)) {
        $LatestOutputRootPath = Join-Path -Path $DataRootPath -ChildPath 'DATA-LAST'
    }
    if ([string]::IsNullOrWhiteSpace($LogRootPath)) {
        $LogRootPath = Join-Path -Path $DataRootPath -ChildPath 'LOG-ALL'
    }

    $DataRootPath = [System.IO.Path]::GetFullPath($DataRootPath)
    $DataAllRootPath = [System.IO.Path]::GetFullPath($DataAllRootPath)
    $LatestOutputRootPath = [System.IO.Path]::GetFullPath($LatestOutputRootPath)
    $LogRootPath = [System.IO.Path]::GetFullPath($LogRootPath)

    return [pscustomobject]@{
        ProfileKey           = $profileKey
        OrganizationKey      = $organizationKeyValue
        EnvironmentKey       = $environmentKeyValue
        TenantKey            = $TenantKey
        TenantId             = $TenantId
        ProjectRootPath      = $projectRoot
        DataRootPath         = $DataRootPath
        DataAllRootPath      = $DataAllRootPath
        LatestOutputRootPath = $LatestOutputRootPath
        LogRootPath          = $LogRootPath
        CmdbLatestPath       = Join-Path -Path $LatestOutputRootPath -ChildPath 'CMDB'
        PowerBILatestPath    = Join-Path -Path $LatestOutputRootPath -ChildPath 'PowerBI'
    }
}

function Initialize-SmartWorkplaceCMDBTenantFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Paths
    )

    foreach ($path in @(
            $Paths.DataRootPath,
            $Paths.DataAllRootPath,
            $Paths.LatestOutputRootPath,
            $Paths.LogRootPath,
            $Paths.CmdbLatestPath,
            $Paths.PowerBILatestPath
        )) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Export-SmartWorkplaceCMDBCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path,

        [string[]]$Columns,
        [string]$TenantKey,
        [string]$OrganizationKey,
        [string]$EnvironmentKey,
        [string]$TenantId
    )

    $identityColumns = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
    if (-not [string]::IsNullOrWhiteSpace($TenantKey)) {
        if ([string]::IsNullOrWhiteSpace($OrganizationKey) -or [string]::IsNullOrWhiteSpace($EnvironmentKey)) {
            throw 'OrganizationKey and EnvironmentKey are required when TenantKey is supplied.'
        }

        if ($InputObject.Count -gt 0) {
            $InputObject = @($InputObject | ForEach-Object {
                $properties = [ordered]@{
                    TenantKey       = $TenantKey
                    OrganizationKey = $OrganizationKey
                    EnvironmentKey  = $EnvironmentKey
                    TenantId        = $TenantId
                }
                foreach ($property in $_.PSObject.Properties) {
                    if ($property.Name -notin $identityColumns) {
                        $properties[$property.Name] = $property.Value
                    }
                }
                [pscustomobject]$properties
            })
        }

        if ($Columns -and $Columns.Count -gt 0) {
            $Columns = @($identityColumns + @($Columns | Where-Object { $_ -notin $identityColumns }))
        }
    }

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $tempPath = '{0}.tmp.{1}.csv' -f $Path, ([guid]::NewGuid().ToString('N'))
    try {
        if ($InputObject.Count -gt 0) {
            $InputObject | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8 -Force
        }
        elseif ($Columns -and $Columns.Count -gt 0) {
            ($Columns -join ',') | Set-Content -LiteralPath $tempPath -Encoding UTF8 -Force
        }
        else {
            '' | Set-Content -LiteralPath $tempPath -Encoding UTF8 -Force
        }

        Move-Item -LiteralPath $tempPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SmartWorkplaceCMDBTableContract {
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Join-Path -Path (Get-SmartWorkplaceCMDBProjectRoot) -ChildPath 'Schema\SmartWorkplaceCMDB.tables.json'
    }

    $contract = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    if ([string]::IsNullOrWhiteSpace([string]$contract.contractVersion)) {
        throw "CSV contract version is missing: $Path"
    }
    if (@($contract.tables).Count -eq 0) {
        throw "CSV contract does not define tables: $Path"
    }

    return $contract
}

function Test-SmartWorkplaceCMDBCsvContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LatestOutputRootPath,

        [string]$ContractPath,

        [switch]$ThrowOnError
    )

    $contract = Get-SmartWorkplaceCMDBTableContract -Path $ContractPath
    $results = @()

    foreach ($table in @($contract.tables)) {
        $folder = Join-Path -Path $LatestOutputRootPath -ChildPath ([string]$table.area)
        $path = Join-Path -Path $folder -ChildPath ([string]$table.name)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $results += [pscustomobject]@{
                Area              = [string]$table.area
                Name              = [string]$table.name
                Path              = $path
                Status            = 'Missing'
                MissingColumns    = ''
                UnexpectedColumns = ''
            }
            continue
        }

        $headerLine = Get-Content -LiteralPath $path -TotalCount 1 -ErrorAction Stop
        $actualColumns = if ([string]::IsNullOrWhiteSpace($headerLine)) {
            @()
        }
        else {
            @($headerLine.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
        }
        $expectedColumns = @($table.columns | ForEach-Object { [string]$_ })
        $missingColumns = @($expectedColumns | Where-Object { $_ -notin $actualColumns })
        $unexpectedColumns = @($actualColumns | Where-Object { $_ -notin $expectedColumns })
        $exactOrder = (($actualColumns -join [char]31) -ceq ($expectedColumns -join [char]31))
        $status = if ($missingColumns.Count -eq 0 -and $unexpectedColumns.Count -eq 0 -and $exactOrder) {
            'Valid'
        }
        else {
            'Incompatible'
        }

        $results += [pscustomobject]@{
            Area              = [string]$table.area
            Name              = [string]$table.name
            Path              = $path
            Status            = $status
            MissingColumns    = ($missingColumns -join ', ')
            UnexpectedColumns = ($unexpectedColumns -join ', ')
        }
    }

    $incompatible = @($results | Where-Object Status -eq 'Incompatible')
    if ($ThrowOnError -and $incompatible.Count -gt 0) {
        $details = @($incompatible | ForEach-Object {
            '{0}: missing=[{1}] unexpected=[{2}]' -f $_.Name, $_.MissingColumns, $_.UnexpectedColumns
        })
        throw ("Incompatible SmartWorkplaceCMDB CSV contract detected. Use a reviewed migration or -ForceInitialize for schema-only outputs. {0}" -f ($details -join '; '))
    }

    return @($results)
}

Export-ModuleMember -Function @(
    'Get-SmartWorkplaceCMDBProjectRoot',
    'Read-SmartWorkplaceCMDBJsonFile',
    'ConvertTo-SmartWorkplaceCMDBKey',
    'Resolve-SmartWorkplaceCMDBContext',
    'Resolve-SmartWorkplaceCMDBTenantPath',
    'Initialize-SmartWorkplaceCMDBTenantFolder',
    'Export-SmartWorkplaceCMDBCsv',
    'Get-SmartWorkplaceCMDBTableContract',
    'Test-SmartWorkplaceCMDBCsvContract'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCYFRBa/vijF/ai
# 66P942o+RFoFbWvEKFLO1dxFVyAeTaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMwKMbJAW7aWzsvMF/al9kN9XXWJ7hFTvf4UKe7rRUZdMA0GCSqG
# SIb3DQEBAQUABIIBgJHxO4oRJSGc8DB6Sdgi0ByPlCfyg1g7xXBJcZnb4SMuPkbs
# Q9hK6x1WTfBPbZz202W4fOiP2Rq8NC2NVpHyIpTboI/ZJNfhBlnIKly2SQCz7Rn7
# s/Xz9tKpxOBgLokMRY03CAOpM2q/FpSSxvVOjesLuX2r9RrkhQG6cDYlBD4u2Pd9
# YaaTpNi+EinL3O7E9y66Tob9Zx1fP3BzbBr0iDb4WivRjWNki3qFUBbFvNPyehqd
# EuF8iPmvTjtvbwNtC+BiTJECNwIQYbfrDn9/7sK+OGLoxJCI/hZVZXmkXYO3Vvnw
# FscP76CvRu5bHZqBiXi6baxX1ViAXjuIC51Wx/tjrlddA9Ijhv9C7OJcaHfAQlup
# UQ/fJmhJ8mhU8K98G8FmEXg0DuEG/hpdiN9brKfwibTBHYFArIT2lJFfMwgvVf5/
# sS02e2FbL6WUBx2v3fAVdp32nva3ZjWh+f1VQKFI/d3e3WqNv6gYPVQInaDLAOzr
# /DKLHOK3vTaRIBHjWaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxMzEw
# MjBaMC8GCSqGSIb3DQEJBDEiBCD1sLbP6vWeKh/WGnPHN18WxdBZyJbSPlU+xgsy
# XMXa4zANBgkqhkiG9w0BAQEFAASCAgBMiUU/yQrqkdb+3dqtLox/W5fcqHzvohk/
# lpVNdU5RSWJ5aJOYMoG6JZZep0jSUzQpCgUEXyxdK35jNOZfE+dXzmmKJKLvYuET
# pjyFqsjvgdi5qfdCCLLOEVMIEhbSZ0wSTc9IqJmMQurfD/fEc2r+pkxELJiPHPVe
# Tpq8+bz+kFjwtAWC3ZCiqJagraRq3olRp07rnv3xrhb089gI4WLVSiD9XlkNMpjS
# m98G7ZV8wm27HCAlk+OyulM27Fxu2IKxdtCc8+85bWj1P83qkOTl6sEx+EGFNLgg
# BAZ6KzAY8TZ681BDOnShX7ZxjtuKOj5Zj+u/C04YE/CZCLT7gONlQF77QHu5SDwe
# OV7YF+7KF4WlOlUIWkxiWLI1Jf2ZuTpSVnMtT9AdiX/h0G1dgFa1xJCGP+L+uHNe
# UPPoNTd2Xy7xidAA+/WxO/1xE7fRX/90qK3wX2MK/PUa2QvagQTbhPMRE1eJNg+0
# DG02FVN01WmCcQ/O+XmjnKqQRHRTDGD7JVkioDo+N4/nqJZs5O8UEHur5MWJ4CXD
# ylRzVmYYg4biUrCaBNO7gZA6ZAs2GkDZTgDBgOGv40QjmG5Ua4TxGXTyihfFQNim
# RUnVWCpB0PllvCXLlglTMM5rC0JRjdQW65z7go3FQrJZD237qtaIRG1C5rmXRnlJ
# k5quXQcPEA==
# SIG # End signature block
