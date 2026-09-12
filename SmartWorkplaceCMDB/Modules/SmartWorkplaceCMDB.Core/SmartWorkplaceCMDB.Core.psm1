# SmartWorkplaceCMDB.Core
# Version: 1.0.0

$script:SmartWorkplaceCMDBCoreVersion = '1.0.0'

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

    if ($raw.TrimStart().StartsWith('[')) {
        $wrapper = ('{"Items":' + $raw + '}') | ConvertFrom-Json -ErrorAction Stop
        return ,$wrapper.Items
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

        # Validate existing identity fields before projecting or replacing any output.
        $expectedIdentity = @{
            TenantKey = $TenantKey; OrganizationKey = $OrganizationKey
            EnvironmentKey = $EnvironmentKey; TenantId = $TenantId
        }
        foreach ($row in $InputObject) {
            foreach ($name in $identityColumns) {
                $property = $row.PSObject.Properties[$name]
                if ($null -ne $property -and [string]$property.Value -ne [string]$expectedIdentity[$name]) {
                    throw "CSV tenant identity mismatch for '$name'."
                }
            }
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
            if ($Columns -and $Columns.Count -gt 0) {
                $InputObject | Select-Object -Property $Columns |
                    Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8 -Force
            }
            else {
                $InputObject | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8 -Force
            }
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

. (Join-Path $PSScriptRoot 'SmartWorkplaceCMDB.Collection.ps1')

Export-ModuleMember -Function @(
    'Resolve-SmartWorkplaceCMDBCollectionPaths',
    'Start-SmartWorkplaceCMDBSourceCollection',
    'Complete-SmartWorkplaceCMDBSourceCollection',
    'Publish-SmartWorkplaceCMDBSourceCsv',
    'Import-SmartWorkplaceCMDBSourceCsv',
    'Get-SmartWorkplaceCMDBSourceHealth',
    'Read-SmartWorkplaceCMDBCollectionFixture',
        'Assert-SmartWorkplaceCMDBCollectionPage',

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAd47HrwI7rJBwU
# JdRriaqIKGt+WsFSsmVpPa5eOPi3kqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIExalrTbtSR3l/RUqH8wq3sxvtoRekFui0ruFQ82G/jSMA0GCSqG
# SIb3DQEBAQUABIIBgE6ekwKg0VM5EEMAsSJ6FfG6JpZ5XoZNJU8IfUy2G/sHRw/+
# /gHEKhcwBdR0+8ZaaswuqoTStcwLXzc3Cp0G0ErV9PA/A5VUTYZbU+yqhpiebVT+
# IC087k3Ti/PRKUo6W6XGkQJF9LdbCAQnozXBzUEcXoGMS8hGnYbOkeOiSbntsI2r
# +/Lv0hI1634d/GUi3xHRJfW8bT8zrIXCtOn9UfsiQUlpTO1gW302238daTyyFwJ4
# V9TmjlXrYbVJ7CHVVo7zAJ6nA0Vl9Anu0koPaT1EISnIxiw93kstru14fn6DsRQH
# YmqAtruskU4qn96qTrElg5E8tcS9z+VwoJzWUJB+e1oZXBk7jLAmE/AyTvKvTGXy
# kSd6trdkx4jtirYlU8h4u9Bl48oYiljOaLTk3Wl8FkZdiR03XDIGKH9TtucLhPmx
# VHQ0mSMrcHGb+WThqBrItoP66OW1qVkjoilv9GaoqyrJ4/ktk3bWmuCIVyqm0KL7
# utRvW8pNZgeoV36yQ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzE3
# MjBaMC8GCSqGSIb3DQEJBDEiBCDe9vgALkpzVLIv9SaYqn0mBFTNoYwBLU5G0YEY
# x3ro2jANBgkqhkiG9w0BAQEFAASCAgCWjwYfo0xGmcdvbca9LfeZq8clJqBbQ42a
# w3P9ig4aIp7iZsx4DqZjXfZOtE08beNtBN3FsLpKgJgpxO79KGPKbTcFgh3oWSoq
# WKpE7Vqg21J0ZrzXI7Lt4WOqbYlNOSsX7M+Tc5QMs65VMXjYp4g98wDglye3Ft/8
# hDFAd14uD0HKh6a+QXhbzWM86snkh1PUGBBSGjhREgtIUqS+3kgHrCLus/GLWQ0+
# PPWzwPwipJKdn7nGZ4BskQfyhhHoQwgLotYQlnguIvdMA+YVLJjbXhgT74Nbyyoo
# cBZM0En25WI+5It2kQggYzY4kNv24PWw6GGuImEB8eQISLBnOY6dHxo8l55tjTWT
# Udl9MHIN/25xOVRNWzIWZJ3kTczDnjH+8cKotyn2bJuaw+ZHR3BEdf4ZayfgS7nV
# x2ofBa+M0VOGcUtnFbDX9n6BmG5QWf0Y8+rGL+JBY8l6l98FylcFBJBdzk0v2Hac
# j7Mhl6HeyShdKa5+A/zRFjf+aKL2NcE60aJTQrEJJC1SYbKKHNGv8sZ302PTW8x8
# qXEedrxiM58P3GEI96QP5JCrTgQPVHdORMCdsZ7FuAw4J9qiIqAoL+OX4kTlo8se
# QwkbFl9hO/yWmSwi89NpQKxm/8m4S6DEyfhZY3x4JJU2SlEttt1EnsFlFKxMswqp
# dXCh30lSfw==
# SIG # End signature block
