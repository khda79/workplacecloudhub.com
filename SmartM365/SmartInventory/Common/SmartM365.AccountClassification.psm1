# Account-classification rules are tenant configuration: the real rules live in the git-ignored
# Config/AccountClassification.local.json(.txt); only the neutral .template is published.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SmartM365AccountClassificationDefaultPath {
    [CmdletBinding()]
    param()
    return (Join-Path (Split-Path -Parent $PSScriptRoot) 'Config\AccountClassification.local.json')
}

function Resolve-SmartM365AccountClassificationPath {
    # Returns the existing file for a logical or physical path. As with the SmartM365 JSON
    # transport, <name>.json.txt is preferred to <name>.json.
    [CmdletBinding()]
    param([string]$Path = (Get-SmartM365AccountClassificationDefaultPath))

    $candidates = if ($Path -match '\.json$') { @("$Path.txt", $Path) } else { @($Path) }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).ProviderPath }
    }
    # Never fall back to the template: generic rules would silently reclassify the population.
    throw ("Account classification configuration not found: {0}. Create the private file from AccountClassification.local.json.template and adapt its rules to the tenant before use." -f ($candidates -join ' or '))
}

function ConvertTo-SmartM365CaseInsensitiveValue {
    # ConvertFrom-Json -AsHashtable returns case-sensitive dictionaries; the former
    # Import-PowerShellDataFile result was case-insensitive, so lookups keep that behavior.
    param([AllowNull()]$Value)

    if ($Value -is [System.Collections.IDictionary]) {
        $result = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $Value.Keys) { $result[[string]$key] = ConvertTo-SmartM365CaseInsensitiveValue $Value[$key] }
        return $result
    }
    if ($Value -is [System.Collections.IList] -and $Value -isnot [string]) {
        return ,@(foreach ($item in $Value) { ConvertTo-SmartM365CaseInsensitiveValue $item })
    }
    return $Value
}

function Read-SmartM365AccountClassification {
    [CmdletBinding()]
    param([string]$Path = (Get-SmartM365AccountClassificationDefaultPath))

    $resolvedPath = Resolve-SmartM365AccountClassificationPath -Path $Path
    $text = [IO.File]::ReadAllText($resolvedPath, [Text.UTF8Encoding]::new($false, $true))
    $config = ConvertTo-SmartM365CaseInsensitiveValue ($text | ConvertFrom-Json -AsHashtable -Depth 20)
    if ($config -isnot [hashtable]) { throw "Account classification configuration must be a JSON object: $resolvedPath" }
    if ([string]$config['SchemaVersion'] -ne '1.0') {
        throw ("Unsupported account classification configuration schema: {0}" -f $config['SchemaVersion'])
    }
    if ([string]::IsNullOrWhiteSpace([string]$config['RuleVersion'])) { throw "Account classification RuleVersion is required: $resolvedPath" }
    $config['SourcePath'] = $resolvedPath
    return $config
}

Export-ModuleMember -Function Get-SmartM365AccountClassificationDefaultPath, Resolve-SmartM365AccountClassificationPath, Read-SmartM365AccountClassification
