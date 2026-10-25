<#
.SYNOPSIS
    Option catalog and effective option resolution.

.DESCRIPTION
    Every option is declared once in Config\SmartEndpointMigrationToolkit.options.json.txt.
    Resolution order, highest first:
      1. explicit CLI options
      2. GUI choices captured for the Run
      3. SEMT_* environment variables
      4. Wave configuration (Wave.local.json)
      5. global configuration (Config\SmartEndpointMigrationToolkit.local.json)
      6. campaign default from the catalog
    In Wave and global files, null, an empty string, __INHERIT__, __USE_GLOBAL__ and
    __USE_DEFAULT__ mean "inherit from the next layer".

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

$script:SemtInheritMarkers = @('__INHERIT__','__USE_GLOBAL__','USE_GLOBAL','__USE_DEFAULT__')
$script:SemtOptionLayers = @('Default','Global','Wave','Environment','Gui','Cli','Rule')

function Import-SemtOptionCatalog {
    [CmdletBinding()]
    param([string]$PackageRoot)

    if ([string]::IsNullOrWhiteSpace($PackageRoot)) { $PackageRoot = Get-SemtPackageRoot }
    $path = Join-Path -Path $PackageRoot -ChildPath 'Config\SmartEndpointMigrationToolkit.options.json.txt'
    $catalog = Read-SemtJsonFile -Path $path -Required
    $byName = @{}
    foreach ($definition in @($catalog.Options)) {
        $name = [string]$definition.Name
        if ([string]::IsNullOrWhiteSpace($name)) { throw "Option catalog entry without Name in $path" }
        if ($byName.ContainsKey($name)) { throw "Duplicate option '$name' in $path" }
        if ($definition.Scope -ne 'Common' -and -not $name.StartsWith(([string]$definition.Scope) + '.')) {
            throw "Option '$name' must be prefixed by its scope '$($definition.Scope)'."
        }
        $byName[$name] = $definition
    }
    return [pscustomobject]@{
        Path = $path
        Sha256 = Get-SemtFileSha256 -Path $path
        SchemaVersion = [int]$catalog.SchemaVersion
        Options = @($catalog.Options)
        ByName = $byName
    }
}

function Get-SemtCampaignOptionDefinitions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Catalog,
        [Parameter(Mandatory = $true)][string]$Campaign
    )

    [void](Assert-SemtCampaign -Campaign $Campaign)
    return @($Catalog.Options | Where-Object { $_.Scope -eq 'Common' -or $_.Scope -eq $Campaign })
}

function Get-SemtOptionDefault {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Definition,
        [Parameter(Mandatory = $true)][string]$Campaign
    )

    $campaignDefaults = Get-SemtPropertyValue -InputObject $Definition -Name 'CampaignDefaults' -Default $null
    if ($null -ne $campaignDefaults) {
        $campaignDefault = Get-SemtPropertyValue -InputObject $campaignDefaults -Name $Campaign -Default $null
        if ($null -ne $campaignDefault) { return $campaignDefault }
    }
    return $Definition.Default
}

function Test-SemtInheritValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) {
        $text = $Value.Trim()
        return ([string]::IsNullOrWhiteSpace($text) -or $script:SemtInheritMarkers -contains $text)
    }
    return $false
}

function ConvertTo-SemtOptionValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Definition,
        [AllowNull()]$Value,
        [string]$Source = 'Unknown'
    )

    $name = [string]$Definition.Name
    $type = [string]$Definition.Type
    switch ($type) {
        'Boolean' {
            if ($Value -is [bool]) { return [bool]$Value }
            $text = ([string]$Value).Trim()
            if ($text -in @('1','true','True','TRUE','yes','Yes','YES','on','On')) { return $true }
            if ($text -in @('0','false','False','FALSE','no','No','NO','off','Off')) { return $false }
            throw ("SEMT_OPTION_INVALID: {0} from {1} expects a boolean (true/false/1/0); received '{2}'." -f $name,$Source,$text)
        }
        'Integer' {
            $number = 0
            if ($Value -is [int] -or $Value -is [long]) { $number = [int64]$Value }
            elseif (-not [int64]::TryParse(([string]$Value).Trim(), [ref]$number)) {
                throw ("SEMT_OPTION_INVALID: {0} from {1} expects an integer; received '{2}'." -f $name,$Source,$Value)
            }
            $minimum = Get-SemtPropertyValue -InputObject $Definition -Name 'Minimum' -Default $null
            $maximum = Get-SemtPropertyValue -InputObject $Definition -Name 'Maximum' -Default $null
            if ($null -ne $minimum -and $number -lt [int64]$minimum) {
                throw ("SEMT_OPTION_INVALID: {0} from {1} is below the minimum {2}{3}; received {4}." -f $name,$Source,$minimum,$(if ($Definition.Unit) { " $($Definition.Unit)" } else { '' }),$number)
            }
            if ($null -ne $maximum -and $number -gt [int64]$maximum) {
                throw ("SEMT_OPTION_INVALID: {0} from {1} is above the maximum {2}{3}; received {4}." -f $name,$Source,$maximum,$(if ($Definition.Unit) { " $($Definition.Unit)" } else { '' }),$number)
            }
            return [int]$number
        }
        'Choice' {
            $text = ([string]$Value).Trim()
            foreach ($choice in @($Definition.Choices)) {
                if ([string]::Equals([string]$choice, $text, [System.StringComparison]::OrdinalIgnoreCase)) { return [string]$choice }
            }
            throw ("SEMT_OPTION_INVALID: {0} from {1} expects one of {2}; received '{3}'." -f $name,$Source,(@($Definition.Choices) -join ', '),$text)
        }
        'Path' {
            $text = ([string]$Value).Trim().Trim('"')
            if ($text -match '[\r\n]') { throw ("SEMT_OPTION_INVALID: {0} from {1} must be a single path." -f $name,$Source) }
            return $text
        }
        'String' {
            $text = ([string]$Value).Trim()
            if ($text -match '[\r\n"]') { throw ("SEMT_OPTION_INVALID: {0} from {1} cannot contain quotes or line breaks." -f $name,$Source) }
            return $text
        }
        default { throw ("Option catalog type '{0}' is not supported for {1}." -f $type,$name) }
    }
}

function ConvertTo-SemtOptionLayer {
    <#
    Normalizes a layer (hashtable, PSCustomObject or file object with an Options member) to a
    case-insensitive hashtable keyed by option name. Inherit markers are kept so the caller decides.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Layer)

    $result = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $Layer) { return $result }
    $source = $Layer
    $optionsMember = Get-SemtPropertyValue -InputObject $Layer -Name 'Options' -Default $null
    if ($null -ne $optionsMember) { $source = $optionsMember }
    if ($source -is [System.Collections.IDictionary]) {
        foreach ($key in @($source.Keys)) { $result[[string]$key] = $source[$key] }
    }
    else {
        foreach ($property in $source.PSObject.Properties) { $result[$property.Name] = $property.Value }
    }
    return $result
}

function Get-SemtEnvironmentOptionLayer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Catalog,
        [Parameter(Mandatory = $true)][string]$Campaign
    )

    $layer = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($definition in @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign)) {
        $envName = [string](Get-SemtPropertyValue -InputObject $definition -Name 'EnvVar' -Default '')
        if ([string]::IsNullOrWhiteSpace($envName)) { continue }
        $value = [Environment]::GetEnvironmentVariable($envName, 'Process')
        if (-not [string]::IsNullOrWhiteSpace($value)) { $layer[[string]$definition.Name] = $value }
    }
    return $layer
}

function Resolve-SemtEffectiveOptions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Catalog,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [AllowNull()]$GlobalLayer,
        [AllowNull()]$WaveLayer,
        [AllowNull()]$EnvironmentLayer,
        [AllowNull()]$GuiLayer,
        [AllowNull()]$CliLayer
    )

    [void](Assert-SemtCampaign -Campaign $Campaign)
    $definitions = @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign)
    $known = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($definition in $definitions) { $known[[string]$definition.Name] = $definition }

    $layers = [ordered]@{
        Global = ConvertTo-SemtOptionLayer -Layer $GlobalLayer
        Wave = ConvertTo-SemtOptionLayer -Layer $WaveLayer
        Environment = ConvertTo-SemtOptionLayer -Layer $EnvironmentLayer
        Gui = ConvertTo-SemtOptionLayer -Layer $GuiLayer
        Cli = ConvertTo-SemtOptionLayer -Layer $CliLayer
    }

    foreach ($layerName in @($layers.Keys)) {
        foreach ($key in @($layers[$layerName].Keys)) {
            if ($known.ContainsKey($key)) { continue }
            $otherCampaign = $Catalog.ByName.ContainsKey($key)
            if ($otherCampaign) { continue }
            throw ("SEMT_OPTION_UNKNOWN: '{0}' from {1} is not a known option. Check the option catalog." -f $key,$layerName)
        }
    }

    $effective = [ordered]@{}
    foreach ($definition in $definitions) {
        $name = [string]$definition.Name
        $value = ConvertTo-SemtOptionValue -Definition $definition -Value (Get-SemtOptionDefault -Definition $definition -Campaign $Campaign) -Source 'Default'
        $source = 'Default'
        foreach ($layerName in @('Global','Wave','Environment','Gui','Cli')) {
            $layer = $layers[$layerName]
            if (-not $layer.ContainsKey($name)) { continue }
            $raw = $layer[$name]
            $inheritAllowed = ($layerName -in @('Global','Wave','Environment'))
            if ($inheritAllowed -and (Test-SemtInheritValue -Value $raw)) { continue }
            if (-not $inheritAllowed -and $null -eq $raw) { continue }
            $value = ConvertTo-SemtOptionValue -Definition $definition -Value $raw -Source $layerName
            $source = $layerName
        }
        $effective[$name] = [pscustomobject]@{
            Name = $name
            Value = $value
            Source = $source
            Scope = [string]$definition.Scope
            Category = [string]$definition.Category
            Unit = [string]$definition.Unit
            Target = [string]$definition.Target
        }
    }

    if ($effective['DryRun'].Value -and $effective['MaxCycles'].Value -eq 0) {
        $effective['MaxCycles'] = [pscustomobject]@{ Name = 'MaxCycles'; Value = 1; Source = 'Rule:DryRunSingleCycle'; Scope = 'Common'; Category = 'Schedule'; Unit = 'cycles'; Target = 'Technician' }
    }
    return $effective
}

function Get-SemtOption {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (-not $Options.Contains($Name)) { throw "Option '$Name' is not resolved for this campaign." }
    return $Options[$Name].Value
}

function Set-SemtRuleOption {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value,
        [Parameter(Mandatory = $true)][string]$Rule
    )

    $current = $Options[$Name]
    $Options[$Name] = [pscustomobject]@{ Name = $Name; Value = $Value; Source = ('Rule:{0}' -f $Rule); Scope = $current.Scope; Category = $current.Category; Unit = $current.Unit; Target = $current.Target }
}

function ConvertTo-SemtEndpointArguments {
    <#
    Builds the endpoint argument list from options marked Target=Endpoint. Per-computer and
    engine arguments (DataRoot, CorrelationId, IgnoreRunGuard) are added by the engine.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Catalog,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options
    )

    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($definition in @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign)) {
        $parameter = [string](Get-SemtPropertyValue -InputObject $definition -Name 'EndpointParameter' -Default '')
        if ([string]::IsNullOrWhiteSpace($parameter)) { continue }
        $value = $Options[[string]$definition.Name].Value
        switch ([string]$definition.EndpointEmit) {
            'Switch' { if ([bool]$value) { $arguments.Add("-$parameter") } }
            'Value' { $arguments.Add("-$parameter"); $arguments.Add([string]$value) }
            'ValueWhenNotEmpty' { if (-not [string]::IsNullOrWhiteSpace([string]$value)) { $arguments.Add("-$parameter"); $arguments.Add([string]$value) } }
            'ValueWhenPositive' { if ([int]$value -gt 0) { $arguments.Add("-$parameter"); $arguments.Add([string]$value) } }
            default { throw ("Option {0} has an unsupported EndpointEmit value '{1}'." -f $definition.Name,$definition.EndpointEmit) }
        }
    }
    return @($arguments.ToArray())
}

function Get-SemtOptionRows {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options)

    return @($Options.Values | ForEach-Object {
        [pscustomobject]@{
            Category = $_.Category
            Option = $_.Name
            Value = if ($_.Value -is [bool]) { [string]$_.Value } else { [string]$_.Value }
            Unit = $_.Unit
            Source = $_.Source
        }
    })
}

function Read-SemtGlobalConfig {
    <#
    Loads Config\SmartEndpointMigrationToolkit.local.json, self-healing missing keys from the
    committed template. A missing runtime file is not created by read-only callers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [switch]$CreateWhenMissing
    )

    $template = Join-Path -Path $PackageRoot -ChildPath 'Config\SmartEndpointMigrationToolkit.local.json.template'
    $runtime = Join-Path -Path $PackageRoot -ChildPath 'Config\SmartEndpointMigrationToolkit.local.json'
    return (Sync-SemtJsonWithTemplate -Path $runtime -TemplatePath $template -CreateWhenMissing:$CreateWhenMissing)
}

function Read-SemtWaveConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PackageRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$WavePath,
        [switch]$CreateWhenMissing
    )

    $template = Join-Path -Path $PackageRoot -ChildPath ("Campaigns\{0}\Templates\Wave\Wave.local.json.template" -f $Campaign)
    $runtime = Join-Path -Path $WavePath -ChildPath 'Wave.local.json'
    return (Sync-SemtJsonWithTemplate -Path $runtime -TemplatePath $template -CreateWhenMissing:$CreateWhenMissing)
}

function Read-SemtGuiSelection {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][string]$Campaign
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $selection = Read-SemtJsonFile -Path $Path -Required
    $selectionCampaign = [string](Get-SemtPropertyValue -InputObject $selection -Name 'Campaign' -Default '')
    if ($selectionCampaign -ne $Campaign) {
        throw ("SEMT_GUI_SELECTION_CAMPAIGN_MISMATCH: selection {0} targets Campaign '{1}', not '{2}'." -f $Path,$selectionCampaign,$Campaign)
    }
    return $selection
}
