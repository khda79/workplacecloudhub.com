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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9nfchabHP4g7X
# Oa0r6k27Y/7FyxFmNy97uV5CBKrCVqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFaK3Ti21oy4SNk9LtjBGBy+iyn4tZuGcKqGqTMe4x1HMA0GCSqG
# SIb3DQEBAQUABIIBgAnwtr+R8UNSgoMPXQXDILr34q66p18ZqSr8E5ss2+8Sx7u9
# TzR76hMOBCqMYeSqXonLM6vyeCfN/mcuM8YhF80yd5baJgpjdOKSzQFkbXhJ2tpv
# CZXz1HkF/IQ+saaM8mQuKk1YhdIA6K8ZXgJFNMpWLmu1EX2FJLHS5jBWMOz34aEL
# nT4w+BAFqUORqNpFIHdtZ6BL6GfE8Kre/KSZIXQmAyZyeTgmFyvRtDtuMBz6UUaW
# 7DUFSUtDH+qHzR7R+UXAN3ZhnUNFds8TuewLbeMoEtoayZfWrxHr8PEyvqjI9mFo
# 7n0aHFppqv18ATIuwVTVc4kvHt/QQA6fyTxbj39Qim4nefsZQeB2E812scsBkL92
# NDCHzcJ3clzWdB3PLzVRO5aR2K1kgvcsXhXaw8fejFS8B0afZ9vHET9CORWX4b7E
# ut89zL+zUsBS8ugpAln5j4BCt9YuThZGrKX37o6KUIDXis4fbY+jPPb67KeOAV3/
# T/2INq9Q5O9Mgd8DSKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTJaMC8GCSqGSIb3DQEJBDEiBCArUme+sFTuxE/00sy7W++rX/s+e4fGjvcOUS3K
# BANegDANBgkqhkiG9w0BAQEFAASCAgCV8jnnYQWi0prO+0ubF+//id+55vUdTuXQ
# nW/V/4ClzKWILfQLqNmXwnMrMnXd2RfGJUa7TlMFrQVH0qdAmp8uqsj2qv7lBE1y
# WAihNOeFN8VyDQVaGl+ox0ST/va5nKShGQI3mkxiNkYPL2VcIdPOw80P8vnwx5QA
# Vh2zYoZxt07QR0VcJUIFTsp0JGQ/yVILMTn/8cTSTG2RSvtG2YjdKFiQWh8HSAJ6
# 6SSNu0CTeg2o+bD7XU8Tr9t/D3RYD7a4UcmEypuDmK/fNAwA44ckT0H2mcRJID5P
# dN5jDLsq//u0kOmvZXvn0G1MUnfcR59N8to41H/D4Me/1x8bqbDC8gZe0EnNDZfB
# tQGo4/C12VJ2JFR0cdrwZwowUFqrel22By3PCV4f5p3D1F1SjDhA8SUiGPaI9QJh
# YJgjW8lk/cqCxjO3DycR5OjSD9ONdOA6qQjcgrnsGtncKJtO8y5hJTAWmxidZQk7
# 6meijEqe426iays6BbI2ncaEDn3vS1BRAWc4BD7jHfdkf9ewyq5Vl3yUgEi5jpQQ
# 0aGjIiXFuUNrlTSIkVt08QwEQ07+Y3xt7F2ScGNQ+7D27LfjiKZoiw/sAvg5TrWT
# PSpYm68t27Ac9xlctiday0dk/G1jh4wsxGjgIwxICv060EMKhx26QdfGyQNSBvBC
# fywWaHwuRQ==
# SIG # End signature block
