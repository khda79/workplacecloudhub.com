<#
.SYNOPSIS
Pure helpers of the read-only Orchestrator GUI Configuration view.
.DESCRIPTION
Flattening, descriptions, sensitive masking, effective values and filtering. No file access.
Dot-sourced by SmartM365.Orchestrator.ConfigView.psm1 and by the demo, which must not import modules.
.VERSION
1.0.0
#>

$script:ConfigViewPlaceholders = @('__USE_GLOBAL__', 'USE_GLOBAL')
$script:ConfigViewMask = '******** (hidden)'

function ConvertTo-SmartM365OrchestratorConfigEntries {
    <# Flattens a JSON object to dotted paths. Arrays, scalars and empty objects are leaves. #>
    [CmdletBinding()]
    param([AllowNull()]$Document, [string]$Prefix = '')
    if ($Document -is [System.Management.Automation.PSCustomObject]) {
        $properties = @($Document.PSObject.Properties)
        if ($properties.Count -eq 0 -and $Prefix) { return [pscustomobject]@{ Path = $Prefix; Value = $Document; IsEmptyObject = $true } }
        foreach ($property in $properties) {
            $path = if ($Prefix) { $Prefix + '.' + $property.Name } else { $property.Name }
            ConvertTo-SmartM365OrchestratorConfigEntries -Document $property.Value -Prefix $path
        }
        return
    }
    [pscustomobject]@{ Path = $Prefix; Value = $Document; IsEmptyObject = $false }
}

function Test-ConfigViewPathPattern {
    # '*' matches exactly one path segment.
    param([Parameter(Mandatory)][string]$Pattern, [Parameter(Mandatory)][string]$Path)
    $patternParts = $Pattern.Split('.'); $pathParts = $Path.Split('.')
    if ($patternParts.Count -ne $pathParts.Count) { return $false }
    for ($i = 0; $i -lt $patternParts.Count; $i++) {
        if ($patternParts[$i] -ne '*' -and $patternParts[$i] -ine $pathParts[$i]) { return $false }
    }
    return $true
}

function Get-SmartM365OrchestratorConfigDescription {
    <# Exact description first, then the most specific '*' pattern, then the nearest described parent. #>
    [CmdletBinding()]
    param([AllowNull()]$Descriptions, [Parameter(Mandatory)][string]$Path)
    if ($null -eq $Descriptions) { return '' }
    $exact = $Descriptions.PSObject.Properties[$Path]
    if ($null -ne $exact) { return [string]$exact.Value }
    $best = $null; $bestWildcards = [int]::MaxValue
    foreach ($property in @($Descriptions.PSObject.Properties | Where-Object { $_.Name.Contains('*') })) {
        if (Test-ConfigViewPathPattern -Pattern $property.Name -Path $Path) {
            $wildcards = @($property.Name.Split('.') | Where-Object { $_ -eq '*' }).Count
            if ($wildcards -lt $bestWildcards) { $best = $property; $bestWildcards = $wildcards }
        }
    }
    if ($null -ne $best) { return [string]$best.Value }
    $parts = $Path.Split('.')
    for ($length = $parts.Count - 1; $length -ge 1; $length--) {
        $parent = $Descriptions.PSObject.Properties[($parts[0..($length - 1)] -join '.')]
        if ($null -ne $parent) { return [string]$parent.Value }
    }
    return ''
}

function Test-SmartM365OrchestratorConfigSensitive {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Patterns = @(), [Parameter(Mandatory)][string]$Path)
    $leaf = $Path.Split('.')[-1]
    foreach ($pattern in $Patterns) { if ($leaf -like $pattern -or $Path -like $pattern) { return $true } }
    return $false
}

function Format-ConfigViewValue {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [array] -or $Value -is [System.Collections.IList]) {
        return (ConvertTo-Json -InputObject $Value -Depth 20 -Compress)
    }
    return [string]$Value
}

function Test-ConfigViewInherits {
    # Same fallback rule as Get-SmartM365EffectiveGlobalConfig: blank, null and placeholders inherit.
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { $text = $Value.Trim(); return ([string]::IsNullOrWhiteSpace($text) -or $text -in $script:ConfigViewPlaceholders) }
    return $false
}

function Get-ConfigViewTopValue {
    param([AllowNull()]$Document, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Document) { return [pscustomobject]@{ Found = $false; Value = $null } }
    $property = $Document.PSObject.Properties[$Name]
    if ($null -eq $property) { return [pscustomobject]@{ Found = $false; Value = $null } }
    return [pscustomobject]@{ Found = $true; Value = $property.Value }
}

function Resolve-ConfigViewInherited {
    # Effective value of a top-level key that the tenant leaves to the global configuration.
    param([AllowNull()]$GlobalDocument, [AllowNull()]$GlobalTemplate, [Parameter(Mandatory)][string]$Name)
    $global = Get-ConfigViewTopValue $GlobalDocument $Name
    # The global file is final when it holds the key, even blank (it has no further fallback).
    if ($global.Found) { return [pscustomobject]@{ Value = $global.Value; Source = 'global' } }
    $default = Get-ConfigViewTopValue $GlobalTemplate $Name
    if ($default.Found) { return [pscustomobject]@{ Value = $default.Value; Source = 'default (global template)' } }
    return [pscustomobject]@{ Value = $null; Source = 'default (script)' }
}

function New-SmartM365OrchestratorConfigurationView {
    <# Pure builder: one row per key of each file plus template keys missing from it. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProfileKey, [Parameter(Mandatory)]$Files, [AllowNull()]$Descriptions = $null
    )
    $patterns = @(if ($null -ne $Descriptions -and $Descriptions.PSObject.Properties['SensitiveKeyPatterns']) { @($Descriptions.SensitiveKeyPatterns | ForEach-Object { [string]$_ }) })
    $describedFiles = if ($null -ne $Descriptions -and $Descriptions.PSObject.Properties['Files']) { $Descriptions.Files } else { $null }
    $globalFile = @($Files | Where-Object Kind -eq 'Global')[0]
    $tenantFile = @($Files | Where-Object Kind -eq 'Tenant')[0]
    $globalDocument = if ($null -ne $globalFile) { $globalFile.Document } else { $null }
    $globalTemplate = if ($null -ne $globalFile) { $globalFile.Template } else { $null }
    $tenantDocument = if ($null -ne $tenantFile) { $tenantFile.Document } else { $null }
    $views = foreach ($file in @($Files)) {
        $descriptions = if ($null -ne $describedFiles -and $describedFiles.PSObject.Properties[$file.DescriptionKey]) { $describedFiles.($file.DescriptionKey) } else { $null }
        $entries = @(if ($null -ne $file.Document) { ConvertTo-SmartM365OrchestratorConfigEntries -Document $file.Document })
        $templateEntries = @(if ($null -ne $file.Template) { ConvertTo-SmartM365OrchestratorConfigEntries -Document $file.Template })
        $templatePaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $templateEntries) { [void]$templatePaths.Add($entry.Path) }
        $freeMaps = @($templateEntries | Where-Object IsEmptyObject | ForEach-Object Path)
        $filePaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $rows = [Collections.Generic.List[object]]::new()
        foreach ($entry in $entries) {
            [void]$filePaths.Add($entry.Path)
            $inFreeMap = @($freeMaps | Where-Object { $entry.Path.StartsWith($_ + '.', [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
            $model = if ($null -eq $file.Template) { 'No template' } elseif ($templatePaths.Contains($entry.Path) -or $inFreeMap) { 'OK' } else { 'Not in template' }
            $effective = $entry.Value; $source = $file.SourceName
            if ($file.Kind -eq 'Tenant' -and (Test-ConfigViewInherits $entry.Value)) {
                $topName = $entry.Path.Split('.')[0]
                if ($topName -eq $entry.Path) {
                    $inherited = Resolve-ConfigViewInherited -GlobalDocument $globalDocument -GlobalTemplate $globalTemplate -Name $topName
                    $effective = $inherited.Value; $source = $inherited.Source
                }
            }
            elseif ($file.Kind -eq 'Mail' -and (Test-ConfigViewInherits $entry.Value)) {
                $field = $entry.Path.Split('.')[-1]
                if ($field -in @('From', 'Cc')) {
                    $tenantValue = Get-ConfigViewTopValue $tenantDocument $field
                    if ($tenantValue.Found -and -not (Test-ConfigViewInherits $tenantValue.Value)) { $effective = $tenantValue.Value; $source = 'tenant' }
                    else { $inherited = Resolve-ConfigViewInherited -GlobalDocument $globalDocument -GlobalTemplate $globalTemplate -Name $field; $effective = $inherited.Value; $source = $inherited.Source }
                }
            }
            $rows.Add([pscustomobject]@{
                Parameter = $entry.Path; RawValue = (Format-ConfigViewValue $entry.Value); RawEffective = (Format-ConfigViewValue $effective)
                Source = $source; Model = $model; Description = (Get-SmartM365OrchestratorConfigDescription -Descriptions $descriptions -Path $entry.Path)
                Sensitive = (Test-SmartM365OrchestratorConfigSensitive -Patterns $patterns -Path $entry.Path)
            })
        }
        if ($null -ne $file.Document) {
            foreach ($entry in $templateEntries) {
                if ($filePaths.Contains($entry.Path)) { continue }
                if ($entry.IsEmptyObject -and @($filePaths | Where-Object { $_.StartsWith($entry.Path + '.', [StringComparison]::OrdinalIgnoreCase) }).Count) { continue }
                $rows.Add([pscustomobject]@{
                    Parameter = $entry.Path; RawValue = '(missing)'; RawEffective = ''; Source = ''; Model = 'Missing from file'
                    Description = (Get-SmartM365OrchestratorConfigDescription -Descriptions $descriptions -Path $entry.Path)
                    Sensitive = (Test-SmartM365OrchestratorConfigSensitive -Patterns $patterns -Path $entry.Path)
                })
            }
        }
        [pscustomobject]@{
            Kind = $file.Kind; Title = $file.Title; Path = $file.Path; Status = $file.Status; Error = $file.Error; Rows = @($rows)
            NotInTemplate = @($rows | Where-Object Model -eq 'Not in template').Count; MissingFromFile = @($rows | Where-Object Model -eq 'Missing from file').Count
        }
    }
    [pscustomobject]@{ ProfileKey = $ProfileKey; Files = @($views); SensitivePatterns = $patterns }
}

function Select-SmartM365OrchestratorConfigRows {
    <# Display rows: sensitive values masked unless revealed; the filter never matches a hidden value. #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Rows = @(), [string]$Filter = '', [switch]$ShowSensitive)
    foreach ($row in @($Rows)) {
        $masked = $row.Sensitive -and -not $ShowSensitive -and $row.Model -ne 'Missing from file'
        $value = if ($masked) { $script:ConfigViewMask } else { $row.RawValue }
        $effective = if ($masked -and $row.RawEffective) { $script:ConfigViewMask } else { $row.RawEffective }
        if ($Filter) {
            $haystack = @($row.Parameter, $row.Description, $row.Source, $row.Model) + $(if ($masked) { @() } else { @($value, $effective) })
            if (-not @($haystack | Where-Object { ([string]$_).IndexOf($Filter, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count) { continue }
        }
        [pscustomobject]@{
            Parameter = $row.Parameter; Value = $value; Effective = $effective; Source = $row.Source; Model = $row.Model
            Description = $row.Description; Sensitive = [bool]$row.Sensitive; Masked = $masked
        }
    }
}
