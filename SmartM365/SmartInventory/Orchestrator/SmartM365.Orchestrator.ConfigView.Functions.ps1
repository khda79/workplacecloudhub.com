<#
.SYNOPSIS
Pure helpers of the read-only Orchestrator GUI Configuration view.
.DESCRIPTION
Flattening, descriptions, sensitive masking, effective values and filtering. No file access.
Dot-sourced by SmartM365.Orchestrator.ConfigView.psm1 and by the demo, which must not import modules.
.VERSION
1.0.2
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
        # Alphabetical by full parameter path, case-insensitive and culture-independent.
        $rows.Sort([Comparison[object]] { param($left, $right) [string]::Compare([string]$left.Parameter, [string]$right.Parameter, [StringComparison]::OrdinalIgnoreCase) })
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
    param([AllowEmptyCollection()][object[]]$Rows = @(), [string]$Filter = '', [switch]$ShowSensitive, [switch]$HideNotInTemplate)
    foreach ($row in @($Rows)) {
        # Keys that the template does not describe (old or custom keys) can be hidden from the view.
        if ($HideNotInTemplate -and $row.Model -eq 'Not in template') { continue }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDbwukuMvI9u5HH
# S0NEzmWE2bAgkiZResKOcoQfBJg4tKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILmiZte88+oG6RhWL11H1dDuJREBHoc5s4rlVBjQHr/9MA0GCSqG
# SIb3DQEBAQUABIIBgGlm6Z8ZoqLPzTW85D7tEG0uAfz3Vl3RMoIW6ZOl+JNLDX4+
# +pWkKsBC6FBltyo4V12Cm5Z45GIb0WtYzq/4BiUeW7kabaluIbCEuWYkXar1zpNE
# Aeisx8kU2c6/7xtFu0n93TW4ZTg2P1NjxVN4QN8ZQ6WXyYXG6+2UFZ/jKJxoWSIZ
# JHKCuhL8nu53QVnTEcf9BbOCcGt8TbJB1AUEjDUqWaciFJ31N24eWDuh8Uc8vPP1
# VXu3KZNO+5rT6L2bGDPSVZJkKW0+1ieqHUQ3b8Na2tt4RxZmnp63vJt4bftPDtf7
# p2GJt2unkuCb+s/uny5iOIFVI1HDfQw+lXL4eSmMk5Ymtcv19DyywEOiJMNmuX9e
# lz625dmJcdxTxTCRtqn6ow3vEtknHIgCGqea3JVceuJB2F05kxXaz6FwuWy5jHK3
# BcOmNrylYV/LTiRCgF+ivTZ5RUGG0fNhFWXgHPpG2uoqpKvqQIklZb2FBhtSAnNa
# w6CzuTONK7Dy/SYroKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgyMDQz
# MzZaMC8GCSqGSIb3DQEJBDEiBCAyxa1eq0sPpH1gb8GFSQznPvl8bFEU7fRiLChn
# UxEqkzANBgkqhkiG9w0BAQEFAASCAgA1VPN+wNPorWd/W9LKU4rFhz4nNSGkj/6a
# ygxYOg2bCkRbZccqjmYV7NY33IcKWA+6la4BHnPgG+CKBGPT6q4Nm0fq35DEB575
# jKcNgmJ8fh0PzvV7nlQVuRXUrZi82f7/psySjCHj6pik36MN6faMcSEAPhezJrgB
# GUPEgRezAvBQvgZEWYRv7aMVwfyPjNtiRwDd7TaEqj/6lDOfEDqmnNjtPgmij1FE
# 13KGXg3/7CvQE6DmNIpiolvFO1BWo9MC78wXvu3cbctY3uD4lpkeObNix4+X9EXg
# ZxYh2xvf05ah/Y0jDFqe27C3d49r75uDeA3klwhUaU1ID4vuqkqoDQ1riO7xG5w9
# 3fivG/s94o99FYWLljCuVBAgeB5ntwpIxDCfEIl5A9a175Nwpq+5YZk31JJZmvuQ
# XJuRvtSvztRyF6OnwKQTYd6VZ/xzQUT0bDsi6SQgrZzxOFsTKV2Ub8wISgRRksLP
# EuZ1CfLOhbpWRRCXQDI2xr2AwlctC1Tox23T2Dzssp6kTEr1SulWXCr0Z9gldb8S
# MAwPUbKndwDiZDS7kD5bg7bzmcHVl3Fzji7OZbcVl6zvSyCyZPIy2VmgTzWybDCv
# VpxAShiAf3Yn0mBFdePfG+U2rB+rT6D2JWaW/2zAZkweM7vpLweYGiNJdkLcU9rn
# VT1KG246WQ==
# SIG # End signature block
