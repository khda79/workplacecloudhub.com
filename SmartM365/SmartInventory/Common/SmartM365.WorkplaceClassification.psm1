# Site and persona classification from the private SmartWorkplaceIntelligence workbooks
# (SharePoint DATA root). Same rules as the prepared workforce evidence; no rule is hard-coded.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:PersonaWorkbookName = 'SmartWorkplaceIntelligence-PersonaClassification.xlsx'
$script:SiteWorkbookName = 'SmartWorkplaceIntelligence-SiteClassification.xlsx'
$script:RequiredWorksheets = @{
    $script:PersonaWorkbookName = @('Personas','Rules','Exclusions')
    $script:SiteWorkbookName = @('Settings','SiteTypes','Sites')
}
$script:Dash = [string][char]0x2014
$script:PersonaLabels = [pscustomobject]@{
    NonHuman = "Review required $($script:Dash) possible non-human"
    Ambiguous = "Review required $($script:Dash) multiple matches"
    MissingTitle = "Unclassified $($script:Dash) missing job title"
    NoMatch = "Unclassified $($script:Dash) no match"
}
$script:CountryCodes = @{ France = 'FR'; Germany = 'DE'; Spain = 'ES'; Belgium = 'BE'; Poland = 'PL'; Switzerland = 'CH'; Portugal = 'PT'; Italy = 'IT'; Austria = 'AT'; Ireland = 'IE'; Luxembourg = 'LU'; Netherlands = 'NL' }

function Get-SmartM365WorkplaceClassificationWorkbookName {
    [CmdletBinding()]
    param()
    [pscustomobject]@{ Persona = $script:PersonaWorkbookName; Site = $script:SiteWorkbookName }
}

function Get-SmartM365WorkplacePersonaLabel {
    [CmdletBinding()]
    param()
    $script:PersonaLabels
}

function Test-SmartM365WorkplaceClassificationWorkbook {
    # Structural check before any rule is read: a valid package with the required worksheets.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)

    if (-not $script:RequiredWorksheets.ContainsKey($Name)) { throw "Unknown classification workbook: $Name" }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf) -or (Get-Item -LiteralPath $Path).Length -eq 0) { throw "Classification workbook missing or empty: $Path" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $archive.GetEntry('xl/workbook.xml')
        if (-not $entry -or $entry.Length -gt 1048576) { throw 'Missing or oversized workbook metadata.' }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $xmlReader = [Xml.XmlReader]::Create([IO.StringReader]::new($xml), $settings)
        $document = [Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        try { $document.Load($xmlReader) } finally { $xmlReader.Dispose() }
        $sheets = @($document.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
        foreach ($sheet in $script:RequiredWorksheets[$Name]) { if ($sheet -notin $sheets) { throw "Required worksheet missing: $sheet" } }
    }
    catch { throw "Invalid classification workbook '$Name': $($_.Exception.Message)" }
    finally { if ($archive) { $archive.Dispose() } }
}

function Receive-SmartM365WorkplaceClassificationWorkbook {
    # Downloads both workbooks into Folder and returns their paths only when both are valid.
    # No cached or older copy is ever used as a fallback.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][scriptblock]$DownloadFile)

    New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    $paths = [ordered]@{}
    foreach ($name in @($script:PersonaWorkbookName, $script:SiteWorkbookName)) {
        $path = Join-Path $Folder $name
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        $receipt = & $DownloadFile $path $name
        if (-not $receipt -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required SharePoint classification workbook download failed: $name. No cached fallback is allowed." }
        Test-SmartM365WorkplaceClassificationWorkbook -Path $path -Name $name
        $paths[$name] = $path
    }
    [pscustomobject]@{ PersonaPath = $paths[$script:PersonaWorkbookName]; SitePath = $paths[$script:SiteWorkbookName] }
}

function Import-SmartM365ClassificationSheet {
    param([string]$Path, [string]$Worksheet)
    if (-not (Get-Module -Name ImportExcel)) {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) { throw 'ImportExcel PowerShell module is required to read the classification workbooks.' }
        Import-Module ImportExcel -ErrorAction Stop
    }
    @(Import-Excel -Path $Path -WorksheetName $Worksheet -ErrorAction Stop)
}

function ConvertTo-SmartM365ClassificationSearchText {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)
    $valueText = ([string]$Value).Trim().ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $builder = [Text.StringBuilder]::new()
    foreach ($character in $valueText.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($character) -eq [Globalization.UnicodeCategory]::NonSpacingMark) { continue }
        if ([char]::IsLetterOrDigit($character)) { [void]$builder.Append($character) }
        else { [void]$builder.Append(' ') }
    }
    return (' ' + (($builder.ToString() -replace '\s+', ' ').Trim()) + ' ')
}

function Read-SmartM365WorkplaceSiteClassification {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    Test-SmartM365WorkplaceClassificationWorkbook -Path $Path -Name $script:SiteWorkbookName
    $settings = @{}
    foreach ($row in (Import-SmartM365ClassificationSheet $Path 'Settings')) { $settings[([string]$row.Parameter).Trim()] = ([string]$row.Value).Trim() }
    $siteAttribute = $settings['Directory Site Attribute']
    $headquartersPersona = $settings['Headquarters Persona']
    if (-not $siteAttribute) { throw 'Directory Site Attribute is missing from site Settings.' }
    if (-not $headquartersPersona) { throw 'Headquarters Persona is missing from site Settings.' }
    $allowedSiteTypes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in (Import-SmartM365ClassificationSheet $Path 'SiteTypes')) {
        $allowedSiteType = ([string]$row.'Site Type Code').Trim()
        if ($allowedSiteType) { [void]$allowedSiteTypes.Add($allowedSiteType) }
    }
    if (-not $allowedSiteTypes.Contains('HQ') -or -not $allowedSiteTypes.Contains('UNKNOWN')) { throw 'SiteTypes must include HQ and UNKNOWN.' }
    $sites = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach ($row in (Import-SmartM365ClassificationSheet $Path 'Sites')) {
        if (([string]$row.'Use for Classification').Trim() -ne 'Yes') { continue }
        $code = ([string]$row.'Directory Site Code').Trim()
        $siteType = ([string]$row.'Site Type').Trim().ToUpperInvariant()
        if (-not $code -or $sites.ContainsKey($code)) { throw "Blank or duplicate enabled Directory Site Code: $code" }
        if (-not $allowedSiteTypes.Contains($siteType)) { throw "Unsupported Site Type for $code`: $siteType" }
        $integration = if ($row.PSObject.Properties['Integration Status']) { ([string]$row.'Integration Status').Trim().ToUpperInvariant() } else { '' }
        $sites[$code] = [pscustomobject]@{ Code = $code; Type = $siteType; Name = ([string]$row.'Site Display Name').Trim(); IntegrationStatus = $integration }
    }
    if ($sites.Count -eq 0) { throw 'No enabled site found in the site classification workbook.' }
    [pscustomobject]@{ SourcePath = $Path; SiteAttribute = $siteAttribute; HeadquartersPersona = $headquartersPersona; Sites = $sites }
}

function Get-SmartM365WorkplaceSite {
    # Returns the enabled site for a directory site code, or $null (blank or unmapped code).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$SiteClassification, [AllowEmptyString()][string]$SiteCode)
    $code = ([string]$SiteCode).Trim()
    if ($code -and $SiteClassification.Sites.ContainsKey($code)) { return $SiteClassification.Sites[$code] }
    return $null
}

function Read-SmartM365WorkplacePersonaClassification {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$HeadquartersPersona)

    Test-SmartM365WorkplaceClassificationWorkbook -Path $Path -Name $script:PersonaWorkbookName
    $personaNames = @{}
    foreach ($row in (Import-SmartM365ClassificationSheet $Path 'Personas')) {
        if (([string]$row.Enabled).Trim() -ne 'Yes') { continue }
        $personaId = ([string]$row.'Persona ID').Trim()
        $personaName = ([string]$row.'Persona name').Trim()
        if (-not $personaId -or -not $personaName -or $personaNames.ContainsKey($personaId)) { throw "Invalid or duplicate persona ID: $personaId" }
        $personaNames[$personaId] = $personaName
    }
    if ($personaNames.Count -eq 0) { throw 'No enabled personas found in the persona classification workbook.' }
    if ($personaNames.Values -notcontains $HeadquartersPersona) { throw "Headquarters Persona is not enabled in the persona workbook: $HeadquartersPersona" }
    $rules = @(Import-SmartM365ClassificationSheet $Path 'Rules' | ForEach-Object {
        if (([string]$_.Enabled).Trim() -ne 'Yes') { return }
        $keyword = ([string]$_.Keyword).Trim()
        if (-not $keyword) { return }
        $personaId = ([string]$_.'Persona ID').Trim()
        if (-not $personaNames.ContainsKey($personaId)) { throw "Unknown persona ID in rule $($_.'Rule ID'): $personaId" }
        if (([string]$_.'Match type').Trim() -ne 'Contains phrase') { throw "Unsupported match type in rule $($_.'Rule ID')" }
        [pscustomobject]@{ Keyword = $keyword; SearchKeyword = ConvertTo-SmartM365ClassificationSearchText $keyword; Category = $personaNames[$personaId]; Country = ([string]$_.Country).Trim().ToUpperInvariant(); Priority = [int]$_.Priority }
    })
    if ($rules.Count -eq 0) { throw 'No enabled persona rules found in the persona classification workbook.' }
    $exclusions = @(Import-SmartM365ClassificationSheet $Path 'Exclusions' | ForEach-Object {
        if (([string]$_.Enabled).Trim() -ne 'Yes') { return }
        $keyword = ([string]$_.Keyword).Trim()
        if (-not $keyword) { return }
        [pscustomobject]@{ Keyword = $keyword; SearchKeyword = ConvertTo-SmartM365ClassificationSearchText $keyword; Country = ([string]$_.Country).Trim().ToUpperInvariant() }
    })
    [pscustomobject]@{
        SourcePath = $Path; HeadquartersPersona = $HeadquartersPersona; Rules = $rules; Exclusions = $exclusions
        Cache = @{}; TextCache = @{}
    }
}

function Resolve-SmartM365PersonaTextMatch {
    param($Classification, [string]$Value, [string]$Field, [string]$CountryCode)
    if (-not $Value -or $Value -eq 'Unknown') {
        return [pscustomobject]@{ Persona = $script:PersonaLabels.MissingTitle; State = 'Missing job title'; Keyword = ''; Source = $Field }
    }
    $textCacheKey = "$CountryCode|$Value"
    if ($Classification.TextCache.ContainsKey($textCacheKey)) {
        $cached = $Classification.TextCache[$textCacheKey]
        return [pscustomobject]@{ Persona = $cached.Persona; State = $cached.State; Keyword = $cached.Keyword; Source = $Field }
    }
    $searchText = ConvertTo-SmartM365ClassificationSearchText $Value
    $exclusionKeywords = [Collections.Generic.HashSet[string]]::new()
    foreach ($rule in $Classification.Exclusions) {
        if (($rule.Country -eq 'ALL' -or $rule.Country -eq $CountryCode) -and $searchText.Contains($rule.SearchKeyword)) { [void]$exclusionKeywords.Add($rule.Keyword) }
    }
    if ($exclusionKeywords.Count) {
        $matched = [pscustomobject]@{ Persona = $script:PersonaLabels.NonHuman; State = 'Non-human keyword'; Keyword = ($exclusionKeywords -join ' | ') }
    }
    else {
        $topPriority = [int]::MinValue
        $categories = [Collections.Generic.HashSet[string]]::new()
        $keywords = [Collections.Generic.HashSet[string]]::new()
        foreach ($rule in $Classification.Rules) {
            if (($rule.Country -ne 'ALL' -and $rule.Country -ne $CountryCode) -or -not $searchText.Contains($rule.SearchKeyword)) { continue }
            if ($rule.Priority -gt $topPriority) { $topPriority = $rule.Priority; $categories.Clear(); $keywords.Clear() }
            if ($rule.Priority -eq $topPriority) { [void]$categories.Add($rule.Category); [void]$keywords.Add($rule.Keyword) }
        }
        $matched = if ($categories.Count -eq 0) { [pscustomobject]@{ Persona = $script:PersonaLabels.NoMatch; State = 'Unmatched'; Keyword = '' } }
        elseif ($categories.Count -gt 1) { [pscustomobject]@{ Persona = $script:PersonaLabels.Ambiguous; State = 'Ambiguous'; Keyword = ($keywords -join ' | ') } }
        else { [pscustomobject]@{ Persona = @($categories)[0]; State = 'Classified'; Keyword = ($keywords -join ' | ') } }
    }
    $Classification.TextCache[$textCacheKey] = $matched
    return [pscustomobject]@{ Persona = $matched.Persona; State = $matched.State; Keyword = $matched.Keyword; Source = $Field }
}

function Get-SmartM365WorkplacePersona {
    # Exclusions first, then HQ site, then job title, then AD description when the title has no unique match.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$PersonaClassification,
        [AllowEmptyString()][string]$JobTitle,
        [AllowEmptyString()][string]$Description,
        [AllowEmptyString()][string]$Country,
        [AllowEmptyString()][string]$SiteType
    )
    $countryText = ([string]$Country).Trim()
    $countryCode = if ($script:CountryCodes.ContainsKey($countryText)) { $script:CountryCodes[$countryText] } else { $countryText.ToUpperInvariant() }
    $titleText = ConvertTo-SmartM365ClassificationSearchText $JobTitle
    $descriptionText = ConvertTo-SmartM365ClassificationSearchText $Description
    foreach ($rule in $PersonaClassification.Exclusions) {
        if ($rule.Country -ne 'ALL' -and $rule.Country -ne $countryCode) { continue }
        if ($titleText.Contains($rule.SearchKeyword)) { return [pscustomobject]@{ Persona = $script:PersonaLabels.NonHuman; State = 'Non-human keyword'; Keyword = $rule.Keyword; Source = 'Job title' } }
        if ($descriptionText.Contains($rule.SearchKeyword)) { return [pscustomobject]@{ Persona = $script:PersonaLabels.NonHuman; State = 'Non-human keyword'; Keyword = $rule.Keyword; Source = 'AD description' } }
    }
    if ($SiteType -eq 'HQ') { return [pscustomobject]@{ Persona = $PersonaClassification.HeadquartersPersona; State = 'Classified'; Keyword = ''; Source = 'Site' } }
    $cacheKey = "$countryCode|$JobTitle|$Description"
    if ($PersonaClassification.Cache.ContainsKey($cacheKey)) { return $PersonaClassification.Cache[$cacheKey] }
    $result = Resolve-SmartM365PersonaTextMatch -Classification $PersonaClassification -Value $JobTitle -Field 'Job title' -CountryCode $countryCode
    if ($result.State -ne 'Classified' -and $result.State -ne 'Non-human keyword' -and $Description) {
        $fallback = Resolve-SmartM365PersonaTextMatch -Classification $PersonaClassification -Value $Description -Field 'AD description' -CountryCode $countryCode
        if ($fallback.State -eq 'Classified' -or $result.State -in @('Unmatched', 'Missing job title')) { $result = $fallback }
    }
    $PersonaClassification.Cache[$cacheKey] = $result
    return $result
}

Export-ModuleMember -Function Get-SmartM365WorkplaceClassificationWorkbookName, Get-SmartM365WorkplacePersonaLabel, Test-SmartM365WorkplaceClassificationWorkbook, Receive-SmartM365WorkplaceClassificationWorkbook, ConvertTo-SmartM365ClassificationSearchText, Read-SmartM365WorkplaceSiteClassification, Get-SmartM365WorkplaceSite, Read-SmartM365WorkplacePersonaClassification, Get-SmartM365WorkplacePersona

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCBIX3Tfg1yk8rv
# 7lVnF9snWPKb0aW8+VtFRJV4xiRtoqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBFWyVlRzWA3I3S6EMEsCsAUx9ehuoosl9ogE8dc7tU1MA0GCSqG
# SIb3DQEBAQUABIIBgBj83/fjKalkSINOqAQwT/cQYqEmHv9DF1WXW0SBMYBFNp6Q
# bRIsyyMk9VngfdOEqHT222NN6Mvb7w0zUz01mmpqBqTNwalvOCLsnl7Jz3nV5882
# uOrB4m+5m8mywJyJycfxc07Q4DVZD6s15wXNaFp/Y06aKjKuuQxSAZ9eGGddlxvV
# oXZj5FFto81Nv1Hi7hg4OH/RvRTms2j4djvDsME3mab2QDqhdVJMtQP9Cg2jK/wM
# y4yZjWvtrAyBgt8iVuKWzUDOhCrFpPOxLDQyLaAKB4eR1KP1rurEwfTN9af5acsQ
# I3L96lq6v8vdUzMN8sjmxN+xSH7e/uksz0+E5eh3m2olLY6Boab60ugJp0qoc0/R
# 4K+a26uZlBekNP9FuH/nUo1Z96BjYapqfE6H9Hi5BOGNCrOVuXpRzRUqD+XTzR7b
# AcKb8bs4QgGLkb7lNSGYKEnmvZXhj9jdUHvyiMfjO/JKuQzIxB4Tv03EYBD91DPY
# 6ExQx4y8aMgerMSg5KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDExNzUw
# MjFaMC8GCSqGSIb3DQEJBDEiBCBiktSBzT7YRmi0FnBt0H4xX9YY2OCU3lg5EgOL
# Ke3sFTANBgkqhkiG9w0BAQEFAASCAgBH5TBkjb1e3l7oPhGJvyQUpwd4r+VnDE84
# h0PFPgXS1/28TTpCr/CIqJw7UvgqKgxAn1MuT1T5mOQm9IaneMzw49Jh7XNtWg7s
# mTZ9w82OnV/5nA5BmZ/CnCl603jaoL691pEqpp8niE/zgHBoedyWRO0UxHwwUEYF
# R5MXtJxwuPJLB60Es2BrX6jOE9c/YLSvvpGfXcjxuj5cOjsEFttpZjWCwJU4rJcT
# 0eJ6u4S1YkT/jCMj8GoZL1OIeTNMiDvUzEeKrUKyL0JPYiKBcOn4c/TK7shiW8dL
# IF6ZdVpnU9YbJ0CZP69aoC38W4uhVp9t3PcwhJCUzl1GLPNWSKPwO++hQO/N4gKi
# Z+3fAh5CNXUgtBpXscigAZZLS9SlPOF+1LmbyVPWZddFnLNlBvflnCwTqu2o5qnt
# 6QGr89VSctE+rJ76X3N8V5ECLJhrSqJEvgq/+WrDKe0bNIJHCXrh53C6gjPb7iFD
# GeCwjcA44GZCAXGtJPe6bhGJ0FYCtXj9GA4fmWE29tUbSmA8sWzXrf56hRy3JOJX
# KDtrdbO2awR9UvaR+mnK18TUxzHWztdXH2JncexDv9JHuELn8zGZri/RWdV4KdLJ
# 107OT0/0JL6W9hDls5peji3cKJEZIEeHY4YJSsMM4Z/pIw2PEsjkew9/EESmIE6t
# TnG2VsQByw==
# SIG # End signature block
