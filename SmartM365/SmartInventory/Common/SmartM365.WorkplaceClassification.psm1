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
