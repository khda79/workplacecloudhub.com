[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$UserOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserInventoryEvidence.csv'),

    [string]$HistoryOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserActivityHistoryEvidence.csv'),

    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\WorkforceTrendEvidence.csv'),

    [string]$IdentityOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\IdentityReconciliationEvidence.csv'),

    [string]$SignalsOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\WorkforceOperationalSignals.csv'),

    [Alias('PersonaKeywordsPath')]
    [string]$PersonaClassificationPath,

    [string]$SiteClassificationPath,

    [switch]$SkipHistory,

    [string]$AccountClassificationConfigPath = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'SmartM365\SmartInventory\Config\AccountClassification.psd1')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $PersonaClassificationPath) {
    $PersonaClassificationPath = Join-Path $DataRoot 'SmartWorkplaceIntelligence-PersonaClassification.xlsx'
}
if (-not $SiteClassificationPath) {
    $SiteClassificationPath = Join-Path $DataRoot 'SmartWorkplaceIntelligence-SiteClassification.xlsx'
}

if (-not (Test-Path -LiteralPath $AccountClassificationConfigPath -PathType Leaf)) {
    throw "Account classification configuration not found: $AccountClassificationConfigPath"
}
$accountClassificationConfig = Import-PowerShellDataFile -LiteralPath $AccountClassificationConfigPath
if ([string]$accountClassificationConfig.SchemaVersion -ne '1.0') {
    throw "Unsupported account classification configuration schema: $($accountClassificationConfig.SchemaVersion)"
}
$accountPopulationRules = $accountClassificationConfig.Population
$accountClassificationRuleVersion = [string]$accountClassificationConfig.RuleVersion

function Get-ConfiguredAccountPopulation {
    param([AllowNull()][object]$AccountType)
    $normalizedType = ([string]$AccountType).Trim()
    if ($accountPopulationRules.Human -contains $normalizedType) { return 'Human' }
    if ($accountPopulationRules.NonHuman -contains $normalizedType) { return 'Non-human' }
    return 'Review required'
}

function Get-RequiredCsv {
    param([string]$RelativePath)

    $path = Join-Path $DataRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required source CSV was not found: $path"
    }
    return $path
}

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Get-OfficeDesktopUsageState {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return 'Unknown - no app usage report row' }
    $fields = @(
        'Outlook (Windows)', 'Word (Windows)', 'Excel (Windows)', 'PowerPoint (Windows)', 'OneNote (Windows)'
    )
    $observed = 0
    foreach ($field in $fields) {
        $property = $Row.PSObject.Properties[$field]
        if ($null -eq $property) { continue }
        $value = ([string]$property.Value).Trim().ToLowerInvariant()
        if ($value -in @('true', 'yes', '1')) { return 'Used on PC in 30D' }
        if ($value -in @('false', 'no', '0')) { $observed++ }
    }
    if ($observed -eq $fields.Count) { return 'No PC app use in 30D' }
    return 'Unknown - incomplete app usage row'
}

function Convert-ToSearchText {
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

function Get-PersonaClassification {
    param([string]$JobTitle, [string]$Description, [string]$Country, [string]$SiteType)
    $countryCode = if ($script:personaCountryCodes.ContainsKey($Country)) { $script:personaCountryCodes[$Country] } else { $Country.ToUpperInvariant() }
    $titleText = Convert-ToSearchText $JobTitle
    $descriptionText = Convert-ToSearchText $Description
    foreach ($rule in $script:personaExclusions) {
        if ($rule.Country -ne 'ALL' -and $rule.Country -ne $countryCode) { continue }
        if ($titleText.Contains($rule.SearchKeyword)) {
            return [pscustomobject]@{ Persona = 'Review required — possible non-human'; State = 'Non-human keyword'; Keyword = $rule.Keyword; Source = 'Job title' }
        }
        if ($descriptionText.Contains($rule.SearchKeyword)) {
            return [pscustomobject]@{ Persona = 'Review required — possible non-human'; State = 'Non-human keyword'; Keyword = $rule.Keyword; Source = 'AD description' }
        }
    }
    if ($SiteType -eq 'HQ') {
        return [pscustomobject]@{ Persona = $script:headquartersPersona; State = 'Classified'; Keyword = ''; Source = 'Site' }
    }
    $cacheKey = "$countryCode|$JobTitle|$Description"
    if ($script:personaCache.ContainsKey($cacheKey)) { return $script:personaCache[$cacheKey] }
    function Resolve-TextMatch {
        param([string]$Value, [string]$Field)
        if (-not $Value -or $Value -eq 'Unknown') {
            return [pscustomobject]@{ Persona = 'Unclassified — missing job title'; State = 'Missing job title'; Keyword = ''; Source = $Field }
        }
        $textCacheKey = "$countryCode|$Value"
        if ($script:personaTextCache.ContainsKey($textCacheKey)) {
            $cached = $script:personaTextCache[$textCacheKey]
            return [pscustomobject]@{ Persona = $cached.Persona; State = $cached.State; Keyword = $cached.Keyword; Source = $Field }
        }
        $searchText = Convert-ToSearchText $Value
        $exclusionKeywords = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($rule in $script:personaExclusions) {
            if (($rule.Country -eq 'ALL' -or $rule.Country -eq $countryCode) -and $searchText.Contains($rule.SearchKeyword)) {
                [void]$exclusionKeywords.Add($rule.Keyword)
            }
        }
        if ($exclusionKeywords.Count) {
            $matched = [pscustomobject]@{ Persona = 'Review required — possible non-human'; State = 'Non-human keyword'; Keyword = ($exclusionKeywords -join ' | ') }
            $script:personaTextCache[$textCacheKey] = $matched
            return [pscustomobject]@{ Persona = $matched.Persona; State = $matched.State; Keyword = $matched.Keyword; Source = $Field }
        }
        $topPriority = [int]::MinValue
        $categories = [System.Collections.Generic.HashSet[string]]::new()
        $keywords = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($rule in $script:personaKeywords) {
            if (($rule.Country -ne 'ALL' -and $rule.Country -ne $countryCode) -or -not $searchText.Contains($rule.SearchKeyword)) { continue }
            if ($rule.Priority -gt $topPriority) {
                $topPriority = $rule.Priority
                $categories.Clear()
                $keywords.Clear()
            }
            if ($rule.Priority -eq $topPriority) {
                [void]$categories.Add($rule.Category)
                [void]$keywords.Add($rule.Keyword)
            }
        }
        if ($categories.Count -eq 0) {
            $matched = [pscustomobject]@{ Persona = 'Unclassified — no match'; State = 'Unmatched'; Keyword = '' }
        } elseif ($categories.Count -gt 1) {
            $matched = [pscustomobject]@{ Persona = 'Review required — multiple matches'; State = 'Ambiguous'; Keyword = ($keywords -join ' | ') }
        } else {
            $matched = [pscustomobject]@{ Persona = @($categories)[0]; State = 'Classified'; Keyword = ($keywords -join ' | ') }
        }
        $script:personaTextCache[$textCacheKey] = $matched
        return [pscustomobject]@{ Persona = $matched.Persona; State = $matched.State; Keyword = $matched.Keyword; Source = $Field }
    }
    $result = Resolve-TextMatch -Value $JobTitle -Field 'Job title'
    if ($result.State -ne 'Classified' -and $result.State -ne 'Non-human keyword' -and $Description) {
        $fallback = Resolve-TextMatch -Value $Description -Field 'AD description'
        if ($fallback.State -eq 'Classified' -or $result.State -in @('Unmatched', 'Missing job title')) { $result = $fallback }
    }
    $script:personaCache[$cacheKey] = $result
    return $result
}

function Get-PropertyValue {
    param(
        [AllowNull()][object]$Row,
        [string]$PropertyName
    )
    if ($null -eq $Row) { return $null }
    $property = $Row.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Convert-ToDateTimeOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = [datetime]::MinValue
    foreach ($culture in @(
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.CultureInfo]::GetCultureInfo('en-US'),
        [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    )) {
        if ([datetime]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
            return $parsed
        }
    }
    return $null
}

function Convert-AdLogonDateTimeOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $text = ([string]$Value).Trim()
    if ($text -match '^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$') {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact(
            $text, 'dd/MM/yyyy HH:mm:ss',
            [Globalization.CultureInfo]::GetCultureInfo('fr-FR'),
            [Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed
        )) { return $parsed }
        return $null
    }
    return Convert-ToDateTimeOrNull $Value
}

function Get-WorkloadUsageState {
    param(
        [AllowNull()][object]$ActivityRow,
        [ValidateSet('Exchange', 'Teams', 'SharePoint', 'OneDrive')][string]$Workload
    )
    if ($null -eq $ActivityRow -or (Get-NormalizedKey (Get-PropertyValue $ActivityRow 'IsDeleted')) -eq 'true') {
        return 'Unknown — no current report row'
    }
    $licensed = Get-NormalizedKey (Get-PropertyValue $ActivityRow "Has${Workload}License")
    if ($licensed -ne 'true') { return 'Not entitled' }
    $reportDate = Convert-ToDateTimeOrNull (Get-PropertyValue $ActivityRow 'ReportRefreshDate')
    if ($null -eq $reportDate) { return 'Unknown — missing report date' }
    $lastActivity = Convert-ToDateTimeOrNull (Get-PropertyValue $ActivityRow "${Workload}LastActivityDate")
    if ($null -ne $lastActivity -and $lastActivity.Date -ge $reportDate.Date.AddDays(-30)) { return 'Active in 30D' }
    return 'No activity in 30D'
}

function Convert-ToIntOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = 0
    if ([int]::TryParse([string]$Value, [Globalization.NumberStyles]::Integer, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Convert-ToCountryName {
    param(
        [AllowNull()][object]$Country,
        [AllowNull()][object]$UsageLocation
    )

    $countryText = ([string]$Country).Trim()
    if ($countryText) { return $countryText }

    $usageText = ([string]$UsageLocation).Trim().ToUpperInvariant()
    if (-not $usageText) { return 'Unknown' }
    try {
        return ([Globalization.RegionInfo]::new($usageText)).EnglishName
    } catch {
        return $usageText
    }
}

function Get-UniqueIndex {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )

    $groups = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey (Get-PropertyValue $row $PropertyName)
        if (-not $key) { continue }
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [System.Collections.Generic.List[object]]::new()
        }
        $groups[$key].Add($row)
    }

    $index = @{}
    foreach ($entry in $groups.GetEnumerator()) {
        if ($entry.Value.Count -eq 1) { $index[$entry.Key] = $entry.Value[0] }
    }
    return $index
}

function Get-MultiIndex {
    param(
        [object[]]$Rows,
        [string]$PropertyName
    )

    $index = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey (Get-PropertyValue $row $PropertyName)
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key)) {
            $index[$key] = [System.Collections.Generic.List[object]]::new()
        }
        $index[$key].Add($row)
    }
    return $index
}

function Get-LatestIndex {
    param(
        [object[]]$Rows,
        [string]$KeyProperty,
        [string]$DateProperty
    )

    $index = @{}
    foreach ($row in $Rows) {
        $key = Get-NormalizedKey (Get-PropertyValue $row $KeyProperty)
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key)) {
            $index[$key] = $row
            continue
        }

        $candidateDate = Convert-ToDateTimeOrNull (Get-PropertyValue $row $DateProperty)
        $currentDate = Convert-ToDateTimeOrNull (Get-PropertyValue $index[$key] $DateProperty)
        if ($null -ne $candidateDate -and ($null -eq $currentDate -or $candidateDate -gt $currentDate)) {
            $index[$key] = $row
        }
    }
    return $index
}

function Import-SelectedCsvColumns {
    param(
        [string]$Path,
        [string[]]$Columns
    )

    $sample = Import-Csv -LiteralPath $Path | Select-Object -First 1
    $headers = $sample.PSObject.Properties.Name
    foreach ($column in $Columns) {
        if ($column -notin $headers) { throw "Required column '$column' was not found in $Path" }
    }
    # Import-Csv performs the parsing in the compiled cmdlet and streams rows
    # through Select-Object. This is substantially faster than a PowerShell loop
    # over TextFieldParser while retaining bounded memory for wide AD exports.
    return @(Import-Csv -LiteralPath $Path | Select-Object -Property $Columns)
}

function Get-ActivityClassification {
    param(
        [AllowNull()][object]$ActivityRow,
        [AllowNull()][object]$AdRow,
        [datetime]$AsOfDate
    )

    $m365LastActivity = if ($null -ne $ActivityRow) { Convert-ToDateTimeOrNull (Get-PropertyValue $ActivityRow 'LastActivityDate') } else { $null }
    $adLastActivity = if ($null -ne $AdRow) { Convert-AdLogonDateTimeOrNull (Get-PropertyValue $AdRow 'LastLogonDate') } else { $null }
    $observedDates = @(@($m365LastActivity, $adLastActivity) | Where-Object { $null -ne $_ })
    $lastActivity = if ($observedDates.Count -gt 0) { $observedDates | Sort-Object -Descending | Select-Object -First 1 } else { $null }
    $days = if ($null -ne $lastActivity) { [math]::Max(0, [int][math]::Floor(($AsOfDate.Date - $lastActivity.Date).TotalDays)) } else { $null }

    if ($null -ne $m365LastActivity -and $null -ne $adLastActivity) {
        $activitySource = 'Active Directory + Microsoft 365'
    } elseif ($null -ne $adLastActivity) {
        $activitySource = 'Active Directory'
    } elseif ($null -ne $m365LastActivity) {
        $activitySource = 'Microsoft 365'
    } else {
        $activitySource = 'None'
    }

    if ($null -ne $days -and $days -le 30) {
        $state = 'Active <=30D'
        $workforceStatus = 'Active'
        $evidenceStatus = 'Observed'
    } elseif ($null -ne $days -and $days -le 90) {
        $state = 'Dormant 31-90D'
        $workforceStatus = 'Inactive'
        $evidenceStatus = 'Observed'
    } elseif ($null -ne $days) {
        $state = 'Stale >90D'
        $workforceStatus = 'Inactive'
        $evidenceStatus = 'Observed'
    } elseif ($null -ne $ActivityRow -and (Get-NormalizedKey (Get-PropertyValue $ActivityRow 'HasAnyM365Activity')) -eq 'false') {
        $state = 'Observed never used'
        $workforceStatus = 'Inactive'
        $evidenceStatus = 'Observed'
    } else {
        $state = 'No activity evidence'
        $workforceStatus = 'Unknown'
        $evidenceStatus = 'Unknown'
    }

    return [pscustomobject]@{
        ActivityState = $state
        WorkforceStatus = $workforceStatus
        EvidenceStatus = $evidenceStatus
        DaysSinceLastActivity = $days
        LastActivityDate = $lastActivity
        AdLastActivityDate = $adLastActivity
        M365LastActivityDate = $m365LastActivity
        ActivitySource = $activitySource
    }
}

function Get-WeekLabelFromPath {
    param([string]$Path)
    $match = [regex]::Match($Path, 'WeeklyHistory[\\/](?<Week>\d{4}-W\d{2})[\\/]')
    if (-not $match.Success) { return $null }
    return $match.Groups['Week'].Value
}

function Write-CsvAtomically {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $temporaryPath = "$Path.tmp"
    $Rows | Export-Csv -LiteralPath $temporaryPath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
}

$activeUsersPath = Get-RequiredCsv 'DATA-LAST\M365_Users_Active.csv'
$activityPath = Get-RequiredCsv 'DATA-LAST\M365_Users_Activity.csv'
$adUsersPath = Get-RequiredCsv 'DATA-LAST\AD_Users_AllDomains.csv'
$duplicateUpnPath = Get-RequiredCsv 'DATA-LAST\AD_Users_DuplicateUPN.csv'
$hybridIssuesPath = Get-RequiredCsv 'DATA-LAST\Exchange_HybridIdentity_Issues.csv'
$licenseUsersPath = Get-RequiredCsv 'DATA-LAST\M365_Licenses_Users.csv'
$officeActivationsPath = Get-RequiredCsv 'DATA-LAST\M365_Apps_Activations.csv'
$officeUsagePath = Get-RequiredCsv 'DATA-LAST\M365_Apps_Usage_30D.csv'
$authenticationMethodsPath = Get-RequiredCsv 'DATA-LAST\M365_Entra_AuthenticationMethodsRegistration.csv'
if (-not (Test-Path -LiteralPath $PersonaClassificationPath -PathType Leaf)) { throw "Persona classification workbook not found: $PersonaClassificationPath" }
if (-not (Test-Path -LiteralPath $SiteClassificationPath -PathType Leaf)) { throw "Site classification workbook not found: $SiteClassificationPath" }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) { throw 'ImportExcel PowerShell module is required to read the persona classification workbook.' }

Write-Verbose 'Loading current identity and activity sources.'
$entraUsers = @(Import-Csv -LiteralPath $activeUsersPath)
$activityRows = @(Import-Csv -LiteralPath $activityPath)
$adUsers = @(Import-Csv -LiteralPath $adUsersPath)
$siteSettings = @{}
foreach ($row in @(Import-Excel -Path $SiteClassificationPath -WorksheetName 'Settings')) {
    $siteSettings[([string]$row.Parameter).Trim()] = ([string]$row.Value).Trim()
}
$siteAttribute = $siteSettings['Directory Site Attribute']
$headquartersPersona = $siteSettings['Headquarters Persona']
if (-not $siteAttribute -or -not $adUsers[0].PSObject.Properties[$siteAttribute]) {
    throw "Configured site attribute is not present in AD user evidence: $siteAttribute"
}
if (-not $headquartersPersona) { throw 'Headquarters Persona is missing from site Settings.' }
$allowedSiteTypes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in @(Import-Excel -Path $SiteClassificationPath -WorksheetName 'SiteTypes')) {
    $allowedSiteType = ([string]$row.'Site Type Code').Trim()
    if ($allowedSiteType) { [void]$allowedSiteTypes.Add($allowedSiteType) }
}
if (-not $allowedSiteTypes.Contains('HQ') -or -not $allowedSiteTypes.Contains('UNKNOWN')) {
    throw 'SiteTypes must include HQ and UNKNOWN.'
}
$siteByCode = @{}
foreach ($row in @(Import-Excel -Path $SiteClassificationPath -WorksheetName 'Sites')) {
    if (([string]$row.'Use for Classification').Trim() -ne 'Yes') { continue }
    $code = ([string]$row.'Directory Site Code').Trim()
    $siteType = ([string]$row.'Site Type').Trim().ToUpperInvariant()
    if (-not $code -or $siteByCode.ContainsKey($code)) { throw "Blank or duplicate enabled Directory Site Code: $code" }
    if (-not $allowedSiteTypes.Contains($siteType)) { throw "Unsupported Site Type for $code`: $siteType" }
    $siteByCode[$code] = [pscustomobject]@{ Type = $siteType; Name = ([string]$row.'Site Display Name').Trim() }
}
$duplicateUpnRows = @(Import-Csv -LiteralPath $duplicateUpnPath)
$hybridIssueRows = @(Import-Csv -LiteralPath $hybridIssuesPath)
$licenseRows = @(Import-Csv -LiteralPath $licenseUsersPath)
$officeActivationRows = @(Import-Csv -LiteralPath $officeActivationsPath)
$officeUsageByUpn = @{}
foreach ($row in @(Import-Csv -LiteralPath $officeUsagePath)) {
    if (([string](Get-PropertyValue $row 'ReportPeriodRequested')).Trim() -ne 'D30') { throw 'M365_Apps_Usage_30D.csv does not contain the expected D30 report period.' }
    $usageUpn = Get-NormalizedKey (Get-PropertyValue $row 'User Principal Name')
    $usageDate = Convert-ToDateTimeOrNull (Get-PropertyValue $row 'Report Refresh Date')
    if (-not $usageUpn -or $null -eq $usageDate -or $usageDate.Date -lt (Get-Date).Date.AddDays(-10)) { continue }
    if (-not $officeUsageByUpn.ContainsKey($usageUpn) -or $usageDate -gt $officeUsageByUpn[$usageUpn].ReportDate) {
        $officeUsageByUpn[$usageUpn] = [pscustomobject]@{ ReportDate = $usageDate; Row = $row }
    }
}
$authenticationMethods = @(Import-Csv -LiteralPath $authenticationMethodsPath)
$authenticationByUserId = Get-UniqueIndex $authenticationMethods 'UserId'
$officeActivationByUpn = @{}
foreach ($row in $officeActivationRows) {
    if (([string](Get-PropertyValue $row 'Product Type')).Trim() -ne 'MICROSOFT 365 APPS FOR ENTERPRISE') { continue }
    $activationUpn = Get-NormalizedKey (Get-PropertyValue $row 'User Principal Name')
    if (-not $activationUpn) { continue }
    $reportDate = Convert-ToDateTimeOrNull (Get-PropertyValue $row 'Report Refresh Date')
    if ($null -eq $reportDate) { continue }
    $existing = if ($officeActivationByUpn.ContainsKey($activationUpn)) { $officeActivationByUpn[$activationUpn] } else { $null }
    if ($null -eq $existing -or $reportDate -gt $existing.ReportDate) {
        $officeActivationByUpn[$activationUpn] = [pscustomobject]@{ ReportDate = $reportDate; LastActivationDate = Convert-ToDateTimeOrNull (Get-PropertyValue $row 'Last Activated Date'); PcCount = [int](Get-PropertyValue $row 'Windows') + [int](Get-PropertyValue $row 'Mac') }
    }
}
$personaCountryCodes = @{ France = 'FR'; Germany = 'DE'; Spain = 'ES'; Belgium = 'BE'; Poland = 'PL'; Switzerland = 'CH'; Portugal = 'PT'; Italy = 'IT'; Austria = 'AT'; Ireland = 'IE'; Luxembourg = 'LU'; Netherlands = 'NL' }
$personaCache = @{}
$personaTextCache = @{}
$personaNames = @{}
foreach ($row in @(Import-Excel -Path $PersonaClassificationPath -WorksheetName 'Personas')) {
    if (([string]$row.Enabled).Trim() -ne 'Yes') { continue }
    $personaId = ([string]$row.'Persona ID').Trim()
    $personaName = ([string]$row.'Persona name').Trim()
    if (-not $personaId -or -not $personaName -or $personaNames.ContainsKey($personaId)) { throw "Invalid or duplicate persona ID: $personaId" }
    $personaNames[$personaId] = $personaName
}
if ($personaNames.Count -eq 0) {
    throw 'No enabled personas found. Copy and complete the blank persona classification template in a private data folder.'
}
if ($personaNames.Values -notcontains $headquartersPersona) { throw "Headquarters Persona is not enabled in the persona workbook: $headquartersPersona" }
$personaKeywords = @(Import-Excel -Path $PersonaClassificationPath -WorksheetName 'Rules' | ForEach-Object {
    if (([string]$_.Enabled).Trim() -ne 'Yes') { return }
    $keyword = ([string]$_.Keyword).Trim()
    if (-not $keyword) { return }
    $personaId = ([string]$_.'Persona ID').Trim()
    if (-not $personaNames.ContainsKey($personaId)) { throw "Unknown persona ID in rule $($_.'Rule ID'): $personaId" }
    if (([string]$_.'Match type').Trim() -ne 'Contains phrase') { throw "Unsupported match type in rule $($_.'Rule ID')" }
    [pscustomobject]@{
        Keyword = $keyword
        SearchKeyword = Convert-ToSearchText $keyword
        Category = $personaNames[$personaId]
        Country = ([string]$_.Country).Trim().ToUpperInvariant()
        Priority = [int]$_.Priority
    }
})
if ($personaKeywords.Count -eq 0) {
    throw 'No enabled persona rules found. Add and review rules before generating workforce evidence.'
}
$personaExclusions = @(Import-Excel -Path $PersonaClassificationPath -WorksheetName 'Exclusions' | ForEach-Object {
    if (([string]$_.Enabled).Trim() -ne 'Yes') { return }
    $keyword = ([string]$_.Keyword).Trim()
    if (-not $keyword) { return }
    [pscustomobject]@{
        Keyword = $keyword
        SearchKeyword = Convert-ToSearchText $keyword
        Country = ([string]$_.Country).Trim().ToUpperInvariant()
    }
})

$enabledMemberUsers = @($entraUsers | Where-Object {
    (Get-NormalizedKey (Get-PropertyValue $_ 'AccountEnabled')) -eq 'true' -and
    (Get-NormalizedKey (Get-PropertyValue $_ 'UserType')) -eq 'member'
})
$activityByUpn = Get-LatestIndex $activityRows 'UserPrincipalName' 'ReportRefreshDate'
$adByDn = Get-UniqueIndex $adUsers 'DistinguishedName'
$entraBySid = Get-MultiIndex $entraUsers 'OnPremisesSecurityIdentifier'
$entraByImmutableId = Get-MultiIndex $entraUsers 'OnPremisesImmutableId'
$entraByUpn = Get-MultiIndex $entraUsers 'User principal name'

# Resolve the governed account population through strong identity keys. A
# conflicting classification or an unmatched cloud-only account is kept in the
# review-required population instead of being silently treated as human.
$accountPopulationByEntraId = @{}
$accountTypeByEntraId = @{}
$matchedAnyAdEntraIds = [System.Collections.Generic.HashSet[string]]::new()
$explicitNonHumanEntraIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($adUser in $adUsers) {
    $accountType = ([string](Get-PropertyValue $adUser 'AccountType')).Trim()
    if (-not $accountType) { $accountType = 'Unclassified Account' }
    $accountPopulation = Get-ConfiguredAccountPopulation $accountType

    foreach ($lookup in @(
        [pscustomobject]@{ Index = $entraBySid; Value = (Get-PropertyValue $adUser 'ObjectSID') },
        [pscustomobject]@{ Index = $entraByImmutableId; Value = (Get-PropertyValue $adUser 'ImmutableId_AD') },
        [pscustomobject]@{ Index = $entraByUpn; Value = (Get-PropertyValue $adUser 'UserPrincipalName') }
    )) {
        $key = Get-NormalizedKey $lookup.Value
        if (-not $key -or -not $lookup.Index.ContainsKey($key)) { continue }
        foreach ($candidate in $lookup.Index[$key]) {
            $candidateId = Get-NormalizedKey (Get-PropertyValue $candidate 'Object Id')
            if (-not $candidateId) { continue }
            [void]$matchedAnyAdEntraIds.Add($candidateId)
            if ($accountPopulation -eq 'Non-human') { [void]$explicitNonHumanEntraIds.Add($candidateId) }
            if ($accountPopulationByEntraId.ContainsKey($candidateId) -and $accountPopulationByEntraId[$candidateId] -ne $accountPopulation) {
                $accountPopulationByEntraId[$candidateId] = 'Review required'
                $accountTypeByEntraId[$candidateId] = 'Conflicting AD classifications'
            } elseif (-not $accountPopulationByEntraId.ContainsKey($candidateId)) {
                $accountPopulationByEntraId[$candidateId] = $accountPopulation
                $accountTypeByEntraId[$candidateId] = $accountType
            }
        }
    }
}

$eligibleAdUsers = @($adUsers | Where-Object {
    (Get-NormalizedKey (Get-PropertyValue $_ 'Enabled')) -eq 'true' -and
    (Get-ConfiguredAccountPopulation (Get-PropertyValue $_ 'AccountType')) -ne 'Non-human'
})
$eligibleAdByUpn = Get-UniqueIndex $eligibleAdUsers 'UserPrincipalName'

$populationCounts = [ordered]@{ Human = 0; 'Non-human' = 0; 'Review required' = 0 }
$workforceUserList = [System.Collections.Generic.List[object]]::new()
$confirmedHumanEntraIds = [System.Collections.Generic.HashSet[string]]::new()
$confirmedHumanUpns = [System.Collections.Generic.HashSet[string]]::new()
$workforceEntraIds = [System.Collections.Generic.HashSet[string]]::new()
$workforceUpns = [System.Collections.Generic.HashSet[string]]::new()
foreach ($memberUser in $enabledMemberUsers) {
    $candidateId = Get-NormalizedKey (Get-PropertyValue $memberUser 'Object Id')
    $candidateUpn = Get-NormalizedKey (Get-PropertyValue $memberUser 'User principal name')
    $population = if ($candidateId -and $accountPopulationByEntraId.ContainsKey($candidateId)) { $accountPopulationByEntraId[$candidateId] } else { 'Review required' }
    $populationCounts[$population]++
    if ($population -eq 'Human') {
        if ($candidateId) { [void]$confirmedHumanEntraIds.Add($candidateId) }
        if ($candidateUpn) { [void]$confirmedHumanUpns.Add($candidateUpn) }
    }
    # Enabled workforce accounts must be linked to AD. Only explicit service,
    # shared-mailbox, room-mailbox and system populations are excluded.
    if (-not $candidateId -or -not $matchedAnyAdEntraIds.Contains($candidateId) -or $explicitNonHumanEntraIds.Contains($candidateId)) { continue }
    $workforceUserList.Add($memberUser)
    [void]$workforceEntraIds.Add($candidateId)
    if ($candidateUpn) { [void]$workforceUpns.Add($candidateUpn) }
}
$enabledWorkforceUsers = @($workforceUserList)
$excludedNonHumanAccounts = $populationCounts['Non-human']
$reviewRequiredAccounts = $populationCounts['Review required']

$licenseSkuByUpn = @{}
$m365PlanByUpn = @{}
foreach ($row in $licenseRows) {
    $upn = Get-NormalizedKey (Get-PropertyValue $row 'User principal name')
    if (-not $upn) { continue }
    if (-not $licenseSkuByUpn.ContainsKey($upn)) {
        $licenseSkuByUpn[$upn] = [System.Collections.Generic.HashSet[string]]::new()
    }
    $sku = Get-NormalizedKey (Get-PropertyValue $row 'SkuId')
    if (-not $sku) { $sku = Get-NormalizedKey (Get-PropertyValue $row 'SkuPartNumber') }
    if ($sku) { [void]$licenseSkuByUpn[$upn].Add($sku) }
    $partNumber = (Get-NormalizedKey (Get-PropertyValue $row 'SkuPartNumber')).ToUpperInvariant()
    if ($partNumber -in @('SPE_F1', 'SPE_E3', 'SPE_E5')) {
        if (-not $m365PlanByUpn.ContainsKey($upn)) { $m365PlanByUpn[$upn] = [System.Collections.Generic.HashSet[string]]::new() }
        [void]$m365PlanByUpn[$upn].Add($partNumber)
    }
}

$matchedAdByEntraId = @{}
$matchMethodByEntraId = @{}
$conflictEntraIds = [System.Collections.Generic.HashSet[string]]::new()
$identityCounts = [ordered]@{ Matched = 0; Conflict = 0; Unmatched = 0 }

foreach ($adUser in $eligibleAdUsers) {
    $candidates = @{}
    $candidateMethods = @{}
    foreach ($lookup in @(
        [pscustomobject]@{ Method = 'SID'; Index = $entraBySid; Value = (Get-PropertyValue $adUser 'ObjectSID') },
        [pscustomobject]@{ Method = 'ImmutableId'; Index = $entraByImmutableId; Value = (Get-PropertyValue $adUser 'ImmutableId_AD') },
        [pscustomobject]@{ Method = 'UPN'; Index = $entraByUpn; Value = (Get-PropertyValue $adUser 'UserPrincipalName') }
    )) {
        $key = Get-NormalizedKey $lookup.Value
        if (-not $key -or -not $lookup.Index.ContainsKey($key)) { continue }
        foreach ($candidate in $lookup.Index[$key]) {
            $candidateId = Get-NormalizedKey (Get-PropertyValue $candidate 'Object Id')
            if (-not $candidateId) { continue }
            $candidates[$candidateId] = $candidate
            if (-not $candidateMethods.ContainsKey($candidateId)) {
                $candidateMethods[$candidateId] = [System.Collections.Generic.HashSet[string]]::new()
            }
            [void]$candidateMethods[$candidateId].Add($lookup.Method)
        }
    }

    if ($candidates.Count -eq 1) {
        $identityCounts.Matched++
        $candidateId = @($candidates.Keys)[0]
        $matchedAdByEntraId[$candidateId] = $adUser
        $methods = $candidateMethods[$candidateId]
        $matchMethodByEntraId[$candidateId] = if ($methods.Contains('SID')) { 'SID' } elseif ($methods.Contains('ImmutableId')) { 'ImmutableId' } else { 'UPN fallback' }
    } elseif ($candidates.Count -gt 1) {
        $identityCounts.Conflict++
        foreach ($candidateId in $candidates.Keys) { [void]$conflictEntraIds.Add($candidateId) }
    } else {
        $identityCounts.Unmatched++
    }
}

$duplicateUpns = [System.Collections.Generic.HashSet[string]]::new()
foreach ($row in $duplicateUpnRows) {
    $upn = Get-NormalizedKey (Get-PropertyValue $row 'UserPrincipalName')
    if ($upn) { [void]$duplicateUpns.Add($upn) }
}

$evidenceDate = (Get-Item -LiteralPath $activityPath).LastWriteTime.Date
$userOutput = [System.Collections.Generic.List[object]]::new()
$activityStateCounts = @{
    'Active <=30D' = 0
    'Dormant 31-90D' = 0
    'Stale >90D' = 0
    'Observed never used' = 0
    'No activity evidence' = 0
}

foreach ($user in $enabledWorkforceUsers) {
    $upn = Get-NormalizedKey (Get-PropertyValue $user 'User principal name')
    $objectId = Get-NormalizedKey (Get-PropertyValue $user 'Object Id')
    if (-not $objectId) { $objectId = $upn }
    $matchedAd = if ($matchedAdByEntraId.ContainsKey($objectId)) { $matchedAdByEntraId[$objectId] } elseif ($eligibleAdByUpn.ContainsKey($upn)) { $eligibleAdByUpn[$upn] } else { $null }
    $activity = if ($activityByUpn.ContainsKey($upn)) { $activityByUpn[$upn] } else { $null }
    $officeActivation = if ($officeActivationByUpn.ContainsKey($upn)) { $officeActivationByUpn[$upn] } else { $null }
    $officeUsage = if ($officeUsageByUpn.ContainsKey($upn)) { $officeUsageByUpn[$upn] } else { $null }
    $officeUsageState = Get-OfficeDesktopUsageState -Row $(if ($null -ne $officeUsage) { $officeUsage.Row } else { $null })
    $officeActivationState = if ($null -eq $officeActivation) { 'Unknown — no activation report row' } elseif ($officeActivation.PcCount -gt 0 -and $null -ne $officeActivation.LastActivationDate -and $officeActivation.LastActivationDate.Date -ge $officeActivation.ReportDate.Date.AddDays(-30) -and $officeActivation.LastActivationDate.Date -le $officeActivation.ReportDate.Date) { 'Activated on PC in 30D' } else { 'No PC activation in 30D' }
    $classification = Get-ActivityClassification -ActivityRow $activity -AdRow $matchedAd -AsOfDate $evidenceDate
    $activityStateCounts[$classification.ActivityState]++
    if ($conflictEntraIds.Contains($objectId)) {
        $identityMatchStatus = 'Conflict'
        $identityMatchMethod = 'Multiple candidates'
    } elseif ($null -ne $matchedAd) {
        $identityMatchStatus = 'Matched'
        $identityMatchMethod = $matchMethodByEntraId[$objectId]
    } else {
        $identityMatchStatus = 'Cloud only / not matched'
        $identityMatchMethod = 'Not applicable'
    }

    $department = ([string](Get-PropertyValue $user 'Department')).Trim()
    if (-not $department -and $null -ne $matchedAd) { $department = ([string](Get-PropertyValue $matchedAd 'Department')).Trim() }
    $departmentState = if ($department) { 'Observed' } else { 'Missing' }
    if (-not $department) { $department = 'Unknown' }

    $jobTitle = ([string](Get-PropertyValue $user 'Title')).Trim()
    if (-not $jobTitle -and $null -ne $matchedAd) { $jobTitle = ([string](Get-PropertyValue $matchedAd 'Title')).Trim() }
    if (-not $jobTitle -and $null -ne $matchedAd) { $jobTitle = ([string](Get-PropertyValue $matchedAd 'JobTitle')).Trim() }
    $jobTitleState = if ($jobTitle) { 'Observed' } else { 'Missing' }
    if (-not $jobTitle) { $jobTitle = 'Unknown' }
    $country = Convert-ToCountryName (Get-PropertyValue $user 'CountryOrRegion') (Get-PropertyValue $user 'Usage location')
    $siteCode = ([string](Get-PropertyValue $matchedAd $siteAttribute)).Trim()
    $site = if ($siteCode -and $siteByCode.ContainsKey($siteCode)) { $siteByCode[$siteCode] } else { $null }
    $siteType = if ($null -ne $site) { $site.Type } else { 'UNKNOWN' }
    $siteMatchState = if (-not $siteCode) { 'No AD site code' } elseif ($null -eq $site) { 'Unmapped site code' } else { 'Mapped' }
    $adDescription = ([string](Get-PropertyValue $matchedAd 'Description')).Trim()
    $persona = Get-PersonaClassification -JobTitle $jobTitle -Description $adDescription -Country $country -SiteType $siteType
    $authentication = if ($authenticationByUserId.ContainsKey($objectId)) { $authenticationByUserId[$objectId] } else { $null }
    $mfaRegistrationState = if ($null -eq $authentication) { 'No evidence' } elseif ((Get-NormalizedKey $authentication.IsMfaRegistered) -eq 'true') { 'Registered' } else { 'Not registered' }

    if ($null -eq $matchedAd) {
        $manager = 'Not observed'
        $managerState = 'Not observed'
    } else {
        $managerDn = Get-NormalizedKey (Get-PropertyValue $matchedAd 'manager')
        if (-not $managerDn) {
            $manager = 'Missing'
            $managerState = 'Missing'
        } elseif ($adByDn.ContainsKey($managerDn)) {
            $managerRow = $adByDn[$managerDn]
            $manager = ([string](Get-PropertyValue $managerRow 'DisplayName')).Trim()
            if (-not $manager) { $manager = ([string](Get-PropertyValue $managerRow 'UserPrincipalName')).Trim() }
            if (-not $manager) { $manager = 'Observed in AD' }
            $managerState = 'Observed'
        } else {
            $manager = 'Observed in AD'
            $managerState = 'Observed'
        }
    }

    $licenseCount = if ($licenseSkuByUpn.ContainsKey($upn)) { $licenseSkuByUpn[$upn].Count } else { 0 }
    $m365Plans = [System.Collections.Generic.HashSet[string]]::new()
    if ($m365PlanByUpn.ContainsKey($upn)) { $m365Plans = $m365PlanByUpn[$upn] }
    $attentionPriority = switch ($classification.ActivityState) {
        'Observed never used' { 'High' }
        'Stale >90D' { 'High' }
        'Dormant 31-90D' { 'Medium' }
        'No activity evidence' { 'Evidence gap' }
        default { 'Healthy' }
    }
    if ($identityMatchStatus -eq 'Conflict') { $attentionPriority = 'Critical' }
    if ($duplicateUpns.Contains($upn)) { $attentionPriority = 'Critical' }

    $userOutput.Add([pscustomobject][ordered]@{
        'User Source ID' = $objectId
        'User Principal Name' = ([string](Get-PropertyValue $user 'User principal name')).Trim()
        'Display Name' = ([string](Get-PropertyValue $user 'Display name')).Trim()
        'Country' = $country
        'Department' = $department
        'Department Evidence State' = $departmentState
        'Job Title' = $jobTitle
        'Job Title Evidence State' = $jobTitleState
        'Directory Site Code' = $siteCode
        'Site Type' = $siteType
        'Site Match State' = $siteMatchState
        'Workforce Persona' = $persona.Persona
        'Persona Match State' = $persona.State
        'Persona Matched Keywords' = $persona.Keyword
        'Persona Classification Source' = $persona.Source
        'MFA Registration State' = $mfaRegistrationState
        'Manager' = $manager
        'Manager Evidence State' = $managerState
        'Account State' = 'Enabled'
        'User Type' = 'Member'
        'Account Type' = if ($accountTypeByEntraId.ContainsKey($objectId)) { $accountTypeByEntraId[$objectId] } else { 'Unclassified Account' }
        'Account Population' = if ($confirmedHumanEntraIds.Contains($objectId)) { 'Human' } else { 'Workforce - review classification' }
        'Account Classification Rule Version' = $accountClassificationRuleVersion
        'Workforce Status' = $classification.WorkforceStatus
        'Activity State' = $classification.ActivityState
        'Activity Evidence Status' = $classification.EvidenceStatus
        'Days Since Last Activity' = $classification.DaysSinceLastActivity
        'Last Activity Date' = if ($null -eq $classification.LastActivityDate) { $null } else { $classification.LastActivityDate.ToString('yyyy-MM-dd') }
        'AD Last Activity Date' = if ($null -eq $classification.AdLastActivityDate) { $null } else { $classification.AdLastActivityDate.ToString('yyyy-MM-dd') }
        'M365 Last Activity Date' = if ($null -eq $classification.M365LastActivityDate) { $null } else { $classification.M365LastActivityDate.ToString('yyyy-MM-dd') }
        'Activity Source' = $classification.ActivitySource
        'Last Activity Workload' = if ($null -ne $classification.AdLastActivityDate -and ($null -eq $classification.M365LastActivityDate -or $classification.AdLastActivityDate -gt $classification.M365LastActivityDate)) { 'Active Directory' } elseif ($null -eq $activity) { $null } else { Get-PropertyValue $activity 'LastActivityWorkload' }
        'Identity Match Status' = $identityMatchStatus
        'Identity Match Method' = $identityMatchMethod
        'M365 License Count' = $licenseCount
        'Has M365 License' = if ($licenseCount -gt 0) { 'Yes' } else { 'No' }
        'Has Microsoft 365 F3' = if ($m365Plans.Contains('SPE_F1')) { 'Yes' } else { 'No' }
        'Has Microsoft 365 E3' = if ($m365Plans.Contains('SPE_E3')) { 'Yes' } else { 'No' }
        'Has Microsoft 365 E5' = if ($m365Plans.Contains('SPE_E5')) { 'Yes' } else { 'No' }
        'Exchange Usage State (30D)' = Get-WorkloadUsageState -ActivityRow $activity -Workload 'Exchange'
        'Teams Usage State (30D)' = Get-WorkloadUsageState -ActivityRow $activity -Workload 'Teams'
        'SharePoint Usage State (30D)' = Get-WorkloadUsageState -ActivityRow $activity -Workload 'SharePoint'
        'OneDrive Usage State (30D)' = Get-WorkloadUsageState -ActivityRow $activity -Workload 'OneDrive'
        'Office Desktop Usage State (30D)' = $officeUsageState
        'Office Desktop Usage Report Date' = if ($null -eq $officeUsage) { '' } else { $officeUsage.ReportDate.ToString('yyyy-MM-dd') }
        'Office Desktop Activation State (30D)' = $officeActivationState
        'M365 Usage Report Date' = if ($null -eq $activity) { '' } else { [string](Get-PropertyValue $activity 'ReportRefreshDate') }
        'Duplicate UPN State' = if ($duplicateUpns.Contains($upn)) { 'Duplicate' } else { 'Unique' }
        'Attention Priority' = $attentionPriority
        'Evidence Date' = $evidenceDate.ToString('yyyy-MM-dd')
    })
}

$userOutputSorted = @($userOutput | Sort-Object 'User Principal Name')
$officeUsageCovered = @($userOutputSorted | Where-Object { $_.'Office Desktop Usage State (30D)' -in @('Used on PC in 30D', 'No PC app use in 30D') }).Count
if ($officeUsageCovered -eq 0) { throw 'The D30 Apps usage report has no usable, non-anonymized workforce matches. Check report refresh, UPN privacy settings and source schema before generating evidence.' }
Write-CsvAtomically $userOutputSorted $UserOutputPath

$identityRows = @(
    [pscustomobject][ordered]@{ 'Evidence Date' = $evidenceDate.ToString('yyyy-MM-dd'); 'Match Status' = 'Matched'; 'Match Method' = 'SID first'; 'Identities' = $identityCounts.Matched; 'Eligible Identities' = $eligibleAdUsers.Count; 'Evidence Status' = 'Observed'; 'Recommended Action' = 'No action required for strongly matched identities.' },
    [pscustomobject][ordered]@{ 'Evidence Date' = $evidenceDate.ToString('yyyy-MM-dd'); 'Match Status' = 'Conflict'; 'Match Method' = 'Multiple strong-key candidates'; 'Identities' = $identityCounts.Conflict; 'Eligible Identities' = $eligibleAdUsers.Count; 'Evidence Status' = 'Observed'; 'Recommended Action' = 'Resolve conflicting SID, immutable ID, or UPN mappings.' },
    [pscustomobject][ordered]@{ 'Evidence Date' = $evidenceDate.ToString('yyyy-MM-dd'); 'Match Status' = 'Unmatched'; 'Match Method' = 'No accepted candidate'; 'Identities' = $identityCounts.Unmatched; 'Eligible Identities' = $eligibleAdUsers.Count; 'Evidence Status' = 'Observed'; 'Recommended Action' = 'Investigate directory synchronization and source identity keys.' }
)
Write-CsvAtomically $identityRows $IdentityOutputPath

$coveredUsers = $enabledWorkforceUsers.Count - $activityStateCounts['No activity evidence']
$inactiveOver30 = $activityStateCounts['Dormant 31-90D'] + $activityStateCounts['Stale >90D'] + $activityStateCounts['Observed never used']
$missingManager = @($eligibleAdUsers | Where-Object { [string]::IsNullOrWhiteSpace([string](Get-PropertyValue $_ 'manager')) }).Count
$missingDepartment = @($userOutputSorted | Where-Object 'Department Evidence State' -eq 'Missing').Count
$missingJobTitle = @($userOutputSorted | Where-Object 'Job Title Evidence State' -eq 'Missing').Count
$unmatchedPersona = @($userOutputSorted | Where-Object 'Persona Match State' -eq 'Unmatched').Count
$ambiguousPersona = @($userOutputSorted | Where-Object 'Persona Match State' -eq 'Ambiguous').Count
$potentialNonHumanPersona = @($userOutputSorted | Where-Object 'Persona Match State' -eq 'Non-human keyword').Count
$disabledLicensedGoverned = @($hybridIssueRows | Where-Object { (Get-NormalizedKey (Get-PropertyValue $_ 'Potential_Issue')) -eq 'disabled account with active m365 license' }).Count

function ConvertTo-InvariantDecimalString {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.######', [System.Globalization.CultureInfo]::InvariantCulture)
}

function New-SignalRow {
    param(
        [string]$Severity,
        [string]$Signal,
        [string]$RiskCategory,
        [int]$AffectedUsers,
        [AllowNull()][object]$CoveredPopulation,
        [string]$EvidenceStatus,
        [string]$RecommendedAction,
        [string]$SourceScope
    )
    $coverage = if ($null -eq $CoveredPopulation -or [int]$CoveredPopulation -le 0) { $null } else { [math]::Round($AffectedUsers / [double]$CoveredPopulation, 6) }
    return [pscustomobject][ordered]@{
        'Severity Label' = $Severity
        'Signal' = $Signal
        'Risk Category' = $RiskCategory
        'Affected Users' = $AffectedUsers
        'Covered Population' = $CoveredPopulation
        'Affected Estate (%)' = ConvertTo-InvariantDecimalString $coverage
        'Evidence Status' = $EvidenceStatus
        'Recommended Action' = $RecommendedAction
        'Source Scope' = $SourceScope
    }
}

$signalRows = @(
    New-SignalRow 'High' 'Accounts without activity for more than 30 days' 'Activity hygiene' $inactiveOver30 $coveredUsers 'Observed' 'Review ownership, disable obsolete accounts, and reclaim licenses where appropriate.' 'Combined AD and Microsoft 365 activity for enabled workforce accounts'
    New-SignalRow 'High' 'Observed never-used accounts' 'Activity hygiene' $activityStateCounts['Observed never used'] $coveredUsers 'Observed' 'Confirm ownership and business need before disabling or licensing changes.' 'Combined AD and Microsoft 365 activity for enabled workforce accounts'
    New-SignalRow 'High' 'Observed activity older than 90 days' 'Activity hygiene' $activityStateCounts['Stale >90D'] $coveredUsers 'Observed' 'Validate employment and service need, then remediate stale access.' 'Combined AD and Microsoft 365 activity for enabled workforce accounts'
    New-SignalRow 'High' 'Disabled accounts with active M365 license' 'License hygiene' $disabledLicensedGoverned $null 'Governed issue export' 'Remove or reassign licenses after validating exception scope.' 'Exchange Hybrid Identity governed rules'
    New-SignalRow 'Critical' 'Identity reconciliation conflicts' 'Identity integrity' $identityCounts.Conflict $eligibleAdUsers.Count 'Observed' 'Resolve conflicting SID, immutable ID, and UPN mappings before automation.' 'Enabled human-classified AD accounts'
    New-SignalRow 'High' 'Duplicate UPN rows' 'Identity integrity' $duplicateUpnRows.Count $eligibleAdUsers.Count 'Observed' 'Correct duplicate UPN values and rerun identity reconciliation.' 'Enabled human-classified AD accounts'
    New-SignalRow 'Medium' 'Users without manager' 'Ownership' $missingManager $eligibleAdUsers.Count 'Observed' 'Assign or correct the manager attribute for accountable access reviews.' 'Enabled human-classified AD accounts'
    New-SignalRow 'Medium' 'Missing activity evidence' 'Data quality' $activityStateCounts['No activity evidence'] $enabledWorkforceUsers.Count 'Partial coverage' 'Restore or validate AD and Microsoft 365 activity collection before classifying these accounts as inactive.' 'Enabled workforce accounts'
    New-SignalRow 'Medium' 'Missing department evidence' 'Data quality' $missingDepartment $enabledWorkforceUsers.Count 'Partial coverage' 'Improve department data before using it for primary segmentation or ownership reporting.' 'Enabled workforce accounts'
    New-SignalRow 'Medium' 'Missing job title evidence' 'Data quality' $missingJobTitle $enabledWorkforceUsers.Count 'Partial coverage' 'Improve job-title data before using it for role segmentation or licensing reviews.' 'Enabled workforce accounts'
    New-SignalRow 'Medium' 'Job titles without persona mapping' 'Persona mapping' $unmatchedPersona $enabledWorkforceUsers.Count 'Observed' 'Review unmatched job titles and enrich the governed keyword workbook where appropriate.' 'Enabled workforce accounts classified from current job title and country'
    New-SignalRow 'High' 'Ambiguous persona keyword matches' 'Persona mapping' $ambiguousPersona $enabledWorkforceUsers.Count 'Review required' 'Resolve overlapping persona keywords before assigning a definitive persona.' 'Enabled workforce accounts classified from current job title and country'
    New-SignalRow 'Medium' 'Potential non-human persona keywords' 'Persona mapping' $potentialNonHumanPersona $enabledWorkforceUsers.Count 'Review required' 'Validate scan-to-mail style titles; do not assign a human persona automatically.' 'Enabled workforce accounts classified from current job title and country'
    New-SignalRow 'Medium' 'Enabled non-human member accounts' 'Identity hygiene' $excludedNonHumanAccounts $enabledMemberUsers.Count "Observed via governed classification $accountClassificationRuleVersion" 'Review ownership, licensing, and workload-identity migration for service, shared-mailbox, room-mailbox, and system accounts.' 'Enabled Entra members matched to governed non-human AD account types'
    New-SignalRow 'High' 'Member accounts requiring classification' 'Data quality' $reviewRequiredAccounts $enabledMemberUsers.Count "Review required under classification $accountClassificationRuleVersion" 'Classify admin, generic, unclassified, conflicting, and cloud-only accounts before including them in workforce KPIs.' 'Enabled Entra members without a confirmed human or non-human classification'
)
Write-CsvAtomically $signalRows $SignalsOutputPath

if ($SkipHistory) {
    [pscustomobject]@{
        UserRows = $userOutputSorted.Count
        PersonaUnmatched = $unmatchedPersona
        PersonaAmbiguous = $ambiguousPersona
        PersonaPotentialNonHuman = $potentialNonHumanPersona
        HistoryUpdated = $false
    }
    return
}

$historyDirectory = Split-Path -Parent $HistoryOutputPath
if (-not (Test-Path -LiteralPath $historyDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $historyDirectory -Force | Out-Null
}
$historyTemporaryPath = "$HistoryOutputPath.tmp"
$resumeHistoryRows = if (Test-Path -LiteralPath $historyTemporaryPath) { @(Import-Csv -LiteralPath $historyTemporaryPath) } else { @() }
if ($resumeHistoryRows.Count -gt 0 -and (
    $null -eq $resumeHistoryRows[0].PSObject.Properties['User Principal Name'] -or
    $null -eq $resumeHistoryRows[0].PSObject.Properties['Confirmed Human'])) {
    throw 'Historical staging schema is obsolete. Use a new output directory; do not resume this temporary file.'
}

$activeWeeklyRoot = Join-Path $DataRoot 'DATA-ALL\M365\Users\ActiveUsers'
$activityWeeklyRoot = Join-Path $DataRoot 'DATA-ALL\M365\Usage'
$adWeeklyRoot = Join-Path $DataRoot 'DATA-ALL\ActiveDirectory\Inventory'
$activeWeeklyFiles = @(Get-ChildItem -LiteralPath $activeWeeklyRoot -Recurse -File -Filter 'M365_Users_Active.csv' | Where-Object { $_.FullName -match 'WeeklyHistory' })
$activityWeeklyFiles = @(Get-ChildItem -LiteralPath $activityWeeklyRoot -Recurse -File -Filter 'M365_Users_Activity.csv' | Where-Object { $_.FullName -match 'WeeklyHistory' })
$adWeeklyFiles = @(Get-ChildItem -LiteralPath $adWeeklyRoot -Recurse -File -Filter 'AD_Users_AllDomains.csv' | Where-Object { $_.FullName -match 'WeeklyHistory' })

$latestActiveByWeek = @{}
foreach ($file in $activeWeeklyFiles) {
    $week = Get-WeekLabelFromPath $file.FullName
    if (-not $week) { continue }
    if (-not $latestActiveByWeek.ContainsKey($week) -or $file.LastWriteTime -gt $latestActiveByWeek[$week].LastWriteTime) {
        $latestActiveByWeek[$week] = $file
    }
}
$latestActivityByWeek = @{}
foreach ($file in $activityWeeklyFiles) {
    $week = Get-WeekLabelFromPath $file.FullName
    if (-not $week) { continue }
    if (-not $latestActivityByWeek.ContainsKey($week) -or $file.LastWriteTime -gt $latestActivityByWeek[$week].LastWriteTime) {
        $latestActivityByWeek[$week] = $file
    }
}
$latestAdByWeek = @{}
foreach ($file in $adWeeklyFiles) {
    $week = Get-WeekLabelFromPath $file.FullName
    if (-not $week) { continue }
    if (-not $latestAdByWeek.ContainsKey($week) -or $file.LastWriteTime -gt $latestAdByWeek[$week].LastWriteTime) {
        $latestAdByWeek[$week] = $file
    }
}

$trendRows = [System.Collections.Generic.List[object]]::new()
$completedHistoryWeeks = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($group in @($resumeHistoryRows | Group-Object 'Week Label')) {
    $week = [string]$group.Name
    if (-not $week) { continue }
    [void]$completedHistoryWeeks.Add($week)
    $rows = @($group.Group)
    $snapshotDate = Convert-ToDateTimeOrNull $rows[0].'Snapshot Date'
    $covered = @($rows | Where-Object 'Activity Evidence Status' -eq 'Observed').Count
    $active = @($rows | Where-Object 'Activity State' -eq 'Active <=30D').Count
    $inactive = @($rows | Where-Object 'Activity State' -in @('Dormant 31-90D', 'Stale >90D', 'Observed never used')).Count
    $confirmed = @($rows | Where-Object {
        $id = Get-NormalizedKey $_.'User Source ID'
        $upn = Get-NormalizedKey $_.'User Principal Name'
        ($id -and $confirmedHumanEntraIds.Contains($id)) -or ($upn -and $confirmedHumanUpns.Contains($upn))
    }).Count
    $trendRows.Add([pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd')
        'Date Key' = [int]$snapshotDate.ToString('yyyyMMdd')
        'Week Label' = $week
        'Enabled Workforce Accounts' = $rows.Count
        'Confirmed Human Users' = $confirmed
        'Activity Covered Users' = $covered
        'Activity Evidence Coverage' = if ($rows.Count -eq 0) { $null } else { ConvertTo-InvariantDecimalString ([math]::Round($covered / [double]$rows.Count, 6)) }
        'Active Workforce' = $active
        'Accounts Without Activity >30D' = $inactive
        'Evidence Status' = 'Observed'
    })
}
$historyHeaderWritten = $null -ne $resumeHistoryRows -and @($resumeHistoryRows).Count -gt 0
foreach ($week in @($latestActiveByWeek.Keys | Sort-Object)) {
    if (-not $latestActivityByWeek.ContainsKey($week)) { continue }
    if ($completedHistoryWeeks.Contains($week)) { continue }
    Write-Verbose "Building user activity history for $week."
    $weekActiveRows = @(Import-SelectedCsvColumns -Path $latestActiveByWeek[$week].FullName -Columns @('Object Id', 'User principal name', 'AccountEnabled', 'UserType') | Where-Object {
        $candidateId = Get-NormalizedKey (Get-PropertyValue $_ 'Object Id')
        $candidateUpn = Get-NormalizedKey (Get-PropertyValue $_ 'User principal name')
        (Get-NormalizedKey (Get-PropertyValue $_ 'AccountEnabled')) -eq 'true' -and
        (Get-NormalizedKey (Get-PropertyValue $_ 'UserType')) -eq 'member' -and
        (($candidateId -and $workforceEntraIds.Contains($candidateId)) -or
         ($candidateUpn -and $workforceUpns.Contains($candidateUpn)))
    })
    $weekActivityRows = @(Import-SelectedCsvColumns -Path $latestActivityByWeek[$week].FullName -Columns @('UserPrincipalName', 'ReportRefreshDate', 'LastActivityDate', 'HasAnyM365Activity', 'LastActivityWorkload'))
    $weekActivityByUpn = Get-LatestIndex $weekActivityRows 'UserPrincipalName' 'ReportRefreshDate'
    # AD snapshots are wide enriched exports. Keep only the two columns required
    # for the historical activity calculation to bound memory usage.
    $weekAdRows = if ($latestAdByWeek.ContainsKey($week)) {
        @(Import-SelectedCsvColumns -Path $latestAdByWeek[$week].FullName -Columns @('UserPrincipalName', 'LastLogonDate'))
    } else { @() }
    $weekAdByUpn = Get-UniqueIndex $weekAdRows 'UserPrincipalName'
    $refreshDates = @($weekActivityRows | ForEach-Object { Convert-ToDateTimeOrNull (Get-PropertyValue $_ 'ReportRefreshDate') } | Where-Object { $null -ne $_ })
    $snapshotDate = if ($refreshDates.Count -gt 0) { ($refreshDates | Sort-Object -Descending | Select-Object -First 1).Date } else { $latestActivityByWeek[$week].LastWriteTime.Date }
    $weekRows = [System.Collections.Generic.List[object]]::new()
    $weekCounts = @{ Active = 0; InactiveOver30 = 0; Covered = 0 }

    foreach ($weekUser in $weekActiveRows) {
        $upn = Get-NormalizedKey (Get-PropertyValue $weekUser 'User principal name')
        $userId = Get-NormalizedKey (Get-PropertyValue $weekUser 'Object Id')
        if (-not $userId) { $userId = $upn }
        $weekActivity = if ($weekActivityByUpn.ContainsKey($upn)) { $weekActivityByUpn[$upn] } else { $null }
        $weekAd = if ($weekAdByUpn.ContainsKey($upn)) { $weekAdByUpn[$upn] } else { $null }
        $weekClassification = Get-ActivityClassification -ActivityRow $weekActivity -AdRow $weekAd -AsOfDate $snapshotDate
        if ($weekClassification.EvidenceStatus -eq 'Observed') { $weekCounts.Covered++ }
        if ($weekClassification.ActivityState -eq 'Active <=30D') { $weekCounts.Active++ }
        if ($weekClassification.ActivityState -in @('Dormant 31-90D', 'Stale >90D', 'Observed never used')) { $weekCounts.InactiveOver30++ }

        $weekRows.Add([pscustomobject][ordered]@{
            'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd')
            'Date Key' = [int]$snapshotDate.ToString('yyyyMMdd')
            'Week Label' = $week
            'User Source ID' = $userId
            'User Principal Name' = $upn
            'Confirmed Human' = $confirmedHumanEntraIds.Contains($userId) -or $confirmedHumanUpns.Contains($upn)
            'Activity State' = $weekClassification.ActivityState
            'Activity Evidence Status' = $weekClassification.EvidenceStatus
            'Days Since Last Activity' = $weekClassification.DaysSinceLastActivity
            'Last Activity Date' = if ($null -eq $weekClassification.LastActivityDate) { $null } else { $weekClassification.LastActivityDate.ToString('yyyy-MM-dd') }
            'AD Last Activity Date' = if ($null -eq $weekClassification.AdLastActivityDate) { $null } else { $weekClassification.AdLastActivityDate.ToString('yyyy-MM-dd') }
            'M365 Last Activity Date' = if ($null -eq $weekClassification.M365LastActivityDate) { $null } else { $weekClassification.M365LastActivityDate.ToString('yyyy-MM-dd') }
            'Activity Source' = $weekClassification.ActivitySource
        })
    }

    if (-not $historyHeaderWritten) {
        $weekRows | Export-Csv -LiteralPath $historyTemporaryPath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        $historyHeaderWritten = $true
    } else {
        $weekRows | Export-Csv -LiteralPath $historyTemporaryPath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded -Append
    }

    $trendRows.Add([pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd')
        'Date Key' = [int]$snapshotDate.ToString('yyyyMMdd')
        'Week Label' = $week
        'Enabled Workforce Accounts' = $weekActiveRows.Count
        'Confirmed Human Users' = @($weekActiveRows | Where-Object {
            $candidateId = Get-NormalizedKey (Get-PropertyValue $_ 'Object Id')
            $candidateUpn = Get-NormalizedKey (Get-PropertyValue $_ 'User principal name')
            ($candidateId -and $confirmedHumanEntraIds.Contains($candidateId)) -or ($candidateUpn -and $confirmedHumanUpns.Contains($candidateUpn))
        }).Count
        'Activity Covered Users' = $weekCounts.Covered
        'Activity Evidence Coverage' = if ($weekActiveRows.Count -eq 0) { $null } else { ConvertTo-InvariantDecimalString ([math]::Round($weekCounts.Covered / [double]$weekActiveRows.Count, 6)) }
        'Active Workforce' = $weekCounts.Active
        'Accounts Without Activity >30D' = $weekCounts.InactiveOver30
        'Evidence Status' = 'Observed'
    })
}

if (-not $historyHeaderWritten) { throw 'No compatible weekly user and activity snapshots were found.' }
Move-Item -LiteralPath $historyTemporaryPath -Destination $HistoryOutputPath -Force
Write-CsvAtomically @($trendRows) $TrendOutputPath

$summary = [pscustomobject][ordered]@{
    EnabledMemberAccounts = $enabledMemberUsers.Count
    EnabledWorkforceAccounts = $enabledWorkforceUsers.Count
    ConfirmedHumanUsers = $confirmedHumanEntraIds.Count
    ConfirmedNonHumanAccounts = $excludedNonHumanAccounts
    ReviewRequiredAccounts = $reviewRequiredAccounts
    AccountClassificationRuleVersion = $accountClassificationRuleVersion
    ActiveWorkforce = $activityStateCounts['Active <=30D']
    AccountsWithoutActivityOver30Days = $inactiveOver30
    ActivityCoveredUsers = $coveredUsers
    ActivityEvidenceCoverage = if ($enabledWorkforceUsers.Count -eq 0) { $null } else { [math]::Round($coveredUsers / [double]$enabledWorkforceUsers.Count, 6) }
    MissingActivityEvidence = $activityStateCounts['No activity evidence']
    ObservedNeverUsed = $activityStateCounts['Observed never used']
    ObservedStaleOver90Days = $activityStateCounts['Stale >90D']
    EligibleAdIdentities = $eligibleAdUsers.Count
    MatchedAdIdentities = $identityCounts.Matched
    IdentityConflicts = $identityCounts.Conflict
    IdentityUnmatched = $identityCounts.Unmatched
    HistoryWeeks = $trendRows.Count
    UserOutputPath = $UserOutputPath
    HistoryOutputPath = $HistoryOutputPath
    TrendOutputPath = $TrendOutputPath
    IdentityOutputPath = $IdentityOutputPath
    SignalsOutputPath = $SignalsOutputPath
}
$summary | ConvertTo-Json -Depth 4

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCVDy/BRrltE8nb
# EV13JaH06lGnNSTfX/ZHKRC2YodQRKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID/6K0s4krF+hZKZc42qvtUPPQil0IRLmvXvPArsxi+sMA0GCSqG
# SIb3DQEBAQUABIIBgFXvv4OOsTigN89UAcOWM+tBFdh1u99+8DAbZCIw7G8+nOWJ
# A/RtcUwRSzrvfcy+PDSVvL5Qs8oGmxWqjeZpCrz0eLqW6g+obxlYq54x962ekX3l
# 9RbchuaSSGyP/FiVr9KyL9lm4m1RO9plQAgMSAaPEc48/QGRwlZRkYgoJFdy0DQh
# eMsGI1lQMas3MkOdxLl9YnM0IdcZTO2FkbgFChA1rba6z6XooMxHx9LnWTPPA2xf
# xNMtvMrRMbp6/OEnUkMY4NTiAYgmKJ8lMKSlPwIIw1NgV02kz/bYeI3fSsJNL09v
# 85qWp+BSyYvOqXtq40Y4q/0MFEm2dTorcPeEJ3N2jcE27ZkmHZf1ygNSc/f6eIlc
# Ankz7InpvgKOnUt9fZ4Fob+1sBghTCRcm13A8ySqEoCqaReWg14ybOv/V9TCtjx5
# Ewsl8sgJnT4NTQ7GIQuaJBMcIa9CG9NqHZ8zzl1W+2CfVTw3urnZNMRukfnLhyaB
# +CvcdaQKn1OLFDDZqqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjYyMzUw
# MDJaMC8GCSqGSIb3DQEJBDEiBCBl4InczuKZmW9jKoIql1bt+H3DO+DWYqzrDhTJ
# VpctTTANBgkqhkiG9w0BAQEFAASCAgBowyEaXEDD+W8JMs/H/4N1HaLi2khZs2Tu
# DZKypBBJLmZXGMaDWn1OmVHXiuGtIF/+1SS4kqB5lp5xOJ2OKyqdJPLsJ8W8I0S2
# VIiIXAkgviUg9nktK1XfNzK64XlwVtYZ3C3lVDG4aXsXA8MtiWG0iBih80PMvBSv
# k+DIbyJnQ17DPomdpu2YeHrt3VhI+tsEGr7YqN0OljGWYaBYxYI+I4W8T98cDdPe
# RGZamFEZZZZzpDcT9tZp+gmGclDtUWrTFe8vrN3yLd7a9NoqFZiDDKcbO7GUqRn0
# JBR6uhwGy3iRokzOp+rrfO7dXkgb2a1WlD4MuPuyH3Er14VPtAJm5+wGgMkLwRzt
# zLDbU/+UAssy6w4GgmD3BfN7clbcyfm0jehPbOSeo3J0Ql63SqwoZDyFPZjMPWwY
# Vd8UFNlFfB8nuO0pfAo80rI89znnalazcjubMNSQ/hfXbkt2fVuZ74jwYKt6uy/q
# Bp2xYyltSGmjlYzzavKJQ/vdQJtqLj58HBZGIDfHAHaXsfn/7IhC2vq6FqBGC0HK
# yfRYvXrvoko95dImq+9iCM1RGzufjeCovYFz5/L98Ai3hGwZfzR/I3ptZ2cwPck5
# PR88/4GiyviEaBAq7/GDO/L56brZjgwNJb+x65i0RQuskrWqvN+LuJOvbXSy05Ks
# gal0uApWsQ==
# SIG # End signature block
