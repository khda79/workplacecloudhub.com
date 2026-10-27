<#
.SYNOPSIS
    Builds enriched Active Directory computer CSV columns required by the SmartWorkplace Power BI model.
.VERSION
1.13
#>

. (Join-Path $PSScriptRoot '../Common/SmartM365.IdentityDateEvidence.ps1')

function Invoke-SmartM365AdComputersEnrichedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CombinedComputersCsv,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $false)][string]$LatestFolderPath,
        [Parameter(Mandatory = $false)][string]$WindowsUpdateAnchorPolicyId,
        [Parameter(Mandatory = $false)][string]$WindowsUpdate24H2PolicyId,
        [Parameter(Mandatory = $false)][string]$WindowsUpdate25H2PolicyId,
        [Parameter(Mandatory = $true)][string]$SiteClassificationPath,
        [Parameter(Mandatory = $false)]
        [string]$AccountClassificationConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Config\AccountClassification.local.json')
    )

    $enrichmentVersion = '1.13'
    $enrichmentTimer = [System.Diagnostics.Stopwatch]::StartNew()
    WriteLog -Message ("AD computers enrichment {0} started" -f $enrichmentVersion)
    if (-not (Test-Path -LiteralPath $CombinedComputersCsv)) {
        WriteLog -Message ("AD computers enriched CSV skipped because source CSV is missing: {0}" -f $CombinedComputersCsv) -Level "WARNING"
        return $null
    }

    $columnsFile = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-ActiveDirectory-EnrichedColumns.ps1'
    if (Test-Path -LiteralPath $columnsFile) {
        $calculatedColumns = @(. $columnsFile | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    else {
        $calculatedColumns = @('ObjectGuidNormalized')
        WriteLog -Message ("AD enrichment column list not found, using ObjectGuidNormalized only: {0}" -f $columnsFile) -Level "WARNING"
    }

    # Private site workbook shared with SmartWorkplaceIntelligence; a missing or invalid workbook stops the enrichment.
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Common\SmartM365.WorkplaceClassification.psd1') -MinimumVersion '1.0.0' -ErrorAction Stop
    $siteClassification = Read-SmartM365WorkplaceSiteClassification -Path $SiteClassificationPath
    WriteLog -Message ("AD computers enrichment site classification loaded: {0} site(s)" -f $siteClassification.Sites.Count)
    # Organization-specific domain and OU rules are private configuration (git-ignored AccountClassification.local.json).
    Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'Common\SmartM365.AccountClassification.psd1') -MinimumVersion '1.1.0' -ErrorAction Stop
    $accountClassificationConfig = Read-SmartM365AccountClassification -Path $AccountClassificationConfigPath
    Assert-SmartM365DirectoryRuleConfiguration -Configuration $accountClassificationConfig
    $integrationFallback = $accountClassificationConfig.DirectoryRules.IntegrationFallback
    $headquartersOuFragments = @($accountClassificationConfig.DirectoryRules.HeadquartersOuContains | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpperInvariant() })
    $windows11OutOfScopeDomains = @($accountClassificationConfig.DirectoryRules.Windows11OutOfScopeDomains | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpperInvariant() })

    function V($Row, [string[]]$Names) {
        if ($null -eq $Row) { return '' }
        if ($Row -is [System.Collections.IDictionary]) {
            foreach ($name in $Names) { if ($null -ne $Row[$name]) { return [string]$Row[$name] } }
            return ''
        }
        foreach ($name in $Names) {
            $property = $Row.PSObject.Properties[$name]
            if ($null -ne $property -and $null -ne $property.Value) { return [string]$property.Value }
        }
        return ''
    }

    function K($Value) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        return $text.ToUpperInvariant()
    }

    function NK($Value) {
        $text = ([string]$Value).Trim()
        if ($text.EndsWith('$')) { $text = $text.Substring(0, $text.Length - 1) }
        return (K $text)
    }

    function GK($Value) {
        $text = ([string]$Value).Trim().Trim('{','}')
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        $parsedGuid = [guid]::Empty
        if ([guid]::TryParse($text, [ref]$parsedGuid)) { return $parsedGuid.ToString('D').ToLowerInvariant() }
        return $text.ToLowerInvariant()
    }

    function B($Value) {
        if ($null -eq $Value) { return '' }
        if ($Value -is [bool]) { return [string]$Value }
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        switch -Regex ($text.ToLowerInvariant()) {
            '^(true|yes|y|1|compliant)$' { return 'True' }
            '^(false|no|n|0|noncompliant)$' { return 'False' }
            default { return $text }
        }
    }

    function Num($Value) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $text = $text.Replace(',', '.')
        $number = 0.0
        if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
        return $null
    }

    function Gb($Value) {
        $number = Num $Value
        if ($null -eq $number) { return '' }
        if ($number -gt 1048576) { $number = $number / 1GB }
        return $number.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Dt($Value) {
        if ([string]$Value -notmatch '(?:Z|[+-]\d{2}:\d{2})$') { return $null }
        $valueDate = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$valueDate)) { return $valueDate.UtcDateTime }
        return $null
    }

    function Fmt($Value) {
        $date = Dt $Value
        if ($null -eq $date) { return '' }
        return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Days($Value) {
        $date = Dt $Value
        if ($null -eq $date) { return '' }
        return ([datetime]::UtcNow.Date - $date.ToUniversalTime().Date).Days.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }

function Build($AdVersion, $IntuneVersion) {
        $intune = ([string]$IntuneVersion).Trim()
        if ($intune -match '^\d+\.\d+\.(\d+)') { return $Matches[1] }
        $raw = ([string]$AdVersion).Trim()
        if ($raw -match '\((\d{4,5})') { return $Matches[1] }
        if ($raw -match '(^|\D)(\d{4,5})(\D|$)') { return $Matches[2] }
        return ''
    }

    function FriendlyOs($AdVersion, $IntuneVersion) {
        $build = Build $AdVersion $IntuneVersion
        switch ($build) {
            '28000' { return 'Windows 11 26H1' }
            '26200' { return 'Windows 11 25H2' }
            '26100' { return 'Windows 11 24H2' }
            '22631' { return 'Windows 11 23H2' }
            '22621' { return 'Windows 11 22H2' }
            '22000' { return 'Windows 11 21H2' }
            '19045' { return 'Windows 10 22H2' }
            '19044' { return 'Windows 10 21H2' }
            '19043' { return 'Windows 10 21H1' }
            '19042' { return 'Windows 10 20H2' }
            '19041' { return 'Windows 10 2004' }
            '18363' { return 'Windows 10 1909' }
            '18362' { return 'Windows 10 1903' }
            '17763' { return 'Windows 10 1809' }
            '17134' { return 'Windows 10 1803' }
            '16299' { return 'Windows 10 1709' }
            '15063' { return 'Windows 10 1703' }
            '14393' { return 'Windows 10 1607' }
            '10586' { return 'Windows 10 1511' }
            '10240' { return 'Windows 10 1507' }
            '9600'  { return 'Windows 8.1' }
            '9200'  { return 'Windows 8' }
            '7601'  { return 'Windows 7 SP1' }
            '7600'  { return 'Windows 7' }
            default {
                if ($build -match '^28') { return ("Windows 11 Insider ({0})" -f $build) }
                if ($build -match '^26') { return ("Windows 11 Insider ({0})" -f $build) }
                $source = if (-not [string]::IsNullOrWhiteSpace($IntuneVersion)) { $IntuneVersion } else { $AdVersion }
                if ([string]::IsNullOrWhiteSpace($source)) { return '?' }
                return ("OS Unknown ({0})" -f $source)
            }
        }
    }

    function Eligibility($Value) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        switch ($text.ToUpperInvariant()) {
            '0' { return 'Upgraded' }
            '1' { return 'Unknown' }
            '2' { return 'Not Capable' }
            '3' { return 'Capable' }
            '4' { return 'Unknown Future Value' }
            'UPGRADED' { return 'Upgraded' }
            'UNKNOWN' { return 'Unknown' }
            'NOTCAPABLE' { return 'Not Capable' }
            'NOT CAPABLE' { return 'Not Capable' }
            'CAPABLE' { return 'Capable' }
            default { return $text }
        }
    }

    function GetDateCategory($Value, [int]$LimitDays) {
        $date = Dt $Value
        if ($null -eq $date) { return 'Unknown / unqualified activity' }
        if ($date.ToUniversalTime().Date -lt ([datetime]::UtcNow.Date.AddDays(-1 * $LimitDays))) { return ("Older than {0} days" -f $LimitDays) }
        return ("Within last {0} days" -f $LimitDays)
    }

    function Get-SourceCandidatePaths($Name) {
        $candidates = New-Object System.Collections.Generic.List[string]
        $roots = New-Object System.Collections.Generic.List[string]
        foreach ($root in @($LatestFolderPath, $OutputFolder)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$root)) { [void]$roots.Add([string]$root) }
        }
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) {
            $latestParent = Split-Path -Path $LatestFolderPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($latestParent)) { [void]$roots.Add((Join-Path -Path $latestParent -ChildPath 'DATA-ALL')) }
        }
        if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
            $currentRoot = $OutputFolder
            for ($i = 0; $i -lt 6 -and -not [string]::IsNullOrWhiteSpace($currentRoot); $i++) {
                if ([System.IO.Path]::GetFileName($currentRoot).Equals('DATA-ALL', [System.StringComparison]::OrdinalIgnoreCase)) { [void]$roots.Add($currentRoot); break }
                $currentRoot = Split-Path -Path $currentRoot -Parent
            }
        }
        foreach ($rootForTenantLookup in @($LatestFolderPath, $OutputFolder, (Split-Path -Path $CombinedComputersCsv -Parent))) {
            $currentRoot = [string]$rootForTenantLookup
            for ($i = 0; $i -lt 8 -and -not [string]::IsNullOrWhiteSpace($currentRoot); $i++) {
                if ([System.IO.Path]::GetFileName($currentRoot).Equals("Tenants", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $dataRoot = Split-Path -Path $currentRoot -Parent
                    if (-not [string]::IsNullOrWhiteSpace($dataRoot)) { [void]$roots.Add((Join-Path -Path $dataRoot -ChildPath "DATA-ALL")) }
                    break
                }
                $currentRoot = Split-Path -Path $currentRoot -Parent
            }
        }
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($root in $roots) {
            if ([string]::IsNullOrWhiteSpace([string]$root)) { continue }
            $candidate = Join-Path -Path $root -ChildPath $Name
            if ($seen.Add($candidate)) { [void]$candidates.Add($candidate) }
        }
        return @($candidates)
    }

    function LoadCsv($Name, [string[]]$Columns, [string]$PreferredPath = '') {
        $candidatePaths = New-Object System.Collections.Generic.List[string]
        $seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        if (-not [string]::IsNullOrWhiteSpace($PreferredPath) -and $seenPaths.Add($PreferredPath)) {
            [void]$candidatePaths.Add($PreferredPath)
        }
        foreach ($candidate in (Get-SourceCandidatePaths $Name)) {
            if ($seenPaths.Add($candidate)) { [void]$candidatePaths.Add($candidate) }
        }

        # Only the columns read by this enrichment are kept in memory; a column missing from the file
        # becomes a null value, which V() skips exactly like an absent property. Projected rows are hashtables:
        # PowerShell row objects made every garbage collection measurably slower.
        $projection = New-Object System.Collections.Generic.List[string]
        $projectionSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($column in $Columns) { if ($projectionSet.Add($column)) { [void]$projection.Add($column) } }
        $projectionNames = $projection.ToArray()
        foreach ($candidate in $candidatePaths) {
            if (Test-Path -LiteralPath $candidate) {
                $projectedRows = New-Object System.Collections.Generic.List[object]
                Import-Csv -LiteralPath $candidate -ErrorAction Stop | & {
                    process {
                        $projected = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
                        foreach ($name in $projectionNames) {
                            $property = $_.PSObject.Properties[$name]
                            $projected[$name] = if ($null -ne $property) { $property.Value } else { $null }
                        }
                        [void]$projectedRows.Add($projected)
                    }
                }
                $rows = $projectedRows.ToArray()
                WriteLog -Message ("AD enrichment source loaded: {0} ({1} row(s))" -f $candidate, $rows.Count)
                return $rows
            }
        }
        WriteLog -Message ("AD enrichment optional source missing; related columns will be blank: {0}" -f $Name)
        return @()
    }

    function AddMap($Map, $Key, $Row) { if (-not [string]::IsNullOrWhiteSpace($Key) -and -not $Map.ContainsKey($Key)) { $Map[$Key] = $Row } }
    function AddMulti($Map, $Key, $Row) { if ([string]::IsNullOrWhiteSpace($Key)) { return }; if (-not $Map.ContainsKey($Key)) { $Map[$Key] = New-Object System.Collections.Generic.List[object] }; [void]$Map[$Key].Add($Row) }
    function GetMap($Map, $Key) { if (-not [string]::IsNullOrWhiteSpace($Key) -and $Map.ContainsKey($Key)) { return $Map[$Key] }; return $null }
    function MapCount($Map, $Key) {
        if ([string]::IsNullOrWhiteSpace($Key) -or -not $Map.ContainsKey($Key)) { return 0 }
        $value = $Map[$Key]
        if ($null -eq $value) { return 0 }
        if ($value -is [System.Collections.ICollection]) { return [int]$value.Count }
        return 1
    }

    function AddLatest($Map, $Key, $Row) {
        if ([string]::IsNullOrWhiteSpace($Key)) { return }
        if (-not $Map.ContainsKey($Key)) { $Map[$Key] = @() }
        $Map[$Key] = @($Map[$Key]) + @($Row)
    }

    function SetWu($Output, $Row, [string]$PolicyName) {
        $Output[("WindowsUpdate{0}AggregateState" -f $PolicyName)] = V $Row @('AggregateState_loc','AggregateState')
        $Output[("WindowsUpdate{0}CurrentDeviceUpdateStatus" -f $PolicyName)] = V $Row @('CurrentDeviceUpdateStatus_loc','CurrentDeviceUpdateStatus')
        $Output[("WindowsUpdate{0}LatestAlertMessage" -f $PolicyName)] = V $Row @('LatestAlertMessage_loc','LatestAlertMessage')
        $Output[("WindowsUpdate{0}BlockingReason" -f $PolicyName)] = V $Row @('BlockingReason','LatestAlertMessage_loc','LatestAlertMessage')
        $Output[("WindowsUpdate{0}RiskBucket" -f $PolicyName)] = V $Row @('RiskBucket')
    }

    function BoolValue($Value) {
        $text = B $Value
        if ($text -eq 'True') { return $true }
        if ($text -eq 'False') { return $false }
        return $false
    }

    function GetOsMajor($Value) {
        $text = ([string]$Value).Trim().ToUpperInvariant()
        if ($text.Contains('WINDOWS 11')) { return 'Windows 11' }
        if ($text.Contains('WINDOWS 10')) { return 'Windows 10' }
        return ''
    }

    function GetBuildNumber($AdVersion, $IntuneVersion) {
        $buildText = Build $AdVersion $IntuneVersion
        $buildNumber = 0
        if ([int]::TryParse(([string]$buildText), [ref]$buildNumber)) { return $buildNumber }
        return $null
    }

    function IsUnknownUser($Value) {
        $text = ([string]$Value).Trim().ToUpperInvariant()
        return @('', 'UNKNOWN', 'N/A', 'NA', '-', 'SYSTEM', 'LOCAL SYSTEM') -contains $text
    }

    function NormalizeDomainSam($Value) {
        $text = ([string]$Value).Trim().Replace([string][char]160, '').Replace([string][char]9, '').Replace(' ', '')
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        return $text.ToLowerInvariant()
    }

    function GetLastOuName($Value) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        if ($text -like '*,*') {
            $ouParts = @($text -split ',' | Where-Object { $_ -match '^(?i)OU=' } | ForEach-Object { ($_ -replace '^(?i)OU=', '').Trim() })
            if ($ouParts.Count -gt 0) { return $ouParts[0] }
        }
        $parts = @($text -split '[,/]' | Where-Object { $_ })
        if ($parts.Count -ge 2) { return $parts[$parts.Count - 2] }
        return ''
    }

    function GetOrganizationalUnitCode($Row) {
        $raw = (V $Row @('OrganizationalUnit','extensionAttribute13')).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { return '999999' }
        $numVal = 0.0
        if ([double]::TryParse($raw.Replace(',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numVal)) {
            return ([math]::Floor($numVal)).ToString('000000', [System.Globalization.CultureInfo]::InvariantCulture)
        }
        return '999998'
    }

    function GetSubnet($Value) {
        $text = ([string]$Value).Trim()
        if ($text -match '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$') { return ('{0}.{1}.{2}.0' -f $Matches[1], $Matches[2], $Matches[3]) }
        return ''
    }

    function GetTypeEntityFallback($DomainNameShort, $DistinguishedName) {
        $domainClean = ([string]$DomainNameShort).Trim().ToUpperInvariant()
        $dnClean = ([string]$DistinguishedName).Trim().ToUpperInvariant()
        # Country: the two characters before the configured country OU suffix, when present.
        $countryFromDn = ''
        $countrySuffix = ([string]$integrationFallback.ComputerCountryOuSuffix).ToUpperInvariant()
        if ($countrySuffix -and $dnClean.Contains($countrySuffix)) {
            $beforeCountries = $dnClean.Substring(0, $dnClean.IndexOf($countrySuffix))
            if ($beforeCountries.Length -ge 2) { $countryFromDn = $beforeCountries.Substring($beforeCountries.Length - 2, 2) }
        }
        return Get-SmartM365IntegrationFallback -Rules $integrationFallback.Computers -DomainNameShort $domainClean -Country $countryFromDn -Default $integrationFallback.Default
    }

    function GetAccountToDeleteFromAd($Row) {
        $sam = V $Row @('SamAccountName')
        if ($sam -match '(?i)-K-') { return 'NO_CLEAN_DEVICES_EXCLUDED' }
        $lastLogonUnknown = V $Row @('IsLastLogonUnknown')
        if (-not $lastLogonUnknown) { $lastLogonUnknown = if (Dt (V $Row @('LastLogonDate'))) { 'Known' } else { 'Unknown' } }
        $lastLogonDate = Dt (V $Row @('LastLogonDate'))
        if ($null -eq $lastLogonDate) { $lastLogonDate = [datetime]'1900-01-01' }
        $whenChangedDate = Dt (V $Row @('WhenChanged'))
        if ($null -eq $whenChangedDate) { $whenChangedDate = [datetime]'1900-01-01' }
        $cutoffDate = [datetime]'2025-01-01'
        if ($lastLogonUnknown -eq 'Unknown' -and $whenChangedDate -lt $cutoffDate) { return 'TOCLEAN_STEP1' }
        if ($lastLogonUnknown -eq 'Known' -and $lastLogonDate -lt $cutoffDate) { return 'TOCLEAN_STEP1' }
        return 'NO_CLEAN'
    }

    function GetMigrationPhase($Eligibility, [bool]$OsMinOk, [bool]$IsLtsc, [bool]$IsEnabled, $ShortOs, $TypeEntity, [bool]$InDisabledOu, $DomainShort, [bool]$IsActiveLast45Days, [bool]$IsInIntune, $ModelName) {
        $eligibilityText = ([string]$Eligibility).Trim().ToUpperInvariant()
        if (-not $eligibilityText) { $eligibilityText = 'UNKNOWN' }
        $shortOsText = ([string]$ShortOs).Trim()
        $isWindows11Value = $shortOsText -eq 'Windows 11'
        $isWindows10Value = $shortOsText -eq 'Windows 10'
        $isSupportedClientOs = $isWindows10Value -or $isWindows11Value
        $typeEntityText = ([string]$TypeEntity).Trim().ToUpperInvariant()
        $domainShortText = ([string]$DomainShort).Trim().ToUpperInvariant()
        $modelText = ([string]$ModelName).Trim().ToUpperInvariant()
        $isVmValue = $modelText -match 'VMWARE|VIRTUAL|HYPER-V|VIRTUALBOX|KVM|XEN'
        if (-not $isSupportedClientOs) { return 'OUT OF SCOPE' }
        if ($isVmValue) { return 'OUT OF SCOPE' }
        if ($typeEntityText -eq 'NOT-INTEGRATED') { return 'OUT OF SCOPE' }
        if (-not $IsEnabled) { return 'OUT OF SCOPE' }
        if ($windows11OutOfScopeDomains -contains $domainShortText) { return 'OUT OF SCOPE' }
        if ($InDisabledOu) { return 'OUT OF SCOPE' }
        if ((-not $IsActiveLast45Days) -and (-not ($isWindows11Value -and $IsInIntune))) { return 'PHASE 2' }
        if ($IsLtsc -and $isWindows10Value) { return 'PHASE 2' }
        if ($eligibilityText -eq 'UPGRADED' -and -not $isWindows11Value) { return 'PHASE 1' }
        if ($isWindows11Value) { return 'PHASE 1' }
        if ($eligibilityText -eq 'UPGRADED') { return 'PHASE 1' }
        if (@('UNKNOWN','UNDETERMINED') -contains $eligibilityText) { return 'PHASE 2' }
        if ($eligibilityText -eq 'CAPABLE' -and $OsMinOk) { return 'PHASE 1' }
        if ($eligibilityText -eq 'CAPABLE' -and -not $OsMinOk) { return 'PHASE 2' }
        if ($eligibilityText -eq 'NOT CAPABLE') { return 'PHASE 2' }
        return 'PHASE 2'
    }

    function GetMigrationPhaseReason($Eligibility, [bool]$OsMinOk, [bool]$IsLtsc, [bool]$IsEnabled, $ShortOs, $TypeEntity, [bool]$InDisabledOu, $DomainShort, [bool]$IsActiveLast45Days, [bool]$IsInIntune, $ModelName) {
        $eligibilityText = ([string]$Eligibility).Trim().ToUpperInvariant()
        if (-not $eligibilityText) { $eligibilityText = 'UNKNOWN' }
        $shortOsText = ([string]$ShortOs).Trim()
        $isWindows11Value = $shortOsText -eq 'Windows 11'
        $isWindows10Value = $shortOsText -eq 'Windows 10'
        $isSupportedClientOs = $isWindows10Value -or $isWindows11Value
        $typeEntityText = ([string]$TypeEntity).Trim().ToUpperInvariant()
        $domainShortText = ([string]$DomainShort).Trim().ToUpperInvariant()
        $modelText = ([string]$ModelName).Trim().ToUpperInvariant()
        $isVmValue = $modelText -match 'VMWARE|VIRTUAL|HYPER-V|VIRTUALBOX|KVM|XEN'
        if (-not $isSupportedClientOs) { return 'OS != Windows 10/11 -> OUT OF SCOPE' }
        if ($isVmValue) { return 'Virtual machine detected from Model -> OUT OF SCOPE' }
        if ($typeEntityText -eq 'NOT-INTEGRATED') { return 'TypeEntity = NOT-INTEGRATED -> OUT OF SCOPE' }
        if ($typeEntityText -eq 'EXCLUDED') { return 'TypeEntity = EXCLUDED -> OUT OF SCOPE' }
        if (-not $IsEnabled) { return 'Device disabled -> OUT OF SCOPE' }
        if ($windows11OutOfScopeDomains -contains $domainShortText) { return ('Domain = {0} -> OUT OF SCOPE' -f $domainShortText) }
        if ($InDisabledOu) { return 'Disabled Objects OU -> OUT OF SCOPE' }
        if ((@('UNKNOWN','UNDETERMINED','NOT CAPABLE','CAPABLE','UPGRADED') -contains $eligibilityText) -and (-not $IsActiveLast45Days) -and (-not ($isWindows11Value -and $IsInIntune))) { return ('Eligibility = {0} + Inactive (no sign-in last 45 days) -> PHASE 2' -f $eligibilityText) }
        if ((@('UNKNOWN','UNDETERMINED','NOT CAPABLE','CAPABLE','UPGRADED') -contains $eligibilityText) -and $IsLtsc -and $isWindows10Value) { return ('Eligibility = {0} + Windows 10 LTSC -> PHASE 2' -f $eligibilityText) }
        if ($isWindows11Value) { return 'Already Windows 11 (pre-existing) -> PHASE 1' }
        if ($eligibilityText -eq 'UPGRADED' -and -not $isWindows11Value) { return 'Data inconsistency: Eligibility = UPGRADED but OS != Windows 11 -> PHASE 1' }
        if (@('UNKNOWN','UNDETERMINED') -contains $eligibilityText) { return ('Eligibility = {0} + OS version {1} -> PHASE 2' -f $eligibilityText, $(if ($OsMinOk) { 'supported' } else { 'not supported' })) }
        if ($eligibilityText -eq 'NOT CAPABLE') { return 'Eligibility = NOT CAPABLE -> PHASE 2' }
        if ($eligibilityText -eq 'CAPABLE' -and $OsMinOk) { return 'Eligibility = CAPABLE + OS version supported -> PHASE 1' }
        if ($eligibilityText -eq 'CAPABLE' -and -not $OsMinOk) { return 'Eligibility = CAPABLE + OS version not supported -> PHASE 2' }
        if ($eligibilityText -eq 'UPGRADED') { return 'Eligibility = UPGRADED -> PHASE 1' }
        return 'Default logic -> PHASE 2'
    }

    function GetMigrationAction($Eligibility, [bool]$OsMinOk, [bool]$IsLtsc, [bool]$IsEnabled, $ShortOs, $TypeEntity, [bool]$InDisabledOu, $DomainShort, [bool]$IsActiveLast45Days, [bool]$IsInIntune, $MemoryGb) {
        $eligibilityText = ([string]$Eligibility).Trim().ToUpperInvariant()
        if (-not $eligibilityText) { $eligibilityText = 'UNKNOWN' }
        $shortOsText = ([string]$ShortOs).Trim()
        $isWindows11Value = $shortOsText -eq 'Windows 11'
        $isWindows10Value = $shortOsText -eq 'Windows 10'
        $isSupportedClientOs = $isWindows10Value -or $isWindows11Value
        $typeEntityText = ([string]$TypeEntity).Trim().ToUpperInvariant()
        $domainShortText = ([string]$DomainShort).Trim().ToUpperInvariant()
        $ram = Num $MemoryGb
        $ramLow = $null -ne $ram -and $ram -lt 8
        if (-not $isSupportedClientOs) { return '6 - To qualify' }
        if ($typeEntityText -eq 'NOT-INTEGRATED') { return '6 - To qualify' }
        if ($typeEntityText -eq 'EXCLUDED') { return '6 - To qualify' }
        if (-not $IsEnabled) { return '6 - To qualify' }
        if ($windows11OutOfScopeDomains -contains $domainShortText) { return '6 - To qualify' }
        if ($InDisabledOu) { return '6 - To qualify' }
        if ($isWindows11Value) { return '6 - To qualify' }
        if ((-not $IsActiveLast45Days) -and (-not ($isWindows11Value -and $IsInIntune))) { return '5 - To check inactivity' }
        if ($eligibilityText -eq 'NOT CAPABLE') { return '4 - To replace' }
        if (@('UNKNOWN','UNDETERMINED') -contains $eligibilityText) { return '6 - To qualify' }
        if ($IsLtsc -and $isWindows10Value) { return '2 - To reinstall (no Intune)' }
        if ($eligibilityText -eq 'UPGRADED') { return '6 - To qualify' }
        if ($eligibilityText -eq 'CAPABLE' -and $ramLow) { return '3 - To upgrade RAM + reinstall' }
        if ($eligibilityText -eq 'CAPABLE' -and $OsMinOk) { return '1 - To update (Intune)' }
        if ($eligibilityText -eq 'CAPABLE' -and -not $OsMinOk) { return '2 - To reinstall (no Intune)' }
        return '6 - To qualify'
    }
    # Column lists must cover every name read through V() for the rows of each source.
    $mailboxColumns = @('UserPrincipalName','PrimarySmtpAddress','TotalItemSizeGB','TotalItemSizeMB_Integer')
    $onPremMailboxColumns = @('UserPrincipalName','PrimarySMTPaddress','PrimarySmtpAddress','WindowsEmailAddress','TotalItemSize-In-MB','TotalItemSizeGB')
    $intuneRows = @(LoadCsv 'Intune_Devices_Inventory.csv' @('AzureADDeviceId_Norm','Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId','Device ID','DeviceId','OS version','OSVersion','OperatingSystemVersion','UpgradeEligibility','UpgradeEligibilityLabel','Total storage','TotalStorage','TotalStorageGB','Free storage','FreeStorage','FreeStorageGB','Primary user UPN','Primary user email address','UserPrincipalName','UPN','Last check-in','LastSyncDateTime','Model','Manufacturer','PhysicalMemoryGB','Intune registered','Enrollment date','EnrollmentDate','Ownership','Compliance','IsCompliant'))
    $entraRows = @(LoadCsv 'M365_Entra_Devices.csv' @('DeviceId_Norm','DeviceId','DisplayName','DeviceName','Name','HardwareId','RegistrationDateTime','ApproximateLastSignInDateTime','ObjectId','IsPending','TrustType','trustType'))
    $nativeEntraById = @{}; $nativeManagedByAad = @{}
    foreach ($native in @(LoadCsv 'M365_EntraDevices_All.csv' @('DeviceId','RegistrationDateTime','ApproximateLastSignInDateTime'))) {
        $proofRow = [pscustomobject]@{RegistrationDateTimeIso8601=(V $native @('RegistrationDateTime')); ApproximateLastSignInDateTimeIso8601=(V $native @('ApproximateLastSignInDateTime'))}
        $proofKey = GK (V $native @('DeviceId'))
        if ($proofKey) { if ($nativeEntraById.ContainsKey($proofKey)) { $nativeEntraById[$proofKey] = $null } else { $nativeEntraById[$proofKey] = $proofRow } }
    }
    foreach ($native in @(LoadCsv 'Intune_ManagedDevices_All.csv' @('AzureADDeviceId','EnrolledDateTime','LastSyncDateTime'))) {
        $proofRow = [pscustomobject]@{EnrolledDateTimeIso8601=(V $native @('EnrolledDateTime'));LastSyncDateTimeIso8601=(V $native @('LastSyncDateTime'))}
        $proofKey = GK (V $native @('AzureADDeviceId'))
        if ($proofKey) { if ($nativeManagedByAad.ContainsKey($proofKey)) { $nativeManagedByAad[$proofKey] = $null } else { $nativeManagedByAad[$proofKey] = $proofRow } }
    }
    $hardwareConflictRows = @(LoadCsv 'M365_Entra_Devices_HardwareIdConflicts.csv' @('HardwareId','DeviceCount'))
    $localSystemRows = @(LoadCsv 'Intune_Devices_LocalSystem.csv' @('AzureADDeviceId_Norm','AzureADDeviceId','Azure AD Device ID','SecureBootStatus','BIOSVersion','LastRebootDate','LastBootUpTime'))
    $windowsUpdateRows = @(LoadCsv 'Intune_WindowsUpdate_Status.csv' @('DeviceId_Norm','DeviceId','PolicyId','ExportDateTime','ExportDateTimeIso8601','ReadinessExportDateTime','LastSyncDateTime','Last check-in','AggregateState_loc','AggregateState','CurrentDeviceUpdateStatus_loc','CurrentDeviceUpdateStatus','LatestAlertMessage_loc','LatestAlertMessage','BlockingReason','RiskBucket','UpgradeEligibility','UpgradeEligibilityLabel'))
    $currentRunAdUsersCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Users_AllDomains.csv'
    if (-not (Test-Path -LiteralPath $currentRunAdUsersCsv)) { $currentRunAdUsersCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Users_AllDomains_Brut.csv' }
    $adUserRows = @(LoadCsv -Name 'AD_Users_AllDomains.csv' -Columns @('UserPrincipalName','EmailAddress','TargetAddress','SamAccountName','DomainAndSam','DomainNameShort','CanonicalName','DistinguishedName') -PreferredPath $currentRunAdUsersCsv)
    $licenseRows = @(LoadCsv 'M365_Licenses_Users.csv' @('User principal name','UserPrincipalName','primarysmtp','SKU name','SkuPartNumber','GroupsAssigningSku'))
    $exoMailboxRows = @(LoadCsv 'Exchange_EXO_Mailboxes_AllDomains.csv' $mailboxColumns)
    $exoMailboxStatsRows = @(LoadCsv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv' $mailboxColumns)
    $localMailboxRows = @(LoadCsv 'Exchange_OnPrem_Mailboxes_AllDomains.csv' $onPremMailboxColumns)
    $remoteMailboxRows = @(LoadCsv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' $onPremMailboxColumns)
    WriteLog -Message ("AD computers enrichment sources loaded in {0:N1} min; managed memory {1:N0} MB" -f $enrichmentTimer.Elapsed.TotalMinutes, ([GC]::GetTotalMemory($false) / 1MB))

    $intuneByAad = @{}
    foreach ($row in $intuneRows) {
        AddMap $intuneByAad (GK (V $row @('AzureADDeviceId_Norm','Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId'))) $row
    }

    $entraByDeviceId = @{}; $entraByDeviceIdAll = @{}; $entraByNameAll = @{}; $hardwareCounts = @{}
    foreach ($row in $hardwareConflictRows) { $hid = K (V $row @('HardwareId')); $count = Num (V $row @('DeviceCount')); if ($hid -and $null -ne $count) { $hardwareCounts[$hid] = [int]$count } }
    foreach ($row in $entraRows) {
        $deviceKey = GK (V $row @('DeviceId_Norm','DeviceId'))
        $nameKey = NK (V $row @('DisplayName','DeviceName','Name'))
        AddMap $entraByDeviceId $deviceKey $row
        AddMulti $entraByDeviceIdAll $deviceKey $row
        AddMulti $entraByNameAll $nameKey $row
        $hid = K (V $row @('HardwareId'))
        if ($hid) { if (-not $hardwareCounts.ContainsKey($hid)) { $hardwareCounts[$hid] = 0 }; $hardwareCounts[$hid] = [int]$hardwareCounts[$hid] + 1 }
    }

    $localByAad = @{}
    foreach ($row in $localSystemRows) {
        AddMap $localByAad (GK (V $row @('AzureADDeviceId_Norm','AzureADDeviceId','Azure AD Device ID'))) $row
    }

    $anchorKey = GK $WindowsUpdateAnchorPolicyId
    $policy24Key = GK $WindowsUpdate24H2PolicyId
    $policy25Key = GK $WindowsUpdate25H2PolicyId
    $wuAnyByDevice = @{}; $wuAnchorByDevice = @{}; $wu24ByDevice = @{}; $wu25ByDevice = @{}
    foreach ($row in $windowsUpdateRows) {
        $deviceKey = GK (V $row @('DeviceId_Norm','DeviceId'))
        $policyKey = GK (V $row @('PolicyId'))
        AddLatest $wuAnyByDevice $deviceKey $row
        if ($anchorKey -and $policyKey -eq $anchorKey) { AddLatest $wuAnchorByDevice $deviceKey $row }
        if ($policy24Key -and $policyKey -eq $policy24Key) { AddLatest $wu24ByDevice $deviceKey $row }
        if ($policy25Key -and $policyKey -eq $policy25Key) { AddLatest $wu25ByDevice $deviceKey $row }
    }

    $usersByUpn = @{}; $licensedUsersByUpn = @{}
    foreach ($row in $adUserRows) {
        foreach ($candidate in @((V $row @('UserPrincipalName')), (V $row @('EmailAddress')), (V $row @('TargetAddress')))) {
            $key = K $candidate
            if ($key) { AddMap $usersByUpn $key $row }
        }
    }
    foreach ($row in $licenseRows) {
        $key = K (V $row @('User principal name','UserPrincipalName','primarysmtp'))
        if ($key) { $licensedUsersByUpn[$key] = $true }
    }

    $usersBySam = @{}; $usersByDomainAndSam = @{}
    foreach ($row in $adUserRows) {
        $samKey = K (V $row @('SamAccountName'))
        AddMap $usersBySam $samKey $row
        $domainSamKey = NormalizeDomainSam (V $row @('DomainAndSam'))
        if (-not $domainSamKey) {
            $domain = V $row @('DomainNameShort')
            $sam = V $row @('SamAccountName')
            if ($domain -and $sam) { $domainSamKey = NormalizeDomainSam (('{0}\{1}' -f $domain, $sam)) }
        }
        AddMap $usersByDomainAndSam $domainSamKey $row
    }

    $licenseByUpn = @{}; $licenseGroupByUpn = @{}
    foreach ($row in $licenseRows) {
        $key = K (V $row @('User principal name','UserPrincipalName','primarysmtp'))
        if ($key) {
            $licenseByUpn[$key] = $true
            $sku = V $row @('SKU name','SkuPartNumber')
            $groups = V $row @('GroupsAssigningSku')
            if (-not $licenseGroupByUpn.ContainsKey($key)) { $licenseGroupByUpn[$key] = New-Object System.Collections.Generic.List[string] }
            if ($groups) { [void]$licenseGroupByUpn[$key].Add($groups) } elseif ($sku) { [void]$licenseGroupByUpn[$key].Add($sku) }
        }
    }

    $mailboxByUpn = @{}; $mailboxSizeByUpn = @{}
    foreach ($row in $exoMailboxRows) {
        foreach ($candidate in @((V $row @('UserPrincipalName')), (V $row @('PrimarySmtpAddress')))) {
            $key = K $candidate
            AddMap $mailboxByUpn $key $row
            $size = V $row @('TotalItemSizeGB','TotalItemSizeMB_Integer')
            if ($key -and $size -and -not $mailboxSizeByUpn.ContainsKey($key)) { $mailboxSizeByUpn[$key] = $size }
        }
    }
    foreach ($row in $exoMailboxStatsRows) {
        foreach ($candidate in @((V $row @('UserPrincipalName')), (V $row @('PrimarySmtpAddress')))) {
            $key = K $candidate
            $size = V $row @('TotalItemSizeGB','TotalItemSizeMB_Integer')
            if ($key -and $size) { $mailboxSizeByUpn[$key] = $size }
        }
    }
    foreach ($row in @($localMailboxRows + $remoteMailboxRows)) {
        foreach ($candidate in @((V $row @('UserPrincipalName')), (V $row @('PrimarySMTPaddress','PrimarySmtpAddress')), (V $row @('WindowsEmailAddress')))) {
            $key = K $candidate
            AddMap $mailboxByUpn $key $row
            $size = V $row @('TotalItemSize-In-MB','TotalItemSizeGB')
            if ($key -and $size -and -not $mailboxSizeByUpn.ContainsKey($key)) { $mailboxSizeByUpn[$key] = $size }
        }
    }


    $serverAdEntraByName = @{}
    foreach ($row in $entraRows) {
        if ((V $row @('TrustType','trustType')) -ne 'ServerAd') { continue }
        $nameKey = NK (V $row @('DisplayName','DeviceName','Name'))
        if (-not $nameKey) { continue }
        $current = V $row @('DeviceId')
        if (-not $serverAdEntraByName.ContainsKey($nameKey) -or ([string]$current).CompareTo([string](V $serverAdEntraByName[$nameKey] @('DeviceId'))) -gt 0) { $serverAdEntraByName[$nameKey] = $row }
    }
    $sourceRows = @(Import-Csv -LiteralPath $CombinedComputersCsv -ErrorAction Stop)
    $sourceColumnSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($sourceRows.Count -gt 0) {
        foreach ($property in $sourceRows[0].PSObject.Properties) { [void]$sourceColumnSet.Add($property.Name) }
    }
    $calculatedColumns += @('ActivityEvidenceCoverage','ActivityEvidenceStatus','WindowsUpdateTemporalSelectionStatus','WindowsUpdateAnchorSelectionStatus','WindowsUpdate24H2SelectionStatus','WindowsUpdate25H2SelectionStatus','ActivityScopeExclusions','ActivityConclusionScope')
    $calculatedColumnsToAdd = @($calculatedColumns | Where-Object { -not $sourceColumnSet.Contains([string]$_) })
    $rawCalculatedColumnRenames = [ordered]@{
        'ConfiguredGroupMatches' = 'MatchedConfiguredGroups'
        'Has_ConfiguredGroup01'  = 'IsMemberOfConfiguredGroup01'
        'Has_ConfiguredGroup02'  = 'IsMemberOfConfiguredGroup02'
        'Has_ConfiguredGroup03'  = 'IsMemberOfConfiguredGroup03'
        'Has_ConfiguredGroup04'  = 'IsMemberOfConfiguredGroup04'
        'Has_ConfiguredGroup05'  = 'IsMemberOfConfiguredGroup05'
        'Has_ConfiguredGroup06'  = 'IsMemberOfConfiguredGroup06'
        'Has_ConfiguredGroup07'  = 'IsMemberOfConfiguredGroup07'
        'Has_ConfiguredGroup08'  = 'IsMemberOfConfiguredGroup08'
        'Has_ConfiguredGroup09'  = 'IsMemberOfConfiguredGroup09'
        'Has_ConfiguredGroup10'  = 'IsMemberOfConfiguredGroup10'
    }
    $enrichedRows = New-Object System.Collections.Generic.List[object]
    $progressInterval = 2000
    $sourceRowCount = $sourceRows.Count
    $loopTimer = [System.Diagnostics.Stopwatch]::StartNew()
    WriteLog -Message ("AD computers enrichment lookups built in {0:N1} min; enriching {1} computer row(s)" -f $enrichmentTimer.Elapsed.TotalMinutes, $sourceRowCount)

    foreach ($row in $sourceRows) {
        # Periodic progress tells a slow run apart from a stalled one.
        if ($enrichedRows.Count -gt 0 -and ($enrichedRows.Count % $progressInterval) -eq 0) {
            WriteLog -Message ("AD computers enrichment progress: {0}/{1} row(s) ({2:N0}%); loop {3:N1} min; managed memory {4:N0} MB" -f $enrichedRows.Count, $sourceRowCount, (100.0 * $enrichedRows.Count / $sourceRowCount), $loopTimer.Elapsed.TotalMinutes, ([GC]::GetTotalMemory($false) / 1MB))
        }
        $out = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) { $out[$property.Name] = $property.Value }
        foreach ($column in $calculatedColumnsToAdd) { $out[$column] = '' }
        foreach ($legacyColumnName in $rawCalculatedColumnRenames.Keys) {
            if (-not $out.Contains($legacyColumnName)) { continue }
            $canonicalColumnName = $rawCalculatedColumnRenames[$legacyColumnName]
            if (-not $out.Contains($canonicalColumnName) -or [string]::IsNullOrWhiteSpace([string]$out[$canonicalColumnName])) {
                $out[$canonicalColumnName] = V $row @($legacyColumnName)
            }
            [void]$out.Remove($legacyColumnName)
        }

        $computerName = NK (V $row @('Name','SamAccountName','DNSHostName'))
        $objectGuid = GK (V $row @('ObjectGUID'))
        $intune = GetMap $intuneByAad $objectGuid
        $intuneDeviceId = GK (V $intune @('Device ID','DeviceId'))
        $entra = GetMap $entraByDeviceId $objectGuid
        $local = GetMap $localByAad $objectGuid
        $wuAnySelection = Resolve-IdentityWindowsUpdateRows @(GetMap $wuAnyByDevice $intuneDeviceId)
        $wuAny = $wuAnySelection.Row
        $wuAnchorSelection = Resolve-IdentityWindowsUpdateRows @(GetMap $wuAnchorByDevice $intuneDeviceId)
        $wuAnchor = $wuAnchorSelection.Row
        $wu24Selection = Resolve-IdentityWindowsUpdateRows @(GetMap $wu24ByDevice $intuneDeviceId)
        $wu24 = $wu24Selection.Row
        $wu25Selection = Resolve-IdentityWindowsUpdateRows @(GetMap $wu25ByDevice $intuneDeviceId)
        $wu25 = $wu25Selection.Row

        $osVersionM365 = V $intune @('OS version','OSVersion','OperatingSystemVersion')
        $friendly = FriendlyOs (V $row @('operatingSystemVersion','OperatingSystemVersion')) $osVersionM365
        $osUpper = $friendly.ToUpperInvariant()
        $isWin11 = $osUpper.Contains('WINDOWS 11')
        $isWin10 = $osUpper.Contains('WINDOWS 10')
        $is24 = ($friendly -eq 'Windows 11 24H2')
        $is25 = ($friendly -eq 'Windows 11 25H2')
        $upgradeRaw = V $wuAny @('UpgradeEligibility','UpgradeEligibilityLabel')
        if (-not $upgradeRaw) { $upgradeRaw = V $intune @('UpgradeEligibility','UpgradeEligibilityLabel') }
        $upgradeLabel = Eligibility $upgradeRaw
        if ($isWin11) { $upgradeLabel = 'Upgraded' }
        if (-not $upgradeLabel) { $upgradeLabel = 'Unknown' }

        $totalGb = Gb (V $intune @('Total storage','TotalStorage','TotalStorageGB'))
        $freeGb = Gb (V $intune @('Free storage','FreeStorage','FreeStorageGB'))
        $freeNum = Num $freeGb
        $totalNum = Num $totalGb
        $freePercent = ''
        if ($null -ne $freeNum -and $null -ne $totalNum -and $totalNum -gt 0) { $freePercent = (($freeNum / $totalNum) * 100).ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture) }
        $freeNotEnough = if ($null -eq $freeNum) { '' } elseif ($freeNum -lt 40) { 'True' } else { 'False' }

        $primaryUser = V $intune @('Primary user UPN','Primary user email address','UserPrincipalName','UPN')
        $primaryKey = K $primaryUser
        $primaryUserRow = GetMap $usersByUpn $primaryKey
        $licenseValue = ''
        if ($primaryKey) { if ($licensedUsersByUpn.ContainsKey($primaryKey)) { $licenseValue = 'Licensed' } elseif ($licenseRows.Count -gt 0) { $licenseValue = 'No License (Unlicensed)' } }

        $deviceActivity = @((Read-IdentityDate $row 'LastLogonDate' 'LastLogonTimestamp'))
        $adLastLogon = if ($deviceActivity[0].Utc) { $deviceActivity[0].Utc.ToString('o') } else { '' }
        # Legacy device enrichment CSVs lack independent source-zone proof.
        $nativeEntra = GetMap $nativeEntraById $objectGuid
        $nativeManaged = GetMap $nativeManagedByAad $objectGuid
        $entraProof = Read-IdentityDate $nativeEntra 'ApproximateLastSignInDateTime'
        $intuneProof = Read-IdentityDate $nativeManaged 'LastSyncDateTime'
        $deviceActivity = @(Get-IdentityDeviceActivitySet $row $nativeEntra $nativeManaged)
        $device90 = Merge-IdentityActivity $deviceActivity 90
        $device45 = Merge-IdentityActivity $deviceActivity 45
        $device30 = Merge-IdentityActivity $deviceActivity 30
        $registrationProof = Read-IdentityDate $nativeEntra 'RegistrationDateTime'
        $registration = if ($registrationProof.Utc) { $registrationProof.Utc.ToString('o') } else { '' }
        $entraSignIn = if ($entraProof.Utc) { $entraProof.Utc.ToString('o') } else { '' }
        $intuneLastCheckIn = if ($intuneProof.Utc) { $intuneProof.Utc.ToString('o') } else { '' }
        $lastActive = if ($device90.Utc) { $device90.Utc.ToString('o') } else { '' }
        $out['ActivityEvidenceCoverage'] = $device90.Coverage
        $out['ActivityEvidenceStatus'] = $device90.Status
        $out['WindowsUpdateTemporalSelectionStatus'] = $wuAnySelection.Status
        $out['WindowsUpdateAnchorSelectionStatus'] = $wuAnchorSelection.Status
        $out['WindowsUpdate24H2SelectionStatus'] = $wu24Selection.Status
        $out['WindowsUpdate25H2SelectionStatus'] = $wu25Selection.Status
        $out['ActivityScopeExclusions'] = $device90.Exclusions
        $out['ActivityConclusionScope'] = $device90.ConclusionScope

        # SWI owns current readiness evidence; missing is never healthy.
        $issueLevel = 'Not evaluated'

        $entraExists = $null -ne $entra
        $entraByNameCount = MapCount $entraByNameAll $computerName
        $entraByGuidCount = MapCount $entraByDeviceIdAll $objectGuid
        $entraNameList = ''
        if ($entraByDeviceIdAll.ContainsKey($objectGuid)) { $entraNameValues = New-Object System.Collections.Generic.List[string]; foreach ($entraMatch in $entraByDeviceIdAll[$objectGuid]) { $display = V $entraMatch @('DisplayName'); if ($display) { [void]$entraNameValues.Add($display) } }; $entraNameList = ($entraNameValues -join ';') }
        $correlation = if (-not $entraExists) { 'Not found in Entra' } elseif ($registration) { 'Found (correlated, registered)' } else { 'Found (correlated, missing registration date)' }

        $canonical = V $row @('CanonicalName','DistinguishedName')
        $lastOu = ''
        if ($canonical) { $parts = @($canonical -split '[,/]' | Where-Object { $_ }); if ($parts.Count -ge 2) { $lastOu = $parts[$parts.Count - 2] } }
        $isDisabledOu = if ($canonical -match '(?i)disabled|desactive|disabled objects') { 'True' } else { 'False' }
        $model = V $intune @('Model')
        $manufacturer = V $intune @('Manufacturer')
        $isVirtual = if (("{0} {1}" -f $manufacturer, $model) -match '(?i)vmware|virtualbox|hyper-v|qemu|kvm|virtual machine') { 'True' } else { 'False' }

        $distinguishedName = V $row @('DistinguishedName')
        $buildNumber = GetBuildNumber (V $row @('operatingSystemVersion','OperatingSystemVersion')) $osVersionM365
        $operatingSystemMajorCurrent = GetOsMajor (V $row @('OperatingSystem'))
        $isLtscOs = (([string](V $row @('OperatingSystem'))).Trim().ToLowerInvariant() -match 'ltsc|ltsb')
        $isLtscBuildFallback = ($operatingSystemMajorCurrent -eq 'Windows 10' -and $null -ne $buildNumber -and @(10240,14393,17763,19044) -contains [int]$buildNumber)
        $isWindows10Ltsc = $isLtscOs -or $isLtscBuildFallback
        $osMinToUpdateW11 = ($operatingSystemMajorCurrent -eq 'Windows 11') -or ($operatingSystemMajorCurrent -eq 'Windows 10' -and $null -ne $buildNumber -and [int]$buildNumber -ge 19041)
        $orgUnit = GetOrganizationalUnitCode $row
        $domainAndOu = if ((V $row @('DomainNameShort')) -and $orgUnit) { ('{0}-{1}' -f (V $row @('DomainNameShort')), $orgUnit) } else { '' }
        $site = Get-SmartM365WorkplaceSite -SiteClassification $siteClassification -SiteCode $orgUnit
        $labelFromEntity = if ($null -ne $site) { $site.Name } else { 'NOTFOUND' }
        $typeEtablissement = if ($null -ne $site) { $site.Type } else { 'NOTFOUND' }
        $typeEntity = if ($null -ne $site) { $site.IntegrationStatus } else { '' }
        if (-not $typeEntity) { $typeEntity = GetTypeEntityFallback (V $row @('DomainNameShort')) $distinguishedName }
        $subnet = GetSubnet (V $row @('IPv4Address'))
        $isInOrganizationOu = [string]($distinguishedName.ToUpperInvariant().Contains('OU=ORGANIZATION'))
        $typeEtablissementUpper = $typeEtablissement.Trim().ToUpperInvariant()
        $inHeadquartersOu = $false
        $distinguishedNameUpper = $distinguishedName.ToUpperInvariant()
        foreach ($fragment in $headquartersOuFragments) { if ($distinguishedNameUpper.Contains($fragment)) { $inHeadquartersOu = $true; break } }
        $isTargetOuHq = [string](($typeEtablissementUpper -eq 'HQ') -or $inHeadquartersOu)
        $isNotTargetOuHq = [string](($typeEtablissementUpper -ne 'HQ') -and ((B $isTargetOuHq) -eq 'False') -and ($distinguishedName -notmatch 'OU=DISABLED_OBJECTS'))
        $serverAdEntra = GetMap $serverAdEntraByName $computerName
        $serverAdDeviceId = GK (V $serverAdEntra @('DeviceId'))
        $hasDifferentDeviceId = if (-not $serverAdDeviceId) { 'Not found in Entra' } elseif ($serverAdDeviceId -ne $objectGuid) { 'Different GUID' } else { 'Match' }

        $primaryUserSam = if ($primaryUserRow) { V $primaryUserRow @('SamAccountName') } else { '' }
        $lastLoggedUserFromM365 = $primaryUserSam
        $lastLoggedUserFromSentinel = ''
        $lastLoggedUser = if (-not (IsUnknownUser $lastLoggedUserFromM365)) { (K $lastLoggedUserFromM365) } elseif (-not (IsUnknownUser $lastLoggedUserFromSentinel)) { (K $lastLoggedUserFromSentinel) } else { '' }
        $samUserRow = GetMap $usersBySam $lastLoggedUser
        $lastLoggedUserDomain = ''
        if ($lastLoggedUser) {
            if ($primaryUserRow) { $lastLoggedUserDomain = NormalizeDomainSam (V $primaryUserRow @('DomainAndSam')) }
            if (-not $lastLoggedUserDomain -and $samUserRow) { $lastLoggedUserDomain = NormalizeDomainSam (V $samUserRow @('DomainAndSam')) }
            if (-not $lastLoggedUserDomain) { $lastLoggedUserDomain = NormalizeDomainSam ('{0}\{1}' -f (V $row @('DomainNameShort')), $lastLoggedUser) }
        }
        $lastUserRow = GetMap $usersByDomainAndSam $lastLoggedUserDomain
        if (-not $lastUserRow) { $lastUserRow = $primaryUserRow }
        $lastUserUpnKey = K (V $lastUserRow @('UserPrincipalName','EmailAddress'))
        $lastUserMailbox = GetMap $mailboxByUpn $lastUserUpnKey
        $lastUserMailboxSize = if ($lastUserUpnKey -and $mailboxSizeByUpn.ContainsKey($lastUserUpnKey)) { [string]$mailboxSizeByUpn[$lastUserUpnKey] } else { '' }
        $lastUserLicenseGroups = if ($lastUserUpnKey -and $licenseGroupByUpn.ContainsKey($lastUserUpnKey)) { (($licenseGroupByUpn[$lastUserUpnKey] | Select-Object -Unique) -join ';') } else { '?' }
        $lastUserPrimarySmtp = V $lastUserRow @('EmailAddress','UserPrincipalName')
        $lastUserOuPath = V $lastUserRow @('CanonicalName','DistinguishedName')
        $lastUserLastOu = GetLastOuName $lastUserOuPath

        $ueMap = @{ '0'='UPGRADED'; '2'='NOT CAPABLE'; '3'='CAPABLE' }
        $win11GlobalResultOld = if ($isWin11) { 'UPGRADED' } elseif ($ueMap.ContainsKey(([string]$upgradeRaw).Trim())) { $ueMap[([string]$upgradeRaw).Trim()] } else { 'UNKNOWN' }
        $issueScore = ''
        $shortOsValue = if ($isWin11) { 'Windows 11' } elseif ($isWin10) { 'Windows 10' } else { V $row @('OperatingSystem') }
        $w11EligibilityValue = if ($isWin11) { 'Upgraded' } else { $upgradeLabel }
        $lastSignInMergedValue = $lastActive
        $mergedDaysValue = Days $lastSignInMergedValue
        $lastSignInActive45Value = ($device45.Recent -eq $true)
        $physicalMemoryGb = V $intune @('PhysicalMemoryGB')
        $migrationPhase = GetMigrationPhase $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $model
        $migrationReason = GetMigrationPhaseReason $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $model
        $migrationAction = GetMigrationAction $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $physicalMemoryGb
        if ($null -eq $device45.Recent -and $migrationPhase -ne 'OUT OF SCOPE') {
            $migrationPhase = 'UNDETERMINED'
            $migrationReason = 'Activity date evidence is incomplete or unqualified'
            $migrationAction = '6 - To qualify'
        }
        if ($wuAnySelection.Status -like 'Ambiguous:*' -and -not $isWin11 -and $migrationPhase -ne 'OUT OF SCOPE') {
            $migrationPhase = 'UNDETERMINED'
            $migrationReason = 'Windows Update observations are ambiguous; independent inventory eligibility only'
            $migrationAction = '6 - To qualify'
        }
        $out['ObjectGuidNormalized'] = $objectGuid
        $out['OperatingSystemBuildNumber'] = if ($null -ne $buildNumber) { [string]$buildNumber } else { '' }
        $out['IsWindows10LtscByName'] = [string]$isLtscOs
        $out['IsWindows10LtscByBuild'] = [string]$isLtscBuildFallback
        $out['IsWindows10Ltsc'] = [string]$isWindows10Ltsc
        $out['MeetsWindows11MinimumOsBuild'] = [string]$osMinToUpdateW11
        $out['OrganizationalUnitCode'] = $orgUnit
        $out['DomainOrganizationalUnitKey'] = $domainAndOu
        $out['EntityLabel'] = $labelFromEntity
        $out['Subnet'] = $subnet
        $out['EntitySiteType'] = $typeEtablissement
        $out['IsInProductionOU'] = $isInOrganizationOu
        $out['IsInTargetHeadquartersOU'] = $isTargetOuHq
        $out['IsOutsideTargetHeadquartersOU'] = $isNotTargetOuHq
        $out['HasDifferentDeviceId'] = $hasDifferentDeviceId
        $out['OperatingSystemDisplayName'] = $friendly
        $out['OperatingSystemShortName'] = $shortOsValue
        $out['IsWindows11'] = [string]$isWin11
        $out['IsWindows1124H2'] = [string]$is24
        $out['IsWindows1125H2'] = [string]$is25
        $out['NeedsWindows11Upgrade'] = [string]($isWin10 -and $upgradeLabel -ne 'Not Capable')
        $out['Windows11UpgradeEligibility'] = $w11EligibilityValue
        $out['Windows11UpgradeEligibilitySource'] = if ($null -ne $wuAny) { 'Intune Windows Update status' } elseif ($null -ne $intune) { 'Intune devices inventory' } else { '' }
        $out['IntuneWindows11EligibilityCode'] = $upgradeRaw
        $out['IntuneWindows11EligibilityLabel'] = $upgradeLabel
        $out['IntuneOperatingSystemVersion'] = $osVersionM365
        $out['ExistsInIntune'] = [string]($null -ne $intune)
        $out['IntuneDeviceId'] = V $intune @('Device ID','DeviceId')
        $out['IsRegisteredInIntune'] = V $intune @('Intune registered')
        $out['IntuneLastSyncDateTime'] = $intuneLastCheckIn
        $out['IntuneEnrollmentDateTime'] = $( $enrollmentProof = Read-IdentityDate $nativeManaged 'EnrolledDateTime'; if ($enrollmentProof.Utc) { $enrollmentProof.Utc.ToString('o') } else { '' } )
        $out['IntuneEnrollmentDate'] = if ($out['IntuneEnrollmentDateTime']) { ([datetime]::Parse($out['IntuneEnrollmentDateTime'])).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { '' }
        $out['IntuneDeviceModel'] = $model
        $out['Manufacturer'] = $manufacturer
        $out['IntuneOwnership'] = V $intune @('Ownership')
        $out['IntuneComplianceState'] = V $intune @('Compliance','IsCompliant')
        $out['IntunePrimaryUserPrincipalName'] = V $intune @('Primary user email address','Primary user UPN')
        $out['IntuneLastLoggedOnUser'] = $lastLoggedUserFromM365
        $out['ResolvedPrimarySmtpAddress'] = if ($primaryUserRow) { V $primaryUserRow @('EmailAddress','UserPrincipalName') } else { $primaryUser }
        $out['M365LicenseForAdUser'] = $licenseValue
        $out['M365LicenseForIntunePrimaryUser'] = $licenseValue
        $out['IntunePrimaryUserOrganizationalUnitPath'] = V $primaryUserRow @('CanonicalName','DistinguishedName')
        $out['IntunePrimaryUserOrganizationalUnitName'] = if ($lastUserLastOu) { $lastUserLastOu } else { '?' }
        $out['HasMailboxForIntunePrimaryUser'] = [string]($null -ne $lastUserMailbox)
        $out['IntunePrimaryUserMailboxSizeGB'] = $lastUserMailboxSize

        $out['IntuneDiskTotalGB'] = $totalGb
        $out['IntuneDiskFreeGB'] = $freeGb
        $out['DiskFreePercent'] = $freePercent
        $out['HasInsufficientStorageForWindows11Upgrade'] = $freeNotEnough
        $out['PhysicalMemoryGB'] = $physicalMemoryGb

        $out['EntraObjectId'] = V $entra @('ObjectId')
        $out['EntraDeviceIdNormalized'] = GK (V $entra @('DeviceId'))
        $out['EntraDeviceIdMatchCount'] = [string]$entraByGuidCount
        $out['EntraRegistrationDateTime'] = $registration
        $out['EntraApproximateLastSignInDateTime'] = $entraSignIn
        $out['EntraDaysSinceLastSignIn'] = Days $entraSignIn
        $out['EntraHardwareIdDeviceCount'] = $hardwareCount
        $out['EntraCorrelationStatus'] = $correlation
        $out['ExistsInEntraByObjectGuid'] = [string]$entraExists
        $out['IsEntraRegistrationPending'] = [string]($entraExists -and ((B (V $entra @('IsPending'))) -eq 'True'))
        $out['EntraNameMatchCount'] = [string]$entraByNameCount
        $out['EntraMatchCountByObjectGuid'] = [string]$entraByGuidCount
        $out['EntraNameMatchesByObjectGuidList'] = $entraNameList
        $out['EntraLastSignInStatus'] = if ($entraSignIn) { 'Known' } else { 'Unknown / unqualified activity' }
        $out['IsEntraLastSignInOlderThan90Days'] = if ($out['EntraDaysSinceLastSignIn'] -eq '') { '' } else { [string]([int]$out['EntraDaysSinceLastSignIn'] -gt 90) }
        $out['IsEntraLastSignInWithin90Days'] = if ($out['EntraDaysSinceLastSignIn'] -eq '') { '' } else { [string]([int]$out['EntraDaysSinceLastSignIn'] -le 90) }

        $out['ResolvedLastSignInDateTime'] = $lastSignInMergedValue
        $mergedDays = $mergedDaysValue
        $out['IsActiveInLast90Days'] = if ($null -eq $device90.Recent) { '' } else { [string]$device90.Recent }
        $out['IsActiveInLast45Days'] = if ($null -eq $device45.Recent -and $migrationPhase -ne 'OUT OF SCOPE') { '' } else { [string]$device45.Recent }
        $out['IsActiveInLast30Days'] = if ($null -eq $device30.Recent) { '' } else { [string]$device30.Recent }

        SetWu $out $wuAnchor 'AutopatchFeatureUpdateAnchor'
        SetWu $out $wu24 'Windows1124H2'
        SetWu $out $wu25 'Windows1125H2'

        $out['SecureBootStatus'] = V $local @('SecureBootStatus')
        $out['BiosVersion'] = V $local @('BIOSVersion')
        $out['BiosKey'] = $out['BiosVersion']
        $out['BiosModel'] = $model
        $out['LastRebootDateTime'] = V $local @('LastRebootDate','LastBootUpTime')
        $out['LastActivityDateTime'] = $lastActive
        $out['OrganizationalUnitLeafName'] = $lastOu
        $out['IsInDisabledObjectsOU'] = $isDisabledOu
        $out['IsVirtualMachine'] = $isVirtual
        $out['EntityIntegrationType'] = $typeEntity
        $out['Windows11MigrationPhase'] = $migrationPhase
        $out['Windows11MigrationPhaseReason'] = $migrationReason
        $out['Windows11MigrationAction'] = $migrationAction
        $out['MigrationIssueScore'] = $issueScore
        $out['DeviceMigrationIssueSeverity'] = $issueLevel
        $out['IsWindows11AndManagedByIntune'] = [string]($isWin11 -and ($null -ne $intune))
        $out['AdLastLogonCategory'] = Format-IdentityActivityCategory (Merge-IdentityActivity @($deviceActivity[0]) 90) 90
        $out['AdLastLogon30DayCategory'] = Format-IdentityActivityCategory (Merge-IdentityActivity @($deviceActivity[0]) 30) 30
        $out['EntraLastSignInCategory'] = GetDateCategory $entraSignIn 90
        $out['EntraLastSignIn30DayCategory'] = GetDateCategory $entraSignIn 30
        $out['LastActivityCategory'] = Format-IdentityActivityCategory $device90 90
        $out['LastActivity30DayCategory'] = Format-IdentityActivityCategory $device30 30
        $out['LastRebootCategory'] = 'Unknown / unqualified activity'
        $out['LastReboot30DayCategory'] = 'Unknown / unqualified activity'

        [void]$enrichedRows.Add([pscustomobject]$out)
    }

    # Sources and lookups are no longer needed; release them before the CSV writer copies the enriched rows.
    $sourceRows = $null
    $intuneRows = $null; $entraRows = $null; $hardwareConflictRows = $null; $localSystemRows = $null; $windowsUpdateRows = $null
    $adUserRows = $null; $exoMailboxRows = $null; $exoMailboxStatsRows = $null; $localMailboxRows = $null; $remoteMailboxRows = $null; $win11IssueRows = $null
    $intuneByAad = $null; $entraByDeviceId = $null; $entraByDeviceIdAll = $null; $entraByNameAll = $null; $hardwareCounts = $null; $localByAad = $null
    $wuAnyByDevice = $null; $wuAnchorByDevice = $null; $wu24ByDevice = $null; $wu25ByDevice = $null
    $usersByUpn = $null; $licensedUsersByUpn = $null; $usersBySam = $null; $usersByDomainAndSam = $null; $licenseByUpn = $null; $licenseGroupByUpn = $null
    $mailboxByUpn = $null; $mailboxSizeByUpn = $null; $serverAdEntraByName = $null
    $row = $null; $out = $null; $intune = $null; $entra = $null; $local = $null; $wuAny = $null; $wuAnchor = $null; $wu24 = $null; $wu25 = $null
    $primaryUserRow = $null; $samUserRow = $null; $lastUserRow = $null; $lastUserMailbox = $null; $serverAdEntra = $null
    [System.GC]::Collect()
    WriteLog -Message ("AD computers enrichment loop completed in {0:N1} min; writing CSV; managed memory {1:N0} MB" -f $loopTimer.Elapsed.TotalMinutes, ([GC]::GetTotalMemory($false) / 1MB))

    $enrichedCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Computers_AllDomains.csv'
    $exportColumns = @()
    if ($enrichedRows.Count -eq 0) {
        $nativeColumns = @(Get-SmartM365AdCsvColumns -Path $CombinedComputersCsv)
        $exportColumns = @($nativeColumns | Where-Object { -not $rawCalculatedColumnRenames.Contains($_) })
        $exportColumns += @($calculatedColumns | Where-Object { $exportColumns -inotcontains $_ })
        $exportColumns += @($rawCalculatedColumnRenames.Values | Where-Object { $exportColumns -inotcontains $_ })
    }
    Write-SmartM365CsvAtomically -Data $enrichedRows.ToArray() -Path $enrichedCsv -Columns $exportColumns -Encoding UTF8
    if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
        $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) { if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) { [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath) } }
    }
    [void]$global:csvGeneratedPaths.Add($enrichedCsv)
    WriteLog -Message ("AD computers enriched CSV generated: {0} ({1} row(s), {2} enriched column(s))" -f $enrichedCsv, $enrichedRows.Count, $calculatedColumns.Count)
    return $enrichedCsv
}
