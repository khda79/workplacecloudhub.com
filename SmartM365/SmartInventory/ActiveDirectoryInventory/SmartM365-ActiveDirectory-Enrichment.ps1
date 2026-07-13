<#
.SYNOPSIS
    Builds enriched Active Directory computer CSV columns required by the SmartWorkplace Power BI model.
#>

function Invoke-SmartM365AdComputersEnrichedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CombinedComputersCsv,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $false)][string]$LatestFolderPath,
        [Parameter(Mandatory = $false)][string]$WindowsUpdateAnchorPolicyId,
        [Parameter(Mandatory = $false)][string]$WindowsUpdate24H2PolicyId,
        [Parameter(Mandatory = $false)][string]$WindowsUpdate25H2PolicyId
    )

    if (-not (Test-Path -LiteralPath $CombinedComputersCsv)) {
        WriteLog -Message ("AD computers enriched CSV skipped because source CSV is missing: {0}" -f $CombinedComputersCsv) -Level "WARNING"
        return $null
    }

    $columnsFile = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-ActiveDirectory-EnrichedColumns.ps1'
    if (Test-Path -LiteralPath $columnsFile) {
        $calculatedColumns = @(. $columnsFile | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    else {
        $calculatedColumns = @('ObjectGUID_Norm')
        WriteLog -Message ("AD enrichment column list not found, using ObjectGUID_Norm only: {0}" -f $columnsFile) -Level "WARNING"
    }

    function V($Row, [string[]]$Names) {
        if ($null -eq $Row) { return '' }
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
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $date = [datetime]::MinValue
        if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
        if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
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
        if ($null -eq $date) { return 'Unknown / Never logged in' }
        if ($date.ToUniversalTime().Date -lt ([datetime]::UtcNow.Date.AddDays(-1 * $LimitDays))) { return ("Older than {0} days" -f $LimitDays) }
        return ("Within last {0} days" -f $LimitDays)
    }

    function LoadCsv($Name) {
        $candidates = New-Object System.Collections.Generic.List[string]
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) { [void]$candidates.Add((Join-Path $LatestFolderPath $Name)) }
        if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) { [void]$candidates.Add((Join-Path $OutputFolder $Name)) }
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate) {
                $rows = @(Import-Csv -LiteralPath $candidate -ErrorAction Stop)
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
        if (-not $Map.ContainsKey($Key)) { $Map[$Key] = $Row; return }
        $oldDate = Dt (V $Map[$Key] @('ExportDateTime','ReadinessExportDateTime','LastSyncDateTime','Last check-in'))
        $newDate = Dt (V $Row @('ExportDateTime','ReadinessExportDateTime','LastSyncDateTime','Last check-in'))
        if ($null -eq $oldDate -or ($null -ne $newDate -and $newDate -gt $oldDate)) { $Map[$Key] = $Row }
    }

    function SetWu($Output, $Row, [string]$Suffix) {
        $Output[("WU_AggregateState_loc_Policy_{0}" -f $Suffix)] = V $Row @('AggregateState_loc','AggregateState')
        $Output[("WU_CurrentDeviceUpdateStatus_loc_Policy_{0}" -f $Suffix)] = V $Row @('CurrentDeviceUpdateStatus_loc','CurrentDeviceUpdateStatus')
        $Output[("WU_LatestAlertMessage_loc_Policy_{0}" -f $Suffix)] = V $Row @('LatestAlertMessage_loc','LatestAlertMessage')
        $Output[("WU_BlockingReason_loc_Policy_{0}" -f $Suffix)] = V $Row @('BlockingReason','LatestAlertMessage_loc','LatestAlertMessage')
        $Output[("WU_RiskBucket_loc_Policy_{0}" -f $Suffix)] = V $Row @('RiskBucket')
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
        $countryFromDn = ''
        if ($domainClean -eq 'GRP' -and $dnClean.Contains(',OU=COUNTRIES')) {
            $beforeCountries = $dnClean.Substring(0, $dnClean.IndexOf(',OU=COUNTRIES'))
            if ($beforeCountries.Length -ge 2) { $countryFromDn = $beforeCountries.Substring($beforeCountries.Length - 2, 2) }
        }
        if ($domainClean -eq 'ORPEA_01') { return 'CORP' }
        if ($domainClean -eq 'GRP' -and @('BE','LU') -notcontains $countryFromDn) { return 'NOT-INTEGRATED' }
        if (@('BE','CH','DE','ES','FR','IT','LU','PL','PT') -contains $domainClean) { return 'INTEGRATED' }
        if (@('AT','CZ','NL') -contains $domainClean) { return 'NOT-INTEGRATED' }
        return 'UNKNOWN'
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
        if ($domainShortText -eq 'ORPEA_01') { return 'OUT OF SCOPE' }
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

    function GetMigrationPhaseReason($Eligibility, [bool]$OsMinOk, [bool]$IsLtsc, [bool]$IsEnabled, $ShortOs, $TypeEntity, [bool]$InDisabledOu, $DomainShort, [bool]$IsActiveLast45Days, [bool]$IsInIntune, $FriendlySnapshot, $ModelName) {
        $eligibilityText = ([string]$Eligibility).Trim().ToUpperInvariant()
        if (-not $eligibilityText) { $eligibilityText = 'UNKNOWN' }
        $shortOsText = ([string]$ShortOs).Trim()
        $isWindows11Value = $shortOsText -eq 'Windows 11'
        $isWindows10Value = $shortOsText -eq 'Windows 10'
        $isSupportedClientOs = $isWindows10Value -or $isWindows11Value
        $typeEntityText = ([string]$TypeEntity).Trim().ToUpperInvariant()
        $domainShortText = ([string]$DomainShort).Trim().ToUpperInvariant()
        $snapshotIsW10 = ([string]$FriendlySnapshot).Trim().ToUpperInvariant().StartsWith('WINDOWS 10')
        $modelText = ([string]$ModelName).Trim().ToUpperInvariant()
        $isVmValue = $modelText -match 'VMWARE|VIRTUAL|HYPER-V|VIRTUALBOX|KVM|XEN'
        if (-not $isSupportedClientOs) { return 'OS != Windows 10/11 -> OUT OF SCOPE' }
        if ($isVmValue) { return 'Virtual machine detected from Model -> OUT OF SCOPE' }
        if ($typeEntityText -eq 'NOT-INTEGRATED') { return 'TypeEntity = NOT-INTEGRATED -> OUT OF SCOPE' }
        if ($typeEntityText -eq 'EXCLUDED') { return 'TypeEntity = EXCLUDED -> OUT OF SCOPE' }
        if (-not $IsEnabled) { return 'Device disabled -> OUT OF SCOPE' }
        if ($domainShortText -eq 'ORPEA_01') { return 'Domain = ORPEA_01 -> OUT OF SCOPE' }
        if ($InDisabledOu) { return 'Disabled Objects OU -> OUT OF SCOPE' }
        if ((@('UNKNOWN','UNDETERMINED','NOT CAPABLE','CAPABLE','UPGRADED') -contains $eligibilityText) -and (-not $IsActiveLast45Days) -and (-not ($isWindows11Value -and $IsInIntune))) { return ('Eligibility = {0} + Inactive (no sign-in last 45 days) -> PHASE 2' -f $eligibilityText) }
        if ((@('UNKNOWN','UNDETERMINED','NOT CAPABLE','CAPABLE','UPGRADED') -contains $eligibilityText) -and $IsLtsc -and $isWindows10Value) { return ('Eligibility = {0} + Windows 10 LTSC -> PHASE 2' -f $eligibilityText) }
        if ($isWindows11Value -and $snapshotIsW10) { return 'Migrated from Windows 10 -> PHASE 1' }
        if ($isWindows11Value -and -not $snapshotIsW10) { return 'Already Windows 11 (pre-existing) -> PHASE 1' }
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
        if ($domainShortText -eq 'ORPEA_01') { return '6 - To qualify' }
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
    $intuneRows = @(LoadCsv 'Intune_Devices_Inventory.csv')
    $entraRows = @(LoadCsv 'M365_Entra_Devices.csv')
    $hardwareConflictRows = @(LoadCsv 'M365_Entra_Devices_HardwareIdConflicts.csv')
    $localSystemRows = @(LoadCsv 'Intune_Devices_LocalSystem.csv')
    $windowsUpdateRows = @(LoadCsv 'Intune_WindowsUpdate_Status.csv')
    $adUserRows = @(LoadCsv 'AD_Users_AllDomains.csv')
    $licenseRows = @(LoadCsv 'M365_Licenses_Users.csv')
    $exoMailboxRows = @(LoadCsv 'Exchange_EXO_Mailboxes_AllDomains.csv')
    $exoMailboxStatsRows = @(LoadCsv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv')
    $localMailboxRows = @(LoadCsv 'Exchange_OnPrem_Mailboxes_AllDomains.csv')
    $remoteMailboxRows = @(LoadCsv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv')
    $entityRows = @(LoadCsv 'EntityDirectories.csv')
    $win11IssueRows = @(LoadCsv 'Mig_Win11Migration_Issues_Expanded_ByScript.csv')
    if ($win11IssueRows.Count -eq 0) { $win11IssueRows = @(LoadCsv 'Intune_Windows11_Readiness_Issues.csv') }

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
        if ($policyKey -eq $anchorKey) { AddLatest $wuAnchorByDevice $deviceKey $row }
        if ($policyKey -eq $policy24Key) { AddLatest $wu24ByDevice $deviceKey $row }
        if ($policyKey -eq $policy25Key) { AddLatest $wu25ByDevice $deviceKey $row }
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

    $entityByCode = @{}
    foreach ($row in $entityRows) {
        $entityKey = K (V $row @('Entity code Text 6 digits','EntityCode','OrganizationalUnit'))
        AddMap $entityByCode $entityKey $row
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


    $issuesByGuid = @{}; $issueScoreByGuid = @{}
    foreach ($row in $win11IssueRows) {
        $guidKey = GK (V $row @('ObjectGUID_Norm','ObjectGUID'))
        if (-not $guidKey) { continue }
        AddMulti $issuesByGuid $guidKey $row
        $category = V $row @('IssueCategory','Category')
        $score = 0
        if ($category -match '^1') { $score = 3 }
        elseif ($category -match '^2') { $score = 2 }
        elseif ($category -match '^3') { $score = 1 }
        if (-not $issueScoreByGuid.ContainsKey($guidKey)) { $issueScoreByGuid[$guidKey] = 0 }
        $issueScoreByGuid[$guidKey] = [int]$issueScoreByGuid[$guidKey] + $score
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
    $calculatedColumnsToAdd = @($calculatedColumns | Where-Object { -not $sourceColumnSet.Contains([string]$_) })
    $enrichedRows = New-Object System.Collections.Generic.List[object]

    foreach ($row in $sourceRows) {
        $out = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) { $out[$property.Name] = $property.Value }
        foreach ($column in $calculatedColumnsToAdd) { $out[$column] = '' }

        $computerName = NK (V $row @('Name','SamAccountName','DNSHostName'))
        $objectGuid = GK (V $row @('ObjectGUID'))
        $intune = GetMap $intuneByAad $objectGuid
        $intuneDeviceId = GK (V $intune @('Device ID','DeviceId'))
        $entra = GetMap $entraByDeviceId $objectGuid
        $local = GetMap $localByAad $objectGuid
        $wuAny = GetMap $wuAnyByDevice $intuneDeviceId
        $wuAnchor = GetMap $wuAnchorByDevice $intuneDeviceId
        $wu24 = GetMap $wu24ByDevice $intuneDeviceId
        $wu25 = GetMap $wu25ByDevice $intuneDeviceId

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

        $registration = Fmt (V $entra @('RegistrationDateTime'))
        $entraSignIn = Fmt (V $entra @('ApproximateLastSignInDateTime'))
        $adLastLogon = Fmt (V $row @('LastLogonDate'))
        $intuneLastCheckIn = Fmt (V $intune @('Last check-in','LastSyncDateTime'))
        $lastActive = $entraSignIn
        if (-not $lastActive) { $lastActive = $intuneLastCheckIn }
        if (-not $lastActive) { $lastActive = $adLastLogon }

        $hardwareId = K (V $entra @('HardwareId'))
        $hardwareCount = if ($hardwareId -and $hardwareCounts.ContainsKey($hardwareId)) { [string]$hardwareCounts[$hardwareId] } else { '' }
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
        $entity = GetMap $entityByCode (K $orgUnit)
        $labelFromEntity = if ($entity) { V $entity @('Entity label','EntityLabel') } else { 'NOTFOUND' }
        $typeEtablissement = if ($entity) { V $entity @('TypeEtablissement') } else { 'NOTFOUND' }
        $typeEntity = if ($entity) { V $entity @('Entity Type','EntityType') } else { '' }
        if (-not $typeEntity) { $typeEntity = GetTypeEntityFallback (V $row @('DomainNameShort')) $distinguishedName }
        $subnet = GetSubnet (V $row @('IPv4Address'))
        $isInOrganizationOu = [string]($distinguishedName.ToUpperInvariant().Contains('OU=ORGANIZATION'))
        $isTargetOuHq = [string](($distinguishedName -match 'OU=020001-MADRID') -or ($distinguishedName -match 'OU=120001-HEAD_QUARTER') -or ($distinguishedName -match 'OU=HEADQUARTER'))
        $isNotTargetOuHq = [string](((B $isTargetOuHq) -eq 'False') -and ($distinguishedName -notmatch 'OU=DISABLED_OBJECTS'))
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

        $os20250918 = ''
        $os20251217 = ''
        $osVersion20250918 = ''
        $osVersion20251217 = ''
        $snapshotOsPriority = ''
        $snapshotVersionPriority = ''
        $snapshotMajorPriority = ''
        $friendlySnapshot = ''
        $osChanged = 'False'
        $osChangedText = 'No snapshot'
        $os10To11Changed = 'False'
        $os10To11ChangedText = 'No snapshot'
        $snapshotUsed = 'None'
        $ueMap = @{ '0'='UPGRADED'; '2'='NOT CAPABLE'; '3'='CAPABLE' }
        $win11GlobalResultOld = if ($isWin11) { 'UPGRADED' } elseif ($ueMap.ContainsKey(([string]$upgradeRaw).Trim())) { $ueMap[([string]$upgradeRaw).Trim()] } else { 'UNKNOWN' }
        $issueScore = if ($issueScoreByGuid.ContainsKey($objectGuid)) { [string]$issueScoreByGuid[$objectGuid] } else { '' }
        $issueLevel = 'None'
        if ($issuesByGuid.ContainsKey($objectGuid)) {
            $severityCodes = @($issuesByGuid[$objectGuid] | ForEach-Object { $cat = V $_ @('IssueCategory','Category'); if ($cat -match '^([1-4])') { [int]$Matches[1] } })
            if ($severityCodes.Count -gt 0) {
                switch (($severityCodes | Measure-Object -Minimum).Minimum) { 1 { $issueLevel = 'Critical' } 2 { $issueLevel = 'High' } 3 { $issueLevel = 'Medium' } default { $issueLevel = 'Low' } }
            }
        }

        $shortOsValue = if ($isWin11) { 'Windows 11' } elseif ($isWin10) { 'Windows 10' } else { V $row @('OperatingSystem') }
        $w11EligibilityValue = if ($isWin11) { 'Upgraded' } else { $upgradeLabel }
        $lastSignInMergedValue = if ($entraSignIn) { $entraSignIn } else { $adLastLogon }
        $mergedDaysValue = Days $lastSignInMergedValue
        $lastSignInActive45Value = ($mergedDaysValue -ne '' -and [int]$mergedDaysValue -le 45)
        $physicalMemoryGb = V $intune @('PhysicalMemoryGB')
        $migrationPhase = GetMigrationPhase $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $model
        $migrationReason = GetMigrationPhaseReason $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $friendlySnapshot $model
        $migrationAction = GetMigrationAction $w11EligibilityValue $osMinToUpdateW11 $isWindows10Ltsc (BoolValue (V $row @('Enabled'))) $shortOsValue $typeEntity (BoolValue $isDisabledOu) (V $row @('DomainNameShort')) $lastSignInActive45Value ($null -ne $intune) $physicalMemoryGb
        $out['ObjectGUID_Norm'] = $objectGuid
        $out['BuildNumber'] = if ($null -ne $buildNumber) { [string]$buildNumber } else { '' }
        $out['IsWindows10LTSC_OS'] = [string]$isLtscOs
        $out['IsWindows10LTSC_BuildFallback'] = [string]$isLtscBuildFallback
        $out['IsWindows10LTSC'] = [string]$isWindows10Ltsc
        $out['OSMinToUpdateW11'] = [string]$osMinToUpdateW11
        $out['OSMinToUpdateW11_Num'] = if ($osMinToUpdateW11) { '1' } else { '0' }
        $out['OrganizationalUnit'] = $orgUnit
        $out['LastLogonDateConverted'] = if ($adLastLogon) { $adLastLogon.Substring(0, 10) } else { '' }
        $out['DomainAndOrganizationalUnit'] = $domainAndOu
        $out['OperatingSystem_Major_Current'] = $operatingSystemMajorCurrent
        $out['AccountToDeleteFromAD'] = GetAccountToDeleteFromAd $row
        $out['LabelFromEntity'] = $labelFromEntity
        $out['Subnet'] = $subnet
        $out['TypeEtablissement'] = $typeEtablissement
        $out['Is_In_OrganizationOU'] = $isInOrganizationOu
        $out['IsTargetOU_HQ'] = $isTargetOuHq
        $out['IsNotTargetOU_HQ'] = $isNotTargetOuHq
        $out['HasDifferentDeviceId'] = $hasDifferentDeviceId
        $out['FriendlyOSVersionName'] = $friendly
        $out['FriendlyOSVersionName_Snapshot'] = if ($friendlySnapshot) { $friendlySnapshot } else { $friendly }
        $out['ShortOSName'] = $shortOsValue
        $out['IsWindows11'] = [string]$isWin11
        $out['IsWindows11-Bool'] = [string]$isWin11
        $out['IsWindows1124H2-Bool'] = [string]$is24
        $out['IsWindows1125H2-Bool'] = [string]$is25
        $out['NeedToBeUpgrade'] = [string]($isWin10 -and $upgradeLabel -ne 'Not Capable')
        $out['W11Eligibilty'] = $w11EligibilityValue
        $out['WIN11_GLOBALRESULT'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_INTUNE_ONLY'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_M365_ONLY'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_SOURCE'] = if ($null -ne $wuAny) { 'Intune Windows Update status' } elseif ($null -ne $intune) { 'Intune devices inventory' } else { '' }
        $out['UpgradeEligibility_From_M365'] = $upgradeRaw
        $out['UpgradeEligibility_Label_From_M365'] = $upgradeLabel
        $out['OS_version_From_M365'] = $osVersionM365
        $out['OperatingSystem_Snapshot_Priority'] = $snapshotOsPriority
        $out['OperatingSystemVersion_Snapshot_Priority'] = $snapshotVersionPriority

        $out['ExistsInIntune'] = [string]($null -ne $intune)
        $out['DeviceID_From_M365'] = V $intune @('Device ID','DeviceId')
        $out['Intune_registered_From_M365'] = V $intune @('Intune registered')
        $out['LastcheckIn_date_From_M365'] = $intuneLastCheckIn
        $out['Enrollment_date_From_M365'] = Fmt (V $intune @('Enrollment date','EnrollmentDate'))
        $out['Enrollment_date_only_From_M365'] = if ($out['Enrollment_date_From_M365']) { ([datetime]::Parse($out['Enrollment_date_From_M365'])).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { '' }
        $out['Model_From_M365'] = $model
        $out['Model'] = $model
        $out['Manufacturer'] = $manufacturer
        $out['Win11_Manufacturer'] = $manufacturer
        $out['Ownership_From_M365'] = V $intune @('Ownership')
        $out['Compliance_From_M365'] = V $intune @('Compliance','IsCompliant')
        $out['Primary_user_email_address_From_M365'] = V $intune @('Primary user email address','Primary user UPN')
        $out['Primary_username_address_From_M365'] = $primaryUser
        $out['Last_Logged_User_FromM365'] = $lastLoggedUserFromM365
        $out['PrimarySMTPaddressUser'] = $primaryUser
        $out['PrimarySMTPaddress_From_AD_Or_M365'] = if ($primaryUserRow) { V $primaryUserRow @('EmailAddress','UserPrincipalName') } else { $primaryUser }
        $out['LicenseFromM365_Users_From_AD'] = $licenseValue
        $out['LicenseFromM365_Users_From_PrimaryIntuneUser'] = $licenseValue
        $out['OU_Path_Users_From_AD'] = V $primaryUserRow @('CanonicalName','DistinguishedName')
        $out['OrganizationalUnit_User_From_AD'] = $out['OU_Path_Users_From_AD']
        $out['OrganizationalUnit_PrimaryUser_From_AD'] = $out['OU_Path_Users_From_AD']
        $out['Enabled_From_Last_Logged_UserDomain'] = B (V $lastUserRow @('Enabled'))
        $out['Last_Logged_User'] = $lastLoggedUser
        $out['Last_Logged_User_FromSentinel'] = $lastLoggedUserFromSentinel
        $out['Last_Logged_User_Source'] = if ($lastLoggedUserFromM365) { 'M365' } else { 'Unknown' }
        $out['Last_Logged_UserDomain'] = $lastLoggedUserDomain
        $out['LastOUName_Users_From_AD'] = if ($lastUserLastOu) { $lastUserLastOu } else { '?' }
        $out['DistinguishedName_Last_Logged_UserDomain'] = V $lastUserRow @('DistinguishedName')
        $out['OrganizationalUnit_LastUserS1_From_AD'] = V $lastUserRow @('CanonicalName','DistinguishedName')
        $out['PrimarySMTPaddress_From_S1_Or_M365'] = if ($out['Primary_user_email_address_From_M365']) { $out['Primary_user_email_address_From_M365'] } else { $lastUserPrimarySmtp }
        $out['LicenseFromM365_Users_From_LastUserS1'] = $lastUserLicenseGroups
        $out['IsMailBox_From_Last_Logged_UserDomain'] = [string]($null -ne $lastUserMailbox)
        $out['IsMailBox_From_PrimaryIntuneUser'] = [string]($null -ne $lastUserMailbox)
        $out['IsMailBoxSize_From_Last_Logged_UserDomain'] = $lastUserMailboxSize
        $out['IsMailBoxSize_From_PrimaryIntuneUser'] = $lastUserMailboxSize
        $out['MailboxSize_From_Last_Logged_UserDomain'] = $lastUserMailboxSize

        $out['DiskTotalStorage_Go_From_M365'] = $totalGb
        $out['DiskFreeStorage_Go_From_M365'] = $freeGb
        $out['DiskTotalSize'] = $totalGb
        $out['DiskTotalFreeSpace'] = $freeGb
        $out['DiskFreePercent'] = $freePercent
        $out['IsFreeStorageNotEnoughForWin11Update'] = $freeNotEnough
        $out['PhysicalMemoryGB_From_M365'] = $physicalMemoryGb
        $out['Memory_GB_Number'] = $physicalMemoryGb
        $out['Memory'] = $physicalMemoryGb

        $out['AzureEntra_ObjectId'] = V $entra @('ObjectId')
        $out['entraDeviceId_norm'] = GK (V $entra @('DeviceId'))
        $out['entraDeviceId_match_count'] = [string]$entraByGuidCount
        $out['AzureEntra_RegistrationDateTime'] = $registration
        $out['AzureEntra_ApproximateLastSignInDateTime'] = $entraSignIn
        $out['AzureEntra_DaysSinceLastSignIn'] = Days (V $entra @('ApproximateLastSignInDateTime'))
        $out['AzureEntra_HardwareId_DeviceCount'] = $hardwareCount
        $out['AzureEntra_CorrelationStatus_Robust'] = $correlation
        $out['EntraExistsByGUID'] = [string]$entraExists
        $out['EntraRegisteredPending'] = [string]($entraExists -and ((B (V $entra @('IsPending'))) -eq 'True'))
        $out['EntraNameMatchesCountByName'] = [string]$entraByNameCount
        $out['EntraMatchesCountByGUID'] = [string]$entraByGuidCount
        $out['EntraNameMatchesByGUIDList'] = $entraNameList
        $out['AzureEntra_LastSignIn_Status'] = if ($entraSignIn) { 'Known' } else { 'Unknown / Never logged in' }
        $out['AzureEntra_LastSignIn_OlderThan3M'] = if ($out['AzureEntra_DaysSinceLastSignIn'] -eq '') { '' } else { [string]([int]$out['AzureEntra_DaysSinceLastSignIn'] -gt 90) }
        $out['AzureEntra_LastSignIn_Recent_3M'] = if ($out['AzureEntra_DaysSinceLastSignIn'] -eq '') { '' } else { [string]([int]$out['AzureEntra_DaysSinceLastSignIn'] -le 90) }

        $out['LastSignIn_Merged'] = $lastSignInMergedValue
        $mergedDays = $mergedDaysValue
        $out['LastSignIn_Merged_Active_Last90Days'] = if ($mergedDays -eq '') { '' } else { [string]([int]$mergedDays -le 90) }
        $out['LastSignIn_Merged_Active_Last45Days'] = if ($mergedDays -eq '') { '' } else { [string]([int]$mergedDays -le 45) }
        $out['LastSignIn_Merged_Active_Last30Days'] = if ($mergedDays -eq '') { '' } else { [string]([int]$mergedDays -le 30) }

        SetWu $out $wuAnchor 'Autopatch_FeatureUpdate_Anchor'
        SetWu $out $wu24 'Windows_11_24H2'
        SetWu $out $wu25 'Windows_11_25H2'

        $out['SecureBootStatus'] = V $local @('SecureBootStatus')
        $out['BIOS_Version'] = V $local @('BIOSVersion')
        $out['BIOS_Key'] = $out['BIOS_Version']
        $out['BIOS_Model'] = $model
        $out['Last Reboot Date'] = V $local @('LastRebootDate','LastBootUpTime')
        $out['Last Active Date'] = $lastActive
        $out['LastOUName'] = $lastOu
        $out['IsInDisabledObjectsOU'] = $isDisabledOu
        $out['IsVirtualMachine'] = $isVirtual
        $out['TypeEntity'] = $typeEntity
        $out['Mig-MigrationPlanW11-Phase'] = $migrationPhase
        $out['Mig-MigrationPlanW11-Phase-Reason'] = $migrationReason
        $out['Mig-MigrationPlanW11-Action'] = $migrationAction
        $out['Mig-MigrationPlanW11-StartDate'] = ''
        $out['Mig_Potential_Issues_Score'] = $issueScore
        $out['Mig_Potential_Issues_Devices_Level2'] = $issueLevel
        $out['WIN11_GLOBALRESULT_OLD'] = $win11GlobalResultOld
        $out['OperatingSystem_20250918_170952'] = $os20250918
        $out['OperatingSystem_20251217_111133'] = $os20251217
        $out['operatingSystemVersion_20250918_170952'] = $osVersion20250918
        $out['operatingSystemVersion_20251217_111133'] = $osVersion20251217
        $out['OperatingSystem_Major_20251217'] = GetOsMajor $os20251217
        $out['OperatingSystem_Major_Snapshot_Priority'] = $snapshotMajorPriority
        $out['OperatingSystem_Changed'] = $osChanged
        $out['OperatingSystem_Changed_Text'] = $osChangedText
        $out['OperatingSystem_10to11_Changed'] = $os10To11Changed
        $out['OperatingSystem_10to11_Changed_Text'] = $os10To11ChangedText
        $out['OperatingSystem_SnapshotUsed'] = $snapshotUsed
        $out['ModelReleaseDate'] = ''
        $out['IsWindows11AndIntune'] = [string]($isWin11 -and ($null -ne $intune))
        $out['Mig-FlagMatchExistsW1124H2'] = [string]$is24
        $out['Mig-FlagMatchExistsW1124H2-Bool'] = [string]$is24
        $out['LastLogonDate_Category'] = GetDateCategory (V $row @('LastLogonDate')) 90
        $out['LastLogonDate_Category30Days'] = GetDateCategory (V $row @('LastLogonDate')) 30
        $out['AzureEntraLastSignInDate_Category'] = GetDateCategory $entraSignIn 90
        $out['AzureEntraLastSignInDate_Category2'] = $out['AzureEntraLastSignInDate_Category']
        $out['AzureEntraLastSignInDate_Category30Days'] = GetDateCategory $entraSignIn 30
        $out['LastActiveDate_Category'] = GetDateCategory $lastActive 90
        $out['LastActiveDate_Category2'] = $out['LastActiveDate_Category']
        $out['LastActiveDate_Category30Days'] = GetDateCategory $lastActive 30
        $out['LastRebootDate_Category'] = GetDateCategory $out['Last Reboot Date'] 90
        $out['LastRebootDate_Category2'] = $out['LastRebootDate_Category']
        $out['LastRebootDate_Category30Days'] = GetDateCategory $out['Last Reboot Date'] 30

        [void]$enrichedRows.Add([pscustomobject]$out)
    }

    $enrichedCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Computers_AllDomains_Enriched.csv'
    $enrichedRows | Add-SmartM365TenantKey | Export-Csv -LiteralPath $enrichedCsv -NoTypeInformation -Encoding UTF8
    if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
        $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) { if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) { [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath) } }
    }
    [void]$global:csvGeneratedPaths.Add($enrichedCsv)
    WriteLog -Message ("AD computers enriched CSV generated: {0} ({1} row(s), {2} enriched column(s))" -f $enrichedCsv, $enrichedRows.Count, $calculatedColumns.Count)
    return $enrichedCsv
}





# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCrr+xpX4kiOEW0
# 4aP3tSqKasC+7PBkbP1zqzHpVzqvZqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC8sxGN2YmlcamWRDZulKqf
# U9NXL1lCkza7q4AZPOerRTANBgkqhkiG9w0BAQEFAASCAYBjDav7BK/GgfLk8BwV
# 7iOtcBAoOqwV6kFwAjB3cKwu/a/sFORWkXIE0JQH06K90PUAeHUgFg9jp2Ji97Xv
# 2a+DwI1oIjDOjyqSiN/jtf1zilOikhc7IWwVeFp5anKE6qrqWdMa65+8RLMhlhxv
# f6GxEr6n5ehrKAoGla8DXaD0tlbm+MGsBZ5B31nKctA8hmcYwpr0KYZIZ7SzZtBh
# yxFfDAx9MkjGnqWbbxstgBjVceTog4n3uIeG3YG+9mFDnFsjg2gQIwzMmJlyBxB8
# D/TXAtwu0AUsu1YvKgM51mMg7NaL1Seun8AAQqwd3LXLqTGbVautb7Rk8YFDA7hl
# Lo19McOlbGlfNjPuwJpLKRxC0iL+/XFZUO8WhwdmvA87bRU7wqVFIDy0diXDV0Qo
# F7uYTFr6xFd7tD/GQHfrL8Cl0kgDo0mS+IYVAUwd80EkdOyJSHQd29pyclvkU5Hq
# VIpm4oik1F7Zwd2l4/prCQKwaMIlPoVhGA9+VMowkLJFZUw=
# SIG # End signature block
