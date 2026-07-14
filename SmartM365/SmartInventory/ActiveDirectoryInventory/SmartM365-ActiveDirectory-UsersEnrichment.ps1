<#
.SYNOPSIS
    Builds enriched Active Directory user CSV columns required by the SmartWorkplace Power BI model.
#>

function Invoke-SmartM365AdUsersEnrichedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CombinedUsersCsv,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $false)][string]$LatestFolderPath
    )

    if (-not (Test-Path -LiteralPath $CombinedUsersCsv)) {
        WriteLog -Message ("WARNING: AD users enrichment source CSV not found: {0}" -f $CombinedUsersCsv)
        return $null
    }

    function Get-OuPathFromDistinguishedName {
        param([string]$DistinguishedName)
        if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return '' }
        $endCnPosition = $DistinguishedName.IndexOf(',')
        if ($endCnPosition -lt 0) { return '' }
        $textAfterCn = $DistinguishedName.Substring($endCnPosition + 1)
        $domainStartPosition = $textAfterCn.IndexOf(',DC=', [System.StringComparison]::OrdinalIgnoreCase)
        if ($domainStartPosition -gt 0) { return $textAfterCn.Substring(0, $domainStartPosition).Trim() }
        return $textAfterCn.Trim()
    }

    function Get-TypeEntityFallback {
        param([string]$DomainNameShort, [string]$Country)
        $domainClean = ([string]$DomainNameShort).Trim().ToUpperInvariant()
        $countryClean = ([string]$Country).Trim().ToUpperInvariant()
        if ($domainClean -eq 'ORPEA_01') { return 'CORP' }
        if ($domainClean -eq 'GRP' -and @('BE','LU') -contains $countryClean) { return 'INTEGRATED' }
        if ($domainClean -eq 'GRP') { return 'NOT-INTEGRATED' }
        if (@('BE','CH','DE','ES','FR','IT','LU','PL','PT') -contains $domainClean) { return 'INTEGRATED' }
        if (@('AT','CZ','NL') -contains $domainClean) { return 'NOT-INTEGRATED' }
        return 'UNKNOWN'
    }

    function Convert-SmartM365NameClean {
        param([string]$Value)
        $clean = ([string]$Value).ToUpperInvariant()
        foreach ($pair in @(@('À','A'),@('Â','A'),@('Ç','C'),@('É','E'),@('È','E'),@('Ê','E'),@('Ë','E'),@('Î','I'),@('Ï','I'),@('Ô','O'),@('Ù','U'),@('Û','U'),@('Ü','U'))) { $clean = $clean.Replace($pair[0], $pair[1]) }
        return $clean.Replace(' ', '')
    }

    function Get-SmartM365BasePersona {
        param([string]$JobFamily)
        switch (([string]$JobFamily).Trim().ToUpperInvariant()) {
            'HEADQUARTERS STAFF' { return 'M365E3' }
            'MANAGEMENT AND ADMINISTRATIVE STAFF' { return 'M365E3' }
            'HEALTHCARE STAFF' { return 'M365F3' }
            'SERVICE STAFF' { return 'M365F3' }
            'EXT STAFF WITH DEVICE' { return 'M365E3' }
            'EXT STAFF WITHOUT DEVICE' { return 'M365F3' }
            'EXT STAFF' { return 'M365E3' }
            'UNCLASSIFIED' { return 'M365F3' }
            'NON-IT STAFF' { return 'None' }
            'NOT INTEGRATED BUS' { return 'None' }
            'NOT INTEGRATED BUSS' { return 'None' }
            default { return 'M365F3' }
        }
    }

    function Get-SmartM365AccountType {
        param([string]$UPN,[string]$SAM,[string]$DN,[string]$RecipientType,[string]$GivenName,[string]$Surname,[string]$GivenNameClean,[string]$SurnameClean)
        $samUc = ([string]$SAM).Trim().ToUpperInvariant(); $samLc = ([string]$SAM).Trim().ToLowerInvariant(); $upnLc = ([string]$UPN).Trim().ToLowerInvariant(); $dnUc = ([string]$DN).ToUpperInvariant(); $rtUc = ([string]$RecipientType).Trim().ToUpperInvariant()
        if ($rtUc.Contains('SHARED')) { return 'Shared Mailbox' }
        if ($rtUc.Contains('ROOM')) { return 'Room Mailbox' }
        if ($upnLc.Contains('.ext@') -or $upnLc.Contains('-ext@')) { return 'Ext Account' }
        if ($SAM.StartsWith('DefaultAccount') -or $SAM.StartsWith('$') -or $SAM.EndsWith('$') -or $samUc.Contains('HEALTHMAILBOX') -or $SAM.StartsWith('MSOL') -or $SAM.StartsWith('KRBTGT') -or $SAM.StartsWith('ASPNET') -or $SAM.StartsWith('GUEST') -or $SAM.StartsWith('SUPPORT_') -or $SAM.StartsWith('SQL') -or $SAM.StartsWith('__') -or $SAM.StartsWith('SSHD') -or $samUc.Contains('IUSR_')) { return 'System Account' }
        if ($SAM.StartsWith('SVC_') -or $SAM.StartsWith('SVC-') -or $samLc.StartsWith('svc_') -or $samLc.StartsWith('svc-') -or $SAM.EndsWith('-SVC') -or $SAM.EndsWith('_SVC') -or $samLc.EndsWith('-svc') -or $samLc.EndsWith('_svc') -or $SAM.StartsWith('S_') -or $samLc.StartsWith('s_') -or $dnUc.Contains('SERVICE_ACCOUNTS') -or $dnUc.Contains('SERVICE ACCOUNTS')) { return 'Service Account' }
        if ($SAM.StartsWith('A_') -or $samLc.StartsWith('a_') -or $samUc.StartsWith('ADMINISTRATOR') -or $samUc.StartsWith('ADMINISTRATEUR') -or $SAM.StartsWith('ADMINDOM') -or $SAM.StartsWith('D_') -or $samLc.StartsWith('d_') -or $SAM.StartsWith('E_') -or $samLc.StartsWith('e_') -or $dnUc.Contains('OU=ADMIN')) { return 'Admin Account' }
        if (@('SharedMailbox','RoomMailbox') -contains $RecipientType -or $SAM.Contains('TABLET') -or $dnUc.Contains('SHARED MAILBOX') -or $SAM.Contains('WB0') -or $SAM.Contains('ATENDI') -or $samUc.StartsWith('CITRIX_TEST') -or $samUc.Contains('SCANER') -or $samUc.Contains('SCHULUNG') -or ([string]$GivenName).Trim().ToUpperInvariant() -eq 'SCAN' -or $SAM -match '^[0-9]' -or [string]::IsNullOrWhiteSpace($GivenName) -or [string]::IsNullOrWhiteSpace($Surname)) { return 'Generic Account' }
        $gnc = ([string]$GivenNameClean).Trim().ToUpperInvariant(); $snc = ([string]$SurnameClean).Trim().ToUpperInvariant()
        if ($gnc.Length -ge 3 -and $snc.Length -ge 3 -and $samUc.StartsWith($gnc.Substring(0,3) + $snc.Substring(0,3))) { return 'Named Account' }
        if ($gnc.Length -eq 2 -and $snc.Length -ge 3 -and $samUc.StartsWith($gnc.Substring(0,2) + $snc.Substring(0,3))) { return 'Named Account' }
        $gn = ([string]$GivenName).Trim(); $sn = ([string]$Surname).Trim(); if ($gn.Length -gt 0 -and $sn.Length -gt 0 -and $upnLc.StartsWith(($gn + '.' + $sn).ToLowerInvariant())) { return 'Named Account' }
        if ($samUc.Length -ge 10) { $part1 = $samUc.Substring(0,6); $part2 = $samUc.Substring(6,4); if ($part1 -notmatch '[0-9]' -and $part2 -match '^[0-9]+$') { return 'Named Account' } }
        return 'Unclassified Account'
    }

    function Get-SmartM365LatestDateText {
        param([string[]]$Values)
        $latest = $null
        foreach ($value in $Values) {
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $date = [datetime]::MinValue
            if ([datetime]::TryParse($value, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date) -or [datetime]::TryParse($value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) {
                if ($null -eq $latest -or $date.ToUniversalTime() -gt $latest) { $latest = $date.ToUniversalTime() }
            }
        }
        if ($null -eq $latest) { return '' }
        return $latest.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    function Get-PrimarySmtpFromProxyAddressText {
        param([string]$ProxyAddresses)
        if ([string]::IsNullOrWhiteSpace($ProxyAddresses)) { return '' }
        $normalized = $ProxyAddresses.Replace("`r", ';').Replace("`n", ';')
        while ($normalized.Contains(';;')) { $normalized = $normalized.Replace(';;', ';') }
        $match = [regex]::Match($normalized, '(^|;)SMTP:(?<Mail>[^;]+)')
        if (-not $match.Success) { return '' }
        return $match.Groups['Mail'].Value.Trim().ToLowerInvariant()
    }

    function Get-MailDomain {
        param([string]$Mail)
        if ([string]::IsNullOrWhiteSpace($Mail)) { return '' }
        $pos = $Mail.IndexOf('@')
        if ($pos -lt 0) { return '' }
        return $Mail.Substring($pos + 1).ToLowerInvariant()
    }

    function Get-DateCategory {
        param([string]$Value, [int]$LimitDays)
        if ([string]::IsNullOrWhiteSpace($Value)) { return 'Unknown / Never logged in' }
        $date = [datetime]::MinValue
        if (-not [datetime]::TryParse($Value, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date) -and
            -not [datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) {
            return 'Unknown / Never logged in'
        }
        if ($date.ToUniversalTime().Date -lt ([datetime]::UtcNow.Date.AddDays(-1 * $LimitDays))) { return ("Older than {0} days" -f $LimitDays) }
        return ("Within last {0} days" -f $LimitDays)
    }
    function Get-NumberInvariant {
        param([object]$Value)
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $text = $text.Replace([string][char]0x00A0, '').Replace([string]' ', '')
        $number = 0.0
        if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
        if ([double]::TryParse($text.Replace(',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
        return $null
    }

    function Format-NumberInvariant {
        param([object]$Value)
        if ($null -eq $Value) { return '' }
        return ([double]$Value).ToString('0.########', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Normalize-SmartM365PersonaeText {
        param([string]$Value)
        $text = ([string]$Value).Trim().ToLowerInvariant()
        foreach ($pair in @(@('é','e'),@('è','e'),@('ê','e'),@('ë','e'),@('à','a'),@('â','a'),@('î','i'),@('ï','i'),@('ô','o'),@('û','u'),@('ù','u'),@('ä','a'),@('ö','o'),@('ü','u'),@('ú','u'),@('ó','o'),@('á','a'),@('í','i'),@('ñ','n'),@('ç','c'),@('ß','ss'),@('œ','oe'),@('æ','ae'))) { $text = $text.Replace($pair[0], $pair[1]) }
        return $text
    }

    function Normalize-SmartM365JobTitleText {
        param([string]$ManagedTitle, [string]$DescriptionTitle)
        $managed = ([string]$ManagedTitle).Trim()
        $description = ([string]$DescriptionTitle).Trim()
        if ($description.ToLowerInvariant().Contains('deprovision')) { $description = '' }
        if (-not [string]::IsNullOrWhiteSpace($description) -and $managed.ToLowerInvariant() -eq $description.ToLowerInvariant()) { $text = $managed.ToLowerInvariant() }
        else { $text = ($managed + $(if ([string]::IsNullOrWhiteSpace($description)) { '' } else { ' ' + $description })).Trim().ToLowerInvariant() }
        foreach ($separator in @('(',')','/','-',',',';',':','.',"'",'&','+','|','_',[string][char]160)) { $text = $text.Replace($separator, ' ') }
        $text = $text.Replace([string][char]0x2013, ' ').Replace([string][char]0x2014, ' ')
        $text = Normalize-SmartM365PersonaeText -Value $text
        while ($text.Contains('  ')) { $text = $text.Replace('  ', ' ') }
        return $text.Trim()
    }

    function Get-SmartM365JobTitleFromDescription {
        param([string]$Description)
        $text = ([string]$Description)
        $position = $text.IndexOf('-')
        if ($position -lt 0) { return '' }
        return $text.Substring($position + 1).Trim()
    }

    function New-SmartM365PersonaeKeywordRows {
        param([object[]]$Rows, [bool]$StrictCategory)
        $result = New-Object System.Collections.Generic.List[object]
        foreach ($row in $Rows) {
            $keyword = Normalize-SmartM365PersonaeText -Value (Get-Value $row @('Keyword'))
            $category = [string](Get-Value $row @('Category'))
            if ([string]::IsNullOrWhiteSpace($keyword) -or [string]::IsNullOrWhiteSpace($category)) { continue }
            [void]$result.Add([pscustomobject]@{ Keyword = $keyword; KeywordPadded = (' {0} ' -f $keyword); KeywordLength = $keyword.Length; Category = $category; StrictCategory = $StrictCategory })
        }
        return @($result | Sort-Object -Property @{ Expression = 'KeywordLength'; Descending = $true }, Keyword)
    }

    function Get-SmartM365JobFamilyFromKeywords {
        param([string]$PaddedTitle, [object[]]$KeywordRows, [bool]$StrictCategory)
        if ([string]::IsNullOrWhiteSpace($PaddedTitle) -or @($KeywordRows).Count -eq 0) { return 'Unclassified' }
        foreach ($keywordRow in $KeywordRows) {
            if ($PaddedTitle.Contains([string]$keywordRow.KeywordPadded)) {
                $category = [string]$keywordRow.Category
                if (-not $StrictCategory) { return $category }
                switch ($category) {
                    'Healthcare staff' { return 'Healthcare staff' }
                    'Management and Administrative staff' { return 'Management and Administrative staff' }
                    'Non-IT staff' { return 'Non-IT staff' }
                    'Service staff' { return 'Service staff' }
                    default { return 'Unclassified' }
                }
            }
        }
        return 'Unclassified'
    }

    function Get-SmartM365OrganizationalUnitCode {
        param($Row)
        $raw = (Get-Value $Row @('OrganizationalUnit','extensionAttribute13')).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { return '999999' }
        $numVal = 0.0
        if ([double]::TryParse($raw.Replace(',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numVal)) { return ([math]::Floor($numVal)).ToString('000000', [System.Globalization.CultureInfo]::InvariantCulture) }
        return '999998'
    }

    function Get-SmartM365TypeEtablissement {
        param($EntityRow)
        if ($null -eq $EntityRow) { return 'NOTFOUND' }
        $hqText = ([string](Get-Value $EntityRow @('HQ'))).Trim().ToLowerInvariant()
        if (@('1','true','vrai','oui') -contains $hqText) { return 'HQ' }
        $service = ([string](Get-Value $EntityRow @('Entity (Service)'))).ToUpperInvariant()
        if ($service.Contains('RESIDENCE')) { return 'RESIDENCE' }
        if ($service.Contains('CLINIQUE')) { return 'CLINIC' }
        if ($service.Contains('UNKNOWN')) { return 'UNKNOWN' }
        return 'OTHER'
    }

    function Get-SmartM365CombinedJobFamily {
        param([string]$AccountType, [string]$TypeEtablissement, [string]$TypeEntity, [string]$UserPrincipalName, [bool]$NoLastLogonInOnComputer, [string]$KeywordCategory)
        $accountTypeText = ([string]$AccountType).Trim()
        $typeEtablissementText = ([string]$TypeEtablissement).Trim()
        $typeEntityText = ([string]$TypeEntity).Trim()
        $upnLower = ([string]$UserPrincipalName).ToLowerInvariant()
        $notIntegrated = if ($typeEntityText -eq 'NOT-INTEGRATED') { 'Not Integrated BUs' } else { '' }
        $immediate = ''
        if ([string]::IsNullOrWhiteSpace($notIntegrated)) {
            switch ($accountTypeText) {
                'Ext Account' { $immediate = 'Ext staff' }
                'Service Account' { $immediate = 'Non-IT staff' }
                'Admin Account' { $immediate = 'Non-IT staff' }
                'System Account' { $immediate = 'Non-IT staff' }
                'MDM Account' { $immediate = 'Non-IT staff' }
                'Shared Mailbox' { $immediate = 'Non-IT staff' }
                'Room Mailbox' { $immediate = 'Non-IT staff' }
                'Test Account' { $immediate = 'Non-IT staff' }
            }
        }
        $extDetected = ''
        if ([string]::IsNullOrWhiteSpace($notIntegrated) -and ($upnLower.Contains('.ext') -or $upnLower.Contains('#ext#') -or $upnLower.Contains('-ext'))) { $extDetected = 'Ext staff' }
        $extFinal = if ($extDetected) { if ($NoLastLogonInOnComputer) { 'Ext staff without device' } else { 'Ext staff with device' } } else { '' }
        $hqStaff = if ([string]::IsNullOrWhiteSpace($notIntegrated) -and [string]::IsNullOrWhiteSpace($extDetected) -and $typeEtablissementText -eq 'HQ') { 'Headquarters staff' } else { '' }
        $keyword = if ([string]::IsNullOrWhiteSpace($notIntegrated) -and [string]::IsNullOrWhiteSpace($extDetected) -and [string]::IsNullOrWhiteSpace($hqStaff)) { $KeywordCategory } else { '' }
        if ($notIntegrated) { return $notIntegrated }
        if ($extFinal) { return $extFinal }
        if ($immediate) { return $immediate }
        if ($hqStaff) { return $hqStaff }
        if ($keyword) { return $keyword }
        return 'Unclassified'
    }

    function Get-Value {
        param($Row, [string[]]$Names)
        if ($null -eq $Row) { return '' }
        foreach ($name in $Names) {
            $property = $Row.PSObject.Properties[$name]
            if ($null -ne $property -and $null -ne $property.Value) { return [string]$property.Value }
        }
        return ''
    }

    function Get-Key {
        param([object]$Value)
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        return $text.ToUpperInvariant()
    }

    function Get-SourceCandidatePaths {
        param([string[]]$Names)
        $roots = New-Object System.Collections.Generic.List[string]
        foreach ($root in @($LatestFolderPath, $OutputFolder, (Split-Path -Path $CombinedUsersCsv -Parent))) { if (-not [string]::IsNullOrWhiteSpace([string]$root)) { [void]$roots.Add([string]$root) } }
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) { $latestParent = Split-Path -Path $LatestFolderPath -Parent; if (-not [string]::IsNullOrWhiteSpace($latestParent)) { [void]$roots.Add((Join-Path -Path $latestParent -ChildPath 'DATA-ALL')) } }
        if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
            $currentRoot = $OutputFolder
            for ($i = 0; $i -lt 6 -and -not [string]::IsNullOrWhiteSpace($currentRoot); $i++) {
                if ([System.IO.Path]::GetFileName($currentRoot).Equals('DATA-ALL', [System.StringComparison]::OrdinalIgnoreCase)) { [void]$roots.Add($currentRoot); break }
                $currentRoot = Split-Path -Path $currentRoot -Parent
            }
        }
        foreach ($rootForTenantLookup in @($LatestFolderPath, $OutputFolder, (Split-Path -Path $CombinedUsersCsv -Parent))) {
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
        $candidates = New-Object System.Collections.Generic.List[string]
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($root in $roots) { if ([string]::IsNullOrWhiteSpace([string]$root)) { continue }; foreach ($name in $Names) { $candidate = Join-Path -Path $root -ChildPath $name; if ($seen.Add($candidate)) { [void]$candidates.Add($candidate) } } }
        return @($candidates)
    }

    function Find-SourceCsv {
        param([string[]]$Names)
        foreach ($candidate in (Get-SourceCandidatePaths -Names $Names)) { if (Test-Path -LiteralPath $candidate) { return $candidate } }
        return ''
    }

    function Import-SourceCsv {
        param([string[]]$Names)
        $csvPath = Find-SourceCsv -Names $Names
        if ([string]::IsNullOrWhiteSpace($csvPath)) { return @() }
        return @(Import-Csv -LiteralPath $csvPath -Encoding UTF8)
    }

    function Import-SourceXlsx {
        param([string[]]$Names, [string[]]$WorksheetNames)
        foreach ($candidate in (Get-SourceCandidatePaths -Names $Names)) {
            if (-not (Test-Path -LiteralPath $candidate)) { continue }
            try { Import-Module ImportExcel -ErrorAction Stop }
            catch { WriteLog -Message ("AD users enrichment Excel source skipped because ImportExcel module is not available: {0}" -f $candidate); return @() }
            $rowsOut = New-Object System.Collections.Generic.List[object]
            foreach ($worksheetName in @($WorksheetNames)) {
                if ([string]::IsNullOrWhiteSpace($worksheetName)) { continue }
                try { foreach ($row in @(Import-Excel -Path $candidate -WorksheetName $worksheetName -ErrorAction Stop)) { [void]$rowsOut.Add($row) } }
                catch { WriteLog -Message ("AD users enrichment Excel worksheet skipped: {0} [{1}] ({2})" -f $candidate, $worksheetName, $PSItem.Exception.Message) }
            }
            WriteLog -Message ("AD users enrichment Excel source loaded: {0} ({1} row(s))" -f $candidate, $rowsOut.Count)
            return $rowsOut.ToArray()
        }
        WriteLog -Message ("AD users enrichment optional Excel source missing; related columns will be blank: {0}" -f ($Names -join ', '))
        return @()
    }
    function Add-SmartM365MapValue {
        param([hashtable]$Map, [string]$Key, $Row)
        if ([string]::IsNullOrWhiteSpace($Key)) { return }
        if (-not $Map.ContainsKey($Key)) { $Map[$Key] = $Row }
    }

    function Add-SmartM365MapListValue {
        param([hashtable]$Map, [string]$Key, $Row)
        if ([string]::IsNullOrWhiteSpace($Key)) { return }
        if (-not $Map.ContainsKey($Key)) { $Map[$Key] = New-Object System.Collections.Generic.List[object] }
        [void]$Map[$Key].Add($Row)
    }

    function Join-UniqueText {
        param([object[]]$Values, [string]$Separator)
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($value in $Values) {
            $text = ([string]$value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) { [void]$set.Add($text) }
        }
        return @($set | Sort-Object) -join $Separator
    }
    $calculatedColumns = @(
        'OU_Path','OU_Category','TypeEtablissement','TypeEntity','IsServiceAccountOrSimilar','IsSpecificAdminOrService','IsTargetOU_HQ','IsNotTargetOU_HQ','NoLastLogonAD','EnabledAndNoLastLogonMerge',
        'GivenNameClean','SurnameClean','AccountType','EXO_Mapping_Status','IsShared_Status','IsLargeSharedMailbox','IsInLocalEXODuplicates','Is_In_SHAREDMAILBOX','TotalItemSizeNumeric_From_Mailboxes','TotalItemSizeNumeri_From_RemoteMailboxes','M365LicenseType','M365LicenseType_ByPersonae','M365LicenseType_ByPersonae2','M365LicenseType_Target1_ByPersona','M365LicenseType_Target1_ByPersona 2','M365LicenseType_Target2_ByPersona','M365LicenseType_Target3_ByPersona','HeavyMailboxF3User','M365License_F3_E3_SizeBucket','M365License_F3_E3_SizeBucket_Sort','M365License_F3_NonCompliant_Size_Remaining_Flag',
        'AccountToCleanExcluded','AccountToCleanStatus1','AccountToCleanStatus2','ActivityCategory','FullAccess_From_Mailboxes_UPN','FullAccess_From_MailboxesEXO','FullAccess_From_MailboxesEXO_Raw','GrantSendOnBehalfTo_UPN','IsJobTitleValid','Job Family Combined','Job Family Combined 2','Job Family from Keywords','Job Family from Keywords 2','Job Title Combined','Job Title From Desc','Job Title From Desc 2','Job Title Normalized','Job Title Normalized 2','Job Title Padded','Job Title Padded 2','LabelFromEntity','LastLoggedComputerSamAccountNames','LastLoggedInOnALLComputers','LastLoggedInOnComputer','LastLogonEntra','LastLogonEntra_Category','LastLogonEntraNonInteractive','LastLogonMergeADExchEntra','LocalMailboxPermissionCount_NotInM365','LogonCountStatus','M365License_E3Persona_MissingE3_Flag','M365License_F3_AllOversized_SizeBucket','M365License_F3_AllOversized_SizeBucket_Sort','M365License_NoLicenseF3_EnabledMailbox_BlockingMigration_Flag','Mig_Potential_Issues_Mailbox_Level2','Mig_Potential_Issues_Score','NoLastLogonEntra','NoLastLogonEntraNonInteractive','NoLastLogonExch','NoLastLogonInOnComputer','SendAs_From_Mailboxes_UPN',
        'PrimarySmtp','PrimarySmtp_Lower','PrimarySMTP_Domain','Containsorpeamailonmicrosoft','ExistsInM365ActiveUsers',
        'RecipientType_From_Mailboxes','RecipientTypeDetails_From_Mailboxes','ExistsInRemoteMailboxes','IsMailBox','MailboxSource','MailboxLocation','MailboxSourceHealthStatus',
        'ItemCount_From_Mailboxes','IsItemCountOver999999','TotalItemSizeNumeric','TotalItemSize_GB','TotalItemSize_TB','TotalArchiveItemSizeNumeric','TotalArchiveItemSize_GB','TotalItemSizeAndArchiveNumeric','TotalItemSizeAndArchive_GB','IsLargerThan2GB','IsLargerThan100GB','IsSmallerThan2GB','MailboxSize_From_Mailboxes',
        'LastLogonDateFromExch','LastLogonDateFromExchDate','LastLogonMerge','IsLastLogonMerge','NoLastLogonMerge','LastLogonDateFromExchDate_Category','LastLogonDate_Category','LastLogonFromExch_OlderThan_3M','LastLogonFromExch_OlderThan_6M',
        'LicenseFromM365','LicenseGroupFromM365','LicenseExchangePlanEnabledFromM365','LicenseGroupCount','LicenseGroupFound','LicenseFromM365_GroupsAssigningSku','LicenseFromM365_HasDirect','LicenseFromM365_HasGroupOnly','LicenseFromM365_HasDirectAndGroup','LicenseGroupHasDoubleAssignment',
        'DomainAllowed_EXO','DomainNotAllowed_EXO','DomainType_EXO','DomainIsOnMicrosoft_EXO','DomainNonCompliant_EXO','DomainUsersNotAllowed_EXO','DomainUsersNonCompliant_EXO','UPNDomainVerified_AAD','UPNEqualsPrimarySMTPDomain',
        'RemoteRoutingAddress_From_MailboxesRemote','WhenRemoteMailboxCreated','DaysSinceRemoteMailboxCreated','WhenMailboxCreated_From_Mailboxes','WhenMailboxCreated_From_EXO','WhenMailboxCreated_From_EXO_Date',
        'FullAccessOn_From_Mailboxes','SendAsOn_From_Mailboxes','SendOnBehalfOn_From_Mailboxes','FullAccess_From_Mailboxes_Raw','FullAccess_From_Mailboxes','SendAs_From_Mailboxes_Raw','SendAs_From_Mailboxes','GrantSendOnBehalfTo_From_Mailboxes_Raw','GrantSendOnBehalfTo_From_Mailboxes','PermissionsSummary','PermissionsCount','SharedAccessCount','HasFullAccess','HasSendAsBehalf','HasSendOnBehalf','HasAnyDelegPerm','HasCrossPremisesPermissions',
        'MigrationUserStatus_From_MigrationJobs','MigrationUserBatchName_From_MigrationJobs','MigrationUserCompleteAfter_From_MigrationJobs','MigrationUserCompleteAfterUTC_From_MigrationJobs','ProtectedMailboxes_From_EXO_BackupProtection'
    )

    WriteLog -Message ("AD users enrichment loading source CSV: {0}" -f $CombinedUsersCsv)
    $users = @(Import-Csv -LiteralPath $CombinedUsersCsv -Encoding UTF8)
    $localMailboxes = Import-SourceCsv @('Exchange_OnPrem_Mailboxes_AllDomains.csv')
    $remoteMailboxes = Import-SourceCsv @('Exchange_OnPrem_RemoteMailboxes_AllDomains.csv')
    $exoMailboxes = Import-SourceCsv @('Exchange_EXO_Mailboxes_AllDomains.csv')
    $exoStats = Import-SourceCsv @('Exchange_EXO_Mailboxes_AllDomains_Stats.csv')
    $permissionsByUser = Import-SourceCsv @('Exchange_Mailboxes_AllSources_PermissionsByUser.csv')
    $m365Users = Import-SourceCsv @('M365_Users_Active.csv')
    $licenseUsers = Import-SourceCsv @('M365_Licenses_Users.csv')
    $licenseServicePlans = Import-SourceCsv @('M365_Licenses_ServicePlans.csv')
    $acceptedDomains = Import-SourceCsv @('Exchange_EXO_AcceptedDomains.csv')
    $verifiedDomains = Import-SourceCsv @('M365_Entra_VerifiedDomains.csv')
    $backupCoverage = Import-SourceCsv @('M365_BackupPolicyScope_MailboxCoverage.csv')
    $migrationJobs = Import-SourceCsv @('Exchange_EXO_MigrationJobs.csv')
    $hybridIdentityIssues = Import-SourceCsv @('Exchange_HybridIdentity_Issues.csv')
    $personaeKeywordRowsRaw = Import-SourceXlsx @('Personae-Keywords.xlsx') @('PersonaeKeywords')
    $entityRowsRaw = Import-SourceXlsx @('EntityDirectories.xlsx') @('Entities','Entities (2)')
    $personaeKeywordRows1 = @(New-SmartM365PersonaeKeywordRows -Rows $personaeKeywordRowsRaw -StrictCategory:$true)
    $personaeKeywordRows2 = @(New-SmartM365PersonaeKeywordRows -Rows $personaeKeywordRowsRaw -StrictCategory:$false)

    $localByDomainSam = @{}; $localBySmtp = @{}
    foreach ($row in $localMailboxes) { Add-SmartM365MapValue $localByDomainSam (Get-Key (Get-Value $row @('DomainAndSam'))) $row; Add-SmartM365MapValue $localBySmtp (Get-Key (Get-Value $row @('PrimarySMTPaddress','PrimarySmtpAddress'))) $row }
    $remoteByDomainSam = @{}; $remoteBySmtp = @{}
    foreach ($row in $remoteMailboxes) { Add-SmartM365MapValue $remoteByDomainSam (Get-Key (Get-Value $row @('DomainAndSam'))) $row; Add-SmartM365MapValue $remoteBySmtp (Get-Key (Get-Value $row @('PrimarySmtpAddress'))) $row }
    $exoBySmtp = @{}; $exoByExternalId = @{}
    foreach ($row in $exoMailboxes) { Add-SmartM365MapValue $exoBySmtp (Get-Key (Get-Value $row @('PrimarySmtpAddress'))) $row; Add-SmartM365MapValue $exoByExternalId (Get-Key (Get-Value $row @('ExternalDirectoryObjectId'))) $row }
    $exoStatsBySmtp = @{}; foreach ($row in $exoStats) { Add-SmartM365MapValue $exoStatsBySmtp (Get-Key (Get-Value $row @('PrimarySmtpAddress'))) $row }
    $permissionsBySmtp = @{}; foreach ($row in $permissionsByUser) { Add-SmartM365MapValue $permissionsBySmtp (Get-Key (Get-Value $row @('PrimarySMTPaddress','PrimarySmtpAddress'))) $row }
    $m365ByImmutableId = @{}; foreach ($row in $m365Users) { Add-SmartM365MapValue $m365ByImmutableId (Get-Key (Get-Value $row @('OnPremisesImmutableId'))) $row }
    $licenseUsersByUserId = @{}; foreach ($row in $licenseUsers) { Add-SmartM365MapListValue $licenseUsersByUserId (Get-Key (Get-Value $row @('UserId'))) $row }
    $exchangePlanByUserId = @{}
    foreach ($row in $licenseServicePlans) {
        $isEnabled = (Get-Value $row @('IsEnabled')) -match '^(?i:true|1)$'
        $planName = Get-Value $row @('PlanName')
        if ($isEnabled -and @('EXCHANGE_S_ENTERPRISE','EXCHANGE_S_DESKLESS') -contains $planName) {
            Add-SmartM365MapValue $exchangePlanByUserId (Get-Key (Get-Value $row @('UserId'))) (Get-Value $row @('PlanDisplayName','PlanName'))
        }
    }
    $acceptedDomainsByName = @{}; foreach ($row in $acceptedDomains) { Add-SmartM365MapValue $acceptedDomainsByName (Get-Key (Get-Value $row @('DomainName'))) $row }
    $verifiedDomainSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $verifiedDomains) { if (([string](Get-Value $row @('IsVerified'))) -match '^(?i:true|1)$') { [void]$verifiedDomainSet.Add((Get-Key (Get-Value $row @('Id','DomainName')))) } }
    $backupBySmtp = @{}; foreach ($row in $backupCoverage) { Add-SmartM365MapValue $backupBySmtp (Get-Key (Get-Value $row @('PrimarySmtpAddress','MailboxUserPrincipalName','MemberUserPrincipalName'))) $row }
    $migrationBySmtp = @{}; foreach ($row in $migrationJobs) { Add-SmartM365MapValue $migrationBySmtp (Get-Key (Get-Value $row @('MigrationUser','EmailAddress'))) $row }
    $hybridIssuesByObjectGuid = @{}; foreach ($row in $hybridIdentityIssues) { Add-SmartM365MapListValue $hybridIssuesByObjectGuid (Get-Key (Get-Value $row @('ObjectGUID'))) $row }
    $entityByCode = @{}; foreach ($row in $entityRowsRaw) { foreach ($entityKeyCandidate in @((Get-Value $row @('Entity code Text')),(Get-Value $row @('Entity code Text 6 digits')),(Get-Value $row @('Entity code')),(Get-Value $row @('EntityCode')),(Get-Value $row @('OrganizationalUnit')))) { Add-SmartM365MapValue $entityByCode (Get-Key $entityKeyCandidate) $row; $entityNumber = 0.0; if ([double]::TryParse(([string]$entityKeyCandidate).Trim().Replace(',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$entityNumber)) { Add-SmartM365MapValue $entityByCode (([math]::Floor($entityNumber)).ToString('000000', [System.Globalization.CultureInfo]::InvariantCulture)) $row } } }
    $enrichedRows = New-Object System.Collections.Generic.List[object]

    foreach ($user in $users) {
        $out = [ordered]@{}
        foreach ($property in $user.PSObject.Properties) { $out[$property.Name] = $property.Value }
        foreach ($column in $calculatedColumns) { if (-not $out.Contains($column)) { $out[$column] = '' } }

        $primarySmtp = Get-PrimarySmtpFromProxyAddressText -ProxyAddresses ([string]$user.ProxyAddresses)
        $primaryDomain = Get-MailDomain -Mail $primarySmtp
        $upnDomain = Get-MailDomain -Mail ([string]$user.UserPrincipalName)
        $smtpKey = Get-Key $primarySmtp
        $domainAndSamKey = Get-Key (Get-Value $user @('DomainAndSam'))
        $m365User = if ($m365ByImmutableId.ContainsKey((Get-Key (Get-Value $user @('ImmutableId_AD'))))) { $m365ByImmutableId[(Get-Key (Get-Value $user @('ImmutableId_AD')))] } else { $null }
        $aadObjectId = Get-Value $m365User @('Object Id','Id')
        $localMailbox = if ($localByDomainSam.ContainsKey($domainAndSamKey)) { $localByDomainSam[$domainAndSamKey] } elseif ($localBySmtp.ContainsKey($smtpKey)) { $localBySmtp[$smtpKey] } else { $null }
        $remoteMailbox = if ($remoteByDomainSam.ContainsKey($domainAndSamKey)) { $remoteByDomainSam[$domainAndSamKey] } elseif ($remoteBySmtp.ContainsKey($smtpKey)) { $remoteBySmtp[$smtpKey] } else { $null }
        $exoMailbox = if ($exoByExternalId.ContainsKey((Get-Key $aadObjectId))) { $exoByExternalId[(Get-Key $aadObjectId)] } elseif ($exoBySmtp.ContainsKey($smtpKey)) { $exoBySmtp[$smtpKey] } else { $null }
        $exoStat = if ($exoStatsBySmtp.ContainsKey($smtpKey)) { $exoStatsBySmtp[$smtpKey] } else { $null }
        $perms = if ($permissionsBySmtp.ContainsKey($smtpKey)) { $permissionsBySmtp[$smtpKey] } else { $null }
        $acceptedDomain = if ($acceptedDomainsByName.ContainsKey((Get-Key $primaryDomain))) { $acceptedDomainsByName[(Get-Key $primaryDomain)] } else { $null }
        $licenseKey = Get-Key $aadObjectId
        $licenseRows = @()
        if (-not [string]::IsNullOrWhiteSpace($licenseKey) -and $licenseUsersByUserId.ContainsKey($licenseKey)) {
            $licenseRows = @($licenseUsersByUserId[$licenseKey].ToArray())
        }
        $hasDirectLicense = @($licenseRows | Where-Object { (Get-Value $_ @('Source')) -match '^(?i:direct)$' }).Count -gt 0
        $hasGroupLicense = @($licenseRows | Where-Object { (Get-Value $_ @('Source')) -match '^(?i:group)$' }).Count -gt 0
        $exchangePlan = ''
        if ($licenseServicePlans.Count -gt 0) {
            $exchangePlan = if (-not [string]::IsNullOrWhiteSpace($licenseKey) -and $exchangePlanByUserId.ContainsKey($licenseKey)) { [string]$exchangePlanByUserId[$licenseKey] } else { 'Not Exchange Plan Enabled' }
        }
        $backupRow = if ($backupBySmtp.ContainsKey($smtpKey)) { $backupBySmtp[$smtpKey] } else { $null }
        $migrationRow = if ($migrationBySmtp.ContainsKey($smtpKey)) { $migrationBySmtp[$smtpKey] } else { $null }
        $objectGuidKey = Get-Key (Get-Value $user @('ObjectGUID'))
        $hybridIssueRows = @()
        if (-not [string]::IsNullOrWhiteSpace($objectGuidKey) -and $hybridIssuesByObjectGuid.ContainsKey($objectGuidKey)) {
            $hybridIssueRows = @($hybridIssuesByObjectGuid[$objectGuidKey].ToArray())
        }

        $recipientType = Get-Value $localMailbox @('RecipientType')
        if ([string]::IsNullOrWhiteSpace($recipientType)) { $recipientType = Get-Value $remoteMailbox @('RecipientTypeDetails') }
        if ([string]::IsNullOrWhiteSpace($recipientType)) { $recipientType = Get-Value $exoMailbox @('RecipientTypeDetails') }
        if ([string]::IsNullOrWhiteSpace($recipientType)) { $recipientType = 'NoMailboxes' }
        $mailboxSourceParts = @()
        if ($null -ne $localMailbox) { $mailboxSourceParts += 'Local' }
        if ($null -ne $remoteMailbox) { $mailboxSourceParts += 'Remote' }
        if ($null -ne $exoMailbox) { $mailboxSourceParts += 'EXO' }
        $mailboxSource = if ($mailboxSourceParts.Count -gt 0) { $mailboxSourceParts -join '+' } else { 'None' }
        $itemCount = [double]0
        foreach ($candidate in @((Get-Value $localMailbox @('ItemCount','ItemCount_Number')),(Get-Value $exoMailbox @('ItemCount')),(Get-Value $exoStat @('ItemCount')))) { $n = Get-NumberInvariant $candidate; if ($null -ne $n -and $n -gt $itemCount) { $itemCount = $n } }
        $localSizeMb = Get-NumberInvariant (Get-Value $localMailbox @('TotalItemSize-In-MB','TotalItemSizeInMB_Numeric'))
        $exoSizeGb = Get-NumberInvariant (Get-Value $exoMailbox @('TotalItemSizeGB'))
        if ($null -eq $exoSizeGb) { $exoSizeGb = Get-NumberInvariant (Get-Value $exoStat @('TotalItemSizeGB')) }
        $totalSizeMb = if ($null -ne $localSizeMb) { $localSizeMb } elseif ($null -ne $exoSizeGb) { $exoSizeGb * 1024 } elseif ($localMailboxes.Count -gt 0 -or $exoMailboxes.Count -gt 0 -or $exoStats.Count -gt 0) { 0 } else { $null }
        $localArchiveMb = Get-NumberInvariant (Get-Value $localMailbox @('ArchiveTotalItemSize-In-MB','ArchiveTotalItemSizeInMB_Numeric'))
        $exoArchiveGb = Get-NumberInvariant (Get-Value $exoMailbox @('Archive_TotalItemSizeGB','ArchiveTotalItemSizeGB'))
        if ($null -eq $exoArchiveGb) { $exoArchiveGb = Get-NumberInvariant (Get-Value $exoStat @('Archive_TotalItemSizeGB')) }
        $archiveMb = if ($null -ne $localArchiveMb) { $localArchiveMb } elseif ($null -ne $exoArchiveGb) { $exoArchiveGb * 1024 } elseif ($localMailboxes.Count -gt 0 -or $exoMailboxes.Count -gt 0 -or $exoStats.Count -gt 0) { 0 } else { $null }
        $totalAndArchiveMb = if ($null -ne $totalSizeMb -or $null -ne $archiveMb) { ($(if ($null -ne $totalSizeMb) { $totalSizeMb } else { 0 }) + $(if ($null -ne $archiveMb) { $archiveMb } else { 0 })) } else { $null }
        $licenseFromM365 = if ($m365Users.Count -eq 0) { '' } elseif ($null -eq $m365User) { 'Not in M365' } elseif ($licenseRows.Count -eq 0) { 'No licenses' } else { Join-UniqueText -Values @($licenseRows | ForEach-Object { Get-Value $_ @('SKU name','SkuPartNumber') }) -Separator ',' }
        $licenseGroup = if (@('SharedMailbox','RemoteSharedMailbox') -contains $recipientType) { 'No License (Shared)' } elseif (@('RoomMailbox','RemoteRoomMailbox','EquipmentMailbox','RemoteEquipmentMailbox') -contains $recipientType) { 'No License (Resource Mailbox)' } elseif ($licenseFromM365 -eq 'Not in M365') { 'Not in M365' } elseif ($licenseFromM365 -eq 'No licenses') { 'No License (Unlicensed)' } elseif ($licenseFromM365 -match 'Microsoft 365 E3') { 'Microsoft 365 E3' } elseif ($licenseFromM365 -match 'Microsoft 365 F3') { 'Microsoft 365 F3' } elseif ($licenseFromM365 -match 'Microsoft 365 F1') { 'Microsoft 365 F1' } elseif (-not [string]::IsNullOrWhiteSpace($licenseFromM365)) { 'No License (Unlicensed)' } else { '' }
        $licenseGroups = @(); if ($licenseFromM365 -match 'Microsoft 365 E3') { $licenseGroups += 'E3' }; if ($licenseFromM365 -match 'Microsoft 365 F3') { $licenseGroups += 'F3' }; if ($licenseFromM365 -match 'Microsoft 365 F1') { $licenseGroups += 'F1' }
        $fullAccess = Get-Value $perms @('FullAccessOn'); $sendAs = Get-Value $perms @('SendAsOn'); $sendOnBehalf = Get-Value $perms @('SendOnBehalfOn')
        $permissionsSummary = (Join-UniqueText -Values @($fullAccess,$sendAs,$sendOnBehalf) -Separator '; ').Replace(' [Local]','').Replace(' [EXO]','')
        $hasLastLogon = -not [string]::IsNullOrWhiteSpace([string]$user.LastLogonDate)

        $distinguishedName = [string]$user.DistinguishedName
        $distinguishedNameUpper = $distinguishedName.ToUpperInvariant()
        $samAccountName = [string]$user.SamAccountName
        $samAccountNameLower = $samAccountName.ToLowerInvariant()
        $userPrincipalName = [string]$user.UserPrincipalName
        $enabledText = ([string]$user.Enabled).Trim()
        $isEnabled = $enabledText -match '^(?i:true|1)$'
        $orgUnit = Get-SmartM365OrganizationalUnitCode -Row $user
        $entity = if ($entityByCode.ContainsKey((Get-Key $orgUnit))) { $entityByCode[(Get-Key $orgUnit)] } else { $null }
        $typeEtablissement = Get-SmartM365TypeEtablissement -EntityRow $entity
        $typeEtablissementUpper = $typeEtablissement.Trim().ToUpperInvariant()
        $typeEntityValue = if ($null -ne $entity) { Get-Value $entity @('Entity Type','EntityType') } else { '' }
        if ([string]::IsNullOrWhiteSpace($typeEntityValue)) { $typeEntityValue = Get-TypeEntityFallback -DomainNameShort (Get-Value $user @('DomainNameShort')) -Country (Get-Value $user @('Country')) }
        $labelFromEntity = if ($null -ne $entity) { Get-Value $entity @('Entity label','EntityLabel') } else { 'NOTFOUND' }
        $inTargetHqOu = $distinguishedNameUpper.Contains('OU=020001-MADRID') -or $distinguishedNameUpper.Contains('OU=120001-HEAD_QUARTER') -or $distinguishedNameUpper.Contains('OU=HEADQUARTER')
        $inDisabledObjectsOu = $distinguishedNameUpper.Contains('OU=DISABLED_OBJECTS')
        $givenNameClean = Convert-SmartM365NameClean -Value (Get-Value $user @('GivenName'))
        $surnameClean = Convert-SmartM365NameClean -Value (Get-Value $user @('Surname'))
        $accountType = Get-SmartM365AccountType -UPN $userPrincipalName -SAM $samAccountName -DN $distinguishedName -RecipientType $recipientType -GivenName (Get-Value $user @('GivenName')) -Surname (Get-Value $user @('Surname')) -GivenNameClean $givenNameClean -SurnameClean $surnameClean
        $jobTitleFromDesc = Get-SmartM365JobTitleFromDescription -Description (Get-Value $user @('Description'))
        $jobTitleFromDesc2 = $jobTitleFromDesc
        $jobTitleNormalized = Normalize-SmartM365JobTitleText -ManagedTitle '' -DescriptionTitle $jobTitleFromDesc
        $jobTitleNormalized2 = Normalize-SmartM365JobTitleText -ManagedTitle '' -DescriptionTitle $jobTitleFromDesc2
        $jobTitlePadded = if ([string]::IsNullOrWhiteSpace($jobTitleNormalized)) { '' } else { ' ' + $jobTitleNormalized + ' ' }
        $jobTitlePadded2 = if ([string]::IsNullOrWhiteSpace($jobTitleNormalized2)) { '' } else { ' ' + $jobTitleNormalized2 + ' ' }
        $jobFamilyFromKeywords = Get-SmartM365JobFamilyFromKeywords -PaddedTitle $jobTitlePadded -KeywordRows $personaeKeywordRows1 -StrictCategory:$true
        $jobFamilyFromKeywords2 = Get-SmartM365JobFamilyFromKeywords -PaddedTitle $jobTitlePadded2 -KeywordRows $personaeKeywordRows2 -StrictCategory:$false
        $noLastLogonInOnComputer = [string]::IsNullOrWhiteSpace([string]$out['LastLoggedInOnComputer'])
        $jobFamilyCombined = Get-SmartM365CombinedJobFamily -AccountType $accountType -TypeEtablissement $typeEtablissement -TypeEntity $typeEntityValue -UserPrincipalName $userPrincipalName -NoLastLogonInOnComputer $noLastLogonInOnComputer -KeywordCategory $jobFamilyFromKeywords
        $jobFamilyCombined2 = Get-SmartM365CombinedJobFamily -AccountType $accountType -TypeEtablissement $typeEtablissement -TypeEntity $typeEntityValue -UserPrincipalName $userPrincipalName -NoLastLogonInOnComputer $noLastLogonInOnComputer -KeywordCategory $jobFamilyFromKeywords2
        $target1Persona = if (@('UserMailbox','RemoteUserMailbox') -contains $recipientType) { Get-SmartM365BasePersona -JobFamily $jobFamilyCombined } else { 'None' }
        $target1Persona2 = if (@('UserMailbox','RemoteUserMailbox') -contains $recipientType) { Get-SmartM365BasePersona -JobFamily $jobFamilyCombined2 } else { 'None' }
        $target2Persona = if ($target1Persona -eq 'M365F3') { if ($totalSizeMb -ge 50000) { 'M365F3+EXCHPLAN2' } elseif ($totalSizeMb -gt 2000) { 'M365F3+EXCHPLAN1' } else { 'M365F3' } } else { $target1Persona }
        $target3Persona = if ($target1Persona -eq 'M365F3') { if ($totalSizeMb -gt 2000) { 'M365F3 > 2G' } else { 'M365F3 < 2G' } } else { $target1Persona }
        $licenseGroupNorm = $licenseGroup.Trim().ToUpperInvariant()
        $allowedAccountType = @('NAMED ACCOUNT','GENERIC ACCOUNT','UNCLASSIFIED ACCOUNT','EXT ACCOUNT') -contains $accountType.ToUpperInvariant()
        $m365LicenseType = if ($licenseGroupNorm.Contains('MICROSOFT 365 E3')) { 'M365E3' } elseif ($licenseGroupNorm.Contains('MICROSOFT 365 F3')) { 'M365F3' } elseif (-not $allowedAccountType) { 'None' } else { Get-SmartM365BasePersona -JobFamily $jobFamilyCombined }
        $lastLogonEntra = Get-Value $m365User @('LastSignInDateTime')
        $lastLogonEntraNonInteractive = Get-Value $m365User @('LastNonInteractiveSignInDateTime')
        $lastLogonMergeAdExchEntra = Get-SmartM365LatestDateText -Values @([string]$user.LastLogonDate, $out['LastLogonDateFromExch'], $lastLogonEntra, $lastLogonEntraNonInteractive)
        $out['OU_Path'] = Get-OuPathFromDistinguishedName -DistinguishedName $distinguishedName
        $out['OU_Category'] = if ($distinguishedNameUpper.Contains('OU=DISABLED_OBJECTS')) { 'DISABLED OBJECTS' } elseif ($distinguishedNameUpper.Contains('OU=ORGANIZATION')) { 'ORGANIZATION' } elseif ($distinguishedNameUpper.Contains('OU=ADMIN')) { 'ADMIN' } elseif ($distinguishedNameUpper.Contains('OU=FAX')) { 'FAX' } else { 'OTHER' }
        $out['TypeEtablissement'] = $typeEtablissement
        $out['TypeEntity'] = $typeEntityValue
        $out['GivenNameClean'] = $givenNameClean
        $out['SurnameClean'] = $surnameClean
        $out['AccountType'] = $accountType
        $out['EXO_Mapping_Status'] = if ($mailboxSource -ne 'EXO') { 'Not EXO' } elseif (-not [string]::IsNullOrWhiteSpace($aadObjectId)) { 'OK_GUID' } else { 'Missing_AAD_Object' }
        $out['IsShared_Status'] = if ($null -eq $localMailbox) { 'NoMailboxes' } else { [string]((Get-Value $localMailbox @('IsShared')) -match '^(?i:true|1)$') }
        $out['IsLargeSharedMailbox'] = [string](($totalSizeMb -gt 40960) -and $recipientType -eq 'SharedMailbox')
        $out['IsInLocalEXODuplicates'] = [string]($mailboxSource -match 'Local' -and $mailboxSource -match 'EXO')
        $out['Is_In_SHAREDMAILBOX'] = [string]($recipientType -match 'SharedMailbox')
        $out['TotalItemSizeNumeric_From_Mailboxes'] = Format-NumberInvariant $localSizeMb
        $out['TotalItemSizeNumeri_From_RemoteMailboxes'] = Format-NumberInvariant (Get-NumberInvariant (Get-Value $remoteMailbox @('TotalItemSizeGB')))
        $out['M365LicenseType_Target1_ByPersona'] = $target1Persona
        $out['M365LicenseType_Target1_ByPersona 2'] = $target1Persona2
        $out['M365LicenseType_Target2_ByPersona'] = $target2Persona
        $out['M365LicenseType_Target3_ByPersona'] = $target3Persona
        $out['M365LicenseType_ByPersonae'] = $target1Persona
        $out['M365LicenseType_ByPersonae2'] = $target2Persona
        $out['M365LicenseType'] = $m365LicenseType
        $out['HeavyMailboxF3User'] = if ((($licenseGroup -eq 'Microsoft 365 F3') -or ($m365LicenseType -eq 'M365F3')) -and $totalSizeMb -ge 2000) { 'Yes' } else { 'No' }
        $out['M365License_F3_E3_SizeBucket'] = if ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 2000 -and $totalSizeMb -le 2500) { '2-2.5 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 2500 -and $totalSizeMb -le 5000) { '2.5-5 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 5000 -and $totalSizeMb -le 10000) { '5-10 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 10000) { '>10 GB' } else { '' }
        $out['M365License_F3_E3_SizeBucket_Sort'] = switch ($out['M365License_F3_E3_SizeBucket']) { '2-2.5 GB' { '1' } '2.5-5 GB' { '2' } '5-10 GB' { '3' } '>10 GB' { '4' } default { '99' } }
        $out['M365License_F3_NonCompliant_Size_Remaining_Flag'] = if ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $recipientType -eq 'UserMailbox' -and $out['ExistsInRemoteMailboxes'] -eq 'False' -and $totalSizeMb -gt 2000 -and $licenseGroup -ne 'Microsoft 365 E3') { '1' } else { '0' }
        $out['M365License_F3_AllOversized_SizeBucket'] = if ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 2000 -and $totalSizeMb -le 2500) { '2-2.5 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 2500 -and $totalSizeMb -le 5000) { '2.5-5 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 5000 -and $totalSizeMb -le 10000) { '5-10 GB' } elseif ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 10000) { '>10 GB' } else { '' }
        $out['M365License_F3_AllOversized_SizeBucket_Sort'] = switch ($out['M365License_F3_AllOversized_SizeBucket']) { '2-2.5 GB' { '1' } '2.5-5 GB' { '2' } '5-10 GB' { '3' } '>10 GB' { '4' } default { '99' } }
        $out['M365License_NoLicenseF3_EnabledMailbox_BlockingMigration_Flag'] = if ($out['IsMailBox'] -eq 'True' -and $isEnabled -and $recipientType -eq 'UserMailbox' -and $out['ExistsInRemoteMailboxes'] -eq 'False' -and $target1Persona -eq 'M365F3' -and @('No License (Unlicensed)','Not in M365') -contains $licenseGroup) { '1' } else { '0' }
        $out['M365License_E3Persona_MissingE3_Flag'] = if ($out['IsMailBox'] -eq 'True' -and $target1Persona -eq 'M365E3' -and @('UserMailbox','RemoteUserMailbox') -contains $recipientType -and $out['ExistsInRemoteMailboxes'] -eq 'False' -and @('Microsoft 365 F3','No License (Unlicensed)') -contains $licenseGroup) { '1' } else { '0' }
        $out['LastLogonEntra'] = $lastLogonEntra
        $out['LastLogonEntraNonInteractive'] = $lastLogonEntraNonInteractive
        $out['LastLogonEntra_Category'] = Get-DateCategory -Value $lastLogonEntra -LimitDays 90
        $out['LastLogonMergeADExchEntra'] = $lastLogonMergeAdExchEntra
        $out['NoLastLogonEntra'] = [string][string]::IsNullOrWhiteSpace($lastLogonEntra)
        $out['NoLastLogonEntraNonInteractive'] = [string][string]::IsNullOrWhiteSpace($lastLogonEntraNonInteractive)
        $out['NoLastLogonExch'] = [string][string]::IsNullOrWhiteSpace([string]$out['LastLogonDateFromExch'])
        $out['NoLastLogonInOnComputer'] = [string][string]::IsNullOrWhiteSpace([string]$out['LastLoggedInOnComputer'])
        $out['FullAccess_From_Mailboxes_UPN'] = $out['FullAccess_From_Mailboxes']
        $out['SendAs_From_Mailboxes_UPN'] = $out['SendAs_From_Mailboxes']
        $out['GrantSendOnBehalfTo_UPN'] = $out['GrantSendOnBehalfTo_From_Mailboxes']
        $out['FullAccess_From_MailboxesEXO'] = Get-Value $perms @('FullAccessEXO','FullAccess_From_MailboxesEXO')
        $out['FullAccess_From_MailboxesEXO_Raw'] = $out['FullAccess_From_MailboxesEXO']
        $out['LocalMailboxPermissionCount_NotInM365'] = '0'
        $out['LogonCountStatus'] = ''
        $out['ActivityCategory'] = ''
        $out['AccountToCleanExcluded'] = ''
        $out['AccountToCleanStatus1'] = ''
        $out['AccountToCleanStatus2'] = ''
        $out['IsJobTitleValid'] = [string](-not [string]::IsNullOrWhiteSpace($jobTitleNormalized))
        $out['Job Family Combined'] = $jobFamilyCombined
        $out['Job Family Combined 2'] = $jobFamilyCombined2
        $out['Job Family from Keywords'] = $jobFamilyFromKeywords
        $out['Job Family from Keywords 2'] = $jobFamilyFromKeywords2
        $out['Job Title Combined'] = $jobTitleFromDesc
        $out['Job Title From Desc'] = $jobTitleFromDesc
        $out['Job Title From Desc 2'] = $jobTitleFromDesc2
        $out['Job Title Normalized'] = $jobTitleNormalized
        $out['Job Title Normalized 2'] = $jobTitleNormalized2
        $out['Job Title Padded'] = $jobTitlePadded
        $out['Job Title Padded 2'] = $jobTitlePadded2
        $out['LabelFromEntity'] = $labelFromEntity
        $out['LastLoggedComputerSamAccountNames'] = ''
        $out['LastLoggedInOnALLComputers'] = ''
        $out['LastLoggedInOnComputer'] = ''
        if ($hybridIdentityIssues.Count -eq 0 -or [string]::IsNullOrWhiteSpace($primarySmtp)) {
            $out['Mig_Potential_Issues_Mailbox_Level2'] = ''
            $out['Mig_Potential_Issues_Score'] = ''
        }
        else {
            $severityCodes = @()
            $issueScore = 0
            foreach ($hybridIssueRow in $hybridIssueRows) {
                $category = Get-Value $hybridIssueRow @('IssueCategory','Category')
                $match = [regex]::Match($category, '^\s*(?<Code>[1-4])')
                if (-not $match.Success) { continue }
                $code = [int]$match.Groups['Code'].Value
                $severityCodes += $code
                switch ($code) {
                    1 { $issueScore += 3 }
                    2 { $issueScore += 2 }
                    3 { $issueScore += 1 }
                    default { }
                }
            }
            if ($severityCodes.Count -eq 0) {
                $out['Mig_Potential_Issues_Mailbox_Level2'] = 'None'
            }
            else {
                switch (($severityCodes | Measure-Object -Minimum).Minimum) {
                    1 { $out['Mig_Potential_Issues_Mailbox_Level2'] = 'Critical' }
                    2 { $out['Mig_Potential_Issues_Mailbox_Level2'] = 'High' }
                    3 { $out['Mig_Potential_Issues_Mailbox_Level2'] = 'Medium' }
                    default { $out['Mig_Potential_Issues_Mailbox_Level2'] = 'Low' }
                }
            }
            $out['Mig_Potential_Issues_Score'] = [string]$issueScore
        }
        $out['IsServiceAccountOrSimilar'] = [string]([string]::IsNullOrWhiteSpace($userPrincipalName) -or $distinguishedName.Contains('Service_Accounts') -or $samAccountName.Contains('s_') -or $samAccountNameLower.Contains('svc') -or $samAccountName.Contains('HealthMailbox'))
        $out['IsSpecificAdminOrService'] = [string]($samAccountNameLower.StartsWith('a_') -or $samAccountNameLower.Contains('s_') -or $samAccountNameLower.Contains('administrateur') -or $samAccountNameLower.Contains('administrator'))
        $out['IsTargetOU_HQ'] = [string](($typeEtablissementUpper -eq 'HQ') -or $inTargetHqOu)
        $out['IsNotTargetOU_HQ'] = [string](($typeEtablissementUpper -ne 'HQ') -and -not $inTargetHqOu -and -not $inDisabledObjectsOu)
        $out['NoLastLogonAD'] = [string]([string]::IsNullOrWhiteSpace([string]$user.LastLogonDate))
        $out['PrimarySmtp'] = $primarySmtp
        $out['PrimarySmtp_Lower'] = $primarySmtp
        $out['PrimarySMTP_Domain'] = $primaryDomain
        $out['Containsorpeamailonmicrosoft'] = if (([string]$user.ProxyAddresses).ToLowerInvariant().Contains('@orpea.mail.onmicrosoft.com')) { 'Yes' } else { 'No' }
        $out['ExistsInM365ActiveUsers'] = if ($m365Users.Count -gt 0) { [string]($null -ne $m365User) } else { '' }
        $out['RecipientType_From_Mailboxes'] = $recipientType
        $out['RecipientTypeDetails_From_Mailboxes'] = $recipientType
        $out['ExistsInRemoteMailboxes'] = [string](@('RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox') -contains $recipientType)
        $out['IsMailBox'] = [string]($recipientType -ne 'NoMailboxes')
        $out['MailboxSource'] = $mailboxSource
        $out['MailboxLocation'] = $mailboxSource
        $out['MailboxSourceHealthStatus'] = if ($mailboxSource -eq 'None') { 'No mailbox' } elseif ($mailboxSource -eq 'Local+Remote+EXO') { 'Hybrid duplicate' } else { 'OK' }
        $out['ItemCount_From_Mailboxes'] = $itemCount.ToString('0.########',[System.Globalization.CultureInfo]::InvariantCulture)
        $out['IsItemCountOver999999'] = [string]($itemCount -gt 999999)
        $out['TotalItemSizeNumeric'] = Format-NumberInvariant $totalSizeMb
        $out['TotalItemSize_GB'] = if ($null -ne $totalSizeMb) { Format-NumberInvariant ($totalSizeMb / 1024) } else { '' }
        $out['TotalItemSize_TB'] = if ($null -ne $totalSizeMb) { Format-NumberInvariant ($totalSizeMb / 1048576) } else { '' }
        $out['TotalArchiveItemSizeNumeric'] = Format-NumberInvariant $archiveMb
        $out['TotalArchiveItemSize_GB'] = if ($null -ne $archiveMb) { Format-NumberInvariant ($archiveMb / 1024) } else { '' }
        $out['TotalItemSizeAndArchiveNumeric'] = Format-NumberInvariant $totalAndArchiveMb
        $out['TotalItemSizeAndArchive_GB'] = if ($null -ne $totalAndArchiveMb) { Format-NumberInvariant ($totalAndArchiveMb / 1024) } else { '' }
        $out['IsLargerThan2GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -gt 2048) } else { 'False' }
        $out['IsLargerThan100GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -gt 102400) } else { 'False' }
        $out['IsSmallerThan2GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -lt 2048) } else { 'False' }
        $out['MailboxSize_From_Mailboxes'] = $out['TotalItemSize_GB']
        $out['LicenseFromM365'] = $licenseFromM365
        $out['LicenseGroupFromM365'] = $licenseGroup
        $out['LicenseExchangePlanEnabledFromM365'] = $exchangePlan
        $out['LicenseGroupCount'] = [string]$licenseGroups.Count
        $out['LicenseGroupFound'] = $licenseGroups -join ' + '
        $out['LicenseFromM365_GroupsAssigningSku'] = Join-UniqueText -Values @($licenseRows | ForEach-Object { Get-Value $_ @('GroupsAssigningSku') }) -Separator ','
        $out['LicenseFromM365_HasDirect'] = if ([string]::IsNullOrWhiteSpace($licenseKey)) { '' } else { [string]$hasDirectLicense }
        $out['LicenseFromM365_HasGroupOnly'] = if ([string]::IsNullOrWhiteSpace($licenseKey)) { '' } else { [string]($hasGroupLicense -and -not $hasDirectLicense) }
        $out['LicenseFromM365_HasDirectAndGroup'] = [string](@($licenseRows | Where-Object { (Get-Value $_ @('HasDirectAndGroup')) -match '^(?i:true|1)$' }).Count -gt 0)
        $out['LicenseGroupHasDoubleAssignment'] = if ($out['LicenseFromM365_HasDirectAndGroup'] -eq 'True') { 'Yes' } else { 'No' }
        $out['DomainAllowed_EXO'] = if ($acceptedDomains.Count -gt 0) { [string]($null -ne $acceptedDomain) } else { '' }
        $out['DomainNotAllowed_EXO'] = if ($out['DomainAllowed_EXO'] -eq '') { '' } else { [string]($out['DomainAllowed_EXO'] -eq 'False') }
        $out['DomainType_EXO'] = Get-Value $acceptedDomain @('DomainType')
        $out['DomainIsOnMicrosoft_EXO'] = if ([string]::IsNullOrWhiteSpace($primaryDomain)) { '' } else { [string]($primaryDomain -like '*.onmicrosoft.com') }
        $out['DomainNonCompliant_EXO'] = if ($out['DomainAllowed_EXO'] -eq '') { '' } elseif ($out['DomainAllowed_EXO'] -eq 'False' -or $out['DomainIsOnMicrosoft_EXO'] -eq 'True') { 'True' } else { 'False' }
        $out['DomainUsersNotAllowed_EXO'] = $out['DomainNotAllowed_EXO']
        $out['DomainUsersNonCompliant_EXO'] = $out['DomainNonCompliant_EXO']
        $out['UPNDomainVerified_AAD'] = if ([string]::IsNullOrWhiteSpace($upnDomain) -or $verifiedDomains.Count -eq 0) { '' } else { [string]$verifiedDomainSet.Contains((Get-Key $upnDomain)) }
        $out['UPNEqualsPrimarySMTPDomain'] = if ([string]::IsNullOrWhiteSpace($upnDomain) -or [string]::IsNullOrWhiteSpace($primaryDomain)) { '' } else { [string]($upnDomain -eq $primaryDomain) }
        $out['RemoteRoutingAddress_From_MailboxesRemote'] = Get-Value $remoteMailbox @('RemoteRoutingAddress')
        $out['FullAccessOn_From_Mailboxes'] = $fullAccess
        $out['SendAsOn_From_Mailboxes'] = $sendAs
        $out['SendOnBehalfOn_From_Mailboxes'] = $sendOnBehalf
        $out['FullAccess_From_Mailboxes_Raw'] = $fullAccess
        $out['FullAccess_From_Mailboxes'] = $fullAccess
        $out['SendAs_From_Mailboxes_Raw'] = $sendAs
        $out['SendAs_From_Mailboxes'] = $sendAs
        $out['GrantSendOnBehalfTo_From_Mailboxes_Raw'] = $sendOnBehalf
        $out['GrantSendOnBehalfTo_From_Mailboxes'] = $sendOnBehalf
        $out['PermissionsSummary'] = $permissionsSummary
        $out['PermissionsCount'] = if ([string]::IsNullOrWhiteSpace($permissionsSummary)) { '0' } else { [string](([string]$permissionsSummary).Split(';').Count) }
        $out['SharedAccessCount'] = if ([string]::IsNullOrWhiteSpace($sendAs)) { '0' } else { [string](([string]$sendAs).Split(';').Count) }
        $out['HasFullAccess'] = [string](-not [string]::IsNullOrWhiteSpace($fullAccess))
        $out['HasSendAsBehalf'] = [string](-not [string]::IsNullOrWhiteSpace($sendAs))
        $out['HasSendOnBehalf'] = [string](-not [string]::IsNullOrWhiteSpace($sendOnBehalf))
        $out['HasAnyDelegPerm'] = [string]((-not [string]::IsNullOrWhiteSpace($fullAccess)) -or (-not [string]::IsNullOrWhiteSpace($sendAs)) -or (-not [string]::IsNullOrWhiteSpace($sendOnBehalf)))
        $out['HasCrossPremisesPermissions'] = Get-Value $perms @('HasCrossPremisesPermissions')
        $out['MigrationUserStatus_From_MigrationJobs'] = if ($recipientType -like '*Remote*') { 'Completed' } else { Get-Value $migrationRow @('UserStatus') }
        $out['MigrationUserBatchName_From_MigrationJobs'] = Get-Value $migrationRow @('BatchName')
        $out['MigrationUserCompleteAfter_From_MigrationJobs'] = Get-Value $migrationRow @('CompleteAfterUser','CompleteAfter')
        $out['MigrationUserCompleteAfterUTC_From_MigrationJobs'] = Get-Value $migrationRow @('CompleteAfterUTCUser','CompleteAfterUTC')
        $out['ProtectedMailboxes_From_EXO_BackupProtection'] = if ($backupCoverage.Count -gt 0) { [string]($null -ne $backupRow) } else { '' }
        $out['LastLogonDate_Category'] = Get-DateCategory -Value ([string]$user.LastLogonDate) -LimitDays 90
        $out['IsLastLogonMerge'] = if ($hasLastLogon) { 'Known' } else { 'Unknown' }
        $out['NoLastLogonMerge'] = [string](-not $hasLastLogon)
        $out['LastLogonMerge'] = if ($hasLastLogon) { [string]$user.LastLogonDate } else { '' }
        $out['EnabledAndNoLastLogonMerge'] = [string]($isEnabled -and [string]::IsNullOrWhiteSpace([string]$out['LastLogonMerge']))

        [void]$enrichedRows.Add([pscustomobject]$out)
    }

    $enrichedCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Users_AllDomains_Enriched.csv'
    $enrichedRows | Add-SmartM365TenantKey | Export-Csv -LiteralPath $enrichedCsv -NoTypeInformation -Encoding UTF8
    if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
        $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
        $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) {
            if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) { [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath) }
        }
    }
    [void]$global:csvGeneratedPaths.Add($enrichedCsv)
    WriteLog -Message ("AD users enriched CSV generated: {0} ({1} row(s), {2} enriched column(s))" -f $enrichedCsv, $enrichedRows.Count, $calculatedColumns.Count)
    return $enrichedCsv
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAUrYq2LNDxBkoK
# dPI2GtUIETocBkD3RL1nSGJs8K2R+aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMQW8empj6g+l5DT1qSpi/04yNyEansG/q1J0U13mB6kMA0GCSqG
# SIb3DQEBAQUABIIBgAQr8wySMI2i8KAK4i32oP9NAIXmDTkNV6Y3XLJn3V+byAPL
# t+At31LzmXS8TkOBALQYWNx3mH9VbOFURXhvdc4MoZBFDT2Og9XopVWwrvnBIRo2
# U67SVOzdWX47KWRWSFs7eefsVDQ7n59yXAX6Dmm4XcO3xWXUkO24ss1E2D8mjH/0
# hKqOtThE86iAakeyBN0rUmG8XZB8oNfbl8zhm2i1jfDqHgSg+pZmCuXBdVygkDtg
# mzsA47AjucpyPPZTxQ0WGbVnqlXMx59PUJ6wYgowzb9tqseGPuOjA5DMJ09UwVpY
# eHPk+Fvnl/l5N4eZG3IQmN0E2Sn6FqwtJ3kzN9wPytMOw9+AQcPvTA/l9kcHRnFW
# eJIzTbTbR2XhREdEeWHo4dD0dycWadG0EtRvwSaaFpd4zNNf3BDII5CHC+FFbcht
# KllTkEBAOFCSSbfsy8aZ6tJRmbVKY2K2rBiZZmJbvIkvlqlbY9T8DvNNR7+d1jPJ
# s/nRKr0RO2IqLnRW4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwNjI1
# MzZaMC8GCSqGSIb3DQEJBDEiBCBb/hQvnYGXCBN48fmIBK9YGZEd4auMinhX6eAI
# 4wvh0zANBgkqhkiG9w0BAQEFAASCAgBbAiLm2uHAzaB0nUOyD5mC/aUaOda14Hvy
# 7XjolLjDzuFAUZJYcJONMu3Dgl8t+OuIUrNIaths4ttMOMGl59pHWrg8pCPfZQtH
# zVWsxAbO7Xt4baibwo7lqXB77MbLgOnotSSEfyJ2lQuSveAptwGUQAqCNbopX/1O
# Wrn4NRq1KMhigSeR3g9UBaF9tnZsFfVRiGQCBA1Y2sN3vSBy76jpAcAD3dz89baW
# M8AngE1Y98rMaAyWbbrm/jHdTO0yvca5Hcjh8vpGGkglFn1tf56XaEo1GFBX4mac
# n63iD9VKUqOiWbo+zT00ZC3fYm8pZ7PBYLzpTsoM4NfX7ieSkAZM0AEMpWp5HPdu
# mzkw06WZDVDl1apDMptImVRYncAsb2MOVoswJfNmFWGofjtHQ03r7U2D52p/9ChZ
# 8RKbCAyPszfR2S494gJGo04mlbWiYwa55K9g/IPifwLwDIkG2e9C1y83wc1a2IvG
# xmZYDgVrnVUclv1NTUYyDgqe4Xkq+JcuK5K9Hd4nHHVrVkSeYqnNjv2dr29ZLNo5
# 6nsvvpW5CRWQuwSSs6ei9qU7/+VxfLadKe9clxv8XFV5lZbFGL3Zb8kmdR4Bsz+K
# TQpB8Gid7JHYD6jJaGWDhBoBdrPClgaiTa3FDXvFinZgFPWGnR9xCaCKKPhiq4bE
# pzSsr6PcaA==
# SIG # End signature block
