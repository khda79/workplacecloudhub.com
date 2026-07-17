<#
.SYNOPSIS
    Builds enriched Active Directory user CSV columns required by the SmartWorkplace Power BI model.
.VERSION
1.5
#>

function Invoke-SmartM365AdUsersEnrichedCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CombinedUsersCsv,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $false)][string]$LatestFolderPath,
        [Parameter(Mandatory = $true)][string]$RemoteRoutingDomain
    )

    $RemoteRoutingDomain = $RemoteRoutingDomain.Trim().TrimStart('@').ToLowerInvariant()
    if ($RemoteRoutingDomain -notmatch '^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+mail\.onmicrosoft\.com$') {
        throw ("Invalid RemoteRoutingDomain passed to AD user enrichment: {0}" -f $RemoteRoutingDomain)
    }

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
        $resolvedPath = ''
        foreach ($candidate in (Get-SourceCandidatePaths -Names $Names)) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $resolvedPath = $candidate; break }
        }
        if ([string]::IsNullOrWhiteSpace($resolvedPath) -and (Get-Command Resolve-SmartM365AdReferenceXlsx -ErrorAction SilentlyContinue)) {
            foreach ($name in $Names) {
                $sharePointPath = Resolve-SmartM365AdReferenceXlsx -Name $name
                if (-not [string]::IsNullOrWhiteSpace($sharePointPath) -and (Test-Path -LiteralPath $sharePointPath -PathType Leaf)) {
                    $resolvedPath = $sharePointPath
                    break
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
            WriteLog -Message ("AD users enrichment optional Excel source missing; related columns will be blank: {0}" -f ($Names -join ', '))
            return @()
        }
        try { Import-Module ImportExcel -ErrorAction Stop }
        catch { WriteLog -Message ("AD users enrichment Excel source skipped because ImportExcel module is not available: {0}" -f $resolvedPath); return @() }
        $rowsOut = New-Object System.Collections.Generic.List[object]
        foreach ($worksheetName in @($WorksheetNames)) {
            if ([string]::IsNullOrWhiteSpace($worksheetName)) { continue }
            try { foreach ($row in @(Import-Excel -Path $resolvedPath -WorksheetName $worksheetName -ErrorAction Stop)) { [void]$rowsOut.Add($row) } }
            catch { WriteLog -Message ("AD users enrichment Excel worksheet skipped: {0} [{1}] ({2})" -f $resolvedPath, $worksheetName, $PSItem.Exception.Message) }
        }
        WriteLog -Message ("AD users enrichment Excel source loaded: {0} ({1} row(s))" -f $resolvedPath, $rowsOut.Count)
        return $rowsOut.ToArray()
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
        'AdOrganizationalUnitPath'
        'AdOrganizationalUnitCategory'
        'EntitySiteType'
        'EntityIntegrationType'
        'IsLikelyServiceAccount'
        'IsLikelyPrivilegedOrServiceAccount'
        'IsInTargetHeadquartersOU'
        'IsOutsideTargetHeadquartersOU'
        'HasNoAdLastLogon'
        'IsEnabledWithNoResolvedLastLogon'
        'GivenNameNormalized'
        'SurnameNormalized'
        'AccountType'
        'ExchangeOnlineMappingStatus'
        'SharedMailboxStatus'
        'IsLargeSharedMailbox'
        'IsInLocalExchangeOnlineDuplicateSet'
        'IsSharedMailbox'
        'OnPremMailboxSizeMB'
        'RemoteMailboxSizeGB'
        'M365LicenseType'
        'M365LicenseTypeByPersona'
        'M365LicenseTargetPersona'
        'M365LicenseTargetPersonaRelaxed'
        'M365LicenseTargetPersona2'
        'M365LicenseTargetPersona3'
        'IsF3UserWithLargeMailbox'
        'F3E3MailboxSizeBucket'
        'F3E3MailboxSizeBucketSortOrder'
        'HasRemainingF3MailboxSizeNonCompliance'
        'ActivityCategory'
        'IsJobTitleValid'
        'JobFamily'
        'JobFamilyRelaxed'
        'JobFamilyFromKeywords'
        'JobFamilyFromKeywordsRelaxed'
        'JobTitle'
        'JobTitleFromDescription'
        'JobTitleFromDescriptionRelaxed'
        'JobTitleNormalized'
        'JobTitleNormalizedRelaxed'
        'JobTitlePadded'
        'JobTitlePaddedRelaxed'
        'EntityLabel'
        'EntraLastSignInDateTime'
        'EntraLastSignInCategory'
        'EntraLastNonInteractiveSignInDateTime'
        'ResolvedLastActivityDateTime'
        'OnPremMailboxPermissionCountNotInM365'
        'IsE3PersonaMissingE3License'
        'F3OversizedMailboxBucket'
        'F3OversizedMailboxBucketSortOrder'
        'IsUnlicensedEnabledF3MailboxBlockingMigration'
        'MailboxMigrationIssueSeverity'
        'MigrationIssueScore'
        'HasNoEntraLastSignIn'
        'HasNoEntraNonInteractiveLastSignIn'
        'HasNoExchangeLastLogon'
        'PrimarySmtpAddress'
        'PrimarySmtpAddressNormalized'
        'PrimarySmtpDomain'
        'HasLegacyRoutingDomainProxyAddress'
        'ExistsInM365Users'
        'MailboxRecipientType'
        'HasRemoteMailbox'
        'HasMailbox'
        'MailboxSource'
        'MailboxSourceHealthStatus'
        'MailboxItemCount'
        'HasMailboxItemCountOver999999'
        'MailboxSizeMB'
        'MailboxSizeGB'
        'MailboxSizeTB'
        'ArchiveMailboxSizeMB'
        'ArchiveMailboxSizeGB'
        'MailboxAndArchiveSizeMB'
        'MailboxAndArchiveSizeGB'
        'IsMailboxLargerThan2GB'
        'IsMailboxLargerThan100GB'
        'IsMailboxSmallerThan2GB'
        'ExchangeLastLogonDateTime'
        'ExchangeLastLogonDate'
        'ResolvedLastLogonDateTime'
        'ResolvedLastLogonStatus'
        'HasNoResolvedLastLogon'
        'ExchangeLastLogonCategory'
        'AdLastLogonCategory'
        'IsExchangeLastLogonOlderThan90Days'
        'IsExchangeLastLogonOlderThan180Days'
        'M365LicenseName'
        'M365LicenseAssignmentGroup'
        'ExchangeServicePlanName'
        'M365LicenseAssignmentGroupCount'
        'M365LicenseAssignmentGroups'
        'M365LicenseAssigningGroups'
        'HasDirectM365License'
        'HasGroupOnlyM365License'
        'HasDirectAndGroupM365License'
        'HasDuplicateM365LicenseAssignment'
        'IsPrimarySmtpDomainAcceptedInExchangeOnline'
        'IsExchangeOnlineDomainNotAllowed'
        'ExchangeOnlineDomainType'
        'IsExchangeOnlineOnMicrosoftDomain'
        'IsExchangeOnlineDomainNonCompliant'
        'AreExchangeOnlineUserDomainsNotAllowed'
        'AreExchangeOnlineUserDomainsNonCompliant'
        'IsUserPrincipalNameDomainVerifiedInEntra'
        'DoesUpnDomainMatchPrimarySmtpDomain'
        'RemoteMailboxRoutingAddress'
        'RemoteMailboxCreatedDateTime'
        'DaysSinceRemoteMailboxCreated'
        'OnPremMailboxCreatedDateTime'
        'ExchangeOnlineMailboxCreatedDateTime'
        'ExchangeOnlineMailboxCreatedDate'
        'FullAccessOnMailboxes'
        'SendAsOnMailboxes'
        'SendOnBehalfOnMailboxes'
        'MailboxDelegatedPermissionSummary'
        'MailboxDelegatedPermissionCount'
        'SendAsMailboxCount'
        'HasFullAccess'
        'HasSendAs'
        'HasSendOnBehalf'
        'HasAnyDelegatedPermission'
        'HasCrossPremisesPermissions'
        'ExchangeMigrationUserStatus'
        'ExchangeMigrationBatchName'
        'ExchangeMigrationCompleteAfterDateTime'
        'ExchangeMigrationCompleteAfterUtcDateTime'
        'IsProtectedByMicrosoft365Backup'
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
    $currentUtcDate = [datetime]::UtcNow.Date
    $exchangeLastLogon90DayCutoff = $currentUtcDate.AddDays(-90)
    $exchangeLastLogon180DayCutoff = $currentUtcDate.AddDays(-180)

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
        $hasRemoteMailbox = @('RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox') -contains $recipientType
        $hasMailbox = $recipientType -ne 'NoMailboxes'
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
        $noLastLogonInOnComputer = $true
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
        $exchangeLastLogonDateTime = Get-SmartM365LatestDateText -Values @(
            (Get-Value $localMailbox @('LastLogonTime')),
            (Get-Value $exoMailbox @('LastLogonTime','LastUserActionTime')),
            (Get-Value $exoStat @('LastLogonTime','LastUserActionTime'))
        )
        $exchangeLastLogonUtc = if ([string]::IsNullOrWhiteSpace($exchangeLastLogonDateTime)) { $null } else { [datetime]::ParseExact($exchangeLastLogonDateTime, 'yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) }
        $remoteMailboxCreatedRaw = Get-Value $remoteMailbox @('WhenMailboxCreated')
        if ([string]::IsNullOrWhiteSpace($remoteMailboxCreatedRaw)) { $remoteMailboxCreatedRaw = Get-Value $remoteMailbox @('WhenCreated') }
        $remoteMailboxCreatedDateTime = Get-SmartM365LatestDateText -Values @($remoteMailboxCreatedRaw)
        $remoteMailboxCreatedUtc = if ([string]::IsNullOrWhiteSpace($remoteMailboxCreatedDateTime)) { $null } else { [datetime]::ParseExact($remoteMailboxCreatedDateTime, 'yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) }
        $onPremMailboxCreatedRaw = Get-Value $localMailbox @('WhenMailboxCreated')
        if ([string]::IsNullOrWhiteSpace($onPremMailboxCreatedRaw)) { $onPremMailboxCreatedRaw = Get-Value $localMailbox @('WhenCreated') }
        $onPremMailboxCreatedDateTime = Get-SmartM365LatestDateText -Values @($onPremMailboxCreatedRaw)
        $exchangeOnlineMailboxCreatedDateTime = Get-SmartM365LatestDateText -Values @((Get-Value $exoMailbox @('WhenMailboxCreated')))
        $out['ExchangeLastLogonDateTime'] = $exchangeLastLogonDateTime
        $out['ExchangeLastLogonDate'] = if ($null -ne $exchangeLastLogonUtc) { $exchangeLastLogonUtc.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { '' }
        $out['ExchangeLastLogonCategory'] = if ($null -eq $exchangeLastLogonUtc) { 'Unknown / Never logged in' } elseif ($exchangeLastLogonUtc.Date -lt $exchangeLastLogon90DayCutoff) { 'Older than 90 days' } else { 'Within last 90 days' }
        $out['IsExchangeLastLogonOlderThan90Days'] = if ($null -ne $exchangeLastLogonUtc) { [string]($exchangeLastLogonUtc.Date -lt $exchangeLastLogon90DayCutoff) } else { '' }
        $out['IsExchangeLastLogonOlderThan180Days'] = if ($null -ne $exchangeLastLogonUtc) { [string]($exchangeLastLogonUtc.Date -lt $exchangeLastLogon180DayCutoff) } else { '' }
        $out['RemoteMailboxCreatedDateTime'] = $remoteMailboxCreatedDateTime
        $out['DaysSinceRemoteMailboxCreated'] = if ($null -ne $remoteMailboxCreatedUtc) { ($currentUtcDate - $remoteMailboxCreatedUtc.Date).Days.ToString([System.Globalization.CultureInfo]::InvariantCulture) } else { '' }
        $out['OnPremMailboxCreatedDateTime'] = $onPremMailboxCreatedDateTime
        $out['ExchangeOnlineMailboxCreatedDateTime'] = $exchangeOnlineMailboxCreatedDateTime
        $out['ExchangeOnlineMailboxCreatedDate'] = if ([string]::IsNullOrWhiteSpace($exchangeOnlineMailboxCreatedDateTime)) { '' } else { $exchangeOnlineMailboxCreatedDateTime.Substring(0, 10) }
        $lastLogonMergeAdExchEntra = Get-SmartM365LatestDateText -Values @([string]$user.LastLogonDate, $out['ExchangeLastLogonDateTime'], $lastLogonEntra, $lastLogonEntraNonInteractive)
        $out['AdOrganizationalUnitPath'] = Get-OuPathFromDistinguishedName -DistinguishedName $distinguishedName
        $out['AdOrganizationalUnitCategory'] = if ($distinguishedNameUpper.Contains('OU=DISABLED_OBJECTS')) { 'DISABLED OBJECTS' } elseif ($distinguishedNameUpper.Contains('OU=ORGANIZATION')) { 'ORGANIZATION' } elseif ($distinguishedNameUpper.Contains('OU=ADMIN')) { 'ADMIN' } elseif ($distinguishedNameUpper.Contains('OU=FAX')) { 'FAX' } else { 'OTHER' }
        $out['EntitySiteType'] = $typeEtablissement
        $out['EntityIntegrationType'] = $typeEntityValue
        $out['GivenNameNormalized'] = $givenNameClean
        $out['SurnameNormalized'] = $surnameClean
        $out['AccountType'] = $accountType
        $out['ExchangeOnlineMappingStatus'] = if ($null -eq $exoMailbox) { 'Not EXO' } elseif (-not [string]::IsNullOrWhiteSpace($aadObjectId)) { 'OK_GUID' } else { 'Missing_AAD_Object' }
        $out['SharedMailboxStatus'] = if ($null -eq $localMailbox) { 'NoMailboxes' } else { [string]((Get-Value $localMailbox @('IsShared')) -match '^(?i:true|1)$') }
        $out['IsLargeSharedMailbox'] = [string](($totalSizeMb -gt 40960) -and $recipientType -eq 'SharedMailbox')
        $out['IsInLocalExchangeOnlineDuplicateSet'] = [string]($mailboxSource -match 'Local' -and $mailboxSource -match 'EXO')
        $out['IsSharedMailbox'] = [string]($recipientType -match 'SharedMailbox')
        $out['OnPremMailboxSizeMB'] = Format-NumberInvariant $localSizeMb
        $out['RemoteMailboxSizeGB'] = if ($null -ne $remoteMailbox -and $null -ne $exoSizeGb) { Format-NumberInvariant $exoSizeGb } else { '' }
        $out['M365LicenseTargetPersona'] = $target1Persona
        $out['M365LicenseTargetPersonaRelaxed'] = $target1Persona2
        $out['M365LicenseTargetPersona2'] = $target2Persona
        $out['M365LicenseTargetPersona3'] = $target3Persona
        $out['M365LicenseTypeByPersona'] = $target1Persona
        $out['M365LicenseType'] = $m365LicenseType
        $out['IsF3UserWithLargeMailbox'] = if ((($licenseGroup -eq 'Microsoft 365 F3') -or ($m365LicenseType -eq 'M365F3')) -and $totalSizeMb -ge 2000) { 'Yes' } else { 'No' }
        $out['F3E3MailboxSizeBucket'] = if ($hasMailbox -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 2000 -and $totalSizeMb -le 2500) { '2-2.5 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 2500 -and $totalSizeMb -le 5000) { '2.5-5 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 5000 -and $totalSizeMb -le 10000) { '5-10 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $licenseGroup -eq 'Microsoft 365 E3' -and $totalSizeMb -gt 10000) { '>10 GB' } else { '' }
        $out['F3E3MailboxSizeBucketSortOrder'] = switch ($out['F3E3MailboxSizeBucket']) { '2-2.5 GB' { '1' } '2.5-5 GB' { '2' } '5-10 GB' { '3' } '>10 GB' { '4' } default { '99' } }
        $out['HasRemainingF3MailboxSizeNonCompliance'] = if ($hasMailbox -and $target1Persona -eq 'M365F3' -and $recipientType -eq 'UserMailbox' -and -not $hasRemoteMailbox -and $totalSizeMb -gt 2000 -and $licenseGroup -ne 'Microsoft 365 E3') { '1' } else { '0' }
        $out['F3OversizedMailboxBucket'] = if ($hasMailbox -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 2000 -and $totalSizeMb -le 2500) { '2-2.5 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 2500 -and $totalSizeMb -le 5000) { '2.5-5 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 5000 -and $totalSizeMb -le 10000) { '5-10 GB' } elseif ($hasMailbox -and $target1Persona -eq 'M365F3' -and $totalSizeMb -gt 10000) { '>10 GB' } else { '' }
        $out['F3OversizedMailboxBucketSortOrder'] = switch ($out['F3OversizedMailboxBucket']) { '2-2.5 GB' { '1' } '2.5-5 GB' { '2' } '5-10 GB' { '3' } '>10 GB' { '4' } default { '99' } }
        $out['IsUnlicensedEnabledF3MailboxBlockingMigration'] = if ($hasMailbox -and $isEnabled -and $recipientType -eq 'UserMailbox' -and -not $hasRemoteMailbox -and $target1Persona -eq 'M365F3' -and @('No License (Unlicensed)','Not in M365') -contains $licenseGroup) { '1' } else { '0' }
        $out['IsE3PersonaMissingE3License'] = if ($hasMailbox -and $target1Persona -eq 'M365E3' -and @('UserMailbox','RemoteUserMailbox') -contains $recipientType -and -not $hasRemoteMailbox -and @('Microsoft 365 F3','No License (Unlicensed)') -contains $licenseGroup) { '1' } else { '0' }
        $out['EntraLastSignInDateTime'] = $lastLogonEntra
        $out['EntraLastNonInteractiveSignInDateTime'] = $lastLogonEntraNonInteractive
        $out['EntraLastSignInCategory'] = Get-DateCategory -Value $lastLogonEntra -LimitDays 90
        $out['ResolvedLastActivityDateTime'] = $lastLogonMergeAdExchEntra
        $out['HasNoEntraLastSignIn'] = [string][string]::IsNullOrWhiteSpace($lastLogonEntra)
        $out['HasNoEntraNonInteractiveLastSignIn'] = [string][string]::IsNullOrWhiteSpace($lastLogonEntraNonInteractive)
        $out['HasNoExchangeLastLogon'] = [string][string]::IsNullOrWhiteSpace([string]$out['ExchangeLastLogonDateTime'])
        $out['OnPremMailboxPermissionCountNotInM365'] = '0'
        $out['ActivityCategory'] = Get-DateCategory -Value $lastLogonMergeAdExchEntra -LimitDays 90
        $out['IsJobTitleValid'] = [string](-not [string]::IsNullOrWhiteSpace($jobTitleNormalized))
        $out['JobFamily'] = $jobFamilyCombined
        $out['JobFamilyRelaxed'] = $jobFamilyCombined2
        $out['JobFamilyFromKeywords'] = $jobFamilyFromKeywords
        $out['JobFamilyFromKeywordsRelaxed'] = $jobFamilyFromKeywords2
        $out['JobTitle'] = $jobTitleFromDesc
        $out['JobTitleFromDescription'] = $jobTitleFromDesc
        $out['JobTitleFromDescriptionRelaxed'] = $jobTitleFromDesc2
        $out['JobTitleNormalized'] = $jobTitleNormalized
        $out['JobTitleNormalizedRelaxed'] = $jobTitleNormalized2
        $out['JobTitlePadded'] = $jobTitlePadded
        $out['JobTitlePaddedRelaxed'] = $jobTitlePadded2
        $out['EntityLabel'] = $labelFromEntity
        if ($hybridIdentityIssues.Count -eq 0 -or [string]::IsNullOrWhiteSpace($primarySmtp)) {
            $out['MailboxMigrationIssueSeverity'] = ''
            $out['MigrationIssueScore'] = ''
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
                $out['MailboxMigrationIssueSeverity'] = 'None'
            }
            else {
                switch (($severityCodes | Measure-Object -Minimum).Minimum) {
                    1 { $out['MailboxMigrationIssueSeverity'] = 'Critical' }
                    2 { $out['MailboxMigrationIssueSeverity'] = 'High' }
                    3 { $out['MailboxMigrationIssueSeverity'] = 'Medium' }
                    default { $out['MailboxMigrationIssueSeverity'] = 'Low' }
                }
            }
            $out['MigrationIssueScore'] = [string]$issueScore
        }
        $out['IsLikelyServiceAccount'] = [string]([string]::IsNullOrWhiteSpace($userPrincipalName) -or $distinguishedName.Contains('Service_Accounts') -or $samAccountName.Contains('s_') -or $samAccountNameLower.Contains('svc') -or $samAccountName.Contains('HealthMailbox'))
        $out['IsLikelyPrivilegedOrServiceAccount'] = [string]($samAccountNameLower.StartsWith('a_') -or $samAccountNameLower.Contains('s_') -or $samAccountNameLower.Contains('administrateur') -or $samAccountNameLower.Contains('administrator'))
        $out['IsInTargetHeadquartersOU'] = [string](($typeEtablissementUpper -eq 'HQ') -or $inTargetHqOu)
        $out['IsOutsideTargetHeadquartersOU'] = [string](($typeEtablissementUpper -ne 'HQ') -and -not $inTargetHqOu -and -not $inDisabledObjectsOu)
        $out['HasNoAdLastLogon'] = [string]([string]::IsNullOrWhiteSpace([string]$user.LastLogonDate))
        $out['PrimarySmtpAddress'] = $primarySmtp
        $out['PrimarySmtpAddressNormalized'] = $primarySmtp
        $out['PrimarySmtpDomain'] = $primaryDomain
        $out['HasLegacyRoutingDomainProxyAddress'] = if (([string]$user.ProxyAddresses).ToLowerInvariant().Contains('@' + $RemoteRoutingDomain)) { 'Yes' } else { 'No' }
        $out['ExistsInM365Users'] = if ($m365Users.Count -gt 0) { [string]($null -ne $m365User) } else { '' }
        $out['MailboxRecipientType'] = $recipientType
        $out['HasRemoteMailbox'] = [string]$hasRemoteMailbox
        $out['HasMailbox'] = [string]$hasMailbox
        $out['MailboxSource'] = $mailboxSource
        $out['MailboxSourceHealthStatus'] = if ($mailboxSource -eq 'None') { 'No mailbox' } elseif ($mailboxSource -eq 'Local+Remote+EXO') { 'Hybrid duplicate' } else { 'OK' }
        $out['MailboxItemCount'] = $itemCount.ToString('0.########',[System.Globalization.CultureInfo]::InvariantCulture)
        $out['HasMailboxItemCountOver999999'] = [string]($itemCount -gt 999999)
        $out['MailboxSizeMB'] = Format-NumberInvariant $totalSizeMb
        $out['MailboxSizeGB'] = if ($null -ne $totalSizeMb) { Format-NumberInvariant ($totalSizeMb / 1024) } else { '' }
        $out['MailboxSizeTB'] = if ($null -ne $totalSizeMb) { Format-NumberInvariant ($totalSizeMb / 1048576) } else { '' }
        $out['ArchiveMailboxSizeMB'] = Format-NumberInvariant $archiveMb
        $out['ArchiveMailboxSizeGB'] = if ($null -ne $archiveMb) { Format-NumberInvariant ($archiveMb / 1024) } else { '' }
        $out['MailboxAndArchiveSizeMB'] = Format-NumberInvariant $totalAndArchiveMb
        $out['MailboxAndArchiveSizeGB'] = if ($null -ne $totalAndArchiveMb) { Format-NumberInvariant ($totalAndArchiveMb / 1024) } else { '' }
        $out['IsMailboxLargerThan2GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -gt 2048) } else { 'False' }
        $out['IsMailboxLargerThan100GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -gt 102400) } else { 'False' }
        $out['IsMailboxSmallerThan2GB'] = if ($null -ne $totalSizeMb) { [string]($totalSizeMb -lt 2048) } else { 'False' }
        $out['M365LicenseName'] = $licenseFromM365
        $out['M365LicenseAssignmentGroup'] = $licenseGroup
        $out['ExchangeServicePlanName'] = $exchangePlan
        $out['M365LicenseAssignmentGroupCount'] = [string]$licenseGroups.Count
        $out['M365LicenseAssignmentGroups'] = $licenseGroups -join ' + '
        $out['M365LicenseAssigningGroups'] = Join-UniqueText -Values @($licenseRows | ForEach-Object { Get-Value $_ @('GroupsAssigningSku') }) -Separator ','
        $out['HasDirectM365License'] = if ([string]::IsNullOrWhiteSpace($licenseKey)) { '' } else { [string]$hasDirectLicense }
        $out['HasGroupOnlyM365License'] = if ([string]::IsNullOrWhiteSpace($licenseKey)) { '' } else { [string]($hasGroupLicense -and -not $hasDirectLicense) }
        $out['HasDirectAndGroupM365License'] = [string](@($licenseRows | Where-Object { (Get-Value $_ @('HasDirectAndGroup')) -match '^(?i:true|1)$' }).Count -gt 0)
        $out['HasDuplicateM365LicenseAssignment'] = if ($out['HasDirectAndGroupM365License'] -eq 'True') { 'Yes' } else { 'No' }
        $out['IsPrimarySmtpDomainAcceptedInExchangeOnline'] = if ($acceptedDomains.Count -gt 0) { [string]($null -ne $acceptedDomain) } else { '' }
        $out['IsExchangeOnlineDomainNotAllowed'] = if ($out['IsPrimarySmtpDomainAcceptedInExchangeOnline'] -eq '') { '' } else { [string]($out['IsPrimarySmtpDomainAcceptedInExchangeOnline'] -eq 'False') }
        $out['ExchangeOnlineDomainType'] = Get-Value $acceptedDomain @('DomainType')
        $out['IsExchangeOnlineOnMicrosoftDomain'] = if ([string]::IsNullOrWhiteSpace($primaryDomain)) { '' } else { [string]($primaryDomain -like '*.onmicrosoft.com') }
        $out['IsExchangeOnlineDomainNonCompliant'] = if ($out['IsPrimarySmtpDomainAcceptedInExchangeOnline'] -eq '') { '' } elseif ($out['IsPrimarySmtpDomainAcceptedInExchangeOnline'] -eq 'False' -or $out['IsExchangeOnlineOnMicrosoftDomain'] -eq 'True') { 'True' } else { 'False' }
        $out['AreExchangeOnlineUserDomainsNotAllowed'] = $out['IsExchangeOnlineDomainNotAllowed']
        $out['AreExchangeOnlineUserDomainsNonCompliant'] = $out['IsExchangeOnlineDomainNonCompliant']
        $out['IsUserPrincipalNameDomainVerifiedInEntra'] = if ([string]::IsNullOrWhiteSpace($upnDomain) -or $verifiedDomains.Count -eq 0) { '' } else { [string]$verifiedDomainSet.Contains((Get-Key $upnDomain)) }
        $out['DoesUpnDomainMatchPrimarySmtpDomain'] = if ([string]::IsNullOrWhiteSpace($upnDomain) -or [string]::IsNullOrWhiteSpace($primaryDomain)) { '' } else { [string]($upnDomain -eq $primaryDomain) }
        $out['RemoteMailboxRoutingAddress'] = Get-Value $remoteMailbox @('RemoteRoutingAddress')
        $out['FullAccessOnMailboxes'] = $fullAccess
        $out['SendAsOnMailboxes'] = $sendAs
        $out['SendOnBehalfOnMailboxes'] = $sendOnBehalf
        $out['MailboxDelegatedPermissionSummary'] = $permissionsSummary
        $out['MailboxDelegatedPermissionCount'] = if ([string]::IsNullOrWhiteSpace($permissionsSummary)) { '0' } else { [string](([string]$permissionsSummary).Split(';').Count) }
        $out['SendAsMailboxCount'] = if ([string]::IsNullOrWhiteSpace($sendAs)) { '0' } else { [string](([string]$sendAs).Split(';').Count) }
        $out['HasFullAccess'] = [string](-not [string]::IsNullOrWhiteSpace($fullAccess))
        $out['HasSendAs'] = [string](-not [string]::IsNullOrWhiteSpace($sendAs))
        $out['HasSendOnBehalf'] = [string](-not [string]::IsNullOrWhiteSpace($sendOnBehalf))
        $out['HasAnyDelegatedPermission'] = [string]((-not [string]::IsNullOrWhiteSpace($fullAccess)) -or (-not [string]::IsNullOrWhiteSpace($sendAs)) -or (-not [string]::IsNullOrWhiteSpace($sendOnBehalf)))
        $out['HasCrossPremisesPermissions'] = Get-Value $perms @('HasCrossPremisesPermissions')
        $out['ExchangeMigrationUserStatus'] = if ($recipientType -like '*Remote*') { 'Completed' } else { Get-Value $migrationRow @('UserStatus') }
        $out['ExchangeMigrationBatchName'] = Get-Value $migrationRow @('BatchName')
        $out['ExchangeMigrationCompleteAfterDateTime'] = Get-Value $migrationRow @('CompleteAfterUser','CompleteAfter')
        $out['ExchangeMigrationCompleteAfterUtcDateTime'] = Get-Value $migrationRow @('CompleteAfterUTCUser','CompleteAfterUTC')
        $out['IsProtectedByMicrosoft365Backup'] = if ($backupCoverage.Count -gt 0) { [string]($null -ne $backupRow) } else { '' }
        $out['AdLastLogonCategory'] = Get-DateCategory -Value ([string]$user.LastLogonDate) -LimitDays 90
        $out['ResolvedLastLogonStatus'] = if ($hasLastLogon) { 'Known' } else { 'Unknown' }
        $out['HasNoResolvedLastLogon'] = [string](-not $hasLastLogon)
        $out['ResolvedLastLogonDateTime'] = if ($hasLastLogon) { [string]$user.LastLogonDate } else { '' }
        $out['IsEnabledWithNoResolvedLastLogon'] = [string]($isEnabled -and [string]::IsNullOrWhiteSpace([string]$out['ResolvedLastLogonDateTime']))

        [void]$enrichedRows.Add([pscustomobject]$out)
    }

    $enrichedCsv = Join-Path -Path $OutputFolder -ChildPath 'AD_Users_AllDomains.csv'
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBSNcEhMw2coE/E
# 9iEh/skFfbNSfDchEt4TnadqaCIo2KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJpg+ImHwjA5DDO6IkJrWxc922NAdbjDVKIMPz/H98E5MA0GCSqG
# SIb3DQEBAQUABIIBgBr8JgG3teNEOo9iBsC5wE/zNGk9JUg7NrqXkXxxgmy480Ou
# E9I8tR13HY1mAf20t7qMNgJuOP6wVsj95tUijp9LGm9Hgr2jHfTcsAUzT6E4AUCJ
# DA2D2SCCuWXcpJzRCwXGFImByX5aiOSi3gUkIODtohWnqfM96Yd3rutVUnOe6PT5
# JfPxYwrQ9g2+lMbdzpu5YDxyyfJeYCXfK7VEnOD9Os/UR0Nezjn0IqKlYEaKipIV
# +AMH1jkcf3fiYnctE36dQLqllRWxRW/NAQWa7cJy+Z4gHuCy+8Hbz3OT+mkeZtww
# XHbGIFYGdzH+4FACCI5w8z2p9mZWcLPbujKp/vC9vyOo5baa0x2XlL3r0xeABBD0
# RrhzKu0BOr7OW1O//yx9dxKOTox0DrIRjEvLdFxafVVA6giVf5zJRXUX9W4Ju4+e
# eTiStLCLNE5kPoq/50fC9BCJ45+w2iG0W+niI22AqMHW5QmKUT3EmY5moQN7piBN
# p8dN8MzsYZouNd5dUaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTcxOTQw
# NDNaMC8GCSqGSIb3DQEJBDEiBCAb7qC6M7n1xKSJa7TIlpK54ZjGj7omLdj3KU2S
# QURfPzANBgkqhkiG9w0BAQEFAASCAgA9fd2rw2L38Z1DpeJUBl4gVYCso9y5XuUB
# c/9LzXY0ZktrDkuQ/577s3NZnMoN+00yMteqkk189Y/ezhtymqKB1U2w8AP72k7N
# aGrhBhCSltocd1hhnuWHxaxLTjJAQd0f+ZgGRGjWNcn1bk2WINh7zfKCSWWck52S
# +X0wE9gE2hBYSpN199uZpVpjDZQV4ea2ZostNzTXm1A5iJNqofSIIyp3KcXxuMPj
# X2Eb2ydvdcsuK1ZRbUK745o8/iv6fAD0xr+Dg+eUE5mEleA2P/ogChTL7P8jRk2f
# BQahZjRQtKv5KMkItYuUut6Li916qlDJTYkdbQez1JA27gr2soTvBc+2cFBp3Fgp
# uVHBbTb2de7J9Jz5Jm2PzgNI5MXUUT+gQuCZ+yEMpetJXAu9f8+lwZGSay2OIcIg
# gJ79EFWqTuTM4L+wCMQZXgkhsmkWCqNr8NOxh+teVWymWAHiXdfrcaZtuwxxLU/6
# CXhECKOJ+RRl/Ch27bp8z3qX3fMdOtYzNKRBsnuXHlpwxI+/DtA0SnkrWTMcfg0N
# Es8MPskmqsTG8el40xLjvq1bMfMzT5hHjruN8iXo5bnXn9GVP8+OBANdQID410O+
# eAn0U4r5cDg02by0GM8LbVg8w3W8kBr00tjsGg5NNQPHb6WXLVOTypijJPJzAW+G
# lxkXJupIiQ==
# SIG # End signature block
