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

    $intuneRows = @(LoadCsv 'Intune_Devices_Inventory.csv')
    $entraRows = @(LoadCsv 'M365_Entra_Devices.csv')
    $hardwareConflictRows = @(LoadCsv 'M365_Entra_Devices_HardwareIdConflicts.csv')
    $localSystemRows = @(LoadCsv 'Intune_Devices_LocalSystem.csv')
    $windowsUpdateRows = @(LoadCsv 'Intune_WindowsUpdate_Status.csv')
    $adUserRows = @(LoadCsv 'AD_Users_AllDomains.csv')
    $licenseRows = @(LoadCsv 'M365_Licenses_Users.csv')

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

        $out['ObjectGUID_Norm'] = $objectGuid
        $out['FriendlyOSVersionName'] = $friendly
        $out['FriendlyOSVersionName_Snapshot'] = $friendly
        $out['ShortOSName'] = if ($isWin11) { 'Windows 11' } elseif ($isWin10) { 'Windows 10' } else { V $row @('OperatingSystem') }
        $out['IsWindows11'] = [string]$isWin11
        $out['IsWindows11-Bool'] = [string]$isWin11
        $out['IsWindows1124H2-Bool'] = [string]$is24
        $out['IsWindows1125H2-Bool'] = [string]$is25
        $out['NeedToBeUpgrade'] = [string]($isWin10 -and $upgradeLabel -ne 'Not Capable')
        $out['W11Eligibilty'] = if ($isWin11) { 'Upgraded' } else { $upgradeLabel }
        $out['WIN11_GLOBALRESULT'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_INTUNE_ONLY'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_M365_ONLY'] = $out['W11Eligibilty']
        $out['WIN11_GLOBALRESULT_SOURCE'] = if ($null -ne $wuAny) { 'Intune Windows Update status' } elseif ($null -ne $intune) { 'Intune devices inventory' } else { '' }
        $out['UpgradeEligibility_From_M365'] = $upgradeRaw
        $out['UpgradeEligibility_Label_From_M365'] = $upgradeLabel
        $out['OS_version_From_M365'] = $osVersionM365
        $out['OperatingSystem_Snapshot_Priority'] = V $row @('OperatingSystem')
        $out['OperatingSystemVersion_Snapshot_Priority'] = if ($osVersionM365) { $osVersionM365 } else { V $row @('operatingSystemVersion') }

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
        $out['Last_Logged_User_FromM365'] = $primaryUser
        $out['PrimarySMTPaddressUser'] = $primaryUser
        $out['PrimarySMTPaddress_From_AD_Or_M365'] = if ($primaryUserRow) { V $primaryUserRow @('EmailAddress','UserPrincipalName') } else { $primaryUser }
        $out['LicenseFromM365_Users_From_AD'] = $licenseValue
        $out['LicenseFromM365_Users_From_PrimaryIntuneUser'] = $licenseValue
        $out['OU_Path_Users_From_AD'] = V $primaryUserRow @('CanonicalName','DistinguishedName')
        $out['OrganizationalUnit_User_From_AD'] = $out['OU_Path_Users_From_AD']
        $out['OrganizationalUnit_PrimaryUser_From_AD'] = $out['OU_Path_Users_From_AD']
        $out['Enabled_From_Last_Logged_UserDomain'] = B (V $primaryUserRow @('Enabled'))

        $out['DiskTotalStorage_Go_From_M365'] = $totalGb
        $out['DiskFreeStorage_Go_From_M365'] = $freeGb
        $out['DiskTotalSize'] = $totalGb
        $out['DiskTotalFreeSpace'] = $freeGb
        $out['DiskFreePercent'] = $freePercent
        $out['IsFreeStorageNotEnoughForWin11Update'] = $freeNotEnough
        $out['PhysicalMemoryGB_From_M365'] = V $intune @('PhysicalMemoryGB')
        $out['Memory_GB_Number'] = $out['PhysicalMemoryGB_From_M365']
        $out['Memory'] = $out['PhysicalMemoryGB_From_M365']

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

        $out['LastSignIn_Merged'] = if ($entraSignIn) { $entraSignIn } else { $adLastLogon }
        $mergedDays = Days $out['LastSignIn_Merged']
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
    $enrichedRows | Export-Csv -LiteralPath $enrichedCsv -NoTypeInformation -Encoding UTF8
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDIU5Fz2FXGsESr
# wT134ldMbbhN0omA5q0dx0w7W2PoNaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCB/6+McGSsggtrw92CTjvjZ+nJ1NC+AxxIszg8kFAwnDTANBgkqhkiG9w0B
# AQEFAASCAYCweYKvxq/5jliFtAC/ujwzrqlmeL4qciUgMo7VSBst0znVLlPAV9zU
# A0flSi2ndioFTFMW437pBe08gClajQhXrh4asUbDnzmF1S3QUSMlypggvfO3scwZ
# 7n0hox4MzCPbmgGNaLX51mLk2e4fz89T4spRwtD7D22njyViMydwQ/jiZqgV7Gug
# l0Ksa7a0G6FZ/qzvJ9H4ppbAUU/c68Wa753vpCX7aNHgvjn+5k2AGCxMfvTcL+Wg
# Vcl4MId8iCem1LuIOlUT8yyPOfYr8/aXXE6mQDrU8YsvbmBhRegTuutm0TVA9Jph
# 3dcYTReyIA88tJzCziZyZyu07v8hQxxEm2VVBXtdZil12pWgH3+uav0ZxaCCQrT1
# C+Hi7/kSTsiXmPTVLbCSShXYLlVgS5BKDERRALD6v14N1zud5+C6inScyyK9/Vxd
# 6vA9Q6DBqXLkz2hQQB3xTa45etY3XbPx0hpGI8AAxu/mMNsVzHz1ZKJ1Vo8mLnX0
# hyALc0FhDUE=
# SIG # End signature block
