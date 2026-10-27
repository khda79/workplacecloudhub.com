#Requires -Version 7.0
# Version 1.1.0
# Offline Device System preparation. No collection, refresh or publication.
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'CountryEvidence.psm1') -Force -ErrorAction Stop

function Get-SystemText {
    param($Row,[string]$Name)
    $property=$Row.PSObject.Properties[$Name]
    if ($null -eq $property) { return '' }
    return ([string]$property.Value).Trim()
}
function Get-SystemKey {
    param($Row,[string]$IdColumn)
    $tenant=(Get-SystemText $Row 'TenantKey').ToLowerInvariant()
    $id=(Get-SystemText $Row $IdColumn).ToLowerInvariant()
    if (-not $tenant -or -not $id -or $tenant.Contains('|') -or $id.Contains('|')) {
        throw 'Device System evidence rejected: missing or invalid tenant-aware identity.'
    }
    return "$tenant|$id"
}
function New-SystemIndex {
    param([object[]]$Rows,[string]$IdColumn,[switch]$AllowAmbiguous)
    $index=@{}
    foreach ($row in $Rows) {
        $key=Get-SystemKey $row $IdColumn
        if ($index.ContainsKey($key)) {
            if (-not $AllowAmbiguous) { throw 'Device System evidence rejected: duplicate tenant-aware identity.' }
            $index[$key]=$null
        } else { $index[$key]=$row }
    }
    return $index
}
function Convert-SystemState {
    param([string]$Value,[string[]]$Known,[Collections.Generic.List[string]]$Issues,[string]$Field)
    if (-not $Value) { return 'Missing' }
    foreach ($state in @($Known)+@('NotCollected','Unknown','Error','NotSupported')) {
        if ($Value -ieq $state) { return $state }
    }
    $Issues.Add("Unrecognized $Field")
    return 'Unknown'
}
function Convert-SystemTime {
    param([string]$Value,[string]$OffsetlessTimeZoneId)
    if (-not $Value) { return [pscustomobject]@{Utc=$null;Quality='Missing'} }
    if ($Value -in @('Error','Unknown','NotCollected','NotSupported')) { return [pscustomobject]@{Utc=$null;Quality=$Value} }
    $dto=[datetimeoffset]::MinValue
    if ($Value -match '(Z|[+-]\d\d:\d\d)$') {
        if ([datetimeoffset]::TryParse($Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dto)) {
            return [pscustomobject]@{Utc=$dto.UtcDateTime;Quality='Explicit offset'}
        }
    } else {
        $date=[datetime]::MinValue
        $formats=[string[]]@('dd/MM/yyyy HH:mm:ss','yyyy-MM-dd HH:mm:ss','yyyy-MM-ddTHH:mm:ss','yyyy-MM-ddTHH:mm:ss.FFFFFFF')
        if ([datetime]::TryParseExact($Value,$formats,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)) {
            if (-not $OffsetlessTimeZoneId) { return [pscustomobject]@{Utc=$null;Quality='Time zone unspecified'} }
            $zone=[TimeZoneInfo]::FindSystemTimeZoneById($OffsetlessTimeZoneId)
            if ($zone.IsAmbiguousTime($date) -or $zone.IsInvalidTime($date)) { return [pscustomobject]@{Utc=$null;Quality='Ambiguous local time'} }
            return [pscustomobject]@{Utc=[TimeZoneInfo]::ConvertTimeToUtc($date,$zone);Quality='Configured time zone'}
        }
    }
    return [pscustomobject]@{Utc=$null;Quality='Invalid date'}
}
function Get-DeviceSystemEvidenceColumns {
    @('TenantKey','Device System Key','Device Source ID','Device Name','Azure AD Device ID','Primary User UPN',
      'Country','Country Basis','Inventory Match','Prepared Inventory Match','Observation Raw','Observation UTC',
      'Observation Time Basis','Observation Age Hours','Freshness','Freshness Max Age Hours','Prepared UTC',
      'Secure Boot State','Secure Boot Raw','BIOS Version','BIOS Date','BIOS Date Quality','Firmware Type','Firmware Raw',
      'Last Boot Time','Last Boot Quality','Local Encryption State','Inventory Encryption State','Encryption Comparison',
      'VBS State','VBS Raw','VBS Hardware Requirements','VBS Hardware Raw','Credential Guard State','Credential Guard Raw',
      'Management State','Management Raw','Compliance State','Compliance Raw','Inventory Compliance State','Compliance Comparison',
      'Last Sync Raw','Last Sync UTC','Last Sync Time Basis','Third Party AV State','Third Party AV Raw','Third Party AV Products',
      'Expected AV Name','Expected AV State','Expected AV Raw','AV Qualification','AV Comparison',
      'Windows Update Reboot Required','Windows Update Reboot Raw','BitLocker OS Protection','BitLocker OS Protection Raw',
      'BitLocker OS Conversion','BitLocker OS Conversion Raw','Data Quality Issues',
      'Local Data Source','Local Data Freshness','Last Console Logon Status','Last Console Logon User',
      'Last Console Logon UTC','Last Console Logon Quality','Last Console Logon SID')
}
function Assert-DeviceSystemCsvSchema {
    param([Parameter(Mandatory)][string]$Path)
    $schema=Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\config\source-schema.json') -Raw | ConvertFrom-Json
    $reader=[IO.File]::OpenText([IO.Path]::GetFullPath($Path))
    try {$header=$reader.ReadLine()} finally {$reader.Dispose()}
    if (-not $header) {throw 'Device System source schema rejected: no CSV header.'}
    $names=@((($header+"`n"+$header | ConvertFrom-Csv)[0]).PSObject.Properties.Name)
    if (@($schema.files.'Intune_Devices_LocalSystem.csv'.columns | Where-Object {$_ -notin $names}).Count) {throw 'Device System source schema rejected: required columns are missing.'}
    $recurring=@('LocalObservationUtc','LocalDataSource','LocalDataFreshness')
    if (@($recurring | Where-Object {$_ -in $names}).Count) {
        if (@($recurring | Where-Object {$_ -notin $names}).Count) {throw 'Device System source schema rejected: incomplete recurring metadata.'}
    } elseif ('PlatformScriptLastUpdate' -notin $names) {throw 'Device System source schema rejected: observation timestamp column missing.'}
}
function New-DeviceSystemEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SystemRows,
          [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$InventoryRows,
          [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$UserRows,
          [object[]]$PreparedInventoryRows=@(),
          [Nullable[double]]$MaxObservationAgeHours,
          [string]$OffsetlessGraphTimeZoneId,
          [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow)
    if ($null -ne $MaxObservationAgeHours -and ($MaxObservationAgeHours -le 0 -or [double]::IsNaN($MaxObservationAgeHours) -or [double]::IsInfinity($MaxObservationAgeHours))) { throw 'Device System observation age policy must be a finite positive number.' }
    if ($OffsetlessGraphTimeZoneId) { $null=[TimeZoneInfo]::FindSystemTimeZoneById($OffsetlessGraphTimeZoneId) }
    $systemIndex=New-SystemIndex $SystemRows 'DeviceId'
    $inventoryIndex=New-SystemIndex $InventoryRows 'Device ID'
    $userIndex=New-SystemIndex $UserRows 'User principal name' -AllowAmbiguous
    $tenants=@(@($SystemRows)+@($InventoryRows)+@($UserRows) | ForEach-Object {(Get-SystemText $_ 'TenantKey').ToLowerInvariant()} | Sort-Object -Unique)
    if ($tenants.Count -ne 1) { throw 'Device System evidence rejected: expected one consistent tenant population.' }
    $preparedIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $PreparedInventoryRows) {
        $tenant=Get-SystemText $row 'TenantKey'
        if ($tenant -and $tenant -ine $tenants[0]) { throw 'Device System evidence rejected: prepared inventory tenant mismatch.' }
        $id=Get-SystemText $row 'Device Source ID'
        if (-not $id -or -not $preparedIds.Add($id)) { throw 'Device System evidence rejected: invalid prepared inventory grain.' }
    }
    foreach ($entry in ($systemIndex.GetEnumerator() | Sort-Object Key)) {
        $row=$entry.Value; $issues=[Collections.Generic.List[string]]::new()
        $deviceId=Get-SystemText $row 'DeviceId'
        $inventory=$inventoryIndex[$entry.Key]
        $match=if ($null -ne $inventory) {'Matched'} else {'Unmatched'}
        $country='Unknown'; $countryBasis='No unique primary-user country'
        $upn=if ($null -ne $inventory) {Get-SystemText $inventory 'Primary user UPN'} else {Get-SystemText $row 'UserPrincipalName'}
        $identityConflict=$null -ne $inventory -and (Get-SystemText $inventory 'Identity Status') -eq 'Identity conflict'
        if($identityConflict){$countryBasis='Identity conflict; attribution indeterminate';$issues.Add('Identity conflict; disputed attribution excluded')}
        if ($match -eq 'Matched' -and $upn -and -not $identityConflict) {
            $userKey=$tenants[0]+'|'+$upn.ToLowerInvariant(); $user=$userIndex[$userKey]
            if ($null -ne $user) {
                $primaryCountry=Resolve-DirectoryCountry (Get-SystemText $user 'CountryOrRegion')
                $country=$primaryCountry.Name
                $countryBasis=if($primaryCountry.Code){'Primary user Entra country'}else{'Primary user Entra country unavailable'}
            } elseif ($userIndex.ContainsKey($userKey)) { $issues.Add('Ambiguous primary-user match') }
        }
        $localSource=Get-SystemText $row 'LocalDataSource'
        if (-not $localSource) {$localSource='LegacyPlatformScript'}
        $localFreshness=Get-SystemText $row 'LocalDataFreshness'
        if ($localSource -eq 'IntuneRemediation') {
            $localFreshness=Convert-SystemState $localFreshness @('Fresh','Stale','Pending','NotApplicable','Failed','ScriptError','InvalidOutput','InvalidTimestamp','FutureObservation') $issues 'LocalDataFreshness'
        } elseif (-not $localFreshness) {$localFreshness='Not reported'}
        $observationRaw=if ($localSource -eq 'IntuneRemediation') {Get-SystemText $row 'LocalObservationUtc'} else {Get-SystemText $row 'PlatformScriptLastUpdate'}
        $observation=Convert-SystemTime $observationRaw $OffsetlessGraphTimeZoneId
        if ($localSource -eq 'IntuneRemediation') {
            # Never repair a recurring local observation using a legacy script-state alias.
            if ($observation.Quality -ne 'Explicit offset') {$observation=[pscustomobject]@{Utc=$null;Quality=if($observation.Quality -eq 'Missing'){'Missing'}else{'Unqualified recurring time'}}}
            $legacyAlias=Get-SystemText $row 'PlatformScriptLastUpdate'
            if ($legacyAlias -and $null -ne $observation.Utc) {
                $aliasTime=Convert-SystemTime $legacyAlias ''
                if ($null -eq $aliasTime.Utc -or $aliasTime.Utc -ne $observation.Utc) {
                    $issues.Add('Conflicting observation aliases')
                    $observation=[pscustomobject]@{Utc=$null;Quality='Conflicting observation aliases'}
                }
            }
        } elseif ($localSource -ne 'LegacyPlatformScript') {
            $issues.Add('Unknown local data source')
            $observation=[pscustomobject]@{Utc=$null;Quality='Unknown local data source'}
        }
        $age=$null; $freshness=$observation.Quality
        if ($null -ne $observation.Utc) {
            $age=($AsOfUtc.UtcDateTime-$observation.Utc).TotalHours
            if ($age -lt 0) { $freshness='Future observation' }
            elseif ($null -eq $MaxObservationAgeHours) { $freshness='Policy not configured' }
            elseif ($age -le $MaxObservationAgeHours) { $freshness='Within policy' } else { $freshness='Stale' }
        }
        if ($localSource -eq 'IntuneRemediation' -and $localFreshness -ne 'Fresh') {
            $freshness=$localFreshness
        }
        $consoleRaw=Get-SystemText $row 'LastConsoleLogonUtc'
        $consoleTime=Convert-SystemTime $consoleRaw ''
        $consoleStatus=Get-SystemText $row 'LastConsoleLogonStatus'
        if (-not $consoleStatus) {$consoleStatus='NotCollected'}
        $consoleStatus=Convert-SystemState $consoleStatus @('Observed','NotFoundInScan') $issues 'LastConsoleLogonStatus'
        $consoleUser=Get-SystemText $row 'LastConsoleLogonUser'
        $consoleSid=Get-SystemText $row 'LastConsoleLogonSid'
        if ($consoleSid -and $consoleSid -notmatch '^S-1-[0-9]+(?:-[0-9]+)+$') {$issues.Add('Invalid console SID');$consoleSid=''}
        $consoleQuality=if($consoleStatus -eq 'Observed'){$consoleTime.Quality}else{$consoleStatus}
        if ($consoleStatus -eq 'Observed' -and (-not $consoleUser -or $consoleTime.Quality -ne 'Explicit offset' -or $null -eq $observation.Utc -or $observation.Utc -gt $AsOfUtc.UtcDateTime -or $consoleTime.Utc -gt $observation.Utc -or $consoleTime.Utc -gt $AsOfUtc.UtcDateTime)) {
            $consoleStatus='Unknown'; $consoleQuality='Unqualified console logon observation'
            $issues.Add('Console logon qualification')
        }
        if ($localSource -eq 'IntuneRemediation' -and $localFreshness -ne 'Fresh') {
            $consoleStatus=if($localFreshness -in @('NotCollected','Error')){$localFreshness}else{'Unknown'}
            $consoleQuality='Local collection '+$localFreshness
        }
        if ($consoleStatus -ne 'Observed') {$consoleUser='';$consoleSid='';$consoleTime=[pscustomobject]@{Utc=$null;Quality=$consoleQuality}}
        $syncRaw=Get-SystemText $row 'LastSyncDateTime'; $sync=Convert-SystemTime $syncRaw $OffsetlessGraphTimeZoneId
        $bootValues=@('LastRebootDate','Last Reboot Date','LastBootUpTime','LastBootTime' | ForEach-Object {Get-SystemText $row $_} | Where-Object {$_})
        $bootDistinct=@($bootValues | Sort-Object -Unique)
        $boot=if ($bootDistinct.Count -eq 1) {$bootDistinct[0]} else {''}
        $bootQuality=if ($bootDistinct.Count -eq 0) {'Missing'} elseif ($bootDistinct.Count -gt 1) {'Conflicting aliases'} elseif ($boot -in @('Error','Unknown','NotCollected','NotSupported')) {$boot} else {'Observed local time; time zone unspecified'}
        if ($bootDistinct.Count -eq 1 -and $boot -notin @('Error','Unknown','NotCollected','NotSupported')) {
            $bootParsed=Convert-SystemTime $boot ''
            if ($bootParsed.Quality -eq 'Invalid date') { $bootQuality='Invalid date' }
        }
        $bios=Get-SystemText $row 'BIOSDate'; $biosDate=[datetime]::MinValue
        $biosQuality=if (-not $bios) {'Missing'} elseif ($bios -in @('Error','Unknown','NotCollected','NotSupported')) {$bios} elseif ([datetime]::TryParseExact($bios,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$biosDate)) {'Observed date; not a health verdict'} else {'Invalid date'}
        $states=@{}
        $definitions=@(
            @('SecureBootStatus','Secure Boot State',@('Enabled','Disabled')),
            @('FirmwareType','Firmware Type',@('Uefi','Bios')),
            @('IsEncrypted','Local Encryption State',@('True','False')),
            @('DeviceGuardVBSState','VBS State',@('running','notEnabled','enabledButNotRunning')),
            @('DeviceGuardVBSHardwareRequirement','VBS Hardware Requirements',@('meetHardwareRequirements','doesNotMeetHardwareRequirements')),
            @('DeviceGuardCredentialGuardState','Credential Guard State',@('running','notEnabled','enabledButNotRunning')),
            @('ManagementState','Management State',@('managed','retirePending','retireFailed','wipePending','wipeFailed','unhealthy','deletePending','retireIssued','wipeIssued','wipeCanceled','retireCanceled','discovered')),
            @('ComplianceState','Compliance State',@('compliant','noncompliant','inGracePeriod','configManager')),
            @('ThirdPartyAVStatus','Third Party AV State',@('Detected','NotDetected')),
            @('ExpectedThirdPartyAVStatus','Expected AV State',@('Detected','NotDetected')),
            @('WindowsUpdateRebootRequired','Windows Update Reboot Required',@('True','False')),
            @('BitLockerOSProtectionStatus','BitLocker OS Protection',@('On','Off')),
            @('BitLockerOSConversionStatus','BitLocker OS Conversion',@('FullyEncrypted','FullyDecrypted','EncryptionInProgress','DecryptionInProgress','EncryptionPaused','DecryptionPaused'))
        )
        foreach ($def in $definitions) { $states[$def[1]]=Convert-SystemState (Get-SystemText $row $def[0]) $def[2] $issues $def[0] }
        $inventoryEncryption=if ($null -ne $inventory) {Convert-SystemState (Get-SystemText $inventory 'Encrypted') @('True','False') $issues 'Inventory.Encrypted'} else {'Missing'}
        $encryptionComparison=if ($states['Local Encryption State'] -in @('True','False') -and $inventoryEncryption -in @('True','False')) {if ($states['Local Encryption State'] -eq $inventoryEncryption) {'Agreement'} else {'Conflict'}} else {'Not comparable'}
        if (($states['Local Encryption State'] -eq 'False' -and $states['BitLocker OS Conversion'] -eq 'FullyEncrypted') -or ($states['Local Encryption State'] -eq 'True' -and $states['BitLocker OS Conversion'] -eq 'FullyDecrypted')) { $encryptionComparison='Conflict' }
        $inventoryCompliance=if ($null -ne $inventory) {Get-SystemText $inventory 'Compliance'} else {''}
        $complianceComparison=if ($states['Compliance State'] -in @('compliant','noncompliant','inGracePeriod') -and $inventoryCompliance -in @('compliant','noncompliant','inGracePeriod')) {if ($states['Compliance State'] -ieq $inventoryCompliance) {'Agreement'} else {'Conflict'}} else {'Not comparable'}
        $avName=Get-SystemText $row 'ExpectedThirdPartyAVName'; $products=Get-SystemText $row 'ThirdPartyAVProducts'
        $avQualification=if (-not $avName) {'Expected product not configured'} elseif ($states['Expected AV State'] -in @('Detected','NotDetected')) {'Evaluable registration observation'} else {$states['Expected AV State']}
        $avComparison='No contradiction observed'
        if (($states['Expected AV State'] -eq 'Detected' -and $states['Third Party AV State'] -eq 'NotDetected') -or ($states['Third Party AV State'] -eq 'Detected' -and -not $products) -or ($states['Third Party AV State'] -eq 'NotCollected' -and $products)) { $avComparison='Conflict'; $avQualification='Conflicting registration evidence'; $issues.Add('AV registration conflict') }
        if ($match -ne 'Matched') {$issues.Add('Inventory match missing')}
        if ($PreparedInventoryRows.Count -gt 0 -and -not $preparedIds.Contains($deviceId)) {$issues.Add('Not in prepared inventory')}
        if ($encryptionComparison -eq 'Conflict') {$issues.Add('Encryption conflict; different observation times possible')}
        if ($complianceComparison -eq 'Conflict') {$issues.Add('Compliance conflict; different observation times possible')}
        if ($bootQuality -in @('Conflicting aliases','Invalid date')) {$issues.Add('Last boot quality')}
        if ($biosQuality -eq 'Invalid date') {$issues.Add('BIOS date quality')}
        if ($observation.Quality -in @('Invalid date','Ambiguous local time') -or $freshness -eq 'Future observation') {$issues.Add('Observation date quality')}
        if ($sync.Quality -in @('Invalid date','Ambiguous local time')) {$issues.Add('Sync date quality')}
        [pscustomobject][ordered]@{
            'TenantKey'=$tenants[0]; 'Device System Key'=$entry.Key; 'Device Source ID'=$deviceId.ToLowerInvariant();
            'Device Name'=(Get-SystemText $row 'DeviceName'); 'Azure AD Device ID'=(Get-SystemText $row 'AzureADDeviceId'); 'Primary User UPN'=$upn;
            'Country'=$country; 'Country Basis'=$countryBasis; 'Inventory Match'=$match;
            'Prepared Inventory Match'=if ($PreparedInventoryRows.Count -eq 0) {'Not checked'} elseif ($preparedIds.Contains($deviceId)) {'Matched'} else {'Unmatched'};
            'Observation Raw'=$observationRaw; 'Observation UTC'=if ($null -ne $observation.Utc) {$observation.Utc.ToString('o')} else {''};
            'Observation Time Basis'=$observation.Quality; 'Observation Age Hours'=if ($null -ne $age) {$age.ToString('0.000000',[Globalization.CultureInfo]::InvariantCulture)} else {''};
            'Freshness'=$freshness; 'Freshness Max Age Hours'=if ($null -ne $MaxObservationAgeHours) {$MaxObservationAgeHours.ToString([Globalization.CultureInfo]::InvariantCulture)} else {''};
            'Prepared UTC'=$AsOfUtc.UtcDateTime.ToString('o');
            'Secure Boot State'=$states['Secure Boot State']; 'Secure Boot Raw'=(Get-SystemText $row 'SecureBootStatus');
            'BIOS Version'=(Get-SystemText $row 'BIOSVersion'); 'BIOS Date'=$bios; 'BIOS Date Quality'=$biosQuality;
            'Firmware Type'=$states['Firmware Type']; 'Firmware Raw'=(Get-SystemText $row 'FirmwareType');
            'Last Boot Time'=$boot; 'Last Boot Quality'=$bootQuality;
            'Local Encryption State'=$states['Local Encryption State']; 'Inventory Encryption State'=$inventoryEncryption; 'Encryption Comparison'=$encryptionComparison;
            'VBS State'=$states['VBS State']; 'VBS Raw'=(Get-SystemText $row 'DeviceGuardVBSState');
            'VBS Hardware Requirements'=$states['VBS Hardware Requirements']; 'VBS Hardware Raw'=(Get-SystemText $row 'DeviceGuardVBSHardwareRequirement');
            'Credential Guard State'=$states['Credential Guard State']; 'Credential Guard Raw'=(Get-SystemText $row 'DeviceGuardCredentialGuardState');
            'Management State'=$states['Management State']; 'Management Raw'=(Get-SystemText $row 'ManagementState');
            'Compliance State'=$states['Compliance State']; 'Compliance Raw'=(Get-SystemText $row 'ComplianceState');
            'Inventory Compliance State'=$inventoryCompliance; 'Compliance Comparison'=$complianceComparison;
            'Last Sync Raw'=$syncRaw; 'Last Sync UTC'=if ($null -ne $sync.Utc) {$sync.Utc.ToString('o')} else {''}; 'Last Sync Time Basis'=$sync.Quality;
            'Third Party AV State'=$states['Third Party AV State']; 'Third Party AV Raw'=(Get-SystemText $row 'ThirdPartyAVStatus'); 'Third Party AV Products'=$products;
            'Expected AV Name'=$avName; 'Expected AV State'=$states['Expected AV State']; 'Expected AV Raw'=(Get-SystemText $row 'ExpectedThirdPartyAVStatus');
            'AV Qualification'=$avQualification; 'AV Comparison'=$avComparison;
            'Windows Update Reboot Required'=$states['Windows Update Reboot Required']; 'Windows Update Reboot Raw'=(Get-SystemText $row 'WindowsUpdateRebootRequired');
            'BitLocker OS Protection'=$states['BitLocker OS Protection']; 'BitLocker OS Protection Raw'=(Get-SystemText $row 'BitLockerOSProtectionStatus');
            'BitLocker OS Conversion'=$states['BitLocker OS Conversion']; 'BitLocker OS Conversion Raw'=(Get-SystemText $row 'BitLockerOSConversionStatus');
            'Data Quality Issues'=($issues -join '; ');
            'Local Data Source'=$localSource; 'Local Data Freshness'=$localFreshness;
            'Last Console Logon Status'=$consoleStatus; 'Last Console Logon User'=$consoleUser;
            'Last Console Logon UTC'=if($consoleStatus -eq 'Observed'){$consoleTime.Utc.ToString('o')}else{''};
            'Last Console Logon Quality'=$consoleQuality; 'Last Console Logon SID'=$consoleSid
        }
    }
}
Export-ModuleMember -Function New-DeviceSystemEvidence,Get-DeviceSystemEvidenceColumns,Assert-DeviceSystemCsvSchema

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBe8KyDI+dIQEan
# bWFAWRAaU4EaCdFAvEHtByzyga3F16CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOQ8Q9hD8vtqJAPFmdhUa1+auwijIPsp7UUDR9dVQqm5MA0GCSqG
# SIb3DQEBAQUABIIBgIbrF+F6wlq1ONJPu0mNmPHpTVQ3pSvjY8qCEmVZkadTnz2B
# HbIxjYVPspOL8Ry8dZa0SuJjg3lJdry694OR0lXyblMYiBWL087gjLyi2QmrZJ5M
# EsCNczBgYRRkTb8EEspTbI1cudKY/bg2wmqqhD+DFrSPeeWJBH9WVnlRlma9MWll
# M+vCjJBDW9LGjuUpHdrEyplj0ZrgRUzUS6NHl38CZPYIxqfTO2Bh8X1QFoKHygZM
# 6z8DG7Ydo+qP5gMquhqYOPRIU6IHSIPgbUPW16hyjCMKfkAcmgabo8aj3vgdRpT9
# b7FYHIYNxO1hhuPtASHskuHAJ5BVcOaHarkd9mUEf6bOEmytC7ls7VCub238oZ6N
# RffcDLxQwCt6C7Yle0TDRvGiu7jclEAdjENfu3mLIw3Ol1GZflC7dAq3AN8Gy82Y
# J/oQpzsYZ4BoZD849itisK1kwZQloII12VGefVA/N2FRR41kSMQxDAz8A/mBR3Fp
# RdHX2uJnLHNXDo3VnKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzIz
# NTNaMC8GCSqGSIb3DQEJBDEiBCDFcRfPeXuxC/lI5PztrcmtnZPD9N8gazTkPnmZ
# 03RjDDANBgkqhkiG9w0BAQEFAASCAgCpSnAFfL5ExUZG9sSxumNiLC/3SED5PDWl
# Tx7PaDLlkbxAWWnRbG6xHH3+sjDNAAZgRonhHxwKzCJdt33nmZBK21KdQ6N0zCUr
# vVPWGboBc6d/+gxvs0dw+FVayw4v0uG3TuBtG6sUNgS9mx5UYe3TZDZhVYOK/+Hq
# E9LcPs7H9IldVmrDXiOMuo27P7TLnPdDfywgQfnOXe59g0J9ypyZ+WltUzWxYSVt
# Sk49/s7g/9mti/Uzh8r0skU/iRnSO9oxDXp02RMf/izW/ILQCDm0ohUPvxcPqUNK
# clALQPqwZzrLBz/rqCmAkepaySo/07IsMNMfldfUs73HhLzBetqQO6t3Ap0hCQCm
# rHPNZ1vD4XrzgdZxgY87PSED8u79SKc8j7ABjAOdDpgk8iFxy5DmObDD6Hd8u+pU
# gy1G/+Zn3OpgSPg/6LGlsWL4AjQ+MadeAX3136Qmba5GsQZYEVg7RH2jE0SsyIeB
# hOzMj3CCxA90jS5LEW2pzqgoJcfsB8G7+isCBDYZ92QvdpPC2fRyp5QQfndQRukZ
# yCGQUqRtq/zaHH6YHOirAwcitgcfvIFxCACCG86n+JX8txUlz7zuh+7LbFmcx+d6
# PK5rEFVYlTSe6/+9B18+Wkm7GZkdD5WlycdsqZL/hPR3IhmGCMzTSv6Gu5ByxGcz
# mdaPm8A5dQ==
# SIG # End signature block
