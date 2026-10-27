# Pure, offline identity projection. Names, UPNs and serial numbers never merge records.
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'CountryEvidence.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'AdCountryEvidence.psm1') -Force
function Get-EstateValue($Row,[string]$Name) {
    if ($null -eq $Row -or -not $Row.PSObject.Properties[$Name]) { return '' }
    return ([string]$Row.$Name).Trim()
}
function Get-EstateGuid([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $parsed=[guid]::Empty
    if (-not [guid]::TryParse($Value.Trim(),[ref]$parsed)) { throw 'Invalid identity GUID.' }
    if ($parsed -eq [guid]::Empty) { return '' }
    return $parsed.ToString('D').ToLowerInvariant()
}
function New-EstateIndex($Rows,[string]$Column,[switch]$Guid,[switch]$CaseSensitive,[switch]$Required) {
    $comparer=if($CaseSensitive){[StringComparer]::Ordinal}else{[StringComparer]::OrdinalIgnoreCase}
    $index=[Collections.Generic.Dictionary[string,object]]::new($comparer)
    foreach($row in $Rows){
        $key=Get-EstateValue $row $Column
        if($Guid){$key=Get-EstateGuid $key}
        if(-not $key){if($Required){throw "Missing identity key: $Column"};continue}
        if($index.ContainsKey($key)){throw "Duplicate identity key: $Column"}
        $index.Add($key,$row)
    }
    return ,$index
}
function Get-EstateCountry($Cloud,$AD) {
    # A linked Entra user with unknown country never falls back to AD or licensing.
    if ($null -ne $Cloud) {
        $result=Resolve-DirectoryCountry (Get-EstateValue $Cloud 'CountryOrRegion')
        return [pscustomobject]@{Country=$result.Name;Basis=$(if($result.Code){'Entra country'}else{'Entra country unavailable'})}
    }
    foreach($field in @('Country','CountryName')) {
        $result=Resolve-DirectoryCountry (Get-EstateValue $AD $field)
        if($result.Code){return [pscustomobject]@{Country=$result.Name;Basis='AD country (AD-only account)'}}
    }
    return [pscustomobject]@{Country='Unknown';Basis='Not observed'}
}
function Get-EstateState($Row,[string]$Column) {
    $s=Get-EstateValue $Row $Column
    if($s -eq 'true'){return 'Enabled'};if($s -eq 'false'){return 'Disabled'};return 'Unknown'
}
function New-EnterpriseEstateProjection {
    [CmdletBinding()]
    param([object[]]$ADUsers,[object[]]$EntraUsers,[object[]]$DirectoryObjects,[object[]]$ManagedDevices,
          [object[]]$ADComputers,[string]$ADObservationUtc,[string]$EntraObservationUtc,[string]$IntuneObservationUtc,
          [string]$ADUserObservationUtc,$AdCountryRules)
    $identityColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')
    if(-not $EntraUsers.Count -or -not $DirectoryObjects.Count -or -not $ManagedDevices.Count){throw 'Required current inventory is empty.'}
    $tenant=[ordered]@{}
    foreach($col in $identityColumns){$tenant[$col]=Get-EstateValue $EntraUsers[0] $col;if(-not $tenant[$col]){throw "Missing tenant identity: $col"}}
    foreach($set in @($ADUsers,$EntraUsers,$DirectoryObjects,$ManagedDevices,$ADComputers)){
        foreach($row in $set){foreach($col in $identityColumns){if((Get-EstateValue $row $col) -cne $tenant[$col]){throw "Cross-tenant or missing source identity: $col"}}}
    }
    $adIds=New-EstateIndex $ADUsers 'ObjectGUID' -Guid -Required
    $null=New-EstateIndex $ADUsers 'ObjectSID' -Required
    $null=New-EstateIndex $ADUsers 'ImmutableId_AD' -CaseSensitive -Required
    $cloudIds=New-EstateIndex $EntraUsers 'Object Id' -Guid -Required
    $cloudSid=New-EstateIndex $EntraUsers 'OnPremisesSecurityIdentifier'
    $cloudImm=New-EstateIndex $EntraUsers 'OnPremisesImmutableId' -CaseSensitive
    $directoryIds=New-EstateIndex $DirectoryObjects 'ObjectGUID' -Guid -Required
    $adComputerIndex=New-EstateIndex $ADComputers 'ObjectGUID' -Guid -Required
    if($null -eq $AdCountryRules){$AdCountryRules=New-AdCountryRules -Configuration @{TenantKey=$tenant.TenantKey} -ExpectedTenantKey $tenant.TenantKey}
    if($AdCountryRules.TenantKey -cne $tenant.TenantKey){throw 'AD country rules tenant differs from estate sources.'}
    $managedIds=New-EstateIndex $ManagedDevices 'ManagedDeviceId' -Guid -Required
    $managedAzure=New-EstateIndex $ManagedDevices 'AzureADDeviceId' -Guid
    $directoryUsers=@($DirectoryObjects|Where-Object { (Get-EstateValue $_ 'ObjectClass') -eq 'user' })
    if($directoryUsers.Count -ne $ADUsers.Count){throw 'AD user inventory disagrees with directory-object coverage.'}
    foreach($row in $directoryUsers){if(-not $adIds.ContainsKey((Get-EstateGuid (Get-EstateValue $row 'ObjectGUID')))){throw 'AD user absent from user inventory.'}}
    foreach($row in $ADComputers){$id=Get-EstateGuid (Get-EstateValue $row 'ObjectGUID');if(-not $directoryIds.ContainsKey($id) -or (Get-EstateValue $directoryIds[$id] 'ObjectClass') -ne 'computer'){throw 'AD computer absent from directory-object inventory.'}}
    $accounts=[Collections.Generic.List[object]]::new();$devices=[Collections.Generic.List[object]]::new()
    $usedCloud=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $usedManaged=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $accountLinks=0;$deviceLinks=0;$upnDifferences=0;$services=0
    $adAccounts=@($ADUsers)+@($DirectoryObjects|Where-Object { (Get-EstateValue $_ 'ObjectClass') -in @('msDS-GroupManagedServiceAccount','msDS-ManagedServiceAccount') })
    foreach($a in $adAccounts){
        $id=Get-EstateGuid (Get-EstateValue $a 'ObjectGUID');$sid=Get-EstateValue $a 'ObjectSID'
        $immutable=Get-EstateValue $a 'ImmutableId_AD'
        $kind=Get-EstateValue $a 'ObjectClass';if(-not $kind){$kind='user'}
        if($kind -ne 'user'){$services++;$immutable=[Convert]::ToBase64String(([guid]$id).ToByteArray())}
        $s=$null;$i=$null
        if($sid){$null=$cloudSid.TryGetValue($sid,[ref]$s)}
        if($immutable){$null=$cloudImm.TryGetValue($immutable,[ref]$i)}
        if($s -and $i -and (Get-EstateGuid (Get-EstateValue $s 'Object Id')) -ne (Get-EstateGuid (Get-EstateValue $i 'Object Id'))){throw 'Account SID and immutable ID point to different Entra identities.'}
        $c=if($s){$s}else{$i};$cid=Get-EstateGuid (Get-EstateValue $c 'Object Id')
        $match=if($s -and $i){'SID + immutable ID'}elseif($s){'SID'}elseif($i){'Immutable ID'}else{'AD only'}
        $upnStatus='Not compared'
        if($c){
            if(-not $usedCloud.Add($cid)){throw 'Multiple AD accounts point to one Entra identity.'};$accountLinks++
            $au=Get-EstateValue $a 'UserPrincipalName';$cu=Get-EstateValue $c 'User principal name'
            if($au -and $cu){$upnStatus=if($au -ieq $cu){'Same'}else{'Different; ID match retained'};if($au -ine $cu){$upnDifferences++}}
        }
        $country=Get-EstateCountry $c $a
        $row=[ordered]@{};foreach($col in $identityColumns){$row[$col]=$tenant[$col]}
        $row['Account Key']='ad:'+$id;$row['Inventory Source']=if($c){'AD + Entra'}else{'AD only'}
        $row['Account Kind']=$kind;$row['AD Object GUID']=$id;$row['Entra Object ID']=$cid
        $row['AD State']=Get-EstateState $a 'Enabled';$row['Entra State']=Get-EstateState $c 'AccountEnabled'
        $row['Entra User Type']=Get-EstateValue $c 'UserType';$row['Country']=$country.Country;$row['Country Basis']=$country.Basis
        $row['Match Basis']=$match;$row['UPN Comparison']=$upnStatus
        $row['AD Observation UTC']=if($kind -eq 'user' -and $ADUserObservationUtc){$ADUserObservationUtc}else{$ADObservationUtc}
        $row['Entra Observation UTC']=if($c){$EntraObservationUtc}else{''}
        $accounts.Add([pscustomobject]$row)
    }
    foreach($c in $EntraUsers){
        $cid=Get-EstateGuid (Get-EstateValue $c 'Object Id');if($usedCloud.Contains($cid)){continue}
        $country=Get-EstateCountry $c $null;$row=[ordered]@{};foreach($col in $identityColumns){$row[$col]=$tenant[$col]}
        $row['Account Key']='entra:'+$cid;$row['Inventory Source']='Entra only';$row['Account Kind']='user'
        $row['AD Object GUID']='';$row['Entra Object ID']=$cid;$row['AD State']='Unknown'
        $row['Entra State']=Get-EstateState $c 'AccountEnabled';$row['Entra User Type']=Get-EstateValue $c 'UserType'
        $row['Country']=$country.Country;$row['Country Basis']=$country.Basis;$row['Match Basis']='Entra only';$row['UPN Comparison']='Not compared'
        $row['AD Observation UTC']='';$row['Entra Observation UTC']=$EntraObservationUtc;$accounts.Add([pscustomobject]$row)
    }
    foreach($a in @($DirectoryObjects|Where-Object { (Get-EstateValue $_ 'ObjectClass') -eq 'computer' })){
        $id=Get-EstateGuid (Get-EstateValue $a 'ObjectGUID');$m=$null;$null=$managedAzure.TryGetValue($id,[ref]$m)
        $mid=Get-EstateGuid (Get-EstateValue $m 'ManagedDeviceId')
        if($m){if(-not $usedManaged.Add($mid)){throw 'Multiple AD computers point to one Intune identity.'};$deviceLinks++}
        $userId=Get-EstateGuid (Get-EstateValue $m 'UserId');$u=$null;if($userId){$null=$cloudIds.TryGetValue($userId,[ref]$u)}
        $country=Get-EstateCountry $u $null
        if(-not $m){
            $detail=$null;$null=$adComputerIndex.TryGetValue($id,[ref]$detail)
            $domain=Get-EstateValue $a 'DomainName';if(-not $domain){$domain=Get-EstateValue $detail 'DomainName'}
            $name=Get-EstateValue $a 'Name';if(-not $name){$name=Get-EstateValue $detail 'Name'}
            if(-not $name){$name=Get-EstateValue $detail 'SamAccountName'}
            $country=Resolve-AdComputerCountry -Domain $domain -ComputerName $name -Rules $AdCountryRules
        }
        $row=[ordered]@{};foreach($col in $identityColumns){$row[$col]=$tenant[$col]}
        $row['Device Key']='ad:'+$id;$row['Inventory Source']=if($m){'AD + Intune'}else{'AD only'}
        $row['AD Object GUID']=$id;$row['Intune Device ID']=$mid;$row['Azure AD Device ID']=Get-EstateGuid (Get-EstateValue $m 'AzureADDeviceId')
        $row['Platform']=Get-EstateValue $m 'OperatingSystem';if(-not $row['Platform']){$row['Platform']='Unknown'}
        $row['Country']=$country.Country;$row['Country Basis']=if(-not $m){$country.Basis}elseif($u){'Primary-user proxy: '+$country.Basis}else{'Not observed'}
        $row['Match Basis']=if($m){'AD GUID = Azure AD device ID'}else{'AD only'}
        $row['AD Observation UTC']=$ADObservationUtc;$row['Intune Observation UTC']=if($m){$IntuneObservationUtc}else{''};$devices.Add([pscustomobject]$row)
    }
    foreach($m in $ManagedDevices){
        $mid=Get-EstateGuid (Get-EstateValue $m 'ManagedDeviceId');if($usedManaged.Contains($mid)){continue}
        $uid=Get-EstateGuid (Get-EstateValue $m 'UserId');$u=$null;if($uid){$null=$cloudIds.TryGetValue($uid,[ref]$u)}
        $country=Get-EstateCountry $u $null;$row=[ordered]@{};foreach($col in $identityColumns){$row[$col]=$tenant[$col]}
        $row['Device Key']='intune:'+$mid;$row['Inventory Source']='Intune only';$row['AD Object GUID']='';$row['Intune Device ID']=$mid
        $row['Azure AD Device ID']=Get-EstateGuid (Get-EstateValue $m 'AzureADDeviceId');$row['Platform']=Get-EstateValue $m 'OperatingSystem'
        if(-not $row['Platform']){$row['Platform']='Unknown'}
        $row['Country']=$country.Country;$row['Country Basis']=if($u){'Primary-user proxy: '+$country.Basis}else{'Not observed'}
        $row['Match Basis']='Intune only';$row['AD Observation UTC']='';$row['Intune Observation UTC']=$IntuneObservationUtc;$devices.Add([pscustomobject]$row)
    }
    $null=New-EstateIndex $accounts 'Account Key' -Required;$null=New-EstateIndex $devices 'Device Key' -Required
    return [pscustomobject]@{Accounts=$accounts.ToArray();Devices=$devices.ToArray();Audit=[ordered]@{
        ADUsers=$ADUsers.Count;EntraUsers=$EntraUsers.Count;ADManagedServiceAccounts=$services;AccountStrongLinks=$accountLinks
        AccountUPNDifferences=$upnDifferences;AllAccounts=$accounts.Count;ADComputers=@($DirectoryObjects|Where-Object ObjectClass -eq 'computer').Count
        IntuneAllPlatforms=$ManagedDevices.Count;DeviceStrongLinks=$deviceLinks;AllDevices=$devices.Count
        IdentityGrain='Directory/management identities, not verified physical machines';NameBasedMerges=0;UPNBasedMerges=0
    }}
}
Export-ModuleMember -Function New-EnterpriseEstateProjection

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASV4cjLbRZ68in
# AR6wmaxouTfyQCyrjHw+tGbDFKowKaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOXvbqP+60XCGDMYEnxJBWaGTIUsxqNTJqqV4ltburu9MA0GCSqG
# SIb3DQEBAQUABIIBgFQ/QOzbCtFFkKtaq2pHPrgQ61ytFgjo8cU0hyR4auucmCHh
# OwACiKbILXfnUCIjNt9cJ4/cseBo6YxcV7fOhYgK0pvpD95kzbfF4x7SZyZefJ9u
# LYzyvbH7rH9GYyihyR0adUBjzNYN1XJT7NYKTrgNjpGqjHVhn0ZcRud67ksuAGb2
# Ae+pedfWbkiRiMfnzzat6EbzBTBYlxbgNyazEhhm9M5qUfpbB4khiSalMm0QNQhS
# XVbhlMSrh5Q83arml/EIRezSPSzQcn4G3xeEH9lgo1ayw7wvQjzUWW7Iw40wobXV
# ZSPJA44BzNBh8py7EPI61SeTc2oedoGxMw0YoKdkwd09wRbd2G3t2qutc72aiOZ2
# 86L6d/c/ZhtaUuvYxyh0/D9Q0LCGCW2IEKQ+VLC3PLAOYllBUBngIB+t+zOYZsy9
# +D10gQaFG1IKQ3O18vwMM6+g7OWe7PwqLzD6ywLXKN5QNHvqQv9aQ/kk58rdktgM
# wF35E/rAyYKTmPOB9qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMzQx
# MDNaMC8GCSqGSIb3DQEJBDEiBCDmBa7XFTMnr8QIshqtjiy+6dQSZ4NTgApBHDRh
# lBUd2jANBgkqhkiG9w0BAQEFAASCAgA8a2o7J4sf7xY9RKmJZJ/+Oe5VcWHRPspC
# goYCKQWTdy7rDYah7nr/wPsf0fpuZjaOx5tgxjgJFd2cnn1doOOZ6VkPf+li+JZA
# cGjIYxANtfJDYZtmldERAiJ/7e0/Dq+8pBz369Dxc0Rlz6eJj2gRC/0UbbYIdFnA
# lUxPF6D8Ya1vZWWf9WTwZKL9ZlIzqVdgVTcVazE6q/VyQPKG6ytFm1bLZM+ckewO
# EbJejbauqNzK8aEtEdGzPcGAHQAaRY5xQTK7JofY9x1Mw2nG+e5grN7YZqJqzXAc
# b0/WwIp+wUNaZMvPSb84UdIe2WxoQhLTH7Ml7OZNFM85eX3C0/02ASF18vPB5fjV
# 7jTymjk1fhIDmxJMC7LlRnuSY15+tlM3xp/+kU/nPWdneRtMgGt+SooU/vJPrtla
# bSISBaCWRr04B/T4Vu8wCmj6YZqBrYeejNJdNkW9HjdIq4XeP5yVbzhilRX8bUEZ
# wN32qkHHVhb7wwp4n7ZUkGLc5t9ia64aNBi4UlVpH3E6AG+EbomsZ8c6/Hd6GCQI
# nqRq2ZfjHrMVUiSENNhHVt91bl1GsUeiK4E67opT1Q+VVyjKhEPTrRLxmUgpQnZ+
# d7rNnDI9+4LwhMMuADHsx8lHX4jS1CxvEyI7MS0dLjfgOXHoMKVpizVwglzAr3bt
# oLKHepNecw==
# SIG # End signature block
