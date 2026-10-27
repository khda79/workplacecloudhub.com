# Version 1.1.0
# Offline migration projection. Inputs are caller-owned, tenant-qualified snapshots.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'AdCountryEvidence.psm1') -Force
function Key($Value) { ([string]$Value).Trim().ToLowerInvariant() }
function DeviceNameKey($Value) { (Key $Value).TrimEnd('$').Split('.')[0] }
function Value($Row, [string]$Column, [string]$Default = '') {
    if ($null -eq $Row -or -not $Row.PSObject.Properties[$Column] -or [string]::IsNullOrWhiteSpace([string]$Row.$Column)) { return $Default }
    [string]$Row.$Column
}
function Index($Rows, [scriptblock]$KeyOf) {
    $result = @{}
    foreach ($row in $Rows) {
        $key = & $KeyOf $row
        if (-not $key) { continue }
        if (-not $result.ContainsKey($key)) { $result[$key] = [Collections.Generic.List[object]]::new() }
        $result[$key].Add($row)
    }
    return $result
}
function Candidates($Index, [string]$Key) { if ($Key -and $Index.ContainsKey($Key)) { $Index[$Key].ToArray() } }
function NumberOrNull($Value) {
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $n = 0.0
    if (-not [double]::TryParse([string]$Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -or [double]::IsNaN($n) -or [double]::IsInfinity($n) -or $n -lt 0) { throw 'Invalid non-negative invariant number.' }
    return $n
}
function New-WindowsMigrationProjection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Inventory, [Parameter(Mandatory)][object[]]$Prepared,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AD,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Readiness,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Policy,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$IntunePublicationUtc, [Parameter(Mandatory)][string]$ADPublicationUtc,
        [Parameter(Mandatory)]$PolicySelectors,$AdCountryRules)
    if($null -eq $AdCountryRules){$AdCountryRules=New-AdCountryRules -Configuration @{TenantKey=$TenantKey} -ExpectedTenantKey $TenantKey}
    if($AdCountryRules.TenantKey -cne $TenantKey){throw 'AD country rules tenant differs from migration sources.'}
    foreach ($rows in @($Inventory,$AD,$Readiness,$Policy)) {
        foreach ($r in $rows) { if ((Value $r 'TenantKey') -cne $TenantKey) { throw 'Missing or incompatible source TenantKey.' } }
    }
    foreach ($r in $Prepared) { if ($r.PSObject.Properties['TenantKey'] -and (Value $r 'TenantKey') -cne $TenantKey) { throw 'Prepared inventory tenant differs.' } }
    $iid = Index $Inventory { param($r) Key (Value $r 'Device ID') }
    $iname = Index $Inventory { param($r) DeviceNameKey (Value $r 'Device name') }
    if (@($Inventory | Where-Object { -not (Key (Value $_ 'Device ID')) }).Count -or @($iid.Values | Where-Object Count -ne 1).Count) { throw 'Missing or duplicate Intune ID.' }
    $did = Index $Prepared { param($r) Key (Value $r 'Device Source ID') }
    if ($Prepared.Count -ne $Inventory.Count -or $did.Count -ne $Prepared.Count) { throw 'Prepared/raw inventory scope or unique IDs differ.' }
    foreach ($r in $Prepared) {
        $raw = @(Candidates $iid (Key (Value $r 'Device Source ID')))
        if ($raw.Count -ne 1 -or (Key (Value $r 'Operating System Version')) -ne (Key (Value $raw[0] 'OS version')) -or (DeviceNameKey (Value $r 'Device Name')) -ne (DeviceNameKey (Value $raw[0] 'Device name'))) { throw 'Prepared inventory identity or OS differs from source.' }
    }
    $guids = Index $AD { param($r) Key (Value $r 'ObjectGUID') }
    $adnames = Index $AD { param($r) DeviceNameKey (Value $r 'Name' (Value $r 'SamAccountName')) }
    $aid = Index $AD { param($r) Key (Value $r 'IntuneDeviceId') }
    $rid = Index $Readiness { param($r) Key (Value $r 'GraphId') }
    $rn = Index $Readiness { param($r) DeviceNameKey (Value $r 'NormalizedDeviceName' (Value $r 'DeviceName')) }
    $pix = Index $Policy { param($r) (Key (Value $r 'PolicyId')) + '|' + (Key (Value $r 'DeviceId')) }
    if ($PolicySelectors.schemaVersion -ne 1 -or -not $PolicySelectors.anchorPolicyNamePattern -or -not $PolicySelectors.windows11PolicyName) { throw 'Invalid migration policy selectors.' }
    $selected = @{}
    foreach ($prefix in 'Autopatch','Windows 11') {
        $ids = @($Policy | Where-Object {
            if ($prefix -eq 'Autopatch') { (Value $_ 'PolicyName') -match $PolicySelectors.anchorPolicyNamePattern }
            else { (Value $_ 'PolicyName') -ceq $PolicySelectors.windows11PolicyName }
        } | ForEach-Object { Key (Value $_ 'PolicyId') } | Sort-Object -Unique)
        if ($ids.Count -gt 1 -or ($ids.Count -eq 1 -and -not $ids[0])) { throw "Missing or ambiguous selected policy ID: $prefix. Review selectors; no arbitrary policy chosen." }
        $selected[$prefix] = if ($ids.Count) { $ids[0] } else { '' }
    }
    $adOnly = [Collections.Generic.List[object]]::new()
    $excluded = @{}; $matched = 0; $matchedWin11 = 0
    $enabledAD = @($AD | Where-Object { (Value $_ 'Enabled') -match '^(true|yes|1)$' -and (Value $_ 'OperatingSystemShortName') -eq 'Windows 10' })
    foreach ($r in $enabledAD) {
        $gid = Key (Value $r 'ObjectGUID'); $nm = DeviceNameKey (Value $r 'Name' (Value $r 'SamAccountName'))
        $direct = @(Candidates $iid (Key (Value $r 'IntuneDeviceId'))); $names = @(Candidates $iname $nm)
        $reason = ''
        if (-not $gid -or @(Candidates $guids $gid).Count -ne 1) { $reason = 'Missing or duplicate AD GUID' }
        elseif (@(@($direct)+@($names) | Where-Object {(Value $_ 'Identity Status') -eq 'Identity conflict'}).Count){$reason='Identity conflict; reconciliation indeterminate'}
        elseif ($direct.Count -eq 1 -or $names.Count -eq 1) {
            $m = if ($direct.Count -eq 1) { $direct[0] } else { $names[0] }
            $matched++
            $parts = (Value $m 'OS version').Split('.')
            if ($parts.Count -gt 2 -and (NumberOrNull $parts[2]) -ge 22000) { $matchedWin11++ }
            continue
        } elseif ($names.Count -gt 1 -or -not $nm -or @(Candidates $adnames $nm).Count -ne 1) { $reason = 'Ambiguous name reconciliation' }
        if ($reason) { if (-not $excluded.ContainsKey($reason)) { $excluded[$reason] = 0 }; $excluded[$reason]++; continue }
        $adOnly.Add($r)
    }
    $windows = [Collections.Generic.List[object]]::new()
    foreach ($r in $Prepared | Where-Object { (Value $_ 'Windows Generation') -eq 'Windows 10' }) {
        $id = Key (Value $r 'Device Source ID'); $nm = DeviceNameKey (Value $r 'Device Name')
        $parts = (Value $r 'Operating System Version').Split('.')
        $build = if ($parts.Count -gt 2) { NumberOrNull $parts[2] } else { $null }
        $free = NumberOrNull (Value $r 'Free Disk Space (GB)')
        $eligibility = Value $r 'Windows 11 Upgrade Eligibility' 'Unknown'
        $hardware = $eligibility -eq 'Not capable'; $unknown = $eligibility -notin 'Capable','Not capable'
        $legacy = $null -ne $build -and $build -lt 19045
        $risk = if ($hardware) { 'Hardware incompatible' } elseif ($eligibility -eq 'Already upgraded') { 'Conflicting OS evidence' } elseif ($unknown) { 'Eligibility unknown' } elseif ($null -ne $free -and $free -lt 20) { 'Disk remediation required' } elseif ($legacy) { 'Legacy build - update path review' } elseif ((Value $r 'Windows Update State') -eq 'Error') { 'Deployment error' } else { 'Autopatch registration unknown' }
        $conflict=(Value (@(Candidates $iid $id)[0]) 'Identity Status') -eq 'Identity conflict'
        $da = @(if(-not $conflict){Candidates $aid $id}); $d = if ($da.Count -eq 1) { $da[0] } else { $null }
        $qa = @(Candidates $rid $id); $na = @(Candidates $rn $nm)
        $re = if ($qa.Count -eq 1) { $qa[0] } elseif (-not $conflict -and $qa.Count -eq 0 -and $na.Count -eq 1) { $na[0] } else { $null }
        $out = [ordered]@{
            'Device Source ID'=$id; 'Device Name'=(Value $r 'Device Name'); 'Country'=(Value $r 'Country' 'Unknown'); 'Primary User UPN'=(Value $r 'Primary User UPN')
            'OS Version'=(Value $r 'Operating System Version'); 'Windows Release'=(Value $r 'Windows Release'); 'Intune State'=(Value $r 'Management State')
            'Autopatch Registration'='Unknown - registration inventory not collected'; 'Upgrade Eligibility'=$eligibility; 'Hardware Reasons'=(Value $r 'Windows 11 Blocking Reasons' 'Not observed')
            'Free Disk GB'=$free; 'Disk Review'=$(if ($null -eq $free) { 'Unknown' } elseif ($free -lt 20) { 'Below 20 GB planning guardrail' } else { 'At least 20 GB - not upgrade proof' })
            'Legacy Build Review'=$(if ($null -eq $build) { 'Unknown' } elseif ($legacy) { 'Pre-22H2 - verify edition/channel and upgrade route' } else { '22H2 - edition/channel still required' })
            'Update State'=(Value $r 'Windows Update State'); 'Update Action'=(Value $r 'Windows Update Action Code'); 'Last Intune Sync'=(Value $r 'Last Sync DateTime')
            'Migration Status'=$risk; 'Risk Priority'=$(if ($hardware -or $risk -eq 'Conflicting OS evidence' -or (Value $r 'Windows Update State') -eq 'Error') { 'High' } else { 'Review' })
            'Next Action'=$(if ($hardware) { 'Review replacement or hardware remediation' } elseif ($risk -eq 'Conflicting OS evidence') { 'Reconcile current OS and readiness evidence' } elseif ($null -ne $free -and $free -lt 20) { 'Free space; verify deployment-specific requirement' } elseif ($legacy) { 'Validate OS edition/channel and supported staged upgrade path' } else { 'Collect Autopatch registration, policy targeting and safeguards' })
            'Evidence Date'=$IntunePublicationUtc.Substring(0,10); 'Country Basis'='Primary user directory country; not physical device location'
            'AD Domain'=$(if ($d) { Value $d 'DomainName' 'Unknown' | ForEach-Object { Key $_ } } else { 'Unknown' })
            'AD Domain Match Status'=$(if ($d) { 'Unique AD match' } elseif ($da.Count -gt 1) { 'Ambiguous AD match' } else { 'No qualified AD match' })
            'Inventory Source'='Intune'; 'Inventory Observation Date'=$IntunePublicationUtc
            'Eligibility Observation Date'=(Value $re 'ExportDateTime' 'Not observed'); 'Eligibility OS Version'=(Value $re 'OSVersion' 'Not observed')
            'Eligibility Match Status'=$(if ($re) { if ($qa.Count -eq 1) { 'Unique readiness ID' } else { 'Unique readiness name' } } elseif ($qa.Count -gt 1 -or $na.Count -gt 1) { 'Ambiguous readiness match' } else { 'Not observed' })
        }
        if ($re -and $qa.Count -eq 0 -and @(Candidates $iname $nm).Count -ne 1) {
            $out['Eligibility Match Status']='Readiness name only - ambiguous Intune identity'
            if ($risk -eq 'Conflicting OS evidence') { $out['Next Action']='Reconcile duplicate Intune identities and readiness name match before judging OS' }
        }
        foreach ($prefix in 'Autopatch','Windows 11') {
            $ps = @(if ($selected[$prefix]) { Candidates $pix ($selected[$prefix]+'|'+$id) })
            $p = if ($ps.Count -eq 1) { $ps[0] } else { $null }
            $missing = if ($ps.Count -gt 1) { 'Ambiguous policy match' } else { 'Not observed' }
            $out[$prefix+' Policy State']=Value $p 'AggregateState' $missing
            $out[$prefix+' Device Status']=Value $p 'CurrentDeviceUpdateStatus_loc' (Value $p 'CurrentDeviceUpdateStatus' $missing)
            $out[$prefix+' Blocking Reason']=Value $p 'BlockingReason' $missing
            $out[$prefix+' Alert']=Value $p 'LatestAlertMessage_loc' (Value $p 'LatestAlertMessage' $missing)
            $out[$prefix+' Observation Date']=Value $p 'ExportDateTime' $missing
        }
        if($conflict){
            $out['Country']='Unknown';$out['Country Basis']='Identity conflict; attribution indeterminate'
            $out['AD Domain']='Unknown';$out['AD Domain Match Status']='Identity conflict; indeterminate'
            $out['Migration Status']='Identity conflict';$out['Next Action']='Resolve independent Intune identity observations before cross-directory decisions'
            if($qa.Count -eq 0){$out['Eligibility Match Status']='Identity conflict; name matching excluded'}
        }
        $windows.Add([pscustomobject]$out)
    }
    # Stable schema even when there are no observed Intune Windows 10.
    $columns = @('Device Source ID','Device Name','Country','Primary User UPN','OS Version','Windows Release','Intune State','Autopatch Registration','Upgrade Eligibility','Hardware Reasons','Free Disk GB','Disk Review','Legacy Build Review','Update State','Update Action','Last Intune Sync','Migration Status','Risk Priority','Next Action','Evidence Date','Country Basis','AD Domain','AD Domain Match Status','Inventory Source','Inventory Observation Date','Eligibility Observation Date','Eligibility OS Version','Eligibility Match Status')
    foreach ($prefix in 'Autopatch','Windows 11') { foreach ($suffix in 'Policy State','Device Status','Blocking Reason','Alert','Observation Date') { $columns += "$prefix $suffix" } }
    $adRows = [Collections.Generic.List[object]]::new()
    foreach ($r in $adOnly) {
        $gid = Key (Value $r 'ObjectGUID'); $nm = Value $r 'Name' (Value $r 'SamAccountName'); $domain = Key (Value $r 'DomainName' 'Unknown')
        $out = [ordered]@{}; foreach ($c in $columns) { $out[$c]='Not observed' }
        $country=Resolve-AdComputerCountry -Domain $domain -ComputerName $nm -Rules $AdCountryRules
        $out['Device Source ID']="ad:${TenantKey}:$gid"; $out['Device Name']=$nm; $out['Country']=$country.Country; $out['OS Version']=Value $r 'operatingSystemVersion' 'Unknown'
        $out['Windows Release']=Value $r 'OperatingSystemDisplayName' 'Windows 10'; $out['Intune State']='Not observed in current Intune export'
        $out['Autopatch Registration']='Unknown - no qualified Intune match'; $out['Upgrade Eligibility']='Not assessed'; $out['Hardware Reasons']='Not assessed'
        $out['Free Disk GB']=$null; $out['Disk Review']='Unknown'; $out['Legacy Build Review']='Unknown - AD only'; $out['Update Action']='COLLECT_INTUNE_EVIDENCE'
        $out['Migration Status']='AD only - no current Intune match'; $out['Risk Priority']='Review'; $out['Next Action']='Validate AD device activity and enroll in Intune if appropriate'
        $out['Evidence Date']=$ADPublicationUtc.Substring(0,10); $out['Country Basis']=$country.Basis
        $out['AD Domain']=$domain; $out['AD Domain Match Status']='AD source record'; $out['Inventory Source']='AD only'; $out['Inventory Observation Date']=$ADPublicationUtc
        $out['Eligibility Match Status']='Not assessed - AD only'; $windows.Add([pscustomobject]$out)
        $adRows.Add([pscustomobject][ordered]@{'AD Object GUID'=$gid; 'Device Name'=$nm; 'Source Domain'=$domain; 'OS Version'=$out['OS Version']; 'Intune Evidence'='No match in current local Intune export'; 'Autopatch Evidence'='Unknown'; 'Country'=$country.Country; 'Qualification'=('Enabled AD Windows 10; unique AD GUID/name; no Intune ID or unique-name match; '+$country.Basis); 'Next Action'='Validate activity and management; absence from export is not proof of non-enrollment'})
    }
    $unionIds = Index $windows { param($r) Key $r.'Device Source ID' }
    if ($unionIds.Count -ne $windows.Count) { throw 'Duplicate union identity.' }
    [pscustomobject]@{ Windows=$windows.ToArray(); ADOnly=$adRows.ToArray(); Columns=$columns
        ADColumns=@('AD Object GUID','Device Name','Source Domain','OS Version','Intune Evidence','Autopatch Evidence','Country','Qualification','Next Action')
        Audit=[ordered]@{ TenantKey=$TenantKey; IntuneRows=$Inventory.Count; IntuneWindows10=$windows.Count-$adRows.Count; ADOnly=$adRows.Count; Total=$windows.Count; EnabledADWindows10=$enabledAD.Count; MatchedAD=$matched; MatchedADIntuneWindows11=$matchedWin11; ExcludedAD=$excluded; Policies=$selected
            Unknown=@($windows | Where-Object { $_.'Upgrade Eligibility' -in 'Not assessed','Unknown' }).Count; Conflicting=@($windows | Where-Object { $_.'Upgrade Eligibility' -eq 'Already upgraded' }).Count
            Hardware=@($windows | Where-Object { $_.'Upgrade Eligibility' -eq 'Not capable' }).Count; AnchorObserved=@($windows | Where-Object { $_.'Autopatch Policy State' -in 'Success','InProgress','Error' }).Count
            Dates='Inventory dates are source file publication times, not device events. Readiness/policy dates are record export times.'
            Scope='Intune plus enabled AD Windows 10 with unique AD GUID/name and no qualified local Intune match. Absence is not proof of non-enrollment.'
            Country='Primary-user proxy for Intune; configured exact AD domain, then configured name prefix for AD only; conflicting rules yield Unknown, not verified physical location.'; History='Unchanged Intune-only historical scope.' }
    }
}
Export-ModuleMember -Function New-WindowsMigrationProjection

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBa3/HfzQySOyER
# Y+kC71N2vByjaOrT0GV45TPOFsGCB6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC2FRuXJ4O13xBcZ/19h5EFeorSfHMFYhOi/qpY/fk99MA0GCSqG
# SIb3DQEBAQUABIIBgIV5WfBP7Xv0HCYlrNjYcXRxd0A9xvYpqyOOIogUQ4oW7j6O
# 1wcdp4YgXglH5GBTb4+Ww2lt5CkdkeJ6JNwf8iQFFz0FLb9p31pjM82Lm1JyuNEc
# CqkIaVMEYVibjRpbx8nV1FznyiKX07ZLjoU8G9mywwydqsuyWJUy/XYHICnm6uev
# kO6Cln4L+lDXILYKwx8ce/1TSFE1qEVgr/k3caAGeU2s5eQFJ6MESu/LNYc2iDvW
# U+tldUbrGN/w/NwjoZAoFBc8yTO71c28YZD01Iru/yHKatWyNwpvFRRa37zIRs2w
# lWn3G2sjSitaj2+H0YMl8kCtLeO9QqdTQB6GqefEIT/rCkw7BM0tGZH5bM1tzWAO
# Rk11xM1i2ns4Ju9PwjYEZxVvN7A8owb/krMmY2Zm4NfWYW8oz/eagxzLT57exlOP
# ZG1XXhyPwW2v4TQxL8R1yz4vYGBwe14O6Pg/ttp1w1vxBYXYIREAzY6luOrY3aXK
# gEvGgGegNMdmG5IAEKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzIz
# NTVaMC8GCSqGSIb3DQEJBDEiBCB6UIJX/Qfudt1ArRy8Z6HR/NTL1ToG86TNwaCM
# IttBSzANBgkqhkiG9w0BAQEFAASCAgBw8at1JCnldgHQwFy4qElRb7sQNoOkOCRs
# 5aA/WAgwmwqzKbuq4yMI6/w1o3A4c1OyRMXyJ8Dq4V9Z6SfswnzJGznTvs4luzXP
# uGk0ph0YO+cRtCZuaTK9SDyR7SknQk+Iz0O/nyjEZzA5ZLh4d1uErgNCyuFBZ54l
# JVFHSGymZ6Hwf5hDlWyQOUTlrCKIWmt8HVJLM3LxUTV8vUOVHTU5ctygdHTJF6oc
# 4yJJXkUfPHoK78ra2kiC5SJKTOjcGiUUOJ9fXW5aV5sKw+Xq0ROB6P3bJ1cLcdB2
# gUeZHrEvxUuvJjyqP9/jmQwE1FbCA/9BXzKwl7hfsWGrwydJ+tFLqYWSnR/HnOhL
# +/r/be5o+Nqk8815EWK/KMKhj2IB9EDVSCd3mkAY0devSo/V7EKgmXhPxz7Gg5X6
# Nn7Y0LF6kU6GKs+UzZOCFYNBjKBrZ8aNtgsRKIyfLxGPw2+2jbLX1R/XsJTCf+dp
# Qi6gvVucIGCTPXom2ljXE2ugkwsmmbybQD14EvHJn1aQY099o11ObxpjwzRq4TK8
# qmiQMneVmEmKAJNsuhhvgUTm/PQ04S35x9KU/zbE7SvXDnEq+KsdjSnnEIozX/5h
# Ns7U6iobfqtOtAFoxhzluEvItjtBFjvNLSfglvEKHjZYqqmLm7mJ/yoW3meu+uzw
# /S0DPQUM5w==
# SIG # End signature block
