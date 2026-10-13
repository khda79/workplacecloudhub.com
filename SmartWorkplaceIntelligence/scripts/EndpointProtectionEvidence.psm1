# Pure offline policy normalization and evaluation. No tenant APIs or filesystem writes.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-EPValue($Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [Collections.IDictionary]) { return ,$Object[$Name] }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -ne $p) { return ,$p.Value }
}
function Get-EPText($Object, [string]$Name) {
    $value=Get-EPValue $Object $Name
    if ($null -eq $value) { return '' }
    ([string]$value).Trim()
}
function Assert-EPObject($Value, [string]$Name) {
    if ($null -eq $Value -or ($Value -isnot [Collections.IDictionary] -and $Value -isnot [pscustomobject])) { throw "$Name must be an object." }
}
function Assert-EPKey([string]$Value, [string]$Name) {
    if (-not $Value -or $Value -match '[|\r\n]') { throw "$Name must be nonempty and must not contain a pipe or newline." }
}
function Get-EPArray($Object, [string]$Name) {
    $v = Get-EPValue $Object $Name
    if ($null -eq $v -or $v -isnot [array]) { throw "$Name must be an array." }
    return ,$v
}
function Convert-EPRequirements([object[]]$Values) {
    $seen = @{}; $activeAV = 0
    foreach ($r in $Values) {
        Assert-EPObject $r 'Requirement'
        $role = Get-EPText $r 'Role'; $product = Get-EPText $r 'Product'; $mode = Get-EPText $r 'ExpectedMode'
        if ($role -cnotin @('Antivirus','EDR')) { throw 'Role must be Antivirus or EDR.' }
        Assert-EPKey $product 'Product'
        if ($mode -cnotin @('Active','Passive','NotRequired')) { throw 'ExpectedMode must be Active, Passive or NotRequired.' }
        if ($mode -eq 'Passive' -and ($role -ne 'Antivirus' -or $product -ine 'Microsoft Defender Antivirus')) { throw 'Passive mode is supported only for Microsoft Defender Antivirus.' }
        $reason = Get-EPText $r 'Reason'
        if ($mode -eq 'NotRequired' -and -not $reason) { throw 'NotRequired requires a documented Reason.' }
        $key = "$role|$product"
        if ($seen.ContainsKey($key)) { throw 'Duplicate role/product requirement.' }; $seen[$key] = $true
        if ($role -eq 'Antivirus' -and $mode -eq 'Active') { $activeAV++ }
        [pscustomobject][ordered]@{ Role=$role; Product=$product; ExpectedMode=$mode; EvidenceProductName=(Get-EPText $r 'EvidenceProductName'); Reason=$reason }
    }
    if ($activeAV -gt 1) { throw 'A scope cannot require more than one active antivirus.' }
}
function ConvertTo-EndpointProtectionPolicy {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Configuration, [Parameter(Mandatory)][string]$TenantKey)
    Assert-EPKey $TenantKey 'TenantKey'
    $configuredTenant = Get-EPText $Configuration 'TenantKey'
    if ($configuredTenant -and $configuredTenant -ine $TenantKey) { throw 'Policy configuration belongs to another tenant.' }
    $raw = Get-EPValue $Configuration 'EndpointProtectionPolicy'
    if ($null -eq $raw) { $raw = @{Enabled=$false;PolicyVersion='';MaxObservationAgeHours=$null;Requirements=@();ScopeRules=@()} }
    Assert-EPObject $raw 'EndpointProtectionPolicy'
    $enabled = Get-EPValue $raw 'Enabled'
    if ($enabled -isnot [bool]) { throw 'Enabled must be a JSON boolean.' }
    $version = Get-EPText $raw 'PolicyVersion'
    if ($enabled) { Assert-EPKey $version 'PolicyVersion' }
    $age = Get-EPValue $raw 'MaxObservationAgeHours'
    if ($null -ne $age) {
        if ($age -is [bool] -or $age -is [string]) { throw 'MaxObservationAgeHours must be a number or null.' }
        $age = [double]$age
        if ($age -le 0 -or [double]::IsNaN($age) -or [double]::IsInfinity($age)) { throw 'MaxObservationAgeHours must be finite and positive.' }
    }
    $requirements = @(Convert-EPRequirements (Get-EPArray $raw 'Requirements'))
    $seen = @{}
    $scopes = @(foreach ($s in (Get-EPArray $raw 'ScopeRules')) {
        Assert-EPObject $s 'Scope rule'
        $id = Get-EPText $s 'ScopeId'; Assert-EPKey $id 'ScopeId'
        if ($id -in @('Default','Ambiguous','Disabled') -or $seen.ContainsKey($id)) { throw 'Reserved or duplicate ScopeId.' }; $seen[$id]=$true
        $ids = Get-EPArray $s 'DeviceSourceIds'; $idSeen=@{}
        if (-not $ids.Count) { throw 'Scope rules require explicit DeviceSourceIds.' }
        $cleanIds = @(foreach ($deviceId in $ids) {
            if ($deviceId -isnot [string]) { throw 'DeviceSourceIds must contain strings.' }
            $deviceId=$deviceId.Trim().ToLowerInvariant(); Assert-EPKey $deviceId 'DeviceSourceId'
            if ($idSeen.ContainsKey($deviceId)) { throw 'Duplicate DeviceSourceId within a scope.' }; $idSeen[$deviceId]=$true; $deviceId
        })
        [pscustomobject][ordered]@{ScopeId=$id;DeviceSourceIds=$cleanIds;Requirements=@(Convert-EPRequirements (Get-EPArray $s 'Requirements'))}
    })
    # Explicit allowlist: never return the complete effective tenant profile.
    [pscustomobject][ordered]@{TenantKey=$TenantKey;Enabled=$enabled;PolicyVersion=$version;MaxObservationAgeHours=$age;Requirements=$requirements;ScopeRules=$scopes}
}
function Get-EndpointProtectionPolicyColumns {
    @('TenantKey','Policy Key','Policy Version','Scope ID','Device Source IDs','Role','Expected Product','Expected Mode','Evidence Product Name','Exception Reason','Max Observation Age Hours')
}
function Get-EndpointProtectionEvidenceColumns {
    @('TenantKey','Endpoint Protection Key','Device Source ID','Country','Policy Version','Scope ID','Role','Expected Product','Expected Mode','Evaluation State','Evaluation Reason','Registration State','Evidence Source','Observation UTC','Freshness','Exception Reason')
}
function New-EPUnknownRequirement { [pscustomobject]@{Role='Unknown';Product='Unknown';ExpectedMode='Unknown';EvidenceProductName='';Reason=''} }
function New-EndpointProtectionPolicyRows {
    param([Parameter(Mandatory)]$Policy)
    if (-not $Policy.Enabled) { return }
    $scopes = @([pscustomobject]@{ScopeId='Default';DeviceSourceIds=@();Requirements=$Policy.Requirements}) + @($Policy.ScopeRules)
    foreach ($s in $scopes) {
        $requirements = @($s.Requirements)
        if (-not $requirements.Count) { $requirements=@(New-EPUnknownRequirement) }
        foreach ($r in $requirements) {
            [pscustomobject][ordered]@{
                TenantKey=$Policy.TenantKey; 'Policy Key'="$($Policy.TenantKey)|$($s.ScopeId)|$($r.Role)|$($r.Product)".ToLowerInvariant()
                'Policy Version'=$Policy.PolicyVersion; 'Scope ID'=$s.ScopeId; 'Device Source IDs'=ConvertTo-Json -InputObject @($s.DeviceSourceIds) -Compress
                Role=$r.Role; 'Expected Product'=$r.Product; 'Expected Mode'=$r.ExpectedMode; 'Evidence Product Name'=$r.EvidenceProductName
                'Exception Reason'=$r.Reason; 'Max Observation Age Hours'=if ($null -ne $Policy.MaxObservationAgeHours) {$Policy.MaxObservationAgeHours.ToString('R',[Globalization.CultureInfo]::InvariantCulture)} else {''}
            }
        }
    }
}
function Get-EPBoolean($Value) {
    if ($null -eq $Value) { return $null }
    switch (([string]$Value).Trim().ToLowerInvariant()) {
        {$_ -in @('true','1','yes','enabled')} { return $true }
        {$_ -in @('false','0','no','disabled')} { return $false }
        default { return $null }
    }
}
function Get-EPFreshness([string]$Timestamp, $MaxAge, [datetimeoffset]$AsOfUtc) {
    $result=@{Utc='';State='Unknown observation time'}
    # No implicit local timezone, collection time or file modification time fallback.
    if ($Timestamp -notmatch '(?i)(Z|[+-]\d{2}:\d{2})$') { return $result }
    $parsed=[datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse($Timestamp,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) { return $result }
    $result.Utc=$parsed.UtcDateTime.ToString('O'); $age=($AsOfUtc.ToUniversalTime()-$parsed.ToUniversalTime()).TotalHours
    $result.State=if ($age -lt 0) {'Future observation'} elseif ($null -eq $MaxAge) {'Age policy not configured'} elseif ($age -gt $MaxAge) {'Stale'} else {'Current'}
    $result
}
function New-EPIndex([object[]]$Rows, [string]$IdColumn, [string]$TenantKey, [switch]$AllowTenantless) {
    $index=@{}
    foreach ($row in $Rows) {
        $tenant=Get-EPText $row 'TenantKey'; $id=Get-EPText $row $IdColumn
        if (($tenant -and $tenant -ine $TenantKey) -or (-not $tenant -and -not $AllowTenantless)) { throw 'Evidence tenant identity is missing or mismatched.' }
        Assert-EPKey $id $IdColumn; $id=$id.ToLowerInvariant()
        if (-not $index.ContainsKey($id)) { $index[$id]=[Collections.Generic.List[object]]::new() }; $index[$id].Add($row)
    }
    return ,$index
}
function New-EndpointProtectionEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Policy, [object[]]$InventoryRows=@(), [object[]]$SystemRows=@(), [object[]]$DefenderRows=@(), [datetimeoffset]$AsOfUtc=[datetimeoffset]::UtcNow)
    if (-not $Policy.Enabled) { return }
    $inventory=New-EPIndex $InventoryRows 'Device Source ID' $Policy.TenantKey -AllowTenantless
    $system=New-EPIndex $SystemRows 'Device Source ID' $Policy.TenantKey
    $defender=New-EPIndex $DefenderRows 'DeviceId' $Policy.TenantKey
    foreach ($id in @($inventory.Keys | Sort-Object)) {
        if ($inventory[$id].Count -ne 1) { throw 'Duplicate prepared inventory device identity.' }
        $device=$inventory[$id][0]
        $scopes=@($Policy.ScopeRules | Where-Object {$id -in $_.DeviceSourceIds})
        $scopeId='Default'; $requirements=@($Policy.Requirements); $policyReason=''
        if ($scopes.Count -gt 1) { $scopeId='Ambiguous'; $requirements=@(); $policyReason='Overlapping policy scopes' }
        elseif ($scopes.Count -eq 1) { $scopeId=$scopes[0].ScopeId; $requirements=@($scopes[0].Requirements) }
        if (-not $requirements.Count) { $requirements=@(New-EPUnknownRequirement); if (-not $policyReason) {$policyReason='Policy requirements not configured'} }
        foreach ($r in $requirements) {
            $state='Unknown'; $reason=$policyReason; $registration='Unknown'; $source='Not collected'; $fresh=@{Utc='';State='Not evaluated'}
            if ($policyReason) { }
            elseif ($r.ExpectedMode -eq 'NotRequired') { $state='NotApplicable'; $reason='Documented policy exception' }
            elseif ($r.Role -eq 'EDR') { $reason='No qualified EDR onboarding and health source; antivirus registration is not EDR health' }
            elseif ($r.ExpectedMode -eq 'Passive') { $reason='Current sources do not prove Defender passive mode' }
            elseif ($r.Product -ieq 'Microsoft Defender Antivirus') {
                $source='Intune endpoint security Defender agents'
                if (-not $defender.ContainsKey($id)) { $reason='Defender health observation missing' }
                elseif ($defender[$id].Count -ne 1) { $reason='Duplicate Defender observations' }
                else {
                    $d=$defender[$id][0]; $fresh=Get-EPFreshness (Get-EPText $d 'LastReportedDateTime') $Policy.MaxObservationAgeHours $AsOfUtc
                    $partial=Get-EPBoolean (Get-EPValue $d 'IsPartialInventory')
                    $malware=Get-EPBoolean (Get-EPValue $d 'MalwareProtectionEnabled'); $realtime=Get-EPBoolean (Get-EPValue $d 'RealTimeProtectionEnabled'); $overdue=Get-EPBoolean (Get-EPValue $d 'SignatureUpdateOverdue')
                    if ((Get-EPText $d 'CollectionStatus') -ine 'Observed' -or $null -eq $partial -or $partial) { $reason='Incomplete or unqualified collection' }
                    elseif ($fresh.State -ne 'Current') { $reason=$fresh.State }
                    elseif ($null -eq $malware -or $null -eq $realtime -or $null -eq $overdue) { $reason='Defender check values missing or invalid' }
                    elseif ($malware -and $realtime -and -not $overdue) { $state='Healthy'; $reason='Required Defender antivirus checks passed' }
                    else { $state='Gap'; $reason='Required Defender antivirus checks failed' }
                }
            } else {
                $source='Windows antivirus registration (Device System Evidence)'
                $reason='Antivirus registration alone does not prove product health'
                if (-not $system.ContainsKey($id)) { $reason='Registration observation missing' }
                elseif ($system[$id].Count -ne 1) { $reason='Duplicate registration observations' }
                else {
                    $s=$system[$id][0]; $fresh=Get-EPFreshness (Get-EPText $s 'Observation UTC') $Policy.MaxObservationAgeHours $AsOfUtc
                    $expectedName=if ($r.EvidenceProductName) {$r.EvidenceProductName} else {$r.Product}
                    if ((Get-EPText $s 'Local Data Source') -eq 'IntuneRemediation' -and (Get-EPText $s 'Local Data Freshness') -ne 'Fresh') {
                        $fresh.State=Get-EPText $s 'Local Data Freshness'
                        if (-not $fresh.State) {$fresh.State='Unknown'}
                        $reason='Local collection state: '+$fresh.State
                    }
                    elseif ((Get-EPText $s 'Expected AV Name') -ine $expectedName) { $reason='Expected product does not match registration evidence; configure an explicit alias if appropriate' }
                    elseif ($fresh.State -ne 'Current') { $reason=$fresh.State }
                    elseif ((Get-EPText $s 'AV Qualification') -ne 'Evaluable registration observation' -or (Get-EPText $s 'AV Comparison') -eq 'Conflict' -or (Get-EPText $s 'Prepared Inventory Match') -ne 'Matched' -or (Get-EPText $s 'Inventory Match') -ne 'Matched') { $reason='Unqualified or conflicting registration observation' }
                    elseif ((Get-EPText $s 'Expected AV State') -in @('Detected','NotDetected')) { $registration=Get-EPText $s 'Expected AV State' }
                }
            }
            [pscustomobject][ordered]@{
                TenantKey=$Policy.TenantKey; 'Endpoint Protection Key'="$($Policy.TenantKey)|$id|$($r.Role)|$($r.Product)".ToLowerInvariant()
                'Device Source ID'=$id;Country=(Get-EPText $device 'Country'); 'Policy Version'=$Policy.PolicyVersion; 'Scope ID'=$scopeId;Role=$r.Role
                'Expected Product'=$r.Product; 'Expected Mode'=$r.ExpectedMode; 'Evaluation State'=$state; 'Evaluation Reason'=$reason
                'Registration State'=$registration; 'Evidence Source'=$source; 'Observation UTC'=$fresh.Utc; Freshness=$fresh.State; 'Exception Reason'=$r.Reason
            }
        }
    }
}
function Merge-EndpointProtectionContract {
    param([Parameter(Mandatory)]$BaseContract, [Parameter(Mandatory)]$ExtensionContract)
    $copy=$BaseContract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    foreach ($table in $ExtensionContract.tables) {
        if (@($copy.tables | Where-Object {$_.table -ieq $table.table -or $_.file -ieq $table.file}).Count) { throw 'Duplicate endpoint protection contract table or file.' }
        $copy.tables=@($copy.tables)+@($table)
    }
    $copy
}
Export-ModuleMember -Function ConvertTo-EndpointProtectionPolicy,Get-EndpointProtectionPolicyColumns,Get-EndpointProtectionEvidenceColumns,New-EndpointProtectionPolicyRows,New-EndpointProtectionEvidence,Merge-EndpointProtectionContract

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfqNM5lVP/48a8
# bD4dBU/YGQxwDUhwZPoI1qTyOTTYBKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBOHlH80QinCiUYjbmMoQDmToiLBELr0+p/OXgIKL8uhMA0GCSqG
# SIb3DQEBAQUABIIBgCRM3bM8Vvh3UPaWddWJ4GeUaBD4RoKUrg2l60Wi5lTAJwSN
# I2HaEPUP2qKiF0cx7ESLNZSC0Z0Kj/ym1BEh06Ehq8/A7YzyXZQdn528exrCt2Fb
# 6m2527kCBhwcDe9e2EAN5q8qrqaKuTIkjAa6K7LYrBC3jsOF0mCq+DexurHay/Il
# oOZVz2baffzxLmtqNmatfh+SYmzAhAjdGUA8QZgcxEwp5MTrqQee7p7NXOkfBBbu
# lrXCgkBXkXaggqyFFRSdBsYuE9dSClVj4lwHy8BS3N4xT8PW5Ej7BXRUHTrmZOgh
# 6BBGxwIAJx/f+leFOd+vEXz6oKc9Idx4AGzK+CgG9og5CVJXj+9egvSGN+TUUIzG
# 6uGeyQ2s+vhoKbPDcrnxDfRS8sJ+A/bjKPbFYkEhzKOkKtTk5HaLSwO5T1lKjuj0
# KD0e6GR1jbpD6PQ/083bgUY5v/effjrkJYkrFlNdscIUyj6pCCyhVIXYuRZz9rb+
# LECpzuG2UxNa+KTLzqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxNTEy
# MzdaMC8GCSqGSIb3DQEJBDEiBCCctJk3L1oXJUs2qJ0TTsgoMJxCIlj7To5l/ZEQ
# Us4iiDANBgkqhkiG9w0BAQEFAASCAgB8zgxQOj5VK/jVA2322q+0ol9+gKlAWRq/
# zydHLaFzkikhun6doJsfigqaN3VWO9UL6w+gxuiW6MOHQgHJBVmvwHyjSgZ2TPh6
# Md53G+ejuwSqL39yF8IlfwFTx/T2gjQSYOvIDjV2RC4d29ky+bslZQagmuSMD/tg
# dvPP88/ZaEP7g+J3B/OPz16DWPdazQJfYAmIXSlp+Vs65eXEgc2V9c3hkKDBj8FU
# VsDk8dUHfiJvUC33k8Ssd4wkAFoNLmp9zhKLQ6/SfDIbPqtAgJm3w+fIr9Cgu/rr
# jXQf5rExEUGkRaJr0k35V8SS6V2zjnydAIAMyg5vWq0d5W0CKzNuUL5WCz4Q3xxq
# cf7lvRrdwwL5IiJs1P/078EsSBa25h7pHWTWBZ64R3dV5WpBmINhjhlES/7iQgo+
# pVZzOL8cukY4Oeg6cSP6MO/HCacRKWLMlsVx4SbOGZVxCAdYwXdJh6IqWCYlVLkm
# N+nVF0ROnPyJ1d0feILoQoOtw8DOZfM0zf1jOlnawPSN6va6xeA1qxMj5S7k3cQP
# 52kN+HSIeehGYTmH0WwOTaFtaLT7ek5zKE1iRzhwfD+CkvoM5tR4sLCeaRxcsZ/F
# mOk7P520RylupdhElcNnAf5PfQnFA4im5ncDiMuJXX2yRo8NK3qfG02orV//1r41
# xZhqJo++4Q==
# SIG # End signature block
