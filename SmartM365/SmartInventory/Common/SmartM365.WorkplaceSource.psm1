Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-WorkplaceSourceValue {
    param([AllowNull()]$Row, [string[]]$Names, [AllowNull()]$Default = '')
    if ($null -eq $Row) { return $Default }
    foreach ($name in $Names) {
        if ($Row -is [Collections.IDictionary] -and $Row.Contains($name)) { return ,$Row[$name] }
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property) { return ,$property.Value }
        $additional = $Row.PSObject.Properties['AdditionalProperties']
        if ($null -ne $additional -and $additional.Value -and $additional.Value.ContainsKey($name)) { return ,$additional.Value[$name] }
    }
    return $Default
}

function Get-WorkplaceApplicationProductKey {
    param([Parameter(Mandatory)]$Row)
    # JSON tuple avoids separator collisions; version intentionally excluded.
    $parts = @('TenantKey','AppName','AppPublisher','Platform' | ForEach-Object {
        ([string](Get-WorkplaceSourceValue $Row @($_))).Trim().ToLowerInvariant()
    })
    return ConvertTo-Json -InputObject $parts -Compress
}

function Get-WorkplaceApplicationFootprint {
    param([AllowEmptyCollection()][object[]]$Applications,
        [Parameter(Mandatory)][string]$RelationPath)
    $appIndex = @{}; $byProduct = @{}; $byApp = @{}
    foreach ($app in $Applications) {
        $relationScope = [string](Get-WorkplaceSourceValue $app @('RelationCollectionScope'))
        if ($relationScope -and $relationScope -ne 'All') { throw 'Complete All-mode relations are required for product/device footprint.' }
        $tenant = ([string](Get-WorkplaceSourceValue $app @('TenantKey'))).Trim().ToLowerInvariant()
        $id = ([string](Get-WorkplaceSourceValue $app @('AppId'))).Trim().ToLowerInvariant()
        if (-not $tenant -or -not $id) { throw 'Application footprint requires TenantKey and AppId.' }
        $nativeKey = ConvertTo-Json -InputObject @($tenant,$id) -Compress
        if ($appIndex.ContainsKey($nativeKey)) { throw "Duplicate application identity: $id" }
        $productKey = Get-WorkplaceApplicationProductKey $app
        $appIndex[$nativeKey] = $productKey
        $byApp[$nativeKey] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        if (-not $byProduct.ContainsKey($productKey)) {
            $byProduct[$productKey] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        }
    }
    # Stream compact relations; no copy of the multi-million-row source.
    Import-Csv -LiteralPath $RelationPath -ErrorAction Stop | ForEach-Object {
        $tenant = ([string](Get-WorkplaceSourceValue $_ @('TenantKey'))).Trim().ToLowerInvariant()
        $appId = ([string](Get-WorkplaceSourceValue $_ @('AppId'))).Trim().ToLowerInvariant()
        $device = ([string](Get-WorkplaceSourceValue $_ @('DeviceId'))).Trim().ToLowerInvariant()
        if (-not $tenant -or -not $appId -or -not $device) { throw 'Incomplete app/device relation identity.' }
        $nativeKey = ConvertTo-Json -InputObject @($tenant,$appId) -Compress
        if (-not $appIndex.ContainsKey($nativeKey)) { throw "Orphan app/device relation: $appId" }
        [void]$byApp[$nativeKey].Add($device)
        [void]$byProduct[$appIndex[$nativeKey]].Add($device)
    }
    foreach ($app in $Applications) {
        $nativeKey = ConvertTo-Json -InputObject @(([string]$app.TenantKey).Trim().ToLowerInvariant(),([string]$app.AppId).Trim().ToLowerInvariant()) -Compress
        $countText = [string](Get-WorkplaceSourceValue $app @('DeviceCount'))
        $expected = [int64]0
        if (-not [int64]::TryParse($countText,[ref]$expected) -or $expected -lt 0 -or $byApp[$nativeKey].Count -ne $expected) {
            throw "Application relation coverage mismatch: $($app.AppId). Complete All-mode relations are required."
        }
    }
    return $byProduct
}

function ConvertTo-WorkplaceADMembership {
    param([AllowEmptyCollection()][object[]]$Groups,
        [AllowEmptyCollection()][object[]]$DirectoryObjects,
        [Parameter(Mandatory)][string]$CollectedAtUtc)
    $index = @{}
    foreach ($row in $DirectoryObjects) {
        $dn = ([string]$row.DistinguishedName).Trim()
        $tenant = ([string]$row.TenantKey).Trim()
        if (-not $dn -or -not $tenant -or -not $row.ObjectGUID) { throw 'AD identity export is incomplete.' }
        $key = ConvertTo-Json -InputObject @($tenant.ToLowerInvariant(),$dn.ToLowerInvariant()) -Compress
        if ($index.ContainsKey($key)) { throw 'Duplicate tenant/DN in AD native identity export.' }
        $index[$key] = $row
    }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($group in $Groups) {
        if (-not $group.TenantKey -or -not $group.ObjectGUID) { throw 'AD group identity is incomplete.' }
        $membersJson = [string](Get-WorkplaceSourceValue $group @('MembersJson'))
        $members = if ($membersJson) { @($membersJson | ConvertFrom-Json -ErrorAction Stop) } else { @(([string]$group.Members).Split(';') | Where-Object { $_.Trim() }) }
        foreach ($memberDn in $members) {
            $dn = $memberDn.Trim()
            $rowKey = ConvertTo-Json -InputObject @([string]$group.TenantKey,[string]$group.ObjectGUID,$dn) -Compress
            if (-not $seen.Add($rowKey)) { continue }
            $key = ConvertTo-Json -InputObject @(([string]$group.TenantKey).Trim().ToLowerInvariant(),$dn.ToLowerInvariant()) -Compress
            $member = $index[$key]
            [pscustomobject][ordered]@{
                TenantKey = $group.TenantKey; GroupObjectGUID = $group.ObjectGUID
                GroupSID = $group.objectSid; MemberDistinguishedName = $dn
                MemberObjectGUID = if ($member) { $member.ObjectGUID } else { '' }
                MemberSID = if ($member) { $member.ObjectSID } else { '' }
                MemberObjectClass = if ($member) { $member.ObjectClass } else { '' }
                MembershipKind = 'Direct'
                ResolutionStatus = if ($member) { 'Resolved' } else { 'UnresolvedOrExternal' }
                CollectedAtUtc = $CollectedAtUtc
            }
        }
    }
    # Primary groups are absent from the AD group's member attribute.
    $groupBySid = @{}
    foreach ($group in $Groups) {
        if ($group.objectSid) { $groupBySid[([string]$group.TenantKey)+'|'+([string]$group.objectSid)] = $group }
    }
    foreach ($member in $DirectoryObjects) {
        if (-not $member.PrimaryGroupID -or -not $member.ObjectSID) { continue }
        $sid = ([string]$member.ObjectSID) -replace '-\d+$',('-'+[string]$member.PrimaryGroupID)
        $group = $groupBySid[([string]$member.TenantKey)+'|'+$sid]
        [pscustomobject][ordered]@{
            TenantKey = $member.TenantKey
            GroupObjectGUID = if ($group) { $group.ObjectGUID } else { '' }
            GroupSID = $sid; MemberDistinguishedName = $member.DistinguishedName
            MemberObjectGUID = $member.ObjectGUID; MemberSID = $member.ObjectSID
            MemberObjectClass = $member.ObjectClass; MembershipKind = 'PrimaryGroup'
            ResolutionStatus = if ($group) { 'Resolved' } else { 'UnresolvedPrimaryGroup' }
            CollectedAtUtc = $CollectedAtUtc
        }
    }
}

function Get-WorkplaceGraphCollection {
    param([Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][scriptblock]$Invoker)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $next = $Uri
    while ($next) {
        $address = [uri]$next
        if ($address.Scheme -ne 'https' -or $address.Host -ne 'graph.microsoft.com' -or $address.UserInfo -or $address.Port -ne 443) { throw 'Untrusted Graph pagination address.' }
        if (-not $seen.Add($next)) { throw 'Cyclic Graph pagination.' }
        $response = & $Invoker $next
        $values = Get-WorkplaceSourceValue $response @('value') $null
        if ($null -eq $values -or $values -is [string] -or $values -isnot [Collections.IEnumerable]) { throw 'Malformed Graph collection response.' }
        foreach ($value in $values) { $value }
        $next = [string](Get-WorkplaceSourceValue $response @('@odata.nextLink'))
    }
}

function ConvertTo-WorkplacePolicyEvidence {
    param([Parameter(Mandatory)]$Policy,[Parameter(Mandatory)][string]$Family,
        [string]$RunId,[string]$CollectedAtUtc)
    $id = [string](Get-WorkplaceSourceValue $Policy @('id'))
    if (-not $id) { throw 'Policy native identity is missing.' }
    [pscustomobject][ordered]@{
        PolicyId = $id; PolicyFamily = $Family
        DisplayName = Get-WorkplaceSourceValue $Policy @('name','displayName')
        Description = Get-WorkplaceSourceValue $Policy @('description')
        Platforms = Get-WorkplaceSourceValue $Policy @('platforms')
        Technologies = Get-WorkplaceSourceValue $Policy @('technologies')
        FeatureUpdateVersion = Get-WorkplaceSourceValue $Policy @('featureUpdateVersion')
        CreatedDateTime = Get-WorkplaceSourceValue $Policy @('createdDateTime')
        LastModifiedDateTime = Get-WorkplaceSourceValue $Policy @('lastModifiedDateTime')
        NativeEvidenceJson = ConvertTo-Json -InputObject $Policy -Depth 50 -Compress
        AssignmentCollectionStatus = 'Collected'; RunId = $RunId; CollectedAtUtc = $CollectedAtUtc
    }
}

Export-ModuleMember -Function Get-WorkplaceSourceValue,Get-WorkplaceApplicationProductKey,Get-WorkplaceApplicationFootprint,ConvertTo-WorkplaceADMembership,Get-WorkplaceGraphCollection,ConvertTo-WorkplacePolicyEvidence

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD96eF70GSl2E+j
# a8SRyMaOLYoh13KL8x0/9XTXENeTQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIC6uXFGUyRtnPfdhX4zuaslYaxprEIJf4+ZA76z1Iz/MA0GCSqG
# SIb3DQEBAQUABIIBgA4xuL08RENM+4usX+YML38ddN9NrB4i2Eg1SJ+o5HT5ygG6
# sCZjGEB9q/4sqbYPA1f9UbvQoqEHfQEOHB3J1qJnYLuVA8wDFpglSNDBvYnpEAuc
# m48fSWafjrNnU1fYFDpgu7ifeQmcJqS80mB0IRLFarT23eF6okj4u2u0T338Ql/Z
# sNJ3jrE4bUy1uly9r4nSIhVa2CXD15OE09ljdwml9n+JHusKKXr/RZaBTvCXjHa1
# I6pUDXxti/+RH0XOpedfi1nbefHXRB7Vx9Nvn0z7csM2whL3ewYtdHAfLSwHTFGA
# qnxIQ0bkU/lSZFnYMelK1ENcc3+RKc9n+7dvr6So+slmmrofwiV+l5azZ4FGCXeo
# W5YoaN1vJV6ojUyNVHeEpeZKzWHrgImLRbItR0yLZCzmNz/bBBiWEa3suyEx8s/y
# +aTRJENLb8ol6Uyj+LNz74vjaJebEDfBE5L/PrY5wX+3Pl5xfLgiIMFJo9WS2B+R
# BmBmgZ4uPqe7uh1eg6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDIxOTQx
# MDJaMC8GCSqGSIb3DQEJBDEiBCAdiBujHX+DJrJX0xSg6GpUB5e8Gq4sy3HlFOjp
# OnDtUjANBgkqhkiG9w0BAQEFAASCAgCx/VJyD2Vt64SEXidWb+xsPF1xiFxIWCZt
# c8JVj9j5y3cFKoq06E7AKJoR35PX/Lu9mf4YdxfIC/WeQ2ieBuS1xXE9HSsK7MMu
# NQrSMFcy9WQ7kvX4hJsCF1/mvwuk2kChHx8Ddd2hocePFsfOxYxxOp54GLRlj3Ng
# r6XVNatyJiyoTNWcyyC8wrOBlf5NigX64+KkP3ATRrrw+JLNatosYBiMkmbt+2n1
# l62NElw0jV7IMH4VF5mFJSjI9izMj5vgj1iyK5lFDjIIi1T4K+zFchKcc77sEaxA
# w7GabCgcKP21Mw8tToScbAbRWmDxwcbUO2JFvidN5sCtEY8uTNGjV5zG97EEbWpT
# 73yRXTG1sS22wmX/jthblTSgcpLqTtIu3xLrcoyWHSCojywNzFzt9Q0MGS/xw0lu
# qFbSRhaw6G+cVuJtiUe6PxVLkIS/Ermd1y8K0gbMNhP5D/n+oUczq94VF4y1k4wL
# Syqqq+oKCQ8TDQyXbZMP/FFPfJhAZ7/z4B3PPJLh7oXk0LETvr4UVSGAJ+r9cq0V
# R1p34vgUOr7oeKhox1wQjlmc4X5zUu1jVQ4wtDwti8vtoT62v07h5Z071XN1FtjD
# CjQfo3dX0Jc74lgGtGftPcTgxRAoG441Y7iK5375PZxouCYif6go3M5VK0i7itVj
# jOUpZ6KGZw==
# SIG # End signature block
