# Private stable V1 context projection. No inferred location, ownership or hardware.
function Read-CIOrganizationReference {
    param([string]$Path,$Identity,$Columns)
    $result=@{Items=@{};Hash='';Path=$Path}
    if (-not $Path) { return $result }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Missing organization reference file.' }
    $result.Hash=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    Test-CIExactHeader -Path $Path -Columns $Columns
    Import-Csv -LiteralPath $Path | ForEach-Object {
        $row=$_
        foreach ($field in $Identity.Keys) { if ([string]$row.$field -cne [string]$Identity[$field]) { throw 'Organization reference tenant mismatch.' } }
        $id=[string]$row.ReferenceId
        if ([string]::IsNullOrWhiteSpace($id) -or -not $id.StartsWith($Identity.TenantKey+'|',[StringComparison]::Ordinal) -or
            [string]::IsNullOrWhiteSpace([string]$row.Name)) { throw 'Invalid organization reference identity or name.' }
        if ($row.ReferenceType -cnotin @('Country','Entity','Site')) { throw 'Unknown organization reference type.' }
        if ($result.Items.ContainsKey($id)) { throw 'Duplicate organization reference ID.' }
        $result.Items[$id]=@{Id=$id;Type=[string]$row.ReferenceType;Name=[string]$row.Name;Parent=[string]$row.ParentReferenceId}
    }
    foreach ($item in $result.Items.Values) {
        if ($item.Type -ceq 'Country') {
            if ($item.Parent) { throw 'Country reference cannot have a parent.' }
        } else {
            $expected=if ($item.Type -ceq 'Entity') {'Country'} else {'Entity'}
            if (-not $result.Items.ContainsKey($item.Parent) -or $result.Items[$item.Parent].Type -cne $expected) { throw 'Missing or wrong-type organization parent.' }
        }
    }
    return $result
}

function ConvertTo-CIQualifiedTime {
    param([AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @{Utc='';Status='Missing'} }
    if ($Value -cnotmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$') { return @{Utc='';Status='NotOffsetQualified'} }
    $parsed=[DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse($Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) { return @{Utc='';Status='Invalid'} }
    return @{Utc=$parsed.UtcDateTime.ToString('o',[Globalization.CultureInfo]::InvariantCulture);Status='UTCQualified'}
}

function Initialize-CIContext {
    param($Contract,$Identity,[string]$Stage,[bool]$ValidateOnly)
    $names=@('CMDB_CIUserContext.csv','CMDB_CIDeviceContext.csv','CMDB_CIDeviceSourceContext.csv','CMDB_CIOrganization.csv')
    if ($Contract.channel -cne 'stable' -or $Contract.contractVersion -cne '1.0.0' -or
        (@($Contract.tables.name) -join ',') -cne ($names -join ',')) { throw 'Unsupported stable V1 context contract.' }
    $result=@{Identity=$Identity;Columns=@{};Writers=@{};Counts=[ordered]@{};AssociationCounts=@{};DepartmentMissing=0;ActivityQualification=@{};EnrollmentQualification=@{}}
    try {
        foreach ($table in $Contract.tables) {
            $name=[string]$table.name; $result.Columns[$name]=@($table.columns); $result.Counts[$name]=0
            if (-not $ValidateOnly) {
                $writer=[IO.StreamWriter]::new((Join-Path $Stage $name),$false,[Text.UTF8Encoding]::new($true))
                $result.Writers[$name]=$writer
                $writer.WriteLine($table.columns -join ',')
            }
        }
    } catch { Close-CIContext -Context $result; throw }
    return $result
}

function Close-CIContext {
    param($Context)
    if ($null -ne $Context) {
        foreach ($writer in $Context.Writers.Values) { $writer.Dispose() }
        $Context.Writers.Clear()
    }
}

function Write-CIContextRow {
    param($Context,[string]$File,$Fields)
    $record=[ordered]@{}
    foreach ($field in $Context.Identity.Keys) { $record[$field]=$Context.Identity[$field] }
    foreach ($field in $Fields.Keys) { $record[$field]=$Fields[$field] }
    if (($record.Keys -join ',') -cne ($Context.Columns[$File] -join ',')) { throw 'Context row does not match its column contract.' }
    if ($Context.Writers.ContainsKey($File)) { $Context.Writers[$File].WriteLine((@([pscustomobject]$record | ConvertTo-Csv -NoTypeInformation)[1])) }
    $Context.Counts[$File]++
}

function Write-CIDeviceSourceContext {
    param($Context,$SourceRow,$Mapping)
    $intune=$Mapping.SourceSystem -ceq 'MicrosoftIntune'
    $activity=if ($intune) {[string]$SourceRow.LastSyncDateTime} else {[string]$SourceRow.ApproximateLastSignInDateTime}
    $enrollment=if ($intune) {[string]$SourceRow.EnrolledDateTime} else {''}
    $qualifiedActivity=ConvertTo-CIQualifiedTime $activity
    $qualifiedEnrollment=if ($intune) {ConvertTo-CIQualifiedTime $enrollment} else {@{Utc='';Status='NotApplicable'}}
    $activityKey=$Mapping.SourceSystem+':'+$qualifiedActivity.Status
    $enrollmentKey=$Mapping.SourceSystem+':'+$qualifiedEnrollment.Status
    $Context.ActivityQualification[$activityKey]++; $Context.EnrollmentQualification[$enrollmentKey]++
    Write-CIContextRow -Context $Context -File 'CMDB_CIDeviceSourceContext.csv' -Fields ([ordered]@{
        CI_ID=$Mapping.CI_ID;SourceSystem=$Mapping.SourceSystem;SourceID=$Mapping.SourceID;SourceIDKind=$Mapping.SourceIDKind
        ManagementAgent=$(if ($intune) {[string]$SourceRow.ManagementAgent} else {''})
        EnrollmentType=$(if ($intune) {[string]$SourceRow.DeviceEnrollmentType} else {''})
        EnrollmentRaw=$enrollment;EnrollmentUtcDateTime=$qualifiedEnrollment.Utc;EnrollmentStatus=$qualifiedEnrollment.Status
        ActivityKind=$(if ($intune) {'IntuneSync'} else {'ApproximateDeviceSignIn'})
        ActivityRaw=$activity;ActivityUtcDateTime=$qualifiedActivity.Utc;ActivityStatus=$qualifiedActivity.Status
        SourceCollectedDateTime=$Mapping.SourceCollectedDateTime;MappingStatus=$Mapping.MappingStatus
    })
}

function Write-CICuratedContext {
    param($Context,[string]$InputRoot,$CIIndex,$Governance,$Organization)
    $users=@{}
    Import-Csv -LiteralPath (Join-Path $InputRoot 'CMDB_Users.csv') | ForEach-Object {
        $row=$_; $sid=([string]$row.SourceUserId).Trim()
        if ($sid) {
            if ($users.ContainsKey($sid)) { throw 'Ambiguous associated-user source identity.' }
            $users[$sid]=@{Id=[string]$row.CmdbUserId;Department=[string]$row.Department;JobTitle=[string]$row.JobTitle;Date=[string]$row.SourceCollectedDateTime}
        }
        $present=-not [string]::IsNullOrWhiteSpace([string]$row.Department)
        if (-not $present) { $Context.DepartmentMissing++ }
        Write-CIContextRow -Context $Context -File 'CMDB_CIUserContext.csv' -Fields ([ordered]@{
            CI_ID=[string]$row.CmdbUserId;Department=[string]$row.Department;JobTitle=[string]$row.JobTitle
            DepartmentStatus=$(if ($present) {'Present'} else {'Missing'});SourceSystem=[string]$row.SourceSystem
            SourceUserId=[string]$row.SourceUserId;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime
        })
    }
    Import-Csv -LiteralPath (Join-Path $InputRoot 'CMDB_Devices.csv') | ForEach-Object {
        $row=$_; $sid=([string]$row.PrimaryUserId).Trim(); $user=$null
        $status=if (-not $sid) {'NotProvided'} elseif ($users.ContainsKey($sid)) {$user=$users[$sid];'Resolved'} else {'Unresolved'}
        $Context.AssociationCounts[$status]++
        Write-CIContextRow -Context $Context -File 'CMDB_CIDeviceContext.csv' -Fields ([ordered]@{
            CI_ID=[string]$row.CmdbDeviceId;DeviceName=[string]$row.DeviceName;OperatingSystem=[string]$row.OperatingSystem
            OperatingSystemVersion=[string]$row.OperatingSystemVersion;Ownership=[string]$row.Ownership
            ComplianceState=[string]$row.ComplianceState;ManagementState=[string]$row.ManagementState
            PrimaryUserSourceId=[string]$row.PrimaryUserId;AssociatedUserCI_ID=$(if ($user) {$user.Id} else {''});AssociationStatus=$status
            AssociatedUserDepartment=$(if ($user) {$user.Department} else {''});AssociatedUserJobTitle=$(if ($user) {$user.JobTitle} else {''})
            AssociatedUserCollectedDateTime=$(if ($user) {$user.Date} else {''});LastSyncRaw=[string]$row.LastSyncDateTime
            SourceSystem=[string]$row.SourceSystem;SourceCollectedDateTime=[string]$row.SourceCollectedDateTime
        })
    }
    foreach ($refId in $Governance.OrganizationReferences.Keys) {
        if (-not $Organization.Items.ContainsKey($refId)) { throw 'Governance organization reference is missing or foreign.' }
    }
    foreach ($id in $Governance.Items.Keys) {
        $refId=[string]$Governance.Items[$id].OrganizationRefId
        if (-not $refId) { continue }
        if (-not $CIIndex.ContainsKey($id)) { throw 'Organization assignment CI is missing.' }
        $item=$Organization.Items[$refId]; $path=@{}
        $cursor=$item
        # Parent types were already checked; at most Site -> Entity -> Country.
        for ($depth=0; $null -ne $cursor -and $depth -lt 3; $depth++) {
            $path[$cursor.Type]=$cursor
            $cursor=if ($cursor.Parent) {$Organization.Items[$cursor.Parent]} else {$null}
        }
        $fields=[ordered]@{CI_ID=$id;OrganizationRefId=$item.Id;ReferenceType=$item.Type}
        foreach ($type in @('Country','Entity','Site')) {
            $fields[$type+'RefId']=if ($path.ContainsKey($type)) {$path[$type].Id} else {''}
            $fields[$type+'Name']=if ($path.ContainsKey($type)) {$path[$type].Name} else {''}
        }
        $fields.AssignmentStatus='Declared'
        Write-CIContextRow -Context $Context -File 'CMDB_CIOrganization.csv' -Fields $fields
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBALflCvGEHmL/R
# 9+OPNV50HlfYVcpsznfYABzcOmuwDKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICpjQS1xNS3SEBw3F4Iy4rb2LdP+ylphkXezjfsmxIpvMA0GCSqG
# SIb3DQEBAQUABIIBgHDxLrU9Tq7oLYxyUBUMZVHdqH3Z41c6hh5Rd/QvWcYuO9CV
# R7WZ7Q3idA11gAFU4wHyBtgKUMi6W3UD+Dl93lcK53/oVLue5dPsc8ZvlIp0NciL
# OywY5Faps1MwvZDCe4RopvIqZ3pNnjpnEJHhfgQG6SJDFOFd8R1YEW3kMvBHj4M9
# T0grX31WBO/YbezwZF0SD3hNI2eVN8jalppsobJ3nlMKIM+MntJL45jhar/ip7gV
# lczrlfIaZ2CPOHpueuZon7mYKOhfDgDCUBhKWeu8gvOvXxW2VqzFLwJC1iiDv33P
# Kfn7Xbn1wFCkA1bGKXvAFsOPqaTW47/OqHcooSmlB9dVWEHLv2rtmc86qtW8PYWi
# AjL00dasd5ZxWycaRy+ldSQS8mAFckQ9wzQGrnvq4VLb32qzKfs03b95zke06Ogu
# KKDMd39urFvrBSeVcvgcjpjN/H27lRIoNlg2MfCI1sg7ijP5MGu5zfumUZ3vJJu1
# kpNtXAZpuvj+8EvBgqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MTZaMC8GCSqGSIb3DQEJBDEiBCCq/t5txSM0r18x5CUu4somAJKpaMezKJZI9rPT
# KhRv0jANBgkqhkiG9w0BAQEFAASCAgAlgfKaZIkdiLeG1vTrLQT67T1FveMfwHdm
# ogR5J5sTaFZf7xDCCUnwSCI5tDa35tKQzFXIWevL91hA6FR6RiQifjpj1QBUEpv/
# dLk5r9unQr7ALrY3gkFnS4xcxo4bbUDGbZY5AvN+xuJgdHMwzhOLldzdk6HuV60/
# QCqoefkOaLROalezCMhXFPHt+kbJyuLpc/HlTAJbZXGt0sGZoHAriDI8PlzxWBIP
# T9jsONPrD/lwhqMpuMHzJirM0r8ChNoIAyCLHGkK+LDyauTlIPattaOvpXfl1Our
# 3PZCJB+NcspoNF96w3ouqX87ChJmY+0dnct5M53tuINYUcWtGzO0fjt4HMetDpk3
# Xd9WxlhPKKGTadKPacjAE6uIDmEjhrUJKw/PHriE28SNcqs+nuvyQAahNmBBEfJ2
# tK69MLLWWWyGbtryRekcYRpUWcbdY2NHgTl0oEboQu8rXqGijavl/08NsvOQnuBM
# +Bm8cbCxYnhatRkBQTp4QrSHHYkmuCVSdsnZfeXdkKQbNEH/ZLhnPFJrPGPQhID6
# tScsA06sH46sGay4VgGBvPEkBn60SohFVNaZAIHH4iJcLKn3BpywK7avy3EfQr+k
# G2X5lRQAS2WkuIacUS5/ym7P3TxL11NrRDZA3LLt+Mhhoq6Wj5/l6zz3udnKwKML
# WJ05er9gnA==
# SIG # End signature block
