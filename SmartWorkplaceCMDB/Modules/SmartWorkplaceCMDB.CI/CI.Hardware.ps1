# Optional stable V1 projection. Source-reported hardware never defines CI identity.
function Export-CIHardwareContext {
    param([string]$Path, $Identity, $Mapping, [string]$ProjectRoot, [string]$Stage, [bool]$ValidateOnly)
    $hashes=@{}
    $schemaPath=Join-Path $ProjectRoot 'Schema\SmartWorkplaceCMDB.ci.hardware.json'
    $rawSchemaPath=Join-Path $ProjectRoot 'Schema\SmartWorkplaceCMDB.hardware.tables.json'
    $statePath=$Path+'.status.json'
    foreach ($file in @($Path,$statePath,$schemaPath,$rawSchemaPath)) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw 'Missing hardware input, state or contract.' }
        $hashes[$file]=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    }
    $schema=Get-Content -LiteralPath $schemaPath -Raw | ConvertFrom-Json
    $rawSchema=Get-Content -LiteralPath $rawSchemaPath -Raw | ConvertFrom-Json
    foreach ($definition in @($schema,$rawSchema)) {
        if ($definition.channel -cne 'stable' -or $definition.contractVersion -cne '1.0.0') { throw 'Hardware contracts must be the frozen V1 stable contracts.' }
    }
    $rawColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','SourceSystem','ManagedDeviceId','AzureAdDeviceId','SerialNumber','SerialNumberStatus','Manufacturer','ManufacturerStatus','Model','ModelStatus','TotalStorageSpaceInBytes','StorageStatus','SourceCollectedDateTime')
    $columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','CI_ID','SourceSystem','ManagedDeviceId','AzureAdDeviceId','SerialNumber','SerialNumberStatus','Manufacturer','ManufacturerStatus','Model','ModelStatus','TotalStorageSpaceInBytes','StorageStatus','HardwareCollectedDateTime','InventoryCollectedDateTime','CollectionCoverage','CollectionMode')
    if ($schema.name -cne 'CMDB_CIDeviceHardware.csv' -or ($schema.columns -join ',') -cne ($columns -join ',') -or
        @($rawSchema.tables).Count -ne 1 -or ($rawSchema.tables[0].columns -join ',') -cne ($rawColumns -join ',')) { throw 'Unsupported hardware column contract.' }
    Test-CIExactHeader -Path $Path -Columns $rawColumns
    $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    foreach ($field in $Identity.Keys) {
        if ([string]$state.$field -cne [string]$Identity[$field]) { throw 'Hardware state tenant mismatch.' }
    }
    if ($state.Version -ne 1 -or $state.SourceName -cne [IO.Path]::GetFileName($Path) -or
        [string]::IsNullOrWhiteSpace([string]$state.RunId) -or $state.Status -cne 'Completed' -or
        [string]$state.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $state.SHA256 -ine $hashes[$Path]) { throw 'Hardware state is not a matching Completed snapshot.' }
    $expectedCount=[long]0; $limit=[long]0
    if ([string]$state.RowCount -cnotmatch '^\d+$' -or -not [long]::TryParse([string]$state.RowCount,[ref]$expectedCount) -or
        [string]$state.MaxItems -cnotmatch '^\d+$' -or -not [long]::TryParse([string]$state.MaxItems,[ref]$limit)) { throw 'Invalid hardware state row count or limit.' }
    if (($state.Mode -ceq 'Fixture' -and $state.Coverage -cnotin @('Fixture','Bounded')) -or
        ($state.Mode -ceq 'Live' -and $state.Coverage -cnotin @('Complete','Bounded')) -or
        $state.Mode -cnotin @('Fixture','Live') -or
        ($state.Coverage -ceq 'Bounded' -and ($limit -le 0 -or $expectedCount -gt $limit)) -or
        ($state.Coverage -cne 'Bounded' -and $limit -ne 0)) { throw 'Invalid hardware coverage or mode.' }
    $times=@{}
    foreach ($field in @('StartedUtc','CompletedUtc')) {
        $value=$state.$field
        $text=if ($value -is [datetime] -or $value -is [datetimeoffset]) { $value.ToString('o',[Globalization.CultureInfo]::InvariantCulture) } else { [string]$value }
        $qualified=ConvertTo-CIQualifiedTime -Value $text
        if ($qualified.Status -cne 'UTCQualified') { throw 'Invalid hardware state timestamp.' }
        $times[$field]=[datetimeoffset]::Parse($qualified.Utc,[Globalization.CultureInfo]::InvariantCulture)
    }
    $validatedUtc=[datetimeoffset]::UtcNow
    if ($times.StartedUtc -gt $times.CompletedUtc -or $times.CompletedUtc -gt $validatedUtc.AddMinutes(5)) { throw 'Hardware state timestamps are reversed or in the future.' }
    $seen=@{}; $ciSeen=@{}; $serials=@{}; $missing=[ordered]@{SerialNumber=0;Manufacturer=0;Model=0;Storage=0}
    $count=0; $zero=0; $writer=$null; $rowDate=''
    try {
        if (-not $ValidateOnly) {
            $writer=[IO.StreamWriter]::new((Join-Path $Stage $schema.name),$false,[Text.UTF8Encoding]::new($true))
            $writer.WriteLine($columns -join ',')
        }
        Import-Csv -LiteralPath $Path | ForEach-Object {
            $row=$_
            foreach ($field in $Identity.Keys) { if ([string]$row.$field -cne [string]$Identity[$field]) { throw 'Hardware row tenant mismatch.' } }
            if ($row.SourceSystem -cne 'MicrosoftIntune') { throw 'Unexpected hardware SourceSystem.' }
            $native=([string]$row.ManagedDeviceId).Trim().ToLowerInvariant()
            if (-not $native -or $seen.ContainsKey($native)) { throw 'Missing or duplicate hardware native ID.' }
            if (-not $Mapping.ContainsKey($native)) { throw 'Unmapped hardware native ID.' }
            $source=$Mapping[$native]
            $correlation=([string]$row.AzureAdDeviceId).Trim().ToLowerInvariant()
            if (-not $correlation -or $correlation -ceq '00000000-0000-0000-0000-000000000000') { $correlation='intune:'+$native }
            if ($correlation -cne $source.Correlation) { throw 'Hardware correlation differs from mapped inventory.' }
            foreach ($attribute in @('SerialNumber','Manufacturer','Model')) {
                $status=[string]$row.PSObject.Properties[$attribute+'Status'].Value; $value=[string]$row.$attribute
                if (($status -ceq 'Missing' -and $value -cne '') -or
                    ($status -ceq 'Reported' -and [string]::IsNullOrWhiteSpace($value)) -or
                    $status -cnotin @('Missing','Reported')) { throw 'Inconsistent hardware text qualification.' }
                if ($status -ceq 'Missing') { $missing[$attribute]++ }
            }
            $capacity=[long]0; $storage=[string]$row.TotalStorageSpaceInBytes
            if ($row.StorageStatus -ceq 'Missing') {
                if ($storage -cne '') { throw 'Missing hardware storage has a value.' }
                $missing.Storage++
            } else {
                if ($storage -cnotmatch '^\d+$' -or -not [long]::TryParse($storage,[ref]$capacity) -or
                    ($row.StorageStatus -ceq 'Reported' -and $capacity -le 0) -or
                    ($row.StorageStatus -ceq 'ZeroReported' -and $capacity -ne 0) -or
                    $row.StorageStatus -cnotin @('Reported','ZeroReported')) { throw 'Invalid hardware storage value or qualification.' }
                if ($capacity -eq 0) { $zero++ }
            }
            $date=ConvertTo-CIQualifiedTime -Value ([string]$row.SourceCollectedDateTime)
            if ($date.Status -cne 'UTCQualified' -or
                [datetimeoffset]::Parse($date.Utc,[Globalization.CultureInfo]::InvariantCulture) -gt $times.CompletedUtc -or
                ($count -gt 0 -and $date.Utc -cne $rowDate)) { throw 'Invalid or mixed hardware collection timestamp.' }
            $rowDate=$date.Utc
            $seen[$native]=$true; $ciSeen[$source.CI_ID]=$true; $count++
            if ($row.SerialNumberStatus -ceq 'Reported') {
                $serial=([string]$row.SerialNumber).Trim().ToLowerInvariant()
                if (-not $serials.ContainsKey($serial)) { $serials[$serial]=0 }; $serials[$serial]++
            }
            $record=[ordered]@{}
            foreach ($field in $Identity.Keys) { $record[$field]=$Identity[$field] }
            $record.CI_ID=$source.CI_ID
            foreach ($field in $rawColumns[4..14]) { $record[$field]=[string]$row.$field }
            $record.HardwareCollectedDateTime=[string]$row.SourceCollectedDateTime
            $record.InventoryCollectedDateTime=$source.Collected
            $record.CollectionCoverage=[string]$state.Coverage; $record.CollectionMode=[string]$state.Mode
            if ($null -ne $writer) { $writer.WriteLine((@([pscustomobject]$record | ConvertTo-Csv -NoTypeInformation)[1])) }
        }
        if ($count -ne $expectedCount) { throw 'Hardware state row count mismatch.' }
    } finally { if ($null -ne $writer) { $writer.Dispose() } }
    foreach ($file in $hashes.Keys) { if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -cne $hashes[$file]) { throw 'Hardware input or contract changed during build.' } }
    return [ordered]@{
        Status='Validated';ContractVersion=$schema.contractVersion;RowCount=$count;CIsWithHardware=$ciSeen.Count
        InventoryNativeDeviceCount=$Mapping.Count;InventoryNativeDevicesWithoutHardware=$Mapping.Count-$count
        MissingAttributes=$missing;ZeroReportedStorageCount=$zero
        DuplicateSerialValueCount=@($serials.Values | Where-Object { $_ -gt 1 }).Count
        Coverage=[string]$state.Coverage;Mode=[string]$state.Mode;RunId=[string]$state.RunId
        HardwareCollectedUtc=$rowDate;CompletedUtc=$times.CompletedUtc.UtcDateTime.ToString('o')
        ValidatedUtc=$validatedUtc.UtcDateTime.ToString('o');AttributeFreshness='NotAssessed';InputHashes=$hashes
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPKj9glrE+IYOh
# BqEb+IOygR8CBIe6guofIAjfJOFSZqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO6IJf3knz13UmELFEiO483q7CmBKhSZDz2KzemOsKJIMA0GCSqG
# SIb3DQEBAQUABIIBgHq3rZeuQGv89SEbQh1Eny9Klrx+6zkUOShOnG9GtH/1v4iu
# fgMnbjzBcj0ucnVtDd55ee0tfz+fVuzSjlOPXTGDeTZmELdZi0mf4ToPeCxfoQOa
# Z5niNuP4L6asul3O015yK+ruPsJq3lsCYX01PDtsn+tJIx0kMHoMSk3xK9L1BNgQ
# ZXd1DFus+Kq3pEacZmXuo42qvTUPkAs1lUOJD70+i4t3fdvvPs1d4guy5o8er5Mu
# qlh8pYA9P7GPacYsdqsOd1wM2sJwpW+CVaJ5/Zy0dRs3iXoRKsciCvdpnknXWawY
# ILFVvJrLm1j27ZBx82nholgGv0CG/E7HhSU94a4mxuEdE5YljFoEILsJtQgHxegk
# srzUUCXRe30YdCtBgS/+0GbfZ+rTpxAN49kXTmSK5QQGAMT3gqdXkfoQJt7iWEzL
# vHtl96SykPkQZfkNuoe+mIqd84xTxio4fUp73QZmW6/7FwT57o56EXQAEzS7epV2
# HF2VnC8eFfG+7lZeAKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTExNTE1
# MTZaMC8GCSqGSIb3DQEJBDEiBCCKnN6EDcPEV1RTtCSWPgXQqnP+AfoK1MCUyUjp
# yOQRVjANBgkqhkiG9w0BAQEFAASCAgB2U8tBZa3OP0zqwy+BJ5dbcXEoxIY6Cu4r
# pqmwW1j4/38/BGiagmmVsUVD75Yn1Jsa7hcplc3iqM4ur/a3IdBBk+N45iPLGJye
# H31kigfj3D/V8OEc0b3ndJTf1VnKQeYbr5U+mnGInr6tqNNSTjS317bqn0Y6xoCf
# TKO1Xj1W/iAm2H9hJVbs5SkSrXcomR9FdkOlNv1lKigVVD1DoENAIcfICn+cGWrf
# 2QVsUZN1aTMoTeBjmaEF8TgXehmGCx2gXd4ZyprUwL41+42KJbg2yu3p4IG8MOJV
# VrnzVlhMXt68m6QLrnUWZ7ju+90QghGuQg3cD6nkcD+g9CRnxPO0oFJjt+sAtifX
# L+5ShFgAoEXOioLNL5FMn/ugRMsCy3QgjP7+5SX6hBq+oZ6ztuITAHwGDuL0fP4U
# j3xuC7BbrdWmZfGNK23dGrmLbZDGIfGAeN/iz9fIX1yZ86bT0+YwgeFoiwXZy6fn
# rdCLs1OcuipJX+cYL+PLzhZX513ATiWl03g3+qFJm6hagaM257i4gM6JhYk+gpXf
# NTwSbiYD8SQMbQ5hzeMADh9j8uUBihS7v/J7+CyhHBCLTJPLZOwplYYTfmloEqYn
# 0U4gBo7WRH33k/PGeRKGf5mrK2F5oAsIFkGtQhIM7Csko9CVswHp114M8LSpRGzc
# 9+UCI6lolg==
# SIG # End signature block
