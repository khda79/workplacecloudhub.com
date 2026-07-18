Set-StrictMode -Version 2.0

function Test-SmartFinOpsExcludedCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FileName)

    foreach ($pattern in @($script:SmartFinOpsExcludedFileNamePatterns)) {
        if ([string]::IsNullOrWhiteSpace([string]$pattern)) { continue }
        if ($FileName -match [regex]::Escape([string]$pattern)) { return $true }
    }
    return $false
}

function Get-SmartFinOpsSourceContract {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Key)

    $matches = @($script:SmartFinOpsSourceContract.Sources | Where-Object { $_.Key -eq $Key })
    if ($matches.Count -ne 1) {
        throw "Expected exactly one SmartFinOps source contract for key '$Key'; found $($matches.Count)."
    }
    return $matches[0]
}

function Get-SmartFinOpsCsvHeaderColumns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($firstLine)) { return @() }

    $commaCount = @($firstLine.ToCharArray() | Where-Object { $_ -eq ',' }).Count
    $semicolonCount = @($firstLine.ToCharArray() | Where-Object { $_ -eq ';' }).Count
    $delimiter = if ($semicolonCount -gt $commaCount) { ';' } else { ',' }

    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string]$delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        return @($parser.ReadFields())
    }
    finally {
        $parser.Dispose()
    }
}

function Resolve-FirstExistingCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string[]]$FileNames
    )

    foreach ($fileName in $FileNames) {
        if (Test-SmartFinOpsExcludedCsv -FileName $fileName) { continue }
        $path = Join-Path -Path $FolderPath -ChildPath $fileName
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    return ''
}

function Import-SmartFinOpsSourceCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string[]]$FileNames,
        [string]$SemanticRole = '',
        [string[]]$RequiredColumns = @(),
        [switch]$ValidationOnly,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DataQualityRows
    )

    $eligibleFileNames = @($FileNames | Where-Object { -not (Test-SmartFinOpsExcludedCsv -FileName $_) })
    if ($eligibleFileNames.Count -eq 0) {
        Write-SmartFinOpsLog -Message ("Skipped source {0}; all candidates match the excluded filename patterns." -f $SourceName)
        return @()
    }

    $path = Resolve-FirstExistingCsv -FolderPath $script:SmartM365LatestCsvFolderPath -FileNames $eligibleFileNames
    if ([string]::IsNullOrWhiteSpace($path)) {
        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            Status = 'Missing'
            ContractStatus = 'NotChecked'
            FreshnessStatus = 'NotChecked'
            Path = ($eligibleFileNames -join ' | ')
            FileName = ''
            RowCount = 0
            ColumnCount = 0
            LastWriteTime = ''
            AgeHours = ''
            RequiredColumnsMissing = ''
            Notes = 'Source CSV not found in SmartM365 DATA-LAST.'
        }) | Out-Null
        return @()
    }

    try {
        $item = Get-Item -LiteralPath $path
        $ageHours = [math]::Round(((Get-Date) - $item.LastWriteTime).TotalHours, 1)
        $freshnessStatus = if ($ageHours -gt $script:SmartFinOpsMaxSourceAgeHours) { 'Stale' } else { 'Fresh' }
        $headerColumns = @(Get-SmartFinOpsCsvHeaderColumns -Path $path)
        $missingColumns = @($RequiredColumns | Where-Object { $_ -notin $headerColumns })
        $contractStatus = if ($missingColumns.Count -eq 0) { 'Valid' } else { 'Invalid' }
        $rows = if ($ValidationOnly) { @() } else { @(Import-Csv -LiteralPath $path) }
        $rowCount = if ($ValidationOnly) { '' } else { $rows.Count }
        $notes = New-Object System.Collections.Generic.List[string]
        if ($freshnessStatus -eq 'Stale') {
            $notes.Add(("Source is older than {0} hours." -f $script:SmartFinOpsMaxSourceAgeHours)) | Out-Null
        }
        if ($missingColumns.Count -gt 0) { $notes.Add('Required columns are missing.') | Out-Null }
        if ($ValidationOnly) { $notes.Add('Header and freshness validation only; source rows were not imported.') | Out-Null }

        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            Status = 'Loaded'
            ContractStatus = $contractStatus
            FreshnessStatus = $freshnessStatus
            Path = $path
            FileName = $item.Name
            RowCount = $rowCount
            ColumnCount = $headerColumns.Count
            LastWriteTime = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            AgeHours = $ageHours
            RequiredColumnsMissing = ($missingColumns -join ' | ')
            Notes = ($notes -join ' ')
        }) | Out-Null

        if ($missingColumns.Count -gt 0) {
            Write-SmartFinOpsLog -Level WARN -Message ("Source contract invalid for {0}: missing {1}" -f $SourceName, ($missingColumns -join ', '))
        }
        if ($freshnessStatus -eq 'Stale') {
            Write-SmartFinOpsLog -Level WARN -Message ("Source is stale: {0}, age={1}h" -f $item.Name, $ageHours)
        }
        if (-not $ValidationOnly) {
            Write-SmartFinOpsLog -Message ("Loaded source {0}: {1} row(s)" -f $SourceName, $rows.Count)
        }
        return $rows
    }
    catch {
        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            Status = 'Error'
            ContractStatus = 'Error'
            FreshnessStatus = 'NotChecked'
            Path = $path
            FileName = [System.IO.Path]::GetFileName($path)
            RowCount = 0
            ColumnCount = 0
            LastWriteTime = ''
            AgeHours = ''
            RequiredColumnsMissing = ''
            Notes = $_.Exception.Message
        }) | Out-Null
        Write-SmartFinOpsLog -Level WARN -Message ("Failed to load source {0}: {1}" -f $SourceName, $_.Exception.Message)
        return @()
    }
}

function Import-SmartFinOpsContractSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Key,
        [switch]$ValidationOnly,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DataQualityRows
    )

    $contract = Get-SmartFinOpsSourceContract -Key $Key
    $contractValidationOnly = $contract.PSObject.Properties['ValidationOnly'] -and [bool]$contract.ValidationOnly
    return Import-SmartFinOpsSourceCsv `
        -SourceName ([string]$contract.DisplayName) `
        -FileNames @($contract.FileNames) `
        -SemanticRole ([string]$contract.SemanticRole) `
        -RequiredColumns @($contract.RequiredColumns) `
        -ValidationOnly:($ValidationOnly -or $contractValidationOnly) `
        -DataQualityRows $DataQualityRows
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD80ApIW6lVq/5D
# LKVFvi8d1cZdJxaJ2Iau5snfoVCOGKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPfuaTtQ2x22dSocCXQdeQRZ/86ENgwgcFa2Mo6InsPVMA0GCSqG
# SIb3DQEBAQUABIIBgBcZ+DUgmHjkjZ4YJK0yxsHr4Eg/TNf2NUKfB+er2P3iP96K
# tt0emgMTSi+ieW7WW78b41u4rUYpbMyVv+aokVpEOhBYtwnBfbsMHKu2Rpays829
# xtvU1tWao6nuNys9GWZ/Md5HPk+zqEyoGmGepwnNMR+/aAUT9Y7JChYLhq/+bQBv
# oB7taaDA9S+HNh5QGt6Gn0+anfw01sB0XjR1dnX0zIZoeAI5zw7Zk4faAH9aAco4
# MSFYqyv0TCpt3JSIVmihl2ohpj9jEuIsBQgSp9zCiZLBBTe4chrt3aAOaecJlUju
# E/7qCyqDJqWLfdjCKnOeZpWAuBkP06vrJv89x9WhcM8KHTpW/JRcJik+soztU2kO
# o9IXOJoqinjhbOwbsWj3WeUgPF5lJG9qgxlNF0SstDrkHZv6JVygpnapSh0m7Dmq
# 2aQKARQ3Clyh9gy9h/3oOkBAsMeGCDAD7kC23a0Y1/6mpyNidT8enXdBo4iOgNLp
# RA4J0C6oKfVlbwWxtqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgxNTE5
# NDNaMC8GCSqGSIb3DQEJBDEiBCCE94Xj03Q4KkwyGADbt9jc5ujJ6kwxB2EsuhG0
# 6p+VbDANBgkqhkiG9w0BAQEFAASCAgCNEljH28Je6/pJ1MzRMibskrSSGiN8gqMX
# G0oR8ZJAV5WUehw2X2NGKVCbIhXl9iB3icc6agkeaIotJdw5oDvVU5CHZE01lBp7
# 0CDw4jaMKlmr5US9cgQfHMGg6tPRPQ4Unvgf9fYNah9u+aXHFVmBv21MVgfNHuvs
# 0eMIxIjMjvyxwKaXMC6Kve6Ui/k223V+TQHo46vI3j1ugeLuK1wig6IEUU/0yd5P
# v9B8c1NUY33Rzcc5nnTMMFL307JbsCQbDlKEKUvnLsLYLSkvDuKdGUCHGrX5n6Y/
# JqkyXYu8fQdDPzc18dMxcoJVENScZE0rTXW1Yl1Rhu2p7i9D8TiQ8rFynLhEQpLl
# g/k47RD5M3V5r5Ti2KJnRgiyG4J7JHqfGO5d4/ob/l57tYAqjN87VU+ybmKxojiV
# bmq+AC3bYEyShtMryNi7DuUQ1Zsddw13CidhaqS4ekgM4Kmlm9msuGejPl2D0SLU
# Tg8Mi/GhOhz37mbYDbm6I4fHlH3pWSNtz+AtRhUfV0UcCJwxkg28tueSaVjNiUqL
# DAlpEMnWiuLEdBy8CXZ4FTceMs584GS6DLaqeaScA0CQEFp8BJjgHe+lR4VQ43ot
# z+yrtAdGpGQXGD5Hpb2JR49MPuMjIMGQWRcOwwJZF0Isz4xQVqVcFlGm8sWSRvlW
# rkURj/gmyg==
# SIG # End signature block
