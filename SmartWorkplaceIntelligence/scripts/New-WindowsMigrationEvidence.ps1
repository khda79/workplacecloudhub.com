[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$PreparedInventoryPath,
    [Parameter(Mandatory)][string]$OutputRoot,
    [string]$PolicySelectorsPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config/windows-migration-policy-selectors.json.txt')
)
# Staging only. No collection, publication, upload or Power BI refresh.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
Import-Module (Join-Path $PSScriptRoot 'WindowsMigrationEvidence.psm1') -Force
$data=(Resolve-Path -LiteralPath $DataRoot).ProviderPath
$last=Join-Path $data 'DATA-LAST'
$output=[IO.Path]::GetFullPath($OutputRoot).TrimEnd('\','/')
if ($output -eq $data.TrimEnd('\','/') -or $output -eq $last.TrimEnd('\','/') -or $output.StartsWith($last.TrimEnd('\','/')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Outputs cannot overwrite raw sources.' }
$sources=[Collections.Generic.List[object]]::new()
function Read-StrictCsv([string]$Path,[string[]]$Required,[string[]]$Optional=@()) {
    $file=Get-Item -LiteralPath $Path
    $receipt=[ordered]@{Path=$file.FullName;Bytes=$file.Length;PublicationUtc=$file.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ');SHA256=(Get-FileHash -LiteralPath $Path).Hash}
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($file.FullName)
    try {
        $parser.SetDelimiters(','); $parser.HasFieldsEnclosedInQuotes=$true; $parser.TrimWhiteSpace=$false
        $head=$parser.ReadFields()
        if (-not $head -or @($head | Sort-Object -Unique).Count -ne $head.Length -or @($head | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) { throw 'Missing or duplicate CSV headers.' }
        foreach ($c in $Required) { if ($c -cnotin $head) { throw "Required CSV column missing in $($file.Name): $c" } }
        # Keep only the narrow projection, not hundreds of unused AD attributes.
        $keep=@($Required+$Optional | Select-Object -Unique | Where-Object { $_ -cin $head })
        $indices=@{}; foreach($c in $keep){$indices[$c]=[array]::IndexOf($head,$c)}
        if('TenantKey' -cin $head -and 'TenantKey' -cnotin $keep){$keep+='TenantKey';$indices['TenantKey']=[array]::IndexOf($head,'TenantKey')}
        $rows=[Collections.Generic.List[object]]::new()
        while (-not $parser.EndOfData) {
            $fields=$parser.ReadFields()
            if ($fields.Length -ne $head.Length) { throw "Malformed CSV record: $($file.Name)" }
            $row=[ordered]@{}; foreach($c in $keep){$row[$c]=$fields[$indices[$c]]}
            $rows.Add([pscustomobject]$row)
        }
        $sources.Add([pscustomobject]$receipt)
        Write-Host ('Validated {0}: {1} records, {2} selected columns.' -f $file.Name,$rows.Count,$keep.Count)
        return ,$rows.ToArray()
    } finally { $parser.Dispose() }
}
$inventory=Read-StrictCsv (Join-Path $last 'Intune_Devices_Inventory.csv') @('TenantKey','Device ID','Device name','OS version')
$prepared=Read-StrictCsv $PreparedInventoryPath @('Device Source ID','Device Name','Operating System Version','Windows Generation','Windows 11 Upgrade Eligibility','Country','Free Disk Space (GB)','Windows Update State','Windows Update Action Code','Last Sync DateTime') @('Primary User UPN','Windows Release','Management State','Windows 11 Blocking Reasons')
$ad=Read-StrictCsv (Join-Path $last 'AD_Computers_AllDomains.csv') @('TenantKey','ObjectGUID','Name','Enabled','OperatingSystemShortName','DomainName','IntuneDeviceId','operatingSystemVersion') @('SamAccountName','OperatingSystemDisplayName')
$readiness=Read-StrictCsv (Join-Path $last 'Intune_Devices_UpgradeEligibility.csv') @('TenantKey','GraphId','DeviceName','NormalizedDeviceName','OSVersion','ExportDateTime')
$policy=Read-StrictCsv (Join-Path $last 'Intune_WindowsUpdate_Status.csv') @('TenantKey','PolicyId','PolicyName','DeviceId','AggregateState','BlockingReason','ExportDateTime') @('CurrentDeviceUpdateStatus','CurrentDeviceUpdateStatus_loc','LatestAlertMessage','LatestAlertMessage_loc')
$tenants=@($inventory | Select-Object -ExpandProperty TenantKey -Unique)
if ($tenants.Count -ne 1 -or [string]::IsNullOrWhiteSpace($tenants[0])) { throw 'A non-empty single-tenant Intune inventory is required.' }
$selectorBytes=[IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $PolicySelectorsPath).ProviderPath)
$selectors=[Text.Encoding]::UTF8.GetString($selectorBytes).TrimStart([char]0xfeff) | ConvertFrom-Json
$result=New-WindowsMigrationProjection -Inventory $inventory -Prepared $prepared -AD $ad -Readiness $readiness -Policy $policy -TenantKey $tenants[0] -IntunePublicationUtc $sources[0].PublicationUtc -ADPublicationUtc $sources[2].PublicationUtc -PolicySelectors $selectors
# Confirm caller did not feed a changing source set; normal pipeline uses verified copies.
foreach ($source in $sources) {
    $now=Get-Item -LiteralPath $source.Path
    if ($now.Length -ne $source.Bytes -or $now.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ') -cne $source.PublicationUtc -or (Get-FileHash -LiteralPath $source.Path).Hash -ne $source.SHA256) { throw 'Source changed during migration preparation. No outputs committed.' }
}
New-Item -ItemType Directory -Path $output -Force | Out-Null
function Export-Projection([string]$Name,$Rows,[string[]]$Columns) {
    $destination=Join-Path $output $Name
    $temporary=$destination+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
    if ($Rows.Count) {
        $Rows | Select-Object $Columns | ForEach-Object {
            if ($_.PSObject.Properties['Free Disk GB'] -and $null -ne $_.'Free Disk GB') { $_.'Free Disk GB'=([double]$_.'Free Disk GB').ToString('R',[Globalization.CultureInfo]::InvariantCulture) }
            $_
        } | Export-Csv -LiteralPath $temporary -NoTypeInformation -Encoding utf8
    } else {
        [IO.File]::WriteAllText($temporary,('"'+($Columns -join '","')+'"'+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
    }
    [IO.File]::Move($temporary,$destination,$true)
}
Export-Projection 'WindowsMigrationEvidence.csv' $result.Windows $result.Columns
Export-Projection 'WindowsMigrationADObservations.csv' $result.ADOnly $result.ADColumns
$audit=[ordered]@{SchemaVersion=1;CreatedUtc=[datetime]::UtcNow.ToString('O');Projection=$result.Audit;Sources=$sources.ToArray();PolicySelectorsSHA256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($selectorBytes));Publication=$false}
[IO.File]::WriteAllText((Join-Path $output 'WindowsMigrationAudit.json.txt'),($audit | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
$result.Audit | ConvertTo-Json -Depth 6

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlkH5QKhBvaK2F
# kLEIC+GH5tHfngFuTC+SQQxmPe+JYKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPXbxwN1NQanuHOCEKCJf1LmDBngZ7poEbe8346wDlZVMA0GCSqG
# SIb3DQEBAQUABIIBgCyiiLn7yFwJwGWxco6cEfne/FPUD42eLWJEn4iCO+y+6Ndf
# tFn6RDaFCr2ztpXMe2dcoGJJzX/UPP3uS9nmb20ZWRY3An4MfBUPvwJCgVe/V1na
# yPfxhMmybXJ0ys/ovx14WR0ShcbHlxMPsjdX3/LuIVwajsRHgPQVutS6HjuRqUQN
# XsmxReGrDrFd+EAOQh3huXDICgiRyhY3Ga5fiitj1vvTpcz9lzdu+Eg2oPjw3sng
# IIJGQ7mdbNqORQB6wwtK19oXOedtLVsCLTWp2ATmYrQudMobSRQatkYsbZ8eF4Ho
# CxxaD1qbCMxqcAXqn1IembKOd9r6RxWfCk9uceGhPqudkJoF3LrKM5C6Nv6zRGjq
# u/UFwZZ+Q/NT1kPuEzivhdEpg/7W723miBXE1I0RueP6jxjh0/3EPUQHaCvzOCrD
# epDM5t7zszxT/b9xsp1ZSiKxFAzDhjpf7DtTQRJcMMaLgMLSILulGD6mkPKEy3Eo
# azu3eFq0sVFEv/tN3KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkxNDA4
# MzZaMC8GCSqGSIb3DQEJBDEiBCDrldt/AbffOQu8wLwwziS2HxhUCIr9HVsK7TO2
# uwORbTANBgkqhkiG9w0BAQEFAASCAgCTNB+SCPbWtuIRoW6wLpMLvZGMIeYE/6WY
# WKJlcGMc9hDmndPLQuT1JM0iPSL9U7oA0cxnkmbLEEnoAru+KeA5QNB86NMyAWlV
# mf05EU+uJcrA74O9sx3d7CIcsIJPNsMl6BJE8EiHFttcPxQXD+MWGd61Bd/WrWa9
# wAGOX86Z6OqyKngLFKe1QmG/hQHM8kzJwmd9b8Fxcq/oHmKNdoKf46jfoJvwx0eR
# YP35NQ5Uy8mO+nzF+E7H4pNT+9WWsUkNU5cqMnsa2PL0cs9UXmI5y12KWUM4k59F
# T01kE5y9F6rKoyhP9eNOZXM/BmIWxmIwOFHxbxtz2oA3lCHaOrzosmYu1stsNidp
# 2IHTUHWU7/B7qbEaYT+cqledvR7G5kFL6rrhvs2wr/KL1pI5ZH4ZkAaZCgLYpvaA
# Wl65iqhtEnWqLDjgCzC6n+mpFk+E9dFFl4Cj7YhXCsIB65VuPXdxN5CN9u3UuEuu
# uVE/4xSUkoK9ryninealyw+cODK3uyLydVp3k0ysJhoqQGfDWecS1XEpQ2YU9V3x
# gH7Uiq3OavL4cf7GFmQG168tD3VbgVmPK1hdAQ6NTWXYkuGjaFOeneoUg8T3uNaS
# m9gieJ2c2G3+pOBUZlyNISHDQ6bIR5QBMXU8EneymLaOAl2mZFrfN3KWHmszLSgP
# 8Au7wGmcmg==
# SIG # End signature block
