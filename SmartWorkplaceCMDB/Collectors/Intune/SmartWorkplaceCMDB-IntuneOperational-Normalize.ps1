<#
.SYNOPSIS
Normalizes Intune operational inventory into dedicated Power BI tables.

.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant='default',
    [string]$OrganizationKey,[string]$EnvironmentKey,[string]$TenantKey,[string]$TenantId,
    [string]$DataRootPath,[string]$DataAllRootPath,[string]$LatestOutputRootPath,[string]$LogRootPath,
    [string]$GlobalConfigPath,[string]$TenantConfigPath,
    [switch]$NoConfigWrite,[switch]$ValidateOnly
)

$ScriptVersion='1.0.0'
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0

function Get-HeaderStatus { param([string]$Path,[string[]]$Columns) if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){return 'Missing'};$line=Get-Content -LiteralPath $Path -TotalCount 1;$actual=if([string]::IsNullOrWhiteSpace($line)){@()}else{@($line.Split(',') | ForEach-Object {$_.Trim().Trim('"')})};if(($actual -join [char]31) -ceq ($Columns -join [char]31)){return 'Valid'};return 'Incompatible' }
function Get-KeyText { param([AllowNull()]$Value) if($null -eq $Value){return ''};return ([string]$Value).Trim().ToLowerInvariant() }

$scriptRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot=Split-Path -Parent (Split-Path -Parent $scriptRoot)
$module=Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$rawContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
$curatedContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json'
Import-Module $module -Force
$bound=@{};foreach($key in $PSBoundParameters.Keys){$bound[$key]=$PSBoundParameters[$key]}
$context=Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths=$context.Paths
$rawContract=Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$curatedContract=Get-SmartWorkplaceCMDBTableContract -Path $curatedContractPath
$mappings=@(
    @{Raw='Intune_AutopilotDevices.csv';Curated='FactAutopilotDevice.csv';Key='AutopilotDeviceId';Prefix='autopilot'},
    @{Raw='Intune_DetectedApps.csv';Curated='DimDetectedApplication.csv';Key='AppId';Prefix='app'},
    @{Raw='Intune_ConfigurationPolicies.csv';Curated='DimIntuneConfigurationPolicy.csv';Key='PolicyId';Prefix='configuration-policy'},
    @{Raw='Intune_WindowsUpdatePolicies.csv';Curated='DimWindowsUpdatePolicy.csv';Key='PolicyId';Prefix='update-policy'}
)
$definitions=@()
foreach($mapping in $mappings){
    $raw=@($rawContract.tables | Where-Object name -eq $mapping.Raw);$target=@($curatedContract.tables | Where-Object name -eq $mapping.Curated)
    if($raw.Count -ne 1 -or $target.Count -ne 1){throw "Operational contract definition missing or duplicated: $($mapping.Raw) / $($mapping.Curated)"}
    $input=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$raw[0].area) ([string]$raw[0].name))))
    $output=[IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$target[0].area) ([string]$target[0].name))))
    $inputStatus=Get-HeaderStatus $input @($raw[0].columns|ForEach-Object{[string]$_});$outputStatus=Get-HeaderStatus $output @($target[0].columns|ForEach-Object{[string]$_})
    if($inputStatus -eq 'Incompatible' -or $outputStatus -eq 'Incompatible'){throw "Operational CSV contract is incompatible: $($mapping.Raw) / $($mapping.Curated)"}
    $definitions+=@{Mapping=$mapping;Raw=$raw[0];Target=$target[0];Input=$input;Output=$output;InputStatus=$inputStatus;OutputStatus=$outputStatus}
}
if($ValidateOnly){foreach($definition in $definitions){if($definition.InputStatus -eq 'Valid'){Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $definition.Input -Paths $paths|Out-Null}};[pscustomobject]@{Status='Valid';ScriptVersion=$ScriptVersion;DatasetCount=$definitions.Count;RawContractVersion=[string]$rawContract.contractVersion;CuratedContractVersion=[string]$curatedContract.contractVersion}|Format-List;return}
$identity=@{TenantKey=$paths.TenantKey;OrganizationKey=$paths.OrganizationKey;EnvironmentKey=$paths.EnvironmentKey;TenantId=$paths.TenantId}
$published=@()
foreach($definition in $definitions){
    if($definition.InputStatus -eq 'Missing'){throw "Required raw operational inventory is missing: $($definition.Input)"}
    $rows=@(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $definition.Input -Paths $paths)
    $keyName=[string]$definition.Mapping.Key
    $duplicates=@(if($definition.Mapping.Raw -eq 'Intune_WindowsUpdatePolicies.csv'){$rows | Group-Object {"$($_.PolicyType)|$($_.PolicyId)"} | Where-Object Count -gt 1}else{$rows | Group-Object $keyName | Where-Object Count -gt 1})
    if($duplicates.Count){throw "Duplicate operational keys found in $($definition.Mapping.Raw): $($duplicates.Name -join ', ')"}
    $targetRows=@($rows|ForEach-Object{
        $row=$_
        $key=Get-KeyText $row.$keyName;if([string]::IsNullOrWhiteSpace($key)){throw "Empty $keyName in $($definition.Mapping.Raw)."}
        $properties=[ordered]@{}
        switch($definition.Mapping.Raw){
            'Intune_AutopilotDevices.csv'{$properties.TenantAutopilotDeviceKey=('{0}|autopilot|{1}'-f$paths.TenantKey,$key);foreach($name in @('AutopilotDeviceId','DisplayName','SerialNumber','Manufacturer','Model','GroupTag','EnrollmentState','LastContactedDateTime','AzureAdDeviceId','ManagedDeviceId','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_DetectedApps.csv'{$properties.TenantApplicationKey=('{0}|detected-app|{1}'-f$paths.TenantKey,$key);foreach($name in @('AppId','DisplayName','Version','Publisher','DeviceCount','Platform','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_ConfigurationPolicies.csv'{$properties.TenantPolicyKey=('{0}|configuration-policy|{1}'-f$paths.TenantKey,$key);foreach($name in @('PolicyId','DisplayName','Description','Platforms','Technologies','TemplateId','TemplateFamily','CreatedDateTime','LastModifiedDateTime','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
            'Intune_WindowsUpdatePolicies.csv'{$type=Get-KeyText $row.PolicyType;$properties.TenantUpdatePolicyKey=('{0}|update-policy|{1}|{2}'-f$paths.TenantKey,$type,$key);foreach($name in @('PolicyType','PolicyId','DisplayName','TargetVersion','ReleaseDateTime','DaysUntilForcedReboot','CreatedDateTime','LastModifiedDateTime','SourceCollectedDateTime')){$properties[$name]=[string]$row.$name}}
        }
        [pscustomobject]$properties
    })
    Export-SmartWorkplaceCMDBCsv -InputObject $targetRows -Path $definition.Output -Columns @($definition.Target.columns|ForEach-Object{[string]$_}) @identity
    if((Get-HeaderStatus $definition.Output @($definition.Target.columns|ForEach-Object{[string]$_})) -ne 'Valid'){throw "Curated operational output failed validation: $($definition.Output)"}
    $published+=[pscustomobject]@{Table=$definition.Mapping.Curated;Count=$targetRows.Count;Path=$definition.Output}
}
Write-Information ("SmartWorkplaceCMDB Intune operational normalization completed. Tables={0}; Rows={1}." -f $published.Count, (($published | Measure-Object Count -Sum).Sum)) -InformationAction Continue
[pscustomobject]@{Status='Completed';ScriptVersion=$ScriptVersion;Published=$published;RawContractVersion=[string]$rawContract.contractVersion;CuratedContractVersion=[string]$curatedContract.contractVersion}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC4yWSgm+PyR+j9
# DdYDm+90navEKazE65JXWYN6FakjoKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC4/To8iLMhLHnGm0+N7M0LXKSFZvQuioNBaZ2nN1PkoMA0GCSqG
# SIb3DQEBAQUABIIBgFzMHGWZdlbRoIC8dOHrsDw+d9Mb/Gys0Dct+GGLP4uJ4jq7
# WW1tAQ5gta9y0/wXaWiUqckSfvy6Qodb4F+W/LdXLe/B8CEmtwNMCR6aMfrIG35V
# 7cQRWxagVGnln+6yVZWt9ib0GUH+3rBbuFS8mC03ablVrFLiiSMa95fscq2kqtYz
# d7etmsJQV+NRjtLDu6+zPvxNvOpgOeLPMNBTvn5kSOQOaHYtcUk9B56RuWjnfbF/
# wpVrsNAPM+uJNImeeYfUk5EoO+b4E9GVuEQmi6OkoGKU8JSG/PvDeGcx3EmHKJ8f
# PfBgOZy4I4IVFG2HxSVxafVq505W4ZQNiDfP2eeMAaXVabsKYQvG+FlZB+NWtCkF
# J9ubXOtYr8AtYY2yTe9u3j5bUFgb/MTEKWFaEbPTN4dwsS+kuNzAg5lp4d7GeCFA
# 8OTUHOpST++fuBDKYz/DWSNtMLNT7+2H5r/gT97s//oQyf3dKLmEgIcTxZRyy6qg
# SgYO7ehN6GsyKUx+2KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxODI0
# NDhaMC8GCSqGSIb3DQEJBDEiBCDIcvt1+U7lRXccODyYGfEpPjy7zboH0rIKlX9b
# vmLphjANBgkqhkiG9w0BAQEFAASCAgBy9NZGQCIgeopnhOrmeHsxf1de5QCFGZIO
# WOz9WWXrSEwRFU/j7XKXPwB+s3Lc324QvJJi7z/P80ukdBKx4u2Shtw+ndgWcDhr
# yL0IH30DZO2PRBF2bzR+JuCzLVXf+22zzUsLKwUR9+v/apiYRwiwdwFIST5s/2GL
# aT3rNX8/mzDTYp7gxKZXBFLo/+bVMrtvGQpSw1wyjDoSGb0dP96u8lw9OLnvZS9I
# wm1TFrGZ5n7aA8S9xrCFoUFzwBzYnJjfJ3crGtzwiDG0GpBIXVjeC+qu1erc4sf3
# j9sUA+yav0pbxRiPyaFM1iGlvppQTM9O7ikoISWby307fpgwrD9kLvStViw54mgI
# rnt3E2I/gJsbRVQww3zhBegYwCuNmhDK4tuuzoSEpBsnfAY3cTTQlC848yZXSm1h
# CzsORL0vTChAGIazTuJnst17V87CYiow5rwnLxsUDxsR0Sxk3sgdeElqBzYVMWDW
# L+sDKhJrmpkPg6PFdh78AJM1/3VcAUdR2a9AhiXjB+qBXtCQ8Yw4cravJZTrVeOd
# /S2UjVU+GEJPxX7xr5p8jQG/H2Mn1ifEbaFxOdQdG+u5TIVqASI39YQBPmttP9Th
# 9wI7Zl0Y2FzduJU/o/BJVr29Jtj28gC/Y5nSUGLAI41FbAOVJsMqq18701SQslU9
# VZBGtrasMQ==
# SIG # End signature block
