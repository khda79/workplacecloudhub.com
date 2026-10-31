<#
.SYNOPSIS
Verifies SharePoint configuration keys and inheritance without tenant access.
.VERSION
1.0.3
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$inventory = Join-Path $root 'SmartInventory'
$keys = @('EnableSharePointUploadDataAll','EnableSharePointUploadDataLast','EnableSharePointUploadDataPowerBI','EnableSharePointUploadLogAll')
$templates = @((Join-Path $root 'Config/SmartM365.global.local.json.template'), (Join-Path $root 'Config/Tenants/tenant.local.json.template'))
$templates += @(Get-ChildItem $inventory -Recurse -File | Where-Object {
    $_.Name -match '^(.*)\.local\.json(?:\.txt)?\.template$' -and (Test-Path -LiteralPath (Join-Path $_.DirectoryName ($Matches[1]+'.ps1')))
} | ForEach-Object FullName)
$templates += Join-Path $inventory 'M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.local.json.template'
$checks=0
foreach($path in @($templates | Select-Object -Unique)) {
    $config=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    foreach($key in $keys) {
        $property=$config.PSObject.Properties[$key]
        if($null -eq $property -or $property.Value -cne '__USE_GLOBAL__') { throw "Missing inheritance marker: $path / $key" }
        $checks++
    }
}
# Execute the real global/tenant merge loop on synthetic configuration only.
$tokens=$null; $errors=$null
$tenantAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Config/SmartM365-TenantContext.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Tenant context parser failed.'}
$merge=$tenantAst.Find({param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text -match '^foreach \(\$key in \$tenantConfig.Keys\)'},$true)
if(-not $merge){throw 'Tenant merge loop was not found.'}
foreach($key in $keys){
    foreach($value in @($null,'','__USE_GLOBAL__','USE_GLOBAL')) {
        $globalConfig=@{}; $globalConfig[$key]=$true; $tenantConfig=@{}; $tenantConfig[$key]=$value
        & ([scriptblock]::Create($merge.Extent.Text))
        if($globalConfig[$key] -ne $true){throw "Tenant inheritance failed: $key"};$checks++
    }
    $globalConfig=@{}; $globalConfig[$key]=$true; $tenantConfig=@{}; $tenantConfig[$key]=$false
    & ([scriptblock]::Create($merge.Extent.Text))
    if($globalConfig[$key] -ne $false){throw "Tenant override failed: $key"};$checks++
}
# Execute both Exchange local merge functions against null, blank and explicit values.
foreach($relative in @(
    'ExchangeInventory/Hybrid/FreeBusy/SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1',
    'ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
)) {
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $inventory $relative),[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Exchange collector parser failed: $relative"}
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Merge-SmartM365ScriptLocalConfig'},$true)
    if(-not $node){throw "Local merge missing: $relative"}
    $testModule=New-Module -ScriptBlock ([scriptblock]::Create($node.Extent.Text))
    $passed=& $testModule {
        function Read-SmartM365JsonConfig {param($Path,[switch]$Required) return $script:FakeLocal}
        function ConvertTo-SmartM365Hashtable {param($InputObject) $copy=@{};foreach($p in $InputObject.PSObject.Properties){$copy[$p.Name]=$p.Value};return $copy}
        $n=0
        foreach($key in @('EnableSharePointUploadDataAll','EnableSharePointUploadDataLast','EnableSharePointUploadDataPowerBI','EnableSharePointUploadLogAll')){
            foreach($value in @($null,'','  ','__USE_GLOBAL__','USE_GLOBAL',$false)){
                $script:SmartM365EffectiveConfig=[pscustomobject]@{$key=$true}
                $script:FakeLocal=@{};$script:FakeLocal[$key]=$value
                Merge-SmartM365ScriptLocalConfig -ConfigPath 'synthetic'
                $expected=if($value -is [bool]){$false}else{$true}
                if($script:SmartM365EffectiveConfig.$key -ne $expected){throw "Exchange local inheritance failed: $key"}
                $n++
            }
        }
        $n
    }
    $checks += $passed
}
# Execute the real initializer calls, capturing their local and tenant arguments.
foreach($item in @(
    @{File='SmartM365-CopilotUsage-Inventory.ps1';Local='SmartM365LocalConfig'},
    @{File='SmartM365-M365UserActivity-Inventory.ps1';Local='SmartM365UploadLocalConfig'}
)) {
    $path=Join-Path (Join-Path $inventory 'M365Inventory/Usage') $item.File
    $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Collector parser failed: $($item.File)"}
    $call=$ast.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Initialize-SmartM365CollectorUploadContext'},$true)
    if(-not $call){throw "Collector upload initializer missing: $($item.File)"}
    $script:SmartM365LocalConfig=@{EnableSharePointUploadDataAll=$false}
    $script:SmartM365UploadLocalConfig=@{EnableSharePointUploadDataAll=$false}
    $script:SmartM365EffectiveConfig=[pscustomobject]@{EnableSharePointUploadDataAll=$true}
    $OutputPath='synthetic-all';$LatestCsvFolderPath='synthetic-last';$logFolder='synthetic-log';$ValidateOnly=$false
    function Initialize-SmartM365CollectorUploadContext {param($LocalConfig,$GlobalConfig,$DataAllRoot,$DataLastRoot,$LogAllRoot,$ExternalActionsAllowed) return @($LocalConfig,$GlobalConfig)}
    $captured=@(& ([scriptblock]::Create($call.Extent.Text)))
    if($captured.Count -ne 2 -or $captured[0].EnableSharePointUploadDataAll -ne $false -or $captured[1].EnableSharePointUploadDataAll -ne $true){throw "Local/tenant policy not passed by $($item.File)"}
    if($item.Local -eq 'SmartM365UploadLocalConfig' -and $ast.Extent.Text -notmatch 'Read-SmartM365JsonConfig\s+-Path\s+\(Join-Path\s+\$PSScriptRoot\s+''SmartM365-M365UserActivity-Inventory.local.json''\)'){throw 'M365 User Activity does not read its local config.'}
    $checks++
}
$endpoint=Join-Path $inventory 'M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1'
$source=Get-Content -LiteralPath $endpoint -Raw
if ($source -notmatch 'Read-SmartM365JsonConfig\s+-Path' -or $source -notmatch 'Initialize-SmartM365CollectorUploadContext\s+-LocalConfig\s+\$localConfig') { throw 'Endpoint Analytics does not load its local upload configuration.' }
$checks++
Import-Module (Join-Path $inventory 'Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.22' -Force
. (Join-Path $root 'Modules/SmartM365.Core/SmartM365-SharePointUploadPolicy.ps1')
# Combine the actual tenant merge with the actual policy resolver. Each tier
# independently enables, disables or inherits; false must never become a fallback.
$inheritanceValues = @($null, '', '  ', '__USE_GLOBAL__', 'USE_GLOBAL', $false, $true)
foreach ($key in $keys) {
    $category = $key.Substring('EnableSharePointUpload'.Length)
    foreach ($legacy in @($false, $true)) {
        foreach ($globalValue in $inheritanceValues) {
            foreach ($tenantValue in $inheritanceValues) {
                foreach ($localValue in $inheritanceValues) {
                    $globalConfig = @{}; $globalConfig[$key] = $globalValue
                    $tenantConfig = @{}; $tenantConfig[$key] = $tenantValue
                    & ([scriptblock]::Create($merge.Extent.Text))
                    $localPolicyConfig = @{}; $localPolicyConfig[$key] = $localValue
                    $expected = $legacy
                    foreach ($candidate in @($localValue, $tenantValue, $globalValue)) {
                        if ($candidate -is [bool]) { $expected = $candidate; break }
                    }
                    $actual = Resolve-SmartM365SharePointUploadPolicy -Category $category -LocalConfig $localPolicyConfig -GlobalConfig $globalConfig -LegacyEnabled $legacy
                    if ($actual.Enabled -ne $expected) { throw "Three-tier upload precedence failed: $key / legacy=$legacy / global='$globalValue' / tenant='$tenantValue' / local='$localValue'." }
                    $checks++
                }
            }
        }
    }
}
$base=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365ConfigTest-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $base | Out-Null
try {
    $scriptPath=Join-Path $base 'Synthetic.ps1'
    $localPath=Join-Path $base 'Synthetic.local.json.txt'
    $effective=[pscustomobject]@{EnableSharePointUpload=$false;EnableSharePointUploadDataAll=$true;EnableSharePointUploadDataLast=$false;EnableSharePointUploadDataPowerBI=$false;EnableSharePointUploadLogAll=$true}
    foreach($marker in @('__USE_GLOBAL__','USE_GLOBAL','',$null)) {
        $local=@{EnableSharePointUploadDataAll=$marker;EnableSharePointUploadDataLast=$true;EnableSharePointUploadDataPowerBI=$false;EnableSharePointUploadLogAll=$false}
        [IO.File]::WriteAllText($localPath,($local | ConvertTo-Json))
        $result=Get-SmartM365EvidenceConfig -ScriptPath $scriptPath -EffectiveConfig $effective
        if($result.EnableSharePointUploadDataAll -ne $true -or $result.EnableSharePointUploadDataLast -ne $true -or $result.EnableSharePointUploadDataPowerBI -ne $false -or $result.EnableSharePointUploadLogAll -ne $false) { throw "Collector config precedence failed for '$marker'." }
        $checks++
    }
    $global=[pscustomobject]@{EnableSharePointUploadDataAll=$true}
    foreach($marker in @('__USE_GLOBAL__','USE_GLOBAL','')) {
        $result=Resolve-SmartM365SharePointUploadPolicy -Category DataAll -LocalConfig @{EnableSharePointUploadDataAll=$marker} -GlobalConfig $global -LegacyEnabled $false
        if(-not $result.Enabled -or $result.Source -ne 'Global') { throw 'Global inheritance failed.' };$checks++
        $result=Resolve-SmartM365SharePointUploadPolicy -Category DataAll -LocalConfig @{EnableSharePointUploadDataAll=$false} -GlobalConfig $global -LegacyEnabled $true
        if($result.Enabled -or $result.Source -ne 'Local') { throw 'Local override failed.' };$checks++
    }
} finally {
    if (Test-Path -LiteralPath $localPath) { Remove-Item -LiteralPath $localPath -Force }
    $lockPath=Join-Path $base 'Synthetic.local.json.transport.lock'
    if (Test-Path -LiteralPath $lockPath) { Remove-Item -LiteralPath $lockPath -Force }
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Force }
}
"PASS: $checks configuration checks across $(@($templates | Select-Object -Unique).Count) templates; no tenant or network access."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCkjWLMICOjw1i/
# D8gzYHChEcJ1g+1hDMskM2E6PGq/xKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID2bPJaKVM7Nbni5wVgZsqyjVnvtj3ML0ydhHh+asWTAMA0GCSqG
# SIb3DQEBAQUABIIBgFHjhQITUL6bjnRHVKE1LOOrHQz2VngndNm0/O7eFNsMOz8A
# E9IPXH0J0N0Y8QHWgOB4Cb2pp/OwsjgsD5hxgTwz2HoqlTY1ajRs2Aj+8OYZDghL
# tX4vjt3wQ2Io1kJw1YC4B8Ys4HjcsqTiD6RnXarru2OW8jsDf7GX8GgXaVozZylH
# ggqUI4Bmc2OjsGhpa9uYi99UjxFLJtIeV3HKxYCuTQ8Fe0t4harJ3TqSkyCL1nBt
# HA0Pf3HpBENd8k4TLWc5S9AklU2jnPMpkhgx+lv0BYLUmW+vO9EmVhF9GlkOKGqR
# HWOx9gPYLa/QODCt8PR2w3i8sYBHE6bMPS2YH7Bamc0TCyzWfzoEJXlbTEB4SgYm
# jxMbyBSOywLatBFadF1S253RV2tfpdGdHUOr6PcLq4rvXYNP3MWfzZWQrAcbzo2t
# B4popRLpekBygkHKtQadUXar5gT6rCs/WD0Eo9tx50YF1wXBRZN/4UT7onFvsJ93
# n2l20oFRwIDX/kIyEaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNDQ0
# NDJaMC8GCSqGSIb3DQEJBDEiBCD2G5YCPrpyJacgcHxUYwREvCymhemsbKnoRFwo
# ax/7KTANBgkqhkiG9w0BAQEFAASCAgC19J98+OyIgsmXUo4enzTQMfY4gf0BZZaQ
# SiQHOXuC6TE4afVpCDHuzZsrKpbUzyR0x1brZ8MF1eZ70md0Q1/3Ts5yzWcFoXSK
# 7RNqxrCQNWw0UmsYnCKNZJy3yUudyKbkYOYWDVqunEPkuI/fB207V5gn7EgKI2vA
# hmYF4fwWuqGG63XApsDwG9qjt9Mmrj7UXLxpJOaeEp5QQpkKg9I3oLPDkdUZT7vb
# FKBjLtmE8sUoOthhV5ZvaoKp1OoJZd++L4Zv0Nn44j8zUATeC91+tAfUECrhvlqe
# 6Gxc8+KkJa+zJzZFKLEcQ6rRon1dWiuQn/GCzq4CssIWUini5vPaDaz0rWczLnDB
# iI381+oPIBy9ds2uxg1AiXjjec9ao1ZuFQAwRWtQ3StjYEtSaRrTtdLLikU9cnOw
# CHSvTM0htYmkNMGqVf0CAYrzAPQM3bghKOgGdH5uQtPcGaALkkPZvwq5c/H5BZ2z
# 11ePIimnmH5nXsRWGtJL7DLBk0HUf4Wdfqrn3RXOBicRHC9YRoP4dTPx2ODPiBT7
# gm5itUFUtvMTwflhgWbkMJfPhmhBAFTHs5WrvB4WSsAuUuAu1mk2uEQAf2RcOH0Y
# CedTP2DZ3IHO3XIWHSPODDZSKnliuMQEmQTaMa37u9WL5rVDJe0QFjxaaqHrA1YT
# DCk/6rgUuQ==
# SIG # End signature block
