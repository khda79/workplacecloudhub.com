<#
.SYNOPSIS
Validates Intune operational collection and normalization with synthetic data.

.VERSION
1.3.2
#>
[CmdletBinding()]param()
$ScriptVersion='1.3.2';$ErrorActionPreference='Stop';Set-StrictMode -Version 2.0
$passed=0;$failed=0
function Invoke-Test{param([string]$Name,[scriptblock]$Test)try{&$Test;$script:passed++;Write-Information "PASS $Name" -InformationAction Continue}catch{$script:failed++;Write-Information "FAIL $Name - $($_.Exception.Message)" -InformationAction Continue}}
function Assert-True{param([bool]$Condition,[string]$Message)if(-not$Condition){throw $Message}}
$projectRoot=Split-Path -Parent $PSScriptRoot
$collector=Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneOperational-Collect.ps1'
$normalizer=Join-Path $projectRoot 'Collectors\Intune\SmartWorkplaceCMDB-IntuneOperational-Normalize.ps1'
$fixture=Join-Path $PSScriptRoot 'Fixtures\IntuneOperational.sample.json'
$tempRoot=Join-Path ([IO.Path]::GetTempPath()) ('SmartWorkplaceCMDB-IntuneOperational-{0}'-f[guid]::NewGuid().ToString('N'))
$identity=@{Tenant='audit';OrganizationKey='contoso';EnvironmentKey='test';TenantKey='contoso-test';TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';DataRootPath=$tempRoot;NoConfigWrite=$true}
try{
    New-Item -ItemType Directory -Path $tempRoot -Force|Out-Null
    Invoke-Test 'Use AppInvRawData as the only complete detected-application source' {
        $source=Get-Content -Raw -LiteralPath $collector
        $directCalls=[regex]::Matches($source,"Invoke-SmartWorkplaceCMDBGraphPagedRequest[^\r\n]+deviceManagement/detectedApps\?")
        Assert-True ($directCalls.Count -eq 1) 'The collector must retain exactly one bounded detectedApps query.'
        Assert-True ($source -match 'if\(\$MaxItems-gt0\)\{\s*\$apps=@\(Invoke-SmartWorkplaceCMDBGraphPagedRequest[^\r\n]+deviceManagement/detectedApps\?') 'The direct detectedApps query is not restricted to bounded collection.'
        $completeBranchMatch=[regex]::Match($source,'if\(\$MaxItems\s*-eq\s*0\)\{')
        $completeBranchIndex=if($completeBranchMatch.Success){$completeBranchMatch.Index}else{-1}
        $exportCallIndex=$source.IndexOf('$exportCsv=Get-AppInventoryRawExport',[StringComparison]::Ordinal)
        Assert-True ($completeBranchIndex -ge 0 -and $exportCallIndex -gt $completeBranchIndex) 'Complete collection is not routed through AppInvRawData.'
    }
    Invoke-Test 'Validate without creating output' {&$collector @identity -InputJsonPath $fixture -ValidateOnly|Out-Null;Assert-True(-not(Test-Path(Join-Path $tempRoot 'DATA-LAST')))'ValidateOnly created output.'}
    Invoke-Test 'Collect five independent source tables' {
        & $collector @identity -InputJsonPath $fixture|Out-Null
        $root=Join-Path $tempRoot 'DATA-LAST\Raw\Intune'
        Assert-True (@(Import-Csv (Join-Path $root 'Intune_AutopilotDevices.csv')).Count -eq 2) 'Expected two Autopilot rows.'
        Assert-True (@(Import-Csv (Join-Path $root 'Intune_DetectedApps.csv')).Count -eq 2) 'Expected two detected applications.'
        Assert-True (@(Import-Csv (Join-Path $root 'Intune_DetectedAppDeviceRelationships.csv')).Count -eq 3) 'Expected three exact application-device relationships.'
        Assert-True (@(Import-Csv (Join-Path $root 'Intune_ConfigurationPolicies.csv')).Count -eq 1) 'Expected one configuration policy.'
        Assert-True (@(Import-Csv (Join-Path $root 'Intune_WindowsUpdatePolicies.csv')).Count -eq 2) 'Expected feature and quality update policies.'
    }
    Invoke-Test 'Normalize dedicated Power BI grains' {
        & $normalizer @identity|Out-Null
        $root=Join-Path $tempRoot 'DATA-LAST\PowerBI'
        $autopilot=@(Import-Csv(Join-Path $root 'FactAutopilotDevice.csv'))
        $apps=@(Import-Csv(Join-Path $root 'DimDetectedApplication.csv'))
        $appDevices=@(Import-Csv(Join-Path $root 'FactDeviceApplication.csv'))
        $configuration=@(Import-Csv(Join-Path $root 'DimIntuneConfigurationPolicy.csv'))
        $updates=@(Import-Csv(Join-Path $root 'DimWindowsUpdatePolicy.csv'))
        Assert-True ($autopilot.Count -eq 2 -and $autopilot[0].TenantAutopilotDeviceKey -match '\|autopilot\|') 'Autopilot key or count is invalid.'
        Assert-True ($apps.Count -eq 2 -and $apps[0].SourceApplicationKey -eq 'app-001' -and $apps[0].DeviceCount -eq '2' -and $apps[0].ReportedDeviceCount -eq '12' -and $apps[0].RelationshipCoverageStatus -eq 'ReconciledCountMismatch') 'Detected application identity or count reconciliation is invalid.'
        Assert-True ($appDevices.Count -eq 3 -and @($appDevices.TenantDeviceApplicationKey|Select-Object -Unique).Count -eq 3) 'Application-device fact grain is invalid.'
        Assert-True ($configuration.Count -eq 1 -and $configuration[0].TemplateFamily -eq 'endpointSecurityAntivirus') 'Configuration policy mapping is invalid.'
        Assert-True ($updates.Count -eq 2 -and (@($updates.PolicyType|Sort-Object) -join ',') -eq 'Feature,Quality') 'Update policy types are invalid.'
    }
    Invoke-Test 'Upgrade an older curated header without weakening the raw contract guard' {
        $legacyPath=Join-Path $tempRoot 'DATA-LAST\PowerBI\DimDetectedApplication.csv'
        $legacyRows=@(Import-Csv -LiteralPath $legacyPath)
        $legacyRows|Select-Object TenantKey,OrganizationKey,EnvironmentKey,TenantId,TenantApplicationKey,AppId,DisplayName,Version,Publisher,DeviceCount,Platform,SourceCollectedDateTime|Export-Csv -LiteralPath $legacyPath -NoTypeInformation -Encoding UTF8
        & $normalizer @identity|Out-Null
        $header=Get-Content -LiteralPath $legacyPath -TotalCount 1
        Assert-True ($header -match 'SourceApplicationKey' -and $header -match 'RelationshipCoverageStatus') 'The current curated contract did not replace the older header.'
        Assert-True (-not @(Get-ChildItem -LiteralPath (Split-Path -Parent $legacyPath) -Filter 'DimDetectedApplication.csv.tmp.*.csv').Count) 'A curated staging file was left behind.'
    }
    Invoke-Test 'Consolidate strictly equivalent detected application duplicates' {
        $duplicateRoot=Join-Path $tempRoot 'EquivalentDuplicate'
        $duplicateIdentity=@{}+$identity;$duplicateIdentity.DataRootPath=$duplicateRoot
        $duplicateFixture=Join-Path $tempRoot 'equivalent-duplicate.json'
        $fixtureObject=Get-Content -Raw -LiteralPath $fixture|ConvertFrom-Json
        $fixtureApps=[object[]]$fixtureObject.detectedApps
        $fixtureObject.detectedApps=[object[]]@($fixtureApps+$fixtureApps[0])
        ConvertTo-Json -InputObject $fixtureObject -Depth 12|Set-Content -LiteralPath $duplicateFixture -Encoding UTF8
        & $collector @duplicateIdentity -InputJsonPath $duplicateFixture|Out-Null
        $apps=@(Import-Csv (Join-Path $duplicateRoot 'DATA-LAST\Raw\Intune\Intune_DetectedApps.csv'))
        Assert-True ($apps.Count -eq 2) 'Equivalent duplicate application rows were not consolidated.'
    }
    Invoke-Test 'Reconcile conflicting detected application duplicates deterministically' {
        $conflictIdentity=@{}+$identity;$conflictIdentity.DataRootPath=Join-Path $tempRoot 'ConflictingDuplicate'
        $conflictFixture=Join-Path $tempRoot 'conflicting-duplicate.json'
        $fixtureObject=Get-Content -Raw -LiteralPath $fixture|ConvertFrom-Json
        $fixtureApps=[object[]]$fixtureObject.detectedApps
        $conflict=ConvertFrom-Json (ConvertTo-Json -InputObject $fixtureApps[0] -Depth 8)
        $conflict.deviceCount=99
        $conflict.displayName='Z Example Browser Alias'
        $fixtureObject.detectedApps=[object[]]@($fixtureApps+$conflict)
        ConvertTo-Json -InputObject $fixtureObject -Depth 12|Set-Content -LiteralPath $conflictFixture -Encoding UTF8
        & $collector @conflictIdentity -InputJsonPath $conflictFixture|Out-Null
        $apps=@(Import-Csv (Join-Path $conflictIdentity.DataRootPath 'DATA-LAST\Raw\Intune\Intune_DetectedApps.csv'))
        $reconciled=@($apps|Where-Object AppId -eq 'app-001')[0]
        Assert-True ($apps.Count -eq 2) 'Conflicting duplicate application rows were not consolidated.'
        Assert-True ($reconciled.DisplayName -eq 'Example Browser' -and $reconciled.DeviceCount -eq '2' -and $reconciled.ReportedDeviceCount -eq '99') 'Conflicting duplicate application values were not reconciled deterministically.'
    }
    Invoke-Test 'Bound each source family independently' {
        $bounded=@{}+$identity;$bounded.DataRootPath=Join-Path $tempRoot 'Bounded'
        $result = & $collector @bounded -InputJsonPath $fixture -MaxItems 1
        foreach($name in @('Intune_AutopilotDevices.csv','Intune_DetectedApps.csv','Intune_ConfigurationPolicies.csv')){
            $path = @($result.PublishedPath | Where-Object { [IO.Path]::GetFileName($_) -eq $name })[0]
            Assert-True (@(Import-Csv -LiteralPath $path).Count -eq 1) "$name was not bounded."
        }
    }
}finally{if(Test-Path -LiteralPath $tempRoot){Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue}}
Write-Information "SmartWorkplaceCMDB Intune operational tests completed. Version=$ScriptVersion; Passed=$passed; Failed=$failed" -InformationAction Continue
if($failed -gt 0){exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCASo1TyPymqb0Da
# iifLcikHp0G0iGo8UUfZEa+iyxq6KqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFOeTfbdw6FqqxzpkSgdHKu0GQWX9NtKRU7/QIV8Ie+aMA0GCSqG
# SIb3DQEBAQUABIIBgG+K6isx3tS7EGxnlSMr3C3MwWMJBCNkNUeEr214Xs6dAxp6
# f93/lnotOCdw4GBEYZ5GDlRvEHsYTV3OHVLMq7EvSnGHDvrvkPg3ip2L9f4iTUUJ
# gSwfOtaGEEw64nuU30b5DHfmghvVSEY464fnlhPVIDZjsethrheg4bGYSkxd165t
# zG9ZfiB9n3mm2+e4P8X0XQxpFqTotPPLYQHG/3AOnuFhCA9Zpe4nytncnjQcIfK9
# DzZcLpV99mXKkk/HnpDPFykXwZ8wNrBuqW/xD/BxZ2GQxf4ePdmv7kA6CKA71IYP
# uB8JP76h/u96Dx0ptVma3rtcejYszHjy5NRh3fEI+dyDGeHD+8Gh3JmFcbKecqbP
# sKlQuhzW+a5qCMExrGGQzhy4TVjNncMuUKGC4i+cyNObCHLADF8ddH3wbo8FVneZ
# rNgkoeYk1Dyl/xwrF/Gef9Xdgf9rjkx39lNnpguTYn66sZyF1oO0No6bqx3SENHJ
# fRo+Uwnsml8q0WRZoKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwODE3
# NTlaMC8GCSqGSIb3DQEJBDEiBCA4ZapxgQ+GxnOp1sBp2LoTcEIke77KHsqvbxu1
# f1EMojANBgkqhkiG9w0BAQEFAASCAgBk8p9tbM1nmmym24CBRSzOca0LkHnVOtdD
# 90/zoaiuJ7XRriKElvKtZDzCRH22oOOaj/m+sdQOujs3ESkOpEDfXA26fffzKzqm
# xQDUq4DtxCpEW+40u7bGT3JK6SINRg09aKhpwxKv3sM+V/16zzuWQ3eV2/MCT0Nh
# oV9Gbd7kMXwHX2UGUEFX6mDMnxOgHdCt3lN0ak81DXunMQEwTKL//v/WKkuzwRhb
# odNRPK0Ncmy+rZSkhAr0JAW0LZ7owAfvxKDylq7SJ0UWGSXcslQwpuTp5o8SOb1q
# 6IcGJ1OaYDDqf0DrK73KHpf0u2fRuyQ3VRaDelhz0aaadHXlj5HY/tQqJnKrnNfQ
# nJOpa6X5KQf2YYW6Ey02FZMYgmWJVkh6iQ8cW5dWeaJ6kMKACX8pOkS3ruXPEW34
# rx3GcRS7iFtgS597zWUU8cPkowWMZf+8K8D/TrTxPLFLNT4mNqStbs94uLL9IOVC
# SbaMlipDOCoC9dq9G1v5Xg+TyPrYwxszTm9Dl4SKLSpZcwtGCEZuyjCyWwyrC+yv
# X3ECF1NQDqL0ey0sOtgYwsHtKXqyHUVwpQmxlDE5Nsw6SALN1QGgj99QvjUPybTC
# qhzov3tr1N1DPhLMDSZE/vyp8vmfdIZmBIDPS9Lzt5aPkfZHxBVEf7OPp6Eqm7ys
# 9d58Nwhx6Q==
# SIG # End signature block
