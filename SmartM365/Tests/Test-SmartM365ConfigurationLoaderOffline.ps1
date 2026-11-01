<#
.SYNOPSIS
    Offline common configuration contract and complete collector wiring checks.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Config/SmartM365-TenantContext.ps1')
. (Join-Path $root 'Modules/SmartM365.Core/SmartM365-SharePointUploadPolicy.ps1')
$script:checks=0
function Check([bool]$Pass,[string]$Message) { $script:checks++; if(-not $Pass){throw $Message} }
function Reject([scriptblock]$Action,[string]$Message) { $failed=$false;try { & $Action | Out-Null }catch{$failed=$true};Check $failed $Message }
$values=@($null,'','  ','__USE_GLOBAL__','USE_GLOBAL',$false,$true)
foreach($category in @('DataAll','DataLast','DataPowerBI','LogAll')) {
    $key='EnableSharePointUpload'+$category
    foreach($globalValue in $values){foreach($tenantValue in $values){foreach($localValue in $values){foreach($legacy in @($false,$true)){
        $effective=Merge-SmartM365Configuration -EffectiveConfig @{$key=$globalValue} -LocalConfig @{$key=$tenantValue}
        $merged=Merge-SmartM365Configuration -EffectiveConfig $effective -LocalConfig @{$key=$localValue}
        $expected=$legacy
        foreach($candidate in @($localValue,$tenantValue,$globalValue)){if($candidate -is [bool]){$expected=$candidate;break}}
        $policy=Resolve-SmartM365SharePointUploadPolicy -Category $category -LocalConfig $merged -LegacyEnabled $legacy
        Check ($policy.Enabled -eq $expected) "Three-tier precedence: $category"
    }}}}
}
foreach($config in @(@{Flag=$false;Zero=0;Array=@('a','b');Empty=@();Map=@{Nested=$false}},[pscustomobject]@{Flag=$false;Zero=0;Array=@('a','b');Empty=@();Map=@{Nested=$false}})) {
    Check ((Get-SmartM365ConfigurationValue Flag -LocalConfig $config -DefaultValue $true) -eq $false) 'False lost'
    Check ((Get-SmartM365ConfigurationValue Zero -LocalConfig $config -DefaultValue 9) -eq 0) 'Zero lost'
    Check ((Get-SmartM365ConfigurationValue Array -LocalConfig $config).Count -eq 2) 'Array lost'
    $empty=Get-SmartM365ConfigurationValue Empty -LocalConfig $config
    Check ($null -ne $empty -and $empty.Count -eq 0) 'Empty array lost'
    Check ((Get-SmartM365ConfigurationValue Map -LocalConfig $config).Nested -eq $false) 'Nested object lost'
}
$tokens=[pscustomobject]@{Root='fixture';Child='{{Root}}/child';False=$false}
Check ((Resolve-SmartM365ConfigurationTokens '{{Child}}/file' $tokens) -eq 'fixture/child/file') 'Recursive tokens'
Check ((Resolve-SmartM365ConfigurationTokens '{{Missing}}' $tokens) -eq '{{Missing}}') 'Legacy unresolved token preservation'
Reject { Resolve-SmartM365ConfigurationTokens '{{Missing}}' $tokens -Strict } 'Strict unresolved token accepted'
Reject { Resolve-SmartM365ConfigurationTokens '{{A}}' @{A='{{B}}';B='{{A}}'} } 'Cycle accepted'
Reject { Resolve-SmartM365ConfigurationTokens '{{A}}' @{A='{{A}}suffix'} } 'Expanding cycle accepted'
$local=@{Map=@{OnlyLocal=1};Blank='';Null=$null;Flag=$false}
$base=@{Map=@{OnlyBase=1};Blank='base';Null='base';Flag=$true}
$merged=Merge-SmartM365Configuration $base $local
Check ($merged.Map.Contains('OnlyLocal') -and -not $merged.Map.Contains('OnlyBase')) 'Unexpected deep merge'
Check ($merged.Blank -eq 'base' -and $merged.Null -eq 'base' -and -not $merged.Flag) 'Default merge contract'
$overlay=Merge-SmartM365Configuration $base $local -Mode Overlay
Check ($overlay.Blank -eq '' -and $null -eq $overlay.Null) 'Exchange overlay contract'
Check ((Get-SmartM365ConfigurationValue Missing -EffectiveConfig $tokens -DefaultValue '{{Root}}') -eq '{{Root}}') 'Literal default changed'
Check ((Get-SmartM365ConfigurationValue Missing -EffectiveConfig $tokens -DefaultValue '{{Root}}' -ResolveDefault) -eq 'fixture') 'Resolved default changed'
for($i=0;$i -lt 3;$i++) {
    $first=Merge-SmartM365Configuration @{TenantKey='a';Flag=$true} @{Flag=$false}
    $second=Merge-SmartM365Configuration @{TenantKey='b';Flag=$true} @{}
    Check ($first.TenantKey -eq 'a' -and -not $first.Flag -and $second.TenantKey -eq 'b' -and $second.Flag) 'Cross-tenant state leak'
}

$fixture=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CommonConfig-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixture
try {
    $scriptPath=Join-Path $fixture 'Synthetic.ps1'
    $path=Join-Path $fixture 'Synthetic.local.json.txt'
    $template=Join-Path $fixture 'Synthetic.local.json.template'
    [IO.File]::WriteAllText($template,'{"Nested":{"Keep":1,"Add":2},"Array":["one"],"Empty":[],"EnableSharePointUploadDataAll":"__USE_GLOBAL__"}')
    $onlyTemplate=Read-SmartM365ScriptConfiguration -ScriptPath $scriptPath -MissingPolicy Template
    Check (-not (Test-Path -LiteralPath $path)) 'Template policy created a private file'
    Check ($onlyTemplate.Nested -is [pscustomobject] -and $onlyTemplate.Nested.PSObject.Properties['Keep'].Value -eq 1) 'Nested JSON object changed type'
    Check ($onlyTemplate.Array -is [array] -and $onlyTemplate.Array.Count -eq 1 -and $onlyTemplate.Empty.Count -eq 0) 'JSON arrays changed shape'
    $created=Read-SmartM365ScriptConfiguration -ScriptPath $scriptPath
    Check (Test-Path -LiteralPath $path -PathType Leaf) 'Create policy did not initialize configuration'
    [IO.File]::WriteAllText($path,'{"Nested":{"Keep":false},"EnableSharePointUploadDataAll":false}')
    $enriched=Read-SmartM365ScriptConfiguration -ScriptPath $scriptPath
    Check ($enriched.Nested.Keep -eq $false -and $enriched.Nested.Add -eq 2 -and -not $enriched.EnableSharePointUploadDataAll) 'Template enrichment overwrote explicit values'
    Check ((Read-SmartM365ScriptConfiguration -ScriptPath $scriptPath -AsHashtable) -is [Collections.IDictionary]) 'Dictionary reader contract'
    [IO.File]::WriteAllText($path,'invalid')
    Reject {Read-SmartM365ScriptConfiguration -ScriptPath $scriptPath -MissingPolicy Template} 'Invalid private JSON fell back to template'
    $missing=Join-Path $fixture 'Missing.ps1'
    $optional=Read-SmartM365ScriptConfiguration -ScriptPath $missing -MissingPolicy Optional
    Check (@($optional.PSObject.Properties).Count -eq 0) 'Optional missing config not empty'
    $optional=Read-SmartM365ScriptConfiguration -ScriptPath $missing -MissingPolicy CreateIfTemplate
    Check (@($optional.PSObject.Properties).Count -eq 0) 'HealthCheck optional legacy contract'

    # Exercise actual global/tenant loading and source metadata in an isolated repository.
    $configDir=Join-Path $fixture 'Config'
    $null=New-Item -ItemType Directory -Path (Join-Path $configDir 'Tenants'),(Join-Path $fixture 'SmartInventory/Area') -Force
    $globalDocument=@{WorkspaceRootPath=$fixture;GlobalOnly='global';TenantOverride='global';LocalOverride='global';EnableSharePointUpload=$true;EnableSharePointUploadDataAll=$true}
    $tenantDocument=@{ProfileKey='fixture';OrganizationKey='example';EnvironmentKey='test';TenantKey='example-test';TenantOverride='tenant';EnableSharePointUploadDataAll=$false}
    [IO.File]::WriteAllText((Join-Path $configDir 'SmartM365.global.local.json.txt'),($globalDocument|ConvertTo-Json))
    [IO.File]::WriteAllText((Join-Path $configDir 'Tenants/fixture.local.json.txt'),($tenantDocument|ConvertTo-Json))
    $entry=Join-Path $fixture 'SmartInventory/Area/Fixture.ps1'
    [IO.File]::WriteAllText((Join-Path $fixture 'SmartInventory/Area/Fixture.local.json.txt'),'{"LocalOverride":"local","EnableSharePointUploadDataAll":"__USE_GLOBAL__"}')
    $context=Get-SmartM365ScriptConfigurationContext -ScriptPath $entry -Tenant fixture
    Check ($context.Sources.GlobalOnly -eq 'Global' -and $context.Sources.TenantOverride -eq 'Tenant' -and $context.Sources.LocalOverride -eq 'Local') 'Configuration provenance'
    Check ($context.EffectiveConfig.LocalOverride -eq 'local' -and $context.EffectiveConfig.TenantOverride -eq 'tenant' -and -not $context.EffectiveConfig.EnableSharePointUploadDataAll) 'Actual loader precedence'
    Check ($context.LocalPath.EndsWith('Fixture.local.json.txt') -and $context.TenantPath.EndsWith('fixture.local.json.txt')) 'Resolved source paths'
} finally {
    $full=[IO.Path]::GetFullPath($fixture)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')+[IO.Path]::DirectorySeparatorChar
    if(-not $full.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($full) -notmatch '^SmartM365-CommonConfig-[a-f0-9]{32}$'){throw 'Unsafe test cleanup'}
    Remove-Item -LiteralPath $full -Recurse -Force
}

# Inspect every publication route, without running any collector or loading private JSON.
$collectors=0;$shared=0;$direct=0
foreach($file in Get-ChildItem (Join-Path $root 'SmartInventory') -Filter '*.ps1' -Recurse -File) {
    if($file.FullName -match '[\\/]Tests[\\/]|[\\/]Orchestrator[\\/]|[\\/]PreparedEvidence[\\/]'){continue}
    $source=[IO.File]::ReadAllText($file.FullName)
    if($source -notmatch 'Initialize-SmartM365EvidenceRuntime|Initialize-SmartM365CollectorUploadContext|Resolve-SmartM365CollectorUploadPolicies'){continue}
    $collectors++
    if($source -match 'Initialize-SmartM365EvidenceRuntime'){$shared++;continue}
    $direct++
    Check ($source -match 'Read-SmartM365ScriptConfiguration') "Missing common reader: $($file.Name)"
    Check ($source -notmatch 'Get-Content[^\r\n]*\$(?:script:)?\w*Config\w*[^\r\n]*ConvertFrom-Json') "Independent reader: $($file.Name)"
    $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$t,[ref]$e)
    Check ($e.Count -eq 0) "Parse: $($file.Name)"
    foreach($node in $ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Get-ScriptLocalConfig','Get-ScriptLocalConfigValue','Resolve-SmartM365ConfigValue','Merge-SmartM365ScriptLocalConfig','Get-SmartM365ConfigValue','Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Resolve-ConfigToken','Get-LocalConfigValue','Get-ConfigValue','Get-TeamsPhoneConfigValue','Get-QuarantineConfigValue','Get-PowerBIActivityRawConfigValue','Get-EARawConfigValue')},$true)) {
        Check ($node.Body.Extent.Text -match 'SmartM365(?:ScriptConfiguration|Configuration)') "Independent resolver: $($file.Name)/$($node.Name)"
        Check ($node.Body.Extent.Text -notmatch '\b(for|foreach|while)\s*\(') "Duplicated resolver loop: $($file.Name)/$($node.Name)"
    }
}
Check ($collectors -eq 51 -and $shared -eq 11 -and $direct -eq 40) "Coverage changed: $collectors collectors, $shared shared, $direct direct"
foreach($relative in @('SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1','SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1','../SmartWorkplaceIntelligence/scripts/Invoke-PreparedSharePointRetention.ps1','SmartInventory/Common/SmartM365.EvidenceCollector.Common.psm1','Modules/SmartM365.Core/SmartM365.Core.psm1','Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
    $source=[IO.File]::ReadAllText((Join-Path $root $relative))
    Check ($source -match 'Read-SmartM365ScriptConfiguration') "Missing infrastructure reader: $relative"
}
"PASS: $script:checks common configuration checks; 51 collectors, orchestrator, Prepare, retention and both Core editions; no tenant access."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCACH6wJxC+wawcs
# lEamLGYkoqjgkUp9jBQS4kshuWbGoKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPzZmnTIKQa1yUFCIyFJSDqvxE705Qq0JGNz6uJbc95oMA0GCSqG
# SIb3DQEBAQUABIIBgGDEkNlIutn0irpDsuASbpafbXTTsVgIN3AzhfpeAegCl8vc
# bG304U+VPNJbgINsNh9e2N9jBPvnSVfwP8f5BA/UsAOLm7ZPx34UYFmiumBRgoHL
# 5JaO9bOp3etDFI6t9PPBFqhvCOjmPmP6zYwpu/h+aaTJDnkEJ94WA/ci77GcyE+U
# 9p+ZBiChHNRzrg/NrVtRPJaYIICFge6lMTHb3bLv4nXcQ/EjyG6GRFKYwtfURvW2
# joHq4HDeOz2YC/4ezVAU7vZULW96DhLw2UoVHrqS+C/nKvmWdJshvLSS4Npu6KcX
# 8aUniWG+Ax7vmpE2VM83Hrkse5OA6IOZOIZPgY/CALWu/MGXU53qu/vNVo9BjNDK
# ejRReioMAmVIdF1DdWhQKeMaNd2lyaBBK4+vPVybdjYbaIMzAQi3HDtcFU/cB0TJ
# Cui8CtPlaXmuaZ6SxIbkjbcFnEkDReTtvUAyobNzOwnmXR1KoHr+chlyXMASJspJ
# PVQy5ewgAcLuB4gP1qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDExMzA1
# MzVaMC8GCSqGSIb3DQEJBDEiBCB5+9aToHbaOr7olwZ1IlpiQxxtpJlA04Tk1HwK
# l7LhWDANBgkqhkiG9w0BAQEFAASCAgArY4bUj0LtSZF/r7LvQ77UveRNPa42hV81
# L5t1phUmHEvmJ0yXAFeSnT01kup0P699sPGw65W+o9hE6LaBFuVIthz/rc+zvHbQ
# KTPCPAiB3q+r4TTmw7MhNWCHF3p8pMn+qIMYdyKylIRwlooFN0iTVnsoGlQXydMh
# rZkWA8Djuj3tdRhqCm9zxLrYJ0vxH3Xou5VeIR/kbBTQRpB+dW+C92fGTRz6b7ot
# UBEbMaDFuUlIqkJvWmHjmD+Nn4T1Lm/x13py2ToKTufgN7TlBCIiJk3DI+LhflaX
# OfKxSiphhRGHuBuBEdxpdIsf96D90zktdg/9Ilpvz5z8ufuGJT3iKOQfbgAkpWYV
# e68UG1U8bl1Zt3UJeCqLGpKz+KWqE3xchJVlGV27qjKS0Oy47xEwQzOSvlIb1/wu
# 6IpZMaWMmiNXTUbBxWAqqPEdHES2qXWi9PB9wr6JYyq5Q3yC4cIpVKvrjr/z2xNy
# mjw0wqIUpvpyd5Vhex0Yxb39OdyIZW0clAW4kSca37MI9nh+oJJfQ0p9Z5uCsD/T
# IJvFOsc5lJ0cFvc9aHNzMnNdOYMcY91DKERbNZhKXFz+6AixY9+z3pu7pi3LwbaL
# EoWjGReGiOdzHzTj73Qbo7wlNfve/6QwTfBlhp8BG5htSjEOGbIsmpzIuuynJj8k
# 3KtKLPoOjA==
# SIG # End signature block
