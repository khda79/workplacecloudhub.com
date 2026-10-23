#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests for profile-specific script mail configuration and inheritance.
.VERSION
1.0.0
.NOTES
Uses synthetic TEMP profiles and extracts the HealthCheck config resolver only.
No collector, authentication, mail, upload or production configuration is run.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='TenantContext and ScriptConfig are consumed by the extracted HealthCheck resolver.')]
[CmdletBinding()]
param([string]$FixtureParent = [IO.Path]::GetTempPath())
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
. (Join-Path $projectRoot 'Config/SmartM365-TenantContext.ps1')
$module = Get-Module SmartM365.JsonTransport
$originalPolicy = & $module { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$fixtureRoot = Join-Path ([IO.Path]::GetFullPath($FixtureParent)) ('SM-Mail-' + [guid]::NewGuid().ToString('N').Substring(0,12))
$tenantFolder = Join-Path $fixtureRoot 'Config/Tenants'
$null = New-Item -ItemType Directory -Path $tenantFolder -Force
$script:passed = 0
function Check([bool]$Value, [string]$Message) {
    if (-not $Value) { throw $Message }
    $script:passed++
}
function Write-FixtureJson([string]$Path, $Document) {
    [IO.File]::WriteAllText($Path, ($Document | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}
function Reject([scriptblock]$Action, [string]$Message) {
    $failed = $false
    try { & $Action | Out-Null } catch { $failed = $true }
    Check $failed $Message
}
$healthName = 'SmartM365-ActiveDirectory-HealthCheck'
$teamsName = 'SmartM365-Teams-Inventory'
$mailPath = Join-Path $tenantFolder 'prod.mail.local.json.txt'
$tenantPath = Join-Path $tenantFolder 'prod.local.json.txt'
$globalPath = Join-Path $fixtureRoot 'Config/SmartM365.global.local.json.txt'
$mailTemplate = Join-Path $tenantFolder 'tenant.mail.local.json.template'
try {
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
    Write-FixtureJson $globalPath @{WorkspaceRootPath=$fixtureRoot;From='global@example.com';Cc='global-copy@example.com'}
    $tenantFixture = @{ProfileKey='prod';OrganizationKey='example';EnvironmentKey='prod';TenantKey='example-prod';From='prod@example.com';Cc='prod-copy@example.com'}
    Write-FixtureJson $tenantPath $tenantFixture
    Write-FixtureJson $mailTemplate @{Scripts=@{
        $healthName=@{From='__USE_GLOBAL__';Cc='__USE_GLOBAL__'}
        $teamsName=@{From='__USE_GLOBAL__';Cc='__USE_GLOBAL__'}
    }}
    Check ((Get-SmartM365JsonTemplatePath $mailPath) -eq $mailTemplate) 'Mail file selected the tenant identity template.'
    $effective = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
    Check ($effective.From -eq 'prod@example.com' -and $effective.Cc -eq 'prod-copy@example.com') 'Absent file changed common mail settings.'
    Check (-not (Test-Path -LiteralPath $mailPath)) 'Optional mail file was created implicitly.'
    $tenantHash = (Get-FileHash $tenantPath).Hash
    $globalHash = (Get-FileHash $globalPath).Hash
    $null = Initialize-SmartM365LocalJsonFromTemplate -Path $mailPath
    Check ((Get-Content $mailPath -Raw | ConvertFrom-Json).Scripts.$healthName.From -eq '__USE_GLOBAL__') 'Mail initialization did not use its own template.'
    Check (-not ((Read-SmartM365JsonConfig $mailPath).Contains('ProfileKey'))) 'Mail file enriched with tenant identity keys.'
    $settings = @{Scripts=@{$healthName=@{From=' ad@example.com ';Cc='ad-copy@example.com';To='ignored@example.com'}}}
    Write-FixtureJson $mailPath $settings
    $effective = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName "$healthName.ps1"
    Check ($effective.From -eq 'ad@example.com' -and $effective.Cc -eq 'ad-copy@example.com') 'Dedicated fields were not applied independently or trimmed.'
    Check ($null -eq $effective.PSObject.Properties['To']) 'Mail entry changed To.'
    Check ((Get-Content $mailPath -Raw | ConvertFrom-Json).Scripts.$teamsName.From -eq '__USE_GLOBAL__') 'Missing script template entry was not added.'
    $mailHash = (Get-FileHash $mailPath).Hash
    $null = Read-SmartM365JsonConfig $mailPath
    Check ((Get-FileHash $mailPath).Hash -eq $mailHash) 'Unchanged mail file was rewritten.'
    $other = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $teamsName
    Check ($other.From -eq 'prod@example.com' -and $other.Cc -eq 'prod-copy@example.com') 'One script contaminated another script context.'
    foreach ($marker in @('__USE_GLOBAL__', 'USE_GLOBAL', '', '  ', $null)) {
        foreach ($field in @('From', 'Cc')) {
            $entry = @{From='ad@example.com';Cc='ad-copy@example.com'}
            $entry[$field] = $marker
            Write-FixtureJson $mailPath @{Scripts=@{$healthName=$entry}}
            $result = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
            Check ($result.$field -eq $tenantFixture[$field]) "Dedicated $field marker did not inherit."
            $otherField = if ($field -eq 'From') {'Cc'} else {'From'}
            Check ($result.$otherField -eq $entry[$otherField]) 'Independent mail field was lost.'
        }
    }
    foreach ($field in @('From', 'Cc')) {
        Write-FixtureJson $mailPath @{Scripts=@{$healthName=@{}}}
        $result = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
        Check ($result.$field -eq $tenantFixture[$field]) 'Missing dedicated key did not inherit.'
        foreach ($marker in @('__USE_GLOBAL__', 'USE_GLOBAL', '', '  ', $null)) {
            $tenantFixture[$field] = $marker
            Write-FixtureJson $tenantPath $tenantFixture
            $result = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
            $expected = if ($field -eq 'From') {'global@example.com'} else {'global-copy@example.com'}
            Check ($result.$field -eq $expected) 'Common tenant marker did not inherit global value.'
        }
        $tenantFixture.Remove($field)
        Write-FixtureJson $tenantPath $tenantFixture
        $result = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
        Check ($result.$field -eq $expected) 'Missing common key did not inherit global value.'
        $tenantFixture[$field] = if ($field -eq 'From') {'prod@example.com'} else {'prod-copy@example.com'}
    }
    Write-FixtureJson $tenantPath $tenantFixture
    Write-FixtureJson $mailPath $settings
    # Exercise real call-stack discovery in a function inside a script, without a collector.
    $callerPath = Join-Path $fixtureRoot "$healthName.ps1"
    [IO.File]::WriteAllText($callerPath, @'
param($Root)
function Load-Context { Initialize-SmartM365TenantContext -StartPath $Root -Tenant prod }
Load-Context
'@)
    $result = & $callerPath $fixtureRoot
    $sharedConfig = Get-Variable -Name SmartM365GlobalConfig -Scope Global -ValueOnly
    Check ($result.From -eq 'ad@example.com' -and $sharedConfig.Cc -eq 'ad-copy@example.com') 'Caller discovery or shared global context failed.'
    $directCaller = Join-Path $fixtureRoot "$teamsName.ps1"
    [IO.File]::WriteAllText($directCaller, 'param($Root) Get-SmartM365EffectiveGlobalConfig -StartPath $Root -ProfileKey prod')
    $result = & $directCaller $fixtureRoot
    Check ($result.From -eq 'prod@example.com') 'Direct caller reused previous script identity.'
    $testProfile = @{ProfileKey='test';OrganizationKey='example';EnvironmentKey='test';TenantKey='example-test';From='test@example.com';Cc='test-copy@example.com'}
    Write-FixtureJson (Join-Path $tenantFolder 'test.local.json.txt') $testProfile
    Write-FixtureJson (Join-Path $tenantFolder 'test.mail.local.json.txt') @{Scripts=@{$healthName=@{From='test-ad@example.com'}}}
    $result = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey test -ScriptName $healthName
    Check ($result.From -eq 'test-ad@example.com' -and $result.Cc -eq 'test-copy@example.com') 'Profile isolation failed.'
    # Run only the real HealthCheck's pure local-config resolver against synthetic objects.
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $projectRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1'), [ref]$tokens, [ref]$errors)
    $functions = $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in @('Resolve-ConfigToken','Get-LocalConfigValue')}, $true)
    $resolver = [scriptblock]::Create(($functions | ForEach-Object {$_.Extent.Text}) -join "`n")
    $effective = Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName
    & {
        . $resolver
        $TenantContext = $effective
        foreach ($marker in @('__USE_GLOBAL__','USE_GLOBAL','', ' ', $null)) {
            $ScriptConfig = [pscustomobject]@{From=$marker;Cc=$marker}
            Check ((Get-LocalConfigValue From '') -eq 'ad@example.com') 'HealthCheck local From marker did not inherit dedicated value.'
            Check ((Get-LocalConfigValue Cc '') -eq 'ad-copy@example.com') 'HealthCheck local Cc marker did not inherit dedicated value.'
        }
        $ScriptConfig = [pscustomobject]@{From='local@example.com';Cc='local-copy@example.com'}
        Check ((Get-LocalConfigValue From '') -eq 'local@example.com') 'Explicit local From lost priority.'
        Check ((Get-LocalConfigValue Cc '') -eq 'local-copy@example.com') 'Explicit local Cc lost priority.'
        $ScriptConfig = [pscustomobject]@{}
        Check ((Get-LocalConfigValue From '') -eq 'ad@example.com') 'Absent local From did not inherit dedicated value.'
        Check ((Get-LocalConfigValue Cc '') -eq 'ad-copy@example.com') 'Absent local Cc did not inherit dedicated value.'
    }
    foreach ($invalid in @('{','[]','null','{"Scripts":null}','{"Scripts":[]}',('{"Scripts":{"'+$healthName+'":null}}'),('{"Scripts":{"'+$healthName+'":{"From":[]}}}'),('{"Scripts":{"'+$healthName+'":{"Cc":false}}}'))) {
        [IO.File]::WriteAllText($mailPath, $invalid)
        Reject { Get-SmartM365EffectiveGlobalConfig -StartPath $fixtureRoot -ProfileKey prod -ScriptName $healthName } 'Invalid mail config silently fell back.'
    }
    Check ((Get-FileHash $tenantPath).Hash -eq $tenantHash) 'Common profile contents changed.'
    Check ((Get-FileHash $globalPath).Hash -eq $globalHash) 'Global config contents changed.'
    [pscustomobject]@{Passed=$script:passed;Evidence='Synthetic configuration and extracted HealthCheck resolver only'}
} finally {
    & $module {param($original) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $original} $originalPolicy
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath($FixtureParent).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $resolved -Leaf) -notmatch '^SM-Mail-[a-f0-9]{12}$') { throw 'Unsafe fixture cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD2x5NzEh8ueLKa
# VsA2cLdZNyZRwb+bFvoGyikgvEYx1qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE30v7CiasoFSsJ9NRY7kn94scn2yrZf5z08PqBp7KlpMA0GCSqG
# SIb3DQEBAQUABIIBgFVdWVlSsigVWf5xb/n/SA4h2L+E5ZH9eZQz2bSZxBIJAnkK
# MhNTHmlsC2kSWhwVFmtLVw3ZXOBsFad1cP/m4Pj79/I5j7jp7Q11QgwcPlK9rhVL
# B+p8Tjdy8NI0ErmbfpxH8GrSUfamHaS/9zDGmXy19tFK4JKJFbcgZ9f6ms67aNXa
# NRicDpe9YqtQBTcxy7DPLtDB0cW8H9xEEH9gy6qiEF7JH8PB+lpKBgdrVB477bR3
# Fn7CSOqwUI1iXhaX5OBwI2wHnN5bKARqFy19wsERoe3tetXcnKKe9YsRUGl4WEp8
# jiF0ecvNgqDtoo/o301fR9f93+iuDDtnNKeoY9ZeNFxL/gYOjI5q20TvYPfVrWP9
# 8P7wiqw01SebtlXqMtPKeciGuRxOW+7JAi2MwPvLVpZf6NduOVUZHo7RD3OJ74p8
# d7qEgvgC3i1tAXhh6oV3d2N25XLeLW8aG+H0QKveXP21hLNvdIrB+FLkGC/I030f
# 1Fj4JZJE2WDRkEmIvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIyMzAz
# MzRaMC8GCSqGSIb3DQEJBDEiBCCYj490iNjap+1uP4kyO3YdxwbJaw/LCecuIrpx
# 9wcW3TANBgkqhkiG9w0BAQEFAASCAgCnEmIBpnb6GModrhiVKgwo2+kUxEXAJWgG
# jgZFDRpTF9IKvUXVNeFIXPLuKwA3roJxB05GuEwxtATf2aOqJTC4jj1/lLddTUig
# KmlYwSI46JH0lU+9PbrO+vEYlNOEPSvUx80zbowImPJItYCNEoIe3Cp16lyWUa+v
# JS85qsH+9bOKTuKexvxUAWVK2ShudjbkYTl5QwHQrrGFZgvSpN57D9POYfTmP/z+
# fkR/gYfCNFDep865hSKXKvuZYzo6w5zExeAQDDGvfuX8GdA1WDd/wb8nMGHgXOEP
# X/OndjsgBi8jr8qsccxD8KakFmA2yGzhejyQ9zIzUALu8ypXGjPrNXe7gjK9P1O+
# zcNiC3t6+faf3/arLhFuzND4EbBVGbhJN4rwS+v2m9h0vEgpVAcJCDUyokfcioM0
# LiH9E42tHpwsLtJ7sFBZKj4EaXFZNr+Ui/2HAICfuchQAMbnifzAqAo8XXuhj7D4
# 45cc+mUiGv7GJeXcl5M0RyELsaKZjt18TKFQDjvFaheLQAqve5m/Fy3Tacc5Gxzy
# DeLJ4snubwXcpK7Pn78Wp3tVwhRrz/GTARvn6I57a7VQYYziT5tBHUQ7kSApKF+9
# EMKpYrm4yN5kC2YGzAM2RuS6IQX1awHCwRX0ECOF7PUHZ/OCMqr+bH6R4L3fwxlP
# WFb+K5p9Wg==
# SIG # End signature block
