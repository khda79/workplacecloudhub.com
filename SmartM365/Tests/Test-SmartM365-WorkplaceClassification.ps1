<#
.SYNOPSIS
    Validates the shared site and persona classification module with synthetic workbooks and,
    when given, the TestCases worksheets of the private SmartWorkplaceIntelligence workbooks.
.VERSION
1.0
#>

[CmdletBinding()]
param(
    [string]$PersonaClassificationPath = '',
    [string]$SiteClassificationPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartInventoryRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory'
Import-Module (Join-Path $smartInventoryRoot 'Common\SmartM365.WorkplaceClassification.psd1') -MinimumVersion '1.0.0' -Force -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
$labels = Get-SmartM365WorkplacePersonaLabel
$names = Get-SmartM365WorkplaceClassificationWorkbookName

function Assert-Classification {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAILED: $Message" }
    Write-Output "PASS: $Message"
}

function Assert-Rejected {
    param([scriptblock]$Action, [string]$Message, [string]$Like = '*')
    $rejected = $false
    try { & $Action | Out-Null } catch { $rejected = $_.Exception.Message -like $Like }
    Assert-Classification $rejected $Message
}

function Export-SyntheticWorkbook {
    param([string]$Path, [hashtable]$Sheets)
    # Site codes stay text ('000001'), as in the governed workbooks.
    foreach ($sheet in $Sheets.Keys) { $Sheets[$sheet] | Export-Excel -Path $Path -WorksheetName $sheet -NoNumberConversion '*' }
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('WorkplaceClassification-{0}' -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    $sitePath = Join-Path $root $names.Site
    Export-SyntheticWorkbook -Path $sitePath -Sheets ([ordered]@{
        Settings = @([pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' }, [pscustomobject]@{ Parameter = 'Headquarters Persona'; Value = 'Headquarters staff' })
        SiteTypes = @('HQ','CLINIC','NURSING HOME','OTHER','UNKNOWN' | ForEach-Object { [pscustomobject]@{ 'Site Type Code' = $_ } })
        Sites = @(
            [pscustomobject]@{ 'Directory Site Code' = '000001'; 'Site Display Name' = 'Head office'; 'Site Type' = 'HQ'; 'Integration Status' = 'INTEGRATED'; 'Use for Classification' = 'Yes' }
            [pscustomobject]@{ 'Directory Site Code' = '000002'; 'Site Display Name' = 'Clinic A'; 'Site Type' = 'CLINIC'; 'Integration Status' = 'NOT-INTEGRATED'; 'Use for Classification' = 'Yes' }
            [pscustomobject]@{ 'Directory Site Code' = '000003'; 'Site Display Name' = 'Retired site'; 'Site Type' = 'OTHER'; 'Integration Status' = 'INTEGRATED'; 'Use for Classification' = 'No' }
        )
    })
    $personaPath = Join-Path $root $names.Persona
    Export-SyntheticWorkbook -Path $personaPath -Sheets ([ordered]@{
        Personas = @(
            [pscustomobject]@{ 'Persona ID' = 'HQ'; 'Persona name' = 'Headquarters staff'; Enabled = 'Yes' }
            [pscustomobject]@{ 'Persona ID' = 'ADMIN'; 'Persona name' = 'Management and Administrative staff'; Enabled = 'Yes' }
            [pscustomobject]@{ 'Persona ID' = 'CARE'; 'Persona name' = 'Healthcare staff'; Enabled = 'Yes' }
        )
        Rules = @(
            [pscustomobject]@{ 'Rule ID' = 'R1'; Keyword = 'accountant'; Country = 'ALL'; 'Persona ID' = 'ADMIN'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R2'; Keyword = 'nurse'; Country = 'ALL'; 'Persona ID' = 'CARE'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R3'; Keyword = 'manager'; Country = 'FR'; 'Persona ID' = 'ADMIN'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R4'; Keyword = 'infirmiere'; Country = 'FR'; 'Persona ID' = 'CARE'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
        )
        Exclusions = @([pscustomobject]@{ 'Rule ID' = 'X1'; Keyword = 'scan to mail'; Country = 'ALL'; Enabled = 'Yes' })
    })

    $site = Read-SmartM365WorkplaceSiteClassification -Path $sitePath
    $persona = Read-SmartM365WorkplacePersonaClassification -Path $personaPath -HeadquartersPersona $site.HeadquartersPersona
    $hq = Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode ' 000001 '
    Assert-Classification ($hq.Type -eq 'HQ' -and $hq.Name -eq 'Head office' -and $hq.IntegrationStatus -eq 'INTEGRATED') 'Mapped site code returns its type, name and integration status.'
    Assert-Classification ($null -eq (Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode '000003')) 'A site with Use for Classification = No is not mapped.'
    Assert-Classification ($null -eq (Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode '')) 'A blank site code is not mapped.'

    $cases = @(
        @{ Title = 'Accountant'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = 'Management and Administrative staff'; Name = 'Job title keyword' }
        @{ Title = 'Accountant'; Description = ''; Country = 'FR'; Site = 'HQ'; Expected = 'Headquarters staff'; Name = 'HQ site has priority over job title' }
        @{ Title = 'Scan to mail'; Description = ''; Country = 'FR'; Site = 'HQ'; Expected = $labels.NonHuman; Name = 'Non-human exclusion has priority over the HQ site' }
        @{ Title = 'Nurse manager'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.Ambiguous; Name = 'Equal-priority matches across personas stay ambiguous' }
        @{ Title = 'Nurse manager'; Description = ''; Country = 'DE'; Site = 'CLINIC'; Expected = 'Healthcare staff'; Name = 'Country-specific rule applies only to its country' }
        @{ Title = 'Infirmiere'; Description = ''; Country = 'France'; Site = 'CLINIC'; Expected = 'Healthcare staff'; Name = 'Country name is converted to its code' }
        @{ Title = 'Unmapped position'; Description = 'Site - Accountant'; Country = 'FR'; Site = 'CLINIC'; Expected = 'Management and Administrative staff'; Name = 'AD description is used when the title has no match' }
        @{ Title = 'Unmapped position'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.NoMatch; Name = 'No match stays unclassified' }
        @{ Title = ''; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.MissingTitle; Name = 'Missing title stays unclassified' }
    )
    foreach ($case in $cases) {
        $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle $case.Title -Description $case.Description -Country $case.Country -SiteType $case.Site
        Assert-Classification ($result.Persona -eq $case.Expected) ("{0} ({1})." -f $case.Name, $result.Persona)
    }

    $broken = Join-Path $root 'broken.xlsx'
    [pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' } | Export-Excel -Path $broken -WorksheetName 'Settings'
    $brokenSite = Join-Path -Path $root -ChildPath 'site-broken' -AdditionalChildPath $names.Site
    New-Item -ItemType Directory -Path (Split-Path $brokenSite) -Force | Out-Null
    Copy-Item -LiteralPath $broken -Destination $brokenSite
    Assert-Rejected -Action { Read-SmartM365WorkplaceSiteClassification -Path $brokenSite } -Message 'A workbook without the required worksheets is rejected.' -Like '*Required worksheet missing*'
    $duplicateSite = Join-Path -Path $root -ChildPath 'site-duplicate' -AdditionalChildPath $names.Site
    New-Item -ItemType Directory -Path (Split-Path $duplicateSite) -Force | Out-Null
    Export-SyntheticWorkbook -Path $duplicateSite -Sheets ([ordered]@{
        Settings = @([pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' }, [pscustomobject]@{ Parameter = 'Headquarters Persona'; Value = 'Headquarters staff' })
        SiteTypes = @('HQ','UNKNOWN' | ForEach-Object { [pscustomobject]@{ 'Site Type Code' = $_ } })
        Sites = @('000001','000001' | ForEach-Object { [pscustomobject]@{ 'Directory Site Code' = $_; 'Site Display Name' = 'Duplicate'; 'Site Type' = 'HQ'; 'Use for Classification' = 'Yes' } })
    })
    Assert-Rejected -Action { Read-SmartM365WorkplaceSiteClassification -Path $duplicateSite } -Message 'A duplicate enabled site code is rejected.' -Like '*duplicate*'
    Assert-Rejected -Action { Read-SmartM365WorkplacePersonaClassification -Path $personaPath -HeadquartersPersona 'Disabled persona' } -Message 'An HQ persona missing from the enabled personas is rejected.' -Like '*Headquarters Persona*'
    $receiveFolder = Join-Path $root 'received'
    Assert-Rejected -Action { Receive-SmartM365WorkplaceClassificationWorkbook -Folder $receiveFolder -DownloadFile { param($destination, $name) Write-Verbose ("Simulated failed download: {0} -> {1}" -f $name, $destination) } } -Message 'A failed download stops without any cached fallback.' -Like '*No cached fallback*'
    $received = Receive-SmartM365WorkplaceClassificationWorkbook -Folder $receiveFolder -DownloadFile { param($destination, $name) Copy-Item -LiteralPath (Join-Path $root $name) -Destination $destination; Get-Item -LiteralPath $destination }.GetNewClosure()
    Assert-Classification ((Test-Path -LiteralPath $received.PersonaPath) -and (Test-Path -LiteralPath $received.SitePath)) 'Both downloaded workbooks are validated and returned.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

# Private workbooks: their own TestCases worksheets.
if ($PersonaClassificationPath -or $SiteClassificationPath) {
    if (-not $PersonaClassificationPath -or -not $SiteClassificationPath) { throw 'Give both PersonaClassificationPath and SiteClassificationPath.' }
    $site = Read-SmartM365WorkplaceSiteClassification -Path $SiteClassificationPath
    $persona = Read-SmartM365WorkplacePersonaClassification -Path $PersonaClassificationPath -HeadquartersPersona $site.HeadquartersPersona
    foreach ($case in @(Import-Excel -Path $PersonaClassificationPath -WorksheetName 'TestCases')) {
        $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle ([string]$case.'Job title') -Description '' -Country ([string]$case.Country) -SiteType 'UNKNOWN'
        Assert-Classification ($result.Persona -eq ([string]$case.'Expected result').Trim()) ("Persona TestCase '{0}' ({1})." -f $case.'Job title', $result.Persona)
    }
    foreach ($case in @(Import-Excel -Path $SiteClassificationPath -WorksheetName 'TestCases')) {
        $mapped = Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode ([string]$case.'Directory Site Code')
        $type = if ($null -ne $mapped) { $mapped.Type } else { 'UNKNOWN' }
        Assert-Classification ($type -eq ([string]$case.'Expected Site Type').Trim()) ("Site TestCase {0} ({1})." -f $case.'Directory Site Code', $type)
        $expectedPersona = ([string]$case.'Expected Persona').Trim()
        if ($expectedPersona) {
            $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle 'Unmapped position' -Description '' -Country '' -SiteType $type
            Assert-Classification ($result.Persona -eq $expectedPersona) ("Site TestCase {0} persona ({1})." -f $case.'Directory Site Code', $result.Persona)
        }
    }
}
else { Write-Output 'Private classification workbooks not given; only the synthetic workbooks were validated.' }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA7CtCATGxFFdFY
# 8SZfRHsUsrLZZ+hB6+WpFOA2LRWtKKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILG+aYi6MGyqa+hrkZ83Gb+x3Due7/m0IQ8qGXHsKSkcMA0GCSqG
# SIb3DQEBAQUABIIBgEpALepEQugeW4A/q4hk+45iN5E+LFoYCBmKcDrhguWvf4ot
# S/gMeVPXuu8lcRwjndMrYht2kVW7rDY/vGbGuw3e0kJm3DYyDhfYQ6RtE40lfd5r
# 04Sn2Ey61wmzrM1SSOxEpA1ZzLu5xgKGr2pJXyyqb7Bea5rg0pxvaADRTXWAu8bj
# V7SPg/3kn4H7sjSdThh6B4kfCjtRICU2g4C3hh/qnhEoNYIeD6nTD2rLsLVxeR1+
# bVR242GNdA3Sp4y9hZ1A92YxqOsZfJ4lahfH1+nkXh6nGhXk7k3DzzAdKjBFVJT0
# 6Y3YjNXvXZWAfHTzgVwbqgRGzfEcot/AYJ7b0nItj4YhBP2OHuPUPnD+BWLknfMe
# LjyNZeCNEPZg0ZTKtMgOrraAIl+O715aWzibGfOoD45hEaxCK9rSf9KqIdiQkJn0
# 7sMZfsP1oxRzPeEXU5RwEs01m7O3GFS4qHia6028C+KykEzGQxoZvbL2rsSdrA6E
# eTIsDj59R9JkWeVJm6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDExNzUw
# MjVaMC8GCSqGSIb3DQEJBDEiBCCwCUd7ENuUhjZ9C1N+Z9nTGANiccreafUoG8OI
# gvl5RjANBgkqhkiG9w0BAQEFAASCAgBa56m6dObJK0usCxA7k7Dp5ERI3njRAnvm
# glunEsYGf4/4tKSruIAS14UL1x9M9s0X+GmZk0qjmJQexucJJISWQDhJF/EphJnl
# 5v+nP28vvMvvGvKSDYM3vQ3tJG6n9fgAQVrEOOv1g9Y1Lv8tQbryMF3tQ2xnNrSz
# ddgkBnSQ1G3IJxXiaZCIWJV8eFIyJKP6mqZthatZkLrkyyK2h7tyiRPB88S4MxET
# LaCGkVZzbezxdNaV3IwEPNKuzPfUm9lqQK1X5TkkIFDX7M03k4u3Ag0ukHL/IiYg
# a0qe3IB+Nw/Jwc6OItw9tcRyFsqjg7yNPlTPdiOZTtLuprahYohIM4ig7a8R/AXi
# FmypEdqkmGOjO9GkdJjjyJoS7KMnpN31A1Scwia62lIbDT06vH30T9+/L11GS7bA
# HJvh9dIMeFTkTITBQ6DegqnzVzHwzk2g4TXlQSwBpxtEf95FWdJ0aZAWhVYP1ywF
# 1+Gi29K39/ww2DiI2sv5UiXJ3RJOoqvfrUnT/TdBmf6xGXSWuBT+IqHg3GQyyVW4
# 63k0xEzj9hitdLmgPfTY65WLizL+WGqkev5V3HQKDQDvuLDzLiKhR7FnVAXRPtwD
# 6fPzwcMpE7g4GZRaO0A3/y5ACFr9YR1isAO5s08e77o3IfVwnwkbv6vZSgqwsnQu
# YDofALuX1w==
# SIG # End signature block
