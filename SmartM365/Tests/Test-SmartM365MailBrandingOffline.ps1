#Requires -Version 5.1
<#
.SYNOPSIS
Offline mail branding checks for both Core editions; no tenant loading or mail transport.
.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$PreviewPath,
    [string]$PreviewClientLogoPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$sourceRoot = Split-Path $PSScriptRoot -Parent
$script:Checks = 0
function Assert-Branding {
    param([bool]$Condition, [string]$Message)
    $script:Checks++
    if (-not $Condition) { throw $Message }
}

$assetPath = Join-Path $sourceRoot 'Config/Branding/Assets/SmartPowerBI-logo.png'
$assetData = 'data:image/png;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes($assetPath))
$editions = @(
    'Modules/SmartM365.Core/SmartM365.Core.psm1',
    'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'
)
$names = @(
    'Get-ModuleLocalConfigValue', 'ConvertTo-SmartM365ConfigBoolean',
    'Get-SmartM365MailBrandingConfig', 'ConvertTo-SmartM365MailLogoDataUri',
    'ConvertTo-SmartM365EmailHtmlText', 'Add-SmartM365MailBranding', 'New-SmartM365EmailBody'
)
$previousHtml = $null
foreach ($edition in $editions) {
    $path = Join-Path $sourceRoot $edition
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    Assert-Branding ($errors.Count -eq 0) "Parse failed: $edition"
    $definitions = foreach ($name in $names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) { throw "Missing function: $name" }
        # AST extraction loses the defining file's automatic PSScriptRoot binding.
        $node.Extent.Text.Replace('$PSScriptRoot', '$script:BrandingModuleRoot')
    }
    # Execute only the named production functions, with synthetic config and logging.
    $runtime = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
    try {
        & $runtime {
            param($ModuleRoot)
            $script:BrandingModuleRoot = $ModuleRoot
            $script:GlobalConfig = [pscustomobject]@{}
            $script:CallerConfig = [pscustomobject]@{}
            $script:Warnings = @()
            function script:Get-ModuleLocalConfig { [pscustomobject]@{} }
            function script:Get-SmartM365CallerLocalConfig { $script:CallerConfig }
            function script:Get-SmartM365EffectiveModuleGlobalConfig { $script:GlobalConfig }
            function script:Resolve-SmartM365ConfigValue { param($Value) $Value }
            function script:WriteLog {
                param($Message, $Level)
                if ($Level -eq 'WARNING') { $script:Warnings += $Message }
            }
            function script:Add-SmartM365MailExecutionFooter { param($BodyHtml) $BodyHtml }
        } (Split-Path $path -Parent)

        $defaults = & $runtime { Get-SmartM365MailBrandingConfig }
        Assert-Branding (-not $defaults.PowerBiEnabled) 'Power BI must default to disabled.'
        Assert-Branding ($defaults.PowerBiUrl -eq '') 'The default report URL must be empty.'
        Assert-Branding ($defaults.PowerBiLogoPath -eq 'Config/Branding/Assets/SmartPowerBI-logo.png') 'The shipped logo must use a portable path.'
        Assert-Branding ($defaults.PowerBiLinkText -eq 'Click here to open the Power BI dashboard') 'Default link text changed.'

        & $runtime {
            param($Logo)
            $script:GlobalConfig = [pscustomobject]@{
                MailBrandingEnabled = $true
                MailClientName = 'DEMO & Client'
                MailClientLogoPath = $Logo
                MailClientLogoMaxKB = 200
                MailPowerBiEnabled = $true
                MailPowerBiUrl = 'https://example.invalid/dashboard?a=1&b="quoted"&literal=$1'
                MailPowerBiLinkText = 'Open <dashboard> & "details" $1'
            }
            $script:CallerConfig = [pscustomobject]@{
                MailPowerBiEnabled = '__USE_GLOBAL__'
                MailPowerBiLogoPath = '__USE_GLOBAL__'
                MailPowerBiLinkText = '__USE_GLOBAL__'
            }
        } $assetPath
        $body = '<html><body><table role="presentation"><tr><td>Body remains unchanged</td></tr></table></body></html>'
        # A caller can run from any directory, including outside the checkout.
        Push-Location ([IO.Path]::GetTempPath())
        try { $html = & $runtime { param($Body) Add-SmartM365MailBranding $Body } $body }
        finally { Pop-Location }
        $safeUrl = [Net.WebUtility]::HtmlEncode('https://example.invalid/dashboard?a=1&b="quoted"&literal=$1')
        Assert-Branding ($html.Contains('href="' + $safeUrl + '"')) 'The escaped report URL is missing or was interpreted as a regex replacement.'
        Assert-Branding (([regex]::Matches($html, [regex]::Escape('href="' + $safeUrl + '"'))).Count -eq 2) 'Image and text must link to the same URL.'
        Assert-Branding ($html.Contains('Open &lt;dashboard&gt; &amp; &quot;details&quot; $1')) 'Link text must be escaped and preserve dollar characters.'
        Assert-Branding ($html.Contains('<td align="left"') -and $html.Contains('<td align="right"')) 'Both aligned cells are required.'
        Assert-Branding ($html -match 'align="left"[^>]*><img[^>]*width="92"') 'Client image must remain 92px and left aligned.'
        Assert-Branding ($html.Contains('alt="Power BI" height="48"')) 'Power BI image height or alt text is missing.'
        Assert-Branding ($html.Contains($assetData)) 'The shipped PNG was not embedded from the module-relative location.'
        Assert-Branding (([regex]::Matches($html, 'target="_blank" rel="noopener"')).Count -eq 2) 'Both links must have safe target attributes.'
        Assert-Branding ($html.Contains('color:#0f6cbd;text-decoration:underline;')) 'Link appearance changed.'
        Assert-Branding ($html.Contains('Body remains unchanged')) 'The original body was lost.'
        Assert-Branding ((& $runtime { $script:Warnings.Count }) -eq 0) 'Valid branding should not warn.'
        if ($null -ne $previousHtml) { Assert-Branding ($html -ceq $previousHtml) 'PS7 and PS5.1 branding HTML differ.' }
        $previousHtml = $html
        $again = & $runtime { param($Body) Add-SmartM365MailBranding $Body } $html
        Assert-Branding ($again -ceq $html) 'Already branded mail must not gain another banner.'
        Assert-Branding (([regex]::Matches($again, '<!-- SmartM365MailBranding:v1 -->')).Count -eq 1) 'The branding marker was duplicated.'

        & $runtime { $script:CallerConfig = [pscustomobject]@{ MailPowerBiUrl = 'https://example.invalid/tenant-dashboard'; MailPowerBiLinkText = 'Tenant dashboard' } }
        $override = & $runtime { Get-SmartM365MailBrandingConfig }
        Assert-Branding ($override.PowerBiUrl -eq 'https://example.invalid/tenant-dashboard' -and $override.PowerBiLinkText -eq 'Tenant dashboard') 'Explicit overrides must win.'
        & $runtime { $script:CallerConfig = [pscustomobject]@{}; $script:GlobalConfig.MailPowerBiEnabled = 'false' }
        $disabled = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding (-not $disabled.Contains('alt="Power BI"') -and -not $disabled.Contains('example.invalid')) 'Disabled Power BI must emit no report link or image.'
        Assert-Branding ($disabled -match 'align="left"[^>]*><img') 'Disabling Power BI must keep the client logo left aligned.'
        & $runtime { $script:GlobalConfig.MailPowerBiEnabled = 'true' }
        foreach ($invalidUrl in @('http://example.invalid/report', 'javascript:alert(1)', 'https://', '//example.invalid/report', "https://example.invalid/`r`ninvalid")) {
            & $runtime { param($Url) $script:GlobalConfig.MailPowerBiUrl = $Url; $script:Warnings = @() } $invalidUrl
            $invalidHtml = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
            Assert-Branding ($invalidHtml -ceq $disabled) 'Invalid URL must degrade to the client-only banner.'
            Assert-Branding ((& $runtime { $script:Warnings.Count }) -eq 1) 'Invalid URL must log one WARNING.'
        }
        & $runtime { $script:GlobalConfig.MailPowerBiUrl = ''; $script:Warnings = @() }
        $emptyHtml = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding ($emptyHtml -ceq $disabled) 'Empty URL must degrade to the client-only banner.'
        Assert-Branding ((& $runtime { $script:Warnings.Count }) -eq 0) 'An unconfigured URL must not warn.'

        & $runtime {
            $script:GlobalConfig.MailPowerBiUrl = 'https://example.invalid/dashboard'
            $script:CallerConfig = [pscustomobject]@{ MailPowerBiLogoPath = 'Config/Branding/Local/nonexistent-test-logo.png' }
            $script:Warnings = @()
        }
        $missing = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding (-not $missing.Contains('alt="Power BI"') -and $missing.Contains('href="https://example.invalid/dashboard"')) 'Missing logo must keep the text link.'
        Assert-Branding ((& $runtime { $script:Warnings.Count }) -eq 1) 'Missing logo must warn.'
        & $runtime { $script:CallerConfig = [pscustomobject]@{ MailClientLogoMaxKB = 1 }; $script:Warnings = @() }
        $oversize = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding (-not $oversize.Contains('<img') -and $oversize.Contains('href="https://example.invalid/dashboard"')) 'Both images must obey the shared size limit.'
        Assert-Branding ($oversize.Contains('DEMO &amp; Client')) 'Missing client logo must fall back to the escaped client name.'
        & $runtime { $script:CallerConfig = [pscustomobject]@{ MailClientLogoPath = 'missing-client.png' } }
        $label = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding ($label -match 'align="left"[^>]*><span[^>]*>DEMO &amp; Client</span>') 'Client name must be left aligned.'
        & $runtime { $script:GlobalConfig.MailClientName = ''; $script:GlobalConfig.MailClientLogoPath = ''; $script:CallerConfig = [pscustomobject]@{} }
        $powerBiOnly = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding ($powerBiOnly.Contains('alt="Power BI"') -and $powerBiOnly -match 'align="left"[^>]*></td>') 'Power BI must render even without client branding.'
        & $runtime { $script:GlobalConfig.MailBrandingEnabled = $false; $script:Warnings = @() }
        $noBranding = & $runtime { Add-SmartM365MailBranding '<p>Body</p>' }
        Assert-Branding (-not $noBranding.Contains('<img') -and -not $noBranding.Contains('example.invalid')) 'Branding disabled must suppress both header columns.'
        Assert-Branding ($noBranding.Contains('workplacecloudhub.com') -and $noBranding.Contains('SmartM365MailBranding:v1')) 'Existing footer/marker behavior must remain when branding is disabled.'
        Assert-Branding ((& $runtime { $script:Warnings.Count }) -eq 0) 'Disabled branding must not read or warn about images.'

        if ($PreviewPath -and $edition -eq $editions[0]) {
            $preview = & $runtime {
                param($ClientLogo)
                $script:GlobalConfig = [pscustomobject]@{
                    MailBrandingEnabled = $true; MailClientName = 'DEMO Client'; MailClientLogoPath = $ClientLogo
                    MailPowerBiEnabled = $true; MailPowerBiUrl = 'https://example.invalid/dashboard'
                }
                $script:CallerConfig = [pscustomobject]@{}
                $body = New-SmartM365EmailBody -Title 'Quarantaine EXO' -Category 'Exchange Online - DEMO' -Severity Warning `
                    -HostName 'DEMO-HOST' -GeneratedAt '2026-10-20 12:00:00 +02:00' -StatusBadge 'DEMO' `
                    -Message 'Apercu local de demonstration. Donnees fictives ; aucun message envoye.' `
                    -SummaryRows @(@{ Label = 'Messages en quarantaine'; Value = '12' }, @{ Label = 'Utilisateurs concernes'; Value = '8' })
                Add-SmartM365MailBranding $body
            } $PreviewClientLogoPath
            [IO.File]::WriteAllText([IO.Path]::GetFullPath($PreviewPath), $preview, [Text.UTF8Encoding]::new($false))
        }
    }
    finally { Remove-Module $runtime -ErrorAction SilentlyContinue }
}
Write-Output ("PASS: {0} mail branding assertions; both module editions." -f $script:Checks)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBiDaSYXQ0s+WED
# n0DwYWP40QW/zu0nTFwCGbO669nfqKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDb1ujF+JH3iCQN+l7Jb7toUcOauWnehFu/NcGK+E2eDMA0GCSqG
# SIb3DQEBAQUABIIBgEkmireE3NtOaSGkEVMX8aVzQ8dpB4yeHH+29YIPlXSPB4Ip
# NeYO7DwxN0hRFWFEmyzcQrExvtQ4r3GEWEfNn4bs2TAGHcjWJdqKXz3a2JZQ8dNP
# oFjcQUZ6qHvgBWoRBRS670tq5W0iUP4Lej7JtTqpkkBiyv2QJ4i73O0xmADRveGk
# iebECGLCktxsHYDP+4DwDqA3apJAPCGO8++z54wzvEb4Q4iIaPjnljAkEyKOjKsc
# NL2hPyCnfDIT7pPtRxocaj3DCxrjhvukQJCorz8UWDxOMLL9hmYyKlQ2+KkFkSc5
# eMoHvpkZpiYpl8x+Wb+HPaf7H0htPi9RjeU93jOQNo1I0epSPoGbAblSHiFhwVGQ
# m9FNttFu/pEi9ysnLWKZYlQOn4S59egX6UYL2ry1+6Z4av1wucNot3WOPNaPHjEI
# JLLHnr7lUQ4b8Xt+kCuMjDZyeZaCAdIDaXJC4WKGl3bfQZeIq4HnX2YCF02xjIHr
# dlFjbIr1s4B0ptNeQ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAxMzM4
# MTBaMC8GCSqGSIb3DQEJBDEiBCAGNOh/DYz6NtLPPAm7CNkpHhpDnaXpKMu6BSKv
# WCOabzANBgkqhkiG9w0BAQEFAASCAgB6XPU1O3YFT1MyXloQXaBdFPQgssqiovMf
# 47fYOItDUSMpM1yL5BVcoFsb1MxR5rCT272ULKjkYsLo+V2GfnSERFU8ljd7rzOU
# E1svoCmwUgw+cXzLVTWLwA2lM63zKalkSj9oBXjHyz6cwXUsEOAaT4fYDXoZKF+a
# l9joesfosBT29rMgMlPdNLGe2fHFrpGFgWm4nhcOKfdZKPb/OflmIW+z44y5jKHV
# 8tLhLeP5ROGJuExP4AZ9eRwclMTUYQ86MnrNd+WeKHL3LlgvJHalXAbtNIWqihTa
# lu44TZdKk7snPm5EpHCFPdzrLgYkOVqKbM049rf+/2PRQzpuN+SW5H3e2jVv/mAz
# 5rf1ywPjZbO0TwkEyAwvCz5tolBUUF//UFDPbm3grlUoHpiJY3EMksftB0JflD8y
# AMNP9ONahdcrf7KZI3RTg2mGrFATM7p2mK9W+NL6QVfy6lIn8yDuYAx+ctUju6IH
# jb+Q91RzTtS4P8notPbhICnBgwLAhPF9h2bew3jC+oChMX/cTaueWriSVr9GcnmU
# 38W4NB2XrV4TqHoY9kcVCIlzu0T7g33gZlpQicNhJ2uUsxUpLBS6xvmQhErVgnQS
# /hTPmcA2z5BRrVVVObdf96XdCHHVd+ZN0F0alyEaZxuCOO3Sp3tv4Ln/coAyuHdZ
# Hibcq08EMg==
# SIG # End signature block
