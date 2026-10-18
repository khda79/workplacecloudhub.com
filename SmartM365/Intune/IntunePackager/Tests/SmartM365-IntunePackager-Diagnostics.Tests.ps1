#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for Tools/SmartM365-IntunePackager-Win32AppDiagnostics.ps1 (offline, fictional data).

.DESCRIPTION
    Functions are extracted from the script through its AST; Graph is replaced by a stub that
    returns fictional Contoso apps. The tests check the read-only contract (GET only), the output
    location guard and that the shared output carries no names, ids, commands, URLs or scripts.

.VERSION
    1.0.0
#>

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
    $script:ScriptPath = Join-Path $script:AppRoot 'Tools\SmartM365-IntunePackager-Win32AppDiagnostics.ps1'
    $tokens = $null; $parseErrors = $null
    $script:Ast = [Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw 'Diagnostic script parse failed.' }
    foreach ($assignment in @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -in @('$script:KeptStringProperties', '$script:Base64Properties') }, $false))) {
        . ([scriptblock]::Create($assignment.Extent.Text))
    }
    foreach ($name in @('Split-IpDiagAppList', 'Get-IpDiagRepositoryRoot', 'Test-IpDiagOutputRoot', 'ConvertTo-IpDiagShape', 'Get-IpDiagValue', 'Get-IpDiagShortType', 'Get-IpDiagAppSummary', 'Get-IpDiagScripts', 'Get-IpDiagPropertyMatrix', 'Write-IpDiagReport', 'Invoke-IpDiagSample')) {
        $definition = $script:Ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        if (-not $definition) { throw "Function not found: $name" }
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    function ConvertTo-Base64 { param([string]$Text) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)) }
    $script:Secrets = @('Contoso Sample App', 'Contoso Field Tools', 'Contoso VPN Client', 'ContosoSampleApp.msi', 'msiexec /i', 'https://www.contoso.example', 'C:\Program Files\Contoso', 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso', 'Get-ItemProperty', 'Test-Path', 'aaaaaaaa-1111-4222-8333-000000000001')
    $script:Fixtures = @{
        'aaaaaaaa-1111-4222-8333-000000000001' = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'aaaaaaaa-1111-4222-8333-000000000001'; displayName = 'Contoso Sample App'; description = 'Fictional'
            informationUrl = 'https://www.contoso.example/app'; installCommandLine = 'msiexec /i "ContosoSampleApp.msi" /qn'; setupFilePath = 'ContosoSampleApp.msi'
            installExperience = [pscustomobject]@{ runAsAccount = 'system'; deviceRestartBehavior = 'suppress'; maxRunTimeInMinutes = 60 }
            allowedArchitectures = 'x64,arm64'; applicableArchitectures = 'none'; minimumSupportedWindowsRelease = 'Windows11_23H2'; committedContentVersion = '2'
            detectionRules = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppProductCodeDetection'; productCode = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'; productVersionOperator = 'greaterThanOrEqual'; productVersion = '1.0.0' })
            requirementRules = @(); rules = @(); msiInformation = [pscustomobject]@{ productCode = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'; packageType = 'perMachine' }
            largeIcon = $null; lastModifiedDateTime = [datetime]::new(2026, 10, 18, 9, 31, 0, [DateTimeKind]::Utc)
        }
        'aaaaaaaa-1111-4222-8333-000000000002' = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'aaaaaaaa-1111-4222-8333-000000000002'; displayName = 'Contoso Field Tools'
            detectionRules = @(
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppFileSystemDetection'; path = 'C:\Program Files\Contoso'; fileOrFolderName = 'Tools.exe'; detectionType = 'exists'; operator = 'notConfigured' }
                [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppRegistryDetection'; keyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso'; valueName = 'Version'; detectionType = 'version'; operator = 'greaterThanOrEqual'; detectionValue = '5.0' }
            )
            requirementRules = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRequirement'; scriptContent = (ConvertTo-Base64 'Test-Path C:\Contoso'); detectionType = 'boolean'; operator = 'equal'; detectionValue = 'true' })
            minimumFreeDiskSpaceInMB = 2048; rules = @(); allowedArchitectures = $null; applicableArchitectures = 'x64'
        }
        'aaaaaaaa-1111-4222-8333-000000000003' = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'aaaaaaaa-1111-4222-8333-000000000003'; displayName = 'Contoso VPN Client'
            detectionRules = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptDetection'; scriptContent = (ConvertTo-Base64 "Get-ItemProperty 'HKLM:\SOFTWARE\Contoso'"); enforceSignatureCheck = $false; runAs32Bit = $false })
            rules = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRule'; ruleType = 'detection'; scriptContent = '' })
            requirementRules = @(); largeIcon = [pscustomobject]@{ type = 'image/png'; value = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==' }
        }
    }

    function Invoke-IpDiagGet {
        param([string]$Uri, [string]$Step)
        $script:Requests.Add($Uri)
        $id = ([regex]::Match($Uri, 'mobileApps/(?<Id>[0-9a-f-]{36})')).Groups['Id'].Value
        if ($Uri -match 'contentVersions$') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = '1' }, [pscustomobject]@{ id = '2' }) } }
        if ($Uri -match 'select=largeIcon') { return [pscustomobject]@{ largeIcon = (Get-IpDiagValue $script:Fixtures[$id] 'largeIcon') } }
        return $script:Fixtures[$id]
    }
}

AfterAll { Set-IpLogSink -Sink $null -Console $false }

Describe 'Read-only contract' {
    It 'sends Graph requests only through Invoke-IpDiagGet with GET' {
        $calls = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -in @('Invoke-IpGraphRequest', 'Invoke-MgGraphRequest', 'Invoke-RestMethod', 'Invoke-WebRequest') }, $true))
        $calls.Count | Should -Be 1
        $calls[0].Extent.Text | Should -Match '-Method GET'
        # Every function of the Graph module that writes to Intune or Azure Storage.
        $writers = @('New-IpWin32App', 'New-IpContentVersion', 'New-IpContentFile', 'Request-IpSasRenewal', 'Send-IpIntuneWinContent', 'Send-IpAzureBlobWithRetry', 'Submit-IpContentFileCommit', 'Complete-IpAppContentVersion', 'Publish-IpWin32App')
        $mutating = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -in $writers }, $true))
        @($mutating | ForEach-Object { $_.GetCommandName() }) | Should -BeNullOrEmpty
    }

    It 'requests the read scope by default' {
        $script:Ast.Extent.Text | Should -Match "\`$script:ReadScope = 'DeviceManagementApps\.Read\.All'"
    }
}

Describe 'Helpers' {
    It 'splits the app list on semicolons' {
        Split-IpDiagAppList -Text ' Contoso A ; ;Contoso B;Contoso A' | Should -Be @('Contoso A', 'Contoso B')
        @(Split-IpDiagAppList -Text '').Count | Should -Be 0
    }

    It 'refuses an output folder inside the application folder or a repository' {
        Test-IpDiagOutputRoot -Path (Join-Path $script:AppRoot 'Diagnostics') -AppRoot $script:AppRoot | Should -Match 'outside the application folder'
        Test-IpDiagOutputRoot -Path (Join-Path $TestDrive 'Diagnostics') -AppRoot $script:AppRoot | Should -Be ''
    }

    It 'anonymizes strings, ids and base64 but keeps enums, flags and numbers' {
        $shape = ConvertTo-IpDiagShape -Value $script:Fixtures['aaaaaaaa-1111-4222-8333-000000000001']
        $shape.displayName | Should -Be '<string:18>'
        $shape.id | Should -Be '<guid>'
        $shape.installExperience.runAsAccount | Should -Be 'system'
        $shape.installExperience.maxRunTimeInMinutes | Should -Be 60
        $shape.allowedArchitectures | Should -Be 'x64,arm64'
        $shape.detectionRules[0].'@odata.type' | Should -Be '#microsoft.graph.win32LobAppProductCodeDetection'
        $shape.detectionRules[0].productCode | Should -Be '<guid>'
        $shape.lastModifiedDateTime | Should -Be '<datetime>'
        (ConvertTo-IpDiagShape -Value $script:Fixtures['aaaaaaaa-1111-4222-8333-000000000003']).largeIcon.value | Should -Match '^<base64:\d+>$'
    }

    It 'summarizes detection, rule and requirement types' {
        $summary = Get-IpDiagAppSummary -App $script:Fixtures['aaaaaaaa-1111-4222-8333-000000000002']
        $summary.DetectionRules | Should -Be 'FileSystemDetection,RegistryDetection'
        $summary.RequirementRules | Should -Be 'PowerShellScriptRequirement'
        $summary.ExtraRequirements | Should -Be 'minimumFreeDiskSpaceInMB'
        $summary.ScriptContentPresent | Should -Be '1/1'
        (Get-IpDiagAppSummary -App $script:Fixtures['aaaaaaaa-1111-4222-8333-000000000003']).ScriptContentPresent | Should -Be '1/2'
    }

    It 'decodes embedded scripts' {
        $scripts = @(Get-IpDiagScripts -App $script:Fixtures['aaaaaaaa-1111-4222-8333-000000000003'])
        $scripts.Count | Should -Be 1
        $scripts[0].Kind | Should -Be 'detection'
        $scripts[0].Text | Should -Match 'Get-ItemProperty'
    }
}

Describe 'Sample export' {
    BeforeAll {
        $script:Requests = [Collections.Generic.List[string]]::new()
        $script:Run = Join-Path $TestDrive 'run'
        [IO.Directory]::CreateDirectory($script:Run) | Out-Null
        $sample = @($script:Fixtures.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{ id = $_; displayName = '' } })
        $script:Result = Invoke-IpDiagSample -Sample $sample -RunFolder $script:Run
    }

    It 'exports every sampled app' {
        $script:Result.Exported | Should -Be 3
        @(Get-ChildItem (Join-Path $script:Run 'shared') -Filter '*.shape.json').Count | Should -Be 3
        Test-Path (Join-Path $script:Run 'shared\report.md') | Should -BeTrue
        Test-Path (Join-Path $script:Run 'mapping.local.csv') | Should -BeTrue
    }

    It 'keeps scripts and icons in the local raw folder' {
        @(Get-ChildItem (Join-Path $script:Run 'raw') -Filter '*.ps1').Count | Should -Be 2
        @(Get-ChildItem (Join-Path $script:Run 'raw') -Filter '*-icon.png').Count | Should -Be 1
        (Get-Content (Join-Path $script:Run 'raw\App03.json') -Raw) | Should -Match '<saved as icon file>'
    }

    It 'writes no name, id, command, URL or script content to the shared folder' {
        $shared = (Get-ChildItem (Join-Path $script:Run 'shared') -File | ForEach-Object { Get-Content $_.FullName -Raw }) -join "`n"
        foreach ($secret in $script:Secrets) { $shared | Should -Not -Match ([regex]::Escape($secret)) }
        $shared | Should -Match 'ProductCodeDetection'
        $shared | Should -Match 'PowerShellScriptRequirement'
    }

    It 'reports content versions and icon availability' {
        $report = Get-Content (Join-Path $script:Run 'shared\report.md') -Raw
        $report | Should -Match '\| App03 \|'
        $report | Should -Match '\| installExperience \| 1 \| 1 \|'
    }

    It 'only read apps, icons and content versions' {
        @($script:Requests | Where-Object { $_ -notmatch '^deviceAppManagement/mobileApps/[0-9a-f-]{36}(\?\$select=largeIcon|/microsoft\.graph\.win32LobApp/contentVersions)?$' }).Count | Should -Be 0
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1sj68b3WxXfSS
# ugrmU3/ahg00nlGZ6bYBCv1RFR1Nz6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILO9FBb03ZwVFguAfPZSEA2wX9Ofup7Ghcl7CYbcbyqdMA0GCSqG
# SIb3DQEBAQUABIIBgLDR7mkA/psNzcWWeU96+lH26z3CJBj87ekeN3hKobogQ6ek
# BPMB95w6xT+zw5hjjcAKqKuLeRUDXh9b9JQVbIbVeTqfeJCFalzOW3p21SXFCAez
# V4Yby9yhXUMP75KYWVSI88T3V5HsbQJ+sVP8UJX/XFBeRqxWFAg/3ahRBDNuT5/q
# V1/tJGvTAYptUv2muBuSQSlStjqVQkKCxnP8GnAe02T7+YjgeeuQQlbPiTYuQUub
# cxCalwciMUlBKnHb/WW/gPgN7y4Lz+jrkChKtI7E/sZKuvwszRjVaQZJEqVA9DIZ
# DNlZ7RlIXoDd/dY6ETu+X7k5HW3gfSHFWX+nU36Zpm/hn196r5UL7j1ArXVU7qTL
# X4iJCOIgUzCdXeF1ziwSGm1J2gZg7+4r/oWO2Jn4g5ke34gxtVCtJ78tKZ9qRCNC
# 0wDFqAdiuPbJiLzzb7d13JZsUBxmENum7ONG5r5O17bwgjbRUUs/IiffInAfmcq3
# rAGd4Rn/k6tkcahu4aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzZaMC8GCSqGSIb3DQEJBDEiBCBgpVwiYvNtfw60RIjKuLkAjU3qcVg8HjOhv7ir
# mYh3hzANBgkqhkiG9w0BAQEFAASCAgCKvNrsx6JL2wp9EEyVPV9W/IRLqJSWmHsG
# zaSkAgj62Aq2XppFwXr4z5AvZnGSIf9mWnIFc9Pnj4MfMVXQC8Wr6uGCZBM3rtZF
# YaRuiFd1T3dXJW+J34JPGaF3m4VkD2dz8HC1OfxGQCBdnUIwUk8VqMGxQQSQ+Obj
# 9hEb6WMjl01nVqBac3a1G3GZTq2Ww8CiErL8tCxy2JBTdSfO8qyvjoPlTxoSP/7f
# qJrFMo1wmgkJhMsIt8eA5ZwqALpgVtEAZzoGr+ryq6kQM/V9IvcYInUH9wPqC2ky
# CSDCjD1MhFtl0BL0IVYVuWmjHsEHQECftuBraCuohvsHW/Cz+zPjKA5ZFfOhgY1g
# YXLYTf7W+tbccGSh3HpStculxk3GlekIruKDmvtuumy4SBs54UhpnCMUumP8tu/i
# rL5b397ZnEe0jA30dkDCtA5IH/2Lin+2iXCZjVCGPgSQc3xwONABGPJ5kq5vFfNO
# GxCRbGkedNSOJcy9821AIqWtkiPFG5wDDpLnjk75Jy4flSArIAhWZ3fle8DwgI3u
# HJqcNjy3FCCfOuNyfhSNJvO0mwlSntD44SngRyF8msIHcu6mi59w25LPi3pzVx4i
# 0Pgf+KmqT8UsEMp2DscRAJnYTHKlR7OSv1t3ZiuoZ6mn9gs4V9DllxivaacJ429T
# 3nhjKJLxFg==
# SIG # End signature block
