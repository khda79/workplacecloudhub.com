#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for SmartM365.IntunePackager.Core.psm1 (offline, synthetic data).

.VERSION
    1.0.3
#>

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    . (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-TestHelpers.ps1')
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
}

Describe 'Logging' {
    It 'prefixes every physical line with yyyy-MM-dd HH:mm:ss' {
        $lines = Format-IpLogLines -Message "first`r`nsecond`nthird" -Timestamp ([datetime]'2026-01-02 03:04:05')
        $lines.Count | Should -Be 3
        $lines | ForEach-Object { $_ | Should -Match '^2026-01-02 03:04:05 ' }
        $lines[2] | Should -Be '2026-01-02 03:04:05 third'
    }

    It 'tags warning and error levels after the timestamp' {
        (Format-IpLogLines -Message 'x' -Level WARN -Timestamp ([datetime]'2026-01-02 03:04:05')) | Should -Be '2026-01-02 03:04:05 [WARN] x'
    }

    It 'writes timestamped lines to the log file and to the sink' {
        $received = [Collections.Generic.List[string]]::new()
        $logPath = Join-Path $TestDrive 'log.txt'
        Set-IpLogPath -Path $logPath
        Set-IpLogSink -Sink { param($line, $level) $received.Add($line) }.GetNewClosure() -Console $false
        try { Write-IpLog "alpha`nbeta" }
        finally { Set-IpLogSink -Sink $null -Console $false; Set-IpLogPath -Path '' }
        $fileLines = Get-Content -LiteralPath $logPath
        $fileLines.Count | Should -Be 2
        $fileLines | ForEach-Object { $_ | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ' }
        $received.Count | Should -Be 2
    }

    It 'keeps file-only details out of the sink' {
        $received = [Collections.Generic.List[string]]::new()
        $logPath = Join-Path $TestDrive 'fileonly.txt'
        Set-IpLogPath -Path $logPath
        Set-IpLogSink -Sink { param($line, $level) $received.Add($line) }.GetNewClosure() -Console $false
        try { Write-IpLog 'visible'; Write-IpLog 'Error location: line 42' -Level ERROR -FileOnly }
        finally { Set-IpLogSink -Sink $null -Console $false; Set-IpLogPath -Path '' }
        $received.Count | Should -Be 1
        (Get-Content -LiteralPath $logPath -Raw) | Should -Match 'Error location: line 42'
    }

    It 'formats byte sizes with an adapted unit' -ForEach @(
        @{ Bytes = 512; Text = '512 B' }, @{ Bytes = 2144; Text = '2.1 KB' }, @{ Bytes = 157286400; Text = '150.0 MB' }, @{ Bytes = 5GB; Text = '5.0 GB' }
    ) {
        Format-IpByteSize -Bytes $Bytes | Should -Be $Text
    }

    It 'converts Graph timestamps to local time' {
        $utc = [datetime]::new(2026, 10, 18, 9, 31, 0, [DateTimeKind]::Utc)
        $expected = $utc.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
        Format-IpLocalDateTime '2026-10-18T09:31:00Z' | Should -Be $expected
        Format-IpLocalDateTime $utc | Should -Be $expected
        Format-IpLocalDateTime ([datetime]::SpecifyKind($utc, [DateTimeKind]::Unspecified)) | Should -Be $expected
        Format-IpLocalDateTime ([datetimeoffset]::new($utc)) | Should -Be $expected
        Format-IpLocalDateTime $null | Should -Be ''
        Format-IpLocalDateTime 'not a date' | Should -Be 'not a date'
    }

    It 'archives the previous log and keeps the newest archives' {
        $directory = Join-Path $TestDrive 'Logs'
        New-Item -ItemType Directory -Path $directory | Out-Null
        1..4 | ForEach-Object { Set-Content -LiteralPath (Join-Path $directory ("App-2026010{0}-000000.log" -f $_)) -Value 'old' }
        Set-Content -LiteralPath (Join-Path $directory 'App.log') -Value 'current'
        $path = Initialize-IpLog -Directory $directory -BaseName 'App' -MaxFiles 3
        try {
            $path | Should -Be (Join-Path $directory 'App.log')
            Test-Path -LiteralPath $path | Should -BeFalse
            @(Get-ChildItem -LiteralPath $directory -Filter 'App-*.log').Count | Should -Be 3
        }
        finally { Set-IpLogPath -Path '' }
    }
}

Describe 'Configuration merge' {
    It 'adds missing nested keys without overwriting local values' {
        $template = '{"A":1,"Section":{"Keep":"template","New":"added"},"List":[1,2]}' | ConvertFrom-Json
        $local = '{"A":5,"Section":{"Keep":"local"}}' | ConvertFrom-Json
        $added = Merge-IpObjectFromTemplate -Target $local -Template $template
        $added | Should -Be @('Section.New', 'List')
        $local.A | Should -Be 5
        $local.Section.Keep | Should -Be 'local'
        $local.Section.New | Should -Be 'added'
        @($local.List).Count | Should -Be 2
    }

    It 'merges every tenant profile with the template tenant schema' {
        $template = '{"Tenants":{"test":{"TenantId":"","AuthMode":"Delegated","ClientId":""}}}' | ConvertFrom-Json
        $local = '{"Tenants":{"prod":{"TenantId":"b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e"}}}' | ConvertFrom-Json
        $added = Merge-IpObjectFromTemplate -Target $local -Template $template
        $added | Should -Be @('Tenants.prod.AuthMode', 'Tenants.prod.ClientId')
        $local.Tenants.PSObject.Properties.Name | Should -Be @('prod')
        $local.Tenants.prod.TenantId | Should -Be 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e'
    }

    It 'creates the local configuration from the template when missing' {
        $templatePath = Join-Path $TestDrive 'app.config.json.template'
        $configPath = Join-Path $TestDrive 'app.config.json'
        '{"Value":"template"}' | Set-Content -LiteralPath $templatePath
        $config = Import-IpConfiguration -ConfigPath $configPath -TemplatePath $templatePath
        Test-Path -LiteralPath $configPath | Should -BeTrue
        $config.Value | Should -Be 'template'
    }

    It 'saves added keys and keeps local values on load' {
        $templatePath = Join-Path $TestDrive 'merge.config.json.template'
        $configPath = Join-Path $TestDrive 'merge.config.json'
        '{"Keep":"template","Upload":{"BlockSizeMB":16,"MaxRetries":5}}' | Set-Content -LiteralPath $templatePath
        '{"Keep":"local","Upload":{"BlockSizeMB":32}}' | Set-Content -LiteralPath $configPath
        $config = Import-IpConfiguration -ConfigPath $configPath -TemplatePath $templatePath
        $config.Keep | Should -Be 'local'
        $config.Upload.BlockSizeMB | Should -Be 32
        $saved = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        $saved.Upload.MaxRetries | Should -Be 5
        $saved.Upload.BlockSizeMB | Should -Be 32
    }

    It 'ships a template without tenant identifiers or secrets' {
        $template = Read-IpJsonFile -Path (Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.config.json.template')
        foreach ($tenant in $template.Tenants.PSObject.Properties) {
            $tenant.Value.TenantId | Should -BeNullOrEmpty
            $tenant.Value.ClientId | Should -BeNullOrEmpty
            $tenant.Value.CertificateThumbprint | Should -BeNullOrEmpty
        }
        $template.Tenants.PSObject.Properties.Name | Should -Contain 'test'
        $template.ContentPrepTool.Sha256 | Should -Match '^[0-9A-F]{64}$'
        $template.Upload.BlockSizeMB | Should -BeGreaterThan 0
    }

    It 'resolves AppRoot tokens, environment variables and relative paths' {
        $base = Join-Path $TestDrive 'root'
        Resolve-IpPath -Path '{{AppRoot}}\Logs' -BaseDirectory $base | Should -Be ([IO.Path]::GetFullPath((Join-Path $base 'Logs')))
        Resolve-IpPath -Path 'Output' -BaseDirectory $base | Should -Be ([IO.Path]::GetFullPath((Join-Path $base 'Output')))
        Resolve-IpPath -Path '%TEMP%\x' -BaseDirectory $base | Should -Be ([IO.Path]::GetFullPath((Join-Path $env:TEMP 'x')))
        Resolve-IpPath -Path '' -BaseDirectory $base | Should -Be ''
    }
}

Describe 'Tenant profiles' {
    BeforeAll {
        $script:TenantConfig = '{"Tenants":{"test":{"TenantId":"","AuthMode":"Delegated","ClientId":"","CertificateThumbprint":""},"cert":{"TenantId":"6e1f3a9c-2b8d-4c7e-a5f0-9d3b7c1e8a24","AuthMode":"Certificate","ClientId":"3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c","CertificateThumbprint":"ab12 cd34 ef56 7890 ab12 cd34 ef56 7890 ab12 cd34"}}}' | ConvertFrom-Json
    }

    It 'returns the delegated test profile' {
        (Get-IpTenantProfile -Config $script:TenantConfig -Tenant 'test').AuthMode | Should -Be 'Delegated'
    }

    It 'normalizes the certificate thumbprint' {
        (Get-IpTenantProfile -Config $script:TenantConfig -Tenant 'cert').CertificateThumbprint | Should -Be 'AB12CD34EF567890AB12CD34EF567890AB12CD34'
    }

    It 'fails clearly for an unknown tenant' {
        { Get-IpTenantProfile -Config $script:TenantConfig -Tenant 'missing' } | Should -Throw "*tenant profile 'missing' was not found*"
    }

    It 'defaults the delegated sign-in mode to Auto' {
        (Get-IpTenantProfile -Config $script:TenantConfig -Tenant 'test').DelegatedSignIn | Should -Be 'Auto'
    }

    It 'rejects an invalid or placeholder TenantId' -ForEach @(
        @{ Value = '<TenantId>' }, @{ Value = 'xxx' }, @{ Value = 'your-tenant-id' }, @{ Value = '00000000-0000-0000-0000-000000000000' },
        @{ Value = '11111111-1111-1111-1111-111111111111' }, @{ Value = 'contoso.onmicrosoft.com' }, @{ Value = 'not a domain' }, @{ Value = 'b4c9e2d1-7a3f-4e8b-9c21' }
    ) {
        $config = ('{"Tenants":{"bad":{"TenantId":"' + $Value + '","AuthMode":"Delegated","ClientId":"","CertificateThumbprint":""}}}') | ConvertFrom-Json
        { Get-IpTenantProfile -Config $config -Tenant 'bad' } | Should -Throw "*TenantId '$Value' is invalid*"
    }

    It 'accepts a tenant GUID or a tenant domain' {
        Test-IpTenantIdentifier -Value 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e' | Should -BeNullOrEmpty
        Test-IpTenantIdentifier -Value 'northwind-it.onmicrosoft.com' | Should -BeNullOrEmpty
        $config = '{"Tenants":{"prod":{"TenantId":"northwind-it.onmicrosoft.com","AuthMode":"Delegated"}}}' | ConvertFrom-Json
        (Get-IpTenantProfile -Config $config -Tenant 'prod').TenantIdKind | Should -Be 'Domain'
    }

    It 'rejects a placeholder ClientId and a malformed thumbprint' {
        $config = '{"Tenants":{"a":{"TenantId":"","AuthMode":"Delegated","ClientId":"00000000-0000-0000-0000-000000000000"},"b":{"TenantId":"b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e","AuthMode":"Certificate","ClientId":"3f2a1b0c-9d8e-4f7a-8b6c-5d4e3f2a1b0c","CertificateThumbprint":"ABCDEF"}}}' | ConvertFrom-Json
        { Get-IpTenantProfile -Config $config -Tenant 'a' } | Should -Throw '*ClientId*placeholder GUID*'
        { Get-IpTenantProfile -Config $config -Tenant 'b' } | Should -Throw '*40 hexadecimal characters*'
    }

    It 'requires a ClientId for browser sign-in' {
        $config = '{"Tenants":{"browser":{"TenantId":"","AuthMode":"Delegated","DelegatedSignIn":"Browser","ClientId":""}}}' | ConvertFrom-Json
        { Get-IpTenantProfile -Config $config -Tenant 'browser' } | Should -Throw '*DelegatedSignIn=Browser*ClientId*'
    }

    It 'requires identifiers for certificate authentication' {
        $config = '{"Tenants":{"bad":{"TenantId":"","AuthMode":"Certificate","ClientId":"","CertificateThumbprint":""}}}' | ConvertFrom-Json
        { Get-IpTenantProfile -Config $config -Tenant 'bad' } | Should -Throw '*requires TenantId*'
    }
}

Describe 'App definition profiles' {
    It 'reports required fields on an empty definition' {
        $errors = Test-IpAppDefinition -Definition (New-IpAppDefinition)
        $errors | Should -Contain 'App.DisplayName is required.'
        $errors | Should -Contain 'Program.InstallCommandLine is required.'
    }

    It 'accepts a complete definition' {
        Test-IpAppDefinition -Definition (New-TestDefinition) | Should -BeNullOrEmpty
    }

    It 'round-trips through a JSON profile' {
        $path = Join-Path $TestDrive 'profile.json'
        Export-IpAppProfile -Definition (New-TestDefinition) -Path $path
        $loaded = Import-IpAppProfile -Path $path
        $loaded.App.DisplayName | Should -Be 'Contoso Sample App'
        $loaded.Requirements.Architectures | Should -Be @('x64', 'arm64')
        @($loaded.ReturnCodes).Count | Should -Be 5
    }

    It 'adds missing keys with defaults when loading an older profile' {
        $path = Join-Path $TestDrive 'old.profile.json'
        '{"SchemaVersion":1,"App":{"DisplayName":"Old"}}' | Set-Content -LiteralPath $path
        $loaded = Import-IpAppProfile -Path $path
        $loaded.App.DisplayName | Should -Be 'Old'
        $loaded.Program.RunAsAccount | Should -Be 'system'
        $loaded.Detection.Msi.ProductVersionOperator | Should -Be 'notConfigured'
    }

    It 'offers the Graph values observed for Windows 10 21H2 and 22H2' {
        $releases = Get-IpWindowsReleases
        $releases | Should -Contain 'Windows10_22H2'
        $releases | Should -Contain 'Windows10_21H2'
        $releases | Should -Contain '2H20'
        $releases | Should -Not -Contain '22H2'
        $releases | Should -Not -Contain '21H2'
    }

    It 'converts a legacy 22H2 value when loading an older profile' {
        $path = Join-Path $TestDrive 'legacy-release.profile.json'
        '{"SchemaVersion":1,"Requirements":{"Architectures":["x64"],"MinimumWindowsRelease":"22H2"}}' | Set-Content -LiteralPath $path
        (Import-IpAppProfile -Path $path).Requirements.MinimumWindowsRelease | Should -Be 'Windows10_22H2'
    }

    It 'rejects a profile with a newer schema' {
        $path = Join-Path $TestDrive 'new.profile.json'
        '{"SchemaVersion":99}' | Set-Content -LiteralPath $path
        { Import-IpAppProfile -Path $path } | Should -Throw '*newer than supported*'
    }

    It 'rejects duplicate return codes and unsupported architectures' {
        $definition = New-TestDefinition
        $definition.ReturnCodes = @([pscustomobject]@{ ReturnCode = 0; Type = 'success' }, [pscustomobject]@{ ReturnCode = 0; Type = 'retry' })
        $definition.Requirements.Architectures = @('arm')
        $errors = Test-IpAppDefinition -Definition $definition
        $errors | Should -Contain 'Return code 0 is defined more than once.'
        $errors | Should -Contain "Unsupported architecture 'arm'. Use x64, arm64 or x86, or NotRestricted."
    }

    It 'requires an existing app id in update mode' {
        $definition = New-TestDefinition
        $definition.Publish.Mode = 'UpdateExisting'
        Test-IpAppDefinition -Definition $definition | Should -Contain 'Publish.ExistingAppId must be a valid GUID when Publish.Mode is UpdateExisting.'
    }
}

Describe 'MSI metadata' {
    It 'maps Property table values to definition fields and commands' {
        $properties = [ordered]@{ ProductCode = '{11111111-2222-3333-4444-555555555555}'; ProductVersion = '2.4.0'; Manufacturer = 'Contoso Ltd'; ProductName = 'Contoso Sample'; UpgradeCode = '{66666666-7777-8888-9999-000000000000}'; ALLUSERS = '1' }
        $metadata = ConvertTo-IpMsiMetadata -Properties $properties -MsiFileName 'ContosoSetup.msi'
        $metadata.Publisher | Should -Be 'Contoso Ltd'
        $metadata.Name | Should -Be 'Contoso Sample'
        $metadata.PackageType | Should -Be 'perMachine'
        $metadata.RunAsAccount | Should -Be 'system'
        $metadata.InstallCommandLine | Should -Be 'msiexec /i "ContosoSetup.msi" /qn /norestart'
        $metadata.UninstallCommandLine | Should -Be 'msiexec /x {11111111-2222-3333-4444-555555555555} /qn /norestart'
    }

    It 'treats a missing ALLUSERS as a per-user package' {
        (ConvertTo-IpMsiMetadata -Properties ([ordered]@{ ProductCode = '{1}' }) -MsiFileName 'a.msi').RunAsAccount | Should -Be 'user'
    }

    It 'reads a generated MSI database through Windows Installer' {
        $path = Join-Path $TestDrive 'ContosoSetup.msi'
        New-TestMsi -Path $path -Properties @{ ProductCode = '{11111111-2222-3333-4444-555555555555}'; ProductVersion = '2.4.0'; Manufacturer = 'Contoso Ltd'; ProductName = "Contoso O'Sample"; ALLUSERS = '1' } | Out-Null
        $metadata = Get-IpMsiMetadata -Path $path
        $metadata.ProductCode | Should -Be '{11111111-2222-3333-4444-555555555555}'
        $metadata.ProductVersion | Should -Be '2.4.0'
        $metadata.Name | Should -Be "Contoso O'Sample"
    }
}

Describe 'PSADT detection' {
    It 'prefers PSADT v4' {
        $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'psadt4')
        New-Item -ItemType File -Path (Join-Path $folder 'Invoke-AppDeployToolkit.exe') | Out-Null
        New-Item -ItemType File -Path (Join-Path $folder 'Deploy-Application.exe') | Out-Null
        $result = Get-IpPsadtCommands -SourceFolder $folder
        $result.ToolkitVersion | Should -Be 4
        $result.InstallCommandLine | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent'
    }

    It 'detects PSADT v3' {
        $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'psadt3')
        New-Item -ItemType File -Path (Join-Path $folder 'Deploy-Application.exe') | Out-Null
        $result = Get-IpPsadtCommands -SourceFolder $folder
        $result.ToolkitVersion | Should -Be 3
        $result.UninstallCommandLine | Should -Be 'Deploy-Application.exe -DeploymentType Uninstall -DeployMode Silent'
    }

    It 'returns nothing for a plain folder' {
        Get-IpPsadtCommands -SourceFolder (New-Item -ItemType Directory -Path (Join-Path $TestDrive 'plain')) | Should -BeNullOrEmpty
    }
}

Describe 'Detection.xml and .intunewin metadata' {
    It 'parses encryption information and MSI information' {
        $metadata = ConvertFrom-IpDetectionXml -XmlText (Get-TestDetectionXml -WithMsi)
        $metadata.ToolVersion | Should -Be '1.8.7.0'
        $metadata.FileName | Should -Be 'IntunePackage.intunewin'
        $metadata.SetupFile | Should -Be 'ContosoSetup.msi'
        $metadata.UnencryptedContentSize | Should -Be 2048
        $metadata.EncryptionKey | Should -Be 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA='
        $metadata.MsiInfo.ProductCode | Should -Be '{11111111-2222-3333-4444-555555555555}'
        $metadata.MsiInfo.IsMachine | Should -BeTrue
    }

    It 'defaults the profile identifier and digest algorithm' {
        $metadata = ConvertFrom-IpDetectionXml -XmlText (Get-TestDetectionXml -Omit ProfileIdentifier, FileDigestAlgorithm)
        $metadata.ProfileIdentifier | Should -Be 'ProfileVersion1'
        $metadata.FileDigestAlgorithm | Should -Be 'SHA256'
    }

    It 'fails when a required encryption field is missing' {
        { ConvertFrom-IpDetectionXml -XmlText (Get-TestDetectionXml -Omit EncryptionKey) } | Should -Throw '*missing required field: EncryptionKey*'
    }

    It 'fails on a document without ApplicationInfo' {
        { ConvertFrom-IpDetectionXml -XmlText '<Other/>' } | Should -Throw '*ApplicationInfo*'
    }

    It 'locates the encrypted payload inside the package' {
        $path = New-TestIntuneWin -Path (Join-Path $TestDrive 'pkg\ContosoSetup.intunewin') -XmlText (Get-TestDetectionXml) -PayloadSize 5000
        $metadata = Read-IpIntuneWinMetadata -Path $path
        $metadata.ContentEntryName | Should -Be 'IntuneWinPackage/Contents/IntunePackage.intunewin'
        $metadata.SizeEncrypted | Should -Be 5000
        $metadata.IntuneWinPath | Should -Be $path
    }

    It 'rejects files that are not .intunewin' {
        $path = Join-Path $TestDrive 'not-a-package.zip'
        Set-Content -LiteralPath $path -Value 'x'
        { Read-IpIntuneWinMetadata -Path $path } | Should -Throw '*Expected a .intunewin file*'
    }
}

Describe 'Detection rule payloads' {
    It 'builds an MSI product code rule with a version operator' {
        $rule = @(New-IpDetectionRules -Detection (New-TestDefinition).Detection)[0]
        $rule['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppProductCodeDetection'
        $rule.productVersionOperator | Should -Be 'greaterThanOrEqual'
        $rule.productVersion | Should -Be '2.4.0'
    }

    It 'omits the product version when no operator is configured' {
        $detection = (New-TestDefinition).Detection
        $detection.Msi.ProductVersionOperator = 'notConfigured'
        @(New-IpDetectionRules -Detection $detection)[0].productVersion | Should -BeNullOrEmpty
    }

    It 'builds file rules and defaults missing optional fields' {
        $detection = (New-TestDefinition -DetectionType File).Detection
        $detection.File = @([pscustomobject]@{ Path = 'C:\Program Files\Contoso'; FileOrFolderName = 'Sample.exe'; DetectionType = 'exists' })
        $rule = @(New-IpDetectionRules -Detection $detection)[0]
        $rule['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppFileSystemDetection'
        $rule.check32BitOn64System | Should -BeFalse
        $rule.operator | Should -Be 'notConfigured'
        $rule.detectionValue | Should -BeNullOrEmpty
    }

    It 'builds registry rules with a normalized key path' {
        $detection = (New-TestDefinition -DetectionType Registry).Detection
        $detection.Registry = @([pscustomobject]@{ KeyPath = 'HKLM\SOFTWARE\Contoso\Sample'; ValueName = 'Version'; DetectionType = 'version'; Operator = 'greaterThanOrEqual'; DetectionValue = '2.4.0'; Check32BitOn64System = $false })
        $rule = @(New-IpDetectionRules -Detection $detection)[0]
        $rule.keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\Sample'
        $rule.detectionType | Should -Be 'version'
        $rule.detectionValue | Should -Be '2.4.0'
    }

    It 'embeds a custom detection script as base64' {
        $scriptPath = Join-Path $TestDrive 'Detect-Contoso.ps1'
        Set-Content -LiteralPath $scriptPath -Value "Write-Output 'Detected'" -NoNewline
        $detection = (New-TestDefinition -DetectionType Script).Detection
        $detection.Script.ScriptPath = 'Detect-Contoso.ps1'
        $rule = @(New-IpDetectionRules -Detection $detection -BaseDirectory $TestDrive)[0]
        $rule['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppPowerShellScriptDetection'
        [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($rule.scriptContent)) | Should -Be "Write-Output 'Detected'"
    }

    It 'refuses invalid rules' {
        $detection = (New-TestDefinition -DetectionType Registry).Detection
        $detection.Registry = @([pscustomobject]@{ KeyPath = 'SOFTWARE\Contoso'; DetectionType = 'exists' })
        { New-IpDetectionRules -Detection $detection } | Should -Throw '*KeyPath must start with*'
    }
}

Describe 'win32LobApp payloads' {
    BeforeAll {
        $script:Metadata = ConvertFrom-IpDetectionXml -XmlText (Get-TestDetectionXml -WithMsi)
        $script:Metadata | Add-Member -NotePropertyName SizeEncrypted -NotePropertyValue ([int64]4096)
    }

    It 'uses allowedArchitectures and the documented install experience fields' {
        $body = New-IpWin32LobAppBody -Definition (New-TestDefinition) -IntuneWinMetadata $script:Metadata
        $body['@odata.type'] | Should -Be '#microsoft.graph.win32LobApp'
        $body.allowedArchitectures | Should -Be 'x64,arm64'
        $body.applicableArchitectures | Should -Be 'none'
        $body.minimumSupportedWindowsRelease | Should -Be 'Windows11_23H2'
        $body.installExperience.maxRunTimeInMinutes | Should -Be 60
        $body.Contains('installCommandLineTimeoutInMinutes') | Should -BeFalse
        $body.setupFilePath | Should -Be 'ContosoSetup.msi'
        $body.displayVersion | Should -Be '2.4.0'
    }

    It 'includes the Intune default return codes' {
        $body = New-IpWin32LobAppBody -Definition (New-TestDefinition) -IntuneWinMetadata $script:Metadata
        @($body.returnCodes | ForEach-Object { $_.returnCode }) | Should -Be @(0, 1707, 3010, 1641, 1618)
        @($body.returnCodes)[2].type | Should -Be 'softReboot'
    }

    It 'adds MSI information from Detection.xml' {
        $body = New-IpWin32LobAppBody -Definition (New-TestDefinition) -IntuneWinMetadata $script:Metadata
        $body.msiInformation.productCode | Should -Be '{11111111-2222-3333-4444-555555555555}'
        $body.msiInformation.packageType | Should -Be 'perMachine'
    }

    It 'embeds the logo as mimeContent' {
        New-TestPngFile -Path (Join-Path $TestDrive 'logo.png') | Out-Null
        $definition = New-TestDefinition
        $definition.App.LogoPath = 'logo.png'
        $body = New-IpWin32LobAppBody -Definition $definition -IntuneWinMetadata $script:Metadata -BaseDirectory $TestDrive
        $body.largeIcon.type | Should -Be 'image/png'
        $body.largeIcon.value | Should -Match '^iVBOR'
    }

    It 'refuses to build a body from an invalid definition' {
        { New-IpWin32LobAppBody -Definition (New-IpAppDefinition) -IntuneWinMetadata $script:Metadata } | Should -Throw '*Invalid app definition*'
    }

    It 'builds content file, commit and finalize bodies' {
        $file = New-IpContentFileBody -IntuneWinMetadata $script:Metadata
        $file.size | Should -Be 2048
        $file.sizeEncrypted | Should -Be 4096
        $commit = New-IpCommitBody -IntuneWinMetadata $script:Metadata
        $commit.fileEncryptionInfo.profileIdentifier | Should -Be 'ProfileVersion1'
        $commit.fileEncryptionInfo.mac | Should -Be 'DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD='
        $patch = New-IpFinalizePatchBody -AppBody (New-IpWin32LobAppBody -Definition (New-TestDefinition) -IntuneWinMetadata $script:Metadata) -ContentVersionId '7'
        $patch.committedContentVersion | Should -Be '7'
        $patch.Contains('applicableArchitectures') | Should -BeFalse
        $patch.allowedArchitectures | Should -Be 'x64,arm64'
    }

    It 'redacts secrets and truncates large values for display' {
        $body = New-IpCommitBody -IntuneWinMetadata $script:Metadata
        $body.largeIcon = [ordered]@{ value = ('A' * 500) }
        $safe = Protect-IpPayloadForDisplay -Payload $body
        $safe.fileEncryptionInfo.encryptionKey | Should -Be '<redacted>'
        $safe.fileEncryptionInfo.profileIdentifier | Should -Be 'ProfileVersion1'
        $safe.largeIcon.value | Should -Match '<500 characters>$'
        $body.fileEncryptionInfo.encryptionKey | Should -Not -Be '<redacted>'
    }

    It 'builds the Intune admin center URL of an app' {
        Get-IpIntunePortalAppUrl -AppId '00000000-0000-0000-0000-000000000009' | Should -Be 'https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/00000000-0000-0000-0000-000000000009'
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBR8uvVpkbviNzT
# BzXTBrLgqprrFLmUJQphXSAb3pgfoaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMpJeXBo8B5O/yMQDH/UXpYj9ZEuRDWNmLnwPo+TbMUJMA0GCSqG
# SIb3DQEBAQUABIIBgDoDsdT20rOhFIkpsN/6QNeWMvZQxn0AiydGkzv7oSdvVnnb
# 2hRhqq4kvank457+FccJJAVknk8kCaBbf4qB27jy1BPj2BvMroGOM+gAXbWjxw2d
# sJvtKSASQSl0cX/tb7C7p7YuiUsPDStWkuWYwvH82Lq4zzPt3ZZ+fJmfmB4oSnkM
# Mu0HVRXNqXD6q6ByWzll1BTOPo3gybZERMa9itSDnaptn2c+MOfnAtkJvumUatAR
# GIpyLdwOX23zqJZBNSqP3spIeXV7ZqcR9vIIP35TLFvKXn+DleAog3cPoa0ilgXr
# YGkV1zCMBBRG0K+ut2CfY+NK4C0RKalJWDfLadL5nquyQGBSWDzunz2gr6sapfWl
# hmn64bCrE/8uj17gyJ9OstF3qkGMr7BAVuuuTwQCso1W4Z4IAQrvBNFvn2WS8mC1
# L0wGNXv6Grqx1omNhbNC191rJSmq+M8+FX+FoTzb6c2OiLrAIXWMx3qJsTMQuVl/
# J/V6cumWaSPa3uTMbaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzVaMC8GCSqGSIb3DQEJBDEiBCA4wFsRdBdUMX67EHIll3HD0LU231/GlVPtTA1k
# k6JNvzANBgkqhkiG9w0BAQEFAASCAgA4Y8L5anNti2IA6JiohVWSy+0FO0EkWhSu
# PXiN3phPmK3rokqOk3Zwt9gXYJjS2FQKvpL3htbWiaHINiUMJMVo1lRpShg6m4Vm
# QWt8xJp1Mb4X7CdQDJd/zGTE1AHS659LMkdhOLMrz+34NiCoI6sKP/JcJEUWaKbp
# LJEmOo+Qn2+GRa3afwKMdHk3R65stn3NAlSRCzP3t+FllG3RttuuKioWpzIe+R1X
# 0C4RB6PF5d7fh/t36XWuIiceZNf/iNpF6udnXz5xjOECBLk+6rg5b3EqqRM8YKwN
# /KGn3HQMVyJunW/Xfb2ooyuHnjp2j0omLiU9gbO4e20j3laNzJmNe7DplnVAko81
# IKPaaUou1lfxx/KSu67mInGCRCJMUV49pf5z/lBnMAcoyABBdKbAV3tse/EY0HJQ
# riEGWxjGYVRnbif1Msxd8kHAml6l924N2Yj10DM7AicvFmnoHqO9sBVdGc7gwggV
# Iy4chqIyfMddF+D3zgXdyjczPFJ3Zz5pyBkZmFn2osHfxCHmNoOWXnkfakofHCLC
# Ghf23UfF7Q78Zq+JmgOX73dOFdV+iZXmnj3oOU3JT+/DMXRn4eyNoAGq2LUg7Lsz
# xfvvjZ3V3qruFtRScSnhnEKkh7aYCUB1CpAgVby/xpMs7VlF/92NR9VzYRsb/yKC
# Kkg7TLO3Bw==
# SIG # End signature block
