#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for the differential UpdateExisting PATCH, the import from Intune and the
    post-publication check. Fictional Contoso fixtures only, shaped after anonymized Graph responses.

.VERSION
    1.0.3
#>

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Graph.psm1') -Force
    . (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-TestHelpers.ps1')
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
    Set-IpGraphSettings -BaseUri 'https://graph.microsoft.com/beta' -MaxRetries 1 -PollSeconds 1 -PollTimeoutMinutes 1
    $script:ModuleName = 'SmartM365.IntunePackager.Graph'

    function ConvertTo-B64 { param([string]$Text) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)) }
    function Copy-Fixture { param($Object) ($Object | ConvertTo-Json -Depth 30 | ConvertFrom-Json) }
    $script:DetectScript = "`$path = 'HKLM:\SOFTWARE\Contoso\SampleApp'`r`nif (Test-Path `$path) { Write-Output 'Installed' }`r`n"
    $script:RequirementScript = "Write-Output 'OK'"
    $script:Modified = [datetime]::new(2026, 10, 18, 9, 31, 0, [DateTimeKind]::Utc)

    function New-ContosoApp {
        param([ValidateSet('Msi', 'ScriptWithRequirement', 'RegistryLeftover', 'LegacyArchitecture', 'NotRestricted', 'Mixed')][string]$Kind)
        $base = [ordered]@{
            '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'c0ffee00-1111-4222-8333-000000000001'; displayName = 'Contoso Sample App'; description = 'Fictional sample app.'
            publisher = 'Contoso Ltd'; displayVersion = '2.4.0'; notes = $null; informationUrl = $null; privacyInformationUrl = $null; developer = ''; owner = ''
            isFeatured = $false; allowAvailableUninstall = $true; roleScopeTagIds = @('0'); committedContentVersion = '2'; fileName = 'IntunePackage.intunewin'; setupFilePath = 'ContosoSetup.msi'
            installCommandLine = 'msiexec /i "ContosoSetup.msi" /qn'; uninstallCommandLine = 'msiexec /x {11111111-2222-3333-4444-555555555555} /qn'
            applicableArchitectures = 'none'; allowedArchitectures = 'x64'; minimumFreeDiskSpaceInMB = $null; minimumMemoryInMB = $null; minimumNumberOfProcessors = $null; minimumCpuSpeedInMHz = $null
            minimumSupportedWindowsRelease = '2H20'; minimumSupportedOperatingSystem = [ordered]@{ v10_1607 = $false; v10_2H20 = $true }
            installExperience = [ordered]@{ runAsAccount = 'system'; maxRunTimeInMinutes = 60; deviceRestartBehavior = 'allow'; inUseBehavior = 'notConfigured' }
            returnCodes = @([ordered]@{ returnCode = 0; type = 'success' }, [ordered]@{ returnCode = 1707; type = 'success' }, [ordered]@{ returnCode = 3010; type = 'softReboot' }, [ordered]@{ returnCode = 1641; type = 'hardReboot' }, [ordered]@{ returnCode = 1618; type = 'retry' })
            msiInformation = [ordered]@{ productCode = '{11111111-2222-3333-4444-555555555555}'; productVersion = '2.4.0'; upgradeCode = '{66666666-7777-8888-9999-000000000000}'; requiresReboot = $false; packageType = 'perMachine'; productName = 'Contoso Sample App'; publisher = 'Contoso Ltd' }
            activeInstallScript = $null; activeUninstallScript = $null; largeIcon = $null; lastModifiedDateTime = $script:Modified
            detectionRules = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppProductCodeDetection'; productCode = '{11111111-2222-3333-4444-555555555555}'; productVersionOperator = 'notConfigured'; productVersion = $null })
            requirementRules = @()
            rules = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppProductCodeRule'; ruleType = 'detection'; productCode = '{11111111-2222-3333-4444-555555555555}'; productVersionOperator = 'notConfigured'; productVersion = $null })
        }
        switch ($Kind) {
            'ScriptWithRequirement' {
                $base.allowedArchitectures = 'x64,arm64'; $base.minimumFreeDiskSpaceInMB = 40960; $base.msiInformation = $null; $base.setupFilePath = 'Install.ps1'
                $base.installCommandLine = 'powershell.exe -ExecutionPolicy Bypass -File .\Install.ps1'; $base.uninstallCommandLine = 'powershell.exe -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall'
                $base.detectionRules = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptDetection'; enforceSignatureCheck = $false; runAs32Bit = $false; scriptContent = (ConvertTo-B64 $script:DetectScript) })
                $requirement = [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRequirement'; operator = 'equal'; detectionValue = 'OK'; displayName = 'Contoso readiness check'; enforceSignatureCheck = $false; runAs32Bit = $false; runAsAccount = 'system'; scriptContent = (ConvertTo-B64 $script:RequirementScript); detectionType = 'string' }
                $base.requirementRules = @($requirement)
                $base.rules = @(
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRule'; ruleType = 'detection'; displayName = $null; enforceSignatureCheck = $false; runAs32Bit = $false; runAsAccount = $null; scriptContent = (ConvertTo-B64 $script:DetectScript); operationType = 'notConfigured'; operator = 'notConfigured'; comparisonValue = $null }
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppPowerShellScriptRule'; ruleType = 'requirement'; displayName = 'Contoso readiness check'; enforceSignatureCheck = $false; runAs32Bit = $false; runAsAccount = 'system'; scriptContent = (ConvertTo-B64 $script:RequirementScript); operationType = 'string'; operator = 'equal'; comparisonValue = 'OK' }
                )
                $base.largeIcon = [ordered]@{ type = 'image/png'; value = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==' }
            }
            'RegistryLeftover' {
                $base.msiInformation = $null
                $base.detectionRules = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppRegistryDetection'; check32BitOn64System = $true; keyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\SampleApp'; valueName = 'Version'; detectionType = 'exists'; operator = 'greaterThanOrEqual'; detectionValue = '' })
                $base.rules = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppRegistryRule'; ruleType = 'detection'; check32BitOn64System = $true; keyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\SampleApp'; valueName = 'Version'; operationType = 'exists'; operator = 'greaterThanOrEqual'; comparisonValue = '' })
            }
            'LegacyArchitecture' { $base.allowedArchitectures = $null; $base.applicableArchitectures = 'x86,x64' }
            'NotRestricted' { $base.allowedArchitectures = $null; $base.applicableArchitectures = 'none' }
            'Mixed' {
                $base.msiInformation = $null
                $base.detectionRules = @(
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppFileSystemDetection'; path = 'C:\Program Files\Contoso'; fileOrFolderName = 'Sample.exe'; check32BitOn64System = $false; detectionType = 'exists'; operator = 'notConfigured'; detectionValue = $null }
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppRegistryDetection'; check32BitOn64System = $false; keyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso'; valueName = 'Edition'; detectionType = 'string'; operator = 'equal'; detectionValue = 'Pro' }
                )
                $base.rules = @(
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppFileSystemRule'; ruleType = 'detection'; path = 'C:\Program Files\Contoso'; fileOrFolderName = 'Sample.exe'; check32BitOn64System = $false; operationType = 'exists'; operator = 'notConfigured'; comparisonValue = $null }
                    [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppRegistryRule'; ruleType = 'detection'; check32BitOn64System = $false; keyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso'; valueName = 'Edition'; operationType = 'string'; operator = 'equal'; comparisonValue = 'Pro' }
                )
            }
        }
        return ($base | ConvertTo-Json -Depth 30 | ConvertFrom-Json)
    }

    function New-ImportedDefinition {
        param($App, [string]$Folder)
        $import = ConvertFrom-IpWin32App -App $App
        return (Save-IpImportAssets -Import $import -Folder $Folder -BaseName 'Contoso-SampleApp')
    }

    function New-PackageMetadata {
        param([string]$SetupFile = 'ContosoSetup.msi', [switch]$NoMsi, [string]$ProductCode = '{11111111-2222-3333-4444-555555555555}')
        $metadata = ConvertFrom-IpDetectionXml -XmlText (Get-TestDetectionXml -WithMsi:(-not $NoMsi) -SetupFile $SetupFile)
        $metadata | Add-Member -NotePropertyName SizeEncrypted -NotePropertyValue ([int64]4096)
        if (-not $NoMsi) { $metadata.MsiInfo.ProductCode = $ProductCode; $metadata.MsiInfo.ProductVersion = '2.4.0'; $metadata.MsiInfo.UpgradeCode = '{66666666-7777-8888-9999-000000000000}' }
        return $metadata
    }

    function Get-Diff {
        param($Definition, $Current, $Metadata, [string]$Base)
        $body = New-IpWin32LobAppBody -Definition $Definition -IntuneWinMetadata $Metadata -BaseDirectory $Base
        return (New-IpUpdatePatchBody -AppBody $body -Current $Current -ContentVersionId '3' -Definition $Definition)
    }
}

AfterAll { Set-IpLogSink -Sink $null -Console $false }

Describe 'Import from Intune' {
    It 'imports an MSI app as an update definition' {
        $import = ConvertFrom-IpWin32App -App (New-ContosoApp Msi)
        $d = $import.Definition
        $d.Publish.Mode | Should -Be 'UpdateExisting'
        $d.Publish.ExistingAppId | Should -Be 'c0ffee00-1111-4222-8333-000000000001'
        $d.Detection.Type | Should -Be 'MsiProductCode'
        $d.Detection.Msi.ProductCode | Should -Be '{11111111-2222-3333-4444-555555555555}'
        $d.Package.SourceFolder | Should -Be ''
        $d.Package.SetupFile | Should -Be 'ContosoSetup.msi'
        @($d.ReturnCodes).Count | Should -Be 5
        $d.Program.DeviceRestartBehavior | Should -Be 'allow'
        $d.ImportedFrom.ContentVersion | Should -Be '2'
        $d.ImportedFrom.LastModifiedDateTime | Should -Be '2026-10-18T09:31:00Z'
        $import.Preserved | Should -Contain 'allowAvailableUninstall: True.'
        Test-IpAppDefinition -Definition $d -ForPublish | Should -BeNullOrEmpty
    }

    It 'lists every preserved property that has a value, false included' {
        $app = New-ContosoApp Msi
        $import = ConvertFrom-IpWin32App -App $app
        $import.Preserved | Should -Contain 'isFeatured: False.'
        $import.Preserved | Should -Contain 'roleScopeTagIds: 1 (0).'
        $import.Preserved | Should -Contain 'installExperience.inUseBehavior: notConfigured.'
        $import.Preserved | Should -Contain 'minimumSupportedOperatingSystem: v10_2H20.'
        ($import.Preserved -join "`n") | Should -Not -Match 'minimumMemoryInMB'
        $app.isFeatured = $true
        $app.minimumMemoryInMB = 4096
        $import = ConvertFrom-IpWin32App -App $app
        $import.Preserved | Should -Contain 'isFeatured: True.'
        $import.Preserved | Should -Contain 'minimumMemoryInMB: 4096.'
    }

    It 'imports a script detection, requirement rules, disk space and logo' {
        $import = ConvertFrom-IpWin32App -App (New-ContosoApp ScriptWithRequirement)
        $import.Definition.Detection.Type | Should -Be 'Script'
        $import.ScriptText | Should -Be $script:DetectScript
        $import.Definition.Requirements.MinimumFreeDiskSpaceInMB | Should -Be 40960
        $import.Definition.Requirements.Architectures | Should -Be @('arm64', 'x64')
        ($import.Preserved -join "`n") | Should -Match "Requirement rules: 1 \(PowerShellScript 'Contoso readiness check'\)"
        $import.IconBytes.Length | Should -BeGreaterThan 0
        $definition = Save-IpImportAssets -Import $import -Folder (Join-Path $TestDrive 'assets') -BaseName 'Contoso Sample App' -Relative
        $definition.Detection.Script.ScriptPath | Should -Be 'Contoso-Sample-App-Detection.ps1'
        $definition.App.LogoPath | Should -Be 'Contoso-Sample-App-Logo.png'
        Test-Path (Join-Path $TestDrive 'assets\Contoso-Sample-App-Detection.ps1') | Should -BeTrue
    }

    It 'maps legacy and unrestricted architectures' {
        (ConvertFrom-IpWin32App -App (New-ContosoApp LegacyArchitecture)).Definition.Requirements.Architectures | Should -Be @('x64', 'x86')
        (ConvertFrom-IpWin32App -App (New-ContosoApp NotRestricted)).Definition.Requirements.Architectures | Should -Be @('NotRestricted')
    }

    It 'keeps a mixed detection as preserved, with a warning' {
        $import = ConvertFrom-IpWin32App -App (New-ContosoApp Mixed)
        $import.Definition.Detection.Type | Should -Be 'Preserved'
        ($import.Warnings -join "`n") | Should -Match 'kept as in Intune'
        Test-IpAppDefinition -Definition $import.Definition -ForPublish | Should -BeNullOrEmpty
        $import.Definition.Publish.Mode = 'Create'
        Test-IpAppDefinition -Definition $import.Definition -ForPublish | Should -Match 'only valid with Publish.Mode UpdateExisting'
    }
}

Describe 'Import is read-only (static check)' {
    BeforeAll {
        $script:Writers = @('New-IpWin32App', 'New-IpContentVersion', 'New-IpContentFile', 'Request-IpSasRenewal', 'Send-IpIntuneWinContent', 'Send-IpAzureBlobWithRetry', 'Submit-IpContentFileCommit', 'Complete-IpAppContentVersion', 'Publish-IpWin32App', 'Invoke-IpGraphRequest', 'Invoke-IpGraphRawRequest', 'Invoke-MgGraphRequest', 'Invoke-RestMethod', 'Invoke-WebRequest')
        function Get-FunctionCommands {
            param([string]$Path, [string]$Name)
            $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
            $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true)
            if (-not $function) { throw "Function not found: $Name" }
            return @($function.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
        }
    }

    It 'converts and saves without any Graph command' -ForEach @(@{ Name = 'ConvertFrom-IpWin32App' }, @{ Name = 'Save-IpImportAssets' }) {
        @(Get-FunctionCommands -Path (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Name $Name | Where-Object { $_ -in $script:Writers }) | Should -BeNullOrEmpty
    }

    It 'reads the app with GET only in the CLI and the GUI import paths' {
        $gui = Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'
        $cli = @(Get-FunctionCommands -Path $gui -Name 'Invoke-IpCliImport')
        $cli | Should -Contain 'Get-IpWin32App'
        @($cli | Where-Object { $_ -in $script:Writers }) | Should -BeNullOrEmpty
        @(Get-FunctionCommands -Path $gui -Name 'Invoke-IpImportResult' | Where-Object { $_ -in $script:Writers -or $_ -eq 'Start-IpJob' }) | Should -BeNullOrEmpty
        $details = @(Get-FunctionCommands -Path $gui -Name 'Start-IpAppDetailsJob')
        $details | Should -Contain 'Get-IpWin32App'
        @($details | Where-Object { $_ -in $script:Writers }) | Should -BeNullOrEmpty
        $graph = Join-Path $script:AppRoot 'SmartM365.IntunePackager.Graph.psm1'
        $read = (Get-Content $graph -Raw)
        $read | Should -Match "Invoke-IpGraphRequest -Method GET -Uri ""deviceAppManagement/mobileApps/\`$AppId"""
    }
}

Describe 'Differential PATCH' {
    It 'sends only the content version and the package properties when nothing changed' -ForEach @(
        @{ Kind = 'Msi' }, @{ Kind = 'ScriptWithRequirement' }, @{ Kind = 'RegistryLeftover' }, @{ Kind = 'LegacyArchitecture' }, @{ Kind = 'NotRestricted' }, @{ Kind = 'Mixed' }
    ) {
        $app = New-ContosoApp $Kind
        $folder = Join-Path $TestDrive "unchanged-$Kind"
        $definition = New-ImportedDefinition -App $app -Folder $folder
        $metadata = if ($app.msiInformation) { New-PackageMetadata -SetupFile $app.setupFilePath } else { New-PackageMetadata -SetupFile $app.setupFilePath -NoMsi }
        $diff = Get-Diff -Definition $definition -Current $app -Metadata $metadata -Base $folder
        @($diff.Changes | Where-Object { $_.Property -notin @('fileName', 'setupFilePath') }) | Should -BeNullOrEmpty
        $diff.DetectionChanged | Should -BeFalse
        @($diff.Body.Keys | Where-Object { $_ -notin @('@odata.type', 'fileName', 'setupFilePath', 'msiInformation', 'committedContentVersion') }) | Should -BeNullOrEmpty
        foreach ($never in @('isFeatured', 'allowAvailableUninstall', 'roleScopeTagIds', 'requirementRules', 'detectionRules', 'minimumSupportedOperatingSystem', 'applicableArchitectures')) { $diff.Body.Contains($never) | Should -BeFalse }
    }

    It 'treats a script saved with a BOM and LF line endings as unchanged' {
        $app = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'bom'
        $definition = New-ImportedDefinition -App $app -Folder $folder
        $path = Join-Path $folder $([IO.Path]::GetFileName($definition.Detection.Script.ScriptPath))
        [IO.File]::WriteAllText($path, ($script:DetectScript.Replace("`r`n", "`n") + "`n`n"), [Text.UTF8Encoding]::new($true))
        (Get-Diff -Definition $definition -Current $app -Metadata (New-PackageMetadata -SetupFile 'Install.ps1' -NoMsi) -Base $folder).DetectionChanged | Should -BeFalse
    }

    It 'sends a changed detection through rules with the existing requirement rules' {
        $app = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'detection'
        $definition = New-ImportedDefinition -App $app -Folder $folder
        Set-Content -LiteralPath (Join-Path $folder 'Contoso-SampleApp-Detection.ps1') -Value "Write-Output 'Detected v3'"
        $diff = Get-Diff -Definition $definition -Current $app -Metadata (New-PackageMetadata -SetupFile 'Install.ps1' -NoMsi) -Base $folder
        $diff.DetectionChanged | Should -BeTrue
        $diff.Body.Contains('detectionRules') | Should -BeFalse
        $rules = @($diff.Body.rules)
        @($rules | Where-Object { $_.ruleType -eq 'detection' }).Count | Should -Be 1
        $requirement = @($rules | Where-Object { $_.ruleType -eq 'requirement' })
        $requirement.Count | Should -Be 1
        (Get-IpRuleSignature -Rule $requirement[0] -RuleType requirement) | Should -Be (Get-IpRuleSignature -Rule $app.requirementRules[0] -RuleType requirement)
        $diff.PreservedRequirementRules | Should -Be 1
    }

    It 'merges installExperience and lists the edited properties' {
        $app = New-ContosoApp Msi
        $folder = Join-Path $TestDrive 'edited'
        $definition = New-ImportedDefinition -App $app -Folder $folder
        $definition.App.DisplayName = 'Contoso Sample App 3'
        $definition.Program.MaxRunTimeInMinutes = 90
        $diff = Get-Diff -Definition $definition -Current $app -Metadata (New-PackageMetadata) -Base $folder
        @($diff.Changes.Property) | Should -Contain 'displayName'
        @($diff.Changes.Property) | Should -Contain 'installExperience.maxRunTimeInMinutes'
        $diff.Body.installExperience.maxRunTimeInMinutes | Should -Be 90
        $diff.Body.installExperience.inUseBehavior | Should -Be 'notConfigured'
        $diff.Body.installExperience.deviceRestartBehavior | Should -Be 'allow'
        $diff.Body.Contains('description') | Should -BeFalse
    }

    It 'sends a disk space change and warns when the new package is not an MSI' {
        $app = New-ContosoApp Msi
        $folder = Join-Path $TestDrive 'disk'
        $definition = New-ImportedDefinition -App $app -Folder $folder
        $definition.Requirements.MinimumFreeDiskSpaceInMB = 2048
        $diff = Get-Diff -Definition $definition -Current $app -Metadata (New-PackageMetadata -NoMsi) -Base $folder
        $diff.Body.minimumFreeDiskSpaceInMB | Should -Be 2048
        $diff.Body.Contains('msiInformation') | Should -BeTrue
        $null -eq $diff.Body.msiInformation | Should -BeTrue
        ($diff.Warnings -join ' ') | Should -Match 'not an MSI'
    }
}

Describe 'Concurrency and post-publication check' {
    It 'detects a modification made in Intune after the import' {
        $app = New-ContosoApp Msi
        $definition = (ConvertFrom-IpWin32App -App $app).Definition
        Test-IpImportConcurrency -Definition $definition -Current $app | Should -Be ''
        $app.lastModifiedDateTime = $script:Modified.AddMinutes(5)
        Test-IpImportConcurrency -Definition $definition -Current $app | Should -Match '^\[Concurrency\]'
        Test-IpImportConcurrency -Definition $definition -Current $app -ExpectedLastModified '2026-10-18T09:36:00Z' | Should -Be ''
    }

    It 'reports lost requirement rules and changed preserved properties' {
        $before = New-ContosoApp ScriptWithRequirement
        $definition = New-ImportedDefinition -App $before -Folder (Join-Path $TestDrive 'verify')
        $update = Get-Diff -Definition $definition -Current $before -Metadata (New-PackageMetadata -SetupFile 'Install.ps1' -NoMsi) -Base (Join-Path $TestDrive 'verify')
        $after = Copy-Fixture $before
        $after.committedContentVersion = '3'
        Test-IpUpdateResult -Before $before -After $after -Update $update -ContentVersionId '3' | Should -BeNullOrEmpty
        $after.rules = @($after.rules | Where-Object { $_.ruleType -eq 'detection' })
        $after.isFeatured = $true
        $issues = Test-IpUpdateResult -Before $before -After $after -Update $update -ContentVersionId '3'
        ($issues -join "`n") | Should -Match 'Requirement rules changed: 1 before, 0 after'
        ($issues -join "`n") | Should -Match 'isFeatured changed'
    }
}

Describe 'UpdateExisting publication with mocked Graph' {
    BeforeAll {
        $script:Package = New-TestIntuneWin -Path (Join-Path $TestDrive 'pkg\Install.intunewin') -XmlText (Get-TestDetectionXml -SetupFile 'Install.ps1') -PayloadSize 2048

        function Initialize-UpdateFake {
            param($Before, [switch]$DropRequirements)
            $script:Fake = @{ Before = $Before; After = $null; Patches = [Collections.Generic.List[object]]::new(); Requests = [Collections.Generic.List[string]]::new(); Committed = $false; DropRequirements = [bool]$DropRequirements }
            Mock -ModuleName $script:ModuleName Start-IpSleep { }
            Mock -ModuleName $script:ModuleName Get-IpGraphContext { [pscustomobject]@{ TenantId = 'b4c9e2d1-7a3f-4e8b-9c21-5d6f7a8b9c0e' } }
            Mock -ModuleName $script:ModuleName Send-IpIntuneWinContent { 1 }
            Mock -ModuleName $script:ModuleName Invoke-IpGraphRawRequest {
                $fake = $script:Fake
                $fake.Requests.Add("$Method $Uri")
                switch -Regex ("$Method $Uri") {
                    '^GET .*/mobileApps/[0-9a-f-]+$' { if ($fake.After) { return $fake.After } else { return $fake.Before } }
                    '^POST .*/contentVersions$' { return [pscustomobject]@{ id = '3' } }
                    '^POST .*/files$' { return [pscustomobject]@{ id = 'file-1' } }
                    '^POST .*/commit$' { $fake.Committed = $true; return $null }
                    '^GET .*/files/file-1$' { if ($fake.Committed) { return [pscustomobject]@{ uploadState = 'commitFileSuccess'; azureStorageUri = $null } }; return [pscustomobject]@{ uploadState = 'azureStorageUriRequestSuccess'; azureStorageUri = 'https://contoso.blob.core.windows.net/c/f?sig=1' } }
                    '^PATCH .*/mobileApps/[0-9a-f-]+$' {
                        $patch = $Body | ConvertFrom-Json
                        $fake.Patches.Add($patch)
                        $after = $fake.Before | ConvertTo-Json -Depth 30 | ConvertFrom-Json
                        foreach ($property in $patch.PSObject.Properties) {
                            if ($property.Name -eq '@odata.type') { continue }
                            if ($after.PSObject.Properties[$property.Name]) { $after.($property.Name) = $property.Value } else { $after | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value }
                        }
                        if ($fake.DropRequirements) { $after.rules = @($after.rules | Where-Object { $_.ruleType -eq 'detection' }); $after.requirementRules = @() }
                        $after.lastModifiedDateTime = [datetime]::new(2026, 10, 18, 10, 0, 0, [DateTimeKind]::Utc)
                        $fake.After = $after
                        return $null
                    }
                    default { throw "Unexpected request: $Method $Uri" }
                }
            }
        }
    }

    It 'publishes a differential PATCH and passes the post-publication check' {
        $before = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'publish-ok'
        $definition = New-ImportedDefinition -App $before -Folder $folder
        Initialize-UpdateFake -Before $before
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Confirm:$false
        $result.ContentVersionId | Should -Be '3'
        @($result.VerificationIssues) | Should -BeNullOrEmpty
        $patch = $script:Fake.Patches[0]
        $patch.committedContentVersion | Should -Be '3'
        foreach ($never in @('isFeatured', 'allowAvailableUninstall', 'roleScopeTagIds', 'requirementRules', 'detectionRules', 'rules', 'installExperience', 'returnCodes', 'minimumSupportedOperatingSystem')) { $patch.PSObject.Properties.Name | Should -Not -Contain $never }
        @($script:Fake.Requests | Where-Object { $_ -match '^POST .*/mobileApps$' }).Count | Should -Be 0
    }

    It 'blocks when the app changed in Intune after the import, and publishes with -Force' {
        $before = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'publish-concurrency'
        $definition = New-ImportedDefinition -App $before -Folder $folder
        $before.lastModifiedDateTime = $script:Modified.AddHours(1)
        Initialize-UpdateFake -Before $before
        { Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Confirm:$false } | Should -Throw '*`[Concurrency`]*'
        @($script:Fake.Requests | Where-Object { $_ -match 'contentVersions' }).Count | Should -Be 0
        Initialize-UpdateFake -Before $before
        (Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Force -Confirm:$false).ContentVersionId | Should -Be '3'
    }

    It 'blocks before the confirmation, flags the error and does not log it itself' {
        $before = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'publish-concurrency-log'
        $definition = New-ImportedDefinition -App $before -Folder $folder
        $before.lastModifiedDateTime = $script:Modified.AddHours(1)
        Initialize-UpdateFake -Before $before
        $script:LogLines = [Collections.Generic.List[string]]::new()
        Set-IpLogSink -Sink { param($line) $script:LogLines.Add([string]$line) } -Console $false
        try {
            $caught = $null
            try { [void](Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Confirm:$false) } catch { $caught = $_ }
        }
        finally { Set-IpLogSink -Sink $null -Console $false }
        $caught.Exception.Message | Should -Match '^\[Concurrency\]'
        $caught.Exception.Data['IpConcurrency'] | Should -BeTrue
        @($script:LogLines | Where-Object { $_ -match 'Concurrency' }).Count | Should -Be 0

        # Static order check: the concurrency test runs before ShouldProcess.
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'SmartM365.IntunePackager.Graph.psm1'), [ref]$null, [ref]$null)
        $publish = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Publish-IpWin32App' }, $true)
        $check = $publish.Find({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Test-IpImportConcurrency' }, $true)
        $confirm = $publish.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Member.Extent.Text -eq 'ShouldProcess' }, $true)
        $check.Extent.StartOffset | Should -BeLessThan $confirm.Extent.StartOffset
    }

    It 'sends a changed detection with the requirement rules and verifies both' {
        $before = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'publish-detection'
        $definition = New-ImportedDefinition -App $before -Folder $folder
        Set-Content -LiteralPath (Join-Path $folder 'Contoso-SampleApp-Detection.ps1') -Value "Write-Output 'Detected v3'"
        Initialize-UpdateFake -Before $before
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Confirm:$false
        @($result.VerificationIssues) | Should -BeNullOrEmpty
        $patch = $script:Fake.Patches[0]
        $patch.PSObject.Properties.Name | Should -Not -Contain 'detectionRules'
        @($patch.rules | Where-Object { $_.ruleType -eq 'requirement' }).Count | Should -Be 1
    }

    It 'reports requirement rules lost by Intune after the update' {
        $before = New-ContosoApp ScriptWithRequirement
        $folder = Join-Path $TestDrive 'publish-lost'
        $definition = New-ImportedDefinition -App $before -Folder $folder
        Set-Content -LiteralPath (Join-Path $folder 'Contoso-SampleApp-Detection.ps1') -Value "Write-Output 'Detected v3'"
        Initialize-UpdateFake -Before $before -DropRequirements
        $result = Publish-IpWin32App -Definition $definition -IntuneWinPath $script:Package -BaseDirectory $folder -Confirm:$false
        (@($result.VerificationIssues) -join "`n") | Should -Match 'Requirement rules changed: 1 before, 0 after'
    }

    Context 'successive publications with the same profile' {
        BeforeAll {
            $guiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'), [ref]$null, [ref]$null)
            $function = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Sync-IpImportedFromAfterPublish' }, $true)
            . ([scriptblock]::Create($function.Extent.Text))
            function New-ProfileOnDisk {
                param([string]$Name)
                $folder = Join-Path $TestDrive $Name
                $before = New-ContosoApp ScriptWithRequirement
                $definition = New-ImportedDefinition -App $before -Folder $folder
                $path = Join-Path $folder 'Contoso-SampleApp.profile.json'
                Export-IpAppProfile -Definition $definition -Path $path -Confirm:$false
                return [pscustomobject]@{ Before = $before; Folder = $folder; Path = $path }
            }
        }

        It 'does not raise [Concurrency] on the second publication when nobody changed the app' {
            $profile = New-ProfileOnDisk 'successive'
            Initialize-UpdateFake -Before $profile.Before
            $first = Publish-IpWin32App -Definition (Import-IpAppProfile -Path $profile.Path) -IntuneWinPath $script:Package -BaseDirectory $profile.Folder -Confirm:$false
            $first.VerifiedState.LastModifiedDateTime | Should -Be '2026-10-18T10:00:00Z'
            Sync-IpImportedFromAfterPublish -Definition (Import-IpAppProfile -Path $profile.Path) -Result $first -ProfilePath $profile.Path | Should -BeTrue
            $saved = Import-IpAppProfile -Path $profile.Path
            ConvertTo-IpIsoUtc $saved.ImportedFrom.LastModifiedDateTime | Should -Be '2026-10-18T10:00:00Z'
            $saved.ImportedFrom.ContentVersion | Should -Be '3'
            $saved.ImportedFrom.SyncedAt | Should -Not -BeNullOrEmpty
            $saved.Publish.Mode | Should -Be 'UpdateExisting'

            Initialize-UpdateFake -Before $script:Fake.After
            $second = Publish-IpWin32App -Definition $saved -IntuneWinPath $script:Package -BaseDirectory $profile.Folder -Confirm:$false
            $second.ContentVersionId | Should -Be '3'
            @($script:Fake.Requests | Where-Object { $_ -match 'contentVersions$' }).Count | Should -Be 1
        }

        It 'still blocks the second publication when the profile was not updated' {
            $profile = New-ProfileOnDisk 'stale'
            Initialize-UpdateFake -Before $profile.Before
            [void](Publish-IpWin32App -Definition (Import-IpAppProfile -Path $profile.Path) -IntuneWinPath $script:Package -BaseDirectory $profile.Folder -Confirm:$false)
            Initialize-UpdateFake -Before $script:Fake.After
            { Publish-IpWin32App -Definition (Import-IpAppProfile -Path $profile.Path) -IntuneWinPath $script:Package -BaseDirectory $profile.Folder -Confirm:$false } | Should -Throw '*`[Concurrency`]*'
        }

        It 'leaves a profile imported from another app, or never imported, unchanged' {
            $profile = New-ProfileOnDisk 'other-app'
            $state = [pscustomobject]@{ AppId = 'c0ffee00-9999-4222-8333-000000000009'; DisplayName = 'Other'; ContentVersion = '7'; LastModifiedDateTime = '2026-10-18T11:00:00Z' }
            $definition = Import-IpAppProfile -Path $profile.Path
            Update-IpImportedFrom -Definition $definition -VerifiedState $state | Should -BeFalse
            $definition.ImportedFrom.ContentVersion | Should -Be '2'
            $other = New-IpAppDefinition
            Update-IpImportedFrom -Definition $other -VerifiedState $state | Should -BeFalse
            $foreign = $definition.ImportedFrom | ConvertTo-Json | ConvertFrom-Json
            $foreign.AppId = $state.AppId
            Save-IpProfileImportedFrom -Path $profile.Path -ImportedFrom $foreign -Confirm:$false | Should -BeFalse
            (Import-IpAppProfile -Path $profile.Path).ImportedFrom.AppId | Should -Be 'c0ffee00-1111-4222-8333-000000000001'
        }
    }
}

Describe 'GUI update confirmation' {
    BeforeAll {
        $guiPath = Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'
        $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$null)
        foreach ($name in @('Get-IpUpdatePreview', 'Get-IpUpdateConfirmationText', 'Get-IpNextContentVersion', 'Get-IpConcurrencyQuestion', 'Confirm-IpUpdatePublication')) {
            $function = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
            if (-not $function) { throw "GUI function not found: $name" }
            . ([scriptblock]::Create($function.Extent.Text))
        }
        $script:MsiPackage = New-TestIntuneWin -Path (Join-Path $TestDrive 'gui\ContosoSetup.intunewin') -XmlText (Get-TestDetectionXml -WithMsi -SetupFile 'ContosoSetup.msi') -PayloadSize 1024
        $script:State = @{ ConnectedTenant = 'contoso.onmicrosoft.com' }
        function Show-IpMessage { param($Text, $Title, $Icon, $Buttons) $script:Shown.Add([pscustomobject]@{ Text = [string]$Text; Buttons = [string]$Buttons }); return $script:Answers.Dequeue() }
        function Start-IpPublishJob { param($Request) $script:Published.Add($Request) }
        function Invoke-IpImportResult { param($App) $script:Reimported++ }
        function Get-IpBaseDirectory { $script:GuiFolder }
        function Reset-GuiStubs { param([string[]]$Answers) $script:Shown = [Collections.Generic.List[object]]::new(); $script:Published = [Collections.Generic.List[object]]::new(); $script:Reimported = 0; $script:Answers = [Collections.Generic.Queue[string]]::new([string[]]$Answers) }
    }

    It 'warns when the new MSI ProductCode differs from the ProductCode rule in Intune' {
        $app = New-ContosoApp Msi
        $script:GuiFolder = Join-Path $TestDrive 'gui-code'
        $definition = New-ImportedDefinition -App $app -Folder $script:GuiFolder
        $app.rules[0].productCode = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}'
        $preview = Get-IpUpdatePreview -Request @{ Definition = $definition; PackagePath = $script:MsiPackage } -Current $app -BaseDirectory $script:GuiFolder
        ($preview.Warnings -join "`n") | Should -Match 'ProductCode \{11111111-2222-3333-4444-555555555555\} differs from the ProductCode detection rule in Intune \(\{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE\}\)'
        $preview.Concurrency | Should -Be ''
    }

    It 'lists the changes and warnings in the confirmation text' {
        $changes = @([pscustomobject]@{ Property = 'displayName'; Before = 'Contoso Sample App'; After = 'Contoso Sample App 3' })
        $text = Get-IpUpdateConfirmationText -Details ([pscustomobject]@{ displayName = 'Contoso Sample App'; id = 'c0ffee00-1111-4222-8333-000000000001'; committedContentVersion = '2' }) -Tenant 'contoso.onmicrosoft.com' -PackagePath 'C:\Out\ContosoSetup.intunewin' -Changes $changes -Warnings @('MSI warning') -Forced
        $text | Should -Match 'Committed content version: 2 -> 3'
        $text | Should -Match '- displayName: Contoso Sample App -> Contoso Sample App 3'
        $text | Should -Match '- MSI warning'
        $text | Should -Match 'Forced publication'
        Get-IpUpdateConfirmationText -Details ([pscustomobject]@{ displayName = 'A'; id = 'b'; committedContentVersion = '2' }) -Tenant 't' -PackagePath 'p.intunewin' -Changes @() | Should -Match 'No metadata change'
    }

    It 'blocks on a concurrent modification: re-import, force or cancel' {
        $app = New-ContosoApp Msi
        $script:GuiFolder = Join-Path $TestDrive 'gui-concurrency'
        $definition = New-ImportedDefinition -App $app -Folder $script:GuiFolder
        $app.lastModifiedDateTime = $script:Modified.AddMinutes(30)
        $request = @{ Definition = $definition; PackagePath = $script:MsiPackage }

        Reset-GuiStubs -Answers @('Yes')
        Confirm-IpUpdatePublication -Request $request -Current $app
        $script:Reimported | Should -Be 1
        $script:Published.Count | Should -Be 0
        $script:Shown[0].Buttons | Should -Be 'YesNoCancel'

        Reset-GuiStubs -Answers @('Cancel')
        Confirm-IpUpdatePublication -Request $request -Current $app
        $script:Published.Count | Should -Be 0

        Reset-GuiStubs -Answers @('No', 'Yes')
        Confirm-IpUpdatePublication -Request $request -Current $app
        $script:Published.Count | Should -Be 1
        $script:Published[0].Force | Should -BeTrue
        $script:Published[0].ExpectedLastModified | Should -Be '2026-10-18T10:01:00Z'
        $script:Shown[1].Text | Should -Match 'Forced publication'
    }

    It 'publishes without the concurrency question when the app is unchanged' {
        $app = New-ContosoApp Msi
        $script:GuiFolder = Join-Path $TestDrive 'gui-plain'
        $definition = New-ImportedDefinition -App $app -Folder $script:GuiFolder
        Reset-GuiStubs -Answers @('Yes')
        Confirm-IpUpdatePublication -Request @{ Definition = $definition; PackagePath = $script:MsiPackage } -Current $app
        $script:Shown.Count | Should -Be 1
        $script:Published[0].Force | Should -BeFalse
    }
}

Describe 'CLI publication of a given package' {
    BeforeAll {
        $guiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'), [ref]$null, [ref]$null)
        foreach ($name in @('Resolve-IpCliOutputFolder', 'Get-IpProfileOutputFolder')) {
            $function = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
            . ([scriptblock]::Create($function.Extent.Text))
        }
        $script:EmptyConfig = [pscustomobject]@{ Paths = [pscustomobject]@{ DefaultOutputFolder = '' } }
    }

    It 'needs no output or source folder when -IntuneWinPath is given' {
        $definition = (ConvertFrom-IpWin32App -App (New-ContosoApp Msi)).Definition
        $definition.Package.SourceFolder | Should -Be ''
        $definition.Package.OutputFolder | Should -Be ''
        Test-IpAppDefinition -Definition $definition -ForPublish | Should -BeNullOrEmpty
        Resolve-IpCliOutputFolder -Definition $definition -Config $script:EmptyConfig -BaseDirectory $TestDrive -NeedsBuild $false -IntuneWinPath 'C:\Packages\ContosoSetup.intunewin' | Should -Be ''
    }

    It 'still requires an output folder to build or to find the package' {
        $definition = (ConvertFrom-IpWin32App -App (New-ContosoApp Msi)).Definition
        { Resolve-IpCliOutputFolder -Definition $definition -Config $script:EmptyConfig -BaseDirectory $TestDrive -NeedsBuild $false -IntuneWinPath '' } | Should -Throw '*No output folder*'
        { Resolve-IpCliOutputFolder -Definition $definition -Config $script:EmptyConfig -BaseDirectory $TestDrive -NeedsBuild $true -IntuneWinPath 'C:\Packages\ContosoSetup.intunewin' } | Should -Throw '*No output folder*'
    }
}

Describe 'CLI concurrency block message' {
    BeforeAll {
        $guiAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.ps1'), [ref]$null, [ref]$null)
        $function = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-IpConcurrencyCommands' }, $true)
        . ([scriptblock]::Create($function.Extent.Text))
        $script:CliText = $guiAst.Extent.Text
    }

    It 'gives the exact re-import and -Force command lines' {
        $commands = Get-IpConcurrencyCommands -ProfilePath 'C:\Profiles\Contoso.profile.json' -AppId 'c0ffee00-1111-4222-8333-000000000001' -PackagePath 'C:\Packages\ContosoSetup.intunewin' -ScriptPath 'C:\Tools\IntunePackager\SmartM365-IntunePackager-GUI.ps1' -TenantName 'test' -PublishAction 'Publish'
        $commands.Reimport | Should -Be 'pwsh -NoProfile -ExecutionPolicy Bypass -File "C:\Tools\IntunePackager\SmartM365-IntunePackager-GUI.ps1" -Tenant test -Action Import -ExistingAppId c0ffee00-1111-4222-8333-000000000001 -ProfilePath "C:\Profiles\Contoso.profile.json" -Force'
        $commands.Force | Should -Be 'pwsh -NoProfile -ExecutionPolicy Bypass -File "C:\Tools\IntunePackager\SmartM365-IntunePackager-GUI.ps1" -Tenant test -ProfilePath "C:\Profiles\Contoso.profile.json" -Action Publish -PublishMode UpdateExisting -ExistingAppId c0ffee00-1111-4222-8333-000000000001 -IntuneWinPath "C:\Packages\ContosoSetup.intunewin" -Force'
        (Get-IpConcurrencyCommands -ProfilePath 'C:\P.json' -AppId 'a' -ScriptPath 's.ps1' -TenantName 't' -PublishAction 'BuildAndPublish').Force | Should -Not -Match 'IntuneWinPath'
    }

    It 'logs the block once in the CLI' {
        $script:CliText | Should -Match "if \(-not \`$_\.Exception\.Data\['IpLogged'\]\) \{ Write-IpErrorDetail"
        $script:CliText | Should -Match "\`$_\.Exception\.Data\['IpLogged'\] = \`$true"
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfdVY7DifibpRJ
# EIYp1D/gBaXqf4puy+IsnT+FEoLcyKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBlWddtt8fVpLnp9NsrLdHIl6/grTB9jXgh5Gz9BL+9qMA0GCSqG
# SIb3DQEBAQUABIIBgD20CzcIMtu/UKdRlQ8AiVdwkXIlrxgKjg14FI8EcIMYoS0T
# ZPl4+vewbzNp+cMV5SGELZyaBi4Sws0xFmdvyiofLxdaB7WjMjAcJPIYYn+QwtSo
# jUTdGh9Ci45mPgW7xs439F1rUisj7E/rzyxB07FuocV0yKo05IK8dTzEsXu8AXnA
# /+EPzayHXqRglbh6vo1tcE3eMMqZFqOa3bQWAdA+NF/5gSM+kz+6mG6QQA803aEJ
# Jmk8/3po0txgDY8HxuKVTCSvh/QabrSr9LDvjGPnpfbaQd5MC8v7QUUYpstmJYQ/
# t9gx2N530EZIVRZbVIN74ke7LNR9PzXmjeI4jIfHzYWsUQIOBqmuQmFg00UcmW+m
# DEfedxE5nDqKqQtChVja9lsPqyKhZlEOkXutzSLIYPi01f3nVlTXHFltbIvqn6+3
# FaSP3WV+ZxeNms3vonz9JAPgCfxgsCftH3Xwt7s3pBATsZpwNa+7V2fowgmVDLN7
# Jc+4GjnxwP4bsgNAS6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzhaMC8GCSqGSIb3DQEJBDEiBCAGg9PaLCY38IXuDRVqf75DvmnteJC+BXEXZO11
# 7TFUsjANBgkqhkiG9w0BAQEFAASCAgBnwCIana1T2YJjHBc4nyQM46WpusuRjcNr
# oUnifAmOff/oAu/hs2WacxzLgpRfQtTEzcnD/SJM8ldRHMsRVW3QzlS7iFlE1wsV
# w3xi9yxFJvtHOs3XsrhdH+c+fjCxFyBZ736sDGXdn2I3kMZuBR2YpHLSib+hODHE
# 6YfGm6/vgBHQCt8k1fZCNy9ObIHCTXkg1y9lC5GxPkrrxi6PdxWGt8fbsLYyKrbq
# 5EC3ncE/aJdMbRR65xPwYStv8f04EIrWrbJqUzOnTkPCgTHrEraIfRhXfOVltcU9
# k+yawu9nWXVhZawwxIPuzu49LnaH9B+zsnQ+kkF26e8lfu+IycKIE3HpDZHBokk/
# RUnyiSlKned3o3Fmhdmr4oAhVfMRFXJhqxEvTHTHNXXP6+WlmiKuRSUbie5RVZWf
# i3KDjQrfLyuGPp8EPpfHRMN5/B1gJIT4cG/jlTgOnN1nGOqhb4tWX86Swi3L0N+7
# ejzB1Sa/qFjgx/OCYrZHwnpwWiq3ceKgBVDQH6f0pFr1nBRfN3y1Fhl2rjcln0/m
# Aaraewie/4oU5hpFMsgcIRyuyaUl4/jJ2MzFBpPY/ZilIZtrR1g1Qo7f5fs1H4dN
# oMX/m7Z/y7YvEj3ywnzMMj7cue759rU2cWU59N7E+6kk3dxA+zz+Y2EUAQWo3I1p
# ihq6psDe/Q==
# SIG # End signature block
