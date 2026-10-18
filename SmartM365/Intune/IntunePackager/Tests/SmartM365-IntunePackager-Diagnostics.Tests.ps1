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
