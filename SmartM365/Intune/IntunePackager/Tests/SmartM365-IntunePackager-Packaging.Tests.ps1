#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
.SYNOPSIS
    Pester 5 tests for SmartM365.IntunePackager.Packaging.psm1 (offline, synthetic data).

.DESCRIPTION
    The optional integration test runs the real Win32 Content Prep Tool only when the
    environment variable IP_TEST_CONTENT_PREP_TOOL points to a verified IntuneWinAppUtil.exe.

.VERSION
    1.0.0
#>

BeforeDiscovery {
    $script:RealToolPath = [string]$env:IP_TEST_CONTENT_PREP_TOOL
}

BeforeAll {
    $script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Core.psm1') -Force
    Import-Module (Join-Path $script:AppRoot 'SmartM365.IntunePackager.Packaging.psm1') -Force
    . (Join-Path $PSScriptRoot 'SmartM365-IntunePackager-TestHelpers.ps1')
    Set-IpLogSink -Sink $null -Console $false
    Set-IpLogPath -Path ''
    $script:Template = Read-IpJsonFile -Path (Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.config.json.template')
}

Describe 'Engine registry' {
    It 'registers the IntuneWinAppUtil engine' {
        (Get-IpPackagingEngine -Name 'IntuneWinAppUtil').Name | Should -Be 'IntuneWinAppUtil'
    }

    It 'fails clearly for an unknown engine' {
        { Get-IpPackagingEngine -Name 'Missing' } | Should -Throw "*Unknown packaging engine 'Missing'*"
    }

    It 'reports missing tool settings as engine issues' {
        $engine = Get-IpPackagingEngine -Name 'IntuneWinAppUtil'
        $settings = '{"ContentPrepTool":{"Version":"1.8.7","DownloadUrl":"","Sha256":"","StoragePath":"x"}}' | ConvertFrom-Json
        $issues = @(& $engine.Test @{ Settings = $settings })
        $issues | Should -Contain 'Configuration error: ContentPrepTool.DownloadUrl is required.'
        $issues | Should -Contain 'Configuration error: ContentPrepTool.Sha256 is required.'
    }
}

Describe 'Build request validation' {
    BeforeAll {
        $script:Source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'Source')
        Set-Content -LiteralPath (Join-Path $script:Source 'setup.cmd') -Value 'exit 0'
    }

    It 'accepts a valid request' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -BeNullOrEmpty
    }

    It 'rejects a missing source folder' {
        Test-IpPackageBuildRequest -SourceFolder (Join-Path $TestDrive 'None') -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'Source folder not found'
    }

    It 'rejects a setup file outside the source folder' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile '..\setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'relative path inside the source folder'
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'C:\setup.cmd' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'relative path inside the source folder'
    }

    It 'rejects a missing setup file' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'missing.exe' -OutputFolder (Join-Path $TestDrive 'Out') | Should -Match 'Setup file not found'
    }

    It 'rejects an output folder inside the source folder' {
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder (Join-Path $script:Source 'Out') | Should -Match 'must not be the source folder'
        Test-IpPackageBuildRequest -SourceFolder $script:Source -SetupFile 'setup.cmd' -OutputFolder $script:Source | Should -Match 'must not be the source folder'
    }

    It 'builds the documented IntuneWinAppUtil arguments' {
        Get-IpIntuneWinAppUtilArguments -SourceFolder 'C:\Src' -SetupFile 'setup.msi' -OutputFolder 'C:\Out' | Should -Be @('-c', 'C:\Src', '-s', 'setup.msi', '-o', 'C:\Out', '-q')
    }
}

Describe 'Content Prep Tool verification' {
    It 'stores the tool per version under the configured path' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = '%TEMP%\IpTools' }
        Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive | Should -Be ([IO.Path]::GetFullPath((Join-Path $env:TEMP 'IpTools\1.8.7\IntuneWinAppUtil.exe')))
    }

    It 'rejects a missing file' {
        $result = Test-IpContentPrepTool -Path (Join-Path $TestDrive 'none.exe') -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Match 'File not found'
    }

    It 'rejects an unsigned file with a different hash' {
        $path = Join-Path $TestDrive 'IntuneWinAppUtil.exe'
        [IO.File]::WriteAllBytes($path, [byte[]](1..64))
        $result = Test-IpContentPrepTool -Path $path -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $result.IsValid | Should -BeFalse
        ($result.Reasons -join ' ') | Should -Match 'SHA256 mismatch'
        ($result.Reasons -join ' ') | Should -Match 'Authenticode signature is not valid'
    }

    It 'refuses a download URL outside the official Microsoft repository' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools1'); DownloadUrl = 'https://downloads.contoso.example/IntuneWinAppUtil.exe'; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { throw 'Download must not be attempted.' }
        { Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false } | Should -Throw '*official Microsoft GitHub repository*'
        Should -Invoke -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest -Times 0
    }

    It 'reuses a verified tool without downloading' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools2'); DownloadUrl = $script:Template.ContentPrepTool.DownloadUrl; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        $toolPath = Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive
        New-Item -ItemType Directory -Path (Split-Path $toolPath) -Force | Out-Null
        Set-Content -LiteralPath $toolPath -Value 'placeholder'
        Mock -ModuleName SmartM365.IntunePackager.Packaging Test-IpContentPrepTool { [pscustomobject]@{ IsValid = $true; Reasons = @() } }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { throw 'Download must not be attempted.' }
        Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false | Should -Be $toolPath
        Should -Invoke -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest -Times 0
    }

    It 'deletes a downloaded file that fails verification' {
        $settings = [pscustomobject]@{ Version = '1.8.7'; StoragePath = (Join-Path $TestDrive 'Tools3'); DownloadUrl = $script:Template.ContentPrepTool.DownloadUrl; Sha256 = $script:Template.ContentPrepTool.Sha256 }
        Mock -ModuleName SmartM365.IntunePackager.Packaging Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](1..16)) }
        { Install-IpContentPrepTool -ToolSettings $settings -AppRoot $TestDrive -Confirm:$false } | Should -Throw '*failed verification*'
        $toolPath = Get-IpContentPrepToolPath -ToolSettings $settings -AppRoot $TestDrive
        Test-Path -LiteralPath $toolPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Split-Path $toolPath) -Force).Count | Should -Be 0
    }
}

Describe 'Package build orchestration' {
    BeforeAll {
        Register-IpPackagingEngine -Name 'TestEngine' -Description 'Synthetic engine for tests.' -Test { param($Request) @() } -Build {
            param($Request)
            $path = Join-Path $Request.OutputFolder ([IO.Path]::GetFileNameWithoutExtension($Request.SetupFile) + '.intunewin')
            New-TestIntuneWin -Path $path -XmlText (Get-TestDetectionXml -SetupFile $Request.SetupFile) -PayloadSize 1024 | Out-Null
            [pscustomobject]@{ IntuneWinPath = $path; EngineVersion = 'test' }
        }
        $script:BuildSource = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'BuildSource')
        Set-Content -LiteralPath (Join-Path $script:BuildSource 'setup.cmd') -Value 'exit 0'
    }

    AfterAll { Unregister-IpPackagingEngine -Name 'TestEngine' }

    It 'returns the package path, hash and metadata' {
        $result = Invoke-IpPackageBuild -EngineName 'TestEngine' -SourceFolder $script:BuildSource -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'BuildOut') -Settings $script:Template -AppRoot $script:AppRoot
        $result.IntuneWinPath | Should -Be (Join-Path $TestDrive 'BuildOut\setup.intunewin')
        $result.Sha256 | Should -Match '^[0-9A-F]{64}$'
        $result.Metadata.SetupFile | Should -Be 'setup.cmd'
        $result.Metadata.SizeEncrypted | Should -Be 1024
        $result.Engine | Should -Be 'TestEngine'
    }

    It 'stops before building when the request is invalid' {
        { Invoke-IpPackageBuild -EngineName 'TestEngine' -SourceFolder $script:BuildSource -SetupFile 'missing.exe' -OutputFolder (Join-Path $TestDrive 'BuildOut') -Settings $script:Template -AppRoot $script:AppRoot } | Should -Throw '*Setup file not found*'
    }

    It 'fails when the engine does not produce the package' {
        Register-IpPackagingEngine -Name 'BrokenEngine' -Description 'Produces nothing.' -Test { param($Request) @() } -Build { param($Request) [pscustomobject]@{ IntuneWinPath = (Join-Path $Request.OutputFolder 'none.intunewin'); EngineVersion = 'x' } }
        try {
            { Invoke-IpPackageBuild -EngineName 'BrokenEngine' -SourceFolder $script:BuildSource -SetupFile 'setup.cmd' -OutputFolder (Join-Path $TestDrive 'BrokenOut') -Settings $script:Template -AppRoot $script:AppRoot } | Should -Throw '*did not produce the expected file*'
        }
        finally { Unregister-IpPackagingEngine -Name 'BrokenEngine' }
    }
}

Describe 'IntuneWinAppUtil integration' -Tag 'Integration' -Skip:([string]::IsNullOrWhiteSpace($script:RealToolPath)) {
    It 'builds a real package from a small folder' {
        $check = Test-IpContentPrepTool -Path $env:IP_TEST_CONTENT_PREP_TOOL -ExpectedSha256 $script:Template.ContentPrepTool.Sha256
        $check.IsValid | Should -BeTrue
        $source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'RealSource')
        Set-Content -LiteralPath (Join-Path $source 'install.cmd') -Value 'exit 0'
        $output = Join-Path $TestDrive 'RealOut'
        New-Item -ItemType Directory -Path $output | Out-Null
        $path = Invoke-IpIntuneWinAppUtil -ToolPath $env:IP_TEST_CONTENT_PREP_TOOL -SourceFolder $source -SetupFile 'install.cmd' -OutputFolder $output -TimeoutMinutes 5
        (Read-IpIntuneWinMetadata -Path $path).SetupFile | Should -Be 'install.cmd'
    }
}
