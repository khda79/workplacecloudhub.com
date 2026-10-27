#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests for the Device Reboot Manager PowerShell Gallery module, build, update helper and publish preview.
.DESCRIPTION
Builds the Gallery package into a temporary folder and validates its content,
file catalog, packaged demo, exported commands, the update-task argument string
(replayed with powershell.exe -File against a stub that has the helper's exact
param block), catalog tamper detection, VM-validation preview and the publish
guards. It never installs the product, registers tasks or publishes.
Use -SkipAuthenticodeValidation when the edited sources are not signed yet.
.VERSION
1.1.0
#>

[CmdletBinding()]
param(
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }
}

function Assert-Throws {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $true)][string]$ExpectedMessageLike,
        [Parameter(Mandatory = $true)][string]$Message
    )
    $thrown = $false
    try {
        & $ScriptBlock | Out-Null
    }
    catch {
        $thrown = $true
        if ($_.Exception.Message -notlike $ExpectedMessageLike) {
            throw ("Assertion failed: {0}. Unexpected error: {1}" -f $Message,$_.Exception.Message)
        }
    }
    Assert-True $thrown $Message
}

function Assert-PackagedDemo {
    param([string]$RuntimeRoot)
    $entry=(Join-Path $RuntimeRoot 'SmartM365-DeviceRebootManager-GUI.ps1').Replace("'","''")
    $command='$ErrorActionPreference=''Stop''; $state=. '''+$entry+''' -Demo -ValidateOnly; if($state.Mode -ne ''Demo'' -or $null -eq $state.Window -or $state.Controls.Count -eq 0){throw ''Packaged demo did not return a window and controls''}; $state.Window.Close(); Write-Output ''DEMO-PASS'''
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $output=& powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and $output -contains 'DEMO-PASS') ('Packaged demo failed: '+($output -join ' | '))
}

function Import-HelperDefinition {
    # Loads selected functions and top-level variables from the update helper
    # without running its elevated deployment body.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$FunctionNames,
        [string[]]$VariableNames = @()
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "Update helper does not parse: $Path"
    $definitions = New-Object Collections.Generic.List[string]
    foreach ($name in $FunctionNames) {
        $functionAst = @($ast.FindAll({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
        }, $false)) | Select-Object -First 1
        Assert-True ($null -ne $functionAst) "Update helper function not found: $name"
        $definitions.Add($functionAst.Extent.Text)
    }
    foreach ($name in $VariableNames) {
        $assignment = @($ast.EndBlock.Statements | Where-Object {
            $_ -is [Management.Automation.Language.AssignmentStatementAst] -and
                $_.Left.Extent.Text -eq ('$script:{0}' -f $name)
        }) | Select-Object -First 1
        Assert-True ($null -ne $assignment) "Update helper variable not found: $name"
        $definitions.Add($assignment.Extent.Text)
    }
    return ($definitions -join "`r`n")
}

function New-BindingStub {
    param(
        [Parameter(Mandatory = $true)][string]$HelperPath,
        [Parameter(Mandatory = $true)][string]$StubPath
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($HelperPath, [ref]$tokens, [ref]$parseErrors)
    $stubLines = New-Object Collections.Generic.List[string]
    foreach ($attribute in @($ast.ParamBlock.Attributes)) {
        $stubLines.Add($attribute.Extent.Text)
    }
    $stubLines.Add($ast.ParamBlock.Extent.Text)
    $stubLines.Add('$bound = [ordered]@{}')
    $stubLines.Add('foreach ($key in @($PSBoundParameters.Keys)) {')
    $stubLines.Add('    $value = $PSBoundParameters[$key]')
    $stubLines.Add('    if ($value -is [Management.Automation.SwitchParameter]) { $value = $value.IsPresent }')
    $stubLines.Add('    $bound[$key] = $value')
    $stubLines.Add('}')
    $stubLines.Add('''BINDING='' + ($bound | ConvertTo-Json -Compress)')
    Set-Content -LiteralPath $StubPath -Value $stubLines.ToArray() -Encoding UTF8
}

function Invoke-BindingProbe {
    # Runs the exact argument string, as Task Scheduler does, and returns the
    # parameters bound by the stub.
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$ArgumentString
    )

    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Executable
    $startInfo.Arguments = $ArgumentString
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($startInfo)
    $standardOutput = $process.StandardOutput.ReadToEnd()
    $standardError = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    $bindingLine = @($standardOutput -split "`r?`n" | Where-Object { $_ -like 'BINDING=*' }) | Select-Object -First 1
    Assert-True ($process.ExitCode -eq 0 -and $null -ne $bindingLine) `
        ("Task arguments did not bind under {0}: exit={1}; stderr={2}; arguments={3}" -f
            $Executable,$process.ExitCode,$standardError.Trim(),$ArgumentString)
    return ($bindingLine.Substring(8) | ConvertFrom-Json)
}

function Assert-Binding {
    param(
        [Parameter(Mandatory = $true)][object]$Bound,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$Expected,
        [Parameter(Mandatory = $true)][string]$Context
    )
    foreach ($name in $Expected.Keys) {
        $property = $Bound.PSObject.Properties[$name]
        $actualValue = if ($property) { $property.Value } else { $null }
        Assert-True ([string]$actualValue -ceq [string]$Expected[$name]) `
            ("{0}: parameter {1} bound to '{2}' instead of '{3}'" -f $Context,$name,$actualValue,$Expected[$name])
    }
}

$productRoot = Split-Path -Path $PSScriptRoot -Parent
$galleryRoot = Join-Path -Path $productRoot -ChildPath 'PowerShellGallery'
$moduleSourceRoot = Join-Path -Path $galleryRoot -ChildPath 'Module'
$testRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-GalleryTest-{0}" -f [guid]::NewGuid().ToString('N'))
$expectedThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114'

try {
    $sourcePowerShellFiles = @(Get-ChildItem -LiteralPath $galleryRoot -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    Assert-True ($sourcePowerShellFiles.Count -ge 5) 'Expected Gallery source files were not found.'

    foreach ($file in $sourcePowerShellFiles) {
        $parseErrors = $null
        [void][Management.Automation.PSParser]::Tokenize(
            (Get-Content -LiteralPath $file.FullName -Raw),
            [ref]$parseErrors
        )
        Assert-True ($parseErrors.Count -eq 0) ("PowerShell syntax errors found in {0}" -f $file.FullName)
    }

    $moduleSourcePath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psm1'
    $moduleSourceText = Get-Content -LiteralPath $moduleSourcePath -Raw
    $updateHelperPath = Join-Path -Path $moduleSourceRoot -ChildPath 'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    $updateHelperText = Get-Content -LiteralPath $updateHelperPath -Raw
    Assert-True ($moduleSourceText -match '\[bool\]\$EnableAutomaticUpdate\s*=\s*\$false') 'Installation must disable automatic updates by default.'
    Assert-True ($updateHelperText -match '\[switch\]\$EnableAutomaticUpdate\s*,') 'Update helper must expose automatic updates as an opt-in switch.'
    Assert-True ($moduleSourceText -match '\[Nullable\[bool\]\]\$EnableAutomaticUpdate\s*=\s*\$null') `
        'Manual update must preserve the existing automatic-update state by default.'
    Assert-True ($moduleSourceText -match '\[bool\]\$IncludePrerelease\s*=\s*\$false') `
        'Stable module must not select prerelease updates by default.'
    Assert-True ($updateHelperText -match '\[switch\]\$IncludePrerelease\s*,') `
        'Stable update helper must expose prerelease selection as an opt-in switch.'
    Assert-True ($moduleSourceText -match '-ModuleVersion \$currentModuleVersion') `
        'Install must deploy the running module version.'
    foreach ($text in @($moduleSourceText, $updateHelperText)) {
        Assert-True ($text -notmatch 'ProgramData') 'Module and helper must not default to the legacy ProgramData path.'
        Assert-True ($text.Contains("-ChildPath 'WCH\SmartDeviceRebootManager'")) 'Module and helper must default to Program Files\WCH\SmartDeviceRebootManager.'
    }

    $scheduledTaskScriptPath = Join-Path -Path $productRoot -ChildPath 'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    $scheduledTaskScriptText = Get-Content -LiteralPath $scheduledTaskScriptPath -Raw
    Assert-True ($scheduledTaskScriptText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Main scheduled-task script must not mutate a missing Repetition property.'
    Assert-True ($updateHelperText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Gallery update helper must not mutate a missing Repetition property.'
    Assert-True ($scheduledTaskScriptText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Main repeat trigger must use the compatible Once parameter set.'
    Assert-True ($scheduledTaskScriptText -notmatch 'LeastPrivilege') `
        'Main scheduled-task script must use a valid RunLevel enumerator.'
    Assert-True ($scheduledTaskScriptText -notmatch '-GroupId[^\r\n]+-LogonType') `
        'Group principal must not mix incompatible GroupId and LogonType parameter sets.'
    Assert-True ($scheduledTaskScriptText -match "-GroupId 'S-1-5-32-545' -RunLevel Limited") `
        'Main task principal must use the Users group with Limited run level.'
    Assert-True ($updateHelperText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Gallery update trigger must use the compatible Once parameter set.'

    $versionManifestPath = Join-Path -Path $productRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
    $versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
    Assert-True ([int]$versionManifest.SchemaVersion -eq 1) 'Unexpected package version schema.'
    $expectedPackageVersion = [string]$versionManifest.PackageVersion
    Assert-True (-not [string]::IsNullOrWhiteSpace($expectedPackageVersion)) 'Version manifest has no PackageVersion.'

    $sourceManifestPath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
    $sourceManifest = Test-ModuleManifest -Path $sourceManifestPath
    Assert-True ($sourceManifest.Name -eq 'SmartM365.DeviceRebootManager') 'Unexpected module name.'
    $sourcePrerelease = [string]$sourceManifest.PrivateData.PSData['Prerelease']
    Assert-True ([string]::IsNullOrWhiteSpace($sourcePrerelease)) 'Stable manifest must not contain a prerelease label.'
    Assert-True ([string]$sourceManifest.Version -eq $expectedPackageVersion) 'Version manifest and stable module version differ.'
    $moduleHeaderVersion = [regex]::Match($moduleSourceText, '(?m)^\.VERSION\s*\r?\n\s*(\S+)').Groups[1].Value
    Assert-True ($moduleHeaderVersion -eq [string]$sourceManifest.Version) 'Module header .VERSION differs from the manifest ModuleVersion.'

    # Update-task argument string: built by the helper's own functions, then
    # replayed with powershell.exe -File (and pwsh when present) against a stub
    # carrying the helper's exact param block.
    . ([scriptblock]::Create((Import-HelperDefinition `
        -Path $updateHelperPath `
        -FunctionNames @('ConvertTo-ProcessArgument', 'Get-GalleryUpdateTaskArgumentString'))))

    $stubRoot = Join-Path -Path $testRoot -ChildPath 'Binding Stub Dir'
    New-Item -ItemType Directory -Path $stubRoot -Force | Out-Null
    $stubPath = Join-Path -Path $stubRoot -ChildPath 'Gallery Update Stub.ps1'
    New-BindingStub -HelperPath $updateHelperPath -StubPath $stubPath

    $bindingCases = @(
        [ordered]@{
            ModuleName               = 'SmartM365.DeviceRebootManager'
            InstallPath              = 'C:\Program Files\WCH\SmartDeviceRebootManager'
            TaskPath                 = '\SmartM365\'
            TaskName                 = 'Device Reboot Manager'
            RepeatIntervalMinutes    = 240
            EnableAutomaticUpdate    = $true
            UpdateTaskPath           = '\SmartM365\'
            UpdateTaskName           = 'Device Reboot Manager Update'
            UpdateIntervalHours      = 24
            IncludePrerelease        = $null
            ExpectedSignerThumbprint = $expectedThumbprint
        }
        [ordered]@{
            ModuleName               = 'SmartM365.DeviceRebootManager'
            InstallPath              = 'C:\Program Files\WCH\Smart Device Reboot Manager\'
            TaskPath                 = '\Smart M365\Pilot "Ring" 1\'
            TaskName                 = 'Device Reboot Manager "Pilot"'
            RepeatIntervalMinutes    = 120
            EnableAutomaticUpdate    = $true
            UpdateTaskPath           = '\Smart M365\Updates\\'
            UpdateTaskName           = 'Device Reboot Manager Update'
            UpdateIntervalHours      = 12
            IncludePrerelease        = $true
            ExpectedSignerThumbprint = $expectedThumbprint
        }
    )
    $bindingExecutables = @('powershell.exe')
    if (Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue) {
        $bindingExecutables += 'pwsh.exe'
    }
    $bindingRuns = 0
    foreach ($case in $bindingCases) {
        $argumentString = Get-GalleryUpdateTaskArgumentString `
            -ScriptPath $stubPath `
            -TargetModuleName $case.ModuleName `
            -TargetInstallPath $case.InstallPath `
            -GuiTaskPath $case.TaskPath `
            -GuiTaskName $case.TaskName `
            -GuiRepeatIntervalMinutes $case.RepeatIntervalMinutes `
            -ScheduledTaskPath $case.UpdateTaskPath `
            -ScheduledTaskName $case.UpdateTaskName `
            -IntervalHours $case.UpdateIntervalHours `
            -Prerelease ([bool]$case.IncludePrerelease) `
            -SignerThumbprint $case.ExpectedSignerThumbprint
        Assert-True ($argumentString -notmatch '\$true|\$false') 'Task arguments must not pass boolean literals.'
        foreach ($executable in $bindingExecutables) {
            $bound = Invoke-BindingProbe -Executable $executable -ArgumentString $argumentString
            Assert-Binding -Bound $bound -Expected $case -Context ("{0} {1}" -f $executable,$case.TaskPath)
            $bindingRuns++
        }
    }

    $buildScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
    $buildResult = & $buildScript `
        -OutputRoot $testRoot `
        -Force `
        -TimestampServer '' `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $buildResult.Ready 'Build did not report Ready.'
    Assert-True ($buildResult.PackageVersion -eq $expectedPackageVersion) 'Stable build package version is incorrect.'
    Assert-True ($buildResult.IntuneSourcePath -eq (Join-Path $buildResult.PackagePath 'Runtime')) `
        'Intune source path must be the packaged Runtime folder.'
    Assert-True (Test-Path -LiteralPath $buildResult.IntuneSourcePath -PathType Container) 'Intune source path was not created.'
    Assert-True (Test-Path -LiteralPath $buildResult.CatalogPath -PathType Leaf) 'Package file catalog was not created.'
    Assert-True ($buildResult.CatalogSigned -eq (-not $SkipAuthenticodeValidation)) 'Catalog signing state is incorrect.'

    Assert-PackagedDemo -RuntimeRoot $buildResult.IntuneSourcePath

    $requiredPackagePaths = @(
        'SmartM365.DeviceRebootManager.psd1'
        'SmartM365.DeviceRebootManager.psm1'
        'SmartM365.DeviceRebootManager.Files.cat'
        'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        'Runtime\SmartM365-DeviceRebootManager.version.json.txt'
        'Runtime\SmartM365-DeviceRebootManager-GUI.ps1'
        'Runtime\SmartM365-DeviceRebootManager-Demo.ps1'
        'Runtime\Start-SmartM365-DeviceRebootManager-Demo.cmd'
        'Runtime\SmartM365-DeviceRebootManager-GUI.config.json.template'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
    )
    foreach ($relativePath in $requiredPackagePaths) {
        Assert-True (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath $relativePath) -PathType Leaf) `
            ("Required package file is missing: {0}" -f $relativePath)
    }

    $forbiddenNames = @(
        'LOCAL_MEMORY.md'
        'SmartM365-DeviceRebootManager-GUI.config.json'
    )
    $packagedFiles = @(Get-ChildItem -LiteralPath $buildResult.PackagePath -Recurse -File)
    foreach ($forbiddenName in $forbiddenNames) {
        Assert-True (-not ($packagedFiles.Name -contains $forbiddenName)) ("Private file was packaged: {0}" -f $forbiddenName)
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath 'Logs'))) 'Logs directory was packaged.'

    # Catalog tamper detection, using the update helper's own integrity check.
    . ([scriptblock]::Create((Import-HelperDefinition `
        -Path $updateHelperPath `
        -FunctionNames @(
            'Get-NormalizedThumbprint'
            'Get-WinVerifyTrustResult'
            'Get-PinnedSignatureState'
            'Assert-PinnedSignature'
            'Get-FileCatalogDifference'
            'Assert-ModulePackageIntegrity'
        ) `
        -VariableNames @('PackageCatalogFileName', 'CatalogFilesToSkip'))))
    $script:UpdateLogLines = New-Object Collections.Generic.List[string]
    function Write-UpdateLog {
        param([string]$Message, [string]$Level = 'INFO')
        $script:UpdateLogLines.Add(("[{0}] {1}" -f $Level,$Message))
    }

    # WinVerifyTrust signer rule, exercised with an in-memory self-signed
    # code-signing certificate whose root is not trusted anywhere: an intact
    # signature is accepted with a warning; a one-byte body change, an
    # unsigned file and a different pinned thumbprint are rejected.
    $signerRoot = Join-Path -Path $testRoot -ChildPath 'SignerRule'
    New-Item -ItemType Directory -Path $signerRoot -Force | Out-Null
    $ephemeralKey = [Security.Cryptography.RSA]::Create(2048)
    $ephemeralRequest = New-Object Security.Cryptography.X509Certificates.CertificateRequest(
        'CN=SmartM365 DeviceRebootManager Ephemeral Test Signer',
        $ephemeralKey,
        [Security.Cryptography.HashAlgorithmName]::SHA256,
        [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $codeSigningUsage = New-Object Security.Cryptography.OidCollection
    [void]$codeSigningUsage.Add((New-Object Security.Cryptography.Oid('1.3.6.1.5.5.7.3.3')))
    $ephemeralRequest.CertificateExtensions.Add(
        (New-Object Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension($codeSigningUsage, $false)))
    $ephemeralCertificate = $ephemeralRequest.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddDays(1))
    $ephemeralPfx = $ephemeralCertificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, 'ephemeral')
    $testSigner = New-Object Security.Cryptography.X509Certificates.X509Certificate2(
        $ephemeralPfx,
        'ephemeral',
        [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
    try {
        $signedScriptPath = Join-Path -Path $signerRoot -ChildPath 'Signed.ps1'
        Set-Content -LiteralPath $signedScriptPath -Value "Write-Output 'signed body'" -Encoding UTF8
        [void](Set-AuthenticodeSignature -LiteralPath $signedScriptPath -Certificate $testSigner -HashAlgorithm SHA256)
        $signedCatalogRoot = Join-Path -Path $signerRoot -ChildPath 'Catalog'
        New-Item -ItemType Directory -Path $signedCatalogRoot -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $signedCatalogRoot 'Launcher.cmd') -Value '@echo off' -Encoding ASCII
        $signedCatalogPath = Join-Path -Path $signedCatalogRoot -ChildPath 'Test.cat'
        New-FileCatalog -Path $signedCatalogRoot -CatalogFilePath $signedCatalogPath -CatalogVersion 2.0 | Out-Null
        [void](Set-AuthenticodeSignature -LiteralPath $signedCatalogPath -Certificate $testSigner -HashAlgorithm SHA256)

        $tamperedScriptPath = Join-Path -Path $signerRoot -ChildPath 'Tampered.ps1'
        $signedBytes = [IO.File]::ReadAllBytes($signedScriptPath)
        $bodyOffset = [Text.Encoding]::ASCII.GetString($signedBytes).IndexOf('signed body')
        Assert-True ($bodyOffset -gt 0) 'Signed test body was not found.'
        $signedBytes[$bodyOffset] = [byte][char]'S'
        [IO.File]::WriteAllBytes($tamperedScriptPath, $signedBytes)
        $unsignedScriptPath = Join-Path -Path $signerRoot -ChildPath 'Unsigned.ps1'
        Set-Content -LiteralPath $unsignedScriptPath -Value 'Write-Output 1' -Encoding UTF8

        $intactState = Get-PinnedSignatureState -Path $signedScriptPath -Thumbprint $testSigner.Thumbprint
        Assert-True ($intactState.Accepted -and -not $intactState.ChainTrusted) `
            ("Intact signature with an untrusted root must be accepted: {0}" -f $intactState.TrustResult)
        $script:UpdateLogLines.Clear()
        Assert-PinnedSignature -Path $signedScriptPath -Thumbprint $testSigner.Thumbprint
        Assert-True (@($script:UpdateLogLines | Where-Object { $_ -like '`[WARN`]*not trusted*' }).Count -eq 1) `
            'An untrusted-chain acceptance must be logged as a warning.'
        $catalogState = Get-PinnedSignatureState -Path $signedCatalogPath -Thumbprint $testSigner.Thumbprint
        Assert-True $catalogState.Accepted ("Intact signed catalog must be accepted: {0}" -f $catalogState.TrustResult)
        foreach ($rejectedCase in @(
            @{ Path = $tamperedScriptPath; Thumbprint = $testSigner.Thumbprint; Name = 'tampered body' }
            @{ Path = $unsignedScriptPath; Thumbprint = $testSigner.Thumbprint; Name = 'unsigned file' }
            @{ Path = $signedScriptPath; Thumbprint = $expectedThumbprint; Name = 'different pinned signer' }
        )) {
            $rejectedState = Get-PinnedSignatureState -Path $rejectedCase.Path -Thumbprint $rejectedCase.Thumbprint
            Assert-True (-not $rejectedState.Accepted) ("Signer rule accepted a {0}: {1}" -f $rejectedCase.Name,$rejectedState.TrustResult)
            Assert-Throws -ScriptBlock {
                Assert-PinnedSignature -Path $rejectedCase.Path -Thumbprint $rejectedCase.Thumbprint
            } -ExpectedMessageLike '*signature validation failed*' -Message ("Assert-PinnedSignature accepted a {0}." -f $rejectedCase.Name)
        }
        $tamperedTrustResult = (Get-PinnedSignatureState -Path $tamperedScriptPath -Thumbprint $testSigner.Thumbprint).TrustResult
        Assert-True ($tamperedTrustResult -eq '0x80096010') "Tampered body must fail with TRUST_E_BAD_DIGEST: $tamperedTrustResult"
    }
    finally {
        $testSigner.Dispose()
        $ephemeralCertificate.Dispose()
        $ephemeralKey.Dispose()
    }

    $tamperRoot = Join-Path -Path $testRoot -ChildPath 'Tamper'
    $tamperCases = @('Clean', 'PackageManagerMetadata', 'ModifiedCmd', 'ExtraScript', 'MissingImage')
    $tamperResults = [ordered]@{}
    foreach ($tamperCase in $tamperCases) {
        $copyPath = Join-Path -Path $tamperRoot -ChildPath $tamperCase
        New-Item -ItemType Directory -Path $copyPath -Force | Out-Null
        Copy-Item -Path (Join-Path $buildResult.PackagePath '*') -Destination $copyPath -Recurse -Force
        switch ($tamperCase) {
            'PackageManagerMetadata' {
                Set-Content -LiteralPath (Join-Path $copyPath 'PSGetModuleInfo.xml') -Value '<Objs />' -Encoding UTF8
            }
            'ModifiedCmd' {
                Add-Content -LiteralPath (Join-Path $copyPath 'Runtime\Start-SmartM365-DeviceRebootManager-GUI.cmd') -Value 'calc.exe'
            }
            'ExtraScript' {
                Set-Content -LiteralPath (Join-Path $copyPath 'Runtime\Deploy\Extra.ps1') -Value 'Write-Output 1' -Encoding UTF8
            }
            'MissingImage' {
                Remove-Item -LiteralPath (Join-Path $copyPath 'Runtime\WorkplaceCloudHub.ico') -Force
            }
        }
        $integrityCall = {
            Assert-ModulePackageIntegrity `
                -ModuleBase $copyPath `
                -Thumbprint $expectedThumbprint `
                -SkipSignatureValidation:$SkipAuthenticodeValidation
        }
        if ($tamperCase -in @('Clean', 'PackageManagerMetadata')) {
            & $integrityCall
            $tamperResults[$tamperCase] = 'Accepted'
        }
        else {
            Assert-Throws -ScriptBlock $integrityCall `
                -ExpectedMessageLike '*file catalog validation failed*' `
                -Message ("Update helper accepted a tampered package: {0}" -f $tamperCase)
            $tamperResults[$tamperCase] = 'Refused'
        }
    }

    $vmPreview = $null
    if (-not $SkipAuthenticodeValidation) {
        $vmValidationScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
        $vmPreview = & $vmValidationScript -PackagePath $buildResult.PackagePath
        Assert-True ($vmPreview.Mode -eq 'Preview') 'VM validation script did not stay in Preview mode.'
        Assert-True (-not $vmPreview.ChangesAttempted) 'VM preview unexpectedly attempted changes.'
        Assert-True (-not $vmPreview.AutomaticUpdateDefault) 'VM preview does not report automatic updates disabled by default.'
        Assert-True $vmPreview.AllPinnedSignerMatches 'VM preview did not validate the pinned package signer.'
        Assert-True $vmPreview.AllSignaturesAccepted 'VM preview did not accept the pinned package signatures.'
        Assert-True $vmPreview.CatalogValid 'VM preview did not validate the package file catalog.'
        Assert-True $vmPreview.ScheduledTaskApiCompatible 'VM preview ScheduledTasks API compatibility preflight failed.'
    }

    $builtManifestPath = Join-Path $buildResult.PackagePath 'SmartM365.DeviceRebootManager.psd1'
    Import-Module -Name $builtManifestPath -Force
    $exportedNames = @((Get-Module SmartM365.DeviceRebootManager).ExportedFunctions.Keys)
    foreach ($expectedCommand in @(
        'Get-SmartM365DeviceRebootManager'
        'Install-SmartM365DeviceRebootManager'
        'Uninstall-SmartM365DeviceRebootManager'
        'Update-SmartM365DeviceRebootManager'
    )) {
        Assert-True ($exportedNames -contains $expectedCommand) ("Exported command is missing: {0}" -f $expectedCommand)
    }

    $fakeInstallPath = Join-Path -Path $testRoot -ChildPath 'NotInstalled'
    $status = Get-SmartM365DeviceRebootManager -InstallPath $fakeInstallPath
    Assert-True (-not $status.Installed) 'Status should report an absent installation.'
    $defaultStatus = Get-SmartM365DeviceRebootManager
    $expectedDefaultInstallPath = Join-Path -Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) -ChildPath 'WCH\SmartDeviceRebootManager'
    Assert-True ($defaultStatus.InstallPath -eq $expectedDefaultInstallPath) 'Default install path is not Program Files\WCH\SmartDeviceRebootManager.'

    $publishScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1'
    $preview = & $publishScript `
        -PackagePath $buildResult.PackagePath `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True ($preview.Mode -eq 'Preview') 'Publish script did not stay in Preview mode.'
    Assert-True (-not $preview.PublicationAttempted) 'Preview unexpectedly attempted publication.'
    Assert-True $preview.PackageValidated 'Publish preview did not validate the supplied package.'
    Assert-True $preview.PublicMetadataComplete 'Stable package must include approved public license metadata.'

    # A supplied -PackagePath is revalidated: tampered content and a missing
    # folder are refused before any publication step.
    $badPackagePath = Join-Path -Path $tamperRoot -ChildPath 'ModifiedCmd'
    Assert-Throws -ScriptBlock {
        & $publishScript -PackagePath $badPackagePath -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    } -ExpectedMessageLike '*file catalog validation failed*' -Message 'Publish accepted a tampered -PackagePath.'
    Assert-Throws -ScriptBlock {
        & $publishScript -PackagePath (Join-Path $testRoot 'DoesNotExist') -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    } -ExpectedMessageLike '*Package folder not found*' -Message 'Publish accepted a missing -PackagePath.'

    $missingApiKeyVariable = "SMARTM365_TEST_MISSING_API_KEY_$([guid]::NewGuid().ToString('N'))"
    $expectedGuardMessage = if ($SkipAuthenticodeValidation) {
        '*cannot be combined with -Execute*'
    }
    else {
        '*API key environment variable is empty*'
    }
    Assert-Throws -ScriptBlock {
        & $publishScript `
            -PackagePath $buildResult.PackagePath `
            -ApiKeyEnvironmentVariable $missingApiKeyVariable `
            -SkipAuthenticodeValidation:$SkipAuthenticodeValidation `
            -Execute `
            -Confirm:$false
    } -ExpectedMessageLike $expectedGuardMessage -Message 'Execute mode was not blocked by its publication guard.'

    [pscustomobject]@{
        Result                    = 'PASS'
        ModuleName                = $buildResult.ModuleName
        Version                   = $buildResult.Version
        Prerelease                = $buildResult.Prerelease
        PackageFileCount          = $buildResult.FileCount
        PowerShellFileCount       = $buildResult.PowerShellFileCount
        SignaturesValidated       = $buildResult.SignaturesValidated
        TaskArgumentBindingRuns   = $bindingRuns
        CatalogTamperResults      = ($tamperResults.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key,$_.Value }) -join '; '
        PublicationAttempted      = $preview.PublicationAttempted
        PublicMetadataComplete    = $preview.PublicMetadataComplete
        PublishRevalidatesPackage = $true
        PublicationGuard          = $true
        VmPreviewRun              = ($null -ne $vmPreview)
    }
}
finally {
    Remove-Module -Name SmartM365.DeviceRebootManager -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) {
        $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
        if ($resolvedTestRoot.StartsWith($resolvedTempRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
        }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDI5digeqa7HDPF
# ZMj9FbNkUiw7RpECvLUteBlpgFx+j6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKtesUFlYSmKdSYb2u998VDCq5EuMHTyUtHnIfJd2TMMMA0GCSqG
# SIb3DQEBAQUABIIBgC1zgr2tXeg+OKnRFaS/umhnYvTV/r61+2ORDxhpFQyNnOPU
# x+jIZL+yAeBoAzaV7tRcbNQ2AA0FZunPOwjFFb/FCU15Ro3wVyHMdNzbB9BUvsk9
# MwIqSbF6/An8P1Kec38NaVzH8tJQ8gKnkv4TmkjwQPbGQq1eyIfXnhNi8ds9QB1X
# sp+EulSCWw8hiLCRyyxGB31t57KkUT6gvDNOEk1F8HmPnEp4qsA1Q4DFv76D8LRy
# AuL4Mi9yyKtdp3UgnUPWmZqKpJnKmd8v9fWXS/iIULqsbP8mSIOy2sG5Y7AbWYVP
# x/TnGla7PrRq2vdUoPDLFUt1FXhdpUZ4uCRrwxKMkRre3OMXPEvZWYXR6V5+ADer
# w8tXrbAbmPm050syUhoHg6NGAFUQCaeevZXI8MXJwd91g/C/cclkRqA3nEJXJpt+
# 5MmzqOC/RO4TwtaC4TKYXveKaLg4FgHo0hJWAeZNlWtc7tlCTpFNK6Ls0FsYjbzN
# 6YjHc7u3kGtsDh8FA6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjRaMC8GCSqGSIb3DQEJBDEiBCBBN3cCkrRtOis+xiIZvbrvEJ4tFNnp+McrG7L3
# X814aTANBgkqhkiG9w0BAQEFAASCAgBor5I8YwERSZdaNz/o7h+pqthK2EYfVxoq
# AsNWeVC71wZf0oKijsVoP/KnrNBcabOfuwA5M06qV5OtpMkIJ+S+rZbPXOSnz2CW
# nXEgKMdJKloDREzdz3/5uoLAzAaKAvI3RH8YD3KmHNJ2g/8daOpIXCAqtMkdp0Ym
# 9r2ECORPL9mmtK1CYJGfbpkR/RFm3QmvT9qRa3N1I39KVFmtPau1eO7jq3ym22Jx
# WUKpuIp4oAxM3S5eqwuts+6HrM0648DUmXfrQljjjzrJM7/9Fb7+ahZ14GYrKKHA
# 170hzkZwyRagXLXa+ZFCivo1b9B7eP7+Zmuvg/oBdbhu4PDIuSsOJE7FuhgzhFOO
# rwINev70Jg7lcmWDi8hZascIYAX/YwhTGAqKLH7tonZ8PM5GEQdrKLglvGH1SHg8
# dPvr5CJ/jsYHcf5oxVH2e1XkddRnvonYEJCSk5m1vewviJh/TTs1ztXlWxuS0L8v
# eUJZizDGClGJstiUj5Wna/zOAyT0u8TUtDuxUAT4wy5ilDBoXkPr99uWbcmicyG1
# RoBFCxLGD5GV/wDyEOQS0/cSX3LKPcqNDXoAtPx5D5zrPVtA/E56STuNMWGUEAcu
# yxIQeWExfNqMqdDtBrZ+P/2xs4QbiYAZFFTyTRI5lGhyjybOcgCoCD/Bdc7mkLb1
# QE39IfqrKA==
# SIG # End signature block
