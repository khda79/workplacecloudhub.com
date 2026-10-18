#Requires -Version 7.0
<#
.SYNOPSIS
Fictional Contoso package definition, packaging, dry-run and publication results for the Smart IntunePackager demo.
.VERSION
1.0.3
.NOTES
Loaded by the local SmartM365.GuiDemo.ps1 host. No module is imported, no package is built and
Microsoft Graph is never called: every business action only updates demo memory or writes
synthetic files to the demo session folder.
#>

function script:Get-IpDemoAppId { '7c1e4b2a-93d5-4f60-a8b1-2e6d9f0c5a37' }

function script:Get-IpDemoPortalUrl { 'https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/' + (Get-IpDemoAppId) }

function script:Set-IpDemoDetectionPanels {
    $c = $script:SmartM365Demo.Controls
    $type = [string]$c.DetectionTypeCombo.SelectedItem
    foreach ($pair in @(@('MsiProductCode', 'MsiPanel'), @('File', 'FilePanel'), @('Registry', 'RegistryPanel'), @('Script', 'ScriptPanel'))) {
        $c[$pair[1]].Visibility = if ($type -eq $pair[0]) { 'Visible' } else { 'Collapsed' }
    }
}

function script:Get-IpDemoReturnCodes {
    @(
        [pscustomobject]@{ ReturnCode = 0; Type = 'success' }
        [pscustomobject]@{ ReturnCode = 1707; Type = 'success' }
        [pscustomobject]@{ ReturnCode = 3010; Type = 'softReboot' }
        [pscustomobject]@{ ReturnCode = 1641; Type = 'hardReboot' }
        [pscustomobject]@{ ReturnCode = 1618; Type = 'retry' }
    )
}

function script:Set-IpDemoResult {
    param([bool]$Published)
    $c = $script:SmartM365Demo.Controls
    if ($Published) {
        $c.ResultAppIdBox.Text = Get-IpDemoAppId
        $c.PortalLinkRun.Text = Get-IpDemoPortalUrl
    }
    else {
        $c.ResultAppIdBox.Text = ''
        $c.PortalLinkRun.Text = 'No app published in this session.'
    }
}

function script:Set-IpDemoUpdateMode {
    param([bool]$Enabled)
    $s = $script:SmartM365Demo; $c = $s.Controls
    if ($Enabled) {
        $app = $s.Data.Apps[0]
        $c.PublishModeCombo.SelectedItem = 'UpdateExisting'
        $c.ExistingAppIdBox.Text = $app.id
        $c.UpdateModeText.Text = 'Update mode: {0} ({1}), current content version {2}. Click Publish package to upload content version {3}.' -f $app.displayName, $app.id, $app.ContentVersion, ([int]$app.ContentVersion + 1)
        $c.UpdateModeBanner.Visibility = 'Visible'
        $c.PublishButton.BorderBrush = [Windows.Media.Brushes]::Goldenrod
        $c.PublishButton.BorderThickness = [Windows.Thickness]::new(3)
    }
    else {
        $c.PublishModeCombo.SelectedItem = 'Create'
        $c.ExistingAppIdBox.Text = ''
        $c.UpdateModeBanner.Visibility = 'Collapsed'
        $c.PublishButton.ClearValue([Windows.Controls.Control]::BorderBrushProperty)
        $c.PublishButton.ClearValue([Windows.Controls.Control]::BorderThicknessProperty)
    }
}

function script:Initialize-SmartM365DemoProfile {
    $s = $script:SmartM365Demo; $c = $s.Controls
    $operators = @('notConfigured', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual')
    Set-SmartM365DemoCombo TenantCombo @('test', 'prod')
    Set-SmartM365DemoCombo RunAsCombo @('system', 'user')
    Set-SmartM365DemoCombo RestartCombo @('suppress', 'allow', 'basedOnReturnCode', 'force')
    Set-SmartM365DemoCombo MinWindowsCombo @('1607', '1809', 'Windows10_21H2', 'Windows10_22H2', 'Windows11_21H2', 'Windows11_22H2', 'Windows11_23H2', 'Windows11_24H2') 6
    Set-SmartM365DemoCombo DetectionTypeCombo @('MsiProductCode', 'File', 'Registry', 'Script')
    Set-SmartM365DemoCombo MsiVersionOperatorCombo $operators 4
    Set-SmartM365DemoCombo PublishModeCombo @('Create', 'UpdateExisting')

    $c.SourceFolderBox.Text = 'C:\Packages\Contoso\SampleApp\Source'
    $c.SetupFileBox.Text = 'ContosoSampleApp.msi'
    $c.OutputFolderBox.Text = 'C:\Packages\Contoso\SampleApp\Output'
    $c.DisplayNameBox.Text = 'Contoso Sample App'
    $c.DescriptionBox.Text = 'DEMO - Fictional line-of-business application packaged by Smart IntunePackager.'
    $c.PublisherBox.Text = 'Contoso Ltd'
    $c.VersionBox.Text = '2.4.0'
    $c.DeveloperBox.Text = 'Contoso Ltd'
    $c.OwnerBox.Text = 'Contoso IT'
    $c.InformationUrlBox.Text = 'https://www.contoso.example/sample-app'
    $c.PrivacyUrlBox.Text = ''
    $c.NotesBox.Text = 'DEMO - Sample data.'
    $c.LogoPathBox.Text = 'C:\Packages\Contoso\SampleApp\contoso-sample.png'
    $c.InstallCommandBox.Text = 'msiexec /i "ContosoSampleApp.msi" /qn /norestart'
    $c.UninstallCommandBox.Text = 'msiexec /x {5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10} /qn /norestart'
    $c.MaxRunTimeBox.Text = '30'
    $c.ArchX64Check.IsChecked = $true; $c.ArchArm64Check.IsChecked = $true; $c.ArchX86Check.IsChecked = $false
    $c.MsiProductCodeBox.Text = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'
    $c.MsiVersionBox.Text = '2.4.0'
    $c.ScriptPathBox.Text = 'C:\Packages\Contoso\SampleApp\Detect-ContosoSampleApp.ps1'
    $c.EnforceSignatureCheck.IsChecked = $false; $c.RunAs32BitCheck.IsChecked = $false
    $c.ExistingAppIdBox.Text = ''
    $c.FindAppBox.Text = 'Contoso'
    $c.DryRunCheck.IsChecked = $false
    $c.ProfilePathText.Text = 'C:\Packages\Contoso\Profiles\Contoso-SampleApp.profile.json (demo)'

    $c.ReturnCodesGrid.Columns[1].ItemsSource = @('success', 'softReboot', 'hardReboot', 'retry', 'failed')
    foreach ($grid in @('FileRulesGrid', 'RegistryRulesGrid')) { $c[$grid].Columns[3].ItemsSource = $operators }
    $c.FileRulesGrid.Columns[2].ItemsSource = @('exists', 'doesNotExist', 'modifiedDate', 'createdDate', 'version', 'sizeInMB')
    $c.RegistryRulesGrid.Columns[2].ItemsSource = @('exists', 'doesNotExist', 'string', 'integer', 'version')
    $s.Data.ReturnCodes = @(Get-IpDemoReturnCodes)
    Set-SmartM365DemoGrid ReturnCodesGrid $s.Data.ReturnCodes
    $s.Data.FileRules = @([pscustomobject]@{ Path = 'C:\Program Files\Contoso\SampleApp'; FileOrFolderName = 'ContosoSample.exe'; DetectionType = 'version'; Operator = 'greaterThanOrEqual'; DetectionValue = '2.4.0'; Check32BitOn64System = $false })
    Set-SmartM365DemoGrid FileRulesGrid $s.Data.FileRules
    $s.Data.RegistryRules = @([pscustomobject]@{ KeyPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Contoso\SampleApp'; ValueName = 'Version'; DetectionType = 'version'; Operator = 'greaterThanOrEqual'; DetectionValue = '2.4.0'; Check32BitOn64System = $false })
    Set-SmartM365DemoGrid RegistryRulesGrid $s.Data.RegistryRules
    $s.Data.Apps = @(
        [pscustomobject]@{ id = Get-IpDemoAppId; displayName = 'Contoso Sample App'; publisher = 'Contoso Ltd'; LastModified = '2026-10-17 18:42'; Version = '2.3.1'; ContentVersion = '3' }
        [pscustomobject]@{ id = 'a4f2c9e1-6b38-4d7a-9e05-3c8b1f7d2e46'; displayName = 'Contoso Field Tools'; publisher = 'Contoso Ltd'; LastModified = '2026-10-02 10:15'; Version = '5.0.2'; ContentVersion = '1' }
        [pscustomobject]@{ id = 'd9b3e7a2-1c54-4f8e-b6a0-8e2f4c1d7b93'; displayName = 'Contoso VPN Client'; publisher = 'Contoso Ltd'; LastModified = '2026-09-21 13:03'; Version = '3.2.0'; ContentVersion = '2' }
    )
    Set-SmartM365DemoGrid AppsGrid $s.Data.Apps
    Set-IpDemoSelectedAppDetails

    $c.PackageInfoText.Text = "DEMO - File: C:\Packages\Contoso\SampleApp\Output\ContosoSampleApp.intunewin`nSHA256: 3F6A0C1D9E8B7A65F4E3D2C1B0A99887766554433221100FFEEDDCCBBAA99887`nEncrypted size: 48.2 MB`nSetup file: ContosoSampleApp.msi`nEngine: IntuneWinAppUtil 1.8.7"
    $c.LogTextBox.Text = @(
        '2026-10-18 09:30:00 Smart IntunePackager DEMO. Fictional Contoso data; Microsoft Graph is not called.'
        '2026-10-18 09:30:01 MSI metadata detected: Name=Contoso Sample App; Version=2.4.0; Publisher=Contoso Ltd'
        '2026-10-18 09:30:12 Packaging completed in 11.4 s. EncryptedSize=48.2 MB'
    ) -join [Environment]::NewLine
    $c.ConnectionStatusText.Text = 'Not connected (demo)'
    $c.OperationProgress.IsIndeterminate = $false
    $c.OperationProgress.Value = 0
    $c.StatusText.Text = 'DEMO - Ready'
    Set-IpDemoResult -Published $false
    Set-IpDemoUpdateMode -Enabled $false

    switch ($s.Scenario) {
        'Issues' {
            $c.UninstallCommandBox.Text = ''
            $c.InformationUrlBox.Text = 'contoso-sample'
            $c.StatusText.Text = 'DEMO - The app definition has 2 issues'
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:31:00 [WARN] Program.UninstallCommandLine is required.')
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:31:00 [WARN] App.InformationUrl must be an absolute http or https URL.')
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:31:20 [WARN] [ContentFile] POST /beta/deviceAppManagement/mobileApps returned HTTP 429. Retry 1/5 in 7 s.')
        }
        'Completed' {
            $c.ConnectionStatusText.Text = 'Connected: test (demo tenant)'
            $c.OperationProgress.Value = 100
            $c.StatusText.Text = 'DEMO - Publish completed'
            Set-IpDemoResult -Published $true
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:32:40 [Upload] Uploaded 4 block(s), 48.2 MB.')
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:32:55 Publication completed. AppId=' + (Get-IpDemoAppId) + '; ContentVersion=1')
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:32:55 Intune admin center: ' + (Get-IpDemoPortalUrl))
            $c.LogTextBox.AppendText([Environment]::NewLine + '2026-10-18 09:32:55 [WARN] The app is not assigned. Create assignments in the Intune admin center.')
        }
    }
    Set-IpDemoDetectionPanels
}

function script:Set-IpDemoSelectedAppDetails {
    $c = $script:SmartM365Demo.Controls
    $app = $c.AppsGrid.SelectedItem
    if ($app) {
        $c.SelectedAppDetailsText.Text = "DEMO - {0}`nVersion: {1}; committed content version: {2}; setup file: ContosoSampleApp.msi; publishing state: published; assigned: False" -f $app.displayName, $app.Version, $app.ContentVersion
    }
    else { $c.SelectedAppDetailsText.Text = 'Select an app to read its version and committed content version.' }
}

function script:Connect-SmartM365DemoProfileEvents {
    $c = $script:SmartM365Demo.Controls
    $c.DetectionTypeCombo.Add_SelectionChanged({ Set-IpDemoDetectionPanels })
    $c.AppsGrid.Add_SelectionChanged({ Set-IpDemoSelectedAppDetails })
}

function script:Save-IpDemoJson {
    param([string]$Name, $Value)
    $path = Get-SmartM365DemoOutputPath $Name
    [IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $Value -Depth 10), [Text.UTF8Encoding]::new($false))
    return $path
}

function script:Get-IpDemoProfile {
    $c = $script:SmartM365Demo.Controls
    [ordered]@{
        SchemaVersion = 1
        Demo          = 'DEMO - Sample data'
        Package       = [ordered]@{ SourceFolder = $c.SourceFolderBox.Text; SetupFile = $c.SetupFileBox.Text; OutputFolder = $c.OutputFolderBox.Text }
        App           = [ordered]@{ DisplayName = $c.DisplayNameBox.Text; Publisher = $c.PublisherBox.Text; Version = $c.VersionBox.Text; Description = $c.DescriptionBox.Text }
        Program       = [ordered]@{ InstallCommandLine = $c.InstallCommandBox.Text; UninstallCommandLine = $c.UninstallCommandBox.Text; RunAsAccount = [string]$c.RunAsCombo.SelectedItem; MaxRunTimeInMinutes = $c.MaxRunTimeBox.Text }
        Requirements  = [ordered]@{ Architectures = @('x64', 'arm64'); MinimumWindowsRelease = [string]$c.MinWindowsCombo.SelectedItem }
        Detection     = [ordered]@{ Type = [string]$c.DetectionTypeCombo.SelectedItem; Msi = [ordered]@{ ProductCode = $c.MsiProductCodeBox.Text; ProductVersionOperator = [string]$c.MsiVersionOperatorCombo.SelectedItem; ProductVersion = $c.MsiVersionBox.Text } }
        Publish       = [ordered]@{ Mode = [string]$c.PublishModeCombo.SelectedItem; ExistingAppId = $c.ExistingAppIdBox.Text }
    }
}

function script:Get-IpDemoCreatePayload {
    $c = $script:SmartM365Demo.Controls
    [ordered]@{
        Step   = 'CreateApp'
        Method = 'POST'
        Uri    = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
        Note   = 'DEMO - Sample payload; Microsoft Graph was not called.'
        Body   = [ordered]@{
            '@odata.type'                  = '#microsoft.graph.win32LobApp'
            displayName                    = $c.DisplayNameBox.Text
            publisher                      = $c.PublisherBox.Text
            displayVersion                 = $c.VersionBox.Text
            setupFilePath                  = $c.SetupFileBox.Text
            installCommandLine             = $c.InstallCommandBox.Text
            uninstallCommandLine           = $c.UninstallCommandBox.Text
            applicableArchitectures        = 'none'
            allowedArchitectures           = 'x64,arm64'
            minimumSupportedWindowsRelease = [string]$c.MinWindowsCombo.SelectedItem
            installExperience              = [ordered]@{ runAsAccount = [string]$c.RunAsCombo.SelectedItem; deviceRestartBehavior = [string]$c.RestartCombo.SelectedItem; maxRunTimeInMinutes = 30 }
            detectionRules                 = @([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppProductCodeDetection'; productCode = $c.MsiProductCodeBox.Text; productVersionOperator = 'greaterThanOrEqual'; productVersion = $c.MsiVersionBox.Text })
            returnCodes                    = @(Get-IpDemoReturnCodes | ForEach-Object { [ordered]@{ returnCode = $_.ReturnCode; type = $_.Type } })
        }
    }
}

function script:Invoke-SmartM365DemoProfileAction {
    param([string]$Name)
    $s = $script:SmartM365Demo; $c = $s.Controls
    switch -Regex ($Name) {
        '^ConnectButton$' { $c.ConnectionStatusText.Text = 'Connected: ' + [string]$c.TenantCombo.SelectedItem + ' (demo tenant)'; Write-SmartM365DemoActivity 'Graph connection simulated. No sign-in was requested.'; return $true }
        '^DisconnectButton$' { $c.ConnectionStatusText.Text = 'Not connected (demo)'; Write-SmartM365DemoActivity 'Graph disconnection simulated.'; return $true }
        '^DetectMetadataButton$' {
            $c.DisplayNameBox.Text = 'Contoso Sample App'; $c.PublisherBox.Text = 'Contoso Ltd'; $c.VersionBox.Text = '2.4.0'
            $c.MsiProductCodeBox.Text = '{5F0C3B2A-6D41-4E8B-9A57-2C1D7E3F9B10}'; $c.DetectionTypeCombo.SelectedItem = 'MsiProductCode'
            Write-SmartM365DemoActivity 'MSI metadata detection simulated: name, version, publisher and product code prefilled.'
            return $true
        }
        '^BuildButton$' { $c.OperationProgress.Value = 100; Write-SmartM365DemoActivity 'Package build simulated; the Win32 Content Prep Tool was not run.'; return $true }
        '^(PublishButton|BuildAndPublishButton)$' {
            if ($c.DryRunCheck.IsChecked) {
                $path = Save-IpDemoJson 'IntunePackager-DryRun-CreateApp-DEMO.json' (Get-IpDemoCreatePayload)
                Write-SmartM365DemoActivity ('Dry run simulated: payload written to ' + $path)
            }
            else {
                Set-IpDemoResult -Published $true
                $c.OperationProgress.Value = 100
                $c.MainTabs.SelectedIndex = 5
                Write-SmartM365DemoActivity ('Publication simulated. AppId=' + (Get-IpDemoAppId) + '. The app is not assigned; Intune was not contacted.')
            }
            return $true
        }
        '^FindAppButton$' {
            $text = $c.FindAppBox.Text.Trim()
            Set-SmartM365DemoGrid AppsGrid @($s.Data.Apps | Where-Object { -not $text -or $_.displayName -like ('*' + $text + '*') })
            Write-SmartM365DemoActivity 'Win32 app search simulated on fictional data.'
            return $true
        }
        '^(UseSelectedAppButton|UseResultForUpdateButton)$' {
            Set-IpDemoUpdateMode -Enabled $true
            Write-SmartM365DemoActivity 'Update mode simulated: the next publication adds content version 4 to the fictional app.'
            return $true
        }
        '^SwitchToCreateButton$' { Set-IpDemoUpdateMode -Enabled $false; Write-SmartM365DemoActivity 'Create mode restored.'; return $true }
        '^CopyAppIdButton$' { Write-SmartM365DemoActivity 'App id copy simulated; the clipboard was not changed.'; return $true }
        '^OpenPortalButton$' { Write-SmartM365DemoActivity ('Intune admin center link (not opened in demo): ' + (Get-IpDemoPortalUrl)); return $true }
        '^Browse' { Write-SmartM365DemoActivity 'File selection simulated; no file system browsing in demo mode.'; return $true }
        '^(SaveProfileButton|SaveProfileAsButton)$' { $path = Save-IpDemoJson 'Contoso-SampleApp-DEMO.profile.json' (Get-IpDemoProfile); Write-SmartM365DemoActivity ('Profile saved to the demo session: ' + $path); return $true }
        '^(NewProfileButton|OpenProfileButton)$' { Initialize-SmartM365DemoProfile; Write-SmartM365DemoActivity 'Fictional profile loaded.'; return $true }
        '^ValidateButton$' {
            $message = if ($s.Scenario -eq 'Issues') { 'Validation simulated: 2 issues (uninstall command, information URL).' } else { 'Validation simulated: the app definition is valid.' }
            Write-SmartM365DemoActivity $message
            return $true
        }
        '^AddReturnCodeButton$' { $s.Data.ReturnCodes = @($s.Data.ReturnCodes) + [pscustomobject]@{ ReturnCode = 1603; Type = 'failed' }; Set-SmartM365DemoGrid ReturnCodesGrid $s.Data.ReturnCodes; Write-SmartM365DemoActivity 'Return code added in demo memory.'; return $true }
        '^RemoveReturnCodeButton$' { $s.Data.ReturnCodes = @($s.Data.ReturnCodes | Select-Object -SkipLast 1); Set-SmartM365DemoGrid ReturnCodesGrid $s.Data.ReturnCodes; Write-SmartM365DemoActivity 'Return code removed in demo memory.'; return $true }
        '^DefaultReturnCodesButton$' { $s.Data.ReturnCodes = @(Get-IpDemoReturnCodes); Set-SmartM365DemoGrid ReturnCodesGrid $s.Data.ReturnCodes; Write-SmartM365DemoActivity 'Intune default return codes restored.'; return $true }
        '^(Add|Remove)(File|Registry)RuleButton$' { Write-SmartM365DemoActivity ($Name + ' simulated in demo memory.'); return $true }
        '^ExportLogButton$' { $path = Get-SmartM365DemoOutputPath 'IntunePackager-Log-DEMO.txt'; [IO.File]::WriteAllText($path, $c.LogTextBox.Text, [Text.UTF8Encoding]::new($false)); Write-SmartM365DemoActivity ('Log exported: ' + $path); return $true }
        '^OpenLogFolderButton$' { Write-SmartM365DemoActivity 'Log folder opening simulated.'; return $true }
        '^ClearLogButton$' { $c.LogTextBox.Clear(); return $true }
    }
    return $false
}

function script:Export-SmartM365DemoProfileOutputs {
    $c = $script:SmartM365Demo.Controls
    $profileFile = Save-IpDemoJson 'Contoso-SampleApp-DEMO.profile.json' (Get-IpDemoProfile)
    $payload = Save-IpDemoJson 'IntunePackager-DryRun-CreateApp-DEMO.json' (Get-IpDemoCreatePayload)
    $log = Get-SmartM365DemoOutputPath 'IntunePackager-Log-DEMO.txt'
    [IO.File]::WriteAllText($log, ('DEMO - Sample data' + [Environment]::NewLine + $c.LogTextBox.Text), [Text.UTF8Encoding]::new($false))
    return @($profileFile, $payload, $log)
}
