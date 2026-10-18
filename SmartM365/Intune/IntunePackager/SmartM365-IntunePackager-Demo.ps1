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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA2BlCsY51iAy16
# D2q64dbuCgMxE1NXHIcdgtIHEc2ddqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMEtIo/TyVW5ljyhqXWLHuMK5eAH6cEqtajLUPGkuoIRMA0GCSqG
# SIb3DQEBAQUABIIBgKiWQ62s853VfnIvdZWnvMBL3nwS/qaxe+glAk5bLyk/Jf7M
# nri0Sxraxu7lI1oKYGdk1wm7Tlcf26+r6WXkQffDP5Dp8lvf3v/oCPBUlHcxRj/G
# Z7/AGzbXyX5Rct0kFoCHtVsRqVTYKefEAuy22mMCEctCMUZwPxchNwqN5BfHywU9
# KGd8ZQz6/YuBpOvNXxNoIe/La9ProEn+xVmLBb+5kche7iVe7LgIkaNnACbb33fn
# pKHhRuTvLvVM7Xte0AMyytRiY2ANSqIVtkAKCOBhgzwTtyoWr1+FZdZ0S9pt9Lit
# do8ztp1RrjG4nC7jlBK6ny48sBucFMdQJG51uW8kY96OhMqcl8oO6+bi+PqrfaCf
# 2OQSWTsQXzfOqaPvXj47kqBpVdAh/HZCXn6N/67z3cGcFgi+a4hNtFHPB6jCzSZd
# o+bbr+7fhKysXj/j56+gFRxmYQ1r9nMxzcv2R1hdndri3S9uHfADLosfyA0p6kIN
# d1XSxnDjLEkBvLcQL6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzNaMC8GCSqGSIb3DQEJBDEiBCC+eNVJVrYjVM7Ds4b7tdNGiEeQEOVNJ/6Oh4Sa
# SKggkDANBgkqhkiG9w0BAQEFAASCAgBQkBU8IJfEmbAM1xjVgEpdVBWxHivJNyoj
# sh5QM3PYW8P+k2cFO5fTgZAfsEIImVVTYWN58HNUWs3Vnkax1U7UFVoNllF5s6Q1
# XDUISsreESh1YQsxKzQsH0Sa/2JVi56v3pJPSsRaDL44EDtJaRP+K+w3pCZQzcSE
# 7ErRvNNKUjlCemY/c+FdZLicYiUIOekVRCwIVWvsphd/4VMPwYRDwZLNBM4mFZMl
# 6Zyj6rG01quStzr2SJmn93APy9nIrtIWsUwHKv3tRgte8CbCvag+6ls+JkrpQezr
# rDfXPaMFjxbFsQjRIt/DjbQ4M+evK93oc0U16o1vcQhpLfiJo1p4WA0gDRMiGV4M
# qE5QGzfONfg95CA5pVw9jmaWOwg67gPNSYG0hDTYSQLtF0+I5HbhMX1ZS++9ypiz
# GA0a5DNe63H+LXjjXfDmqpqlkwlmuMDl3t3/Smk2LI1g+PotFSvmjmphoeljjwxI
# NEbKHgYqsgli1vUY8R2c1QjLhw7UMWAvFCh3FsVxIDmBavDoEyt74COrBcbhKD2H
# RTDNmvuL6Uks58FJSJAk/MhhVp16b/0XPdOzQ0lO0T0xDWijXmOpEIqOZQZRmZMK
# lfRgndLKf98g7vKkHanrLQI2SrKs9mTqmepBfJw9dNjJGe8dQtrlfWUfpbU/GMZe
# MhW9nJrRVQ==
# SIG # End signature block
