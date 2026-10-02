<#
.SYNOPSIS
    WPF dialog for validated creation of a local SharePoint migration.
#>

function Show-SmartM365NewMigrationWizard {
    param(
        $Owner,
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [ValidateSet('Interactive', 'DeviceLogin', 'Certificate')][string]$AuthMode = 'Interactive',
        [string]$ActivityPath = '',
        [switch]$ValidateOnly
    )

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New SharePoint migration" Width="760" Height="700"
        MinWidth="690" MinHeight="620" WindowStartupLocation="CenterOwner"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Margin="0,0,0,12">
      <TextBlock Text="Create a migration" FontSize="20" FontWeight="SemiBold" Foreground="#1F2937"/>
      <TextBlock Text="Enter the migration scope. Every site URL is checked before the folder is created."
                 Foreground="#5F6B7A" Margin="0,4,0,0" TextWrapping="Wrap"/>
    </StackPanel>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel x:Name="form" Margin="0,0,8,0">
        <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
          <StackPanel>
            <TextBlock Text="MIGRATION" Foreground="#0078D4" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <TextBlock Text="Folder and report name" Foreground="#1F2937"/>
            <TextBox x:Name="txtName" Height="29" Margin="0,4,0,0"/>
          </StackPanel>
        </Border>
        <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
          <StackPanel>
            <TextBlock Text="SOURCE" Foreground="#0078D4" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <TextBlock Text="SharePoint type" Foreground="#1F2937"/>
            <ComboBox x:Name="cmbSourceType" Height="29" Margin="0,4,0,9">
              <ComboBoxItem Content="SP2019" IsSelected="True"/>
              <ComboBoxItem Content="SP2016"/>
              <ComboBoxItem Content="SPO"/>
            </ComboBox>
            <TextBlock Text="Source web URL (first mapping)" Foreground="#1F2937"/>
            <TextBox x:Name="txtSourceUrl" Height="29" Margin="0,4,0,9"/>
            <TextBlock Text="Web application URL for SP2016/SP2019 (blank = URL host)" Foreground="#1F2937"/>
            <TextBox x:Name="txtSourceApp" Height="29" Margin="0,4,0,9"/>
            <TextBlock Text="Permission root path (blank = common mapping path)" Foreground="#1F2937"/>
            <TextBox x:Name="txtSourceRoot" Height="29" Margin="0,4,0,0"/>
          </StackPanel>
        </Border>
        <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
          <StackPanel>
            <TextBlock Text="TARGET" Foreground="#0078D4" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <TextBlock Text="SharePoint type" Foreground="#1F2937"/>
            <ComboBox x:Name="cmbTargetType" Height="29" Margin="0,4,0,9">
              <ComboBoxItem Content="SPO" IsSelected="True"/>
              <ComboBoxItem Content="SP2019"/>
              <ComboBoxItem Content="SP2016"/>
            </ComboBox>
            <TextBlock Text="Target web URL (first mapping)" Foreground="#1F2937"/>
            <TextBox x:Name="txtTargetUrl" Height="29" Margin="0,4,0,9"/>
            <TextBlock Text="Tenant admin URL for SPO" Foreground="#1F2937"/>
            <TextBox x:Name="txtTargetAdmin" Height="29" Margin="0,4,0,9"/>
            <TextBlock Text="Web application URL for SP2016/SP2019 (blank = URL host)" Foreground="#1F2937"/>
            <TextBox x:Name="txtTargetApp" Height="29" Margin="0,4,0,9"/>
            <TextBlock Text="Permission root path (blank = common mapping path)" Foreground="#1F2937"/>
            <TextBox x:Name="txtTargetRoot" Height="29" Margin="0,4,0,0"/>
          </StackPanel>
        </Border>
        <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,0,10">
          <StackPanel>
            <TextBlock Text="ADDITIONAL MAPPINGS" Foreground="#0078D4" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <TextBlock Text="Optional: one source URL and one target URL per line, separated by a space. Encode spaces in URLs as %20."
                       TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,0,0,6"/>
            <TextBox x:Name="txtAdditionalMappings" Height="82" AcceptsReturn="True"
                     VerticalScrollBarVisibility="Auto" TextWrapping="NoWrap"/>
          </StackPanel>
        </Border>
        <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14">
          <StackPanel>
            <TextBlock Text="VALIDATION AUTHENTICATION" Foreground="#0078D4" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <ComboBox x:Name="cmbWizardAuth" Height="29">
              <ComboBoxItem Content="Interactive" IsSelected="True"/>
              <ComboBoxItem Content="Device login"/>
              <ComboBoxItem Content="Certificate"/>
            </ComboBox>
            <TextBlock Text="On-premises URLs use the current Windows account. SPO checks may open a sign-in window."
                       Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,7,0,0"/>
          </StackPanel>
        </Border>
      </StackPanel>
    </ScrollViewer>
    <Grid Grid.Row="2" Margin="0,12,0,0">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <TextBlock x:Name="lblStatus" Text="Ready to validate URLs." Foreground="#5F6B7A"
                 TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <StackPanel Grid.Column="1" Orientation="Horizontal">
        <Button x:Name="btnCancel" Content="Cancel" MinWidth="78" Height="30" Margin="0,0,7,0"/>
        <Button x:Name="btnCreate" Content="Validate and create" MinWidth="145" Height="30"
                Background="#0078D4" Foreground="White" BorderThickness="0"/>
      </StackPanel>
    </Grid>
  </Grid>
</Window>
'@
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($ValidateOnly) {
        foreach ($name in @('form', 'lblStatus', 'btnCreate', 'btnCancel', 'txtName',
                'cmbSourceType', 'txtSourceUrl', 'txtSourceApp', 'txtSourceRoot',
                'cmbTargetType', 'txtTargetUrl', 'txtTargetApp', 'txtTargetAdmin',
                'txtTargetRoot', 'txtAdditionalMappings', 'cmbWizardAuth')) {
            if ($null -eq $dialog.FindName($name)) { throw "Wizard control not found: $name" }
        }
        return $true
    }
    if ($null -eq $Owner) { throw 'An owner window is required for the migration wizard.' }
    $dialog.Owner = $Owner
    $state = [pscustomobject]@{
        Dialog = $dialog
        Form = $dialog.FindName('form')
        Status = $dialog.FindName('lblStatus')
        Create = $dialog.FindName('btnCreate')
        Cancel = $dialog.FindName('btnCancel')
        Name = $dialog.FindName('txtName')
        SourceType = $dialog.FindName('cmbSourceType')
        SourceUrl = $dialog.FindName('txtSourceUrl')
        SourceApp = $dialog.FindName('txtSourceApp')
        SourceRoot = $dialog.FindName('txtSourceRoot')
        TargetType = $dialog.FindName('cmbTargetType')
        TargetUrl = $dialog.FindName('txtTargetUrl')
        TargetApp = $dialog.FindName('txtTargetApp')
        TargetAdmin = $dialog.FindName('txtTargetAdmin')
        TargetRoot = $dialog.FindName('txtTargetRoot')
        Additional = $dialog.FindName('txtAdditionalMappings')
        Auth = $dialog.FindName('cmbWizardAuth')
        Process = $null
        RequestPath = $null
        ResultPath = $null
        CreatedName = $null
        Timer = $null
    }
    $state.Auth.SelectedIndex = switch ($AuthMode) {
        'DeviceLogin' { 1 }
        'Certificate' { 2 }
        default { 0 }
    }
    $updateTypeFields = {
        $sourceIsSPO = [string]$state.SourceType.SelectedItem.Content -eq 'SPO'
        $targetIsSPO = [string]$state.TargetType.SelectedItem.Content -eq 'SPO'
        $state.SourceApp.IsEnabled = -not $sourceIsSPO
        $state.TargetApp.IsEnabled = -not $targetIsSPO
        $state.TargetAdmin.IsEnabled = $targetIsSPO
        $state.Auth.IsEnabled = $sourceIsSPO -or $targetIsSPO
    }.GetNewClosure()
    $state.SourceType.Add_SelectionChanged($updateTypeFields)
    $state.TargetType.Add_SelectionChanged($updateTypeFields)
    & $updateTypeFields

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(350)
    $state.Timer = $timer
    $timer.Add_Tick({
        if ($null -eq $state.Process -or -not $state.Process.HasExited) { return }
        $state.Timer.Stop()
        $state.Form.IsEnabled = $true
        $state.Create.IsEnabled = $true
        $state.Cancel.IsEnabled = $true
        try {
            if (-not (Test-Path -LiteralPath $state.ResultPath -PathType Leaf)) {
                throw "Validation process exited with code $($state.Process.ExitCode) without a result."
            }
            $result = Get-Content -LiteralPath $state.ResultPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $result.Success) { throw [string]$result.Message }
            $state.CreatedName = [string]$result.MigrationName
            $state.Status.Text = "Created $($state.CreatedName) after $($result.ValidatedUrls) URL checks."
            $state.Dialog.Close()
        }
        catch {
            $state.Status.Foreground = [System.Windows.Media.Brushes]::Firebrick
            $state.Status.Text = $_.Exception.Message
            if ($ActivityPath) {
                Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status 'ValidationFailed' `
                    -Detail $_.Exception.Message -ExitCode 1
            }
        }
        finally {
            foreach ($path in @($state.RequestPath, $state.ResultPath)) {
                if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                    Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
                }
            }
            $state.Process.Dispose()
            $state.Process = $null
        }
    }.GetNewClosure())

    $state.Create.Add_Click({
        try {
            $name = $state.Name.Text.Trim()
            $source = $state.SourceUrl.Text.Trim()
            $target = $state.TargetUrl.Text.Trim()
            if (-not $name -or -not $source -or -not $target) {
                throw 'Name, source web URL and target web URL are required.'
            }
            $rows = [System.Collections.Generic.List[object]]::new()
            $rows.Add([pscustomobject]@{ SourceUrl = $source; TargetUrl = $target })
            $lineNumber = 0
            foreach ($line in ($state.Additional.Text -split '\r?\n')) {
                $lineNumber++
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $parts = @($line.Trim() -split '\s+')
                if ($parts.Count -ne 2) {
                    throw "Additional mapping line $lineNumber must contain exactly two URLs."
                }
                $rows.Add([pscustomobject]@{ SourceUrl = $parts[0]; TargetUrl = $parts[1] })
            }
            $mode = [string]$state.Auth.SelectedItem.Content
            if ($mode -eq 'Device login') { $mode = 'DeviceLogin' }
            $request = [ordered]@{
                Name = $name
                SourceType = [string]$state.SourceType.SelectedItem.Content
                TargetType = [string]$state.TargetType.SelectedItem.Content
                SourceWebApplicationUrl = $state.SourceApp.Text.Trim()
                TargetWebApplicationUrl = $state.TargetApp.Text.Trim()
                TargetAdminUrl = $state.TargetAdmin.Text.Trim()
                SourcePermissionRootPath = $state.SourceRoot.Text.Trim()
                TargetPermissionRootPath = $state.TargetRoot.Text.Trim()
                AuthMode = $mode
                Mappings = $rows.ToArray()
            }
            $helper = Join-Path $ProjectRoot 'Scripts\SmartM365-SharePointMigration-NewMigration.ps1'
            $pwsh = Join-Path $PSHOME 'pwsh.exe'
            if (-not (Test-Path -LiteralPath $pwsh -PathType Leaf)) {
                throw 'PowerShell 7.4 or later is required to validate SharePoint Online URLs.'
            }
            $state.RequestPath = [System.IO.Path]::GetTempFileName()
            $state.ResultPath = [System.IO.Path]::ChangeExtension($state.RequestPath, '.result.json')
            $json = $request | ConvertTo-Json -Depth 8
            [System.IO.File]::WriteAllText($state.RequestPath, $json, [System.Text.UTF8Encoding]::new($false))
            $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$helper`"",
                '-RequestPath', "`"$($state.RequestPath)`"", '-ResultPath', "`"$($state.ResultPath)`"")
            $start = @{ FilePath = $pwsh; ArgumentList = $arguments; PassThru = $true }
            if ($mode -eq 'Certificate') { $start.WindowStyle = 'Hidden' }
            $state.Process = Start-Process @start
            if ($ActivityPath) {
                Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status 'Validating' `
                    -Detail 'Checking configured SharePoint URLs.'
            }
            $state.Form.IsEnabled = $false
            $state.Create.IsEnabled = $false
            $state.Status.Foreground = [System.Windows.Media.Brushes]::DarkSlateGray
            $state.Status.Text = 'Checking every SharePoint URL. Complete sign-in if prompted; the folder will be created only after all checks succeed.'
            $state.Timer.Start()
        }
        catch {
            $state.Status.Foreground = [System.Windows.Media.Brushes]::Firebrick
            $state.Status.Text = $_.Exception.Message
            foreach ($path in @($state.RequestPath, $state.ResultPath)) {
                if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                    Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }.GetNewClosure())
    $cancelValidation = {
        $state.Timer.Stop()
        if ($state.Process) {
            if (-not $state.Process.HasExited) { $state.Process.Kill(); $state.Process.WaitForExit(5000) | Out-Null }
            $state.Process.Dispose()
            $state.Process = $null
        }
        foreach ($path in @($state.RequestPath, $state.ResultPath)) {
            if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
        if ($ActivityPath) {
            Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status 'Cancelled' `
                -Detail 'User cancelled migration creation.' -ExitCode 2
        }
    }.GetNewClosure()
    $state.Cancel.Add_Click({
        & $cancelValidation
        $state.Dialog.Close()
    }.GetNewClosure())
    $dialog.Add_Closing({
        param($sender, $eventArgs)
        if ($state.Process -and -not $state.Process.HasExited) {
            & $cancelValidation
        }
    }.GetNewClosure())
    [void]$dialog.ShowDialog()
    return $state.CreatedName
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCn7qZPdTC/7P65
# VhyxsG7fok7aiTgMvaNQ63gjnBy3UaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAAIkDuFyzVws4sNoqH7VHd
# ZEooRjuV1YCaCacfejgsWDANBgkqhkiG9w0BAQEFAASCAYBzRTlD7ypPTf4sehdq
# WR3Ha+Am8soWFBK0NIg7KLuyDTo5CaIc0wk1vQzaCEj0Gret3DzZq26GhXcA1fqF
# e0b/x2FcyutFCXPwIqvwkpA8R6PYrj6Jqhv9v3rhoPlE+FQZCC5L7SQqmq9U2hRa
# nHQvaZ1zwLlpgvrQYJfZNV/CChjokC3BY1NpuzCrcLnVoiTR2T95HPOn4nBrOLqJ
# l6cxsWrXPHb6ktYqOevmM7jAWukS/19PfNISzWELKA2+KVwX4NLHXPwzuj2RYLR1
# rVZzN/479VtwFZmbcc5gNGMa9gbE2RwSfmC/w9KQSYj+kKEcYsKw+o8dZO3/4GiG
# ych6CLSFQqW69PY3WSR7v1fgOLeTJS2OyUbTBVYoZeKniVkcGzoo42kuz51WOkhG
# 7ZYzefzj9UVJpQD+Q7oX0NHjhUS1QDVeRHp5rwAEJWvm8HnlI0IHtDbYWoi2Z+Ba
# QSY8bI7HPCWdenu8mW1/CJIIQz4BnbWMuYr3UR/uoeJ2bkQ=
# SIG # End signature block
