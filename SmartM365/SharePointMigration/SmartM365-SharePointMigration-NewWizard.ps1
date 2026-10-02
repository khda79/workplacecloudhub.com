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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAVnuV/xqtRxtP+
# LIGJihYuF6qcGlzO00xYoxnGQdR7D6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAMyndi1PFUD1RcRxUsRDKGYW0Ulv6ee3ES4Z2gBEY4nMA0GCSqG
# SIb3DQEBAQUABIIBgHHdYYCE7ndlS5mGW+9T+qui9L9DYIHG0ld0fYoMowdLMJcd
# R49FDikI4oBzott20TmwxRuiUYoGPIHaw0JNNngoOCtIcomD/e2v4Cfne3ZYfJ/5
# uvfSMLLXw6incssNygl4sixuwveiEiTvujTuXktCmiy9nKTHIedenVoL3d5B2IA4
# 0Ig5GNTO4AokW33P/3k+F9C4EFmzTP9t6ZAl1qjYfuyL0e0BvdsGu4YbURnKTVAw
# RpaqPIOX6r+3APrVlPctBE+hcffZLGEaBp906hq4tCVMxlkmr6U0cr6wzPPbmoCd
# gJehcCkMpoUMjT3FFxk8dueE6YoVGBMgt3DARp9XmxM6j8yTD93a8Pia+XvInVsk
# eyenoHdCYqir1odv6t9IOQ+Ggzht4YfIR6MQO/bI8z6X+5W9Gndgcq4c6P0uw2lw
# PaOPeWxrYCMQbObN9bOrg6o+nf1r/cnV/UKhWynl/PRd0r5LEQbYRgvUpmE4blLt
# ibds9e0qlqky8BXrvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDIxMjIw
# NTNaMC8GCSqGSIb3DQEJBDEiBCBue7CFiT4Ihf/7WeTshRzw1B29GQ5IJ12DYzRV
# 6X6TyTANBgkqhkiG9w0BAQEFAASCAgCH0gYvRK9QTD5g/ot3RzUEMF1WOz0MDuIU
# m1T2G9AgrFjeZ4LiXxRPfhpEzCOUFdywruu3OhdDsq5/vDpjW0eWaWly57cSvJbg
# +X258FscXxK7Sml46FjUb09yz/X7tyqe6NPVcvVo7zEn3HGlf7DvbDmfozsuZOoN
# 4l0WoOeEl86acc4EBVB0IC4gzYEyxqQoGT1w+asMNyqBBENrWEjBnLuihA029LoR
# Z82Mq8x2n3y9MERcIBEPI9Lz+ENoEqX/M+Dh88HaLxTPiiT1slVUxhdzVy0Opa4Z
# E05+HhOw59phxmCstapS/hv1yb+OQOr0k/+MnHDVM0RNXDJ0rNDVdCwqqLuJ9DYI
# yCBm+gIf6dwMH5ZfrxVhcSSPw9dcZ/7X9a4SH0HW9TCIrLn99mLv4KID3x2g7tXv
# Rtybn6rfU8btAMKbmzbwKC7RQ2jQjBk/owWnEETLyQpwKXZtEOf2agXWyYLJA6Si
# k6CU1zAEPDV+CTOa0UTJ0Wx/dpBoO2sige0RqmVBcHki1VcZswMVKByH+mAgRZmA
# 1PCapGqjEI4KyDW3RG4yT18/97k6utFqqC+/yuwGzIZxJun4Al0R5wT/ESiERonK
# 42u8i8XgRVEW8WZznkyV66aOCaTbMtjn1RJwpIIZ/J5xfqdA+6eEXWCsKM/sIWMZ
# FepdH5nqlg==
# SIG # End signature block
