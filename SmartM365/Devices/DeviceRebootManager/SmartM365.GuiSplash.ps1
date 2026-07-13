function Initialize-SmartM365NativeWindowApi {
    if ('SmartM365.NativeWindow' -as [type]) { return }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace SmartM365 {
    public static class NativeWindow {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
'@
}

function Set-SmartM365WpfWindowVisible {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Window)

    try {
        Initialize-SmartM365NativeWindowApi
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        if ($helper.Handle -ne [IntPtr]::Zero) {
            [void][SmartM365.NativeWindow]::ShowWindow($helper.Handle, 5)
            [void][SmartM365.NativeWindow]::SetForegroundWindow($helper.Handle)
        }
        $Window.Visibility = [System.Windows.Visibility]::Visible
        [void]$Window.Activate()
    }
    catch {}
}

function Start-SmartM365GuiSplash {
    [CmdletBinding()]
    param(
        [ValidateSet('Wpf')]
        [string]$Framework = 'Wpf',

        [string]$ProductName = 'SmartM365',

        [int]$MinimumDurationMs = 6000,

        [string]$BadgeText = 'WORKPLACECLOUDHUB.COM',

        [string]$Subtitle = 'Powered by WorkplaceCloudHub',

        [string]$LogoPath = '',

        [string]$WindowIconPath = '',

        [bool]$ShowSiteUrl = $true
    )

    $siteUrl = 'https://workplacecloudhub.com'
    $startedAt = Get-Date

    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName WindowsBase

    $accentBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212))
    $inkBrush = [System.Windows.Media.Brushes]::White
    $mutedBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(214, 233, 250))
    $linkBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(188, 231, 255))
    $borderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(118, 184, 235))
    $surfaceBrush = [System.Windows.Media.LinearGradientBrush]::new()
    $surfaceBrush.StartPoint = [System.Windows.Point]::new(0, 0)
    $surfaceBrush.EndPoint = [System.Windows.Point]::new(1, 1)
    [void]$surfaceBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Color]::FromRgb(5, 32, 66), 0.0))
    [void]$surfaceBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Color]::FromRgb(0, 91, 159), 0.58))
    [void]$surfaceBrush.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Color]::FromRgb(0, 145, 204), 1.0))

    $workplaceLogoPath = $null
    $resolvedWindowIconPath = $null
    if (-not [string]::IsNullOrWhiteSpace($LogoPath)) {
        $candidate = if ([System.IO.Path]::IsPathRooted($LogoPath)) { $LogoPath } else { Join-Path -Path $PSScriptRoot -ChildPath $LogoPath }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $workplaceLogoPath = $candidate
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($WindowIconPath)) {
        $candidate = if ([System.IO.Path]::IsPathRooted($WindowIconPath)) { $WindowIconPath } else { Join-Path -Path $PSScriptRoot -ChildPath $WindowIconPath }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $resolvedWindowIconPath = $candidate
        }
    }

    $logoSearchRoot = $PSScriptRoot
    while ($logoSearchRoot) {
        $candidate = Join-Path -Path $logoSearchRoot -ChildPath 'WorkplaceCloudHub-lockup-WPF.png'
        if (-not $workplaceLogoPath -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            $workplaceLogoPath = $candidate
        }

        $iconCandidate = Join-Path -Path $logoSearchRoot -ChildPath 'WorkplaceCloudHub.ico'
        if (-not $resolvedWindowIconPath -and (Test-Path -LiteralPath $iconCandidate -PathType Leaf)) {
            $resolvedWindowIconPath = $iconCandidate
        }

        if ($workplaceLogoPath -and $resolvedWindowIconPath) { break }

        $parent = Split-Path -Parent $logoSearchRoot
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $logoSearchRoot) {
            break
        }

        $logoSearchRoot = $parent
    }

    if (-not $workplaceLogoPath) {
        $smartM365Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        foreach ($relativePath in @(
            'Devices\IntuneHybridJoinToolkit\WorkplaceCloudHub-lockup-WPF.png',
            'Devices\Windows11UpgradeToolkit\WorkplaceCloudHub-lockup-WPF.png',
            'Devices\DeviceRegistrationTool\WorkplaceCloudHub-lockup-WPF.png',
            'Devices\EndpointDiagnosticsAnalyzer\WorkplaceCloudHub-lockup-WPF.png'
        )) {
            $candidate = Join-Path $smartM365Root $relativePath
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $workplaceLogoPath = $candidate
                break
            }
        }
    }

    if (-not $resolvedWindowIconPath) {
        $smartM365Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        foreach ($relativePath in @(
            'Devices\IntuneHybridJoinToolkit\WorkplaceCloudHub.ico',
            'Devices\Windows11UpgradeToolkit\WorkplaceCloudHub.ico',
            'Devices\DeviceRegistrationTool\WorkplaceCloudHub.ico',
            'Devices\EndpointDiagnosticsAnalyzer\WorkplaceCloudHub.ico',
            'Devices\DeviceRebootManager\WorkplaceCloudHub.ico',
            'Intune\Remediation\GUI\WorkplaceCloudHub.ico',
            'WorkplaceCloudHub.ico'
        )) {
            $candidate = Join-Path $smartM365Root $relativePath
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $resolvedWindowIconPath = $candidate
                break
            }
        }
    }

    $rootGrid = [System.Windows.Controls.Grid]::new()
    $rootGrid.Margin = [System.Windows.Thickness]::new(26, 22, 26, 20)
    [void]$rootGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
    $logoColumn = [System.Windows.Controls.ColumnDefinition]::new()
    $logoColumn.Width = [System.Windows.GridLength]::new(210)
    [void]$rootGrid.ColumnDefinitions.Add($logoColumn)

    $panel = [System.Windows.Controls.StackPanel]::new()
    [System.Windows.Controls.Grid]::SetColumn($panel, 0)
    [void]$rootGrid.Children.Add($panel)

    $logoStack = [System.Windows.Controls.StackPanel]::new()
    $logoStack.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $logoStack.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $logoStack.Margin = [System.Windows.Thickness]::new(18, 0, 0, 0)
    [System.Windows.Controls.Grid]::SetColumn($logoStack, 1)
    [void]$rootGrid.Children.Add($logoStack)

    if ($workplaceLogoPath) {
        try {
            $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bitmap.BeginInit()
            $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmap.UriSource = [Uri]$workplaceLogoPath
            $bitmap.EndInit()
            $bitmap.Freeze()

            $workplaceLogo = [System.Windows.Controls.Image]::new()
            $workplaceLogo.Source = $bitmap
            $workplaceLogo.Width = 178
            $workplaceLogo.Height = 178
            $workplaceLogo.Stretch = [System.Windows.Media.Stretch]::Uniform
            $workplaceLogo.SnapsToDevicePixels = $true
            [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($workplaceLogo, [System.Windows.Media.BitmapScalingMode]::HighQuality)
            [void]$logoStack.Children.Add($workplaceLogo)
        }
        catch {}
    }

    $badge = [System.Windows.Controls.TextBlock]::new()
    $badge.Text = $BadgeText
    $badge.Foreground = $linkBrush
    $badge.FontSize = 11
    $badge.FontWeight = [System.Windows.FontWeights]::SemiBold
    if (-not [string]::IsNullOrWhiteSpace($BadgeText)) {
        [void]$panel.Children.Add($badge)
    }

    $title = [System.Windows.Controls.TextBlock]::new()
    $title.Text = $ProductName
    $title.Foreground = $inkBrush
    $title.FontSize = 24
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.Margin = [System.Windows.Thickness]::new(0, 8, 0, 4)
    [void]$panel.Children.Add($title)

    $subtitleTextBlock = [System.Windows.Controls.TextBlock]::new()
    $subtitleTextBlock.Text = $Subtitle
    $subtitleTextBlock.Foreground = $mutedBrush
    $subtitleTextBlock.FontSize = 13
    $subtitleTextBlock.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    if (-not [string]::IsNullOrWhiteSpace($Subtitle)) {
        [void]$panel.Children.Add($subtitleTextBlock)
    }

    if ($ShowSiteUrl) {
        foreach ($url in @($siteUrl)) {
            $line = [System.Windows.Controls.TextBlock]::new()
            $line.Text = $url
            $line.Foreground = $linkBrush
            $line.FontSize = 12
            $line.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
            [void]$panel.Children.Add($line)
        }
    }

    $border = [System.Windows.Controls.Border]::new()
    $border.Background = $surfaceBrush
    $border.BorderBrush = $borderBrush
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $border.Child = $rootGrid

    $window = [System.Windows.Window]::new()
    $window.Title = 'WorkplaceCloudHub'
    $window.Width = 720
    $window.Height = 292
    $window.WindowStartupLocation = 'CenterScreen'
    $window.ResizeMode = 'NoResize'
    $window.ShowInTaskbar = $false
    $window.Topmost = $true
    $window.UseLayoutRounding = $true
    $window.SnapsToDevicePixels = $true
    $window.Content = $border
    if ($resolvedWindowIconPath) {
        try {
            $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$resolvedWindowIconPath)
        }
        catch {}
    }
    $window.Show()
    Set-SmartM365WpfWindowVisible -Window $window

    try {
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    }
    catch {}

    return [pscustomobject]@{
        Framework = $Framework
        Window = $window
        StartedAt = $startedAt
        MinimumDurationMs = $MinimumDurationMs
    }
}

function Wait-SmartM365GuiSplashMinimum {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Splash)

    while ($true) {
        $elapsedMs = [int]((Get-Date) - [datetime]$Splash.StartedAt).TotalMilliseconds
        $remainingMs = [int]$Splash.MinimumDurationMs - $elapsedMs
        if ($remainingMs -le 0) {
            break
        }

        try {
            if ($Splash.Window -and $Splash.Window.Dispatcher) {
                $Splash.Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
            }
        }
        catch {}

        Start-Sleep -Milliseconds ([math]::Min(100, $remainingMs))
    }
}

function Hide-SmartM365GuiSplash {
    [CmdletBinding()]
    param([AllowNull()]$Splash)

    if ($null -eq $Splash -or $null -eq $Splash.Window) { return }
    Wait-SmartM365GuiSplashMinimum -Splash $Splash

    try {
        $Splash.Window.Topmost = $false
        $Splash.Window.Hide()
    }
    catch {}
}

function Close-SmartM365GuiSplash {
    [CmdletBinding()]
    param([AllowNull()]$Splash)

    if ($null -eq $Splash -or $null -eq $Splash.Window) { return }
    Wait-SmartM365GuiSplashMinimum -Splash $Splash

    try {
        $Splash.Window.Close()
    }
    catch {}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDj4fU0+Ki+Y4jg
# Mb/hGc0OzE+Wka2A8u6MrN8ugBAGNqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJh6u4VD+3o2pg47hyFWuVohHUU1oy9boH3U9CKwELSCMA0GCSqG
# SIb3DQEBAQUABIIBgBlV6FC8wQtc+z7w74iwBi4lp3jaMuhRlPUrGLG3Vr0NF04U
# 295itxcAet4MRczflglT4TjQNlLL3lRCIA8zXNt8TVmfN/ku3+kuXrvsWE/1ZS4Y
# 4SMmUsUQ0hpw7A7MfEt80YtPhOX9deGpFw3uFP0yGYQmgphsw4XPZnIB2qKEdY4Q
# Vjo6l1CMlCN7csAeSrMitf02kQbypP+MSvm6A3fmL7ef61Z6p22TssGZTkzTLEvM
# +Ap/pTJLXWZko5415dvcBAmb/nqfHr9CbtCrjH4eeLfOrCCrAcS3ndt4NhjcrFUl
# Is5r5SuHdo3WTPK79bSZUkISHWph34XNjo1FZu9dAOKOpt8d/oq+KEkNiu6KOWAG
# kG7otKGxyPpQqP2H9crhxgXuK1PCF38y7o5l2g9BftUh/D/UDgAxCS3cBDaRc/La
# nK5KfW3HRC2M2yrGhGX8O3oFb7eaNaTdRQrEE4pUSUajHRnyAE+O4P64E1J+Kcoh
# jimeeR0NwtP4e2shr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MjhaMC8GCSqGSIb3DQEJBDEiBCBsEVHYhL4AT/Di2ID4bGz2RhWMaCSfZvNqCRri
# DFCB+zANBgkqhkiG9w0BAQEFAASCAgBwBqzfJFuElqR2YgGjSkxpljYpEtbVnCzN
# duggim+/+fRyKx0K/mLdtnEIiNKaXkIjxGdloG+ajsjCTVo57PMEDDKOzio+bXX+
# NYOAz6MIiu+FlZMP4P7Be4KeFR+Io/apl3fXJYr9FarTDQLXM8V9Pc6uLgJGqHnu
# F67tn5mVknxQgD5YdKsR31yGvE0FhOp9ft0oHGsKmOgGd7MMPeUupM7CuD4piuu2
# 7ylt6w/RXMP9xT93GsFv9/CbGMcIiUfG0U/9P6qqCcRl4sBfolsz6ir7IUDZjiwo
# jfqkAaUSZNvs1BDKDa2Rb9TapByaxx0FpqF1x9v3Gm5aeN9GOgffyDpJ3GXtZCRc
# leX04Mfi1hZ0x9dG6cio3fzoTBspWro3P2FboDGUSKjHodjH4blSBuztl90CVcdV
# FuXErjn5AZg3//v3Uod5jYdiDhlfYxWF05WQ1uXOQ2XT0Z+uGczdEP0aCm5J1Pvg
# CYYn4Pj2Hd1ou1wZQq77aSqXBkdNHeKqkSL3wtns0UYTqTHZD4BXFwlvvsKxZJfj
# Wvoy0g7Ec06kgEU3XUVrW6Nqtq38F6VieKLY+ZFcy6UGkWmzcmRogs7WYC7mj9a7
# rbncaA8NWzanGujOgjuJhRC8dreL/ouHfVwQ6Klvnp1RW5w4Uc9li+ZeF6JjKH+h
# qN0x8fT2wQ==
# SIG # End signature block
