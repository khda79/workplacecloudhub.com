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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDda3zNe3retFSC
# ZIT8ASREkctL/+ZUODjZ828R9YiAOqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCANQBcTvzh8XEjJ1JUkt2RpPMWkgknmBD3Qfrwji7BK5DANBgkqhkiG9w0B
# AQEFAASCAYBxCJjb/TjJql/kNQwHX9Fw7PBOZqIKmlRFd1F2pnbTQ2XLUICmvCkx
# yvGYx9pkvwrp1rd/WVOKig7s5zFJ0g94Ktqgrt0gsaLsuzv8o6vfC6VYtNTzSm5D
# Drhy9id589xMflv73ohbkBCue5JVDYsgzeYCkoCAE44qsYgyH2o7U2vBziy7B1W2
# NQtLCEppZlkMi+f9nL7CkeIf9VC0UoP5nFmDyrYOvzyeJJXFs1r/wxBUkjWUE2bx
# fIov3qx4ISRUCiKh1pZ+um/27AUq/VqzmDfgphX9Vs/AiZ1imxTKF69aCQO156+n
# 0wo8R+n3hXb0c3MrXzr67selPU+ta8Ovg7MJAbNt1HsBrgZGkzj6jfHAPIWAUNBT
# l+15Dityto1X5Ga0a229PWX9EwGFysBwSpxyUIGXdPXI2ibD523rKm/Plow3qO4N
# e5b+yHM9YYQ4CGWQ8yW8Gmv/WIBj5VAtCJt9tu5lZ0G6ypv8z6quP3RdKWN5Yn5B
# qd6cZMaPLhU=
# SIG # End signature block
