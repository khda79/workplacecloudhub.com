#Requires -Version 7.4
<#
.SYNOPSIS
Renders a loading shield before synchronous local GUI reads.
.VERSION
1.0.0
#>
function script:Invoke-SmartM365SharePointLoading {
    param([Parameter(Mandatory)]$Window,
          [Parameter(Mandatory)][string]$Message,
          [Parameter(Mandatory)][scriptblock]$Action)
    $overlay = $Window.FindName('panelGuiLoading')
    $detail = $Window.FindName('lblGuiLoadingDetail')
    if (-not $overlay -or -not $detail) { throw 'Loading controls are missing.' }
    # A nested request must not hide the shield belonging to its caller.
    if ($overlay.Visibility -eq 'Visible') { & $Action; return }
    $previousCursor = $Window.Cursor
    $previousFocus = [System.Windows.Input.Keyboard]::FocusedElement
    try {
        $detail.Text = $Message
        $overlay.Visibility = 'Visible'
        $Window.Cursor = [System.Windows.Input.Cursors]::Wait
        [void]$overlay.Focus()
        $Window.UpdateLayout()
        # Render priority flushes layout/painting without pumping input or timers.
        $Window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        & $Action
    }
    finally {
        $overlay.Visibility = 'Collapsed'
        $Window.Cursor = $previousCursor
        if ($Window.IsActive -and $previousFocus) {
            [void][System.Windows.Input.Keyboard]::Focus($previousFocus)
        }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDIQBFEJp2LgOdL
# Af4JLxB1dk1Hwz5rcNW05DzmGK1MHqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBJZhgsF2dUgmwi8SEuwFEi
# RzQHTXUrHBrj9wigRQCpGzANBgkqhkiG9w0BAQEFAASCAYBV7J2M7Xk2Mm2PykL+
# bZmK7FooZBPXN8D9i7pdFNF1J9HynwvJcXcqmWwH+AFJNaNjd4gojpariRe4uhKW
# jlTjCnmWB4XMIlFYkLKultVF+1EhFX5tS/E/B0ZoDGmd69uSFVZA8Y5siZoiTwS8
# uHZuuvpH9lkPc8arsU8zmq9wEk753SyaEt6/L40uMesUJM0EPylTaWbJmBSUmVjc
# OoX99T78/LkKuFaIQwTGhfed6t7fMz8Yzwp1spToMh8/HWglqIPupQGIl4ARCjle
# Q6fKaSQoAVYGHmR871XCywbt1zUp6ecQ031dXe9J8qlAogPvrzBhdeiTRcWIa6bD
# mQ6hiWW3Kg/xRTMVsPrYLNASKDrLP/Tgt1hpC8LJMpT7PsCEPxyQ29WtmZG46tp8
# a2s0CNUXkU26vCOqJusbvIoxMtC4rcewNnnFNO6gwG9qwx3OYO2bNTdFSpvOqOdt
# pRCcKm9KE4uAVjBR7hABL7yQuU6X7hqcLxtB8p07yvi2ats=
# SIG # End signature block
