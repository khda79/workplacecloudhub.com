#Requires -Version 7.4
<#
.SYNOPSIS
    Tests diagnostic Copy message controls with synthetic UI objects only.
    Does not execute the GUI, access the real clipboard or read migration data.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$guiPath = Join-Path $project 'SmartM365-SharePointMigration-GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
$script:checks = 0
function Assert-Copy([bool]$Condition, [string]$Message) {
    $script:checks++
    if (-not $Condition) { throw $Message }
}
Assert-Copy ($parseErrors.Count -eq 0) 'GUI source does not parse.'

$xamlNodes = @($ast.FindAll({param($node)
    ($node -is [Management.Automation.Language.StringConstantExpressionAst] -or
     $node -is [Management.Automation.Language.ExpandableStringExpressionAst]) -and
    $node.Value -match 'x:Name="btnDiagCopyMessage"'
}, $true))
Assert-Copy ($xamlNodes.Count -eq 1) 'Copy message must occur in exactly one XAML literal.'
$xml = [xml]$xamlNodes[0].Value
$ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
$button = @($xml.SelectNodes('//*[@x:Name="btnDiagCopyMessage"]', $ns))
Assert-Copy ($button.Count -eq 1 -and $button[0].GetAttribute('IsEnabled') -eq 'False') 'Copy message must start disabled.'
Assert-Copy ($button[0].GetAttribute('Content') -eq 'Copy message') 'Button label changed.'
$names = @($xml.SelectNodes('//*[@x:Name]', $ns) | ForEach-Object { $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') })
Assert-Copy (@($names | Group-Object | Where-Object Count -gt 1).Count -eq 0) 'XAML has duplicate control names.'
$rowGrids = @($xml.SelectNodes('//*[@x:Name="gridDiagRows"]', $ns))
Assert-Copy ($rowGrids.Count -eq 1 -and $rowGrids[0].GetAttribute('SelectionMode') -eq 'Single') 'Clipboard source must be the single selected raw report row.'
$guiText = Get-Content -LiteralPath $guiPath -Raw
Assert-Copy ($guiText.Contains('$btnDiagCopyMessage = ctrl ''btnDiagCopyMessage''')) 'XAML button is not resolved to its GUI variable.'
$readme = Get-Content -LiteralPath (Join-Path $project 'README.md') -Raw
Assert-Copy ($readme.Contains('click **Copy message** to copy its full message')) 'Copy message documentation is missing.'

function Get-EventBody([string]$Variable, [string]$Event) {
    $matches = @($ast.FindAll({param($node)
        $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Expression.Extent.Text -eq $Variable -and $node.Member.Extent.Text -eq $Event
    }, $true))
    Assert-Copy ($matches.Count -eq 1) "Expected one $Variable.$Event handler."
    $text = $matches[0].Arguments[0].ScriptBlock.Extent.Text.Trim()
    $text.Substring(1, $text.Length - 2)
}
$selectionBody = Get-EventBody '$gridDiagRows' 'Add_SelectionChanged'
$clickBody = Get-EventBody '$btnDiagCopyMessage' 'Add_Click'
$clipboardCall = '[System.Windows.Clipboard]::SetText([string]$row.Message)'
Assert-Copy ([regex]::Matches($clickBody, [regex]::Escape($clipboardCall)).Count -eq 1) 'Copy handler must copy the complete selected Message once.'
# Replace only the external API boundary in the in-memory test delegate.
$syntheticClickBody = $clickBody.Replace($clipboardCall, 'Set-SyntheticClipboard ([string]$row.Message)')
Assert-Copy ($syntheticClickBody -notmatch 'System\.Windows\.Clipboard') 'A real clipboard call remains in the synthetic delegate.'
$selectionHandler = [scriptblock]::Create($selectionBody)
$copyHandler = [scriptblock]::Create($syntheticClickBody)

$script:clipboardCalls = [Collections.Generic.List[string]]::new()
$script:clipboardFails = $false
function Set-SyntheticClipboard([string]$Text) {
    if ($script:clipboardFails) { throw 'Synthetic clipboard busy' }
    $script:clipboardCalls.Add($Text)
}
function Reset-SyntheticUI {
    $script:btnDiagCopyMessage = [pscustomobject]@{IsEnabled = $false; Content = 'Copy message'; ToolTip = ''}
    $script:gridDiagRows = [pscustomobject]@{SelectedItem = $null; ItemsSource = $null}
    $script:txtDiagRaw = [pscustomobject]@{Text = ''}
    $script:clipboardCalls.Clear()
    $script:clipboardFails = $false
}
foreach ($message in @("First line`r`nSecond line`nThird line", 'Échec — accès refusé 日本語', ('Synthetic message ' * 4000), '   ')) {
    Reset-SyntheticUI
    $gridDiagRows.SelectedItem = [pscustomobject]@{Message = $message; ItemName = 'synthetic-item'}
    & $selectionHandler
    Assert-Copy $btnDiagCopyMessage.IsEnabled 'A nonempty selected message must enable Copy message.'
    Assert-Copy (($txtDiagRaw.Text | ConvertFrom-Json).Message -ceq $message) 'Selected-row preview changed the message.'
    & $copyHandler
    Assert-Copy ($script:clipboardCalls.Count -eq 1 -and $script:clipboardCalls[0] -ceq $message) 'Full message, Unicode or line breaks were truncated or modified.'
    Assert-Copy ($btnDiagCopyMessage.Content -eq 'Copied' -and $btnDiagCopyMessage.ToolTip -eq 'Message copied to clipboard') 'Copy success feedback is missing.'
    $gridDiagRows.SelectedItem = $null
    & $selectionHandler
    Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $btnDiagCopyMessage.Content -eq 'Copy message' -and $txtDiagRaw.Text -eq '') 'Clearing the selection left stale preview or success feedback.'
    & $copyHandler
    Assert-Copy ($script:clipboardCalls.Count -eq 1) 'A cleared selection copied stale text.'
}
foreach ($row in @($null, [pscustomobject]@{Message = $null}, [pscustomobject]@{Message = ''})) {
    Reset-SyntheticUI
    $gridDiagRows.SelectedItem = $row
    & $selectionHandler
    & $copyHandler
    Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $script:clipboardCalls.Count -eq 0) 'Missing selection or empty message was copied.'
}
Reset-SyntheticUI
$gridDiagRows.SelectedItem = [pscustomobject]@{Message = 'synthetic retry'}
& $selectionHandler
$script:clipboardFails = $true
& $copyHandler
Assert-Copy ($btnDiagCopyMessage.Content -eq 'Copy failed' -and $btnDiagCopyMessage.ToolTip -eq 'Synthetic clipboard busy') 'Clipboard failure was not contained and explained.'
Assert-Copy ($btnDiagCopyMessage.IsEnabled -and $script:clipboardCalls.Count -eq 0) 'Clipboard failure lost the selection or claimed success.'
$script:clipboardFails = $false
& $copyHandler
Assert-Copy ($btnDiagCopyMessage.Content -eq 'Copied' -and $script:clipboardCalls[0] -ceq 'synthetic retry') 'Copy retry did not recover.'

# Extract only pure UI-state functions; never execute the GUI entry script.
foreach ($name in @('Refresh-DiagnosticPatterns', 'Refresh-DiagnosticRows', 'Clear-DiagnosticResult')) {
    $definition = $ast.Find({param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    Assert-Copy ($null -ne $definition) "Missing diagnostic reset function: $name"
    . ([scriptblock]::Create($definition.Extent.Text))
    $controlNames = @($definition.FindAll({param($node)
        $node -is [Management.Automation.Language.VariableExpressionAst] -and
        $node.VariablePath.UserPath -match '^(btn|cmb|lbl|grid|txt|panel)'
    }, $true) | ForEach-Object { $_.VariablePath.UserPath } | Sort-Object -Unique)
    foreach ($controlName in $controlNames) {
        Set-Variable -Scope Script -Name $controlName -Value ([pscustomobject]@{
            IsEnabled = $true; Content = 'Copied'; ToolTip = ''; Text = ''; Visibility = 'Visible'
            ItemsSource = $null; SelectedItem = $null; SelectedIndex = -1; Items = [Collections.Generic.List[object]]::new()
        })
    }
}
$script:DiagSummary = [pscustomobject]@{Patterns = @()}
$script:DiagRows = @()
$cmbDiagStatus.SelectedItem = [pscustomobject]@{Content = 'All statuses'}
Refresh-DiagnosticPatterns
Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $btnDiagCopyMessage.Content -eq 'Copy message') 'Pattern refresh left Copy message enabled.'
$btnDiagCopyMessage.IsEnabled = $true; $btnDiagCopyMessage.Content = 'Copied'
$gridDiagPatterns.SelectedItem = $null
Refresh-DiagnosticRows
Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $btnDiagCopyMessage.Content -eq 'Copy message') 'Missing pattern did not clear Copy message feedback.'
$btnDiagCopyMessage.IsEnabled = $true; $btnDiagCopyMessage.Content = 'Copied'
$gridDiagPatterns.SelectedItem = [pscustomobject]@{PatternKey = 'synthetic-pattern'}
$script:DiagRows = @([pscustomobject]@{PatternKey = 'synthetic-pattern'; ItemName = 'synthetic-item'; Message = 'synthetic match'; SourceUrl = ''; DestinationUrl = ''})
$txtDiagRowFilter.Text = 'match'
Refresh-DiagnosticRows
Assert-Copy (@($gridDiagRows.ItemsSource).Count -eq 1) 'Raw-row filtering regressed.'
Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $btnDiagCopyMessage.Content -eq 'Copy message') 'Raw-row refresh left stale copy feedback.'
$btnDiagCopyMessage.IsEnabled = $true; $btnDiagCopyMessage.Content = 'Copied'
Clear-DiagnosticResult
Assert-Copy (-not $btnDiagCopyMessage.IsEnabled -and $btnDiagCopyMessage.Content -eq 'Copy message') 'Changing migration did not reset Copy message.'
Assert-Copy ($script:DiagRows.Count -eq 0 -and $txtDiagRaw.Text -eq '') 'Diagnostic clear left stale raw messages.'
[pscustomobject]@{Passed = $true; Checks = $script:checks; Source = 'Synthetic UI objects only'; GUIExecuted = $false; RealClipboardAccess = $false; ServerOperations = $false; ClientDataRead = $false; OutputFilesWritten = 0} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOnvkXjkZ+lVdA
# rx9WbfqcPHsGzxG51SzQeod63hUaxqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPtegKd3LyE6El+wKqiEDKBr/RdPTc0hiSzDssXTtqatMA0GCSqG
# SIb3DQEBAQUABIIBgI5Gr/2zrFeK1v7pLhf5GUABg+9BIDeIO9yrPBt5mKfhYx3l
# OAJwIIeK2ONigOtdPWI1el8ynJYAitP98Ihp10c1xcUnWP8fqMYQiMBGABc12kUx
# Ey3W/ZqtLohZ7m8t3vi2h81i8Uf3aLhutTuPTKkcRk9C3K2TCyVtkTdRuJOHfD98
# R87Wc/jfdSff8iUV9c3iJC9dP6bKWTyxg5QwsKjJcm9XV4ADH1F9izmGjaMElnZ4
# k4NFpneHdKSOPFt7D9KrZ2Fw6BKx4Hk+kOtjH2Cdw+Zph+6wxCpfHGAmeNR37sRS
# JUKrsuFj3tlPU2ArkKvcoii/c94VxVm4yROsaCk+qjm6iRJQAvorKC0y1+9bJ0Ac
# TL2Z3+B4WIqoW/9aESzT9KBfdINo7n6xa+A9/tHfowq+5/2ALS5SLHADTcKWhAbl
# oeKeifxw0EgbYovXIkOZDnoNkA8A3cifN6JLk4f/O+QQws0zQXJyYw8xid70LQIL
# ApmpDYkffQB10qrm6aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxMDEx
# MzhaMC8GCSqGSIb3DQEJBDEiBCBMu1Kb6N+nr0uKZPYVpz8QXl4UOL6T55R+2/mP
# 7BT9KzANBgkqhkiG9w0BAQEFAASCAgC1TqMADSFvtVHSnB/QXSfR7FISMDOEUtvw
# SGe75uIMhjxzPb5JBiRSUqahSY0MtWOv6+f1WHzpsgrXvZ91d+Q/tF8722uYxU6P
# msY1ARyeoFoBk8aw9IzRoHf/NRYxOzO0v+YuRwPm9UXR5uFHrG4XhN46X4/RpbqA
# 4aKDp1Jf4up0KQCTlGN7ZC+8owJL9aGP90eTLVtUb5UZ3JDzWXrvg6jwjlg1lula
# T4eZ8LNAVgQ58blz2l0jCS9fES9+mh96WMFo+tJ3nd60AYLc7S8niB5ix8AsZqRx
# xZchVrjIF+Ao2QTOXspSRyKofN2PTVOIuIFNb1XSnetevo0CeuKxajOD4yWECGvN
# QhwZo2X3v9ZLE8UzxAUJbbqLA85s3jIUl3tGcUMj12UYT0SIpOU7DXQXwDnsimez
# 8cwPWBg2vMIKoBzMHc3nXS47WJdd/pE8/QTGOjkUIhqcbyUzn7m+UAZl8uYIuHHv
# z/QuldL6ZDAXcBelQGgI6KZb6QbbrQcnvomDNHqmMYW8LnXr3fgjqx/L1Zf0gyNi
# gwgIJqpUhnr486MFrdWXY2oBL2lkizBbqHJdfUVR3r2uk2HCcyt2t6WsL88TzcIv
# 05oigV+J95o6R+x6k/54tQSEt2ywPNCgIzx1cPvMCwoW8AwlmblPYqDLD8SxRCBq
# uXLFgnVdSw==
# SIG # End signature block
