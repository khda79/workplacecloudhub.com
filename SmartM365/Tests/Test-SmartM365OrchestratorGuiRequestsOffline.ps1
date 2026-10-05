#Requires -Version 7.0
<#
.SYNOPSIS
Offline WPF request-selection and all-request cancellation tests. No tenant or live shared writes.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$path = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator-GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'GUI parsing failed.' }
foreach ($name in @('ConvertFrom-OrchestratorGuiXaml', 'Update-SelectedPipelineRequest', 'Refresh-RequestsView', 'Stop-AllPipelineRequests')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Missing tested GUI function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
$xamlAssignment = $ast.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml' }, $true)
$window = ConvertFrom-OrchestratorGuiXaml -Text $xamlAssignment.Right.Expression.Value
$script:Controls = @{}
foreach ($name in @('RequestsGrid', 'RequestJobsGrid', 'CancelRequestButton', 'CancelAllRequestsButton', 'CancellationReasonBox')) {
    $script:Controls[$name] = $window.FindName($name)
    if (-not $script:Controls[$name]) { throw "Missing XAML control: $name" }
}
$script:Cases = 0
function Assert-Case { param([bool]$Condition, [string]$Message) $script:Cases++; if (-not $Condition) { throw $Message } }
$script:Failures = @()
function Write-GuiException { param($Context, $ErrorRecord) $script:Failures += "$Context : $($ErrorRecord.Exception.Message)" }
$script:RequestRows = @()
$script:RecentRows = @()
function Get-SmartM365OrchestratorRecentPipelineRuns { param($SharedDataFolderPath, $Count) $script:RecentRows }
$script:SharedDataFolderPath = 'synthetic-unused'
$selectionEvent = $ast.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
    $node.Expression.Extent.Text -eq '$script:Controls.RequestsGrid' -and $node.Member.Value -eq 'Add_SelectionChanged' }, $true)
if ($null -eq $selectionEvent) { throw 'Requests selection handler is missing.' }
. ([scriptblock]::Create($selectionEvent.Extent.Text))
try {
    $rows = @(foreach ($count in 0, 1, 2, 4) {
        [pscustomobject]@{ BatchId = "Batch$count"; Status = 'Running'; JobRows = @(foreach ($i in 1..4 | Select-Object -First $count) {
            [pscustomobject]@{ JobName = "Job$i"; Status = 'Pending'; OwnerServer = ''; UpdatedAtUtc = ''; Detail = '' }
        }) }
    })
    $script:Controls.RequestsGrid.ItemsSource = $rows
    foreach ($row in $rows + $rows[1] + $rows[2]) {
        $script:Controls.RequestsGrid.SelectedItem = $row
        Assert-Case ($script:Controls.RequestJobsGrid.ItemsSource -is [Collections.IEnumerable]) 'WPF ItemsSource was unwrapped.'
        Assert-Case ($script:Controls.RequestJobsGrid.Items.Count -eq $row.JobRows.Count) 'Job row count changed.'
        Assert-Case $script:Controls.CancelRequestButton.IsEnabled 'Active request cancellation disabled.'
    }
    $script:Controls.RequestJobsGrid.SelectedItems.Add($rows[2].JobRows[0]) | Out-Null
    $script:Controls.RequestJobsGrid.SelectedItems.Add($rows[2].JobRows[1]) | Out-Null
    Assert-Case ($script:Controls.RequestJobsGrid.SelectedItems.Count -eq 2) 'Two-job selection failed.'
    Assert-Case ($script:Failures.Count -eq 0) 'A selection event failed.'
    $script:Controls.RequestsGrid.SelectedItem = $null
    Update-SelectedPipelineRequest
    Assert-Case ($script:Controls.RequestJobsGrid.Items.Count -eq 0 -and -not $script:Controls.CancelRequestButton.IsEnabled) 'Empty selection retained jobs or enabled cancellation.'
    $script:RecentRows = @($rows[1])
    Refresh-RequestsView
    Assert-Case $script:Controls.CancelAllRequestsButton.IsEnabled 'All-request cancellation not enabled for an active batch.'
    $script:Controls.RequestsGrid.SelectedItem = $rows[1]
    Refresh-RequestsView
    Assert-Case ($script:Controls.RequestsGrid.SelectedItem.BatchId -eq 'Batch1' -and $script:Controls.RequestJobsGrid.Items.Count -eq 1) 'Refresh lost singleton selection.'
    $script:RecentRows = @()
    Refresh-RequestsView
    Assert-Case (-not $script:Controls.CancelAllRequestsButton.IsEnabled -and $script:Controls.RequestJobsGrid.Items.Count -eq 0) 'Empty refresh left stale controls.'
    # Test refusal paths without opening a confirmation dialog or writing shared state.
    $script:Tenant = 'test'
    $script:ActiveRows = @()
    function Get-SmartM365OrchestratorActivePipelineRuns { param($SharedDataFolderPath) $script:ActiveRows }
    $script:PreviewCalls = @()
    function Stop-SmartM365OrchestratorPipelineRequest {
        param($SharedDataFolderPath, $BatchId, $Tenant, $Reason, [switch]$ValidateOnly)
        if (-not $ValidateOnly) { throw 'Unexpected cancellation write in offline refusal test.' }
        $script:PreviewCalls += $BatchId
        throw 'Synthetic readiness refusal'
    }
    $script:Controls.CancellationReasonBox.Text = ' '
    $message = ''
    try { Stop-AllPipelineRequests } catch { $message = $_.Exception.Message }
    Assert-Case ($message -eq 'Enter a cancellation reason.') 'Missing reason was not refused.'
    $script:Controls.CancellationReasonBox.Text = 'Synthetic test'
    $message = ''
    try { Stop-AllPipelineRequests } catch { $message = $_.Exception.Message }
    Assert-Case ($message -eq 'No active pipeline requests remain.') 'No-active-request case was not refused.'
    $script:ActiveRows = @([pscustomobject]@{ BatchId = 'BatchOutsideRecentHistory' })
    $message = ''
    try { Stop-AllPipelineRequests } catch { $message = $_.Exception.Message }
    Assert-Case ($message -eq 'Synthetic readiness refusal' -and $script:PreviewCalls -contains 'BatchOutsideRecentHistory') 'All-request validation depended on recent history or bypassed readiness.'
    $script:ActiveRows = @([pscustomobject]@{ BatchId = 'BatchA' }, [pscustomobject]@{ BatchId = 'BatchB' })
    $script:PreviewCalls = @(); $script:WriteCalls = @(); $script:ConfirmationCalls = 0
    $script:ConfirmResult = $false; $script:FailBatch = ''
    function Stop-SmartM365OrchestratorPipelineRequest {
        param($SharedDataFolderPath, $BatchId, $Tenant, $Reason, [switch]$ValidateOnly)
        Assert-Case ($Tenant -eq 'test' -and $Reason -eq 'Synthetic test') 'Cancellation lost tenant or reason.'
        if ($ValidateOnly) { $script:PreviewCalls += $BatchId }
        else {
            if ($BatchId -eq $script:FailBatch) { throw 'Synthetic batch changed before publication' }
            $script:WriteCalls += $BatchId
        }
        [pscustomobject]@{ BatchId = $BatchId; OverallStatus = 'Cancelling'; CancelledCount = 2; PendingCount = 1 }
    }
    function Confirm-PipelineCancellation {
        param($Message, $Title)
        $script:ConfirmationCalls++
        Assert-Case ($Message -match 'BatchA' -and $Message -match 'BatchB' -and $Message -match 'Running collectors continue' -and $Message -match 'automatic schedules stay unchanged') 'Confirmation scope or safety text is missing.'
        return $script:ConfirmResult
    }
    $script:Activity = @()
    function Write-GuiActivity { param($Level, $Message) $script:Activity += $Message }
    Stop-AllPipelineRequests
    Assert-Case ($script:ConfirmationCalls -eq 1 -and $script:PreviewCalls.Count -eq 2 -and $script:WriteCalls.Count -eq 0) 'Declined confirmation wrote shared state or skipped validation.'
    $script:ConfirmResult = $true; $script:ConfirmationCalls = 0
    $script:PreviewCalls = @(); $script:WriteCalls = @()
    Stop-AllPipelineRequests
    Assert-Case ($script:ConfirmationCalls -eq 1 -and ($script:WriteCalls -join ',') -eq 'BatchA,BatchB' -and $script:Activity.Count -eq 2) 'Bulk cancellation lost a batch, audit record or single confirmation.'
    Assert-Case ($script:Controls.CancellationReasonBox.Text -eq '') 'Successful cancellation did not clear the reason.'
    $script:Controls.CancellationReasonBox.Text = 'Synthetic test'
    $script:WriteCalls = @(); $script:FailBatch = 'BatchB'; $message = ''
    try { Stop-AllPipelineRequests } catch { $message = $_.Exception.Message }
    Assert-Case ($message -eq 'Synthetic batch changed before publication' -and ($script:WriteCalls -join ',') -eq 'BatchA') 'A late publication refusal was ignored or rolled back an audited cancellation.'
    Assert-Case ($script:Controls.CancellationReasonBox.Text -eq 'Synthetic test') 'Partial refusal discarded the operator reason.'
    Assert-Case ($script:Controls.CancelAllRequestsButton.Content -eq 'Cancel All remaining Jobs') 'Bulk cancellation label is missing.'
    # Exercise the real click registration, without publishing any request.
    $script:ClickCalls = 0
    function Stop-AllPipelineRequests { $script:ClickCalls++ }
    $clickEvent = $ast.Find({ param($node) $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Expression.Extent.Text -eq '$script:Controls.CancelAllRequestsButton' -and $node.Member.Value -eq 'Add_Click' }, $true)
    if ($null -eq $clickEvent) { throw 'Bulk cancellation click handler is missing.' }
    . ([scriptblock]::Create($clickEvent.Extent.Text))
    $script:Controls.CancelAllRequestsButton.IsEnabled = $true
    $script:Controls.CancelAllRequestsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    Assert-Case ($script:ClickCalls -eq 1) 'Bulk cancellation button is not wired to the bulk action.'
    $tab = $script:Controls.RequestsGrid.Parent.Parent.Parent
    $tab.IsSelected = $true
    $layoutRoot = $window.Content
    $layoutRoot.Measure([System.Windows.Size]::new(1180, 720))
    $layoutRoot.Arrange([System.Windows.Rect]::new(0, 0, 1180, 720))
    $layoutRoot.UpdateLayout()
    $button = $script:Controls.CancelAllRequestsButton
    $position = $button.TranslatePoint([System.Windows.Point]::new(0, 0), $layoutRoot)
    Assert-Case ($button.ActualWidth -gt 0 -and $position.X + $button.ActualWidth -lt 1180) "Bulk cancellation button is clipped at the minimum window width: X=$($position.X); Width=$($button.ActualWidth)."
    Assert-Case ($script:Failures.Count -eq 0) 'Requests refresh failed.'
    "[{0}] PASS: {1} offline WPF request cases; no collectors, tenant calls or shared control writes." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $script:Cases
}
finally { $window.Close() }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBBs9lsk1ffD7ui
# 4kR9cOeLpmak+99E5dxvrPRUKhqQ76CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPWrvoUfHqlWicDIOKsqTQJCVcpRmcg2rldkXDIeQWOCMA0GCSqG
# SIb3DQEBAQUABIIBgK5L0w4caj2aN4xPOw0MmN9CraVxpLjsVhrnhuYon9xU15vZ
# vUX8SS/Lc66icqWcYCXsFtxIPzd4QmaNqoy/HVI1RbdAGi6ZWeKd0t+7JpAw/j0r
# IOjgiMFvYhiLle2mik3kxS5EI91J+Tus75hC/U1AbGvb3gn2ihUqK8tVHqMrT568
# u54tvhpZgVZDT1973NsmhaCy58XkBHHoU7IDQbWzWtnFjFxfmpnFv/zFYf0nHcV6
# V2dwqOfi5rx64ykQngTPqpZgfIcug53Fv50PNeOFDEA16jKDE3R98/VyCFQPpHTs
# 3jbnAxEfVdTM+rDZtYDOdVOiIDR0/SyMXA+M1VTH3lK+s9QPofpFh86i4pvzRWFH
# rXGIOK+i2qBlUVIapekMtPiQBvc0Uek516/twbzjOgM8m4It2RDg82gaKelozuR3
# Qn04FSVmLy8MuxiJ/DB2FU1VlubjVn10TF6hn3y1gWgLXlW9Ll4C6RUEgD7IXfOH
# 9Qq8+ko80zsKl3Co86GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUxOTM0
# NTFaMC8GCSqGSIb3DQEJBDEiBCAUd5B5ni79BfIQAlEfr4RTi5b/3IRaHoqE431B
# /QpW/jANBgkqhkiG9w0BAQEFAASCAgAKLrGdaBx9CPcJgEz5wv3kkQatI1XY9Xt1
# YpPRcHifz+XWAvdsYLLMNm7tBgJ7/k6cBrePb0bCsx96ptmSKOzddgDn2cBqrTUv
# nMK9dfL/J1K1tOw1E3ztMfOcdmMCMr9eek7HmEO0kb/CRGLO7YE2wInMhZhLtjs4
# YrzDJLbJJs0VW+0SJa+YSTSaV4T/+3WmPY589esxta+MaJGcDjW5uxwx6h5YOVi4
# 9L/6rc5+CY9IKbYHJtKGrqI23pIqwBkbNrqLkGtbLvEdL+in672lPo3jQuv1AbnE
# uctJ2IYGiUTPmjz8UBXq+qylHf/crgKp6PJfdWgaCh3B7c6CpkQQ2B37M7X35WTu
# HTE6j0SkwKq6aw7KSXFfQW27rDPGXfYh6pZ6j+P+ey4LaQhpq3Fq20ky6JgcRXkw
# rLZphJ1tANVRgsfWmKCmYCrBo8VYjUkNdrIo150klbl80Wd8HuhwP9bJy8rDv8CW
# ePMdbIo4+TqAeGJnLhyxy66b1OLqkKfFx+SYIV64wB1toSbt9v33G281ATnkMT/X
# mBmY9uAhpye3VnjrAyQFcZh3L/tggnH5y5Pe66ij1m+WdBD1N5pSsK4KAeom/jWS
# jObL3D/BpJ+G6gy77EINps4nr1XOWNAaD/An3lVCjxD+/uEbQ2rUsJgVVpaoy+zV
# zKlN8H+05A==
# SIG # End signature block
