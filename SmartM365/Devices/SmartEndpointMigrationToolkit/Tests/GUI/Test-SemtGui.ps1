<#
.SYNOPSIS
    Offline tests for the Smart Endpoint Migration Toolkit WPF GUI and its demonstration.
.DESCRIPTION
    Static checks of the window markup and launch helpers, then the GUI -ValidateOnly and the
    demo (-Demo -ValidateOnly with simulated button clicks) in child Windows PowerShell 5.1
    STA processes. Nothing is launched: no CMD, Run, inventory, PsExec or endpoint action.
    The technician state is redirected with SEMT_RUNTIME_ROOT_OVERRIDE and the package is
    copied to a temp folder so no GUI state file is written into the source tree.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$sourceRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-GuiTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$packageRoot = Join-Path $tempRoot 'SmartEndpointMigrationToolkit'
$previousOverride = $env:SEMT_RUNTIME_ROOT_OVERRIDE
$previousUpdate = $env:SEMT_GUI_UPDATE_CHECK
try {
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    foreach ($item in @('Assets','Campaigns','Config','Modules','Scripts')) { Copy-Item -LiteralPath (Join-Path $sourceRoot $item) -Destination $packageRoot -Recurse -Force }
    foreach ($item in @('Runs','Waves\HybridJoin','Waves\Windows11Upgrade')) { New-Item -ItemType Directory -Path (Join-Path $packageRoot $item) -Force | Out-Null }
    Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json' -File | Remove-Item -Force
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = Join-Path $tempRoot 'Runtime'
    $env:SEMT_GUI_UPDATE_CHECK = '0'
    $guiPath = Join-Path $packageRoot 'Scripts\SmartM365-EndpointMigration-GUI.ps1'
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)

    #region Static
    Test-Step 'GUI parses without error' { @($errors).Count -eq 0 }
    $assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml' }, $true)
    $xaml = [string]$assignment.Right.Expression.Value
    Test-Step 'window markup is a literal here-string (demo framework contract)' { $assignment.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $xaml -notmatch '\$\(|\$script:' }
    foreach ($name in @('CampaignCombo','WaveCombo','WaveModeCombo','LaunchWaveButton','LaunchAllButton','StopAllButton','SingleComputerText','LaunchSingleButton','NewWaveNameText','CreateWaveButton','AutomaticPreviewButton','AutomaticCreateButton','AutomaticRefreshInventoryButton','RunsGrid','OpenRunReportButton','StopSelectedRunButton','ForceStopSelectedRunButton','OptionsHost','ResetOptionsButton','HeaderLogoImage')) {
        Write-TestResult -Condition ($xaml -match ('x:Name="{0}"' -f $name)) -Message ("control present: {0}" -f $name)
    }
    Test-Step 'no Batch/LOT wording in the interface' { $xaml -notmatch '\b(LOT|LOTs|Lot|Batch|batch)\b' }
    Test-Step 'Preview and Create are separate controls; Create states it never launches' { $xaml -match 'AutomaticPreviewButton' -and $xaml -match 'never launches the Wave' }
    Test-Step 'Launch All is labelled as limited to the selected Campaign in its confirmation' { [IO.File]::ReadAllText($guiPath) -match 'Launch All Waves of this Campaign only' }
    Test-Step 'stop states remote endpoints are reported separately' { [IO.File]::ReadAllText($guiPath) -match 'Remote endpoints may continue and must be verified' }
    $guiText = [IO.File]::ReadAllText($guiPath)
    foreach ($functionName in @('ConvertTo-PowerShellLiteral','Get-AutomaticWaveCommandText','Get-AutomaticFilterSignature','Split-GuiList','Get-GuiRunCommand','Assert-CmdLiteralValue','ConvertTo-CmdArgument','Get-GuiStatePath','New-GuiSelectionFile')) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    $script:PackageRoot = $packageRoot
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    $filters = @{ WaveName = 'WAVE-AUTO-TEST'; Prefixes = @('CTSO-'); Contains = @("O'Brien"); ExcludeStaleAd = $true; LastLogonDays = 30; Source = 'Both'; ExcludeIntunePresent = $true }
    foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
        $engine = Join-Path $packageRoot ("Campaigns\{0}\Scripts\SmartM365-EndpointMigration-{0}-New-AutomaticWave.ps1" -f $campaign)
        $engineAst = [System.Management.Automation.Language.Parser]::ParseFile($engine, [ref]$null, [ref]$null)
        $engineParameters = @($engineAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $preview = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence'
        $create = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot 'C:\Evidence' -Create
        $used = @([regex]::Matches($create, '(?<=\s)-(?<Name>[A-Za-z]+)') | ForEach-Object { $_.Groups['Name'].Value } | Select-Object -Unique)
        $unknown = @($used | Where-Object { $engineParameters -notcontains $_ })
        Write-TestResult -Condition ($unknown.Count -eq 0) -Message ("{0}: Automatic Wave command uses only engine parameters" -f $campaign) -Detail ($unknown -join ', ')
        Write-TestResult -Condition ($preview -notmatch '-Create\b' -and $create -match '-Create$') -Message ("{0}: Preview never passes -Create; Create does" -f $campaign)
        Write-TestResult -Condition ($create -match "'O''Brien'") -Message ("{0}: filter values are quoted as PowerShell literals" -f $campaign)
    }
    Test-Step 'filter signature changes when a filter changes (Create needs a matching Preview)' {
        $a = Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $filters
        $changed = $filters.Clone(); $changed.LastLogonDays = 31
        $a -ne (Get-AutomaticFilterSignature -Campaign HybridJoin -Filters $changed)
    }
    $run = Get-GuiRunCommand -Campaign Windows11Upgrade -WavePath 'C:\Waves\W 1' -RunMode Once -IgnoreRunGuard $true -SelectionPath 'C:\Sel\s.json'
    Test-Step 'Run command uses Windows PowerShell 5.1, the engine, the GUI selection and the run mode' { $run -match 'WindowsPowerShell\\v1\.0\\powershell\.exe' -and $run -match 'SmartM365-Invoke-EndpointMigrationWithPsExec\.ps1' -and $run -match '-GuiSelectionPath' -and $run -match '-RunMode Once' -and $run -match '-IgnoreRunGuard' -and $run -match '"C:\\Waves\\W 1"' }
    Test-Step 'Single PC command passes -ComputerName and no Wave' { $fictionalComputer = 'PC-1'; $c = Get-GuiRunCommand -Campaign HybridJoin -ComputerName $fictionalComputer -RunMode Once -IgnoreRunGuard $false -SelectionPath 'C:\s.json'; $c -match '-ComputerName PC-1' -and $c -notmatch '-WavePath' -and $c -notmatch '-IgnoreRunGuard' -and $c -match '-GuiSelectionPath' }
    $catalog = Import-SemtOptionCatalog -PackageRoot $packageRoot
    $values = @{ ThrottleLimit = 4; 'Windows11Upgrade.SetupSourcePath' = ''; DryRun = $true }
    $selection = New-GuiSelectionFile -Catalog $catalog -Campaign Windows11Upgrade -Values $values
    $captured = Read-SemtJsonFile -Path $selection
    Test-Step 'GUI selection is stored per Campaign under Runs\<Campaign>\_GuiSelections' { $selection -like (Join-Path $packageRoot 'Runs\Windows11Upgrade\_GuiSelections\*.json') -and $captured.Campaign -eq 'Windows11Upgrade' }
    Test-Step 'GUI selection omits empty text options so Wave/global values still apply' { [int]$captured.Options.ThrottleLimit -eq 4 -and -not $captured.Options.PSObject.Properties['Windows11Upgrade.SetupSourcePath'] }
    Test-Step 'engine accepts the GUI selection only for its own Campaign' {
        $ok = $false
        try { [void](Read-SemtGuiSelection -Path $selection -Campaign HybridJoin) } catch { $ok = $true }
        $layer = Read-SemtGuiSelection -Path $selection -Campaign Windows11Upgrade
        $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign Windows11Upgrade -WaveLayer @{ ThrottleLimit = 2 } -GuiLayer $layer
        $ok -and $resolved['ThrottleLimit'].Value -eq 4 -and $resolved['ThrottleLimit'].Source -eq 'Gui'
    }
    #endregion

    #region Windows PowerShell 5.1 STA
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $validate = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $guiPath -ValidateOnly 2>&1 | ForEach-Object { [string]$_ })
    Test-Step 'GUI -ValidateOnly loads under Windows PowerShell 5.1 STA for both Campaigns' { $LASTEXITCODE -eq 0 -and @($validate | Where-Object { $_ -match '^ValidateOnly: GUI 1\.0\.0 loaded; campaign=(HybridJoin|Windows11Upgrade); options=\d+' }).Count -eq 2 }
    Test-Step 'GUI -ValidateOnly writes no GUI state file' { @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter 'GuiState.*' -File).Count -eq 0 }
    $demoScript = Join-Path $tempRoot 'demo-check.ps1'
    @'
param($Gui)
$ErrorActionPreference = 'Stop'
$state = . $Gui -Demo -ValidateOnly
'VIEWS=' + ($state.ViewNames -join ',')
'OPTIONS1=' + $state.Data.Editors.Count
$state.Controls.CampaignCombo.SelectedIndex = 1
'CAMPAIGN=' + $state.Data.Campaign + ';OPTIONS2=' + $state.Data.Editors.Count
foreach ($name in 'LaunchWaveButton','LaunchAllButton','LaunchSingleButton','CreateWaveButton','AutomaticPreviewButton','AutomaticCreateButton','StopAllButton','ForceStopSelectedRunButton','OpenRunReportButton','ResetOptionsButton','RefreshButton') {
    $state.Controls[$name].RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
}
'FAILED=' + @($state.Activity | Where-Object { $_ -match 'Demo action failed' }).Count
'OUTPUT=' + $state.LastOutput
'NOLAUNCH=' + (@($state.Activity | Where-Object { $_ -match 'No launcher, PsExec, worker' }).Count -gt 0)
$state.Window.Close()
'@ | Set-Content -LiteralPath $demoScript -Encoding ASCII
    $demo = @(& $powershell -NoProfile -STA -ExecutionPolicy Bypass -File $demoScript -Gui $guiPath 2>&1 | ForEach-Object { [string]$_ })
    $demoText = $demo -join "`n"
    Test-Step 'demo loads every view' { $demoText -match 'VIEWS=Waves,Single PC,New Wave,Automatic Wave,Runs,Options' }
    Test-Step 'demo shows the real catalog options for both Campaigns' { $demoText -match 'OPTIONS1=\d{2}' -and $demoText -match 'CAMPAIGN=Windows11Upgrade;OPTIONS2=\d{2}' }
    Test-Step 'every demo action completes without error' { $demoText -match 'FAILED=0' }
    Test-Step 'demo report is generated in the demo temp folder only' { $demoText -match 'OUTPUT=.*SmartM365-Demo.*Run-Report-DEMO\.html' }
    Test-Step 'demo states that nothing was launched' { $demoText -match 'NOLAUNCH=True' }
    Test-Step 'demo wrote no Wave, Run or configuration into the package' { @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Waves') -Directory -Recurse | Where-Object { $_.Parent.Name -in @('HybridJoin','Windows11Upgrade') }).Count -eq 0 -and @(Get-ChildItem -LiteralPath (Join-Path $packageRoot 'Config') -Filter '*.local.json' -File).Count -eq 0 }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = $previousOverride
    $env:SEMT_GUI_UPDATE_CHECK = $previousUpdate
    try { if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force } } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLvpE9/u+VE3yR
# oHEjl1Gp7GAzbimhew+35UYIhWIt6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIALRI+uOLOtuPx6otMWkzpjDr32gj2fvCSC8rRov1MKAMA0GCSqG
# SIb3DQEBAQUABIIBgHzCaKAGnOuiU8AVWSmgWVKaN2ipg84Afk5R0ILEgcyQIHRE
# 4jKXXCW1qI5DrHdnMP9jXtSPySZOk6gWCTH50AKGk2nyrXph/YWwfQoQ5KI0Qrfk
# Devoj90pTKoRGE2NSNJUelzl741ZsGLJpLU+olrVb/sBj3lmRYdll2y/s8PROqWG
# YVDPCrrCWH/jA08hJ3lhv4EblIvOrXjeoubu8mw9e6PoVcUalIyaYjF2nRjIW88O
# /E6NSIYOifiiNFXE2xkC9tPL/IftShe8QfJhnI5dgU8gpZhYRR6YVEU4b61ttUYy
# O8YAI0RMpl8Ft2ZKpcLZ9EnHEck9znwsNIVGScX5SRpYLuPelDKM4kv5UBXuDRlQ
# qU1JIARx+9/YJhsKLgd9mQVcEyZCkaF7NbocsLIbDksQV0JmAY8O0AVd1xtteOG6
# qoUgfaXJX8G2A4X+a/btfj9VSbA8p+x89HLoXstPY1m9KEeLk8V67IOwqBhIzpYL
# dweyiXFFOZQYb/kosaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NThaMC8GCSqGSIb3DQEJBDEiBCDp5j+KxtD550xJv484XMd4dco5/3b89o5C96uF
# U2ueHzANBgkqhkiG9w0BAQEFAASCAgB6z171sBZfCu51nNb2Pg01rSntRwpJ7vzp
# 20vEb3gBz78eJXWxLU8Z18dnPslN0ISYxY0gHry5VlvMk48B7qQS1YoiKzI5tu15
# HQ2wvwm4RZIbl/mA4fA1CNJ9jD5fyjPblBRYuXPeBhYn04tN52YojKHC1bbDhBZe
# rVOWFrUrojOL4Eb5FBKV6Zbimcfyx8qiHwoSoIqAOSmWE2WbORAu81PSveI0kWDQ
# AqTW2gQyov6Lgxg9UI6MxC9CgA7e1XK3pIA+Wg07/A7wkZxos/kEDyLvuXBX4QQU
# 8SpnPleiv3KutLO1TbXzvt4VulvLSGOCspBchxOf+VpZSyifXYwu++LO++PU+3Li
# naMtiZ8+Dw7mUGGDpjCr5oDwcd/QkDzlooCeZnqry3WQ3gPxH8t5k2tGLiF7feJM
# ytWxMIwBMSuUMT3AslfNKFmoUoD2+o1Bxl2wboqjpm9oxGQE/Xf6D7U15W0BcH6o
# Wrkkyz5hcjsb4BB+GtD6YNM1aBXRUJbI0sNjoPgFpQA8ek0Ifofgat9XKJfBsECu
# uvpH0XQ3niZvQTovCXblCI1Sa8pWIItAPaP+74JZxFswdko1dxtnwrOQaYyoHAgh
# ykxfdZ9j51ZvO5tvkNenjhuBo9n1aG2SEAVCuRawQgYQJTYuc4j+dp4BkImZdvvD
# 4k9WDHg8nQ==
# SIG # End signature block
