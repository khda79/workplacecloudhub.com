<#
.SYNOPSIS
Validates scoped endpoint run-guard retries and generated GUI CMD launchers.

.VERSION
1.0.0
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ([string]$Actual -ne [string]$Expected) {
        throw ("ASSERT FAILED: {0}. Expected={1}; Actual={2}" -f $Message,$Expected,$Actual)
    }
}

function Import-ScriptFunction {
    param([string]$Path, [string]$Name)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw ($errors | ForEach-Object Message | Out-String) }
    $functionAst = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true) | Select-Object -First 1
    if (-not $functionAst) { throw "Function not found: $Name" }
    Set-Item -Path ("Function:\global:{0}" -f $Name) -Value $functionAst.Body.GetScriptBlock()
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$orchestrator = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
$gui = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SmartM365-W11UT-RunGuardCmd-{0}' -f [guid]::NewGuid().ToString('N'))

try {
    Import-ScriptFunction -Path $orchestrator -Name 'ConvertTo-TechnicianRunGuardUtcDateTime'
    Import-ScriptFunction -Path $orchestrator -Name 'Test-TechnicianRunGuardEndpointBypassEligible'

    $dueUtc = (Get-Date).ToUniversalTime().AddMinutes(-1).ToString('o')
    $futureUtc = (Get-Date).ToUniversalTime().AddMinutes(10).ToString('o')
    $dueSetup = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'SetupFailure'; RemoteStatus = 'DIRECT_SETUP_UPGRADE_FAILED' }
    Assert-True -Condition (Test-TechnicianRunGuardEndpointBypassEligible -Entry $dueSetup) -Message 'due setup failure permits a one-attempt endpoint bypass'

    $futureSetup = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $futureUtc; FailureCategory = 'SetupFailure'; RemoteStatus = 'DIRECT_SETUP_UPGRADE_FAILED' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $futureSetup)) -Message 'future retry stays blocked'

    $networkFailure = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'NetworkTransient'; RemoteStatus = 'ADMIN_SHARE_UNREACHABLE' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $networkFailure)) -Message 'network failure never bypasses endpoint guard'

    $pendingReboot = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'OperatorAction'; RemoteStatus = 'PENDING_REBOOT_USER_CONNECTED' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $pendingReboot)) -Message 'pending reboot never bypasses endpoint guard'

    $ambiguousExecution = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'ExecutionTransient'; RemoteStatus = 'REMOTE_RESULT_STALE' }
    Assert-True -Condition (-not (Test-TechnicianRunGuardEndpointBypassEligible -Entry $ambiguousExecution)) -Message 'ambiguous execution state never bypasses endpoint guard'

    $endpointTransient = [pscustomobject]@{ State = 'Result'; RetryAfterUtc = $dueUtc; FailureCategory = 'ExecutionTransient'; RemoteStatus = 'SETUP_CACHE_LOCKED' }
    Assert-True -Condition (Test-TechnicianRunGuardEndpointBypassEligible -Entry $endpointTransient) -Message 'recognized endpoint transient can retry after backoff'

    . $gui -ValidateOnly | Out-Null
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticInventoryFileInfo'
    Import-ScriptFunction -Path $gui -Name 'Get-AutomaticInventorySnapshot'
    function global:Get-ConfiguredValue { param([string]$Name); return '' }
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    function global:Show-GuiWarningYesNo { param([string]$Title, [string]$Message); return [bool]$global:SyntheticRootFallbackAccepted }
    $automaticToolkitRoot = Join-Path $testRoot 'AutomaticToolkit'
    $automaticScriptsRoot = Join-Path $automaticToolkitRoot 'Scripts'
    New-Item -ItemType Directory -Path $automaticScriptsRoot -Force | Out-Null
    [pscustomobject]@{ DeviceName = 'ROOT-AD'; Marker = 'ROOT' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesAD.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; InventoryTenantProfile = 'prod'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    $toolkitRoot = $automaticToolkitRoot

    function global:Invoke-GuiPowerShellProcess {
        param([string]$ScriptPath, [string[]]$Arguments, [string]$LogPath, [string]$Activity)
        if ($global:SyntheticInventoryRefreshFails) { throw 'Synthetic inventory refresh failure.' }
        $outputIndex = [array]::IndexOf($Arguments, '-OutputPath')
        if ($outputIndex -lt 0) { throw 'Synthetic exporter did not receive -OutputPath.' }
        $outputPath = $Arguments[$outputIndex + 1]
        if ($ScriptPath -like '*Export-ADDevicesCsv.ps1') {
            [pscustomobject]@{ DeviceName = 'REFRESHED-AD'; Marker = 'SNAPSHOT' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        else {
            [pscustomobject]@{ DeviceName = 'REFRESHED-INTUNE'; Marker = 'SNAPSHOT'; InventoryTenantProfile = 'prod'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        Set-Content -LiteralPath $LogPath -Value $Activity -Encoding UTF8
    }

    $automaticSnapshot = Get-AutomaticInventorySnapshot -Source Both -TenantProfile 'prod'
    Assert-True -Condition ($automaticSnapshot.AdInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesAD.csv')) -Message 'automatic AD selection does not silently reuse root CSV'
    Assert-True -Condition ($automaticSnapshot.IntuneInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesIntune.csv')) -Message 'automatic Intune selection does not silently reuse root CSV'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-AD' -Message 'automatic AD snapshot was freshly generated'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.IntuneInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-INTUNE' -Message 'automatic Intune snapshot was freshly generated'

    $automaticRunPath = Join-Path $testRoot 'AutomaticRun'
    New-Item -ItemType Directory -Path $automaticRunPath -Force | Out-Null
    $copiedSnapshots = @(Copy-AutomaticInventorySnapshotToRun -InventoryContext $automaticSnapshot -RunPath $automaticRunPath)
    Assert-Equal -Actual $copiedSnapshots.Count -Expected 2 -Message 'both automatic snapshots were copied into the LOT run'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesAD.csv') -PathType Leaf) -Message 'run-local AD snapshot exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesIntune.csv') -PathType Leaf) -Message 'run-local Intune snapshot exists'

    $global:SyntheticInventoryRefreshFails = $true
    $fallbackRejected = $false
    try { Get-AutomaticInventorySnapshot -Source AD -TenantProfile 'prod' | Out-Null } catch { $fallbackRejected = $true }
    Assert-True -Condition $fallbackRejected -Message 'fresh root cache is not reused when explicit fallback is rejected'

    $global:SyntheticRootFallbackAccepted = $true
    $acceptedFallback = Get-AutomaticInventorySnapshot -Source AD -TenantProfile 'prod'
    Assert-Equal -Actual $acceptedFallback.AdInventoryCsv -Expected (Join-Path $automaticToolkitRoot 'DevicesAD.csv') -Message 'fresh root cache is used only after explicit fallback acceptance'
    Assert-True -Condition (($acceptedFallback.SourceDetails -join ' ') -match 'explicitly accepted') -Message 'accepted root fallback is marked in source details'
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    Assert-Equal -Actual (ConvertTo-CmdArgument -Value 'C:\Program Files\SmartM365') -Expected '"C:\Program Files\SmartM365"' -Message 'CMD path argument quoting'
    Assert-Equal -Actual (ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value 'value&safe') -Expected 'set "W11UT_TEST=value&safe"' -Message 'quoted SET command keeps ampersand literal'

    $unsafeRejected = $false
    try { ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value '%TEMP%' | Out-Null } catch { $unsafeRejected = $true }
    Assert-True -Condition $unsafeRejected -Message 'percent expansion is rejected in generated CMD values'

    $workingDirectory = Join-Path $testRoot 'Run With Spaces'
    $evidenceDirectory = Join-Path $workingDirectory 'Logs'
    New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
    $launchPath = New-GuiLaunchCommandFile -WorkingDirectory $workingDirectory -Commands @('echo CMD_TEST_OK') -NamePrefix 'Synthetic' -EvidenceDirectory $evidenceDirectory
    Assert-True -Condition (Test-Path -LiteralPath $launchPath -PathType Leaf) -Message 'temporary GUI launcher exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $evidenceDirectory 'GuiLaunchCommand.cmd') -PathType Leaf) -Message 'run-local GUI launcher evidence exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $evidenceDirectory 'GuiLaunchCommandEvidence.txt') -PathType Leaf) -Message 'GUI launcher evidence metadata exists'

    $cmdArguments = '/d /s /c ""{0}""' -f $launchPath
    $cmdProcess = Start-Process -FilePath $env:ComSpec -ArgumentList $cmdArguments -WorkingDirectory $workingDirectory -Wait -PassThru
    Assert-Equal -Actual $cmdProcess.ExitCode -Expected 0 -Message 'generated CMD launcher execution through Start-Process quoting'

    Write-Output 'SmartM365 Windows 11 run guard and CMD synthetic tests passed.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDGPWmFqmDfZ9My
# albXLR8ighR2oMEmIpgzjPecv1zW5qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIO7QN6OV3WCcG8yGwyFKJ9B9ylSiFfJNRSpMg+GkAePlMA0GCSqG
# SIb3DQEBAQUABIIBgHDVpoArujjXjgD3Dt1Hd5d0ciK+0rbIRSbqgZTvfpkXqgSk
# QKYlyeLtq8OY7g/Cj8rq5jD5Hdv3WgxVMzvov+Pt+TJLYdc2oepV1UNVoq8o3WgX
# 5bEg0vHWYeLE4KPWoWDIdqXH8HHvvYcCJgYBUW5EMSUa3FUDnNkgjPixR/yAiGWo
# hTpKNSB0MryR39vpn2/nXVd5zNemQlHLInQ4uwGDhLO+5M8IsoeMlunDgGSka2q9
# kJip0GXyszMUxA/JB8DQgxwyTfB2g3mhu4yg6GoejIo8TIwY+Y5++9qNwoYxJVdC
# NLNXR2FStSM+krWRF3QmT5hG3K60tnsYTXfAky0uHtAJxEx+SLs6I1FAcXQRwRMr
# ci0QotAGKIs61x4FDIHOn2PsIxhf3aGdRVZ68mbtZNJFXyaG318nyUel+uYA1Uli
# vKrx+Aw2XrepUpymqeWd+ifPnBC4sgWIILvzPU/wa8XAvtYcVG4yyTek7Yt/XJMn
# ncQ2KBMWkcOO64HmlqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjQxNDQz
# MThaMC8GCSqGSIb3DQEJBDEiBCAN79r55eN0PmeM86dHh/RyNVGmAiTzwemMUjJa
# 7XEMETANBgkqhkiG9w0BAQEFAASCAgArAdgw0XqDpvmAGlyJEv5Wop6yC9Xr/4Qj
# 1wujICM4VLxy2PuOK5/sFwW528N7hDKyZB2Pn2gI/inyLZzwzKIg7Sa1CniVsItQ
# YtptQ3bLwiv8RUfmb6YFPOXx+DxMr9MYT3/wqv50fj6D8ThPyRezzVwF9HIg8Yy9
# SIBxeLCR1mVU5hyzxLF1dQCnNPmdt+GqQClIl95VYI3my6PkuEarO5A+jk3MjvS+
# Zax5NFcx2258Y/2cq/NXn0hAdM/Nhq/wir0vh1gcAExwHfBc8CD8lCsaudW20qMH
# WXDoZleODEQO3m2yohyT+zNH4A+/9n0KYzXatBEVupJtTD6yGMdNbDr4sVH49eR3
# BH4xwX2uJgG2G7lXa1K1fgn/m/UpM/LQUh8w2DTY7T4m0vSe2ynZfoUJHJzn9x42
# 6UCZi7dtol6fTECvWVJnZ6LpcxydoyXoX9FSe4Ln48ZHFI7vn5MF4owpSaLaBLkq
# BrtyKRc6w0361AjubM6Cfnz3qeOyyZFthQPg3/gbppBc5U2W7GNBoGWABN4VHUCd
# KPv9T1mWirHnHauAx9AR1Dm8UyMXQMF2YzrT+gRpY+i5XNlpk2/PsmrahB/tehSS
# pDcvxpr9hM4R1BVPWx0GzBSDQyxm9z7F6GmS72v4wDSsiT145DbM/rGWcYwNCZbp
# YG2ZnkbnQQ==
# SIG # End signature block
