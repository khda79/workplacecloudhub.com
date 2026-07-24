<#
.SYNOPSIS
Validates scoped endpoint run-guard retries and generated GUI CMD launchers.

.VERSION
1.1.0
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
$sourceToolkitRoot = $toolkitRoot
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
    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $automaticScriptsRoot 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Value '# synthetic exporter' -Encoding UTF8
    $toolkitRoot = $automaticToolkitRoot

    function global:Invoke-GuiPowerShellProcess {
        param([string]$ScriptPath, [string[]]$Arguments, [string]$LogPath, [string]$Activity, [switch]$Interactive)
        if ($global:SyntheticInventoryRefreshFails) { throw 'Synthetic inventory refresh failure.' }
        $outputIndex = [array]::IndexOf($Arguments, '-OutputPath')
        if ($outputIndex -lt 0) { throw 'Synthetic exporter did not receive -OutputPath.' }
        $outputPath = $Arguments[$outputIndex + 1]
        if ($ScriptPath -like '*Export-ADDevicesCsv.ps1') {
            [pscustomobject]@{ DeviceName = 'REFRESHED-AD'; Marker = 'SNAPSHOT' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        else {
            if (-not $Interactive) { throw 'Synthetic Intune export was not launched interactively.' }
            [pscustomobject]@{ DeviceName = 'REFRESHED-INTUNE'; Marker = 'SNAPSHOT'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        }
        Set-Content -LiteralPath $LogPath -Value $Activity -Encoding UTF8
    }

    $automaticSnapshot = Get-AutomaticInventorySnapshot -Source Both
    Assert-True -Condition ($automaticSnapshot.AdInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesAD.csv')) -Message 'automatic AD selection does not silently reuse root CSV'
    Assert-True -Condition ($automaticSnapshot.IntuneInventoryCsv -ne (Join-Path $automaticToolkitRoot 'DevicesIntune.csv')) -Message 'automatic Intune selection does not silently reuse root CSV'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.AdInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-AD' -Message 'automatic AD snapshot was freshly generated'
    Assert-Equal -Actual (Import-Csv -LiteralPath $automaticSnapshot.IntuneInventoryCsv | Select-Object -First 1 -ExpandProperty DeviceName) -Expected 'REFRESHED-INTUNE' -Message 'automatic Intune snapshot was freshly generated'
    Assert-Equal -Actual $automaticSnapshot.TenantId -Expected 'tenant-001' -Message 'automatic snapshot records the connected tenant'
    Assert-Equal -Actual $automaticSnapshot.AuthenticationMode -Expected 'DelegatedInteractive' -Message 'automatic snapshot records delegated interactive authentication'

    $automaticRunPath = Join-Path $testRoot 'AutomaticRun'
    New-Item -ItemType Directory -Path $automaticRunPath -Force | Out-Null
    $copiedSnapshots = @(Copy-AutomaticInventorySnapshotToRun -InventoryContext $automaticSnapshot -RunPath $automaticRunPath)
    Assert-Equal -Actual $copiedSnapshots.Count -Expected 2 -Message 'both automatic snapshots were copied into the LOT run'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesAD.csv') -PathType Leaf) -Message 'run-local AD snapshot exists'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $automaticRunPath 'DevicesIntune.csv') -PathType Leaf) -Message 'run-local Intune snapshot exists'

    $global:SyntheticInventoryRefreshFails = $true
    $fallbackRejected = $false
    try { Get-AutomaticInventorySnapshot -Source AD | Out-Null } catch { $fallbackRejected = $true }
    Assert-True -Condition $fallbackRejected -Message 'fresh root cache is not reused when explicit fallback is rejected'

    $global:SyntheticRootFallbackAccepted = $true
    $acceptedFallback = Get-AutomaticInventorySnapshot -Source AD
    Assert-Equal -Actual $acceptedFallback.AdInventoryCsv -Expected (Join-Path $automaticToolkitRoot 'DevicesAD.csv') -Message 'fresh root cache is used only after explicit fallback acceptance'
    Assert-True -Condition (($acceptedFallback.SourceDetails -join ' ') -match 'explicitly accepted') -Message 'accepted root fallback is marked in source details'

    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'AppOnly'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    $appOnlyFallbackRejected = $false
    try { Get-AutomaticInventorySnapshot -Source Intune | Out-Null } catch { $appOnlyFallbackRejected = $true }
    Assert-True -Condition $appOnlyFallbackRejected -Message 'app-only Intune root cache is rejected even when fallback confirmation is enabled'

    [pscustomobject]@{ DeviceName = 'ROOT-INTUNE'; Marker = 'ROOT'; InventoryTenantId = 'tenant-001'; InventoryAuthenticationMode = 'DelegatedInteractive'; InventoryScope = 'AllManagedDevices' } | Export-Csv -LiteralPath (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -NoTypeInformation -Encoding UTF8
    $delegatedFallback = Get-AutomaticInventorySnapshot -Source Intune
    Assert-Equal -Actual $delegatedFallback.IntuneInventoryCsv -Expected (Join-Path $automaticToolkitRoot 'DevicesIntune.csv') -Message 'delegated Intune root cache can be explicitly accepted'
    Assert-Equal -Actual $delegatedFallback.AuthenticationMode -Expected 'DelegatedInteractive' -Message 'delegated fallback authentication provenance is preserved'
    $global:SyntheticInventoryRefreshFails = $false
    $global:SyntheticRootFallbackAccepted = $false
    Assert-Equal -Actual (ConvertTo-CmdArgument -Value 'C:\Program Files\SmartM365') -Expected '"C:\Program Files\SmartM365"' -Message 'CMD path argument quoting'
    Assert-Equal -Actual (ConvertTo-CmdSetCommand -Name 'W11UT_TEST' -Value 'value&safe') -Expected 'set "W11UT_TEST=value&safe"' -Message 'quoted SET command keeps ampersand literal'

    $exporterText = Get-Content -LiteralPath (Join-Path $sourceToolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1') -Raw
    Assert-True -Condition ($exporterText -match 'DeviceManagementManagedDevices\.Read\.All') -Message 'Intune exporter requests the delegated read scope'
    Assert-True -Condition ($exporterText -notmatch 'Initialize-SmartM365TenantContext|CertificateThumbprint|\[string\]\$Tenant\s*=') -Message 'Intune exporter has no app-only tenant profile dependency'
    Assert-True -Condition ($exporterText -match 'InventoryAuthenticationMode') -Message 'Intune exporter records authentication provenance'
    $rootCmdText = Get-Content -LiteralPath (Join-Path $sourceToolkitRoot 'Export-IntuneDevicesCsv.cmd') -Raw
    Assert-True -Condition ($rootCmdText -notmatch 'W11UT_INTUNE_TENANT_PROFILE|Profile\s*:\s*test') -Message 'root Intune CMD does not force the test profile'
    $orchestratorText = Get-Content -LiteralPath $orchestrator -Raw
    Assert-True -Condition ($orchestratorText -notmatch 'IntuneTenantProfile|InventoryTenantProfile') -Message 'LOT orchestrator has no app-only profile dependency'
    Assert-True -Condition ((Get-Content -LiteralPath $gui -Raw) -match 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH.{0,200}=\s*''1''') -Message 'automatic LOT launch freezes the interactive Intune snapshot for the run'

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCW5z9UnfarMLeB
# uXT1KufO/wUuyi9CEz4ye1bau6QRcaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINrt/hrBBzCpm6vXJgSIdYhFajC4xXxdSsBFxEP9wUpZMA0GCSqG
# SIb3DQEBAQUABIIBgDMYmklAaVoX1dwsnpve7og3Og33Y6u9GXbFoH1Jmnl8sbyt
# s35RVPiY3T9EXcpEKSjapPNr2CBA+YbSgU5Nm1+pZvZq39VHgKEwXJfvLzxb+m04
# 4U/xQnl/XduVNywompFirOWaUK2xNU+cxal6qIgwGEgeTMD5rrx4bMu9wBRMSH/X
# y4wiXwvLrObBq3cC9TG+0yaO7KxWs0Py5jEjnF8EtFl00fgKLf0jzr/9FzDtpZGY
# FSjjjkji+4iF42VV7br+6BV40LT61n210ylB/Cu2GsLKj3UHFtuXeit137FgmFy3
# ZjUdAsalXxpRC9uazGNDRyMLTmUTJnc4KCA/8IJasz8Jezo+zFbh9T+LiwSVCgY2
# aOy0FSwqpjfPtsHorTgWFs3LWWeOIFzq843lPff3R8+NZhEpkKuL+9ahTQUUExDI
# U5/TKfjtNL0kuE3vkSOl7wso9iU1RaHd+0c+DgbVgtW79fAtXmo/lyBgSPq0hse1
# /qy+ZzBv9QsQXcWezqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjQxODMx
# MDlaMC8GCSqGSIb3DQEJBDEiBCDE4GogV9NkbXYJkZvmjPgGSUktx1UunID65Dwg
# Z7tiFzANBgkqhkiG9w0BAQEFAASCAgBDFtyYGZEZ6GkCL7TE7KNhjabTHKxu8mOf
# OCawdc83w2nwrr9T5rruQnYBMxgl2AlcU9K/qgGWxRWApcFPiTJ7ppvhGDH92D+/
# IP69UEgR4e+zPitqJM+GsKOJ8sWFv/LFwqa9fHScoPN4it0lduTJmWUgbpdELAli
# y4vzaQMAH8IBOPOJxf9YV70P9/LlJbzHumrPbZ1/PGwAuHCNqHQ6axH2NSPrJggW
# DGqo0wZeKw0nIeZZRxsyc4qJXFD5CkiTABiNW6nrpL22TIz0ZPPR9WPKYfmUWNXH
# liQ2GEUD7ms4MnFstSLmvFln5z+LDTLgjOC0LgvwNeqljN2AD30tPI2+FYkw5XqN
# tRJKU/2/ytlRJptBsa08wlaWzG8j8gbk7EgUxu8VhaOkBnGjTlXrqwCsNpZHUtlC
# VKIapKGe3w4YHYaNFwQztjn4PZGxnKACULBGYz4AStf0d4Geomvbi9sJtPFH3eh4
# 5hqrP/vAul5hAoX9HtckmGPGb2W0ikyb6EKCyw0j44umKpUdet6E7G0L9rAuo43r
# +CcJfRyqoZ3EZ4BHi2WFP+UH9faNMNDhdBfM9RVcgRvzMWbiiwL+NvaieQGvd8Fx
# Fui+g7cKJe2eFsBFtAjlYhnrIIuNknnPkSHkz+zk4WRpitNfRRUHMInpiLbVBL7P
# Nqu2cAyRQg==
# SIG # End signature block
