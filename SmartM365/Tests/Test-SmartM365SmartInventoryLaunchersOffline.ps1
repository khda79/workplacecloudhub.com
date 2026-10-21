#Requires -Version 7.0
<#
.SYNOPSIS
Validate offline that the hand-written SmartInventory launchers stay aligned with the jobs template.
.VERSION
1.0.0
.NOTES
Every template job is either started through its LauncherPath, mapped to the manual launcher that
reproduces it, or explicitly covered only by the generated Launchers\Jobs launcher.
A LauncherPath launcher must forward the orchestrator Arguments (%*), follow the orchestrator
tenant and use the engine of PowerShellEdition. A mapped manual launcher must run the job script
with that engine and pass every template Argument. Reads files only: nothing is executed.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$inventoryRoot = Join-Path $smartRoot 'SmartInventory'
$launchersRoot = Join-Path $inventoryRoot 'Launchers'
$checks = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

# Manual launcher (relative to Launchers) per job without LauncherPath. $null = only the generated
# Launchers\Jobs\Start-<JobName>.cmd reproduces the job.
$manualLaunchers = [ordered]@{
    'M365-ActiveUsers-Inventory' = 'Cloud\Start-SmartM365-ActiveUsers-Inventory.cmd'
    'M365-EntraDevices-Inventory' = 'Cloud\Start-SmartM365-EntraDevices-Inventory.cmd'
    'Intune-Devices-Inventory' = 'Cloud\Start-SmartM365-Devices-Inventory.cmd'
    'M365-VerifiedDomains-Inventory' = 'Cloud\Start-SmartM365-VerifiedDomains-Inventory.cmd'
    'EXO-AcceptedDomains-Inventory' = 'Cloud\Start-SmartM365-EXO-AcceptedDomains-Inventory.cmd'
    'M365-Licences-Inventory' = 'Cloud\Start-SmartM365-Licences-Inventory.cmd'
    'Intune-Devices-BIOS-Inventory' = 'Cloud\Start-SmartM365-Devices-BIOS-Inventory.cmd'
    'Intune-DeviceSystem-Inventory' = 'Cloud\Start-SmartM365-Device-System-Inventory.cmd'
    'Intune-Devices-Compliance-Inventory' = 'Cloud\Start-SmartM365-Devices-Compliance-Inventory.cmd'
    'Intune-Devices-UpgradeEligibility' = 'Cloud\Start-SmartM365-Devices-UpgradeEligibility.cmd'
    'EXO-Mailboxes-Inventory-Fast' = 'Cloud\Start-SmartM365-EXO-Mailboxes-Inventory.cmd'
    'EXO-Mailboxes-Inventory' = 'Cloud\Start-SmartM365-EXO-Mailboxes-Inventory-IncludeStats.cmd'
    'EXO-Mailboxes-Permissions' = $null
    'Intune-WinUpdate-Status' = 'Cloud\Start-SmartM365-WinUpdate_Status_From_Intune.cmd'
    'Intune-Windows11-Readiness-Issues' = 'Cloud\Start-SmartM365-Intune-Windows11-Readiness-Issues-Inventory.cmd'
    'M365-SecureScore-Inventory' = 'Cloud\Start-SmartM365-SecureScore-Inventory.cmd'
    'M365-AuthenticationMethodsRegistration-Inventory' = 'Cloud\Start-SmartM365-AuthenticationMethodsRegistration-Inventory.cmd'
    'M365-ConditionalAccess-Inventory' = 'Cloud\Start-SmartM365-ConditionalAccess-Inventory.cmd'
    'Intune-EndpointSecurityHealth-Inventory' = 'Cloud\Start-SmartM365-Intune-EndpointSecurityHealth-Inventory.cmd'
    'Intune-EndpointAnalytics-Inventory' = 'Cloud\Start-SmartM365-EndpointAnalytics-Inventory-IncludeStartupProcesses.cmd'
    'Intune-AutopatchAlerts-Inventory' = 'Cloud\Start-SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.cmd'
    'Intune-RBAC-GroupMembers' = 'Cloud\Start-SmartM365-Intune-RBAC-GroupMembers.cmd'
    'Intune-Remediations-Export' = 'Cloud\Start-SmartM365-Intune-ExportRemediationScripts.cmd'
    'M365-BackupProtectedMailboxes-Inventory' = 'Cloud\Start-SmartM365-M365-BackupProtectedMailboxes-Inventory-Interactive.cmd'
    'M365-BackupPolicyScope-Inventory' = 'Cloud\Start-SmartM365-M365-BackupPolicyScope-Inventory.cmd'
    'M365-BackupProtectedSitesAndDrives-Inventory' = 'Cloud\Start-SmartM365-M365-BackupProtectedSitesAndDrives-Inventory-Interactive.cmd'
    'M365-PowerBIFabricActivity-Inventory' = 'Cloud\Start-SmartM365-PowerBIFabricActivity-Inventory.cmd'
    'M365-Usage-Inventory' = 'Cloud\Start-SmartM365-M365UsageReports-Inventory.cmd'
    'M365-CopilotUsage-Inventory' = 'Cloud\Start-SmartM365-CopilotUsage-Inventory.cmd'
    'M365-TeamsPhonePstnUsage-Inventory' = 'Cloud\Start-SmartM365-TeamsPhonePstnUsage-Inventory-IncludePhoneAssignments.cmd'
    'M365-LicensePricing-Inventory' = 'Cloud\Start-SmartM365-LicensePricing-Inventory.cmd'
    'SmartFinOps-Workplace-Valuation' = $null
    'SmartFinOps-Workplace-Report' = $null
    'Exchange-HybridIdentity-Issues' = 'Cloud\Start-SmartM365-Exchange-HybridIdentity-Issues-Inventory.cmd'
    'Mailboxes-PermissionsByUser-Report' = 'Cloud\Start-SmartM365-Mailboxes-AllSources-PermissionsByUserReport.cmd'
    'WorkplaceEvidence-Prepare' = 'Cloud\Start-SmartM365-WorkplaceEvidence-Prepare.cmd'
    'M365-SyncHealth-Inventory' = 'Cloud\Start-SmartM365-AzureADConnect-SyncHealth-Inventory.cmd'
    'EXO-MigrationJobs-Inventory' = 'Cloud\Start-SmartM365-EXO-MigJob-Inventory.cmd'
    'EXO-Mailboxes-CalendarPermissions' = 'Cloud\Start-SmartM365-EXO-Mailboxes-CalPerm_Inventory.cmd'
    'Intune-DiscoveredApps-Inventory' = 'Cloud\Start-SmartM365-Intune-DiscoveredApps-Inventory.cmd'
    'M365-Teams-Inventory' = 'Cloud\Start-SmartM365-Teams-Inventory.cmd'
    'M365-SPO-Inventory' = 'Cloud\Start-SmartM365-SPO-Inventory.cmd'
    'Intune-WindowsAutopilot-Inventory' = 'Cloud\Start-SmartM365-WindowsAutopilot-Inventory.cmd'
    'M365-WorkplaceScope-Inventory' = $null
    'Intune-Configurations-Inventory' = 'Cloud\Start-SmartM365-IntuneConfigurations-Inventory.cmd'
    'Intune-ManagedApps-Inventory' = 'Cloud\Start-SmartM365-IntuneManagedApps-Inventory.cmd'
    'Intune-EnrollmentFailures-Inventory' = 'Cloud\Start-SmartM365-IntuneEnrollmentFailures-Inventory.cmd'
}

function Get-LauncherInvocation {
    # Returns the engine, script leaf, fixed arguments and forwarding of the launcher's final
    # PowerShell call, with the launcher's own set "NAME=value" variables expanded.
    param([string]$Path)
    $lines = @(Get-Content -LiteralPath $Path)
    $variables = @{}
    foreach ($line in $lines) {
        if ($line -match '^\s*set "([A-Za-z0-9_]+)=(.*)"\s*$') { $variables[$Matches[1]] = $Matches[2] }
    }
    $call = @($lines | Where-Object { $_ -match '-File\s+"' -and $_ -notmatch 'Unblock-File' } | Select-Object -Last 1)
    if ($call.Count -eq 0 -or $call[0] -notmatch '^\s*"(%[A-Z0-9_]+%)"\s.*-File\s+"([^"]+)"(.*)$') { throw "No PowerShell call found in launcher: $Path" }
    $engineVariable = $Matches[1]
    $scriptPath = $Matches[2]
    $rest = $Matches[3]
    for ($i = 0; $i -lt 5; $i++) { foreach ($name in $variables.Keys) { $scriptPath = $scriptPath.Replace("%$name%", $variables[$name]) } }
    $engine = switch ($engineVariable) { '%PWSH%' { 'PowerShell7' } '%POWERSHELL5%' { 'WindowsPowerShell' } default { 'Unknown' } }
    $tenant = if ($rest -match '-Tenant\s+"?([^\s"]+)"?') { $Matches[1] } else { '' }
    $fixed = (($rest -replace '%\*', ' ' -replace '-Tenant\s+"?[^\s"]+"?', ' ') -replace '\s+', ' ').Trim()
    [pscustomobject]@{
        Engine = $engine
        ScriptLeaf = ($scriptPath -split '[\\/]')[-1]
        Tenant = $tenant
        FixedArguments = $fixed
        ForwardsArguments = $rest -match '%\*'
    }
}

function Get-ArgumentGroup {
    # Splits "-Name value -Switch" into "-Name value", "-Switch".
    param([string]$Arguments)
    $groups = [System.Collections.Generic.List[string]]::new()
    foreach ($token in @(([string]$Arguments).Trim() -split '\s+' | Where-Object { $_ })) {
        if ($token.StartsWith('-') -or $groups.Count -eq 0) { $groups.Add($token) } else { $groups[$groups.Count - 1] += ' ' + $token }
    }
    return $groups
}

$template = Get-Content -LiteralPath (Join-Path $inventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json -Depth 100
foreach ($job in @($template.Jobs)) {
    $name = [string]$job.Name
    $edition = if ($job.PSObject.Properties['PowerShellEdition'] -and $job.PowerShellEdition) { [string]$job.PowerShellEdition } else { 'PowerShell7' }
    $scriptLeaf = Split-Path -Path ([string]$job.ScriptPath) -Leaf
    $launcherPath = if ($job.PSObject.Properties['LauncherPath'] -and $job.LauncherPath) { [string]$job.LauncherPath } else { '' }
    if ($launcherPath) {
        # The orchestrator runs: cmd /c call "<LauncherPath>" <Arguments>, without -Tenant or -Connect.
        $path = Join-Path $inventoryRoot $launcherPath
        Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "Job $name LauncherPath not found: $launcherPath"
        $invocation = Get-LauncherInvocation -Path $path
        Assert-True ($invocation.ScriptLeaf -eq $scriptLeaf) "Job $name launcher runs $($invocation.ScriptLeaf), not $scriptLeaf."
        Assert-True ($invocation.Engine -eq $edition) "Job $name launcher uses $($invocation.Engine), the job declares $edition."
        Assert-True $invocation.ForwardsArguments "Job $name launcher does not forward %*: orchestrator Arguments would be dropped."
        Assert-True ((Get-Content -LiteralPath $path -Raw) -match 'SMARTM365_ORCHESTRATOR_TENANT' -and $invocation.Tenant -ne 'prod') "Job $name launcher ignores the orchestrator tenant."
        $duplicated = @(Get-ArgumentGroup $job.Arguments | Where-Object { (' ' + $invocation.FixedArguments + ' ').Contains(' ' + ($_ -split ' ')[0] + ' ') })
        Assert-True ($duplicated.Count -eq 0) "Job $name passes $($duplicated -join ', ') twice (template Arguments and launcher)."
        Assert-True (-not $manualLaunchers.Contains($name)) "Job $name has a LauncherPath; remove it from the manual launcher map."
        continue
    }
    Assert-True ($manualLaunchers.Contains($name)) "Job $name has no LauncherPath and no manual launcher decision in this test."
    $manual = $manualLaunchers[$name]
    if ($null -eq $manual) { continue }
    $path = Join-Path $launchersRoot $manual
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "Job $name manual launcher not found: $manual"
    $invocation = Get-LauncherInvocation -Path $path
    Assert-True ($invocation.ScriptLeaf -eq $scriptLeaf) "Job $name manual launcher runs $($invocation.ScriptLeaf), not $scriptLeaf."
    Assert-True ($invocation.Engine -eq $edition) "Job $name manual launcher uses $($invocation.Engine), the job declares $edition."
    foreach ($group in @(Get-ArgumentGroup $job.Arguments)) {
        Assert-True ((' ' + $invocation.FixedArguments + ' ').Contains(' ' + $group + ' ')) "Job $name manual launcher $manual does not pass '$group' (template Arguments: $($job.Arguments))."
    }
}
foreach ($name in $manualLaunchers.Keys) {
    Assert-True (@($template.Jobs | Where-Object Name -eq $name).Count -eq 1) "Manual launcher map names an unknown job: $name"
}
Write-Output ("PASS: {0} SmartInventory launcher alignment checks. No launcher or collector executed." -f $checks)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD9SmpQcsBEI2Tk
# P/HN7u2bVWAFP5Gug6nfi7M3FZ/YlKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAm+hR2XRV6YOFZLqZHhX7oo+hyPVhBpiPMju+3AJt1fMA0GCSqG
# SIb3DQEBAQUABIIBgEBUEOlqsZOE3d/UHhP6t0akKudB/OjSSIm0u1ZQ9ZSm8gkf
# q1pE5ku09vW4leU4lpzbGKFmj6SGq4dr6TDhGxy77FG1K6rvrv0d+eghGrXa6nAf
# nL0qgbrPfJWQYnytLWn4vz+VcQ33MVPKrh0KcrEouGkCJPDqiFMMcr5+AgqWtT/K
# VSTFQ8JjvTh3sYowa5Mmp8BKigVVzWKXUOX5hOtUJo2zHPXWcXv1H0Jl4AA1CeRN
# +ewVulKneVAipVYllbEGGIFIYdpIk3CC/7fWdsticTi9BLmJLU2eVcQWStAFRd27
# EPnXp+vyoXPNsfYbJC2CjC3ysThuoIIm5/x5TieHuLPBh7QqyxWPsis6kCSG2Qnu
# P18ceJLuHgEwSAxas6e1oQvMitLCJYHWcslhM2DNcZoF7dyXtfvPPEtwUqW83f6G
# ZryPn0siBx+UhQ7Kd5GdVSxpk23VkdF55dWGW39WCNuVd0LKMbyQ9eNGdFV0M7zW
# B0g83nowQubnc/0kBqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNTI2
# NDFaMC8GCSqGSIb3DQEJBDEiBCAFoKdKAOogcWIuxkIgOEbe8MhMJzQaRiWiT4nD
# 333XFzANBgkqhkiG9w0BAQEFAASCAgCs0hreUgLZZDcOYvuX80QtO+qoATYclRSw
# No0UuZP64lZ8t4MAA+d38AJpcQ+ElPqVdyKTF2DPI7PPlcttu5LFG23VuGzmQWlo
# DOeumDT9Z1O/2nWw5c0pbpYFEe/rSY8bsmb3myIOafgLhYprI9rzOWvnbGQRL23d
# J6QoH/uAzquPuEpZRchVZmUpy80yHE98uDR0Je0SqUWBNOKRQpVF7peBZoyBADVU
# Wkv0qcrJEg8q29tFWXcn3lqWApRQr4AidAiAjA2UFY7NNisbk9+4KpYU07+m27MZ
# nFgf5/IHpzUXtdv3PPe2vchh4jCqUvPQX44lRWgVhIvc0WG3X/gobJTjvPuxkJwi
# In4VOIzru29MtnpOphoQI2yeC3ZMGkvNiIAyAxZS0+09FqnxDCsg0SUHPaIXoMZN
# rlgyK9RCUEEZ+xeUWl2IOjP8ZCo14Ov5/gAX6QrUSKtTn96x87HbsQUgGyEphYnO
# Lk8xNqcOXfjqNYHMQoEhohTcht+bsYCGnn+IZwAGw76H8J2BEqGOjpvq71IHlgK7
# +O/NHGzZYd7eG44jWnMCicmv8k4Zi/EoX2iCQ/W2hYqUQoAH/JVq02gU5PCUZxeX
# 4fhd+R3/tgS0jENKIr0g9JbnutrQqM/DmNhAb4ii+v79cPdbgkCPLb9jFK+MfVmS
# mH31+ECDSw==
# SIG # End signature block
