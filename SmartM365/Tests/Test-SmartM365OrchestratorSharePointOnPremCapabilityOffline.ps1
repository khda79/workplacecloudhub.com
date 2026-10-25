#Requires -Version 7.0
<#
.SYNOPSIS
Checks the SharePointOnPrem orchestrator capability and server policy, offline.
.DESCRIPTION
Verifies that every capability list accepts SharePointOnPrem, that the static probe reports
it from the snap-in registration, that the read-only probe uses the collector snap-in and a
read-only farm call in Windows PowerShell, and that an OnlyJobsRequiring SharePointOnPrem
policy keeps the other jobs away from that server in the election plan.
.VERSION
1.0.1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1') -Force
$script:Checks = 0
function Assert-Capability {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

# Every hard-coded capability list must stay identical and include SharePointOnPrem.
$expected = "'SharedRuntime', 'Graph', 'EXO', 'AD', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell'"
$orchestratorText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'))
$managementText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Management.psm1'))
$guiText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
$distributedText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1'))
Assert-Capability ([regex]::Matches($orchestratorText, [regex]::Escape('$validCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident job capability list.'
Assert-Capability ([regex]::Matches($orchestratorText, [regex]::Escape('$validJobPolicyCapabilities = @(' + $expected + ')')).Count -eq 1) 'Resident server policy capability list.'
Assert-Capability ($managementText.Contains('$script:ValidCapabilities = @(' + $expected + ')')) 'Management capability list.'
Assert-Capability ($guiText.Contains('<ComboBoxItem Content="SharePointOnPrem"/>')) 'GUI server policy choice.'

# Publication validation accepts the capability for jobs and server policies.
$jobs = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
$sharePointJob = ($jobs.Jobs[0] | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
$sharePointJob.Name = 'SharePointOnPrem-Sample'
$sharePointJob.RequiredCapabilities = @('SharedRuntime', 'SharePointOnPrem')
$jobs.Jobs = @($jobs.Jobs) + $sharePointJob
$jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $jobs
Assert-Capability ($jobsValidation.Valid -and -not @($jobsValidation.Errors | Where-Object { $_ -like '*SharePointOnPrem*' }).Count) "Jobs validation: $($jobsValidation.Errors -join '; ')"
$cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Cluster.json.template')
$cluster.ExpectedOrchestratorServers = @('SRV01', 'SP01')
$cluster.ServerJobPolicies = [pscustomobject]@{ SP01 = [pscustomobject]@{ OnlyJobsRequiring = @('SharePointOnPrem') } }
$clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $cluster
Assert-Capability $clusterValidation.Valid "Cluster policy validation: $($clusterValidation.Errors -join '; ')"
$cluster.ServerJobPolicies = [pscustomobject]@{ SP01 = [pscustomobject]@{ OnlyJobsRequiring = @('SharePointOnline') } }
Assert-Capability (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'Unknown policy capability is still rejected.'

# Static probe reports SharePointOnPrem from the snap-in registration.
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPCapability-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $root
    $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $root -ProbeMode Static
    $probe = @($static.Results | Where-Object Name -eq 'SharePointOnPrem')
    $registered = [bool](& (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoLogo -NoProfile -NonInteractive -Command "if (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) { 'yes' }")
    Assert-Capability ($probe.Count -eq 1 -and $probe[0].Ready -eq $registered) 'Static probe result matches the snap-in registration.'
    Assert-Capability (($static.ReadyCapabilities -contains 'SharePointOnPrem') -eq $registered) 'ReadyCapabilities follows the probe.'
    Assert-Capability (@(Get-ChildItem -LiteralPath $root -Recurse -File).Count -eq 0) 'Static probe leaves no file behind.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }

# Read-only probe: Windows PowerShell, collector snap-in, read-only farm call, bounded timeout.
$readOnly = [regex]::Match($distributedText, "(?s)\`$sharePointScript = @'(?<Body>.*?)'@\s*\r?\n\s*(?<Call>\`$results\.Add\(\(Invoke-IsolatedCapabilityProbe -Name 'SharePointOnPrem'[^\r\n]*)")
Assert-Capability $readOnly.Success 'Read-only SharePoint probe exists.'
Assert-Capability ($readOnly.Groups['Call'].Value -match '-EnginePath \$windowsPowerShellPath' -and $readOnly.Groups['Call'].Value -match '-TimeoutSeconds \$ProbeTimeoutSeconds') 'Probe runs in bounded Windows PowerShell.'
$body = $readOnly.Groups['Body'].Value
Assert-Capability ($body -match 'Add-PSSnapin Microsoft\.SharePoint\.PowerShell' -and $body -match 'Get-SPFarm' -and $body -match 'SMARTM365_CAPABILITY_JSON') 'Probe loads the collector snap-in and reads the farm.'
Assert-Capability ($body -notmatch '(?i)\b(Set|New|Remove|Add|Update|Install|Enable|Disable|Start|Stop|Mount|Dismount)-SP') 'Probe never changes the farm.'
$tokens = $null; $errors = $null
[void][Management.Automation.Language.Parser]::ParseInput($body, [ref]$tokens, [ref]$errors)
Assert-Capability ($errors.Count -eq 0) 'Probe script parses.'

# Election: a SharePointOnPrem-only server receives only SharePoint on-premises jobs.
$sample = @(
    [pscustomobject]@{ Name = 'SP-Content'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'SP-Content'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime', 'SharePointOnPrem'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('02:00') } }
    [pscustomobject]@{ Name = 'Cloud-A'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'Cloud-A'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime', 'Graph'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('03:00') } }
    [pscustomobject]@{ Name = 'Cloud-B'; Enabled = $true; AssignmentMode = 'Elected'; Group = 'Cloud-B'; DependsOn = @(); RequiredCapabilities = @('SharedRuntime'); RequiredGraphAppRoles = @(); EstimatedDurationMinutes = 30; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('04:00') } }
)
$capabilities = @(
    [pscustomobject]@{ ServerName = 'SRV01'; ReadyCapabilities = @('SharedRuntime', 'Graph'); GraphAppRoles = @() }
    [pscustomobject]@{ ServerName = 'SP01'; ReadyCapabilities = @('SharedRuntime', 'Graph', 'SharePointOnPrem'); GraphAppRoles = @() }
)
$plan = Get-SmartM365OrchestratorElectionPlan -Jobs $sample -ServerCapabilities $capabilities -ServerWeights @{ SRV01 = 1.0; SP01 = 5.0 } -ServerJobPolicies @{ SP01 = @{ OnlyJobsRequiring = @('SharePointOnPrem') } }
$owners = @{}; foreach ($assignment in @($plan.Assignments)) { $owners[[string]$assignment.JobName] = [string]$assignment.OwnerServer }
Assert-Capability ($owners['SP-Content'] -eq 'SP01') 'SharePoint on-premises job elected on the SharePoint server.'
Assert-Capability ($owners['Cloud-A'] -eq 'SRV01' -and $owners['Cloud-B'] -eq 'SRV01') 'Policy keeps other jobs off the SharePoint server despite its weight.'
$noSharePoint = Get-SmartM365OrchestratorElectionPlan -Jobs @($sample[0]) -ServerCapabilities @($capabilities[0]) -ServerWeights @{ SRV01 = 1.0 } -ServerJobPolicies @{}
Assert-Capability (@($noSharePoint.Assignments).Count -eq 0 -and @($noSharePoint.UnassignedGroups).Count -eq 1) 'Without a ready SharePointOnPrem server the job stays unassigned.'

# Forced Manual jobs: the resident launches them only where every required capability is ready.
$orchestratorAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$tokens, [ref]$errors)
foreach ($name in @('Get-ManualJobMissingCapabilities', 'Test-JobAllowedOnServer')) {
    $definition = $orchestratorAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Assert-Capability ($null -ne $definition) "$name exists."
    . ([scriptblock]::Create($definition.Extent.Text))
}
$template = Read-SmartM365OrchestratorJson -Path (Join-Path $folder 'Orchestrator-Jobs.json.template')
foreach ($jobName in @('SharePointOnPrem-Infrastructure-Inventory', 'SharePointOnPrem-Content-Inventory')) {
    $job = @($template.Jobs | Where-Object Name -eq $jobName)[0]
    Assert-Capability ($job.AssignmentMode -eq 'Manual' -and @($job.DependsOn).Count -eq 0 -and @($job.RequiredCapabilities) -contains 'SharePointOnPrem') "$jobName is manual, independent and requires SharePointOnPrem."
    $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD') }
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job -AllowManual)) "$jobName refused by -Force without SharePointOnPrem."
    Assert-Capability ((Get-ManualJobMissingCapabilities -Job $job) -join ',' -eq 'SharePointOnPrem') "$jobName reports the missing capability."
    $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = @('SharedRuntime', 'SharePointOnPrem') }
    Assert-Capability (Test-JobAllowedOnServer -Job $job -AllowManual) "$jobName accepted by -Force on a SharePointOnPrem server."
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job)) "$jobName is never scheduled without -Force."
    $script:ServerCapabilities = $null
    Assert-Capability (-not (Test-JobAllowedOnServer -Job $job -AllowManual)) "$jobName fails closed when capabilities are unknown."
}
$script:ServerCapabilities = $null
Assert-Capability (Test-JobAllowedOnServer -Job ([pscustomobject]@{ Name = 'ManualNoCapability'; AssignmentMode = 'Manual'; RequiredCapabilities = @() }) -AllowManual) 'A manual job without required capabilities keeps its previous behaviour.'

# Absent components (snap-in not registered, module not installed) are logged as INFO, real failures as WARN.
foreach ($detail in @('SharePoint Server PowerShell snap-in Microsoft.SharePoint.PowerShell is not registered on this server.', 'Exchange Management Shell snap-in Microsoft.Exchange.Management.PowerShell.SnapIn is not registered on this server.', 'MicrosoftTeams PowerShell module is not installed on this server.')) {
    Assert-Capability (Test-SmartM365OrchestratorCapabilityComponentMissing -Result ([pscustomobject]@{ Ready = $false; Detail = $detail })) "Missing component: $detail"
    Assert-Capability ($distributedText.Contains($detail.TrimEnd('.'))) "Probe emits: $detail"
}
foreach ($detail in @('Cannot connect to the configuration database.', 'Read-only probe timed out after 90 seconds.', 'The specified module MicrosoftTeams with version 4.7.1 was not loaded because no valid module file was found in any module directory.')) {
    Assert-Capability (-not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result ([pscustomobject]@{ Ready = $false; Detail = $detail }))) "Real failure stays a warning: $detail"
}
Assert-Capability (-not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result ([pscustomobject]@{ Ready = $true; Detail = 'x is not installed on this server.' }))) 'A ready capability is never missing.'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPCapability-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $root
    $static = Get-SmartM365OrchestratorServerCapability -ServerName 'OFFLINE' -SharedDataFolderPath $root -ProbeMode Static
    foreach ($result in @($static.Results)) {
        Assert-Capability ($result.PSObject.Properties['ComponentMissing'] -and [bool]$result.ComponentMissing -eq (Test-SmartM365OrchestratorCapabilityComponentMissing -Result $result)) "ComponentMissing published for $($result.Name)."
    }
    $sharePointResult = @($static.Results | Where-Object Name -eq 'SharePointOnPrem')[0]
    Assert-Capability ($sharePointResult.Ready -or $sharePointResult.ComponentMissing) 'An unregistered SharePoint snap-in is reported as a missing component.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
$capabilityRefresh = $orchestratorAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-OrchestratorServerCapabilities' }, $true).Extent.Text
Assert-Capability ($capabilityRefresh.Contains('[bool]$probe.ComponentMissing) {') -and $capabilityRefresh.Contains('"Capability {0} not available on this server (component not installed): {1}" -f $probe.Name, $probe.Detail)') -and $capabilityRefresh.Contains('else { Write-OrchestratorLog -Message ("Capability {0} unavailable: {1}" -f $probe.Name, $probe.Detail) -Level WARN }')) 'Resident logs missing components as INFO and real failures as WARN.'

"[{0}] PASS: SharePointOnPrem capability and policy, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD+fDBASZzEgDi4
# DMEYCdHyhgMkXbsL0+vlCi+0OpNhA6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIF4vSBWAU0hGzQcmurSoshev9ZWhpw9ZB1gFRf0yES5MA0GCSqG
# SIb3DQEBAQUABIIBgAFA0JrIvUt1gYlFFMd6uQT1J+4gJWXmqo0bYnt+ujVhB7/F
# U2rlJVjpzpbUygIJDuRN09w2uYlXoyDQbOPhCQJH4Q8xVmufiZXdaP6KBStsLMRu
# 2A7A/op+7fzko05+Yt3qm4ejUgoIn19zoJ1tinTZX3uK4PKa8/5jm/8nE3pSKVGI
# Z5KzFOELehhnMta0mPA5f1cvBr/URzkBqQN8ssSBDEGJGHJSFoLGqJj1vYvptgHI
# btQeFDO2Zv7FkdfPyBKgPz4h1aDOGIsK11wcG0mwB1akveZT8drvw0Aa0ac3EWQB
# 9H694zOv51+LadSMXL4ZIhJn3bWKSkYFT83Zhh4Ta5alPUQkBRSYy9d0nsRlWVJZ
# UPcCwnw+pwNw9FuIKL/LSIIvMsQm3+jZQCZEPGhAIIRlg7Vul8tojxus2XviDt+m
# CS3TTV/yLei3/MTY7g4+VmZUL3d77/JOzImqKCXlE9SmS47eM0LPcdJk8JbAIf+u
# GDj7ugibFAtPQYGe8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUyMDIx
# NTdaMC8GCSqGSIb3DQEJBDEiBCBVPlrOYLqoxQlBdH3ZvvCHug9XZClnRduvYK02
# sExF5TANBgkqhkiG9w0BAQEFAASCAgCs+IYegeOajvhfoPP9CeJmZesOLWlMZgft
# PD1YBt19TQYJRTCzhREUFrsBvMraZyvy60iLGVd3ion4jcsZiq4ctew7Bnzor0fw
# Uvc8I5B8JaS9MGf4SjQUc+SAQchSwK1s1S6JXJyKZqaRucOroHMwuxvo0v7lfCTC
# rFVRMXD6dag7D6t2z0dlyzoijwnvdzxhLK+yF8i/pY74FOTFqlQdAVdAVoH1d/xP
# cFXHXYzpZIAtZaaHS5uVF/sxzHgUiapAvpW5BOD47xp5Vl5EVgFM04iKZwXGdpo/
# R3iCRV7B0bb6W/eq80uh+3gZ35rEFOb93mZz5cWJtGnmOFiBRzCFqWQ69CUOQ/Rb
# +26zyNskb1MzlQg6ZsvI0GHgsu+p+moLxzvTfHDXiJvrJ1Rt1rAD+pJr6FnUw+KB
# +aK2GgQ6aGPLdD0MfJ4q7xbppM3j77qjiNzwxyAtctc0huQeGKXfloMmEG1cFjpY
# mnTONYC1jz+VjJssr6XpwJA0NjcTp3mUGNhihSwrcCfseSWLEoLlCbleSw63521j
# 8IqjGCQaoNZoZ9REkuy+PA1gfGui6fyS+8ter34nbtoTOXvmoQIwcFvEoTU4yW8Z
# kWjct4iUs500h6zZXnOQMng8iXnt2eYr3LN3hUBMjGjpHxiNhCxUxesu458b1vPG
# VyvhxVHLmw==
# SIG # End signature block
