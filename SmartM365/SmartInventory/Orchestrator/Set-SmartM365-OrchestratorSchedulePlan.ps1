<#
.SYNOPSIS
Previews or publishes the committed template schedule plan in the shared Orchestrator jobs configuration.

.DESCRIPTION
For every shared job that also exists in Orchestrator-Jobs.json.template, the planning fields are
taken from the template: Schedule, TimeoutMinutes, DependsOn, ConcurrencyKey, DependencyMode and
DependencyMaxAgeHours (removed when the template has none). Jobs listed in -RemoveJob are removed.
Every other field (Arguments, Enabled, AssignmentMode, AllowedServers, retries, capabilities,
launcher, estimated duration) and every cluster setting are preserved. Shared jobs absent from the
template are kept unchanged and reported; template jobs absent from the shared configuration are
reported (the resident orchestrator appends them at startup). Enabled differences are reported and
kept, unless -SyncEnabled is used. Without -Execute, the script is read-only.
Publication uses the Orchestrator management module's validation, optimistic hash checks,
configuration lock, before/after version snapshots and audit CSV.

.PARAMETER SharedDataFolderPath
Shared tenant Orchestrator root containing Config, Versions and Audit, for example
\\server\share\DATA-ALL\Orchestrator.

.PARAMETER TemplatePath
Jobs template to apply. Defaults to Orchestrator-Jobs.json.template next to this script.

.PARAMETER RemoveJob
Shared jobs to remove. Defaults to Exchange2016-Local-Mailboxes-Fast (same script and arguments
as Exchange2016-Local-Mailboxes-Inventory).

.PARAMETER SyncEnabled
Also takes the Enabled flag from the template for the jobs it contains. The changed flags are
shown in the preview like the other fields.

.PARAMETER Execute
Publishes the displayed changes. Omit this switch for preview only.

.VERSION
1.1
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SharedDataFolderPath,

    [string]$TemplatePath = '',

    [string[]]$RemoveJob = @('Exchange2016-Local-Mailboxes-Fast'),

    [switch]$SyncEnabled,

    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner

$managementModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Management.psm1'
Import-Module -Name $managementModulePath -Force -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($TemplatePath)) { $TemplatePath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json.template' }
$template = Get-Content -LiteralPath $TemplatePath -Raw | ConvertFrom-Json -Depth 100
$templateValidation = Test-SmartM365OrchestratorJobsDocument -Document $template
if (-not $templateValidation.Valid) { throw "The template is invalid: $($templateValidation.Errors -join '; ')" }
$templateJobs = @{}
foreach ($templateJob in @($template.Jobs)) { $templateJobs[[string]$templateJob.Name] = $templateJob }

$snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $SharedDataFolderPath
$jobsDocument = $snapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
$changes = [System.Collections.Generic.List[object]]::new()
$notes = [System.Collections.Generic.List[string]]::new()

function ConvertTo-PlanText {
    # Canonical, readable text: property order and empty-versus-absent lists never count as changes.
    param([Parameter(Mandatory = $true)][string]$Name, [AllowNull()]$Value)
    if ($Name -eq 'DependsOn') {
        $names = @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Sort-Object)
        if ($names.Count -eq 0) { return '<none>' }
        return ($names -join ', ')
    }
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return '<none>' }
    if ($Name -eq 'Schedule') {
        $days = if ($Value.PSObject.Properties['DaysOfWeek'] -and [string]$Value.Type -eq 'Weekly') { ' ' + ((@($Value.DaysOfWeek) | ForEach-Object { ([string]$_).Substring(0, 3) }) -join '/') } else { '' }
        $missed = if ($Value.PSObject.Properties['MissedRunPolicy'] -and $Value.MissedRunPolicy) { [string]$Value.MissedRunPolicy } else { 'RunOnce' }
        return ('{0}{1} {2} [{3}]' -f $Value.Type, $days, ((@($Value.Times) | Sort-Object) -join ','), $missed)
    }
    return [string]$Value
}

function Merge-PlanField {
    param([Parameter(Mandatory = $true)]$Job, [Parameter(Mandatory = $true)][string]$Name, [AllowNull()]$Value)
    $current = if ($Job.PSObject.Properties[$Name]) { $Job.$Name } else { $null }
    $before = ConvertTo-PlanText -Name $Name -Value $current
    $after = ConvertTo-PlanText -Name $Name -Value $Value
    if ($before -ceq $after) { return }
    if ($null -eq $Value) { [void]$Job.PSObject.Properties.Remove($Name) }
    elseif ($Job.PSObject.Properties[$Name]) { $Job.$Name = $Value }
    else { $Job | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
    $changes.Add([pscustomobject][ordered]@{ Job = [string]$Job.Name; Field = $Name; Before = $before; After = $after })
}

$keptJobs = foreach ($job in @($jobsDocument.Jobs)) {
    $name = [string]$job.Name
    if ($name -in $RemoveJob) {
        $changes.Add([pscustomobject][ordered]@{ Job = $name; Field = '<job>'; Before = 'present'; After = 'removed' })
        continue
    }
    if (-not $templateJobs.ContainsKey($name)) {
        $notes.Add("Kept unchanged (absent from the template): $name")
        $job
        continue
    }
    $templateJob = $templateJobs[$name]
    foreach ($field in @('Schedule', 'TimeoutMinutes', 'DependsOn', 'ConcurrencyKey', 'DependencyMode', 'DependencyMaxAgeHours')) {
        $value = if ($templateJob.PSObject.Properties[$field]) { $templateJob.$field } else { $null }
        if ($field -eq 'DependsOn' -and $null -eq $value) { $value = @() }
        Merge-PlanField -Job $job -Name $field -Value $value
    }
    if ($SyncEnabled) { Merge-PlanField -Job $job -Name 'Enabled' -Value ([bool]$templateJob.Enabled) }
    elseif ([bool]$job.Enabled -ne [bool]$templateJob.Enabled) {
        $notes.Add(("Enabled differs and is kept: {0} shared={1}, template={2}" -f $name, [bool]$job.Enabled, [bool]$templateJob.Enabled))
    }
    $job
}
$jobsDocument.Jobs = @($keptJobs)
foreach ($name in @($templateJobs.Keys | Where-Object { $_ -notin @($jobsDocument.Jobs.Name) -and $_ -notin $RemoveJob } | Sort-Object)) {
    $notes.Add("Template job absent from the shared configuration (appended by the orchestrator at startup): $name")
}

$validation = Test-SmartM365OrchestratorJobsDocument -Document $jobsDocument
$consistency = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $jobsDocument -ClusterDocument $snapshot.Cluster
$errors = @($validation.Errors) + @($consistency.Errors)
if ($errors.Count -gt 0) { throw "The resulting configuration is invalid: $($errors -join '; ')" }

if ($changes.Count -eq 0) {
    Write-Output 'The shared jobs configuration already matches the template schedule plan.'
}
else {
    foreach ($group in @($changes | Group-Object Job | Sort-Object Name)) {
        Write-Output $group.Name
        foreach ($change in $group.Group) { Write-Output ("    {0,-22} {1}  ->  {2}" -f $change.Field, $change.Before, $change.After) }
    }
    Write-Output ''
    Write-Output ("{0} change(s) on {1} job(s)." -f $changes.Count, @($changes.Job | Sort-Object -Unique).Count)
}
foreach ($note in $notes) { Write-Output $note }
foreach ($warning in @($validation.Warnings) + @($consistency.Warnings)) { Write-Warning $warning }
if ($changes.Count -eq 0) { return }

if (-not $Execute) {
    Write-Warning 'Preview only. Re-run the same command with -Execute to publish these changes.'
    return
}

$result = Publish-SmartM365OrchestratorConfiguration `
    -SharedDataFolderPath $SharedDataFolderPath `
    -JobsDocument $jobsDocument `
    -ClusterDocument $snapshot.Cluster `
    -ExpectedJobsHash $snapshot.JobsHash `
    -ExpectedClusterHash $snapshot.ClusterHash `
    -ChangeSummary ("Apply template schedule plan: {0} change(s) on {1} job(s)" -f $changes.Count, @($changes.Job | Sort-Object -Unique).Count)

Write-Output ("Published configuration version {0}." -f $result.VersionId)
Write-Output ("New jobs hash: {0}" -f $result.JobsHash)
if (@($result.Warnings).Count -gt 0) {
    $result.Warnings | ForEach-Object { Write-Warning $_ }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDT54wqAzYvOR9V
# KJhEmIA8B4dIgkuPBVbLk44QKThNcKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGmtGq3Nzpr1tl8rl2C3EkQ875lV5XiIy/fsDbrAyvcjMA0GCSqG
# SIb3DQEBAQUABIIBgGQxYww4aFW3uYsCxlsP1vfObJpGApPgJKTOnMp1PqwLChm6
# u8ZDo5RLgUic7iQsyzLRi/VQdyvIGdsZe5CgDAjN7++RyyvaRc+u2CPr6U8Z1XtS
# n4GOY8suw1YvCR1xv5d6Ht25M9ChLAb/H0ULCdYPYMMn5HFQfSLfcsAThONaYbgj
# GYLRurMjDABQcz4MKdG+zoNF2cJmAnk6RlKz9tAOJol0hvHRzsi2hLIO4ePkA0Gn
# 8ff773Mi0J8WG9MLm/sh6HynIoAEmwyndN5HsODw8fy6zo8lvJ3nrfvNb+9hEizj
# npulNu5mcpMIqBrxK/g+Bc1CBeZW+1Z5US0hGIg8HAUJCp5MI4zzUhU8Koekc4P8
# kqjhUK/Xv0KoV0T++6hsXnqHg6QevQXYl0N4XjTDNe9PKvRYAAE3OBcRqeh4nV9+
# h0T9CkedDva2+bgug13KlmSFRIMEz1y2oNx9uHJw7+kuKm3YrqWgGQIhCCfk+Jc1
# NXkET3ZC/VHB2YNaaKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDEwNjA3
# NDRaMC8GCSqGSIb3DQEJBDEiBCB3m0Kfb+sCINa7jVyIFpqxfvwcxQvHWawBmcdH
# S6va6jANBgkqhkiG9w0BAQEFAASCAgClj2nVSrO4ktO6x3mQLMWXPoUwdi3Gl9w4
# Q9bvpap5qIgMiUuEkKKsCtzSIRTh+YcIy3W3+NoMDURzGV6YBS2nKfO1y3wrhE7D
# uxW2b90UjeS1Zd7a2swFjip2L4j8r/hQrSNtra30XOszijAV+tSkJ+MuLnty8XKq
# frpryLPZZFPu3tOiBI2LRmRWef8sOJQeAWOF/ux+8Njh+hB2Vos6O6SWfzSl1sY3
# 6x2G5XWf5C0Vp03HKEGD2VxRy3t7E+vB603n32G2z212EXr7t2h1NKSQVdYnPF/B
# RlqVOoUieP8EBzgV/X0AWOxR4IdorMOVG28JvPVfeix+ic2aOBUVtqAVen5OyD10
# PwtZjPfLvV1oe1bwXQW2N0VlpHyTeBeDuy2AbEHxZq60/Gma3soouFKiWUWampfI
# 7UmPCH0kAbkahYh+CPQFR0CjInNxNoUbOftStMzuAQtCnA8YviNJE/XyAQ9f8V2Z
# x8QElUhDb6ica7zNW0pSZT0GhbyZiZQ9CcTsD/CjDTbYd9vuw0KTDam2i/1F6l3V
# ZAh4cmZfamZ5gozlVCtHldZLbxEszsPlZqg9ikFcfrMwP/1DmwNGv9ngjmJPcxxJ
# EYbzKIwJSyUA5xksZptx8jGQo+ZE1+oKQ+TaRdLUZzIi9mETYk60p53ekzhQt4Q0
# 8HuwreXDnQ==
# SIG # End signature block
