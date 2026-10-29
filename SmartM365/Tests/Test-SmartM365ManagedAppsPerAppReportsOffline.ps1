#Requires -Version 7.0
<#
.SYNOPSIS
Verifies managed-app report filters and assigned-app status collection without Graph access.
.VERSION
1.2.1
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Join-Path $PSScriptRoot '..'
$commonPath = Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psm1'
$collectorPath = Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-IntuneManagedApps-Inventory.ps1'

function Get-FunctionText {
    param([string]$Path,[string]$Name)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count -gt 0) { throw "Parse errors in $Path." }
    $match = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name },$true))
    if ($match.Count -ne 1) { throw "Expected one definition of $Name in $Path." }
    return $match[0].Extent.Text
}

$reportText = Get-FunctionText $commonPath 'Invoke-SmartM365EvidenceExportReport'
$requestModule = New-Module -ScriptBlock ([scriptblock]::Create($reportText))
& $requestModule {
    function script:Invoke-SmartM365EvidenceGraphRequest {
        param($Method,$Uri,$Body)
        $script:Request = [pscustomobject]@{ Method=$Method; Uri=$Uri; Body=$Body }
        throw 'SyntheticStop'
    }
    foreach ($filter in @('',"(ApplicationId eq '11111111-1111-1111-1111-111111111111')")) {
        $script:Request = $null
        try { Invoke-SmartM365EvidenceExportReport -ReportName 'DeviceInstallStatusByApp' -Filter $filter -ApiVersion beta | Out-Null; throw 'Expected synthetic stop.' }
        catch { if ($_.Exception.Message -ne 'SyntheticStop') { throw } }
        $body = $script:Request.Body | ConvertFrom-Json
        if ($script:Request.Method -ne 'POST' -or $script:Request.Uri -ne 'https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs' -or $body.reportName -ne 'DeviceInstallStatusByApp') {
            throw 'Export request target or report name is wrong.'
        }
        if ($filter) {
            if ($body.filter -cne $filter) { throw 'Required application filter was lost.' }
        } elseif ($body.PSObject.Properties.Name -contains 'filter') {
            throw 'Unfiltered reports must retain their previous request shape.'
        }
    }
}

$definitions = @(
    (Get-FunctionText $commonPath 'Get-SmartM365EvidenceProperty'),
    (Get-FunctionText $collectorPath 'Get-IntuneManagedAppDeploymentState'),
    (Get-FunctionText $collectorPath 'Get-IntuneAssignedManagedApps'),
    (Get-FunctionText $collectorPath 'Get-IntuneManagedAppStatusSource'),
    (Get-FunctionText $collectorPath 'Get-IntuneManagedAppInstallStates')
)
$collectorModule = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
& $collectorModule {
    $script:Warnings = [Collections.Generic.List[string]]::new()
    function script:Write-SmartM365EvidenceLog { param($Message,$Level) if ($Level -eq 'WARNING') { $script:Warnings.Add([string]$Message) } }
    $identity = [pscustomobject]@{ TenantKey='contoso-test'; OrganizationKey='contoso'; EnvironmentKey='test'; TenantId='' }
    $ids = @(
        '11111111-1111-1111-1111-111111111111',
        '22222222-2222-2222-2222-222222222222',
        '33333333-3333-3333-3333-333333333333'
    )
    $apps = @($ids | ForEach-Object { [pscustomobject]@{ ApplicationId=$_ } })
    $assignments = @(
        [pscustomobject]@{ ApplicationId=$ids[0] },
        [pscustomobject]@{ ApplicationId=$ids[0] },
        [pscustomobject]@{ ApplicationId=$ids[2] }
    )
    $assignedApps = @(Get-IntuneAssignedManagedApps -Apps $apps -Assignments $assignments)
    if ($assignedApps.Count -ne 2 -or $assignedApps[0].ApplicationId -ne $ids[0] -or $assignedApps[1].ApplicationId -ne $ids[2]) {
        throw 'Assigned apps were not selected uniquely in catalog order.'
    }
    $unassignedOnly = @(Get-IntuneAssignedManagedApps -Apps $apps -Assignments @())
    if ($unassignedOnly.Count -ne 0) { throw 'An empty assignment set must not trigger installation reports.' }
    $assignedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $assignedApps) { [void]$assignedIds.Add($app.ApplicationId) }
    if ((Get-IntuneManagedAppStatusSource -ApplicationId $ids[0] -ObservedRows 1 -AssignedIds $assignedIds) -ne 'DeviceInstallStatusByApp' -or
        (Get-IntuneManagedAppStatusSource -ApplicationId $ids[2] -ObservedRows 0 -AssignedIds $assignedIds) -ne 'NoReportRows' -or
        (Get-IntuneManagedAppStatusSource -ApplicationId $ids[1] -ObservedRows 0 -AssignedIds $assignedIds) -ne 'NotAssigned') {
        throw 'Managed app status source does not distinguish assignment from absent report rows.'
    }
    $script:ProductionCalls = [Collections.Generic.List[object]]::new()
    function script:Invoke-SmartM365EvidenceExportReport {
        param($ReportName,$Filter,$Select,$ApiVersion,$TimeoutSeconds,$MaxItems)
        $script:ProductionCalls.Add([pscustomobject]@{ ReportName=$ReportName; Filter=$Filter; Select=@($Select); ApiVersion=$ApiVersion })
        $id = $Filter -replace "^\(ApplicationId eq '([0-9a-f-]+)'\)$",'$1'
        return [pscustomobject]@{ ApplicationId=$id; DeviceId='device-1'; InstallState='installed' }
    }
    $productionRows = @(Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $identity -PauseSeconds 0)
    if ($productionRows.Count -ne 2 -or $script:ProductionCalls.Count -ne 2 -or
        @($script:ProductionCalls | Where-Object { $_.ReportName -ne 'DeviceInstallStatusByApp' -or $_.ApiVersion -ne 'beta' -or @($_.Select).Count -lt 9 -or @($_.Select) -notcontains 'ApplicationId' -or @($_.Select) -notcontains 'InstallState' }).Count -gt 0) {
        throw 'Production report path did not query every assigned app with the required report and API version.'
    }
    $expectedColumns = 'TenantKey,OrganizationKey,EnvironmentKey,TenantId,ApplicationId,DeviceId,UserPrincipalName,AppVersion,Platform,InstallState,ReportedDeploymentState,InstallStateDetail,ErrorCode,LastModifiedDateTime,RunId,CollectedAtUtc'
    if (($productionRows[0].PSObject.Properties.Name -join ',') -cne $expectedColumns -or $productionRows[0].TenantKey -cne 'contoso-test' -or $productionRows[0].TenantId -cne '') {
        throw 'Device state rows must carry the tenant identity first, in the published column order.'
    }
    foreach ($missing in @('TenantKey','OrganizationKey','EnvironmentKey')) {
        $partial = $identity.PSObject.Copy(); $partial.$missing = ' '
        try { Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $partial -PauseSeconds 0 | Out-Null; throw 'Expected identity failure.' }
        catch { if ($_.Exception.Message -notlike '*are required for managed-app device states*') { throw } }
    }
    for ($index=0; $index -lt $assignedApps.Count; $index++) {
        if ($script:ProductionCalls[$index].Filter -cne "(ApplicationId eq '$($assignedApps[$index].ApplicationId)')") { throw 'Production report filter does not match the assigned app.' }
    }
    $script:Seen = [Collections.Generic.List[string]]::new()
    $exporter = {
        param($appId,$filter)
        $script:Seen.Add($appId)
        if ($filter -cne "(ApplicationId eq '$appId')") { throw 'Per-app filter does not match the requested catalog id.' }
        if ($appId -eq $ids[0]) { return [pscustomobject]@{ ApplicationId=$appId; DeviceId='device-1'; InstallState='3'; AppInstallState='Not installed' } }
        if ($appId -eq $ids[2]) { return [pscustomobject]@{ ApplicationId=$appId; DeviceId='device-3'; AppInstallState='failed'; HexErrorCode='0x1' } }
    }
    $rows = @(Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $identity -PauseSeconds 0 -ExportReport $exporter)
    if ($script:Seen.Count -ne 2 -or (@($script:Seen) -join ',') -cne (@($ids[0],$ids[2]) -join ',')) { throw 'An assigned app was skipped or an unassigned app was queried.' }
    if ($rows.Count -ne 2 -or $rows[0].ReportedDeploymentState -ne 'ReportedNotInstalled' -or $rows[0].InstallState -cne '3' -or $rows[1].ReportedDeploymentState -ne 'ReportedFailure' -or $rows[1].ErrorCode -ne '0x1' -or $script:Warnings.Count -ne 0) {
        throw 'Per-app device status mapping is wrong.'
    }
    $unknownExporter = { param($appId,$filter) [pscustomobject]@{ ApplicationId=$appId; InstallState='99'; AppInstallState='S99' } }
    $unknownRows = @(Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $identity -PauseSeconds 0 -ExportReport $unknownExporter)
    if ($unknownRows.Count -ne 2 -or @($unknownRows | Where-Object { $_.ReportedDeploymentState -ne 'Unknown' }).Count -gt 0 -or
        $script:Warnings.Count -ne 1 -or $script:Warnings[0] -notlike '*rows=2*' -or $script:Warnings[0] -notlike "*AppInstallState='S99', InstallState='99': 2*") {
        throw 'Unmapped source states were not kept Unknown and logged without row identities.'
    }
    $badExporter = { param($appId,$filter) [pscustomobject]@{ ApplicationId='99999999-9999-9999-9999-999999999999'; InstallState='installed' } }
    try { Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $identity -PauseSeconds 0 -ExportReport $badExporter | Out-Null; throw 'Expected app identity failure.' }
    catch { if ($_.Exception.Message -notlike '*mismatched ApplicationId*') { throw } }
    $failedExporter = { param($appId,$filter) if ($appId -eq $ids[2]) { throw 'Synthetic report failure' } }
    try { Get-IntuneManagedAppInstallStates -Apps $assignedApps -RunId 'run-1' -CollectedAtUtc '2026-10-13T00:00:00Z' -Identity $identity -PauseSeconds 0 -ExportReport $failedExporter | Out-Null; throw 'Expected report failure.' }
    catch { if ($_.Exception.Message -ne 'Synthetic report failure') { throw } }
}

$collectorText = Get-Content $collectorPath -Raw
if ($collectorText -notmatch "Name='Intune_ManagedAppDeviceStates';Rows=\`$deviceStates;Options=@\{NoTenantKey=\`$true;LocalSerialization=\`$true\};Columns=@\('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ApplicationId'" -or
    $collectorText -notmatch "SmartM365\.Core\.psd1'\) -MinimumVersion '1\.0\.113'" -or $collectorText -notmatch "EvidenceCollector\.Common\.psd1'\) -MinimumVersion '1\.0\.15'") {
    throw 'Device state publication must use prepared identity columns, local serialization and the matching module versions.'
}

Write-Output 'PASS: assigned app selection, status source, export filter body, state precedence, tenant identity columns, unknown diagnostics, identity mismatch, and report failure (offline).'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC80V2GeCyvLlZw
# zhl7mmfeuLJTdhFouV3JTlUPgS88V6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPpFU4ez+3NBTbBGC+OcLVsRz0CjjehzTcoxQVg0RFI/MA0GCSqG
# SIb3DQEBAQUABIIBgKBT9nvWp3Wp8KitYo4AGo87fVcDx32P3kwBLsoVQfPEHhJz
# GFl2moS9jbfrH7FdTxaUYOQu9Sdm46ify6/nn2GvjcRewC4i+NW0H71nyPZexRqK
# PtGatMDOE+vk5pr6qCc22OlAIwhxZICih0gENGkC7mJfH3SidFNSnZsrJaFL7mPw
# ptHVONYhuVjyoRR98bvbEBVRXNS8FdAcweJPbO6qgX/3zYkzCqb+ZdWT/Y1uY4Nk
# h0pp0U5U6APE3WzshEUcrlzv9OfpLnNH2pseRSp0HvK2f5gFXtK+adQGa3Zcn1/q
# fVm6mty3EzgXF7C+4Ywp7qCoyCuO2VIljLOUoqgJSe1r3/sf5TBpKpzGEb9w+PGN
# xWHZnLeeonDXo2OdnsL7ZlAzgtfg16j/OpZ0jq0in78boCcduhtz8zJwv0HJMCaD
# XmFG4u0+pvwFpZ0TKrnzJiNkcRK2EWy2psyN1Z0cV7sjSi4JKZFd+MgB9G3M4MM1
# fODXK20QazowJdCKDaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxMDEy
# NTVaMC8GCSqGSIb3DQEJBDEiBCClhEQOM+1X9mWW9oRAudpCT9umL/+Ikulhg+7h
# LcZUHjANBgkqhkiG9w0BAQEFAASCAgAE17ux5BPrcGcFJ5FgkpvPZ8KOXltGfYmt
# WBWKpoeuu4dMosvjVC+n7dgB1DZ2V9mHCzlxFM3iGL1I0U4aOmC+IH3tt5Opij1s
# 97GdJk0BdPaUs/On69B371r7w+Xn7QFA0X5dPwmal2QohWWqmOgqbodooPyaLEMj
# mu8I4pAFPdoJ7ZrbGflxsCW+eWToRcxS++C3eXfYhQBtn5ruo9FcGZ16Gb8DHJfb
# aySuoru5xQx9nJTuPggDIC5XCVNX70tLqCsUmkFr9wtl3ZQCWfKFhh3EUGOzNMya
# 53YrSunpBarcJCjMs0fcU27YaeqixlIpLMMa3KoSCkSX0ukjL/8c4PbsQEZldSPR
# r07o4XSxMHpndUZe7dePO8BqkHrLM6DgxpMcU3eHMLfTtXODQeyMDQtKOB9+wQe3
# 0A+X+2kzMQMyb3cMtM2RQ1g0RWnkd1lPrTxUJBPbHilLnkNu0fXl6UUmx/7D71gP
# 4nXlAt4/Dac7gxutG7hgRFjgLSSjKQF4NjyYapLF/tpcyw4vJf0UntSzGXzCVD4M
# s48xbRVy5VgzIQhCtqdxrizUp99FkL2j0DyPkzReXj8Cgz+TrCB7Ut/IRwr5YY3V
# vRXyBUzOFha7JQUCsZ3ztOm7P97wwYalbbywgmI5RoodtW0FRA6LPs8klZsqRt25
# aJYy4tkleQ==
# SIG # End signature block
