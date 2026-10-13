#Requires -Version 7.0
<#
.SYNOPSIS
Collects Intune managed apps, assignments, and reported installation states.
.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
    [switch]$EnableConfiguredExternalActions,
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [ValidateRange(0,10000000)][int]$MaxItems = 0,
    [ValidateRange(30,7200)][int]$ExportTimeoutSeconds = 1200
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Runtime = $null
$failure = $null

function Get-IntuneManagedAppDeploymentState {
    param([string]$InstallState)
    switch ($InstallState.Trim().ToLowerInvariant()) {
        'installed' { return 'ReportedInstalled' }
        { $_ -in @('failed','error') } { return 'ReportedFailure' }
        'notinstalled' { return 'ReportedNotInstalled' }
        { $_ -in @('pendinginstall','installing','pending') } { return 'Pending' }
        'notapplicable' { return 'NotApplicable' }
        default { return 'Unknown' }
    }
}

try {
    $smartRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.8' -Force
    $script:Runtime = Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effective `
        -DefaultOutputRelativePath 'Intune/WorkplaceScope/ManagedApps' -OutputPath $OutputPath `
        -LatestCsvFolderPath $LatestCsvFolderPath -ValidateOnly:$ValidateOnly `
        -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath `
        -ReadOnly:($ValidateOnly -or $MaxItems -gt 0)
    Connect-SmartM365EvidenceGraph $script:Runtime @('DeviceManagementApps.Read.All','DeviceManagementManagedDevices.ReadWrite.All') -InteractiveAuth:$InteractiveAuth
    Invoke-SmartM365Preflight -ScriptName 'SmartM365-IntuneManagedApps-Inventory' `
        -OutputPaths @($script:Runtime.OutputPath,$script:Runtime.LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('DeviceManagementApps.Read.All','DeviceManagementManagedDevices.ReadWrite.All') `
        -GraphProbeUris @('https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?$top=1') | Out-Null
    if ($ValidateOnly) {
        Write-SmartM365EvidenceLog 'Managed app prerequisites validated. Report availability requires a collection run.'
        return
    }

    $base = 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps'
    $nativeApps = @(Get-SmartM365EvidenceGraphCollection -Uri $base -MaxItems $MaxItems)
    $apps = [Collections.Generic.List[object]]::new()
    $assignments = [Collections.Generic.List[object]]::new()
    $seenApps = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($app in $nativeApps) {
        $id = [string](Get-SmartM365EvidenceProperty $app @('id'))
        if (-not $id -or -not $seenApps.Add($id)) { throw 'Managed app catalog contains a missing or duplicate app id.' }
        $apps.Add([pscustomobject][ordered]@{
            ApplicationId=$id
            DisplayName=[string](Get-SmartM365EvidenceProperty $app @('displayName'))
            Publisher=[string](Get-SmartM365EvidenceProperty $app @('publisher'))
            AppType=[string](Get-SmartM365EvidenceProperty $app @('@odata.type'))
            DisplayVersion=[string](Get-SmartM365EvidenceProperty $app @('displayVersion','version'))
            CreatedDateTime=[string](Get-SmartM365EvidenceProperty $app @('createdDateTime'))
            LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $app @('lastModifiedDateTime'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        })
        $encoded = [uri]::EscapeDataString($id)
        $nativeAssignments = @(Get-SmartM365EvidenceGraphCollection -Uri "$base/$encoded/assignments")
        $seenAssignments = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($assignment in $nativeAssignments) {
            $assignmentId = [string](Get-SmartM365EvidenceProperty $assignment @('id'))
            $target = Get-SmartM365EvidenceProperty $assignment @('target') $null
            if (-not $assignmentId -or -not $seenAssignments.Add($assignmentId) -or $null -eq $target) { throw "Invalid assignment for app $id." }
            $assignments.Add([pscustomobject][ordered]@{
                ApplicationId=$id; AssignmentId=$assignmentId
                Intent=[string](Get-SmartM365EvidenceProperty $assignment @('intent'))
                TargetType=[string](Get-SmartM365EvidenceProperty $target @('@odata.type'))
                GroupId=[string](Get-SmartM365EvidenceProperty $target @('groupId'))
                AssignmentFilterId=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterId'))
                AssignmentFilterType=[string](Get-SmartM365EvidenceProperty $target @('deviceAndAppManagementAssignmentFilterType'))
                RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
            })
        }
    }
    $aggregateRaw = @(Invoke-SmartM365EvidenceExportReport -ReportName 'OrgAppsInstallStatus' -ApiVersion v1.0 -TimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
    $deviceRaw = @(Invoke-SmartM365EvidenceExportReport -ReportName 'OrgDeviceInstallStatus' -ApiVersion v1.0 -TimeoutSeconds $ExportTimeoutSeconds -MaxItems $MaxItems)
    $aggregate = @($aggregateRaw | ForEach-Object {
        $id = [string](Get-SmartM365EvidenceProperty $_ @('ApplicationId'))
        if (-not $id) { throw 'App installation aggregate lacks ApplicationId.' }
        [pscustomobject][ordered]@{
            ApplicationId=$id; AppVersion=[string](Get-SmartM365EvidenceProperty $_ @('AppVersion'))
            Platform=[string](Get-SmartM365EvidenceProperty $_ @('Platform'))
            InstalledDeviceCount=[string](Get-SmartM365EvidenceProperty $_ @('InstalledDeviceCount'))
            FailedDeviceCount=[string](Get-SmartM365EvidenceProperty $_ @('FailedDeviceCount'))
            PendingInstallDeviceCount=[string](Get-SmartM365EvidenceProperty $_ @('PendingInstallDeviceCount'))
            NotInstalledDeviceCount=[string](Get-SmartM365EvidenceProperty $_ @('NotInstalledDeviceCount'))
            NotApplicableDeviceCount=[string](Get-SmartM365EvidenceProperty $_ @('NotApplicableDeviceCount'))
            InstalledUserCount=[string](Get-SmartM365EvidenceProperty $_ @('InstalledUserCount'))
            FailedUserCount=[string](Get-SmartM365EvidenceProperty $_ @('FailedUserCount'))
            PendingInstallUserCount=[string](Get-SmartM365EvidenceProperty $_ @('PendingInstallUserCount'))
            NotInstalledUserCount=[string](Get-SmartM365EvidenceProperty $_ @('NotInstalledUserCount'))
            NotApplicableUserCount=[string](Get-SmartM365EvidenceProperty $_ @('NotApplicableUserCount'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        }
    })
    $deviceStates = @($deviceRaw | ForEach-Object {
        $id = [string](Get-SmartM365EvidenceProperty $_ @('ApplicationId'))
        if (-not $id) { throw 'App device installation status lacks ApplicationId.' }
        $installState = [string](Get-SmartM365EvidenceProperty $_ @('InstallState'))
        [pscustomobject][ordered]@{
            ApplicationId=$id; DeviceId=[string](Get-SmartM365EvidenceProperty $_ @('DeviceId'))
            UserPrincipalName=[string](Get-SmartM365EvidenceProperty $_ @('UserPrincipalName'))
            AppVersion=[string](Get-SmartM365EvidenceProperty $_ @('AppVersion'))
            Platform=[string](Get-SmartM365EvidenceProperty $_ @('Platform'))
            InstallState=$installState
            ReportedDeploymentState=Get-IntuneManagedAppDeploymentState $installState
            InstallStateDetail=[string](Get-SmartM365EvidenceProperty $_ @('InstallStateDetail'))
            ErrorCode=[string](Get-SmartM365EvidenceProperty $_ @('ErrorCode'))
            LastModifiedDateTime=[string](Get-SmartM365EvidenceProperty $_ @('LastModifiedDateTime'))
            RunId=$script:Runtime.RunId; CollectedAtUtc=$script:Runtime.CollectedAtUtc
        }
    })
    if ($MaxItems -gt 0) {
        Write-SmartM365EvidenceLog "Sample complete: apps=$($apps.Count); assignments=$($assignments.Count); aggregate=$($aggregate.Count); deviceStates=$($deviceStates.Count). No canonical CSV published."
        return
    }
    $rules = $global:SmartM365CsvValidationRules
    if ($rules -isnot [hashtable]) { throw 'Core CSV validation rules are unavailable.' }
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedApps' -CriticalFields @('ApplicationId','RunId','CollectedAtUtc') -UniqueFields @('ApplicationId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppAssignments' -CriticalFields @('ApplicationId','AssignmentId','RunId','CollectedAtUtc') -UniqueFields @('ApplicationId','AssignmentId') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppInstallSummary' -CriticalFields @('ApplicationId','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    Add-SmartM365CsvValidationRule -Rules $rules -BaseFileName 'Intune_ManagedAppDeviceStates' -CriticalFields @('ApplicationId','RunId','CollectedAtUtc') -AllowEmptyDataset -CriticalMissingFailMinRows 1 -CriticalMissingFailPercent 0
    $exports = @(
        @{Name='Intune_ManagedApps';Rows=$apps.ToArray();Columns=@('ApplicationId','DisplayName','Publisher','AppType','DisplayVersion','CreatedDateTime','LastModifiedDateTime','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppAssignments';Rows=$assignments.ToArray();Columns=@('ApplicationId','AssignmentId','Intent','TargetType','GroupId','AssignmentFilterId','AssignmentFilterType','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppInstallSummary';Rows=$aggregate;Columns=@('ApplicationId','AppVersion','Platform','InstalledDeviceCount','FailedDeviceCount','PendingInstallDeviceCount','NotInstalledDeviceCount','NotApplicableDeviceCount','InstalledUserCount','FailedUserCount','PendingInstallUserCount','NotInstalledUserCount','NotApplicableUserCount','RunId','CollectedAtUtc')},
        @{Name='Intune_ManagedAppDeviceStates';Rows=$deviceStates;Columns=@('ApplicationId','DeviceId','UserPrincipalName','AppVersion','Platform','InstallState','ReportedDeploymentState','InstallStateDetail','ErrorCode','LastModifiedDateTime','RunId','CollectedAtUtc')}
    )
    foreach ($export in $exports) { Assert-SmartM365CsvDataCompleteness -Data $export.Rows -BaseFileName $export.Name -Columns $export.Columns | Out-Null }
    foreach ($export in $exports) { Export-SmartM365EvidenceDataset $script:Runtime $export.Name $export.Rows $export.Columns | Out-Null }
    Write-SmartM365EvidenceLog "Managed apps collected: apps=$($apps.Count); assignments=$($assignments.Count); aggregate=$($aggregate.Count); deviceStates=$($deviceStates.Count)." -Level SUCCESS
} catch { $failure = $_; throw } finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    if ($script:Runtime) { Complete-SmartM365EvidenceRuntime -Status $(if ($failure) {'Failed'} else {'Success'}) -ErrorRecord $failure -FailureStage 'IntuneManagedAppsInventory' }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWSH7nX7FJl7+Y
# 2dTXWxXEXDpdXzRM4yNsfucnzUJZ/KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC43U/ktl2B/Qb+M+a5v4N3aq6aGiGG/ngvPxFKLaXh/MA0GCSqG
# SIb3DQEBAQUABIIBgDyRv7ABlPZWbb6bzIIcMJ0kC6qBtMtv4CUjoUJdIniPfkTN
# /OZ5O4P5MAxUkVhExbBmduAmAgaqRrZ2a/QH4b/DZvdV5H5nHDTObZT5PvaDK3Ba
# PU+bzYA3cunCBsGJ3trQFlTklX28ofCXGtRxqGEI50hVaLjm7VRx0Yo47apyDNMA
# kLg45rUjGGiGGou52DLmaoCP0LsIXg2EeQNdhTtndiNeu9tzNT0Un6dNccXR7/At
# nMlv2ONuArXQC5pJ+4Z1cd1XdOApH8Gx038tQ/DmqbS+s0J7ziCiR8OM3j8ERUvY
# 8OlPvbuEh3fYIwpuuvBMDNTmMOjaDrE+WAJxJSJbA7bVM34sF2D+o4bPHgOibR7z
# IOe+1Ur07JGWTLJ46ZYpyOvHYXzpAhODp1G5zQuXmPv6BfaMRhCf9S946DQmwVGi
# ZCr+AM4N6jikRXiZB4G6YJJDvAX8boom/IsxpSFQr6Cxq3Bd4pyi4ZeYsMqwCbu4
# Jx4xxe03CaXIiwrUCaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTMxMzEw
# MjRaMC8GCSqGSIb3DQEJBDEiBCBqVOEOcga0z+fpEhlbq4NfObdQf29A5Sb1V47B
# Gd2azzANBgkqhkiG9w0BAQEFAASCAgBPJiVB9dfVby2+HtzNq71a9zqKh2l4Iqir
# cNyZXcDaQOychcw55+idbfgUDYrJEsoy9qPBflb+jo/Bx8rS41rDbegr0bmDundz
# JAi3ZS+0eyeLiW+zccZ9IOgQgAFkY5oJVnNntMipweXAimWJrAUkPRPG/DZHypDV
# v6Kcyxq+i89fkvK2/XEQOFcEUUbTNPLkqsYBe1+l6PHl3HI2b5lD3DU7bDNoxAMe
# l18eLqNlSsX9Rg6lARbskIWyMwbt04Rg/Rus8SlkAic8rqZe2JC94AknMzwtala4
# X7vhvUjyKtVGYhPggyntdmb3BCZXoNfvA+ZrDU51a946o52lBQVr3y50aFzXqjST
# mCitj7hVmI8xV9AolNskMIckFbUwRasQS0iU5azRxXRKzlpDMRZnaBbUexc7EIqr
# cCVLeIbivnhe6IKcNkb+qdfBHxogiiWHOG1WInv0TeJgbvqyeRtBEQNWzHdPGlzS
# avNzjrlhzW5iCiShWMemKrhRk3gBqnemAd7vbHhwtUOtWYHSgtecDqnoOhMVvNFQ
# oAJZ+b2v1YxgUKY1JoAgr5GKeUI+kXXQo/5dxtc6WPJdwqKdW1V+0g/0aSX1QEVR
# RV2KCfMJUGZYrn3RYBb3ah5oHVfbGBdWGFZVXAaOLoZGdQsY3IinRg69FNToutCX
# eKP+r7TGOw==
# SIG # End signature block
