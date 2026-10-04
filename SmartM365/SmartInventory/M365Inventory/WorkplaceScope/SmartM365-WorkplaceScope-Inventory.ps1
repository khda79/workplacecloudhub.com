#Requires -Version 7.0
<#
.SYNOPSIS
Collect native Entra direct membership and Intune policy/assignment evidence.
.VERSION
1.0.2
.NOTES
Candidate collector. Not automatically scheduled until production qualification.
Graph beta membership avoids the documented v1.0 service-principal omission.
Hidden groups require Member.Read.Hidden; missing access is fatal, not empty.
#>
[CmdletBinding()]
param([string]$Tenant='test', [switch]$InteractiveAuth, [switch]$ValidateOnly,
    [switch]$EnableConfiguredExternalActions, [string]$OutputPath, [string]$LatestCsvFolderPath,
    [ValidateRange(0,10000000)][int]$MaxItems=0)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Runtime = $null; $failure = $null; $transcriptStarted = $false
try {
    $smartRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
    . (Join-Path $smartRoot 'Config/SmartM365-TenantContext.ps1')
    $effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.65' -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot '../../Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.4' -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot '../../Common/SmartM365.WorkplaceSource.psd1') -MinimumVersion '1.0.0' -ErrorAction Stop
    $script:Runtime = Initialize-SmartM365EvidenceRuntime -ScriptPath $PSCommandPath -EffectiveConfig $effective `
        -DefaultOutputRelativePath 'M365/WorkplaceScope' -OutputPath $OutputPath -LatestCsvFolderPath $LatestCsvFolderPath `
        -ValidateOnly:$ValidateOnly -EnableConfiguredExternalActions:$EnableConfiguredExternalActions
    Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $script:Runtime.LatestCsvFolderPath -ReadOnly:($ValidateOnly -or $MaxItems -gt 0)
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    $transcriptStarted = $true
    foreach ($module in @('Microsoft.Graph.Beta.Groups','Microsoft.Graph.Beta.DeviceManagement')) {
        if (-not (Get-Module -ListAvailable $module)) { throw "Required SDK module missing. Install-Module $module -Scope CurrentUser" }
        Import-Module $module -ErrorAction Stop
    }
    Connect-SmartM365EvidenceGraph $script:Runtime @('Group.Read.All','DeviceManagementConfiguration.Read.All') -InteractiveAuth:$InteractiveAuth
    Invoke-SmartM365Preflight -ScriptName 'SmartM365-WorkplaceScope-Inventory' -OutputPaths @($script:Runtime.OutputPath) `
        -RequiredGraphApplicationPermissions @('Group.Read.All','DeviceManagementConfiguration.Read.All') `
        -GraphProbeUris @('https://graph.microsoft.com/beta/groups?$top=1','https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$top=1') | Out-Null
    if ($ValidateOnly) { Write-SmartM365EvidenceLog 'Workplace scope prerequisites validated; no full collection or CSV publication.'; return }
    if ($MaxItems -gt 0) {
        $global:SmartM365MaxItems=$MaxItems; $global:SmartM365TestMaxItems=$MaxItems; $global:SmartM365IsMaxItemsRun=$true
    }
    $memberships = [Collections.Generic.List[object]]::new()
    $groupScope = [Collections.Generic.List[object]]::new()
    $groups = @(Get-MgBetaGroup -All -Property id,visibility -ErrorAction Stop)
    if ($MaxItems -gt 0) { $groups = @($groups | Select-Object -First $MaxItems) }
    if (@($groups | Where-Object Visibility -eq 'HiddenMembership').Count) {
        Invoke-SmartM365Preflight -ScriptName 'SmartM365-WorkplaceScope-Inventory' -RequiredGraphApplicationPermissions @('Member.Read.Hidden') | Out-Null
    }
    Write-SmartM365EvidenceLog "Reading direct memberships for $($groups.Count) Entra groups. This can be a long-running inventory."
    foreach ($group in $groups) {
        if (-not $group.Id) { throw 'Entra group identity missing.' }
        $members = @(Get-MgBetaGroupMember -GroupId $group.Id -All -ErrorAction Stop)
        $collected = [datetime]::UtcNow.ToString('o')
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($member in $members) {
            if (-not $member.Id -or -not $seen.Add([string]$member.Id)) { throw 'Missing or duplicate Entra member identity.' }
            $memberships.Add([pscustomobject][ordered]@{
                GroupId=$group.Id; MemberId=$member.Id
                MemberType=Get-WorkplaceSourceValue $member @('@odata.type','OdataType')
                MembershipKind='Direct'; CollectionStatus='Collected'
                RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
            })
        }
        $groupScope.Add([pscustomobject][ordered]@{
            GroupId=$group.Id; Visibility=$group.Visibility; MemberCount=$seen.Count
            MemberCollectionStatus='Collected'; RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
        })
    }
    $policies = [Collections.Generic.List[object]]::new()
    $assignments = [Collections.Generic.List[object]]::new()
    $families = @(
        @{Name='SettingsCatalog';Command='Get-MgBetaDeviceManagementConfigurationPolicy';Path='configurationPolicies'},
        @{Name='DeviceConfiguration';Command='Get-MgBetaDeviceManagementDeviceConfiguration';Path='deviceConfigurations'},
        @{Name='DeviceCompliance';Command='Get-MgBetaDeviceManagementDeviceCompliancePolicy';Path='deviceCompliancePolicies'},
        @{Name='WindowsFeatureUpdate';Command='Get-MgBetaDeviceManagementWindowsFeatureUpdateProfile';Path='windowsFeatureUpdateProfiles'},
        @{Name='WindowsQualityUpdate';Command='Get-MgBetaDeviceManagementWindowsQualityUpdateProfile';Path='windowsQualityUpdateProfiles'}
    )
    foreach ($family in $families) {
        $items = @(& $family.Command -All -ErrorAction Stop)
        if ($MaxItems -gt 0) { $items = @($items | Select-Object -First $MaxItems) }
        $seenPolicies = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($policy in $items) {
            $id = [string](Get-WorkplaceSourceValue $policy @('id'))
            if (-not $id -or -not $seenPolicies.Add($id)) { throw 'Missing or duplicate policy identity.' }
            $uri = 'https://graph.microsoft.com/beta/deviceManagement/'+$family.Path+'/'+[uri]::EscapeDataString($id)+'/assignments'
            $targets = @(Get-WorkplaceGraphCollection -Uri $uri -Invoker { param($address) Invoke-SmartM365EvidenceGraphRequest -Uri $address })
            $collected = [datetime]::UtcNow.ToString('o')
            $policies.Add((ConvertTo-WorkplacePolicyEvidence $policy $family.Name $script:Runtime.RunId $collected))
            $seenAssignments = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($assignment in $targets) {
                $target = Get-WorkplaceSourceValue $assignment @('target') $null
                $assignmentId = [string](Get-WorkplaceSourceValue $assignment @('id'))
                if (-not $assignmentId -or -not $seenAssignments.Add($assignmentId) -or $null -eq $target) { throw 'Malformed policy assignment.' }
                $assignments.Add([pscustomobject][ordered]@{
                    PolicyFamily=$family.Name; PolicyId=$id; AssignmentId=$assignmentId
                    TargetType=Get-WorkplaceSourceValue $target @('@odata.type')
                    GroupId=Get-WorkplaceSourceValue $target @('groupId')
                    AssignmentFilterId=Get-WorkplaceSourceValue $target @('deviceAndAppManagementAssignmentFilterId')
                    AssignmentFilterType=Get-WorkplaceSourceValue $target @('deviceAndAppManagementAssignmentFilterType')
                    NativeTargetJson=ConvertTo-Json -InputObject $target -Depth 30 -Compress
                    RunId=$script:Runtime.RunId; CollectedAtUtc=$collected
                })
            }
        }
    }
    # No canonical publication until every required family and child collection succeeded.
    $exports = @(
        @{Name='M365_EntraGroupMemberships_All';Rows=$memberships.ToArray();Columns=@('GroupId','MemberId','MemberType','MembershipKind','CollectionStatus','RunId','CollectedAtUtc')},
        @{Name='M365_EntraGroupMembershipScope';Rows=$groupScope.ToArray();Columns=@('GroupId','Visibility','MemberCount','MemberCollectionStatus','RunId','CollectedAtUtc')},
        @{Name='Intune_Policies_All';Rows=$policies.ToArray();Columns=@('PolicyId','PolicyFamily','DisplayName','Description','Platforms','Technologies','FeatureUpdateVersion','CreatedDateTime','LastModifiedDateTime','NativeEvidenceJson','AssignmentCollectionStatus','RunId','CollectedAtUtc')},
        @{Name='Intune_PolicyAssignments_All';Rows=$assignments.ToArray();Columns=@('PolicyFamily','PolicyId','AssignmentId','TargetType','GroupId','AssignmentFilterId','AssignmentFilterType','NativeTargetJson','RunId','CollectedAtUtc')}
    )
    foreach ($export in $exports) { Assert-SmartM365CsvDataCompleteness -Data $export.Rows -BaseFileName $export.Name -Columns $export.Columns | Out-Null }
    foreach ($export in $exports) {
        Export-SmartM365EvidenceDataset $script:Runtime $export.Name $export.Rows $export.Columns -NoWeeklyHistory | Out-Null
    }
    Set-SmartM365CmdbSourceScope -CompleteScope ($MaxItems -eq 0) -Scope 'CMDB:group_scope,group_members,policies,policy_assignments' -Qualifications @('Direct membership only; all five required policy families collected.')
    Write-SmartM365EvidenceLog "Workplace scope collected. Groups=$($groups.Count); memberships=$($memberships.Count); policies=$($policies.Count); assignments=$($assignments.Count)." -Level SUCCESS
} catch { $failure=$_; throw } finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    if ($script:Runtime) { Complete-SmartM365EvidenceRuntime -Status $(if ($failure) {'Failed'} else {'Success'}) -ErrorRecord $failure -FailureStage 'WorkplaceScopeInventory' }
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAQzC+VyCkbXzuK
# 2JxNbBHaTLyimuMTH5GUZjNmF0hMpaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB7binBQNgltDkc3W1+Qq2d/Bnjgo3lTxTVZKnkWQqEjMA0GCSqG
# SIb3DQEBAQUABIIBgGuRFv8eAayYTqxvFWPUxWSnGVS8ja/AtlXY72i7oAvqsBSx
# E4itOXdRx9W+nelarZlYsCa3ZlReD2VeH7Y5eix1H6miZlleiNqqoQDut00YilTE
# MnttxTP3iG2UHRpV8LMsAM9skiEFz5TnO/Ys78v2ZJROTXNObb/18kw9n4hyvcIk
# YtHwgT+crnYDp1BfHjpQ+o0Ytw06LhqDeyK0i1DzxVX2GZl7htcXad99spvJDMOC
# nLki+gC/SfSciv8Tf8ciDDJVynkYUEHWmO/HYmTHbQljBddFrDDn4c91ytQ9Juyc
# zC1psEo+NQlN5El//HnVrjMTt190/2jDfuzoE1w6Go/QM2iKTRKy8FJYaDSDFgny
# TKCVsfM/Yw9GBwQXiPDxuT7f8xNHmWjDlgXR5KlAWOghDroYYsK5QdT3XiZJrLm5
# iAnnmZ1sfJiwrpGfHCclO8xT8Wyw5IoRdsSwjVLUGVrarV72lPdamlOPs6VGEbo/
# M14hMzngJYhxp522YaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQxMTAy
# MjVaMC8GCSqGSIb3DQEJBDEiBCDa+c03zEdHEV3Q13640Q3Ws+h4yxow3BUTjj5M
# M+UfIjANBgkqhkiG9w0BAQEFAASCAgBJaSIfbLmglRleC2BMENbLj+vcVN5jT4+k
# KA3SoFpQm52Gw84ZEX/WO2DSl4I75M2RLEN3pEAw2nk4YrnrwyyN6mgFfy4JTP4q
# mA37tZzQcNBVSsZeTG7hRE0QzE086G3P/5n2fjc69Zaw95SlAvcRopvqtJ2FNGFi
# 0TpwM4C5/B3bYXKKlt/+ZDFzfANiMtPNamlU6E7GmrZM3g5cNBxcZPjTcSZ8+1IM
# Wm5MgKA1NareTdmqTmGGkEuat3TCixYrkDN9Y/U6PhEJKY5rl5Up8CqgH11AWWSD
# Wi+9n3rIW6NBWJtrVVYFFReJzSf7AGgB1NvPNcUL5AZ2Ns38s149yz5TSD39/4+0
# Vo8tyjtTom92nRDcVtvBfskcJd7LsAHN9kLX9TwYZKjJrZU1QSD45qWJ3GRhdi8d
# oeAvuYk67U1odkhNsrJT6cZeVFrLzk4xkVf9BhEmB/q8TLjBTG9W4cymR2MCjWR4
# s5ixqJ2f+2CPhLmLybc3JjbzUefBtvJmKPmni15crVNCNK4IOYu219+giShr5OZX
# JzfgAMjCXpP4mw3SHQm1XAdzvxHk9X9hOXlTFIX3gO/9pUZuoIrRI/fUXXNpLEFQ
# ZA4qu+F3rp7Gqn7EW71JN7Tqv8X+U+3022sVLhVIrvfAocVZzOpApgdZMGODelif
# OCDoj6w7lw==
# SIG # End signature block
