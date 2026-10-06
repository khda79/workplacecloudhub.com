#Requires -Version 7.0
<#
.SYNOPSIS
Offline coherent group catalog/membership acquisition tests. No tenant calls.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$collectorPath = Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1'
$collector = Get-Content -LiteralPath $collectorPath -Raw
$start = $collector.IndexOf('    $memberships = ')
$end = $collector.IndexOf('    Write-SmartM365EvidenceLog "Workplace scope collected.', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Collection block boundaries changed; update this offline test.' }
$collectionBlock = [scriptblock]::Create($collector.Substring($start, $end - $start))
$script:checks = 0
function Assert-Catalog([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Invoke-CatalogFixture([object[]]$Groups, [switch]$FailMembers, [switch]$DuplicateMember, [switch]$FailPolicies) {
    $fixture = New-Module -ScriptBlock {
        param($rows, $block, $memberFailure, $duplicate, $policyFailure)
        Set-StrictMode -Version Latest
        $script:Runtime = [pscustomobject]@{RunId='synthetic-runtime-run'}
        $script:Groups = @($rows); $script:Block = $block
        $script:FailMembers = $memberFailure; $script:DuplicateMember = $duplicate
        $script:FailPolicies = $policyFailure
        $script:Exports = @{}; $script:GroupCalls = 0; $script:MemberCalls = 0
        $script:Properties = @(); $script:Scope = ''; $script:ErrorText = ''
        function Get-MgBetaGroup {
            [CmdletBinding()] param([switch]$All, [string[]]$Property)
            if (-not $All) { throw 'Full group enumeration required.' }
            $script:GroupCalls++; $script:Properties = @($Property)
            $script:Groups
        }
        function Get-MgBetaGroupMember {
            [CmdletBinding()] param([string]$GroupId, [switch]$All)
            if (-not $All) { throw 'Full membership enumeration required.' }
            $script:MemberCalls++
            if ($script:FailMembers) { throw 'Synthetic membership failure' }
            if ($GroupId -eq 'group-one') {
                [pscustomobject]@{Id='user-one';OdataType='#microsoft.graph.user'}
                [pscustomobject]@{Id=$(if ($script:DuplicateMember) {'user-one'} else {'principal-one'});OdataType='#microsoft.graph.servicePrincipal'}
            }
        }
        function Get-WorkplaceSourceValue {
            param($Row, [string[]]$Names, $Default='')
            foreach ($name in $Names) {
                $property = $Row.PSObject.Properties[$name]
                if ($null -ne $property) { return $property.Value }
            }
            return $Default
        }
        function Write-SmartM365EvidenceLog { param($Message, $Level) }
        function Invoke-SmartM365Preflight { throw 'Unexpected preflight in synthetic visible group collection.' }
        function Get-MgBetaDeviceManagementConfigurationPolicy {
            [CmdletBinding()] param([switch]$All)
            if ($script:FailPolicies) { throw 'Synthetic policy failure' }
        }
        function Get-MgBetaDeviceManagementDeviceConfiguration { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementDeviceCompliancePolicy { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementWindowsFeatureUpdateProfile { [CmdletBinding()] param([switch]$All) }
        function Get-MgBetaDeviceManagementWindowsQualityUpdateProfile { [CmdletBinding()] param([switch]$All) }
        function Assert-SmartM365CsvDataCompleteness { param($Data, $BaseFileName, $Columns) }
        function Export-SmartM365EvidenceDataset {
            param($Runtime, $Name, $Rows, $Columns, [switch]$NoWeeklyHistory, [switch]$SingleSerialization)
            if (-not $NoWeeklyHistory -or -not $SingleSerialization) { throw 'Export optimization/history contract changed.' }
            $script:Exports[$Name] = [pscustomobject]@{Rows=@($Rows);Columns=@($Columns)}
        }
        function Set-SmartM365CmdbSourceScope { param($CompleteScope, $Scope, $Qualifications) $script:Scope=$Scope }
        function Invoke-SyntheticCatalog {
            $MaxItems = 0
            try { & $script:Block } catch { $script:ErrorText = $_.Exception.Message }
            [pscustomobject]@{Exports=$script:Exports;GroupCalls=$script:GroupCalls;MemberCalls=$script:MemberCalls;
                             Properties=$script:Properties;Scope=$script:Scope;ErrorText=$script:ErrorText}
        }
    } -ArgumentList @($Groups, $collectionBlock, [bool]$FailMembers, [bool]$DuplicateMember, [bool]$FailPolicies)
    try { & $fixture { Invoke-SyntheticCatalog } }
    finally { Remove-Module $fixture -ErrorAction SilentlyContinue }
}
$group = [pscustomobject]@{Id='group-one';Visibility='Private';DisplayName="Group, quoted`nsecond line";
    MailEnabled=$false;SecurityEnabled=$true;GroupTypes=@('DynamicMembership','Unified');OnPremisesSecurityIdentifier='S-1-5-21-1-2-3-513'}
$empty = [pscustomobject]@{Id='group-empty';Visibility='Public';DisplayName='Empty group';MailEnabled=$true;
    SecurityEnabled=$false;GroupTypes=@();OnPremisesSecurityIdentifier=$null}
$result = Invoke-CatalogFixture @($group, $empty)
Assert-Catalog (-not $result.ErrorText) 'Complete synthetic group collection failed.'
Assert-Catalog ($result.GroupCalls -eq 1 -and $result.MemberCalls -eq 2) 'Catalog caused an additional group or membership enumeration.'
Assert-Catalog ($result.Exports.Count -eq 4) 'Existing four CSV exports changed.'
Assert-Catalog ($result.Scope -ceq 'CMDB:group_scope,group_members,policies,policy_assignments') 'Consumer receipt scope changed.'
foreach ($property in @('id','visibility','displayName','mailEnabled','securityEnabled','groupTypes','onPremisesSecurityIdentifier')) {
    Assert-Catalog ($result.Properties -contains $property) "Native property not requested: $property"
}
$scope = $result.Exports['M365_EntraGroupMembershipScope']
$members = $result.Exports['M365_EntraGroupMemberships_All'].Rows
Assert-Catalog ($scope.Rows.Count -eq 2 -and $members.Count -eq 2) 'Group or member rows were lost.'
$first = $scope.Rows[0]
Assert-Catalog ($first.DisplayName -ceq $group.DisplayName) 'Native multiline display name changed.'
Assert-Catalog (-not $first.MailEnabled -and $first.SecurityEnabled) 'Native false/true flags changed.'
Assert-Catalog ($first.GroupTypes -ceq 'DynamicMembership;Unified') 'Native group types changed.'
Assert-Catalog ($first.OnPremisesSecurityIdentifier -ceq $group.OnPremisesSecurityIdentifier) 'Native synchronized SID changed.'
Assert-Catalog ($first.MemberCount -eq 2 -and $scope.Rows[1].MemberCount -eq 0) 'Exact or empty member counts changed.'
Assert-Catalog ($scope.Rows[1].GroupTypes -ceq '' -and $null -eq $scope.Rows[1].OnPremisesSecurityIdentifier) 'Missing evidence was invented.'
foreach ($member in $members) {
    Assert-Catalog ($member.RunId -ceq $first.RunId -and $member.CollectedAtUtc -ceq $first.CollectedAtUtc) 'Catalog and members do not share row lineage.'
}
Assert-Catalog ([datetimeoffset]$first.GroupCollectedAtUtc -le [datetimeoffset]$first.CollectedAtUtc) 'Catalog date is after membership acquisition.'
foreach ($column in @('GroupId','Visibility','MemberCount','MemberCollectionStatus','RunId','CollectedAtUtc',
                     'DisplayName','MailEnabled','SecurityEnabled','GroupTypes','OnPremisesSecurityIdentifier','GroupCollectedAtUtc')) {
    Assert-Catalog ($scope.Columns -contains $column) "Export column missing: $column"
}
$zero = Invoke-CatalogFixture @()
Assert-Catalog (-not $zero.ErrorText -and $zero.Exports['M365_EntraGroupMembershipScope'].Rows.Count -eq 0) 'Successful zero collection failed.'
Assert-Catalog ($zero.Exports['M365_EntraGroupMembershipScope'].Columns.Count -eq 12) 'Empty enriched schema lost headers.'
foreach ($case in @(
    @{Groups=@($group,$group);Pattern='duplicate Entra group'},
    @{Groups=@($group);FailMembers=$true;Pattern='Synthetic membership failure'},
    @{Groups=@($group);DuplicateMember=$true;Pattern='duplicate Entra member'},
    @{Groups=@($group);FailPolicies=$true;Pattern='Synthetic policy failure'}
)) {
    $pattern=$case.Pattern; $arguments=$case.Clone(); $arguments.Remove('Pattern')
    $failed=Invoke-CatalogFixture @arguments
    Assert-Catalog ($failed.ErrorText -match $pattern -and $failed.Exports.Count -eq 0) 'Failed/duplicate acquisition published canonical evidence.'
}
Write-Output "PASS: $script:checks offline catalog/membership checks. No collectors, APIs, mail or live writes."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCTojqe8ylT/WD+
# mHOBXQgrcLSkyj2Es6iXxbe2hiaGxqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINlR6ePpQz1dFXoik2KDNyoqwHsM9Gr0K80Fp+EO7iRqMA0GCSqG
# SIb3DQEBAQUABIIBgHAI8P/fSV03Z2IXP3vrBTNIcMGOhSr+xTgFeGj7swPENt5N
# FVL8zvgxPvyiNfKjp6ntR/C20JFN7rkD1XUDFmRk3EYzcC/QwWZ7H7NK6m4HZfFh
# v3Uh7QIbqb5dxYXI16H1BbBL77drxhebGi0OoIan2WN1HI3prQ/uNpeL28QYEDA1
# 9GfC5Z/Ked7BfEKVS1cL8QVQTBoVL5Id4XoUhoqveAvjHoidl6Dlw7ArDuFb5U91
# NuQ4xPgT6IcnVIIr0W6Y8PKI5RTn96GrMALrRxD66tbiQxOC7M61NkYjERAU3fXV
# DVohRkhFAuEiTWoZbaqczI7tGSbMD4nJbU0oXmI1QnA8gJtAv9PPHSOh89uJQ7vN
# ZZKb9+Ld06yYpljTIMfqxH+g2mGrhRpQk3Q1WgUhO+bXpiF2wZ+kne1fsm1cjlCN
# +4MrsRf58B31RWKg8rvwucB95Ucx3KwGs86EE0cAtfLgsTik3dio/ApgTtpUlFQt
# 9Fn/O6WVLGcfaAsz56GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYwNzMx
# MDhaMC8GCSqGSIb3DQEJBDEiBCC3GlzCnXQoycYnwaILTnIARgsRJnrgV8NRHxDc
# 2vdOGjANBgkqhkiG9w0BAQEFAASCAgBYvXS/97ojsBbAI0dTWc5hUJaJ/NIFLUNl
# GTSoqqcjcyAfYEucBRb21xrM/gAcgLfeXYYe08MBPgajGHb8fxX08GeqZHigMSIJ
# DpeiRhyyU0ZNr96yWo8c7LPK4lqIMoh9bczrdlzsTfsyBsvS7hzvefyb8X0dYDWL
# PwYIhlFnbLCZRVuOQxd9GuzT92jCHC/JZUyNzIENp6wsmNZ1h55Lpm27kGNlsLlu
# Sd2Q3Ifqf5jGnXp+VYxDDp27ihpBay3Eet9YS9svj0054NfL5fjkwAAkWeC2tDOU
# NPwqfmb/DUkpyxHdg7+qgZWjDhfCJO2jRdDXMowKjV1pGll50fM2STq2F8J+DUgC
# KriDk+tydCaezpjEZhL2ggLWGdwfFnPcxq4IR78lNu7ifuRs0v5m2Adj0nl/yEhJ
# SL764T0RSenFyJGMADxCUJJ9m02918ltY6b/nVMnQO4rJ7g4ye8hQQMYUHZ3Vg31
# hm3p/VjcFcg6oYD9fZJkOSh8r45C42YzYTu0wCY7OfZtLCm/ozaMnByUNCvaz5SN
# HRlHjHaQXTBMzz30UFtWXJ0ZL6soEfMDX8zfxXLC6af/rfrDPEfAPxarSXz+Qu8j
# U4kR/QrW1cRAO8e+1j5GjrEgy6yHm4Bin1/oNxxCas1ML/f5E3e4ivHP29YrNWDw
# ESenVuYJ+g==
# SIG # End signature block
