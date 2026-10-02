#Requires -Version 7.0
<#
.SYNOPSIS
Offline native identity, application-grain and consumer regression tests.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.WorkplaceSource.psd1') -MinimumVersion '1.0.0' -Force
$cases = [Collections.Generic.List[object]]::new()
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-WorkplaceSource-Test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
function Assert-Equal($actual,$expected) { if ($actual -cne $expected) { throw "Expected '$expected', got '$actual'." } }
function Assert-Throws([scriptblock]$body,[string]$pattern) {
    $caught=$null; try { & $body | Out-Null } catch {$caught=$_}
    if ($null -eq $caught -or $caught.Exception.Message -notlike $pattern) { throw "Expected failure matching $pattern" }
}
function Case([string]$name,[scriptblock]$body) {
    try { & $body; $cases.Add([pscustomobject]@{Name=$name;Passed=$true;Error=''}) }
    catch { $cases.Add([pscustomobject]@{Name=$name;Passed=$false;Error=$_.Exception.Message}) }
}
function Read-Functions([string]$path,[string[]]$names) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    foreach($name in $names) {
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if ($null -eq $node) { throw "Missing function: $name" }
        Set-Item -Path "Function:script:$name" -Value ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
    }
}
try {
    $apps=@(
        [pscustomobject]@{TenantKey='test';AppId='a1';AppName=' Editor ';AppPublisher='Vendor';Platform='windows';AppVersion='1';DeviceCount=2;CollectionScope='AllPlatforms'},
        [pscustomobject]@{TenantKey='test';AppId='a2';AppName='editor';AppPublisher='vendor';Platform='WINDOWS';AppVersion='2';DeviceCount=2;CollectionScope='AllPlatforms'},
        [pscustomobject]@{TenantKey='test';AppId='a3';AppName='editor';AppPublisher='Other';Platform='windows';AppVersion='1';DeviceCount=1;CollectionScope='AllPlatforms'},
        [pscustomobject]@{TenantKey='test';AppId='a4';AppName='editor';AppPublisher='Vendor';Platform='ios';AppVersion='1';DeviceCount=0;CollectionScope='AllPlatforms'}
    )
    $relations=Join-Path $temporaryRoot 'relations.csv'
    @(
        [pscustomobject]@{TenantKey='test';AppId='a1';DeviceId='d1'},
        [pscustomobject]@{TenantKey='test';AppId='a1';DeviceId='d2'},
        [pscustomobject]@{TenantKey='test';AppId='a2';DeviceId='d2'},
        [pscustomobject]@{TenantKey='test';AppId='a2';DeviceId='d3'},
        [pscustomobject]@{TenantKey='test';AppId='a3';DeviceId='d1'}
    ) | Export-Csv -LiteralPath $relations -NoTypeInformation
    Case 'Same product across versions counts distinct devices, not four installs' {
        $footprint=Get-WorkplaceApplicationFootprint $apps $relations
        Assert-Equal $footprint[(Get-WorkplaceApplicationProductKey $apps[0])].Count 3
        Assert-Equal $footprint.Count 3
    }
    Case 'Publisher and platform separate products; casing and spaces do not' {
        Assert-Equal (Get-WorkplaceApplicationProductKey $apps[0]) (Get-WorkplaceApplicationProductKey $apps[1])
        Assert-Equal ((Get-WorkplaceApplicationProductKey $apps[0]) -eq (Get-WorkplaceApplicationProductKey $apps[2])) $false
        Assert-Equal ((Get-WorkplaceApplicationProductKey $apps[0]) -eq (Get-WorkplaceApplicationProductKey $apps[3])) $false
    }
    Case 'Application tenant identity is isolated' {
        $other=$apps[0].PSObject.Copy();$other.TenantKey='other'
        Assert-Equal ((Get-WorkplaceApplicationProductKey $apps[0]) -eq (Get-WorkplaceApplicationProductKey $other)) $false
    }
    Case 'Partial relations fail instead of inventing product/device counts' {
        Assert-Throws {Get-WorkplaceApplicationFootprint @($apps[0]) $relations} '*Orphan*'
        $wrong=$apps[0].PSObject.Copy();$wrong.DeviceCount=3
        $one=Join-Path $temporaryRoot 'one.csv'
        Import-Csv $relations | Where-Object AppId -eq 'a1' | Export-Csv $one -NoTypeInformation
        Assert-Throws {Get-WorkplaceApplicationFootprint @($wrong) $one} '*coverage mismatch*'
    }
    Case 'Duplicate application native IDs fail' { Assert-Throws {Get-WorkplaceApplicationFootprint @($apps[0],$apps[0]) $relations} '*Duplicate*' }
    Case 'Top-mode application evidence cannot masquerade as complete' {
        $bounded=$apps[0].PSObject.Copy(); $bounded | Add-Member RelationCollectionScope 'Top'
        Assert-Throws {Get-WorkplaceApplicationFootprint @($bounded) $relations} '*All-mode*'
    }
    $objects=@(
        [pscustomobject]@{TenantKey='test';ObjectGUID='user1';ObjectSID='S-1-5-21-1-100';PrimaryGroupID='513';ObjectClass='user';DistinguishedName='CN=NoUPN,DC=test'},
        [pscustomobject]@{TenantKey='test';ObjectGUID='server1';ObjectSID='S-1-5-21-1-200';PrimaryGroupID='';ObjectClass='computer';DistinguishedName='CN=Server,DC=test'},
        [pscustomobject]@{TenantKey='test';ObjectGUID='foreign1';ObjectSID='S-1-5-21-2-100';PrimaryGroupID='';ObjectClass='foreignSecurityPrincipal';DistinguishedName='CN=Foreign,DC=test'}
    )
    $group=[pscustomobject]@{TenantKey='test';ObjectGUID='group1';objectSid='S-1-5-21-1-513';Members='';MembersJson='["CN=NoUPN,DC=test","CN=Server,DC=test","CN=Foreign,DC=test","CN=Missing,DC=other"]'}
    Case 'AD native identities retain no-UPN, server and external members' {
        $rows=@(ConvertTo-WorkplaceADMembership @($group) $objects '2026-01-01T00:00:00Z')
        Assert-Equal $rows.Count 5
        Assert-Equal @($rows | Where-Object ResolutionStatus -eq 'UnresolvedOrExternal').Count 1
        Assert-Equal @($rows | Where-Object MemberObjectClass -eq 'computer').Count 1
        Assert-Equal @($rows | Where-Object MembershipKind -eq 'PrimaryGroup').Count 1
    }
    Case 'JSON membership preserves semicolons in native DNs' {
        $withSeparator=$objects[1].PSObject.Copy();$withSeparator.DistinguishedName='CN=Server\;A,DC=test'
        $g=$group.PSObject.Copy();$g.MembersJson=ConvertTo-Json -InputObject @($withSeparator.DistinguishedName) -Compress
        $rows=@(ConvertTo-WorkplaceADMembership @($g) @($withSeparator) '2026-01-01T00:00:00Z')
        Assert-Equal $rows.Count 1; Assert-Equal $rows[0].MemberObjectGUID 'server1'
    }
    Case 'Duplicate AD tenant/DN identity blocks ambiguous joins' {
        Assert-Throws {ConvertTo-WorkplaceADMembership @($group) @($objects[0],$objects[0]) 'now'} '*Duplicate*'
    }
    Case 'Empty Graph collection is valid' {
        Assert-Equal @(Get-WorkplaceGraphCollection 'https://graph.microsoft.com/beta/test' {param($uri) @{value=@()}}).Count 0
    }
    Case 'Graph malformed, cyclic and untrusted pagination fail' {
        Assert-Throws {Get-WorkplaceGraphCollection 'https://graph.microsoft.com/beta/test' {param($uri) @{other=@()}}} '*Malformed*'
        Assert-Throws {Get-WorkplaceGraphCollection 'https://graph.microsoft.com/beta/test' {param($uri) @{value=@();'@odata.nextLink'=$uri}}} '*Cyclic*'
        Assert-Throws {Get-WorkplaceGraphCollection 'https://example.com/test' {throw 'Should not be called'}} '*Untrusted*'
    }
    Case 'Policy retains native payload, family and scope' {
        $p=ConvertTo-WorkplacePolicyEvidence ([pscustomobject]@{Id='p1';Name='Policy';TemplateReference=@{templateId='template'}}) SettingsCatalog r now
        Assert-Equal $p.PolicyId 'p1';Assert-Equal $p.PolicyFamily 'SettingsCatalog'
        Assert-Equal (($p.NativeEvidenceJson | ConvertFrom-Json).TemplateReference.templateId) 'template'
        Assert-Throws {ConvertTo-WorkplacePolicyEvidence ([pscustomobject]@{name='Bad'}) SettingsCatalog} '*identity*'
    }
    $intelligence=Join-Path (Split-Path $root -Parent) 'SmartWorkplaceIntelligence/scripts/New-ApplicationInventoryEvidence.ps1'
    Read-Functions $intelligence @('Convert-ToInvariantDecimalText','Get-ApplicationProfile')
    Case 'Intelligence current profile uses three products and distinct device evidence' {
        $profile=Get-ApplicationProfile -Rows $apps -SnapshotDate '2026-01-01' -SnapshotWeek '2026-W01' -IncludeDetail -ProductFootprint (Get-WorkplaceApplicationFootprint $apps $relations)
        Assert-Equal $profile.Trend.Applications 3
        Assert-Equal @($profile.Detail | Where-Object 'Application Source ID' -eq 'a1')[0].'Product Distinct Devices' 3
    }
    Case 'Intelligence historical definitions remain numerically unchanged' {
        $profile=Get-ApplicationProfile -Rows $apps -SnapshotDate '2026-01-01' -SnapshotWeek '2026-W01' -LegacyDefinition
        Assert-Equal $profile.Trend.Applications 2
        Assert-Equal $profile.Trend.'Application Install Observations' 5
        Assert-Equal $profile.Trend.'Collection Scope' 'Legacy Windows'
    }
    Case 'Intelligence accepts explicit empty inventory' {
        $profile=Get-ApplicationProfile -Rows @() -SnapshotDate '2026-01-01' -SnapshotWeek '2026-W01' -IncludeDetail
        Assert-Equal $profile.Detail.Count 0; Assert-Equal $profile.Trend.Applications 0
    }
    Case 'Intelligence application entry point exports distinct footprint and qualified legacy history' {
        $data=Join-Path $temporaryRoot 'apps-data'
        $last=Join-Path $data 'DATA-LAST'; $week=Join-Path $data 'DATA-ALL/Intune/Applications/DiscoveredApps/WeeklyHistory/2026-W01'
        New-Item -ItemType Directory -Path $last,$week | Out-Null
        $apps | Export-Csv (Join-Path $last 'Intune_DiscoveredApps_Summary.csv') -NoTypeInformation
        Copy-Item -LiteralPath $relations -Destination (Join-Path $last 'Intune_DiscoveredApps_AppDeviceRelations.csv')
        $apps | Select-Object TenantKey,AppId,AppName,AppPublisher,Platform,AppVersion,DeviceCount | Export-Csv (Join-Path $week 'Intune_DiscoveredApps_Summary.csv') -NoTypeInformation
        $detail=Join-Path $temporaryRoot 'detail.csv'; $trend=Join-Path $temporaryRoot 'trend.csv'
        & $intelligence -DataRoot $data -InventoryOutputPath $detail -TrendOutputPath $trend | Out-Null
        $output=@(Import-Csv $detail); Assert-Equal $output.Count 4
        Assert-Equal @($output | Where-Object 'Application Source ID' -eq 'a1')[0].'Product Distinct Devices' '3'
        Assert-Equal (Import-Csv $trend).'Metric Definition' 'Legacy name / version observations'
        Assert-Equal (Import-Csv $trend).Applications '2'
    }
    Read-Functions (Join-Path (Split-Path $root -Parent) 'SmartWorkplaceIntelligence/scripts/New-LicensingEvidence.ps1') @('Convert-ToDoubleOrNull')
    Case 'Intelligence missing or invalid license capacity remains unavailable' {
        foreach ($value in @($null,'','invalid','NaN','-1')) { Assert-Equal (Convert-ToDoubleOrNull $value) $null }
        Assert-Equal (Convert-ToDoubleOrNull '0') 0.0
        Assert-Equal (Convert-ToDoubleOrNull '15') 15.0
    }
    Read-Functions (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1') @('Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule','Assert-SmartM365CsvDataCompleteness','Add-SmartM365CsvValidationRule','Initialize-SmartM365DefaultCsvValidationRules')
    function WriteLog { param([string]$Message,[string]$Level) }
    $global:SmartM365CsvValidationRules=@{}; $global:SmartM365RequireCsvValidationRules=$true
    Initialize-SmartM365DefaultCsvValidationRules
    foreach ($sourceName in @('AD_DirectoryObjects_AllDomains','AD_GroupMemberships_AllDomains','M365_EntraGroupMemberships_All','M365_EntraGroupMembershipScope','Intune_Policies_All','Intune_PolicyAssignments_All')) {
        Case "Native CSV gate: empty schema, missing identity and duplicate keys / $sourceName" {
            $rule=Get-SmartM365CsvValidationRule -BaseFileName $sourceName
            $columns=@($rule.CriticalFields)+@($rule.RequiredColumns) | Select-Object -Unique
            Assert-SmartM365CsvDataCompleteness -BaseFileName $sourceName -Data @() -Columns $columns
            $row=[ordered]@{}; foreach($column in $columns){$row[$column]='evidence'}
            Assert-SmartM365CsvDataCompleteness -BaseFileName $sourceName -Data @([pscustomobject]$row) -Columns $columns
            Assert-Throws {Assert-SmartM365CsvDataCompleteness -BaseFileName $sourceName -Data @([pscustomobject]$row,[pscustomobject]$row) -Columns $columns} '*duplicate immutable*'
            $row[$rule.CriticalFields[0]]=''
            Assert-Throws {Assert-SmartM365CsvDataCompleteness -BaseFileName $sourceName -Data @([pscustomobject]$row) -Columns $columns} '*critical*'
        }
    }
    $cases | Format-Table -AutoSize
    $failed=@($cases | Where-Object Passed -eq $false)
    if ($failed.Count) { throw "$($failed.Count) workplace-source tests failed." }
    [pscustomobject]@{Cases=$cases.Count;Passed=$cases.Count;ProductionActions=0}
} finally {
    $resolved=[IO.Path]::GetFullPath($temporaryRoot)
    if ($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('SmartM365-WorkplaceSource-Test-')) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCANzbzxHk5hkoTP
# PkyhVpMpEqa+PgAQm1CBR8BAIWs9DqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGg7tlk1sPfGvNJbjbYWKdJyb0F4zYXnS50W/5u1NjrlMA0GCSqG
# SIb3DQEBAQUABIIBgFquZCjMW//USgPg68S4A3UjxOFS4gXEgb6e8q7e+bnQm9iY
# LW2BgjMmBRQ8ssrhgzx/2QwSZiXXVVyVjQZUFllk62ufexOjphhblwVYgDy9HL1W
# JyuMHKjpeluem9GNHkwkgzlIaOX6Kgi9A5fIkr1cm/3rGRO+vIXafQDLY6dwLLPv
# Uc6CMmdiPCqxJYbGFpk1nh9qeA+jk+c/xoSL4ha+Y52SLYNEPWcUDl6IU4K3QO9F
# ZQuoojFmIlTPluFxc51tuTOxQ8zZamOqcTqG1OUaMea3Slpc9keM5oMauHLXzdPt
# 4xTP+1zELVXNp1Bc7KwJg8TCQlzykuXtfQQDZo/fUdHSno9sYbQrQNmDxJ++ajis
# 7Sv/D5FliokGwGc1aYUKJDiSBBjMApcpgAuI1gjLD2siiunF3rxq9fWGJkrjWHB4
# bFzE77XC42hDCEJK+NGel2jqAqGYkQhXa1qdWn96tvKWdkWWuJ2E8NAwsGcz+FmO
# hmvx2G8USV9XgnrcH6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDIxOTQx
# MTBaMC8GCSqGSIb3DQEJBDEiBCAW1O1e8DkWYzv8p2l/kIU9qjaPgdx/2f6UQOkk
# hiusIjANBgkqhkiG9w0BAQEFAASCAgCFlorLwtpB7nlcaS/PyF532QX0hXqSpW0h
# lDY4tly/Sbfgh6j5aUIalZiYVSXynSvCKrSTm6oUs0cgdHvlSmRBr9VfuGZRBVX0
# Rasu1ebu5oIyCXyMeNGlA0fLiH76DToipCEPvE9dZr/xzx9kJXrJjG5w/D0Hk2EC
# 7B9HvBn1eN+XnKjctSlrkChk8LqxEr2FYLEVMEftBHP0KN6zMpr7HThLm+Wu/OZi
# Ok5UM6XSGqt2NfyNEH/1BYkIGGg91Lz87cty0yMO8WuP8+OfHKH6gkx9a3PgtzSb
# CyXBw5eM1Ieb4/RvFqvQafKrHYujeVvc/h8foaTP+wu2TDAjdKQ09z/WtKRlUEUm
# 7B8lk8P4JvXGVHtjv5RanNuIZR1dvNWQwnZc3TzrLSXk5lD9bVjXFsoO0kDJGnap
# fK7Fo2X/xi2U25JTlmxD8n6fiyjtaIPutpearufLcoXauYFfNzCiMlKLzr7huS5D
# 9yvZ40Xb7WrElqaOhjdYjreHoFsvuS+seUroZuihsmZHmpHTy6FBeHwomUh56JD0
# 2zgPzcdC6LV+2eAsCOS78kkfv3uyFC1cHEkQ2R2Vcpy0vsP1C9nblDObN9MN31Rr
# t672Gj+gfcSxuUdzPFagIkRot6yTOYnkmhtbjbTdA2i/VOPPbnoivnp3yNNeN033
# M+/oLBe3sA==
# SIG # End signature block
