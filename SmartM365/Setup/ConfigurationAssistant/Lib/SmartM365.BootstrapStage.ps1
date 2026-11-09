# Isolated bootstrap staging, version 0.2.0. Staging never invokes setup.
function New-SiBootstrapStagePlan {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)]$BootstrapPlan)
    if ($BootstrapPlan.UpdateExisting) { throw 'Updating an existing application by display name is not safe enough for isolated bootstrap.' }
    $profile=Read-SiPrivateJson (Join-Path $Root ("Config/Tenants/$($BootstrapPlan.ProfileKey).local"))
    if ($BootstrapPlan.ProfileSHA256 -eq 'ABSENT') {
        $pending=$BootstrapPlan.PendingProfile
        if ($null -eq $pending -or $pending.Root -ine $Root -or $pending.ProfileKey -cne $BootstrapPlan.ProfileKey) {throw 'Pending profile does not match the bootstrap.'}
        if ($profile.State -eq 'Missing') {
            $fresh=Get-SiTenantCreationPlan -Root $Root -ProfileKey $pending.ProfileKey -OrganizationKey $pending.OrganizationKey -EnvironmentKey $pending.EnvironmentKey
            if ($fresh.TemplateHash -cne $pending.TemplateHash) {throw 'Pending profile template changed.'}
            foreach ($edit in $pending.PendingEdits) {if ($edit.Step -eq 'share' -or $edit.Field -notin @(Get-SiEditableFields $edit.Step) -or $null -eq $fresh.Model.PSObject.Properties[$edit.Field]) {throw 'Invalid pending profile setting.'};$fresh.Model.($edit.Field)=ConvertTo-SiEditValue -Snapshot $edit -Value $edit.Value}
            $profile=[pscustomobject]@{State='Present';Value=$fresh.Model}
        } elseif ($profile.State -eq 'Present') {
            $expectedModel=($pending.Model | ConvertTo-Json -Depth 100) | ConvertFrom-Json
            foreach ($edit in $pending.PendingEdits) {if ($edit.Step -eq 'share' -or $edit.Field -notin @(Get-SiEditableFields $edit.Step) -or $null -eq $expectedModel.PSObject.Properties[$edit.Field]) {throw 'Invalid pending profile setting.'};$expectedModel.($edit.Field)=ConvertTo-SiEditValue -Snapshot $edit -Value $edit.Value}
            $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($expectedModel | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
            $sha=[Security.Cryptography.SHA256]::Create();try {$expected=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','')} finally {$sha.Dispose()}
            if (-not (Test-Path -LiteralPath $pending.Destination -PathType Leaf) -or (Get-FileHash -LiteralPath $pending.Destination -Algorithm SHA256).Hash -cne $expected) {throw 'The newly created profile differs from the reviewed template.'}
            $BootstrapPlan=New-SiBootstrapPlan -Root $Root -ProfileKey $BootstrapPlan.ProfileKey -Directory $BootstrapPlan.Directory -DisplayName $BootstrapPlan.DisplayName -OutsideAddress $BootstrapPlan.OutsideAddress -AutoSelectOutsideAddress $BootstrapPlan.AutoSelectOutsideAddress -UpdateExisting $BootstrapPlan.UpdateExisting -ConfigureTeams $BootstrapPlan.ConfigureTeams
        } else {throw 'The pending profile is invalid or conflicting.'}
    } elseif ($profile.State -ne 'Present') { throw 'The selected profile is unavailable.' }
    $path=Join-Path $Root ("Config/Tenants/$($BootstrapPlan.ProfileKey).local.json")
    if (Test-Path -LiteralPath ($path+'.txt') -PathType Leaf) { $path+='.txt' }
    if ($BootstrapPlan.ProfileSHA256 -ne 'ABSENT' -and (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $BootstrapPlan.ProfileSHA256) { throw 'Profile changed; prepare the review again.' }
    $app=$profile.Value.PSObject.Properties['AppId']
    if ($null -ne $app -and -not [string]::IsNullOrWhiteSpace([string]$app.Value)) {
        $id=[guid]::Empty
        if (-not [guid]::TryParse([string]$app.Value,[ref]$id) -or $id -ne [guid]::Empty) { throw 'An existing application identifier requires an identity-pinned update workflow.' }
    }
    $files=@('Setup/SmartM365-Create-AppRegistration.ps1','Config/SmartM365-TenantContext.ps1','Config/SmartM365-JsonTransport.policy.psd1','Config/SmartM365.global.local.json.template','Config/Tenants/tenant.local.json.template','Config/Tenants/tenant.mail.local.json.template','Modules/SmartM365.Core/SmartM365.JsonTransport.psd1','Modules/SmartM365.Core/SmartM365.JsonTransport.psm1')
    $hashes=[ordered]@{}
    foreach ($file in $files) {
        $source=Join-Path $Root $file
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw ('Missing bootstrap dependency: '+$file) }
        $hashes[$file]=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    }
    if ($hashes[$files[0]] -cne $BootstrapPlan.SetupSHA256) { throw 'Bootstrap script changed; prepare the review again.' }
    [pscustomobject]@{Root=[IO.Path]::GetFullPath($Root); ProfileKey=$BootstrapPlan.ProfileKey; ProfileHash=$BootstrapPlan.ProfileSHA256; Files=$hashes; BootstrapPlan=$BootstrapPlan}
}

function New-SiBootstrapStage {
    param([Parameter(Mandatory)]$Plan,[Parameter(Mandatory)][string]$Parent)
    if (-not [IO.Path]::IsPathFullyQualified($Parent) -or -not (Test-Path -LiteralPath $Parent -PathType Container)) { throw 'An existing absolute staging parent is required.' }
    $fresh=New-SiBootstrapStagePlan -Root $Plan.Root -BootstrapPlan $Plan.BootstrapPlan
    if ($fresh.ProfileHash -eq 'ABSENT') {throw 'The pending profile must be created before staging.'}
    foreach ($file in $fresh.Files.Keys) { if ($fresh.Files[$file] -cne $Plan.Files[$file]) { throw 'A bootstrap dependency changed; review again.' } }
    $parentItem=Get-Item -LiteralPath $Parent -Force
    $drive=[IO.DriveInfo]::new([IO.Path]::GetPathRoot($parentItem.FullName))
    if ($drive.DriveType -notin @([IO.DriveType]::Fixed,[IO.DriveType]::Removable)) { throw 'Staging requires a local drive.' }
    $ancestor=$parentItem
    while ($null -ne $ancestor) {
        if (($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Staging paths cannot traverse reparse points.' }
        $ancestor=$ancestor.Parent
    }
    $stage=Join-Path $parentItem.FullName ('SmartM365-bootstrap-'+[guid]::NewGuid().ToString('N'))
    $security=[Security.AccessControl.DirectorySecurity]::new()
    $security.SetAccessRuleProtection($true,$false)
    $inherit=[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit'
    foreach ($sid in @([Security.Principal.WindowsIdentity]::GetCurrent().User, [Security.Principal.SecurityIdentifier]::new('S-1-5-18'), [Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'))) {
        $security.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($sid,[Security.AccessControl.FileSystemRights]::FullControl,$inherit,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow))
    }
    [IO.FileSystemAclExtensions]::Create([IO.DirectoryInfo]::new($stage),$security)
    foreach ($file in $fresh.Files.Keys) {
        $source=Join-Path $fresh.Root $file; $destination=Join-Path $stage $file
        $null=[IO.Directory]::CreateDirectory((Split-Path -Parent $destination))
        $bytes=[IO.File]::ReadAllBytes($source)
        $sha=[Security.Cryptography.SHA256]::Create()
        try { $hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
        if ($hash -cne $fresh.Files[$file]) { throw 'A bootstrap dependency changed while staging; the incomplete staging directory was retained.' }
        $stream=[IO.File]::Open($destination,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try { $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
    }
    $global=Get-Content -LiteralPath (Join-Path $stage 'Config/SmartM365.global.local.json.template') -Raw | ConvertFrom-Json
    $tenant=Get-Content -LiteralPath (Join-Path $stage 'Config/Tenants/tenant.local.json.template') -Raw | ConvertFrom-Json
    $global.DefaultTenant=$fresh.ProfileKey
    $global.SmartM365RootPath='.'; $global.WorkspaceRootPath='{{SmartM365RootPath}}'; $global.SmartInventoryConfigRootPath=''
    $tenant.ProfileKey=$fresh.ProfileKey; $tenant.OrganizationKey='isolated'; $tenant.EnvironmentKey='bootstrap'; $tenant.TenantKey='isolated-bootstrap'
    $encoding=[Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText((Join-Path $stage 'Config/SmartM365.global.local.json.txt'),(($global | ConvertTo-Json -Depth 100)+[Environment]::NewLine),$encoding)
    [IO.File]::WriteAllText((Join-Path $stage ("Config/Tenants/$($fresh.ProfileKey).local.json.txt")),(($tenant | ConvertTo-Json -Depth 100)+[Environment]::NewLine),$encoding)
    $globalPath=Join-Path $stage 'Config/SmartM365.global.local.json.txt'
    $profilePath=Join-Path $stage ("Config/Tenants/$($fresh.ProfileKey).local.json.txt")
    [pscustomobject]@{Path=$stage; ScriptPath=(Join-Path $stage 'Setup/SmartM365-Create-AppRegistration.ps1'); ProfilePath=$profilePath; ProfileHash=(Get-FileHash -LiteralPath $profilePath -Algorithm SHA256).Hash; GlobalPath=$globalPath; GlobalHash=(Get-FileHash -LiteralPath $globalPath -Algorithm SHA256).Hash; StagePlan=$fresh; State='PreparedOnly'}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfiLdfs4g6UMTr
# D7whl680UT2HneEXmnsUC61LWzsCp6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINr9eWApGxc692Gwc7BXz01JtshEYGQDjhX1Jf4FHDJmMA0GCSqG
# SIb3DQEBAQUABIIBgDMqI4a+k56z9lkT9lWXeT1ynhJGIrA9+GnCUyj1RGfEVsFy
# 0DdD/TJkfWsanmRcX/yk+Ub7cgjezAK0GFiWOQ2O+Ca3QsWbBHKXNJIzdosOD2n+
# GnYsMD1srTfzVK1ZRNK+Diqo+5ch7ubGXCcOZUECqagTODb5TDhI/7GJJ1I7JRym
# AcZ2dNwY8z6liew0t9Qs7IM0AbznwXfNHTFeQ+0BBMuoqoFCfAMrh28cWBd1PwjC
# cJbT5aVx1KbCSr5Ygw0W1EG7K6IFzwtCh/ehMe7kRaRkVfUYMN0Z6HVmyXqXmhaS
# LPli0WSeQ5Ivfxb24MYfD0IibG+CVlZjQ4xjfgM8gzzB4XDDChPv2UYWj+t2JVN5
# 8PTbb829bL4etyO6VG52Hxxd9mrlCWzl3WL3MKJvD9F+Rb+s/wHwOGl801Gi4tME
# KUjPSyy9EYdWIEq3Rq8Z0hCzLNEwNW1dMIlVzo2BYiFsLqtftRW1h67ATU/OE2vt
# U5diqEVsmecC0jpTvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTdaMC8GCSqGSIb3DQEJBDEiBCAmFlMVKgdiNNC7oQ/ER1LNsYqICRhillmL3q4K
# aQFF2jANBgkqhkiG9w0BAQEFAASCAgAZF92TPgMz+D/OI3hBUX3LaurKgBxJtGg9
# q9+Q6tJVIqZblFVeX197JBdYjAcGKSgq/rx7j3bIxrjKtQHYGrIlceOFmguTlRCf
# IiS7xskPpO9w31oQNeQox3b2mpjE5BWE2GlkxIIgiO1ddX56MkqfNuUZZVRwKmyC
# hksQFTj02tcsZqyYr0wXr2oej3uK3V6IGPiG6bVF0Q7QKPz2R2R4+XPEVVLHxJm4
# 1JLCE/ypy+BWti9Tl3peVQCQxRXXhQNKfbQGD0U7zYwfC1GsTrumfMzsglWmp9YZ
# TKwdOSGmvPET0d3kPcHrwFHohchdwF6PQHs5vwldvbY8cyzjhbWRH5Llr89W00BJ
# 22GZE4TYGLCzrY3pXBa+gDtKWvdDpMkhMIudbXvWchSfooFhoYGP+CWdk5e32kqN
# ddPtwhwdrr2UEq9XYaXVksVmUyGPf5K+DMasiqwgoemkqVEuhZAS9nvkS5ZmPmZX
# cUqKu9j+W1xnj6kO7mCb86qkENdOoAt5y6hWo1FOAm5STI7ha8ZKry8tCPGWY5ay
# hSL+7wmeTvfnxUfex/0yH2QAz0jtqE1a8Q/+7sXXgiF6+mJuhSAw6TDXhHRSuG6r
# 0T0bkzY+AIl66FkOw34mfY8AkHYijd/eqf1RjHDBLyExXqf868pe8i4pb7fEL8YG
# oFIYMMckHw==
# SIG # End signature block
