# Passive workflow assessment, version 0.6.0. It does not run setup actions.
function Set-SiAssistantRepositoryRoot {
    param([Parameter(Mandatory)][string]$Root)
    $script:SiAssistantRepositoryUncRoot=''
    if (-not $Root.StartsWith('\\')) {return}
    if ($Root -notmatch '^\\\\[^\\/:*?]+\\[^\\/:*?]+(?:\\[^\\/:*?]+)*\\?\z' -or $Root -match '[\x00-\x1f\x7f]' -or @($Root.Split('\') | Where-Object {$_ -in @('.','..')}).Count) {throw 'The assistant repository UNC path is invalid.'}
    $script:SiAssistantRepositoryUncRoot=([IO.Path]::GetFullPath($Root)).TrimEnd('\')
}
function Assert-SiLocalAssessmentPath {
    param([Parameter(Mandatory)][string]$Path,[scriptblock]$GetAttributes,[switch]$AllowRepository)
    if (-not [IO.Path]::IsPathFullyQualified($Path)) { throw 'An absolute path is required.' }
    if (@(($Path -split '[\\/]') | Where-Object {$_ -in @('.','..')}).Count -or $Path -match '[*?\x00-\x1f\x7f]') {throw 'Path traversal or unsupported characters are not allowed.'}
    $full=[IO.Path]::GetFullPath($Path)
    if ($full.StartsWith('\\')) {
        if ($Path -match '[:/]') {throw 'UNC alternate streams or forward separators are not allowed.'}
        $trusted=[string]$script:SiAssistantRepositoryUncRoot
        if (-not $AllowRepository -or -not $trusted -or ($full.TrimEnd('\') -ine $trusted -and -not $full.StartsWith($trusted+'\',[StringComparison]::OrdinalIgnoreCase))) {throw 'UNC access is limited to the assistant repository when explicitly allowed.'}
    } else {
        $drive=[IO.DriveInfo]::new([IO.Path]::GetPathRoot($full))
        if ($drive.DriveType -notin @([IO.DriveType]::Fixed,[IO.DriveType]::Removable)) { throw 'The assessment requires a local drive or the assistant repository UNC path.' }
    }
    if ($null -eq $GetAttributes) {$GetAttributes={param($candidate) [IO.File]::GetAttributes($candidate)}}
    $driveRoot=[IO.Path]::GetPathRoot($full)
    $current=$driveRoot
    $parts=@($full.Substring($driveRoot.Length) -split '[\\/]' | Where-Object { $_ })
    foreach ($part in @('')+$parts) {
        if ($part) { $current=[IO.Path]::Combine($current,$part) }
        try { $attributes=& $GetAttributes $current }
        catch [IO.FileNotFoundException] { break }
        catch [IO.DirectoryNotFoundException] { break }
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Assessment paths cannot traverse reparse points.' }
    }
    $full
}
function Get-SiAssessmentFingerprint {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Tenant)
    $paths=@('Config/SmartM365.global.local.json','Config/SmartM365.global.local.json.txt',("Config/Tenants/$Tenant.local.json"),("Config/Tenants/$Tenant.local.json.txt"))
    @(
        foreach ($relative in $paths) {
            $path=Assert-SiLocalAssessmentPath (Join-Path $Root $relative) -AllowRepository
            if (Test-Path -LiteralPath $path -PathType Leaf) { [pscustomobject]@{Name=$relative;Hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash} }
            elseif (Test-Path -LiteralPath $path) { throw 'A configuration path is not a regular file.' }
            else { [pscustomobject]@{Name=$relative;Hash='ABSENT'} }
        }
    )
}
function Get-SiLocalWorkflowAssessment {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Tenant,[scriptblock]$CertificateSource,[scriptblock]$TaskSource,[string[]]$ModuleRoots=@($env:PSModulePath -split [IO.Path]::PathSeparator))
    if ($Tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$') { throw 'Invalid profile key.' }
    $safeRoot=Assert-SiLocalAssessmentPath $Root -AllowRepository
    if (-not (Test-Path -LiteralPath $safeRoot -PathType Container)) { throw 'Assessment root is unavailable.' }
    $before=@(Get-SiAssessmentFingerprint -Root $safeRoot -Tenant $Tenant)
    $rows=@(Get-SiStepAudit -SmartM365Root $safeRoot -Tenant $Tenant)
    if ($rows.Count -ne 15) { throw 'The workflow catalog is incomplete.' }
    $profile=Read-SiPrivateJson (Join-Path $safeRoot ("Config/Tenants/$Tenant.local"))
    $mail=$null
    if ($profile.State -eq 'Present') { $mail=Get-SiTenantMailObservation -Root $safeRoot -Tenant $Tenant }
    $certificate=$null
    if ($profile.State -eq 'Present' -and (Get-SiLocalIdentityState -Profile $profile.Value) -eq 'Local format consistent; access not checked') {
        try { $certificate=Get-SiCertificateObservation -Root $safeRoot -Tenant $Tenant -ListCertificates $CertificateSource }
        catch { $certificate=[pscustomobject]@{State='Uncertain';Detail='Local certificate inspection failed. Association, service-account access and permissions remain unverified.'} }
    }
    $task=$null
    try { $task=Get-SiOrchestratorTaskObservation -Tenant $Tenant -SmartM365Root $safeRoot -ListTasks $TaskSource }
    catch { $task=[pscustomobject]@{State='Uncertain';Detail='Local task inspection failed; account, rights, triggers and execution remain unverified.'} }
    $prerequisites=$null
    try { $prerequisites=Get-SiLocalPrerequisiteObservation -Root $safeRoot -ModuleRoots $ModuleRoots }
    catch { $prerequisites=[pscustomobject]@{Observed=@();NotObserved=@();UnscannedLocations=1;Unavailable=$true;Detail='The prerequisite source or local module locations could not be inspected. Module availability and imports remain unverified.'} }
    $catalog=$null
    try {
        $inventoryRoot=Assert-SiLocalAssessmentPath (Join-Path $safeRoot 'SmartInventory') -AllowRepository
        $catalogPath=Assert-SiLocalAssessmentPath (Join-Path $inventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template') -AllowRepository
        if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) { throw 'Source catalog unavailable.' }
        $catalogHash=(Get-FileHash -LiteralPath $catalogPath -Algorithm SHA256).Hash
        $catalog=Get-SiCatalogObservation -CatalogPath $catalogPath -InventoryRoot $inventoryRoot -PathGuard {param($p) Assert-SiLocalAssessmentPath $p -AllowRepository}
        if ((Get-FileHash -LiteralPath $catalogPath -Algorithm SHA256).Hash -cne $catalogHash) { throw 'Source catalog changed during inspection.' }
    } catch { $catalog=$null }
    $after=@(Get-SiAssessmentFingerprint -Root $safeRoot -Tenant $Tenant)
    if (($before | ConvertTo-Json -Compress) -cne ($after | ConvertTo-Json -Compress)) { throw 'Configuration changed during assessment; run it again.' }
    $results=@(foreach ($row in $rows) {
        $state='Manual or external check required'; $detail='No external action or effective-access check was performed.'
        if ($row.Id -in @('discovery','tenant','identity','administrators','teams')) { $state=$row.State; $detail='Local configuration observation only; effective access and remote resources are not verified.' }
        if ($row.Id -eq 'prerequisites') {
            if ($prerequisites.PSObject.Properties['Unavailable'] -and $prerequisites.Unavailable) { $state='Local module scan unavailable'; $detail=$prerequisites.Detail }
            else { $state='Local module names partially inspected'; $detail=$prerequisites.Detail+' Observed: '+$prerequisites.Observed.Count+'; not observed locally: '+$prerequisites.NotObserved.Count+'; skipped path checks: '+$prerequisites.UnscannedLocations+'. Names not observed locally: '+($prerequisites.NotObserved -join ', ')+'.' }
        }
        if ($row.Id -eq 'identity' -and $null -ne $certificate) { $state='Local certificate: '+$certificate.State; $detail=$certificate.Detail }
        if ($row.Id -eq 'sharepoint' -and $profile.State -eq 'Present') {
            $missing=@(foreach ($name in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath')) { if (-not (Test-SiConfiguredValue $profile.Value.$name)) { $name } })
            $state=if ($missing.Count) { 'Local SharePoint fields require review' } else { 'Local SharePoint fields present' }
            $detail='Tenant-level presence only; inheritance, site access and upload are not verified. Fields requiring review: '+($missing -join ', ')+'.'
        }
        if ($row.Id -eq 'orchestrator' -and $null -ne $task) { $state='Local task: '+$task.State; $detail=$task.Detail }
        if ($row.Id -eq 'collectors' -and $null -ne $catalog) { $state=if ($catalog.Issues.Count) { 'Source catalog requires review' } else { 'Source catalog partially checked' }; $detail='Source template only; the effective catalog and access are not verified. Enabled jobs: '+$catalog.EnabledCount+'; issues: '+$catalog.Issues.Count+'; external or unchecked scripts: '+$catalog.UncheckedScripts+'; ignored dependencies: '+$catalog.IgnoredDependencies+'. '+($catalog.Issues -join ' | ') }
        if ($row.Id -eq 'mail' -and $null -ne $mail) { $state=if ($mail.State -eq 'ReviewNeeded') { 'Local mail settings need review' } else { 'Local mail format partially checked' }; $detail=$mail.Detail+' Fields requiring review: '+($mail.Issues -join ', ')+'; unresolved: '+($mail.Unresolved -join ', ')+'.' }
        if ($row.Id -eq 'final') { $state='Not qualified'; $detail='No tenant or network connection, installation or message was attempted; no account or task was created, modified or started.' }
        [pscustomobject]@{Id=$row.Id;State=$state;Detail=$detail}
    })
    [pscustomobject]@{Tenant=$Tenant;Results=$results;OperationallyQualified=$false;Fingerprint=$after}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDTCF7PGOXGvP7a
# GVmQlMSHxvTzot/aXGYeC/ADi1St9KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJnhrWCm4IHRBHE/mvvnI0TK2Q4r/zvjJAgOO18A8leDMA0GCSqG
# SIb3DQEBAQUABIIBgBnkGEAoPdI4vruf6voWM+gxpz8G8XGQOT7YNNUeOIzfbKZ6
# +viUwNXaDyEA03mQOWpzqzjHMChl3hvGd6KtnpE0ikRSo/HfL7AiZed1drIC4PRl
# oQo5mjFVr5Fdh4GXoA5VT7sjXtwsTBDVx/sBtB2KIOfUDJJYVNWL1wiRMl7haMTA
# AAEmQUXklhBQ0pEsq7BokHYB831q6w5syF77uDTrwK+ZHAyReGoMHJdVFuaHiJw5
# p5d9QqEHAjyAsVMI6XlolvC0gHY55bQhTdKzybn8Nlgu8NND6LVBO+QX7quUXVO6
# WicrTc+5dTkET0UbNnaZyae2UZC29SsvmbErxMD2OlEo7+2TBqelAR1fJx+Kj2z1
# tux6lOnfyvLRJIdwY2ZEuelDM1rCG196yHecIPNgzr8Rlq22joPrzPJKgDIt0hIG
# IFUMolbcSPJ0w2p/17KFPIpFklgwPX2tGbzw963jZ3n52ZizHLhIWrGZhL17ToF0
# usSigw7yJ3Muz158c6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjhaMC8GCSqGSIb3DQEJBDEiBCDzu+ZNMmAQ/XGLnxdod2bubYru1C7c4L2Zgv0c
# 9G8f/jANBgkqhkiG9w0BAQEFAASCAgBf4Pxrnxv6UGZyA3OJtrsq28RC92pSCYeU
# qJiOodDgYxnqsZOTVjiTJTTXcQZg0Ba1lxmfblx7ePBALMnvpYV5jwkHvu1IB4Ed
# ktC3chBgkZlxvyEFiI6R9mRUACH/b4d6/jNl7Y7HtjM97ZSduBkH642y5wRZ6Ilx
# xcIOWZKi21oOwH8l/OCfrtdoRVzi3V5qr1RB0XmUVy8Bl8D1kZZua4IydoxBBw+/
# sb/cAanfZNcHupx3e6TSh/Kjah2fYyeuJ1JxWgbVv1plQyOviaGSngsVtgUuibVo
# 3WDklTLYI836MAxw2qU9GPgPuc0j4ZXhQsU9zWNsarNcftq0aD+mc5l3PX3Od5te
# jGlO+ceJfM8/UAdg9QPmH9mY68S0lnYEAljp8APppulR8TioLIgDgxdP3hBgxzSh
# W3yWcnjZnF0TlepNNdWkfF0GVkkfdnJA7il1EYK/Frfi69J/DeSTK2LRzA8Qh8sL
# EvnymkG/HYgNVoEV2h9UB/Z+Psyx+GHxm69/voEbbNLdmmHVyFdsFm3t9sRlJPF2
# 8TyzKmINE4KnIctTDir4OrB97ZYZb2riKayS8ER2kyE/pkuZGtieZTNFZyyjx0Xv
# 8K/WTWm/4IxdaeGDv5KxI1puZqMTZ0kmbjHoIlEUZpJJUQpci2YtSAn/CoQy2aDN
# rKKD80IMww==
# SIG # End signature block
