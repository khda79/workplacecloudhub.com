# Passive assessment against synthetic files only; no real services are touched, version 0.5.0.
$ErrorActionPreference='Stop'
$assistantRoot=Split-Path -Parent $PSScriptRoot
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationAudit.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationSteps.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.MailConfigurationCheck.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationChecks.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationCatalog.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.WorkflowAssessment.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.LocalPrerequisiteScan.ps1')
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('si-assessment-'+[guid]::NewGuid().ToString('N'))
try {
    $folder=Join-Path $fixture 'Config/Tenants'; $null=New-Item -ItemType Directory -Path $folder -Force
    $missing=Join-Path $folder 'missing/child.json'
    if ((Assert-SiLocalAssessmentPath $missing) -ine [IO.Path]::GetFullPath($missing)) {throw 'A missing synthetic descendant was not handled.'}
    $junction=Join-Path $fixture 'synthetic-junction'
    $syntheticAttributes={param($candidate) if ($candidate -ieq $junction) {[IO.FileAttributes]::Directory -bor [IO.FileAttributes]::ReparsePoint} else {[IO.File]::GetAttributes($candidate)}}.GetNewClosure()
    $rejected=$false;try {Assert-SiLocalAssessmentPath (Join-Path $junction 'Tenants') -GetAttributes $syntheticAttributes | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'A synthetic reparse point was accepted.'}
    $denied={param($candidate) throw [UnauthorizedAccessException]::new('synthetic denial')}
    $rejected=$false;try {Assert-SiLocalAssessmentPath $folder -GetAttributes $denied | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'An attribute access failure was treated as a missing path.'}
    $uncRoot='\\synthetic.invalid\share\SmartM365';$uncFile=$uncRoot+'\Config\Tenants\test.local.json'
    Set-SiAssistantRepositoryRoot -Root $uncRoot
    $virtualAttributes={param($candidate) [IO.FileAttributes]::Directory}
    if ((Assert-SiLocalAssessmentPath $uncRoot -AllowRepository -GetAttributes $virtualAttributes) -ine $uncRoot) {throw 'The synthetic UNC repository root was rejected.'}
    if ((Assert-SiLocalAssessmentPath $uncFile -AllowRepository -GetAttributes $virtualAttributes) -ine $uncFile) {throw 'The synthetic UNC repository path was rejected.'}
    foreach ($unsafe in @('\\synthetic.invalid\share\SmartM365-other\file.json','\\synthetic.invalid\other\SmartM365\file.json',($uncRoot+'\..\other\file.json'))) {
        $rejected=$false;try {Assert-SiLocalAssessmentPath $unsafe -AllowRepository -GetAttributes $virtualAttributes | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'A path outside the synthetic UNC repository was accepted.'}
    }
    $rejected=$false;try {Assert-SiLocalAssessmentPath $uncFile -GetAttributes $virtualAttributes | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'A local-only check accepted a UNC repository path.'}
    $linked={param($candidate) if ($candidate -ieq ($uncRoot+'\Config')) {[IO.FileAttributes]::Directory -bor [IO.FileAttributes]::ReparsePoint} else {[IO.FileAttributes]::Directory}}.GetNewClosure()
    $rejected=$false;try {Assert-SiLocalAssessmentPath $uncFile -AllowRepository -GetAttributes $linked | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'A synthetic UNC reparse point was accepted.'}
    Set-SiAssistantRepositoryRoot -Root $fixture
    $global=Join-Path $fixture 'Config/SmartM365.global.local.json'; $profile=Join-Path $folder 'test.local.json'
    [IO.File]::WriteAllText($global,'{}')
    $setup=Join-Path $fixture 'Setup'; $moduleRoot=Join-Path $fixture 'modules'; $inventory=Join-Path $fixture 'SmartInventory'
    $null=New-Item -ItemType Directory -Path $setup,(Join-Path $moduleRoot 'Example.One'),(Join-Path $inventory 'Orchestrator'),(Join-Path $inventory 'jobs') -Force
    [IO.File]::WriteAllText((Join-Path $setup 'Install-SmartM365-SmartInventoryPrerequisites.ps1'),"`$requiredModules += @('Example.One','Example.Two'); throw 'must not execute'")
    [IO.File]::WriteAllText((Join-Path $inventory 'jobs/collector.ps1'),'throw "must not execute"')
    [IO.File]::WriteAllText((Join-Path $inventory 'Orchestrator/Orchestrator-Jobs.json.template'),'{"Jobs":[{"Name":"Example","Enabled":true,"ScriptPath":"jobs/collector.ps1","DependsOn":[],"AssignmentMode":"Elected"}]}')
    [IO.File]::WriteAllText($profile,'{"ProfileKey":"test","OrganizationKey":"sample","EnvironmentKey":"test","TenantKey":"sample-test","From":"not-an-address","TeamsWebhookUrl":"synthetic-private-value"}')
    $before=(Get-FileHash $profile).Hash
    $emptyTasks={@()}
    $report=Get-SiLocalWorkflowAssessment -Root $fixture -Tenant test -TaskSource $emptyTasks -ModuleRoots @($moduleRoot)
    if ($report.Results.Count -ne 15 -or $report.OperationallyQualified -or ($report.Results | ConvertTo-Json -Depth 10) -match 'synthetic-private-value') { throw 'Assessment leaked data or overstated qualification.' }
    if ((@($report.Results | Where-Object Id -eq 'mail')[0]).State -ne 'Local mail settings need review') { throw 'Invalid synthetic mail was not reported.' }
    if ((@($report.Results | Where-Object Id -eq 'prerequisites')[0]).State -ne 'Local module names partially inspected' -or (@($report.Results | Where-Object Id -eq 'collectors')[0]).State -ne 'Source catalog partially checked') { throw 'Synthetic prerequisite or catalog check was not reported.' }
    if ((Get-FileHash $profile).Hash -ne $before) { throw 'Assessment modified the profile.' }
    $worker=[PowerShell]::Create()
    try {
        [void]$worker.AddCommand((Join-Path $assistantRoot 'Lib/SmartM365.LocalAssessmentWorker.ps1')).AddParameter('Root',$fixture).AddParameter('Tenant','test').AddParameter('LibraryRoot',(Join-Path $assistantRoot 'Lib')).AddParameter('OfflineFixture',$true)
        $async=$worker.BeginInvoke()
        if (-not $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds(20))) {throw 'Synthetic background assessment timed out.'}
        $background=@($worker.EndInvoke($async))
        if ($worker.HadErrors -or $background.Count -ne 1 -or $background[0].Results.Count -ne 15 -or $background[0].OperationallyQualified) {throw 'Synthetic background assessment failed.'}
        foreach ($id in @('mail','collectors','discovery')) {if (@($background[0].Results | Where-Object Id -eq $id)[0].State -cne @($report.Results | Where-Object Id -eq $id)[0].State) {throw ('Background assessment changed the '+$id+' result.')}}
        if ((Get-FileHash $profile).Hash -ne $before) {throw 'Background assessment modified the profile.'}
    } finally {$worker.Dispose()}
    $synthetic=@{ProfileKey='test';OrganizationKey='sample';EnvironmentKey='test';TenantKey='sample-test';AppId='11111111-1111-1111-1111-111111111111';TenantId='22222222-2222-2222-2222-222222222222';Thumb=('A'*40);Thumbprint=('A'*40);EnableTeamsNotifications=$true;TeamsAlertsWebhookUrl='https://workflow.example.test/private';TeamsInfosWebhookUrl='https://workflow.example.test/other';SharePointSiteHostname='tenant.example.test';SharePointSitePath='/sites/test';SharePointLibraryDisplayName='Documents';SharePointTargetFolderPath='data'}
    [IO.File]::WriteAllText($profile,($synthetic | ConvertTo-Json))
    $certificates={ [pscustomobject]@{Thumbprint=('A'*40);NotBefore=(Get-Date).AddDays(-1);NotAfter=(Get-Date).AddDays(1);HasPrivateKey=$true} }
    $report=Get-SiLocalWorkflowAssessment -Root $fixture -Tenant test -CertificateSource $certificates -TaskSource $emptyTasks -ModuleRoots @($moduleRoot)
    if ((@($report.Results | Where-Object Id -eq 'identity')[0]).State -ne 'Local certificate: Observed' -or $report.OperationallyQualified -or ($report.Results | ConvertTo-Json -Depth 10) -match 'workflow.example.test') { throw 'Synthetic local checks were misreported or revealed a URL.' }
    $fakeTasks={ [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - test';TaskPath='\WCH\';Principal=[pscustomobject]@{UserId='SYNTHETIC\service'};Actions=@()} }
    $report=Get-SiLocalWorkflowAssessment -Root $fixture -Tenant test -CertificateSource $certificates -TaskSource $fakeTasks -ModuleRoots @($moduleRoot)
    if ((@($report.Results | Where-Object Id -eq 'orchestrator')[0]).State -ne 'Local task: ActionMismatch' -or $report.OperationallyQualified) { throw 'Synthetic task mismatch was not reported.' }
    $customTasks={ [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - client_eu-prod';TaskPath='\WCH\';Principal=[pscustomobject]@{UserId='SYNTHETIC\service'};Actions=@()} }
    $customReport=Get-SiLocalWorkflowAssessment -Root $fixture -Tenant client_eu-prod -CertificateSource $certificates -TaskSource $customTasks -ModuleRoots @($moduleRoot)
    if ((@($customReport.Results | Where-Object Id -eq 'orchestrator')[0]).State -ne 'Local task: ActionMismatch') {throw 'Local workflow skipped custom-profile task observation.'}
    $originalMail=(Get-Command Get-SiTenantMailObservation).ScriptBlock
    try {
        function Get-SiTenantMailObservation { param($Root,$Tenant) [IO.File]::AppendAllText((Join-Path $Root ("Config/Tenants/$Tenant.local.json")),' '); [pscustomobject]@{State='Partial';Issues=@();Unresolved=@();Detail='synthetic'} }
        $rejected=$false; try { Get-SiLocalWorkflowAssessment -Root $fixture -Tenant test -CertificateSource $certificates -TaskSource $emptyTasks -ModuleRoots @($moduleRoot) | Out-Null } catch { $rejected=$true }
        if (-not $rejected) { throw 'Concurrent profile change was accepted.' }
    } finally { Set-Item -Path Function:Get-SiTenantMailObservation -Value $originalMail }
    [IO.File]::WriteAllText(($profile+'.txt'),'{}')
    $report=Get-SiLocalWorkflowAssessment -Root $fixture -Tenant test -TaskSource $emptyTasks -ModuleRoots @($moduleRoot)
    if ((@($report.Results | Where-Object Id -eq 'discovery')[0]).State -ne 'Correction required') { throw 'Conflicting profiles were not reported.' }
    $rejected=$false; try { Get-SiLocalWorkflowAssessment -Root '\\server.invalid\share' -Tenant test | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'A network root was accepted.' }
    'Workflow local assessment fixtures passed.'
} finally {
    $resolved=(Resolve-Path -LiteralPath $fixture).Path
    if ($resolved -ne $fixture -or [IO.Path]::GetFileName($resolved) -notmatch '^si-assessment-[0-9a-f]{32}$') { throw 'Unexpected fixture path; cleanup refused.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDevRQzg0w3965H
# uocpks78jZ5ebG/T9XjzPoSZ3V9Df6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAYa4f2hPmNFYfsGNJSnXjgJb/M4lcE9xsDHR/2vx/0uMA0GCSqG
# SIb3DQEBAQUABIIBgE72VU0SysebwfQ8j/6+TnGIW5gFKV4pBu9wNPVRC1Sbc6W7
# Ze40p9oBQ/CsZ6s+4/dpSZNkxDPvbabEvKbvonGTj2UbvNPwBoWz+Ut0rVcC+WEC
# kBJVdeIJLb4p8S/CdnZ//l4yk6Sd+wHukZAIn2o5W1JCZaj6hCR2d4ZnXP0rj/M5
# PedwlZF5HOpqU8gUW6qGg/dYYqs0evx6ntMi1bGTR3drOOYKP3InoFew40rt51UE
# Oog4eqHuza8E+dhS4u6FrXlPwKN07al8xWJfEpx4ylDjQvR5dlQxicnGHYuhKLtX
# 3+9e3+0LCrZxE/g4rFM3hvNO4+9/LCrGIgZ7NCLugU2jUr2hqFT9lIl+VFZ/F+K/
# /0lNc0MQlvtgIqgwbrlMzTEYSFLVjzKIYSiemGGlHHbBMCV7V0+BfIQzXrQ4o77i
# MLjUCxiuz712lZPG8Iw9KxoHIZH5+qPi3IGqqAubyeRqVHUDl4n/rHSde81PxCg0
# 0mNmuk41t0Z1+IBjS6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTJaMC8GCSqGSIb3DQEJBDEiBCB2SYEHMYMsBZ51RfWO7CyKId4AHaRa0FkV+6RI
# PV/UuDANBgkqhkiG9w0BAQEFAASCAgAEfLuMOR0t9mdTvt7HtQg7UZwIbbVsNcFf
# wP8ik61JZ39v29GyVvVeNOJnXNKUyrxA3nx5p9inBSc5QMwy24HsHaV/9c81THMd
# g/F3fixHrDg0YV8JTpvsQwvrxuTq79DbjEEvEtCbs3E43jRIhFv+RoawF/F+EiNL
# SCrxSvN8RAebZxJvN13g9eP5dYJncTXZqtDe9pdvlVsyPtFhD4VXLwfL2MMWrswu
# FmHKX/YVvwygCXpU1CeDrr80oiTkPhj9GgJJMDc5AZr0AEsHCBbD06MZuUc0tHLS
# y5Q/6BM7sJ3Q/nvKrxhiPHkurobSPb2jCWm8UXQLSOBhMutTp3c/wTE/EDUIxd8R
# lTVTG5w2keVk3Ts0IrotQCtAHW1oliQmVctnmAgUzXUa3g85xHn6cwRB62F9QEKt
# UxfBnjRzeBY0V3/Ss36H0YdCOMxfbCi6yeIqg8SytTxqrJbBVXqXaiz2TdoqPs3x
# 0GnaGtogJBnvEUMU2q/AeFc04KfVla9gqC/WIYWv0oUZ5FFjzngC/4Q5EwT9nCd0
# 7kbgJ+b5gQt0nqV+jPqyxgE64Usjy1klzZAWMcswHxYqyGb6p+OknV1/haLTul0x
# qh4krWrPtQFojmEyC/r4wnEdHDt/kTZ+oaP7Ndew2q9cY1JhdpwTnPToTpKka6cH
# 1lGLjV0w+Q==
# SIG # End signature block
