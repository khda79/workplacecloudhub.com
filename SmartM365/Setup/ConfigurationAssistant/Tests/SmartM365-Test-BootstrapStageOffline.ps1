# Isolated staging fixtures, version 0.4.0. The real setup script is never run.
$ErrorActionPreference='Stop'
$assistantRoot=Split-Path -Parent $PSScriptRoot
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationAudit.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.BootstrapPlan.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.BootstrapStage.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.BootstrapReconcile.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.BootstrapProposal.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.BootstrapRun.Interactive.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.PendingActions.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationEdit.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.PendingEdits.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.OrchestratorPlan.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.OrchestratorRegistration.ps1')
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('si-stage-'+[guid]::NewGuid().ToString('N'))
$root=Join-Path $fixture 'source'; $parent=Join-Path $fixture 'stages'
$null=New-Item -ItemType Directory -Path $root,$parent -Force
try {
    $realRoot=Split-Path -Parent (Split-Path -Parent $assistantRoot)
    $dependencies=@('Config/SmartM365-TenantContext.ps1','Config/SmartM365-JsonTransport.policy.psd1','Config/SmartM365.global.local.json.template','Config/Tenants/tenant.local.json.template','Config/Tenants/tenant.mail.local.json.template','Modules/SmartM365.Core/SmartM365.JsonTransport.psd1','Modules/SmartM365.Core/SmartM365.JsonTransport.psm1')
    foreach ($file in $dependencies) { $dest=Join-Path $root $file; $null=New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force; Copy-Item -LiteralPath (Join-Path $realRoot $file) -Destination $dest }
    $setup=Join-Path $root 'Setup/SmartM365-Create-AppRegistration.ps1'; $null=New-Item -ItemType Directory -Path (Split-Path $setup -Parent) -Force
    [IO.File]::WriteAllText($setup,'throw "fixture setup must never run"')
    $profile=Join-Path $root 'Config/Tenants/test.local.json'; [IO.File]::WriteAllText($profile,'{"ProfileKey":"test","AppId":"00000000-0000-0000-0000-000000000000","TeamsWebhookUrl":"private-secret"}')
    $before=(Get-FileHash $profile).Hash
    $review=New-SiBootstrapPlan -Root $root -ProfileKey test -Directory 'example.test' -DisplayName 'Inventory' -OutsideAddress 'outside@example.test'
    $plan=New-SiBootstrapStagePlan -Root $root -BootstrapPlan $review
    $intent=New-SiPendingBootstrapIntent -Plan $plan -Parent $parent
    if ($intent.Kind -ne 'Bootstrap' -or $intent.State -ne 'Pending' -or @(Get-ChildItem -LiteralPath $parent).Count -ne 0 -or (Get-FileHash $profile).Hash -ne $before) {throw 'Preparing the bootstrap intent had side effects.'}
    $stage=New-SiBootstrapStage -Plan $plan -Parent $parent
    if ($stage.State -ne 'PreparedOnly' -or (Get-FileHash $profile).Hash -ne $before -or (Get-Content $stage.ProfilePath -Raw) -match 'private-secret') { throw 'Original profile was changed or copied.' }
    . (Join-Path $stage.Path 'Config/SmartM365-TenantContext.ps1')
    $context=Get-SmartM365EffectiveGlobalConfig -StartPath (Join-Path $stage.Path 'Setup') -ProfileKey test
    if ($context.SmartM365RootPath -ine $stage.Path -or $context.TenantKey -ne 'isolated-bootstrap') { throw 'The staging context escaped its root.' }
    $seen=@{}
    $fakeLaunch={param($executable,$encoded) $seen.Command=[Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($encoded)); [pscustomobject]@{HasExited=$true;ExitCode=0}}.GetNewClosure()
    $run=Start-SiBootstrapRun -Stage $stage -Launch $fakeLaunch
    if ($seen.Command -match 'UseDeviceCode|UpdateExisting|RemoveAppRegistration') { throw 'Unexpected bootstrap arguments.' }
    $synthetic=Get-Content -LiteralPath $stage.ProfilePath -Raw | ConvertFrom-Json
    $synthetic.AppId='11111111-1111-1111-1111-111111111111'; $synthetic.TenantId='22222222-2222-2222-2222-222222222222'; $synthetic.Thumb='A'*40; $synthetic.Thumbprint='A'*40; $synthetic.From='sender@example.test'; $synthetic.MailSendAccessPolicyGroup='scope@example.test'
    [IO.File]::WriteAllText($stage.ProfilePath,($synthetic | ConvertTo-Json -Depth 100))
    $completed=Complete-SiBootstrapRun -Run $run
    if ($completed.State -ne 'ReviewRequired' -or -not (Test-Path -LiteralPath $completed.ProposalPath) -or (Get-FileHash $profile).Hash -ne $before) { throw 'Synthetic bootstrap result was not isolated.' }
    $proposal=Get-Content -LiteralPath $completed.ProposalPath -Raw | ConvertFrom-Json
    if ($proposal.PSObject.Properties['TeamsWebhookUrl'] -or $proposal.PSObject.Properties['SharePointSitePath']) { throw 'Unrequested values leaked into proposal.' }
    $rejected=$false; try { Start-SiBootstrapRun -Stage $stage -Launch $fakeLaunch | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'A changed staging profile was accepted.' }
    [IO.File]::AppendAllText((Join-Path $root $dependencies[0]),'# changed')
    $rejected=$false; try {New-SiPendingBootstrapIntent -Plan $plan -Parent $parent | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'A bootstrap intent accepted a changed dependency.'}
    $rejected=$false; try { New-SiBootstrapStage -Plan $plan -Parent $parent | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'A changed dependency was accepted.' }
    [IO.File]::WriteAllText($profile,'{"ProfileKey":"test","AppId":"11111111-1111-1111-1111-111111111111"}')
    $review=New-SiBootstrapPlan -Root $root -ProfileKey test -Directory 'example.test' -DisplayName 'Inventory' -OutsideAddress 'outside@example.test'
    $rejected=$false; try { New-SiBootstrapStagePlan -Root $root -BootstrapPlan $review | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'An existing application was accepted for unpinned update.' }
    $newProfile=Get-SiTenantCreationPlan -Root $root -ProfileKey client_eu-prod -OrganizationKey example -EnvironmentKey test
    $newAction=New-SiPendingCreation -Kind Profile -Data $newProfile
    $virtual=Get-SiPendingCreationSnapshot -Action $newAction -Tenant client_eu-prod -Step tenant -Field OrgDomain
    Set-SiPendingCreationSetting -Snapshot $virtual -Value 'example.invalid'
    $newReview=New-SiBootstrapPlan -Root $root -ProfileKey client_eu-prod -Directory example.test -DisplayName Inventory -AutoSelectOutsideAddress $true -ConfigureTeams $true -PendingProfile $newProfile
    $newPlan=New-SiBootstrapStagePlan -Root $root -BootstrapPlan $newReview
    $null=New-SiPendingBootstrapIntent -Plan $newPlan -Parent $parent
    if (Test-Path -LiteralPath $newProfile.Destination) {throw 'Planning a new profile wrote its configuration.'}
    . (Join-Path $root 'Config/SmartM365-TenantContext.ps1')
    $taskDir=Join-Path $root 'SmartInventory/Orchestrator';$null=New-Item -ItemType Directory -Path $taskDir -Force
    [IO.File]::WriteAllText((Join-Path $taskDir 'Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'),'# synthetic installer; never executed')
    [IO.File]::WriteAllText((Join-Path $taskDir 'SmartM365-Inventory-Orchestrator.ps1'),'# synthetic orchestrator; never executed')
    $exe=Join-Path $fixture 'pwsh.exe';[IO.File]::WriteAllText($exe,'synthetic')
    $actions=[Collections.Generic.List[object]]::new();Add-SiPendingCreation -Queue $actions -Action $newAction
    $taskIntent=New-SiPendingTaskIntent -Root $root -Tenant client_eu-prod -Executable $exe -AccountLabel 'TEST\synthetic' -EnableAfterCreation $true
    Add-SiPendingCreation -Queue $actions -Action $taskIntent
    $created=Invoke-SiPendingCreations -Queue $actions
    if ($created.Completed.Count -ne 1 -or $actions.Count -ne 1 -or -not (Test-Path -LiteralPath $newProfile.Destination)) {throw 'Synthetic reviewed creation sequence failed.'}
    $resolvedPlan=New-SiBootstrapStagePlan -Root $root -BootstrapPlan $newReview
    if ($resolvedPlan.ProfileHash -eq 'ABSENT') {throw 'New profile bootstrap was not resolved after creation.'}
    $newStage=New-SiBootstrapStage -Plan $newPlan -Parent $parent
    $newRun=Start-SiBootstrapRun -Stage $newStage -Launch $fakeLaunch
    $staged=Get-Content -LiteralPath $newStage.ProfilePath -Raw | ConvertFrom-Json
    $proposalValues=@{AppId='11111111-1111-1111-1111-111111111111';TenantId='22222222-2222-2222-2222-222222222222';Thumb=('A'*40);Thumbprint=('A'*40);From='sender@example.test';SmtpServer='';MailSendAccessPolicyGroup='scope@example.test';SharePointSiteHostname='tenant.example.test';SharePointSitePath='/sites/synthetic';SharePointLibraryDisplayName='Documents';SharePointTargetFolderPath='DATA'}
    foreach ($key in $proposalValues.Keys) {$staged.$key=$proposalValues[$key]}
    [IO.File]::WriteAllText($newStage.ProfilePath,($staged | ConvertTo-Json -Depth 100))
    $completion=Complete-SiBootstrapRun -Run $newRun
    if ($completion.State -ne 'ReviewRequired') {throw 'Synthetic bootstrap did not produce a reviewed proposal.'}
    $newProposal=$completion.ProposalPath
    $reconciliation=New-SiBootstrapReconciliation -Root $root -Tenant client_eu-prod -ProposalPath $newProposal -NewProfilePlan $newReview
    if ($reconciliation.Replacements.Count -lt 4 -or $reconciliation.BlockedGroups.Count -ne 0) {throw 'Unchanged new-profile defaults were not prepared for review.'}
    $queue=[Collections.Generic.List[object]]::new();Add-SiPendingEdit -Queue $queue -Edit (New-SiPendingReconciliationEdit -Plan $reconciliation)
    $result=Invoke-SiPendingEdits -Queue $queue
    $newSaved=Get-Content -LiteralPath $newProfile.Destination -Raw | ConvertFrom-Json
    if ($result.Applied -ne 1 -or $newSaved.AppId -ne $proposalValues.AppId -or $newSaved.OrgDomain -ne 'example.invalid') {throw 'Reviewed bootstrap values were not applied or the explicit profile value changed.'}
    $sid='S-1-5-21-1-2-3-1001';$identity=[pscustomobject]@{Account='TEST\synthetic';Sid=$sid;Managed=$false}
    $fakeStart={param($p) [pscustomobject]@{Plan=$p;Process=[pscustomobject]@{HasExited=$true;ExitCode=0};Lock=[IO.MemoryStream]::new()}}
    $taskRun=Start-SiPendingTaskAction -Action $taskIntent -Identity $identity -ListTasks {@()} -Start $fakeStart -AllowedExecutables @($exe)
    $task=[pscustomobject]@{TaskName=$taskRun.Plan.TaskName;TaskPath=$taskRun.Plan.TaskPath;Settings=[pscustomobject]@{Enabled=$false};Actions=@([pscustomobject]@{Execute=$taskRun.Plan.Execute;Arguments=$taskRun.Plan.Arguments});Principal=[pscustomobject]@{UserId=$identity.Account};Triggers=@([pscustomobject]@{CimClass=[pscustomobject]@{CimClassName='MSFT_TaskBootTrigger'};Enabled=$true;Delay='PT5M'},[pscustomobject]@{CimClass=[pscustomobject]@{CimClassName='MSFT_TaskDailyTrigger'};Enabled=$true;DaysInterval=1;StartBoundary='2026-01-01T00:00:00';Repetition=[pscustomobject]@{Interval='PT5M';Duration='P1D';StopAtDurationEnd=$false}})}
    $list={$task}.GetNewClosure();$resolved={param($name) [pscustomobject]@{Sid=$sid}}.GetNewClosure();$enable={param($name,$path) $task.Settings.Enabled=$true}.GetNewClosure()
    $registered=Complete-SiOrchestratorRegistration -Run $taskRun -ListTasks $list
    if ($registered.State -ne 'CreatedDisabled') {throw 'Synthetic task creation did not complete.'}
    $enabled=Enable-SiCreatedOrchestratorTask -Plan $taskRun.Plan -ListTasks $list -ResolveAccount $resolved -Enable $enable
    if ($enabled.State -ne 'CreatedEnabled' -or -not $task.Settings.Enabled) {throw 'Synthetic task activation did not complete.'}
    [IO.File]::AppendAllText($newProfile.Destination,' ');$rejected=$false
    try {New-SiBootstrapStagePlan -Root $root -BootstrapPlan $newReview | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'Changed new profile was accepted.'}
    $protected=Get-SiTenantCreationPlan -Root $root -ProfileKey protected -OrganizationKey example -EnvironmentKey test
    $protectedAction=New-SiPendingCreation -Kind Profile -Data $protected
    $explicit=Get-SiPendingCreationSnapshot -Action $protectedAction -Tenant protected -Step mail -Field From
    Set-SiPendingCreationSetting -Snapshot $explicit -Value 'chosen@example.test'
    $protectedReview=New-SiBootstrapPlan -Root $root -ProfileKey protected -Directory example.test -DisplayName Inventory -AutoSelectOutsideAddress $true -PendingProfile $protected
    $null=New-SiTenantProfile -Root $root -ProfileKey protected -OrganizationKey example -EnvironmentKey test -ExpectedTemplateHash $protected.TemplateHash -Edits @($protected.PendingEdits)
    $mailProposal=Join-Path $fixture 'protected-proposal.json'
    [IO.File]::WriteAllText($mailProposal,(@{ProfileKey='protected';From='other@example.test';SmtpServer='';MailSendAccessPolicyGroup='scope@example.test'} | ConvertTo-Json))
    $protectedPlan=New-SiBootstrapReconciliation -Root $root -Tenant protected -ProposalPath $mailProposal -NewProfilePlan $protectedReview
    if ($protectedPlan.BlockedGroups -notcontains 'Mail' -or @($protectedPlan.Replacements | Where-Object Field -eq 'From').Count) {throw 'An explicitly selected sender was treated as a replaceable default.'}
    [IO.File]::AppendAllText($protected.Destination,' ');$rejected=$false
    try {New-SiBootstrapReconciliation -Root $root -Tenant protected -ProposalPath $mailProposal -NewProfilePlan $protectedReview | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'Changed new-profile provenance was accepted.'}
    'Bootstrap staging fixtures passed.'
} finally {
    $resolved=(Resolve-Path -LiteralPath $fixture).Path
    if ($resolved -ne $fixture -or [IO.Path]::GetFileName($resolved) -notmatch '^si-stage-[0-9a-f]{32}$') { throw 'Unexpected fixture path; cleanup refused.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFAoJGz1BzdvVY
# TfwAq/14pIXuQ7EkJujs0E49rkV01aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC9NPpbJ5ALWzRS46BDdWIRxkwggXg7leF27vPndZZKrMA0GCSqG
# SIb3DQEBAQUABIIBgDda3SthQy16AuwQjoWUx7cn/1P3q7xtq2WoNkkey3dM95Pc
# oWogEyoYB+yraWj1Lq7sly5aZFwDaNpiienFH/9D4yiTcmOEY1pxykNoGNGRqheD
# MlJBs/cggRE8VXwXBHgTxrxBmAyrLJ2eegFSO3unFgLDT5Pu5+p21TUmX/pwYg3c
# d3k272cHeABffist7IqYDOXFA8OOihB5NFCBEYlkQGkYljMkpew5z8wwMlS2lQs8
# T3hYF0yhJXRpmFSzYu0seVRteTPBc2qDORm+Z7+4RrQmWw848TcCd1UlPIn2iedg
# z8Kxk1nzi+MZymKvTTvQeOhIVit7zTt28V5j54N3M/Dzif1ds+2lWPgtlLIRnPt1
# 59/PYjHUJB1cmnm5hGuII+t8EuN9RDOh1Y67t8LcplBKpKOXB0JHAD04pwpsimP0
# VnwxzRqVs+prqVgqM8IoPwX5V1iB5I9/rNhCnhUmQq/zoGlstk/DIiam2QhqjWb0
# 2sOnbzjLLJm1IXDMY6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MDJaMC8GCSqGSIb3DQEJBDEiBCDHoqLjarxVkJ2CUFWd97x5ne8IdBomGdHdUssu
# JpJ3YzANBgkqhkiG9w0BAQEFAASCAgAZDxuM6xAfVkyD+mjTm3H4n7qdB0zUn18k
# PpSnXlUNxPRS9Hurwwp+4736f6Z9C86rJTA2nTnAQYI9//VjD0YWwVPaQALkfII5
# JHM+YoakMmsxvCREmXPZamNSVJSJHS0WxObJfOkkXOl3bmiOcu586nfTJDXggD7z
# 8o1W6uRlTEPFrn2qhkVXU7NwWfuzTtjzaAfr+WuFjwuRus9KAolPmqyWmXneHPVc
# 4HQf5mJVHEi6qs2o98m4pKaK87Hc/4gQv0SyoKS4SOEt+To0Bj4MCCnSy8MfjiWO
# a/I7TMplM3o2gV50qV41vbvrNH8I644OzgRlrfW0lLUgDRHBkke4VONZYgkxHNS1
# gJH+gGgcqJFrhf/0nbPp74adNVTjC6/C6yQeCglaQbLKDJ05VOD8AIOiAm4WBSBL
# jD3xJDkI/I1YcJhwxpR3ECCYdampbE1XbsqWsLMBBCxOHdSirhDtPkHaW9wGZdWB
# gRl5cCPW9MQqtnb6LQSkZRNIpum6MK889betjuYckPgfqcYdcwg2TAF7CZk04TMF
# hrRXEaE5Ip+bsKhzcMrlFnzBrEumTx5L0avZiBX0iaeHnQ1Iudww4uJePAkD/6Rv
# WKGOxjSGAiLdVxNvTpaiS609zAQrYAHLSbYLWicGBx4eiXOkC5AewDZgRegCgymQ
# K6ThaOygFw==
# SIG # End signature block
