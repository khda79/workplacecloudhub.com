# Inspect and exercise isolated installer expressions with synthetic adapters only, version 0.2.0.
$ErrorActionPreference='Stop'
$assistant=Split-Path -Parent $PSScriptRoot;$root=Split-Path -Parent (Split-Path -Parent $assistant)
$installer=Join-Path $root 'SmartInventory/Orchestrator/Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($installer,[ref]$tokens,[ref]$errors)
if ($errors.Count) {throw 'Task installer did not parse.'}
$source=$ast.Extent.Text
if ($source.IndexOf('Invoke-SiServiceCertificateImport -PfxPath') -lt 0 -or $source.IndexOf('Invoke-SiServiceCertificateImport -PfxPath') -gt $source.IndexOf('Register-ScheduledTask @registrationParameters')) {throw 'Certificate import must precede task registration.'}
if ($source.IndexOf('Ensure-SiBatchLogonRight -Credential') -lt 0 -or $source.IndexOf('Ensure-SiBatchLogonRight -Credential') -gt $source.IndexOf('Invoke-SiServiceCertificateImport -PfxPath')) {throw 'Batch logon must be verified before certificate import.'}
if ($source.IndexOf('Ensure-SiInteractiveLogonRight -Credential') -lt 0 -or $source.IndexOf('Ensure-SiInteractiveLogonRight -Credential') -gt $source.IndexOf('Invoke-SiServiceCertificateImport -PfxPath')) {throw 'Interactive logon must be verified before certificate import.'}
if ($source -match 'Add-LocalGroupMember|Add-ADGroupMember|net\s+localgroup') {throw 'The installer must not add the account to an administrator group.'}
if ($source -notmatch '\$credentialSid -ne \$serviceAccountSid' -or $source -notmatch 'PfxPassword \$pfxPassword') {throw 'Service account or PFX import guard is missing.'}
foreach ($pattern in @('CreateOnlyDisabled -and \(\$StartNow -or \$Uninstall -or \$Interactive\)','\$settingsParameters\.Disable\s*=\s*\$true','Force\s*=\s*\(-not \$CreateOnlyDisabled\)','if \(-not \$CreateOnlyDisabled -and \$null -ne \$legacyTask','if \(-not \$CreateOnlyDisabled\) \{\s*foreach \(\$guiShortcutPath in \$guiShortcutPaths\)','\$serviceAccountSid -ne \$ExpectedServiceSid')) {
    if ($source -notmatch $pattern) {throw 'A safe-mode guard is missing.'}
}
$settingGuard=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(-not \$CreateOnlyDisabled -and \(-not \$script:InteractiveMode -or \$interactiveTaskEnabled\)\)'},$true)
if (-not $settingGuard -or $source -notmatch 'Enable the scheduled task after installation\?' -or $source -match 'Start the task immediately after installation\?') {throw 'Interactive activation choice was not implemented.'}
if ($source.IndexOf('Register-SmartM365OrchestratorServer -SharedDataFolderPath') -gt $source.IndexOf('Enable-ScheduledTask -TaskPath')) {throw 'Task enabled before server registration.'}
function Enable-ScheduledTask { $script:enabledCalled=$true }
function Write-InstallerMessage {param($Message)}
foreach ($case in @(@($true,$false,$false,$true),@($true,$true,$false,$false),@($false,$true,$true,$true),@($false,$true,$false,$false))) {
    $script:InteractiveMode=$case[0];$interactiveTaskEnabled=$case[1];$CreateOnlyDisabled=$case[2];$script:enabledCalled=$false;$taskPath='synthetic';$TaskName='synthetic'
    . ([scriptblock]::Create($settingGuard.Extent.Text))
    if ($script:enabledCalled -eq $case[3]) {throw 'Task activation choice did not map to the correct enabled state.'}
}
$assignment=$ast.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$registrationParameters'},$true)
if ($null -eq $assignment) {throw 'Registration parameters not found.'}
$taskPath='\WCH\';$TaskName='synthetic';$action='fake';$startupTrigger='boot';$dailyTrigger='daily';$settings='disabled';$description='synthetic';$ServiceAccount='TEST\svc';$plainTextPassword='synthetic'
$CreateOnlyDisabled=$true;. ([scriptblock]::Create($assignment.Extent.Text))
if ($registrationParameters.Force -ne $false) {throw 'Safe mode would request replacement.'}
$CreateOnlyDisabled=$false;. ([scriptblock]::Create($assignment.Extent.Text))
if ($registrationParameters.Force -ne $true) {throw 'Legacy default changed.'}
$output=@(& pwsh -NoProfile -File $installer -Tenant 'client_eu-prod' -CreateOnlyDisabled -StartNow -ExpectedServiceSid 'S-1-5-21-1-2-3-1001' 2>&1 | ForEach-Object {[string]$_})
if ($LASTEXITCODE -eq 0 -or ($output -join ' ') -notmatch 'cannot be combined') {throw 'Contradictory safe-mode options were not rejected before execution.'}
foreach ($invalid in @('../other','bad name','a;whoami',('x'*65))) {
    $output=@(& pwsh -NoProfile -File $installer -Tenant $invalid -CreateOnlyDisabled -StartNow -ExpectedServiceSid 'S-1-5-21-1-2-3-1001' 2>&1 | ForEach-Object {[string]$_})
    if ($LASTEXITCODE -eq 0 -or ($output -join ' ') -match 'cannot be combined') {throw 'The installer accepted an invalid profile before its safety guard.'}
}
foreach ($functionName in @('Get-InstallerHostLaunchMode','Get-InstallerAvailableProfiles','Read-InstallerProfile','Read-InstallerProfileDocument','Read-InstallerServiceAccount')) {
    $functionAst=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName},$true)
    if (-not $functionAst) {throw "Missing function $functionName"}
    . ([scriptblock]::Create($functionAst.Extent.Text))
}
foreach ($case in @(@($false,5,'Elevate'),@($false,7,'Elevate'),@($true,5,'PowerShell7'),@($true,7,'Current'))) {
    if ((Get-InstallerHostLaunchMode -IsAdministrator $case[0] -PowerShellMajor $case[1]) -cne $case[2]) {throw 'Incorrect PowerShell host or elevation decision.'}
}
if ($source -notmatch "(?s)\`$hostLaunchMode -ne 'Current'.*Get-PowerShell7Path.*Start-Process @launchOptions" -or $source -match 'using Windows PowerShell for UAC elevation') {throw 'The installer may continue in Windows PowerShell instead of PowerShell 7.'}
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('si-profile-'+[guid]::NewGuid().ToString('N'))
try {
    $folder=Join-Path $fixture 'Config/Tenants';$null=New-Item -ItemType Directory -Path $folder -Force
    [IO.File]::WriteAllText((Join-Path $folder 'client_eu-prod.local.json.txt'),'{"ProfileKey":"client_eu-prod"}')
    [IO.File]::WriteAllText((Join-Path $folder 'test.local.json'),'{"ProfileKey":"wrong"}')
    [IO.File]::WriteAllText((Join-Path $folder 'duplicate.local.json'),'{"ProfileKey":"duplicate"}')
    [IO.File]::WriteAllText((Join-Path $folder 'duplicate.local.json.txt'),'{"ProfileKey":"duplicate"}')
    [IO.File]::WriteAllText((Join-Path $folder 'invalid.local.json'),'{invalid')
    [IO.File]::WriteAllText((Join-Path $folder 'client_eu-prod.mail.local.json'),'{"ProfileKey":"client_eu-prod"}')
    Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force -ErrorAction Stop
    $before=@(Get-ChildItem -LiteralPath $folder -File | Sort-Object Name | ForEach-Object {$_.Name+':'+(Get-FileHash -LiteralPath $_.FullName).Hash}) -join '|'
    $readProfile=Read-InstallerProfileDocument -Root $fixture -ProfileKey 'client_eu-prod'
    $after=@(Get-ChildItem -LiteralPath $folder -File | Sort-Object Name | ForEach-Object {$_.Name+':'+(Get-FileHash -LiteralPath $_.FullName).Hash}) -join '|'
    if ($readProfile.ProfileKey -cne 'client_eu-prod' -or $before -cne $after) {throw 'Profile reading modified the configuration or returned the wrong profile.'}
    $legacy=Join-Path $folder 'client_eu-prod.local.json'
    [IO.File]::WriteAllText($legacy,'{"ProfileKey":"other"}')
    $rejected=$false;try {Read-InstallerProfileDocument -Root $fixture -ProfileKey 'client_eu-prod' | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'Conflicting JSON profile files were accepted.'}
    Remove-Item -LiteralPath $legacy
    $preferred=Join-Path $folder 'client_eu-prod.local.json.txt'
    Remove-Item -LiteralPath $preferred
    [IO.File]::WriteAllText($legacy,'{"ProfileKey":"client_eu-prod"}')
    $legacyRead=Read-InstallerProfileDocument -Root $fixture -ProfileKey 'client_eu-prod'
    if ($legacyRead.ProfileKey -cne 'client_eu-prod' -or (Test-Path -LiteralPath $preferred)) {throw 'Legacy profile was not read without migration.'}
    [IO.File]::WriteAllText($preferred,'{"ProfileKey":"client_eu-prod"}')
    $both=Read-InstallerProfileDocument -Root $fixture -ProfileKey 'client_eu-prod'
    if ($both.ProfileKey -cne 'client_eu-prod' -or -not (Test-Path -LiteralPath $legacy)) {throw 'Identical JSON profile names were changed.'}
    Remove-Item -LiteralPath $legacy
    function Write-InstallerMessage {param([string]$Message)}
    $list=@(Get-InstallerAvailableProfiles -Root $fixture)
    if ($list.Count -ne 1 -or $list[0] -cne 'client_eu-prod') {throw 'Profile discovery accepted a missing, invalid or conflicting profile.'}
    $answers=[Collections.Generic.Queue[string]]::new();foreach ($value in @('test','2','','1')) {$answers.Enqueue($value)}
    function Read-Host {param([string]$Prompt) $answers.Dequeue()}
    $selected=Read-InstallerProfile -Root $fixture
    if ($selected -cne 'client_eu-prod' -or $answers.Count -ne 0) {throw 'Profile selection did not reject non-list entries.'}
    [IO.File]::WriteAllText($preferred,'{"ProfileKey":"client_eu-prod","ServiceAccount":"DEMO\\svc-inventory"}')
    $hashBefore=(Get-FileHash -LiteralPath $preferred).Hash
    $answers.Enqueue('')
    if ((Read-InstallerServiceAccount -Root $fixture -ProfileKey 'client_eu-prod') -cne 'DEMO\svc-inventory') {throw 'Configured account was not offered as the default.'}
    $answers.Enqueue('DEMO\other')
    if ((Read-InstallerServiceAccount -Root $fixture -ProfileKey 'client_eu-prod') -cne 'DEMO\other') {throw 'The configured account could not be changed interactively.'}
    if ((Get-FileHash -LiteralPath $preferred).Hash -cne $hashBefore) {throw 'Installer account selection changed the profile.'}
    [IO.File]::WriteAllText($preferred,'{"ProfileKey":"client_eu-prod"}')
    $answers.Enqueue('');$answers.Enqueue('user@example.invalid')
    if ((Read-InstallerServiceAccount -Root $fixture -ProfileKey 'client_eu-prod') -cne 'user@example.invalid' -or $answers.Count) {throw 'A legacy profile did not require a valid account.'}
    [IO.File]::WriteAllText($preferred,'{"ProfileKey":"client_eu-prod","ServiceAccount":"DEMO\\svc$"}')
    $rejected=$false;try {Read-InstallerServiceAccount -Root $fixture -ProfileKey 'client_eu-prod' | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'A managed account in the profile was accepted.'}
    [IO.File]::WriteAllText($preferred,'{"ProfileKey":"client_eu-prod"}')
    $script:observedRead=''
    function Read-SmartM365JsonDocument {param($Path,$Validate) $script:observedRead=$Path;[pscustomobject]@{Document=[pscustomobject]@{ProfileKey='client_eu-prod'}}}
    $uncProfile=Read-InstallerProfileDocument -Root '\\offline.invalid\share\SmartM365' -ProfileKey 'client_eu-prod'
    if ($uncProfile.ProfileKey -cne 'client_eu-prod' -or $script:observedRead -cne '\\offline.invalid\share\SmartM365\Config\Tenants\client_eu-prod.local.json') {throw 'UNC profile did not use the read-only JSON reader.'}
    Remove-Item Function:\Read-SmartM365JsonDocument
    Remove-Item -LiteralPath (Join-Path $folder 'client_eu-prod.local.json.txt')
    $rejected=$false;try {Read-InstallerProfile -Root $fixture | Out-Null} catch {$rejected=$_.Exception.Message -match 'No valid tenant profiles'}
    if (-not $rejected) {throw 'An empty profile list was accepted.'}
} finally {
    if (Test-Path -LiteralPath $fixture) {
        $resolved=(Resolve-Path -LiteralPath $fixture).Path
        if ($resolved -ne $fixture -or [IO.Path]::GetFileName($resolved) -notmatch '^si-profile-[0-9a-f]{32}$') {throw 'Unexpected fixture path.'}
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
if ($source -notmatch '\$Tenant = Read-InstallerProfile -Root \$smartM365Root' -or $source -notmatch 'Get-InstallerAvailableProfiles -Root \$smartM365Root') {throw 'The installer does not enforce profile selection.'}
'Task installer contract fixtures passed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD5iZC08Ai2m6jJ
# 0CK46rPsA0xBLlBMnOEvJ5ydwFDkSaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKHoLADvSpSqG+Dsvs9DabF3xPZS+V6B7MI7Nh+jZxt2MA0GCSqG
# SIb3DQEBAQUABIIBgGzolitVoVRjEX/X8tAJ3fUqNBeY+E/KdB5YjKTqTa8XkeFh
# 2+urvrPCKhlC+iI3DDoY4B83RJqU+eMfexY57AmLHFUnLzkF8g/naRKvXZPGuXU9
# XzN7HyVvsMJRdMaX7B6bGaNw29LH1Fslqd/G1ilTVouCk2YFIFC7Uif6AWxaMngd
# V+QqtpT5iDMRhavZ4mLeLde6OnP8mFd7xqZnO9vZgxACTC+puf7wv8YQMOQK7R4A
# pxwBL05nDI7IaQ4UWbgRdPqK54ZIVs/fSrep3bEnep9igPhc0OhaQaBJtdevJJ/u
# p8E5Mt4p867+PgMbcNv6/TR/mwzpNuxg5ir/opDSaEO8yFwunJNrOxXlSICXeZKo
# f3GFmDR4q7154Iy1TojEOfk6W5vyV2OepYM1DsAdo7MmVZj3SFpIGXW9hbwnfxSI
# G4fMYB8EzoUvkOOa7KUrN41qo8cA1beDQ2IQ+b/swOydRBlMAqSKdO7O5g9rDP4n
# poYSbOJKHJvPgbZU3qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTFaMC8GCSqGSIb3DQEJBDEiBCCovfM0PXuBy34l7saHt1IfHHTkjSKHBQwZwiM3
# 3qECGTANBgkqhkiG9w0BAQEFAASCAgAMlsxKr58HZzyflfgOaWPX+n343SU8eV+Q
# 19eHsGzfOmhSFgFx/eln3zF/nwcxkPWKxmplXgDtUbtBwN3OmT/N1KcZ9ZKyro3Q
# Il0nUXjDIXH2cm3hDPH4sXldLXeca8/E5bNkOlQghfzAUfDpkWOzF0VFuRQGrZYF
# HEYTayXTFcvum88+BT/F6/Ym0HJY0m8TPZz8vjfOaf1JDcom+gdonuWyDKhZsveI
# gTD6AAb93vI9nUNKERxWF+/4zPS4ncHbT34IyFUh6EnWNnU6Vubvw97GAPqfVmW7
# N3njrHYF8yPcg/opqPf+s6iOlTN28vsD1LQ4LDyj+gPYnjALbS1YwK6gqMRVltDx
# P+I+RtmP7PkXP6plYeNFJRv+OtMyLJBD7HDN6dkeCwdIZ7L1YW2hXYSiMEcDvtW+
# HWdHQ3qG66vcPF9fYaH/Scuqd56+dniaNIHDcaAwKYgbOs2Qh+bzk05Tr7Njgx/s
# VXhv/yADrEMZSEh6KH/5TR42XCiZbiP0Eow3ROyuAMVMdCyjOmgpq9sUVRm7mxUX
# 08BUGFaiM+T55tik7HKoxFGuX3f+oxq+4ZBvjWFps4L9fbrbYfXlaXHvLgJ+eyUv
# SExbq49pXab2utxxQf23/rIkS2RAPUQKyar4znD5fYbC9kgPV8hp0rayvj9plvx/
# tig/9joi3w==
# SIG # End signature block
