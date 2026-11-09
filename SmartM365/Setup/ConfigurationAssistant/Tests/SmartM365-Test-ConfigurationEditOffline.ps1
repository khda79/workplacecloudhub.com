# SmartInventory local configuration edit regression, version 0.4.0.
$ErrorActionPreference = 'Stop'
$assistantRoot=Split-Path -Parent $PSScriptRoot
$root = Split-Path -Parent (Split-Path -Parent $assistantRoot)
. (Join-Path $root 'Config/SmartM365-TenantContext.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationAudit.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationEdit.ps1')
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('si-edit-' + [guid]::NewGuid().ToString('N'))
try {
    $folder = Join-Path $fixture 'Config/Tenants'
    $null = New-Item -ItemType Directory -Path $folder -Force
    [IO.File]::WriteAllText((Join-Path $folder 'tenant.local.json.template'), '{"OrgDomain":"","Nested":{"New":true}}')
    $path = Join-Path $folder 'sample.local.json'
    [IO.File]::WriteAllText($path, '{"OrgDomain":"old.invalid","Private":"keep","Nested":{"Existing":42}}')
    $snapshot = Get-SiEditSnapshot -Root $fixture -Tenant 'sample' -Step 'tenant' -Field 'OrgDomain'
    if ($snapshot.Value -ne 'old.invalid') { throw 'Incorrect snapshot.' }
    $preview=Get-SiEditTemplatePreview -Snapshot $snapshot
    if (@($preview.Additions | Where-Object { $_.Field -eq 'Nested.New' -and $_.ValueJson -eq 'true' }).Count -ne 1) { throw 'Template additions were not shown for review.' }
    $before=(Get-FileHash $path).Hash
    [IO.File]::AppendAllText((Join-Path $folder 'tenant.local.json.template'),' ')
    $rejected=$false; try { Set-SiEditSnapshot -Snapshot $snapshot -Value 'new.invalid' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null } catch { $rejected=$true }
    if (-not $rejected -or (Get-FileHash $path).Hash -ne $before) { throw 'Changed template was accepted or target changed.' }
    $preview=Get-SiEditTemplatePreview -Snapshot $snapshot
    Set-SiEditSnapshot -Snapshot $snapshot -Value 'new.invalid' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $actual = Get-Content $path -Raw | ConvertFrom-Json
    if ($actual.OrgDomain -ne 'new.invalid' -or $actual.Private -ne 'keep' -or $actual.Nested.Existing -ne 42 -or $actual.Nested.New -ne $true) { throw 'Edit or template enrichment lost data.' }
    $global=Join-Path $fixture 'Config/SmartM365.global.local.json'
    [IO.File]::WriteAllText($global,'{"RetentionMaxCSV":30,"OnPremExecutionMode":"LocalCache","Private":"keep"}')
    [IO.File]::WriteAllText((Join-Path $fixture 'Config/SmartM365.global.local.json.template'),'{"RetentionMaxCSV":30}')
    $retention=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field RetentionMaxCSV
    $globalHash=(Get-FileHash $global).Hash
    $rejected=$false; try { Set-SiEditSnapshot -Snapshot $retention -Value '-1' | Out-Null } catch { $rejected=$true }
    if (-not $rejected -or (Get-FileHash $global).Hash -ne $globalHash) { throw 'Invalid numeric retention was accepted.' }
    $globalPreview=Get-SiEditTemplatePreview -Snapshot $retention
    Set-SiEditSnapshot -Snapshot $retention -Value '42' -ExpectedTemplateHash $globalPreview.Hash -ExpectedAddedKeys $globalPreview.Keys | Out-Null
    $globalValue=Get-Content -LiteralPath $global -Raw | ConvertFrom-Json
    if ($globalValue.RetentionMaxCSV -isnot [long] -or $globalValue.RetentionMaxCSV -ne 42 -or $globalValue.Private -ne 'keep') { throw 'Numeric retention type or unrelated configuration was not preserved.' }
    $mode=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field OnPremExecutionMode
    $rejected=$false; try { Set-SiEditSnapshot -Snapshot $mode -Value 'Unrecognized' | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'Invalid on-premises launch mode was accepted.' }
    [IO.File]::WriteAllText($global,'{"RetentionMaxCSV":42,"OnPremExecutionMode":"LocalCache","DefaultTenant":"sample","EnableSharePointUploadDataAll":"__USE_GLOBAL__","MailBrandingEnabled":true,"MailPowerBiEnabled":false,"MailClientLogoMaxKB":200,"ReportMailSchedule":{"Enabled":false,"NotBefore":"07:30","TimeZone":"UTC","QueuePath":"","TrustedSids":[]},"Private":"keep"}')
    $default=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field DefaultTenant
    $before=(Get-FileHash -LiteralPath $global).Hash;$rejected=$false;try {Set-SiEditSnapshot -Snapshot $default -Value '../bad' | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $global).Hash -ne $before) {throw 'Invalid default profile was accepted.'}
    $preview=Get-SiEditTemplatePreview -Snapshot $default;Set-SiEditSnapshot -Snapshot $default -Value 'new-test' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $category=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field EnableSharePointUploadDataAll
    $preview=Get-SiEditTemplatePreview -Snapshot $category;Set-SiEditSnapshot -Snapshot $category -Value 'false' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $brandingGlobal=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field MailBrandingEnabled
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $brandingGlobal -Value '__USE_GLOBAL__' | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'A global Boolean accepted an inheritance marker.'}
    $preview=Get-SiEditTemplatePreview -Snapshot $brandingGlobal;Set-SiEditSnapshot -Snapshot $brandingGlobal -Value 'false' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $globalResult=Get-Content -LiteralPath $global -Raw | ConvertFrom-Json
    if ($globalResult.DefaultTenant -ne 'new-test' -or $globalResult.EnableSharePointUploadDataAll -isnot [bool] -or $globalResult.EnableSharePointUploadDataAll -ne $false -or $globalResult.MailBrandingEnabled -isnot [bool] -or $globalResult.MailBrandingEnabled -ne $false -or $globalResult.Private -ne 'keep') {throw 'Global settings were not saved with their types or unrelated data was changed.'}
    $schedule=Get-SiEditSnapshot -Root $fixture -Tenant sample -Step share -Field ReportMailSchedule
    $before=(Get-FileHash -LiteralPath $global).Hash
    foreach ($bad in @('{"Enabled":true,"NotBefore":"07:30","TimeZone":"UTC","QueuePath":"","TrustedSids":[]}','{"Enabled":false,"NotBefore":"25:00","TimeZone":"UTC","QueuePath":"","TrustedSids":[]}','{"Enabled":false,"NotBefore":"07:30","TimeZone":"UTC","QueuePath":"","TrustedSids":[],"Secret":"synthetic"}')) {$rejected=$false;try {Set-SiEditSnapshot -Snapshot $schedule -Value $bad | Out-Null} catch {$rejected=$true};if (-not $rejected -or (Get-FileHash -LiteralPath $global).Hash -ne $before) {throw 'An invalid or sensitive report schedule was accepted.'}}
    $preview=Get-SiEditTemplatePreview -Snapshot $schedule
    Set-SiEditSnapshot -Snapshot $schedule -Value '{"Enabled":true,"NotBefore":"08:00","TimeZone":"UTC","QueuePath":"{{WorkspaceRootPath}}\\Mail","TrustedSids":["S-1-5-18"]}' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $globalResult=Get-Content -LiteralPath $global -Raw | ConvertFrom-Json
    if (-not $globalResult.ReportMailSchedule.Enabled -or $globalResult.ReportMailSchedule.TrustedSids.Count -ne 1 -or $globalResult.Private -ne 'keep') {throw 'Report schedule edit lost data.'}
    $advanced=Join-Path $folder 'advanced.local.json'
    [IO.File]::WriteAllText($advanced,'{"OrgDomain":"example.invalid","Environment":"Test","AdComputerNameCountryPrefixes":{"FR":"France"},"AdDomainCountryMappings":{"old.example.test":"France"},"DeviceSystemMaxObservationAgeHours":null,"DeviceSystemOffsetlessGraphTimeZoneId":"","ExpectedThirdPartyAV":{"DisplayName":"","RegisteredNames":[],"Additional":"preserve"},"EndpointProtectionPolicy":{"Enabled":false,"PolicyVersion":"","MaxObservationAgeHours":null,"Requirements":[],"ScopeRules":[]},"MailBrandingEnabled":"__USE_GLOBAL__","MailPowerBiEnabled":"false","MailClientLogoMaxKB":"__USE_GLOBAL__","Private":"preserve"}')
    $map=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field AdComputerNameCountryPrefixes
    $before=(Get-FileHash -LiteralPath $advanced).Hash
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $map -Value '{"F":"France"}' | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'Invalid country prefix was accepted.'}
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $map -Value '[{"FR":"France"}]' | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'An array was accepted as a country mapping.'}
    $preview=Get-SiEditTemplatePreview -Snapshot $map
    Set-SiEditSnapshot -Snapshot $map -Value '{"FR":"France","GB":"United Kingdom"}' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $result=Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json
    if ($result.AdComputerNameCountryPrefixes.GB -ne 'United Kingdom' -or $result.Private -ne 'preserve') {throw 'Country mapping edit did not preserve unrelated values.'}
    $domainMap=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field AdDomainCountryMappings
    $before=(Get-FileHash -LiteralPath $advanced).Hash
    foreach ($bad in @('{"*.example.test":"France"}','{"a.example.test":"France","A.example.test.":"Germany"}')) {$rejected=$false;try {Set-SiEditSnapshot -Snapshot $domainMap -Value $bad | Out-Null} catch {$rejected=$true};if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'Invalid domain mapping was accepted.'}}
    $preview=Get-SiEditTemplatePreview -Snapshot $domainMap;Set-SiEditSnapshot -Snapshot $domainMap -Value '{"north.example.test":"France"}' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $age=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field DeviceSystemMaxObservationAgeHours
    if ($age.Value -ne 'null') {throw 'Null observation age was not displayed.'}
    foreach ($bad in @('0','-1','NaN','Infinity')) {$rejected=$false;try {Set-SiEditSnapshot -Snapshot $age -Value $bad | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'Invalid observation age was accepted.'}}
    $preview=Get-SiEditTemplatePreview -Snapshot $age;Set-SiEditSnapshot -Snapshot $age -Value '48.5' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $result=Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json
    if ($result.DeviceSystemMaxObservationAgeHours -ne 48.5 -or $result.AdDomainCountryMappings.'north.example.test' -ne 'France' -or $result.Private -ne 'preserve') {throw 'Advanced tenant settings were not saved safely.'}
    $age=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field DeviceSystemMaxObservationAgeHours;$preview=Get-SiEditTemplatePreview -Snapshot $age;Set-SiEditSnapshot -Snapshot $age -Value 'null' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    if ($null -ne (Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json).DeviceSystemMaxObservationAgeHours) {throw 'Observation age was not cleared to null.'}
    $zone=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field DeviceSystemOffsetlessGraphTimeZoneId
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $zone -Value 'Synthetic/Not-A-Timezone' | Out-Null} catch {$rejected=$true};if (-not $rejected) {throw 'Invalid time zone was accepted.'}
    $policy=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field EndpointProtectionPolicy
    $before=(Get-FileHash -LiteralPath $advanced).Hash
    foreach ($bad in @('{"Enabled":true,"PolicyVersion":"","MaxObservationAgeHours":null,"Requirements":[],"ScopeRules":[]}','{"Enabled":true,"PolicyVersion":"v1","MaxObservationAgeHours":0,"Requirements":[],"ScopeRules":[]}','{"Enabled":true,"PolicyVersion":"v1","MaxObservationAgeHours":72,"Requirements":[{"Role":"EDR","Product":"x","ExpectedMode":"Passive"}],"ScopeRules":[]}','{"Enabled":false,"PolicyVersion":"","MaxObservationAgeHours":null,"Requirements":[],"ScopeRules":[],"Secret":"synthetic"}')) {$rejected=$false;try {Set-SiEditSnapshot -Snapshot $policy -Value $bad | Out-Null} catch {$rejected=$true};if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'An invalid or sensitive endpoint policy was accepted.'}}
    $validPolicy='{"Enabled":true,"PolicyVersion":"v1","MaxObservationAgeHours":72,"Requirements":[{"Role":"Antivirus","Product":"Synthetic AV","ExpectedMode":"Active","EvidenceProductName":"Synthetic AV","Reason":""}],"ScopeRules":[]}'
    $preview=Get-SiEditTemplatePreview -Snapshot $policy;Set-SiEditSnapshot -Snapshot $policy -Value $validPolicy -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $policyResult=(Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json).EndpointProtectionPolicy
    if (-not $policyResult.Enabled -or $policyResult.Requirements.Count -ne 1 -or $policyResult.Requirements[0].Role -ne 'Antivirus') {throw 'Endpoint policy edit lost typed values.'}
    $evidenceModule=Join-Path (Split-Path -Parent $root) 'SmartWorkplaceIntelligence/scripts/EndpointProtectionEvidence.psm1'
    if ([IO.File]::Exists($evidenceModule)) {
        Import-Module -Name $evidenceModule -Force -ErrorAction Stop
        $consumerPolicy=ConvertTo-EndpointProtectionPolicy -Configuration @{EndpointProtectionPolicy=$policyResult} -TenantKey 'synthetic'
        if (-not $consumerPolicy.Enabled -or $consumerPolicy.Requirements.Count -ne 1) {throw 'The downstream policy consumer rejected the synthetic saved policy.'}
    } else {Write-Output 'Downstream endpoint policy comparison skipped: module absent.'}
    $av=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step tenant -Field ExpectedThirdPartyAV
    $before=(Get-FileHash -LiteralPath $advanced).Hash;$rejected=$false
    try {Set-SiEditSnapshot -Snapshot $av -Value '{"DisplayName":"Synthetic AV","RegisteredNames":[]}' | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'Incomplete antivirus setting was accepted.'}
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $av -Value '[{"DisplayName":"Synthetic AV","RegisteredNames":["Synthetic AV"]}]' | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $advanced).Hash -ne $before) {throw 'An array was accepted as antivirus configuration.'}
    $preview=Get-SiEditTemplatePreview -Snapshot $av
    Set-SiEditSnapshot -Snapshot $av -Value '{"DisplayName":"Synthetic AV","RegisteredNames":["Synthetic AV"],"Additional":"preserve"}' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $result=Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json
    if ($result.ExpectedThirdPartyAV.RegisteredNames.Count -ne 1 -or $result.ExpectedThirdPartyAV.Additional -ne 'preserve' -or $result.Private -ne 'preserve') {throw 'Antivirus configuration or unrelated values were lost.'}
    $branding=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step mail -Field MailBrandingEnabled
    $preview=Get-SiEditTemplatePreview -Snapshot $branding
    Set-SiEditSnapshot -Snapshot $branding -Value 'true' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $size=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step mail -Field MailClientLogoMaxKB
    $preview=Get-SiEditTemplatePreview -Snapshot $size
    Set-SiEditSnapshot -Snapshot $size -Value '350' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $powerbi=Get-SiEditSnapshot -Root $fixture -Tenant advanced -Step mail -Field MailPowerBiEnabled
    if ($powerbi.Value -ne 'false') {throw 'String boolean setting could not be reviewed.'}
    $result=Get-Content -LiteralPath $advanced -Raw | ConvertFrom-Json
    if ($result.MailBrandingEnabled -isnot [bool] -or -not $result.MailBrandingEnabled -or $result.MailClientLogoMaxKB -ne 350 -or $result.Private -ne 'preserve') {throw 'Mail settings lost their types or unrelated values.'}
    $missing=Join-Path $folder 'missing.local.json'
    [IO.File]::WriteAllText($missing,'{"OrgDomain":"example.invalid","Private":"preserve"}')
    [IO.File]::WriteAllText((Join-Path $folder 'tenant.local.json.template'),'{"OrgDomain":"","MailBrandingEnabled":"__USE_GLOBAL__"}')
    $absent=Get-SiEditSnapshot -Root $fixture -Tenant missing -Step mail -Field MailBrandingEnabled
    if ($absent.Present -or $absent.Value -ne '__USE_GLOBAL__') {throw 'Missing setting did not load its template default.'}
    $preview=Get-SiEditTemplatePreview -Snapshot $absent
    if (@($preview.Additions | Where-Object {$_.Field -eq 'MailBrandingEnabled'}).Count -ne 1) {throw 'Missing setting was not included in the review.'}
    $before=(Get-FileHash -LiteralPath $missing).Hash
    [IO.File]::AppendAllText((Join-Path $folder 'tenant.local.json.template'),' ')
    $rejected=$false;try {Get-SiEditTemplatePreview -Snapshot $absent | Out-Null} catch {$rejected=$true}
    if (-not $rejected) {throw 'Changed template was accepted during missing-setting preview.'}
    $rejected=$false;try {Set-SiEditSnapshot -Snapshot $absent -Value 'true' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null} catch {$rejected=$true}
    if (-not $rejected -or (Get-FileHash -LiteralPath $missing).Hash -ne $before) {throw 'A changed template was accepted for a missing setting.'}
    $absent=Get-SiEditSnapshot -Root $fixture -Tenant missing -Step mail -Field MailBrandingEnabled
    $preview=Get-SiEditTemplatePreview -Snapshot $absent
    Set-SiEditSnapshot -Snapshot $absent -Value '__USE_GLOBAL__' -ExpectedTemplateHash $preview.Hash -ExpectedAddedKeys $preview.Keys | Out-Null
    $result=Get-Content -LiteralPath $missing -Raw | ConvertFrom-Json
    if ($result.MailBrandingEnabled -ne '__USE_GLOBAL__' -or $result.Private -ne 'preserve') {throw 'Missing default was not added safely.'}
    $snapshot = Get-SiEditSnapshot -Root $fixture -Tenant 'sample' -Step 'tenant' -Field 'OrgDomain'
    [IO.File]::WriteAllText($path, '{"OrgDomain":"concurrent.invalid","Private":"keep"}')
    try { Set-SiEditSnapshot -Snapshot $snapshot -Value 'must-not-write.invalid' | Out-Null; throw 'Conflict accepted.' }
    catch { if ($_.Exception.Message -eq 'Conflict accepted.') { throw } }
    if ((Get-Content $path -Raw | ConvertFrom-Json).OrgDomain -ne 'concurrent.invalid') { throw 'Concurrent data was overwritten.' }
    [IO.File]::WriteAllText($path + '.txt', '{}')
    try { Get-SiEditSnapshot -Root $fixture -Tenant 'sample' -Step 'tenant' -Field 'OrgDomain' | Out-Null; throw 'Coexistence accepted.' }
    catch { if ($_.Exception.Message -eq 'Coexistence accepted.') { throw } }
    try { Get-SiEditSnapshot -Root $fixture -Tenant 'sample' -Step 'tenant' -Field 'Private' | Out-Null; throw 'Unlisted field accepted.' }
    catch { if ($_.Exception.Message -eq 'Unlisted field accepted.') { throw } }
    [IO.File]::WriteAllText((Join-Path $folder 'tenant.local.json.template'), '{"ProfileKey":"tenant","OrganizationKey":"placeholder","EnvironmentKey":"test","TenantKey":"placeholder-test","Environment":"Test","AppId":"00000000-0000-0000-0000-000000000000","EnableTeamsNotifications":true,"EnableSharePointUpload":false,"SharePointSitePath":"/sites/synthetic"}')
    $created = New-SiTenantProfile -Root $fixture -ProfileKey 'new-test' -OrganizationKey 'new' -EnvironmentKey 'test'
    $profile = Get-Content -LiteralPath $created.Path -Raw | ConvertFrom-Json
    if ($profile.ProfileKey -ne 'new-test' -or $profile.TenantKey -ne 'new-test') { throw 'New profile identity mismatch.' }
    if ($profile.EnableTeamsNotifications -ne $false -or $profile.EnableSharePointUpload -ne $true -or $profile.SharePointSitePath -ne '/sites/synthetic') {throw 'New profile defaults did not preserve the site template and apply the selected feature defaults.'}
    $before = (Get-FileHash $created.Path).Hash
    try { New-SiTenantProfile -Root $fixture -ProfileKey 'new-test' -OrganizationKey 'other' -EnvironmentKey 'test' | Out-Null; throw 'Existing profile accepted.' }
    catch { if ($_.Exception.Message -eq 'Existing profile accepted.') { throw } }
    if ((Get-FileHash $created.Path).Hash -ne $before) { throw 'Existing profile overwritten.' }
    try { New-SiTenantProfile -Root $fixture -ProfileKey 'invalid' -OrganizationKey 'bad_key' -EnvironmentKey 'test' | Out-Null; throw 'Invalid organization key accepted.' }
    catch { if ($_.Exception.Message -eq 'Invalid organization key accepted.') { throw } }
    'Configuration edit fixtures passed.'
} finally { if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Recurse -Force } }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCV5he9HqySz/Rn
# jQj7zF2E4Thnj7tq4I6FGQyRHwjUuKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBCXm1++Hd9g2OJDr/qzKZF0IM6KmHedqlXWHwTwg1OhMA0GCSqG
# SIb3DQEBAQUABIIBgCqC8jpoqpKE6heNoly1Bn+cNi9+TkYDD/saQ94eA2YXJRiT
# srIdxratMAL89u5kbt4Lx12h58IU2TjF3VGo6kWI8HK3redr7oVmJ7/Cz83ZLFuO
# HUv7rV+YJDCtIUd+Nbqr/nTEo/fPFJVYMU8kSQm9SmaeTzeH4aydKE0haiUOhd0C
# lLNNwOGc680GZx9qStTumMTqtzFtrBvIVJDwudfodvm1FqSdvomPts6jSRIEQlNy
# Jfth9NJ+g4MLwe0CV+zhs+CbYn5DMu5yo5LSodr1ALMGYfqmK3mVpeVUyemP9P6P
# W8FfKgZh/CCK7nI06k180KFHHy/uJmg37QzHxdE/3RUoNEEYVsFW4Ab2NkRfFHd4
# VdrEluzRrgfqX1i6yS2OF0PvegEEqCfRA/zGPhj8vUpYA6zhS/n/MJTW/TeX/srh
# /76g8bHKdl8Yr68LNC5QGZclTkv8qFI3HtUYi3RFKTshlj7x3vgI8XpGcE0pFp6j
# YH8VDAe6qSIhWAh+mKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MDNaMC8GCSqGSIb3DQEJBDEiBCA5V7RQL3x9Ho1hNS9TdN21meBcR4vPp26V15E3
# FchqgjANBgkqhkiG9w0BAQEFAASCAgAG9yo/tTaWpMmFrRJcEtxziSdoUDBYyEuL
# JCwuE43FSju2rmObAWZZ59VhbNnK/fLUh/ycQggHZxuV5MOujYEhYy+xT+uYcm2t
# +m5/GY6bXHxxgLlulzdmE/WXng497RlTU54hXMceSQYrlCmDbxi0+RoIJrmHnAPM
# 9aOVvhNnpeWLNsIqqVpEDRvS/MzG1ZJYIqtg0lDcPrQsJb9yIc8kjkjjTyC6C9I3
# NaiKLlQOPP0rwc+fi9bpHtve60VERiJu/vQspPIdypBJa+CBAy3d8tZoNjHGzA85
# CyjoBDPImJX5mwLYgYzKWy4wF9O06wa+vNZZdRHMPB7mpBtnCxw0aREuqqYKelSj
# MgQwEqj1GNnyqcsmCJTVKGS0hwgLnslTwA9SsFQ8j9tyZJbB92KYpST7DODWj3wW
# sCy6YUrD7NdvakH9l190GqgIhHygQwNriTeVttzlZ/kGT1PvN0JxO3rhpvdpNrZr
# kiexJQo9y9DMftgm/rS3XIzAe8tlHz+Ok/FbKv9ROVF7SNaitD4UzpGzeuyoLpjH
# zXKPtyThdhC2XE5/7YlUYoemvfUD0FggjG/C/Ar3jrXgjIMixkM0mW5dKutHCnUG
# Isyd/68GxRw4EmTL3+wvuXop+O/ZPSo7ZdN4XTVgcfWOWyxeP/0SS/fx84SY5fiL
# T5WTFYCGXg==
# SIG # End signature block
