[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1'
$templatePath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Content-Inventory.local.json.template'
$registryPath = Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SourceReceipts.json.txt'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw "Content parser errors: $($errors.Count)" }
$template = Get-Content $templatePath -Raw | ConvertFrom-Json
if ($template.EnableSharePointUpload -ne $true -or $template.EnableWeeklyHistory -ne $true) { throw 'Content template policy is invalid.' }
if ($ast.Extent.Text -notmatch '(?s)\.VERSION\s+1\.0\.2') { throw 'Content version was not updated for lock-state verification.' }
$registry = Get-Content $registryPath -Raw | ConvertFrom-Json
$producer = @($registry.Producers | Where-Object Script -eq 'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1')
if ($producer.Count -ne 1 -or $producer[0].Files.Count -ne 4) { throw 'Content source receipt registration is invalid.' }
if ($ast.Extent.Text -match '(?im)^\s*(Set-SPSite|Set-SPWeb|Set-SPContentDatabase|Add-SPShellAdmin|Remove-SPSite)\b') { throw 'A SharePoint write command was found.' }
if ($ast.Extent.Text -match '(?im)^\s*Get-SPSiteAdministration\b') { throw 'Lock inspection must use SPSite and SPContentDatabase only.' }
if ($ast.Extent.Text -notmatch '(?m)^\s*\$site\s*=\s*\$_\s*$' -or $ast.Extent.Text -match '(?m)^\s*param\(\$site\)') { throw 'Streaming site pipeline input is not bound to the current site.' }
if ($ast.Extent.Text -notmatch '\$runBase\s*=\s*if\s*\(\$MaxItems\s*-gt\s*0\).*?TEST' -or $ast.Extent.Text -notmatch 'Select-Object\s+-First\s+\$remaining' -or $ast.Extent.Text -notmatch 'Flush-RunRows\s+-Kind\s+CollectionCoverage') { throw 'Limited or per-site coverage path is missing.' }
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-ObservedProperty','Test-MissingObservation','Resolve-SiteLockObservation','Get-LockStateLookup','Assert-Deadline','Get-CollectionFailureStatus','Get-DatabaseCoverageStatus','Write-RunCsv','Flush-RunRows','Invoke-DailySummaryMail','Send-InventorySummaryMail','Publish-QualifiedCsvUploads')) {
    $functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true))
    if ($functionAst.Count -ne 1) { throw "Missing function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$script:EffectiveConfig = [pscustomobject]@{
    SmartM365RootPath='C:\SmartM365'; WorkspaceRootPath='{{SmartM365RootPath}}'; ProfileKey='prod'
    DataAllRootPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'
    OutputRoot='{{DataAllRootPath}}\SharePoint\OnPrem\Content'
    WeeklyHistoryFolderPath='{{DataAllRootPath}}\SharePoint\OnPrem\Content\WeeklyHistory'
}
if ((Get-InventoryConfigValue 'OutputRoot') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Content') { throw 'Nested content OutputRoot tokens were not resolved.' }
if ((Get-InventoryConfigValue 'WeeklyHistoryFolderPath') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Content\WeeklyHistory') { throw 'Nested content history tokens were not resolved.' }
$script:EffectiveConfig.OutputRoot = '{{MissingPath}}\SharePoint'
try { $null = Get-InventoryConfigValue 'OutputRoot'; throw 'Missing token was accepted.' }
catch { if ($_.Exception.Message -ne 'Unresolved configuration value: OutputRoot') { throw } }
$root = 'C:\ProgramData\SmartM365\LauncherCache\SharePointContent\SmartM365'
try { $null = Assert-InventoryPath -Path (Join-Path $root 'Data\Tenants\prod\DATA-ALL') -Name 'OutputRoot'; throw 'Disposable content output path was accepted.' }
catch { if ($_.Exception.Message -notlike 'OutputRoot resolves inside the disposable launcher cache.*') { throw } }
if ((Assert-InventoryPath -Path 'C:\SmartM365\DATA' -Name 'OutputRoot') -ne 'C:\SmartM365\DATA') { throw 'Persistent content output path was rejected.' }
$site = [pscustomobject]@{ ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }
$database = [pscustomobject]@{ IsReadOnly=$false }
$unknown = Resolve-SiteLockObservation -Site $site -ContentDatabase $database
if ($unknown.LockState -ne '' -or $unknown.LockStatus -ne 'Unverified' -or $unknown.IsReadOnly -ne 'False' -or $unknown.ContentDatabaseIsReadOnly -ne 'False') { throw 'Unknown lock state was converted to a false observation.' }
foreach ($state in @('Unlock','NoAdditions','ReadOnly','NoAccess')) {
    $result = Resolve-SiteLockObservation -Site ([pscustomobject]@{ LockState=$state; ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }) -ContentDatabase $database
    if ($result.LockState -ne $state -or $result.LockStatus -ne 'Observed') { throw "Lock state mapping failed: $state" }
}
$writeLocked = Resolve-SiteLockObservation -Site ([pscustomobject]@{ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database
if ($writeLocked.LockState -ne '' -or $writeLocked.LockStatus -ne 'Unverified') { throw 'WriteLocked alone was misclassified as ReadOnly.' }
$noAdditions = Resolve-SiteLockObservation -Site ([pscustomobject]@{ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database -FilteredLockState 'NoAdditions'
if ($noAdditions.LockState -ne 'NoAdditions' -or $noAdditions.LockStatus -ne 'Observed') { throw 'Filtered NoAdditions mapping failed.' }
$conflicting = Resolve-SiteLockObservation -Site ([pscustomobject]@{LockState='ReadOnly';ReadLocked=$false;WriteLocked=$true;ReadOnly=$false}) -ContentDatabase $database -FilteredLockState 'NoAdditions'
if ($conflicting.LockState -ne '' -or $conflicting.LockStatus -ne 'Conflict') { throw 'Conflicting lock-state observations were accepted.' }
$script:GlobalDeadlineUtc = [datetime]::UtcNow.AddMinutes(1)
$script:MockLockSites = @(
    [pscustomobject]@{Id='site-a';Url='https://example.test/a';State='Unlock'},
    [pscustomobject]@{Id='site-b';Url='https://example.test/b';State='NoAdditions'}
)
$script:MockFailState = ''
$script:MockIgnoreFilter = $false
function Get-SPSite {
    [CmdletBinding()]
    param([string]$Identity, $ContentDatabase, [scriptblock]$Filter, [string]$Limit)
    if ($Filter.ToString() -notmatch "'(Unlock|NoAdditions|ReadOnly|NoAccess)'") { throw 'An unsupported lock-state filter was used.' }
    $requestedState = $Matches[1]
    if ($requestedState -eq $script:MockFailState) { throw 'MockFilterFailure' }
    foreach ($entry in $script:MockLockSites) {
        if ($Identity -and $entry.Url -ne $Identity) { continue }
        if (-not $script:MockIgnoreFilter -and $entry.State -ne $requestedState) { continue }
        $result = [pscustomobject]@{ Id=$entry.Id }
        $result | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
        $result
    }
}
$lookup = Get-LockStateLookup -ContentDatabase $database
if ($lookup.States.Count -ne 2 -or $lookup.States['site-a'] -ne 'Unlock' -or $lookup.States['site-b'] -ne 'NoAdditions' -or $lookup.Errors.Count -ne 0) { throw 'Database lock-state lookup failed.' }
$limitedLookup = Get-LockStateLookup -ContentDatabase $database -SiteUrl 'https://example.test/a'
if ($limitedLookup.States.Count -ne 1 -or $limitedLookup.States['site-a'] -ne 'Unlock') { throw 'Limited lock-state lookup traversed the wrong site.' }
$script:MockFailState = 'NoAccess'
$failedLookup = Get-LockStateLookup -ContentDatabase $database
if ($failedLookup.Errors.Count -ne 1 -or $failedLookup.States.Count -ne 2) { throw 'Lock-state query failure was hidden.' }
$script:MockFailState = ''
$script:MockIgnoreFilter = $true
$conflictLookup = Get-LockStateLookup -ContentDatabase $database
if ($conflictLookup.States.Count -ne 0 -or $conflictLookup.Conflicts.Count -ne 2) { throw 'Conflicting filter results were accepted.' }
Remove-Item Function:\Get-SPSite
if ((Get-CollectionFailureStatus 'Access is denied') -ne 'AccessDenied' -or (Get-CollectionFailureStatus 'CollectionTimeout') -ne 'TimedOut' -or (Get-CollectionFailureStatus 'unexpected') -ne 'Failed') { throw 'Per-collection failure classification failed.' }
if ((Get-DatabaseCoverageStatus -ExpectedCount 2 -ObservedCount 2) -ne 'Collected' -or (Get-DatabaseCoverageStatus -ExpectedCount 2 -ObservedCount 1) -ne 'DatabaseCoverageMismatch' -or (Get-DatabaseCoverageStatus -ExpectedCount '' -ObservedCount 0) -ne 'DatabaseCountUnavailable') { throw 'Database coverage qualification failed.' }
$script:GlobalDeadlineUtc = [datetime]::UtcNow.AddMinutes(-1)
$timedOut = $false
try { Assert-Deadline ([datetime]::UtcNow.AddMinutes(1)) } catch { $timedOut = $_.Exception.Message -eq 'GlobalTimeout' }
if (-not $timedOut) { throw 'Global timeout was not enforced.' }
$coreManifest = Join-Path $folder '..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module $coreManifest -Force -ErrorAction Stop
$global:SmartM365TenantKey = 'OFFLINE'
$global:SmartM365OrganizationKey = 'OFFLINE'
$global:SmartM365EnvironmentKey = 'TEST'
$global:SmartM365TenantId = '00000000-0000-0000-0000-000000000001'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPContent-Test-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $path = Write-RunCsv -Name 'Empty.csv' -Rows @() -Columns @('TenantKey','FarmId','LockState') -Folder $tempRoot
    $text = [IO.File]::ReadAllText($path)
    if ($text -notmatch 'TenantKey.*;.*FarmId.*;.*LockState' -or @(Import-Csv $path -Delimiter ';').Count -ne 0) { throw 'Empty content CSV contract failed.' }
    $buffers = @{ SiteCollections = (New-Object 'System.Collections.Generic.List[object]') }
    $paths = @{ SiteCollections = $path }
    $buffers.SiteCollections.Add([pscustomobject][ordered]@{TenantKey='OFFLINE';FarmId='FARM';LockState='NoAdditions'})
    $buffers.SiteCollections.Add([pscustomobject][ordered]@{TenantKey='OFFLINE';FarmId='FARM';LockState='ReadOnly'})
    Flush-RunRows -Kind SiteCollections -Buffers $buffers -Paths $paths -MinimumCount 2
    $actual = @(Import-Csv $path -Delimiter ';')
    if ($actual.Count -ne 2 -or $actual[0].LockState -ne 'NoAdditions' -or $actual[1].LockState -ne 'ReadOnly' -or $buffers.SiteCollections.Count -ne 0) { throw 'Bounded content CSV batch failed.' }
    $latest = Join-Path $tempRoot 'DATA-LAST'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest -ScopeParameters @{MaxItems=0}
    foreach ($name in $producer[0].Files) {
        if ($name -eq 'SharePoint_OnPrem_SiteCollections.csv') {
            Copy-SmartM365FileAtomically -SourcePath $path -DestinationPath (Join-Path $latest $name)
        } else {
            Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $latest $name) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
        }
    }
    $proofPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $proof = Get-Content $proofPath -Raw | ConvertFrom-Json
    if ($proof.Status -ne 'Completed' -or $proof.Files.Count -ne 4) { throw 'Content receipt completion failed.' }
    $proofHash = (Get-FileHash $proofPath -Algorithm SHA256).Hash
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest -ReadOnly -ScopeParameters @{MaxItems=2}
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'MaxItems changed the content receipt.' }
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    $null = Complete-SmartM365SourceReceipt -Status Failed -ErrorCount 1
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'A failed run changed the preceding content receipt.' }
    $marker = Join-Path $tempRoot 'Content-DailySummary.sent'
    $script:sendCount = 0
    function WriteLog { param($Message, $Level) }
    $sendAction = { $script:sendCount++ }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction)) { throw 'First content email was skipped.' }
    if (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction) { throw 'Same-day content email was repeated.' }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction -Force) -or $script:sendCount -ne 2) { throw 'Forced content email failed.' }
    $failedMarker = Join-Path $tempRoot 'Content-Failed.sent'
    try { $null = Invoke-DailySummaryMail -MarkerPath $failedMarker -SendAction { throw 'Simulated SMTP failure' }; throw 'Mail failure was ignored.' }
    catch { if ($_.Exception.Message -ne 'Simulated SMTP failure') { throw } }
    if (Test-Path -LiteralPath $failedMarker) { throw 'Failed content email wrote a marker.' }
    $script:mailParameters = $null
    function SendEmailHtmlReport { [CmdletBinding()] param($From,$To,$Subject,$BodyHtml,$MailPurpose,$SmtpServer,$SmtpPort,$SendMailMode,$Cc) $script:mailParameters = $PSBoundParameters }
    $script:EffectiveConfig = [pscustomobject]@{ From='sender@example.test'; To=''; ErrorMailTo='recipient@example.test'; SmtpServer='smtp.example.test'; SmtpPort=25; SendMailMode='SMTP'; Cc='' }
    Send-InventorySummaryMail -Subject 'Offline summary' -BodyHtml '<p>Summary</p>'
    if ($script:mailParameters.To -ne 'recipient@example.test' -or $script:mailParameters.From -ne 'sender@example.test' -or $script:mailParameters.ContainsKey('Attachments')) { throw 'Content mail routing or attachment policy failed.' }
    $script:uploadCalls = New-Object 'System.Collections.Generic.List[string]'
    function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploadCalls.Add($LocalFilePath); if ($LocalFilePath -eq 'failed.csv') { return }; return [pscustomobject]@{ LocalFilePath=$LocalFilePath } }
    $global:EnableSharePointUpload = $false
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv')) -ne 0 -or $script:uploadCalls.Count -ne 0) { throw 'Disabled content upload was attempted.' }
    $global:EnableSharePointUpload = $true
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv','run.csv','failed.csv')) -ne 1 -or $script:uploadCalls.Count -ne 3) { throw 'Content upload qualification failed.' }
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: content parser, nested configuration, persistent paths, lock mapping, timeouts, per-site errors, empty CSV, receipts, upload policy, daily mail gate and routing.'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBapBSJ46cIsgAy
# ppBUw7bHxHBUBjnV8PwmyMWdFRfTAaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDGHYbazpNfVOlsCcZqYzid
# UNidJPXkBOGWfApmIfpNjDANBgkqhkiG9w0BAQEFAASCAYARBTXHrf++D/qsDzw3
# cyfLMijO9BeRtpReg9rXZ5lgx5WjAGg83IxhE/eDWqnibatVKoOFVBE8jwS1DzZy
# S4XpSeJ7TbRq6FrjXCsKZTYZ4g9hKMj9OvHF6YUK5cPZBaA44v+p9miCH8d/a+Hr
# GTo2HweOjeldE7AAsPzAaeeYcK1EFYPEUt1tc6w7hpwmGur3ay7f1LYh58rh951+
# JP8Re6qcq1vjKOBYsEAm5oTBeOfslQw3ndZYPkuEI+h6Kvv05LJuJWwjPUhw9WGk
# 4gEEm5wxn/CS7+Q3Xt424qsXtpa3m0ARXk4SwswWF7miwzH7qvSk2i6dbniZUK1c
# U/n7wV9P6Z5I+Ftda9k+kkQXGd/LCcsQOumwDwVuW3s1a1I7X5J1Nyg53fRHrHOw
# lrqANqwYxotIyfTpqKrDeVSu6SkUdSk73vGcAyLHjwXG+gEPAJhtXzf1lmEwpQpP
# 88jjV2n6xJMSVThKsQMlsf64RTe1mLpQO4HkAMSvWMBSblU=
# SIG # End signature block
