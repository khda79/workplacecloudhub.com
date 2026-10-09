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
if ($template.EnableSharePointUpload -ne $false -or $template.EnableWeeklyHistory -ne $true) { throw 'Content template policy is invalid.' }
$registry = Get-Content $registryPath -Raw | ConvertFrom-Json
$producer = @($registry.Producers | Where-Object Script -eq 'SmartM365-SharePoint-OnPrem-Content-Inventory.ps1')
if ($producer.Count -ne 1 -or $producer[0].Files.Count -ne 4) { throw 'Content source receipt registration is invalid.' }
if ($ast.Extent.Text -match '(?im)^\s*(Set-SPSite|Set-SPWeb|Set-SPContentDatabase|Add-SPShellAdmin|Remove-SPSite)\b') { throw 'A SharePoint write command was found.' }
if ($ast.Extent.Text -match '(?im)^\s*Get-SPSiteAdministration\b') { throw 'Lock inspection must use SPSite and SPContentDatabase only.' }
if ($ast.Extent.Text -notmatch '(?m)^\s*\$site\s*=\s*\$_\s*$' -or $ast.Extent.Text -match '(?m)^\s*param\(\$site\)') { throw 'Streaming site pipeline input is not bound to the current site.' }
if ($ast.Extent.Text -notmatch '\$runBase\s*=\s*if\s*\(\$MaxItems\s*-gt\s*0\).*?TEST' -or $ast.Extent.Text -notmatch 'Select-Object\s+-First\s+\$remaining' -or $ast.Extent.Text -notmatch 'Flush-RunRows\s+-Kind\s+CollectionCoverage') { throw 'Limited or per-site coverage path is missing.' }
foreach ($name in @('Get-InventoryConfigValue','Get-ObservedProperty','Test-MissingObservation','Resolve-SiteLockObservation','Assert-Deadline','Get-CollectionFailureStatus','Get-DatabaseCoverageStatus','Write-RunCsv','Flush-RunRows','Invoke-DailySummaryMail','Send-InventorySummaryMail')) {
    $functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true))
    if ($functionAst.Count -ne 1) { throw "Missing function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$site = [pscustomobject]@{ ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }
$database = [pscustomobject]@{ IsReadOnly=$false }
$unknown = Resolve-SiteLockObservation -Site $site -ContentDatabase $database
if ($unknown.LockState -ne '' -or $unknown.LockStatus -ne 'Unverified' -or $unknown.IsReadOnly -ne 'False' -or $unknown.ContentDatabaseIsReadOnly -ne 'False') { throw 'Unknown lock state was converted to a false observation.' }
foreach ($state in @('Unlock','NoAdditions','ReadOnly','NoAccess')) {
    $result = Resolve-SiteLockObservation -Site ([pscustomobject]@{ LockState=$state; ReadOnly=$false; ReadLocked=$false; WriteLocked=$false; LockIssue='' }) -ContentDatabase $database
    if ($result.LockState -ne $state -or $result.LockStatus -ne 'Observed') { throw "Lock state mapping failed: $state" }
}
$readLocked = Resolve-SiteLockObservation -Site ([pscustomobject]@{ReadLocked=$true;WriteLocked=$true;ReadOnly=$true}) -ContentDatabase $database
if ($readLocked.LockState -ne 'NoAccess') { throw 'Read lock mapping failed.' }
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
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: content parser, template, lock mapping, timeouts, per-site errors, empty CSV, receipts, daily mail gate and routing.'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH2H+FpSBIJa2M
# h4DLd22BkhnAyDJZ1DMU376DYpRZt6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBl9pBvVcmRGS4xyDDry7Bb
# j55mNRlRtyNh9DEaAu41AzANBgkqhkiG9w0BAQEFAASCAYCO3LzE2N+U0N5lm8Hu
# MBLrnyR+9n+UALMoQCMDI4STpej6BgylAO1AtzGhiOfsJfqnltdYTlzkc+vQX6C7
# 3R/UGANVdf9UbXbKfZlyYfyVMGXYYf76GHSv/3CMyJiyzvCIL6+9XAsQ6p9nTgqR
# RQIILSjrxyHLw7J8RnpINwCupPTE+3k4HCvKzH5LrEztmzLq+7ZEfRZNkdiMEkUS
# VabLli1NcRtMvTuX0/8ah9ZfEG3kjJs2jtVDlR5SItgMuG2nmN02Jr3mByEZZT4a
# 5DE9+nNFfspBpQ3TY2G8BRjKdJyCHDy5XEhuDeQXaOYhSwvs6m6CC4NZjxMpUqvc
# 6+KQAjNWhnAaoyeryufCtONBWkogmBDzSBdaaJCBct96rmrNvLgSxx+7/6JXliS3
# jpTLCyFHSp2ZY/vm1ts/QS+SxbVYh0X8NS54HdhHelIq3Kf7y3Ye2TvaqjHuxG1/
# 92n7eYMS+Sfx99PbIlJ2hMBCD8Jg4dY4lRuuGKUiEQEbt+4=
# SIG # End signature block
