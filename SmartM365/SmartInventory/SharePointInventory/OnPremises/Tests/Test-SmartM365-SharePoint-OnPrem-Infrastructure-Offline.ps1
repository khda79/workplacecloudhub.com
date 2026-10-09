[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1'
$templatePath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.local.json.template'
$registryPath = Join-Path $folder '..\..\..\Modules\SmartM365.Core\SmartM365-SourceReceipts.json.txt'
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw "Infrastructure parser errors: $($errors.Count)" }
if ($ast.Extent.Text -notmatch '(?s)\.VERSION\s+1\.0\.5') { throw 'Infrastructure version was not updated for Graph bootstrap and mail.' }
$firstUploadConfigRead = $ast.Extent.Text.IndexOf('$global:SharePointSiteHostname =', [StringComparison]::Ordinal)
$resolverDefinition = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Resolve-InventoryConfigTokens' }, $true))
if ($firstUploadConfigRead -lt 0 -or $resolverDefinition.Count -ne 1 -or $resolverDefinition[0].Extent.EndOffset -ge $firstUploadConfigRead) { throw 'Infrastructure reads upload configuration before the token resolver is defined.' }
$template = Get-Content $templatePath -Raw | ConvertFrom-Json
if ($template.EnableSharePointUpload -ne $true -or $template.EnableWeeklyHistory -ne $true -or $template.SendMailMode -ne 'Graph') { throw 'Infrastructure template policy is invalid.' }
$registry = Get-Content $registryPath -Raw | ConvertFrom-Json
$producer = @($registry.Producers | Where-Object Script -eq 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1')
if ($producer.Count -ne 1 -or $producer[0].Files.Count -ne 6) { throw 'Infrastructure source receipt registration is invalid.' }
if ($ast.Extent.Text -match '(?im)^\s*(Set-SPSite|Set-SPWeb|Set-SPContentDatabase|Add-SPShellAdmin|Remove-SPSite)\b') { throw 'A SharePoint write command was found.' }
if ($ast.Extent.Text -notmatch '-Rows\s+\$rows\[\$kind\]\.ToArray\(\)') { throw 'Infrastructure CSV buffers must be converted to arrays for Windows PowerShell 5.1.' }
foreach ($name in @('Get-InventoryConfigValue','Resolve-InventoryConfigTokens','Assert-InventoryPath','Get-ConfiguredWebApplications','Write-RunCsv','Ensure-GraphAuthenticationModule','Invoke-DailySummaryMail','Send-InventorySummaryMail','Publish-QualifiedCsvUploads')) {
    $functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true))
    if ($functionAst.Count -ne 1) { throw "Missing function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$script:EffectiveConfig = [pscustomobject]@{
    SmartM365RootPath='C:\SmartM365'; WorkspaceRootPath='{{SmartM365RootPath}}'; ProfileKey='prod'
    DataAllRootPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-ALL'
    OutputRoot='{{DataAllRootPath}}\SharePoint\OnPrem\Infrastructure'
    LatestCsvFolderPath='{{WorkspaceRootPath}}\Data\Tenants\{{ProfileKey}}\DATA-LAST'
}
if ((Get-InventoryConfigValue 'OutputRoot') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-ALL\SharePoint\OnPrem\Infrastructure') { throw 'Nested infrastructure OutputRoot tokens were not resolved.' }
if ((Get-InventoryConfigValue 'LatestCsvFolderPath') -ne 'C:\SmartM365\Data\Tenants\prod\DATA-LAST') { throw 'Nested infrastructure DATA-LAST tokens were not resolved.' }
$script:EffectiveConfig.OutputRoot = '{{MissingPath}}\SharePoint'
try { $null = Get-InventoryConfigValue 'OutputRoot'; throw 'Missing token was accepted.' }
catch { if ($_.Exception.Message -ne 'Unresolved configuration value: OutputRoot') { throw } }
$root = 'C:\ProgramData\SmartM365\LauncherCache\SharePointInfrastructure\SmartM365'
try { $null = Assert-InventoryPath -Path (Join-Path $root 'Data\Tenants\prod\DATA-ALL') -Name 'OutputRoot'; throw 'Disposable infrastructure output path was accepted.' }
catch { if ($_.Exception.Message -notlike 'OutputRoot resolves inside the disposable launcher cache.*') { throw } }
if ((Assert-InventoryPath -Path 'C:\SmartM365\DATA' -Name 'OutputRoot') -ne 'C:\SmartM365\DATA') { throw 'Persistent infrastructure output path was rejected.' }
$script:EffectiveConfig = [pscustomobject]@{ IncludedWebApplicationUrls=@('https://content-a'); ExcludedWebApplicationUrls=@('https://content-b') }
function Get-SPWebApplication { @([pscustomobject]@{Url='https://content-a/'},[pscustomobject]@{Url='https://content-b/'}) }
$selected = @(Get-ConfiguredWebApplications)
if ($selected.Count -ne 1 -or $selected[0].Url -ne 'https://content-a/') { throw 'Web application filtering failed.' }
$coreManifest = Join-Path $folder '..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module $coreManifest -Force -ErrorAction Stop
$global:SmartM365TenantKey = 'OFFLINE'
$global:SmartM365OrganizationKey = 'OFFLINE'
$global:SmartM365EnvironmentKey = 'TEST'
$global:SmartM365TenantId = '00000000-0000-0000-0000-000000000001'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SPInfra-Test-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $path = Write-RunCsv -Name 'Empty.csv' -Rows @() -Columns @('TenantKey','FarmId','CollectionStatus') -Folder $tempRoot
    if (-not $global:csvGeneratedPaths.Contains($path)) { throw 'Infrastructure run CSV was not registered for execution summary.' }
    $text = [IO.File]::ReadAllText($path)
    if ($text -notmatch 'TenantKey.*;.*FarmId.*;.*CollectionStatus' -or @(Import-Csv $path -Delimiter ';').Count -ne 0) { throw 'Empty infrastructure CSV contract failed.' }
    $buffers = @{ Farms = (New-Object 'System.Collections.Generic.List[object]') }
    $buffers.Farms.Add([pscustomobject]@{ TenantKey='OFFLINE'; FarmId='FARM'; CollectionStatus='Collected' })
    $bufferedPath = Write-RunCsv -Name 'Buffered.csv' -Rows $buffers['Farms'].ToArray() -Columns @('TenantKey','FarmId','CollectionStatus') -Folder $tempRoot
    $bufferedRows = @(Import-Csv -LiteralPath $bufferedPath -Delimiter ';')
    if ($bufferedRows.Count -ne 1 -or $bufferedRows[0].FarmId -ne 'FARM') { throw 'Windows PowerShell 5.1 infrastructure buffer serialization failed.' }
    $latest = Join-Path $tempRoot 'DATA-LAST'
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    foreach ($name in $producer[0].Files) {
        Write-SmartM365CsvAtomically -Data @([pscustomobject]@{TenantKey='OFFLINE';FarmId='FARM'}) -Path (Join-Path $latest $name) -Columns @('TenantKey','FarmId') -Delimiter ';' -NoTenantKey
    }
    $proofPath = Complete-SmartM365SourceReceipt -Status Success -ErrorCount 0
    $proof = Get-Content $proofPath -Raw | ConvertFrom-Json
    if ($proof.Status -ne 'Completed' -or $proof.Files.Count -ne 6) { throw 'Infrastructure receipt completion failed.' }
    $proofHash = (Get-FileHash $proofPath -Algorithm SHA256).Hash
    Start-SmartM365SourceReceipt -ScriptPath $scriptPath -SourceRootPath $latest
    $null = Complete-SmartM365SourceReceipt -Status Failed -ErrorCount 1
    if ((Get-FileHash $proofPath -Algorithm SHA256).Hash -ne $proofHash) { throw 'A failed run changed the preceding infrastructure receipt.' }
    $marker = Join-Path $tempRoot 'Infrastructure-DailySummary.sent'
    $script:sendCount = 0
    function WriteLog { param($Message, $Level) }
    $sendAction = { $script:sendCount++ }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction)) { throw 'First infrastructure email was skipped.' }
    if (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction) { throw 'Same-day infrastructure email was repeated.' }
    if (-not (Invoke-DailySummaryMail -MarkerPath $marker -SendAction $sendAction -Force) -or $script:sendCount -ne 2) { throw 'Forced infrastructure email failed.' }
    $failedMarker = Join-Path $tempRoot 'Infrastructure-Failed.sent'
    try { $null = Invoke-DailySummaryMail -MarkerPath $failedMarker -SendAction { throw 'Simulated SMTP failure' }; throw 'Mail failure was ignored.' }
    catch { if ($_.Exception.Message -ne 'Simulated SMTP failure') { throw } }
    if (Test-Path -LiteralPath $failedMarker) { throw 'Failed infrastructure email wrote a marker.' }
    $script:mailParameters = $null
    function SendEmailHtmlReport { [CmdletBinding()] param($From,$To,$Subject,$BodyHtml,$MailPurpose,$SmtpServer,$SmtpPort,$SendMailMode,$Cc) $script:mailParameters = $PSBoundParameters }
    $script:EffectiveConfig = [pscustomobject]@{ From='sender@example.test'; To=''; ErrorMailTo='recipient@example.test'; SmtpServer='smtp.example.test'; SmtpPort=25; SendMailMode='SMTP'; Cc='' }
    Send-InventorySummaryMail -Subject 'Offline summary' -BodyHtml '<p>Summary</p>'
    if ($script:mailParameters.To -ne 'recipient@example.test' -or $script:mailParameters.From -ne 'sender@example.test' -or $script:mailParameters.SendMailMode -ne 'Graph' -or $script:mailParameters.ContainsKey('Attachments') -or $script:mailParameters.ContainsKey('SmtpServer')) { throw 'Infrastructure Graph mail routing or attachment policy failed.' }
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $true
    $script:graphInstallCount = 0
    $script:graphImportCount = 0
    function Get-Module { [CmdletBinding()] param([switch]$ListAvailable,[string]$Name) if ($script:graphAvailable -and $Name -eq 'Microsoft.Graph.Authentication') { [pscustomobject]@{Name=$Name;Version=[version]'2.0.0'} } }
    function Get-Command { [CmdletBinding()] param([string]$Name) if (($Name -eq 'Install-Module' -and $script:graphInstallerAvailable) -or ($Name -in @('Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest') -and $script:graphAvailable)) { [pscustomobject]@{Name=$Name} } }
    function Install-Module { [CmdletBinding()] param([string]$Name,[string]$Scope,[string]$Repository,[switch]$Force,[switch]$AllowClobber) if ($Name -ne 'Microsoft.Graph.Authentication' -or $Scope -ne 'CurrentUser' -or $Repository -ne 'PSGallery' -or -not $Force -or -not $AllowClobber) { throw 'Unexpected Graph installation parameters.' }; $script:graphInstallCount++; $script:graphAvailable = $true }
    function Import-Module { [CmdletBinding()] param([string]$Name) if ($Name -ne 'Microsoft.Graph.Authentication') { throw 'Unexpected Graph import.' }; $script:graphImportCount++ }
    $originalProtocol = [Net.ServicePointManager]::SecurityProtocol
    Ensure-GraphAuthenticationModule
    Ensure-GraphAuthenticationModule
    if ($script:graphInstallCount -ne 1 -or $script:graphImportCount -ne 2 -or [Net.ServicePointManager]::SecurityProtocol -ne $originalProtocol) { throw 'Infrastructure Graph bootstrap was not idempotent or changed the process TLS policy.' }
    $script:graphAvailable = $false
    $script:graphInstallerAvailable = $false
    try { Ensure-GraphAuthenticationModule; throw 'Missing Graph installer was accepted.' }
    catch { if ($_.Exception.Message -ne 'Install-Module is unavailable.') { throw } }
    Remove-Item Function:\Get-Module,Function:\Get-Command,Function:\Install-Module,Function:\Import-Module
    $script:uploadCalls = New-Object 'System.Collections.Generic.List[string]'
    function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploadCalls.Add($LocalFilePath); if ($LocalFilePath -eq 'failed.csv') { return }; return [pscustomobject]@{ LocalFilePath=$LocalFilePath } }
    $global:EnableSharePointUpload = $false
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv')) -ne 0 -or $script:uploadCalls.Count -ne 0) { throw 'Disabled infrastructure upload was attempted.' }
    $global:EnableSharePointUpload = $true
    if ((Publish-QualifiedCsvUploads -Paths @('run.csv','latest.csv','run.csv','failed.csv')) -ne 1 -or $script:uploadCalls.Count -ne 3) { throw 'Infrastructure upload qualification failed.' }
} finally {
    if ([IO.Path]::GetFullPath($tempRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Output 'PASS: infrastructure parser, nested configuration, persistent paths, filtering, read-only commands, PS5 buffered CSV, receipts, upload policy, daily mail gate and routing.'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA/rrN1iuxVZMDr
# 1W/twTeglGrrsWjn7MENOZ71R8KOtKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCnOU64ocnWvwSCyqrDiQJC
# 59roYA5Ck1PzEbm48CIMcTANBgkqhkiG9w0BAQEFAASCAYAdxDrZwymqGPnyIlwz
# NARkGASeL8Ys5IZnVqVM8hZ/bCYFap20uC4LTjIEumWt40BwX18ByDfJakzeTRWM
# vYphWJFoQeSBZiEqSojXmrNQMqFdml3hZKfT6SmxhUczNKYI3ixcEP1xrfXGWNCt
# o7c3NBLWYXHyJLiKPoaF0r40uKrv+lwumb7mvH985MsWSeDkeKEyu2jJGGRN/JOk
# WCvvbDhoD0yt9WsOxTMwv0F1fqTt5GeMgw+cZmOlqGrU5/jm+1d+xIREcxK7ByLZ
# y9INnvEjvaQ5YqOSEu9yqOow6RzrpWhSSqjt8+70DM0k82VSwdBNbQ9tTnGFm/g4
# 3Ll0IO2ehPmHvd5a7Yu0ttNrus/KrV0cFrnbY59o7heg4z1AjQpkVs21NLZN1OxC
# eAZWgMzC7yOqKhpE4jeJlImPWQkShkRLm9Adn4ZBlj8LZs13vl4VXntYtpO/R3LV
# YlLBF2fc2ZUHC68sIWV6GBnDE0Au/VE2lXkr0+wfZexoanM=
# SIG # End signature block
