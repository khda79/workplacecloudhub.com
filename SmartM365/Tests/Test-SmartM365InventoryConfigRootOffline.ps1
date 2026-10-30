#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests of SmartInventoryConfigRootPath: path mapping, templates and fail-closed reads.
.VERSION
1.1.1
.NOTES
Builds a synthetic SmartM365 tree under TEMP (copies of the JSON transport module, the tenant context and the
account-classification module) and runs the checks in Windows PowerShell 5.1 and PowerShell 7 when both are
available. No real configuration or tenant is touched.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

$inner = @'
param([string]$Root)
$ErrorActionPreference = 'Stop'
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed = $false; try { & $Action | Out-Null } catch { $failed = $true }; Check $failed $Message }
function Get-Long([string]$Path) { (Resolve-Path -LiteralPath $Path).ProviderPath }
$ws = Join-Path (Split-Path $Root -Parent) 'WS'
$code = Join-Path $Root 'SmartInventory'
$globalPath = Join-Path $Root 'Config\SmartM365.global.local.json.txt'
function Set-Global([string]$Value) {
    $document = [ordered]@{ WorkspaceRootPath = $ws; SmartInventoryConfigRootPath = $Value }
    [IO.File]::WriteAllText($globalPath, ($document | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
}
Set-Global ''
. (Join-Path $Root 'Config\SmartM365-TenantContext.ps1')
$foo = Join-Path $code 'Area\SmartM365-Foo.local.json'

Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $foo) 'Empty setting changed the path.'
Check ((Resolve-SmartM365JsonConfigurationPath $foo) -eq "$foo.txt") 'Empty setting changed the resolution.'
Check ((Get-SmartM365JsonTemplateName "$foo.txt") -eq "$foo.template") 'Empty setting changed the template name.'

Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
$target = Join-Path $ws 'SmartInventory\Config\Area\SmartM365-Foo.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $target) 'Mapping to the configuration root is wrong.'
Check (Test-Path -LiteralPath (Split-Path $target -Parent) -PathType Container) 'Target folder was not created.'
Check ((ConvertTo-SmartM365InventoryConfigLocation $target) -eq $target) 'Mapping is not idempotent.'
Check ((ConvertFrom-SmartM365InventoryConfigLocation "$target.txt") -eq "$foo.txt") 'Inverse mapping is wrong.'
Check ((Get-SmartM365JsonTemplateName "$target.txt") -eq "$foo.template") 'Module template is not read from the code folder.'
Check ((Get-SmartM365JsonTemplatePath "$target.txt") -eq "$foo.template") 'Tenant context template is not read from the code folder.'
$config = Read-SmartM365JsonConfig -Path $foo -Required
Check ($config.Setting -eq 1) 'First run did not use the code folder template.'
Check ((Test-Path -LiteralPath "$target.txt") -and -not (Test-Path -LiteralPath "$foo.txt") -and -not (Test-Path -LiteralPath $foo)) 'First run created the file in the wrong folder.'
[IO.File]::WriteAllText("$foo.template", '{"Setting":1,"Added":2}')
$null = Read-SmartM365JsonConfig -Path $foo -Required
Check (((Get-Content -LiteralPath "$target.txt" -Raw | ConvertFrom-Json).Added) -eq 2) 'Template enrichment was not written to the configuration root.'
$bar = Join-Path $code 'Area\SmartM365-Bar.local.json'
[IO.File]::WriteAllText("$bar.txt", '{"A":1}')
Reject { Resolve-SmartM365JsonConfigurationPath $bar } 'A file left in the code folder was read from a second source.'
Remove-Item -LiteralPath "$bar.txt"
$tenantPath = Join-Path $Root 'Config\Tenants\x.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $tenantPath) -eq $tenantPath) 'A path outside SmartInventory was mapped.'
Set-Global ($code + '\Private'); Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A root inside the code folder was accepted.'
Set-Global '{{DataAllRootPath}}\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'An unsupported token was accepted.'
Set-Global 'Relative\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A relative root was accepted.'

# Exercise the real Prepare retention wrapper and real path resolver with a local spy.
Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
$entry = Join-Path $code 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
$script:scriptName = 'SmartM365-WorkplaceEvidence-Prepare'
$script:Tenant = 'test'
$script:product = Join-Path (Split-Path $Root -Parent) 'product'
New-Item -ItemType Directory -Path (Join-Path $script:product 'scripts') -Force | Out-Null
$spy = "param(`$Tenant,`$Mode,`$ConfigurationPath,`$ExpectedPublishedBatchId,`$ExpectedPointerSHA256,[switch]`$NoCompletionBanner)`n[IO.File]::WriteAllText((Join-Path (Split-Path `$PSScriptRoot -Parent) 'received.txt'), (@(`$ConfigurationPath,`$ExpectedPublishedBatchId,`$ExpectedPointerSHA256,[string][bool]`$NoCompletionBanner) -join '|'))`n[pscustomobject]@{Status='Preview';PlanPath='plan';JournalPath='journal'}"
[IO.File]::WriteAllText((Join-Path $script:product 'scripts\Invoke-PreparedSharePointRetention.ps1'),$spy)
$tokens=$null; $errors=$null
$entryAst=[Management.Automation.Language.Parser]::ParseFile($entry,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Prepare source has parse errors.' }
$func=@($entryAst.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-PostTransferRetention'},$true))
if ($func.Count -ne 1) { throw 'Prepare retention wrapper not found.' }
$body=$func[0].Extent.Text.Replace('$PSScriptRoot',("'"+(Split-Path $entry -Parent).Replace("'","''")+"'"))
. ([scriptblock]::Create($body))
function Invoke-PreparedStep { param([scriptblock]$Action) & $Action }
function WriteLog { param($Message,$Level) $script:retentionLog = "$Level $Message" }
$receipt=[pscustomobject]@{BatchId='synthetic-batch';PointerSHA256=('a'*64)}
$sourceConfig=Join-Path (Split-Path $entry -Parent) ($script:scriptName+'.local.json')
$centralConfig=ConvertTo-SmartM365InventoryConfigLocation -Path $sourceConfig
[IO.File]::WriteAllText($centralConfig+'.txt','{}')
Invoke-PostTransferRetention $receipt
$record=Get-Content -LiteralPath (Join-Path $script:product 'received.txt') -Raw
Check ($record -eq (($centralConfig+'.txt')+'|synthetic-batch|'+('a'*64)+'|True')) 'Post-transfer retention did not receive the central config and verified pointer.'
Check ($script:retentionLog -match 'INFO Post-transfer retention: Preview') 'Central config did not reach retention.'
Remove-Item -LiteralPath ($centralConfig+'.txt')
[IO.File]::WriteAllText($sourceConfig+'.txt','{}')
Remove-Item -LiteralPath (Join-Path $script:product 'received.txt')
Invoke-PostTransferRetention $receipt
Check (-not (Test-Path -LiteralPath (Join-Path $script:product 'received.txt')) -and $script:retentionLog -match '^WARNING Batch publication remains verified; retention incomplete:') 'Missing central config must fail closed without invoking retention.'
Set-Global ''
Invoke-PostTransferRetention $receipt
$record=Get-Content -LiteralPath (Join-Path $script:product 'received.txt') -Raw
Check ($record -eq (($sourceConfig+'.txt')+'|synthetic-batch|'+('a'*64)+'|True')) 'Legacy non-central config no longer reaches retention.'

if ($PSVersionTable.PSVersion.Major -ge 7) {
    # The account-classification module requires PowerShell 7.
    Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
    Import-Module (Join-Path $code 'Common\SmartM365.AccountClassification.psd1') -Force
    $classification = Join-Path $ws 'SmartInventory\Config\Config\AccountClassification.local.json.txt'
    New-Item -ItemType Directory -Path (Split-Path $classification -Parent) -Force | Out-Null
    [IO.File]::WriteAllText($classification, '{}')
    Check ((Resolve-SmartM365AccountClassificationPath) -eq (Get-Long $classification)) 'Account classification was not read from the configuration root.'
}
'PASSED {0}' -f $script:passed
'@

function New-Fixture {
    $base = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-InvCfgRoot-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $root = Join-Path $base 'SmartM365'
    $copies = @(
        'SmartInventory\PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1',
        'Modules\SmartM365.Core\SmartM365.JsonTransport.psd1', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psm1',
        'Config\SmartM365-JsonTransport.policy.psd1', 'Config\SmartM365-TenantContext.ps1', 'Config\SmartM365.global.local.json.template',
        'Config\Tenants\tenant.local.json.template', 'Config\Tenants\tenant.mail.local.json.template',
        'SmartInventory\Common\SmartM365.AccountClassification.psd1', 'SmartInventory\Common\SmartM365.AccountClassification.psm1',
        'SmartInventory\Config\AccountClassification.local.json.template')
    foreach ($relative in $copies) {
        $destination = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination $destination
    }
    [IO.File]::WriteAllText((Join-Path $root 'Config\Tenants\test.local.json.txt'), '{"ProfileKey":"test","OrganizationKey":"contoso","EnvironmentKey":"test","TenantKey":"contoso-test"}')
    $area = Join-Path $root 'SmartInventory\Area'
    New-Item -ItemType Directory -Path $area -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.ps1'), '# owner script')
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.local.json.template'), '{"Setting":1}')
    $innerPath = Join-Path $base 'inner.ps1'
    [IO.File]::WriteAllText($innerPath, $inner, (New-Object Text.UTF8Encoding($true)))
    [pscustomobject]@{ Base = $base; Root = $root; Inner = $innerPath }
}

$engines = @()
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (Test-Path -LiteralPath $windowsPowerShell) { $engines += $windowsPowerShell }
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) { $engines += $pwsh.Source }
if (-not $engines.Count) { throw 'No PowerShell engine found.' }
foreach ($engine in $engines) {
    $fixture = New-Fixture
    try {
        $output = & $engine -NoProfile -ExecutionPolicy Bypass -File $fixture.Inner -Root $fixture.Root 2>&1
        $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        if ($LASTEXITCODE -ne 0 -or $text -notmatch 'PASSED \d+') { throw ("{0} failed:{1}{2}" -f $engine, [Environment]::NewLine, $text) }
        Write-Host ('{0}: {1}' -f (Split-Path $engine -Leaf), ([regex]::Match($text, 'PASSED \d+').Value)) -ForegroundColor Green
    }
    finally { Remove-Item -LiteralPath $fixture.Base -Recurse -Force -ErrorAction SilentlyContinue }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAbLVN2UZU9HBEo
# 3ckUwd80mI0IEpQd8kOPC0paT3MpJaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCADKAHoLGUILxyR8Pp2Iwa0
# Vqx2lsCeKvOOqjbJjwdSUDANBgkqhkiG9w0BAQEFAASCAYCnKCutWNXaGtoF8vcD
# Axm6EWfGrn/JTE2pPhBoIaEnvh0E5YaCltyMumV5T4aVf1+JPkLl138ExLnOqUBL
# Q/o17rHw03bti9iAnNVu/BzHeq1oeRtMBKj2MRsjWdbxekaRyOH2AE5I8r/vphRC
# I7GG2XFdNKWXK0qI/ZKjoe5+89/26xLBIwNIct5Ef3AEn3p66FbE4HoRp3yDweNe
# ZtFzCuyE6QHVUsQMfmn2ND13gtHit9/H6q2jS/wAny2ed3gF/0ZfuIo9eoZ64p7g
# f+L7bcEUWzONO5kGc/+ELwhc2lbqOFUBo/PV7gN+MzhnJ1KkGelI91anTLwgWQau
# zOPZhwPvi4CY9EMC3AVD9NRYXKZaqHtZ2T3aWOnoGoi+7AhJqwK6YfuS2O9/W5dY
# hVQyVDs/XacbUnZUcR8/YLFyIsOeqOmKyPy3BV7aMhh9V4sqLVvCVqG/nd24YKPk
# jpTb/GOaej4AvKNdqiipw2Y71bTqS/RPVUPlJrpjPanU7e4=
# SIG # End signature block
