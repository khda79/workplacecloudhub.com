#Requires -Version 5.1
<# .SYNOPSIS Offline session ownership, mode and cache regression tests. .VERSION 1.0.0 #>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1')
$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message }; $script:checks++ }
function WriteLog { param($Message,$Level) $script:messages += [string]$Message }
function New-PSSession {
    [CmdletBinding()]
    param($ConfigurationName,$ConnectionUri,$Authentication)
    $script:created++
    Check ($ConfigurationName -eq 'Microsoft.Exchange' -and $Authentication -eq 'Kerberos') 'Wrong endpoint or authentication'
    Check ([uri]$ConnectionUri).AbsolutePath.Equals('/PowerShell/') 'Wrong endpoint path'
    if ($script:failure -eq 'Connect') { throw 'Synthetic connection quota failure' }
    [pscustomobject]@{Id=42;State='Opened'}
}
function Import-PSSession {
    [CmdletBinding()]
    param($Session,[string[]]$CommandName,[switch]$AllowClobber,[switch]$DisableNameChecking)
    Check ($Session.Id -eq 42 -and $AllowClobber -and $DisableNameChecking) 'Wrong session import'
    if ($script:failure -eq 'Import') { throw 'Synthetic import denied' }
    $body = ($CommandName | Where-Object { $script:failure -ne 'MissingCommand' -or $_ -ne 'Get-Mailbox' } | ForEach-Object {
        if ($_ -notmatch '^[A-Za-z]+-[A-Za-z]+$') { throw 'Unsafe synthetic command name' }
        'function ' + $_ + ' { [CmdletBinding()] param([bool]$ViewEntireForest) "SyntheticRemoteResult" }'
    }) -join "`n"
    New-Module -Name ('SyntheticExchange' + [guid]::NewGuid().ToString('N')) -ScriptBlock ([scriptblock]::Create($body))
}
function Remove-PSSession { [CmdletBinding()] param($Session) Check ($Session.Id -eq 42) 'Removed unrelated session'; $script:removed++ }
function Get-PSSnapin { [CmdletBinding()] param($Name,[switch]$Registered) if ($Registered) { [pscustomobject]@{Name=$Name} } }
function Add-PSSnapin { [CmdletBinding()] param($Name) $script:snapins++; function global:Get-Mailbox { 'SyntheticLocalResult' } }
function global:Get-Mailbox { 'SyntheticLocalResult' }
function global:Set-ADServerSettings { [CmdletBinding()] param([bool]$ViewEntireForest) }

$collector = Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$errors)
Check ($errors.Count -eq 0) 'Collector parser failed'
$definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Ensure-SmartM365ExchangeScriptScope'},$true)
. ([scriptblock]::Create($definition.Extent.Text))
$script:created=0; $script:removed=0; $script:snapins=0; $script:failure=''; $script:messages=@()
$script:ExchangeOnPremContext=$null; $script:ExchangeConnectionMode='Remote'
$IncludeRemoteMailboxes=$true; $CollectLargeItemStatistics=$false
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('exchange-session-' + [guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixtureRoot
try {
    Check (Ensure-SmartM365ExchangeScriptScope -ViewEntireForest) 'Remote readiness failed'
    Check (Ensure-SmartM365ExchangeScriptScope -ViewEntireForest) 'Repeated readiness failed'
    Check ($script:created -eq 1 -and $script:snapins -eq 0) 'A run must own one remote session and load no snap-in'
    Check ((Get-Mailbox) -eq 'SyntheticRemoteResult') 'Native command shadowed remote command'
    $moduleName=$script:ExchangeOnPremContext.Module.Name
    Check ((Get-Command Get-RemoteMailbox).ModuleName -eq $moduleName) 'Remote mailbox command was not imported'
    Check (($script:messages -join '|') -match 'ExchangeConnectionMode=Remote; state=Ready') 'Ready mode not logged'
    Stop-SmartM365ExchangeOnPremSession $script:ExchangeOnPremContext
    Check ($script:removed -eq 1 -and -not (Get-Module $moduleName)) 'Owned resources leaked'
    Check ((Get-Mailbox) -eq 'SyntheticLocalResult') 'Cleanup removed the caller command'
    Stop-SmartM365ExchangeOnPremSession $script:ExchangeOnPremContext
    Check ($script:removed -eq 1) 'Cleanup is not idempotent'

    foreach ($failure in @('Connect','Import','MissingCommand')) {
        $script:failure=$failure; $script:ExchangeOnPremContext=$null
        $beforeRemoved=$script:removed; $beforeSnapins=$script:snapins; $caught=$null
        try { Ensure-SmartM365ExchangeScriptScope -ViewEntireForest | Out-Null } catch { $caught=$_ }
        Check ($null -ne $caught -and $caught.Exception.Message -match 'no LocalSnapin fallback') ('Remote failure was hidden: ' + $failure)
        Check ($script:snapins -eq $beforeSnapins) 'Failure fell back to snap-in'
        Check ($script:removed -eq ($beforeRemoved + [int]($failure -ne 'Connect'))) 'Failed startup leaked or removed an unrelated session'
    }
    $script:failure=''; $script:ExchangeConnectionMode='LocalSnapin'; $script:ExchangeOnPremContext=$null
    $beforeCreated=$script:created
    Check (Ensure-SmartM365ExchangeScriptScope -ViewEntireForest) 'LocalSnapin readiness changed'
    Check ($script:created -eq $beforeCreated) 'LocalSnapin created a remote session'
    Remove-Item Function:\Get-Mailbox
    Check (Ensure-SmartM365ExchangeScriptScope -ViewEntireForest) 'LocalSnapin loading failed'
    Check ($script:snapins -eq 1 -and (Get-Mailbox) -eq 'SyntheticLocalResult') 'Legacy snap-in branch did not execute'

    $helperPath=Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1'
    $helperText=[IO.File]::ReadAllText($helperPath)
    Check ($helperText -notmatch 'Invoke-Command|Invoke-Expression|Add-PSSnapin|ScriptBlock\s*=') 'Helper sends language or loads local Exchange'
    Check ($ast.Extent.Text -match "-RequireExchangeOnPrem:\(\x24script:ExchangeConnectionMode -ne 'Remote'\)") 'Remote preflight may load a local snap-in'
    Check ($ast.Extent.Text -match "'ExchangeConnectionMode' -DefaultValue 'Remote'") 'Remote must be the default'
    $modeGuard=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(\$script:ExchangeConnectionMode -notin'},$true)
    $script:ExchangeConnectionMode='InvalidMode'; $invalidRejected=$false
    try { & ([scriptblock]::Create($modeGuard.Extent.Text)) } catch { $invalidRejected=$_.Exception.Message -match 'Remote or LocalSnapin' }
    Check $invalidRejected 'Invalid connection mode was accepted'
    $script:ExchangeConnectionMode='LocalSnapin'
    Check ($ast.Extent.Text -match 'Stop-SmartM365ExchangeOnPremSession -State \x24script:ExchangeOnPremContext') 'Collector final cleanup missing'
    $template=Get-Content ($collector -replace '\.ps1$','.local.json.template') -Raw | ConvertFrom-Json
    Check ($template.ExchangeConnectionMode -eq 'Remote') 'Template default missing'

    foreach ($name in @('Start-SmartM365-Exchange-Local-Mailboxes-Inventory.cmd','Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions.cmd','Start-SmartM365-Exchange-Local-Mailboxes-Inventory-OnlyADPermissions.cmd')) {
        $launcher=[IO.File]::ReadAllText((Join-Path $smartRoot ('SmartInventory/Launchers/OnPremises/' + $name)))
        Check ($launcher -match 'robocopy "%SMARTM365_ROOT%\\Modules\\SmartM365.Core" "%LOCAL_SMARTM365_ROOT%\\Modules\\SmartM365.Core" /MIR') ($name + ': shared helper not covered by cache')
        Check (Test-Path -LiteralPath $helperPath -PathType Leaf) ($name + ': missing dot-sourced helper')
        $sourceRoot=Join-Path $fixtureRoot 'source'
        $cacheRoot=Join-Path $fixtureRoot ([IO.Path]::GetFileNameWithoutExtension($name))
        $sourceModule=Join-Path $sourceRoot 'Modules/SmartM365.Core'
        $null=New-Item -ItemType Directory -Path $sourceModule -Force
        Copy-Item -LiteralPath $helperPath -Destination $sourceModule
        $copyLine=($launcher -split '\r?\n' | Where-Object { $_ -like 'robocopy *Modules*' }) -replace '%SMARTM365_ROOT%', $sourceRoot -replace '%LOCAL_SMARTM365_ROOT%', $cacheRoot -replace '%CACHE_LOG%', (Join-Path $fixtureRoot 'copy.log')
        $copyBatch=Join-Path $fixtureRoot 'copy-only.cmd'
        [IO.File]::WriteAllText($copyBatch,('@echo off' + "`r`n" + $copyLine + "`r`nexit /b %ERRORLEVEL%`r`n"),[Text.Encoding]::ASCII)
        & $env:ComSpec /d /c ('"' + $copyBatch + '"') | Out-Null
        Check ($LASTEXITCODE -lt 8) ($name + ': real cache copy failed')
        $cachedHelper=Join-Path $cacheRoot 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1'
        Check ((Get-FileHash $cachedHelper).Hash -eq (Get-FileHash $helperPath).Hash) ($name + ': cached helper bytes changed')
        $loader=@($ast.EndBlock.Statements | Where-Object { $_.Extent.Text -match '^\$exchangeSessionHelperPath =|^if \(-not \(Test-Path -LiteralPath \$exchangeSessionHelperPath|^\. \$exchangeSessionHelperPath' } | ForEach-Object { $_.Extent.Text })
        # The collector guard is nested in its outer lifecycle try block.
        if ($loader.Count -ne 3) {
            $loader=@($ast.FindAll({param($n) ($n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Extent.Text -match '^\$exchangeSessionHelperPath =') -or ($n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(-not \(Test-Path -LiteralPath \$exchangeSessionHelperPath') -or ($n -is [Management.Automation.Language.PipelineAst] -and $n.Extent.Text -eq '. $exchangeSessionHelperPath')},$true) | Sort-Object { $_.Extent.StartOffset } | ForEach-Object { $_.Extent.Text })
        }
        Check ($loader.Count -eq 3) 'Actual helper loader could not be extracted'
        $smartM365LifecycleRoot=$cacheRoot
        & ([scriptblock]::Create(($loader -join "`n")))
        Remove-Item -LiteralPath $cachedHelper
        $missingRejected=$false
        try { & ([scriptblock]::Create(($loader -join "`n"))) } catch { $missingRejected=$_.Exception.Message -match 'complete runtime cache' }
        Check $missingRejected ($name + ': missing dot-sourced cache dependency was accepted')
    }
    Write-Output ('PASS: {0} session/mode/cache checks; {1}; no Exchange connection or collector.' -f $checks,$PSVersionTable.PSVersion)
}
finally {
    Stop-SmartM365ExchangeOnPremSession $script:ExchangeOnPremContext
    Get-Module SyntheticExchange* | Remove-Module -Force
    Remove-Item Function:\Get-Mailbox,Function:\Set-ADServerSettings -ErrorAction SilentlyContinue
    $resolvedFixture=[IO.Path]::GetFullPath($fixtureRoot)
    $tempPrefix=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedFixture.StartsWith($tempPrefix,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedFixture) -notlike 'exchange-session-*') { throw 'Unsafe fixture cleanup path' }
    Remove-Item -LiteralPath $resolvedFixture -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDmkqeNFNQ+9B5d
# XyvVyb/nLaFPmq/lqFL3Fu3zx2hTsKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID0cXxJpNN/3hh7V8Di5vgcT8pW9sQ0tiG45lRKqHSbNMA0GCSqG
# SIb3DQEBAQUABIIBgJK9YaSTXsywrqCquPGhIIlvvsIKN0VZV1tilU0MZW6T15Oc
# 0mUQIpnkPDQlk4jhUQrwLy8PurqaGGE2vCRZQswrxvdb0ofzOScROTK9m2rjXza2
# wGujktatoLZpQdTvU227/YPFmPXL1OfsDCuTO0zpaTcJwDK1s9X8AbCSl9RLyL6t
# qggXJVTdCqm+EKNP+cmqVbkx063/3RxJWeCdkBRQHCDIjeCz8RnHwza0RVzK4m4+
# UCzKFVNr3xlsVOaL+BHAk8OrKMf9KNO9dyH4cIl2zM6ISbKMpid2hX9yIk8DIac0
# +1Elszbka4eN4OrDSobC/GpsxBBtWAdo3QZe3ehvkBRXNEj5xtQbAFScSioGF4jD
# QhhUg0fYgVpY2ZPWGeMWL6zzPv+6KKYCT/Rn3KcUwMnrtiKj4ojEZ31ABJPgvNmB
# e4acMiXBVMhbvTn4Ryj/sWHc5zazEY8dE5quvPRjCunc0g7vzCgQUC33tX2ZGImM
# B46p86OcyxazwdJhcqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxMTQ3
# MDBaMC8GCSqGSIb3DQEJBDEiBCAvmpyIJMpGFEi1NNVjIyFpwkDO4IWiQZP8v+f4
# UM1xTjANBgkqhkiG9w0BAQEFAASCAgBOtg67v8lFaHjpKZDEfSeF9HB3GZzmT7je
# /YQ7Kpo6PenQbYyDgj9Q5brPdB8vfHT0qam9xJ1n4UO9oAzdflFisOFklHGfrXbF
# esim/6gFu2UJnNyZDeHVlU9uF7uUzWsk/tWDqbwT8JSKR6RV//M3Lo147RpKL6wB
# JhkXe7SRT8/MVU+XrQ1lnAytGQwqTJZd8w37T/kxB9o2Yn5QluhbYbfREKQqCklO
# d1FMJhYsyh1d4v7s3eVJgSIgkG7h8WDW1b+NsZ1azmoobqMz4ttScCD97mEvpKpz
# sIHT/ZMdgOGPAenF3rTJM1sJUAFTcufQLi9jS9i9eqPHItqE4igdVrT01MC69+n6
# GNiQSyKggKFZYoUEkEICPCeDCD40btsiBCSTMyHOwFXS0uHOgStuVc/3bbZptWCG
# 4bljGlIhChe496Y6Yihe5U2R2g/UUU14tJnRkNtGZXs+ucdosDFNQfZoFTAtcRFX
# rRHbJ/b1k+NE/e2pJ4YBRoHitM8Ct5E9zBYUqBBIpdEuocOdcSnzg1yGvIAjCKj9
# D/jvPE8pk0PSmt8pzNWgl2nQLNmrcd6ZyE5E5b1PRYG4Z1zuL8wla1/rkenmCI3W
# f6K9BVodJhvOQKQme+TIVcP2UhdSR7QeO0HNU2GB/QG34s1FG04t2WY/mvY6hoMH
# SwVKHHQFtw==
# SIG # End signature block
