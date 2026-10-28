#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the SWI validity shown in the orchestrator GUI (Jobs and Planning). Real SWI
contracts, CMDB source registry and jobs template; synthetic DATA-LAST files and receipts only.
.VERSION
1.0.1
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$orchestratorRoot = Join-Path $repoRoot 'SmartM365/SmartInventory/Orchestrator'
$swiConfig = Join-Path $repoRoot 'SmartWorkplaceIntelligence/config'
$registryPath = Join-Path $repoRoot 'SmartM365/Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'
$templatePath = Join-Path $repoRoot 'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.local.json.template'
$insights = Import-Module (Join-Path $orchestratorRoot 'SmartM365.Orchestrator.Insights.psm1') -Force -PassThru
$script:Cases = 0
function Assert-Case { param([bool]$Condition, [string]$Message) $script:Cases++; if (-not $Condition) { throw "FAIL: $Message" } }

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SwiValidity-Test-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $root 'Tenant/DATA-LAST'
$now = [datetimeoffset]::UtcNow
$jobs = Get-Content -LiteralPath (Join-Path $orchestratorRoot 'Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
function Write-DataFile {
    param([string]$Name, [double]$AgeHours, [string]$Content = 'Header')
    $path = Join-Path $latest $Name
    [IO.File]::WriteAllText($path, $Content)
    (Get-Item -LiteralPath $path).LastWriteTimeUtc = $now.UtcDateTime.AddHours(-$AgeHours)
}
function Write-Receipt {
    param([string]$Name, [double]$StartAgeHours, [string[]]$Files, [string]$Status = 'Completed', [double]$CompletedAgeHours = -1, [hashtable]$FileStartAge = @{})
    $completed = if ($CompletedAgeHours -ge 0) { $CompletedAgeHours } else { [math]::Max(0, $StartAgeHours - 1) }
    $records = foreach ($file in $Files) {
        $start = if ($FileStartAge.ContainsKey($file)) { $FileStartAge[$file] } else { $StartAgeHours }
        [ordered]@{ File = $file; StartedAtUtc = $now.AddHours(-$start).ToString('o'); CompletedAtUtc = $now.AddHours(-$completed).ToString('o'); Status = 'Success' }
    }
    $document = [ordered]@{ Owner = 'SmartInventory-SourceReceipt'; ContractVersion = '1.2'; Status = $Status; Errors = 0
        StartedAtUtc = $now.AddHours(-$StartAgeHours).ToString('o'); CompletedAtUtc = $now.AddHours(-$completed).ToString('o'); Files = @($records) }
    [IO.File]::WriteAllText((Join-Path $latest $Name), ($document | ConvertTo-Json -Depth 5))
}
function Invoke-Validity {
    param([string]$ConfigPath = '')
    Get-SmartM365OrchestratorSwiValidity -LatestFolderPath $latest -JobsDocument $jobs -SwiConfigRoot $swiConfig -RegistryPath $registryPath -PrepareConfigPath $ConfigPath -PrepareTemplatePath $templatePath -NowUtc $now
}
function Get-Row { param($Result, [string]$Name) @($Result.Rows | Where-Object Name -eq $Name)[0] }

try {
    $null = New-Item -ItemType Directory -Path $latest -Force
    $source = Get-Content -LiteralPath (Join-Path $swiConfig 'prepared-source-contract.json.txt') -Raw | ConvertFrom-Json
    $artifact = Get-Content -LiteralPath (Join-Path $swiConfig 'prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json
    $native = Get-Content -LiteralPath (Join-Path $swiConfig 'native-evidence-contract.json.txt') -Raw | ConvertFrom-Json
    $intune = Get-Content -LiteralPath (Join-Path $swiConfig 'intune-operational-contract.json.txt') -Raw | ConvertFrom-Json
    $registry = Get-Content -LiteralPath $registryPath -Raw | ConvertFrom-Json

    # 1. The producer map covers every SWI input, names real scripts used by jobs, and each listed
    #    script writes (mentions) the file it is mapped to.
    $map = & $insights { $script:SwiFileProducers }
    $registryFiles = @($registry.Producers | ForEach-Object { $_.Files })
    $intuneFiles = @($intune.groups | ForEach-Object { $_.files.file })
    $contractFiles = @(@($source.currentFiles) + @($source.optionalCurrentFiles.file) + @($artifact.currentFiles) + @($artifact.optionalCurrentFiles.file) + @($native.sources.file) + $intuneFiles | Sort-Object -Unique)
    $mappedFiles = @($map.Values | ForEach-Object { $_ })
    $missing = @($contractFiles | Where-Object { $_ -notin $mappedFiles -and $_ -notin $registryFiles -and $_ -notin $intuneFiles })
    Assert-Case ($missing.Count -eq 0) ("SWI inputs without a producer: {0}" -f ($missing -join ', '))
    Assert-Case (@($mappedFiles | Where-Object { $_ -in $registryFiles -or $_ -in $intuneFiles }).Count -eq 0) 'A receipt producer file is also in the manual map.'
    Assert-Case (@($mappedFiles | Group-Object | Where-Object Count -gt 1).Count -eq 0) 'A file is mapped to two producers.'
    $jobScripts = @($jobs.Jobs | ForEach-Object { [IO.Path]::GetFileName([string]$_.ScriptPath) })
    $scriptFiles = @{}
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $repoRoot 'SmartM365/SmartInventory') -Recurse -File -Filter 'SmartM365-*.ps1') { if (-not $scriptFiles.ContainsKey($file.Name)) { $scriptFiles[$file.Name] = $file.FullName } }
    foreach ($script in @($map.Keys) + @($registry.Producers.Script) + @($intune.groups.producer)) {
        Assert-Case ($script -in $jobScripts) "Producer $script is not the script of any job in the template."
        Assert-Case ($scriptFiles.ContainsKey($script)) "Producer script not found: $script"
    }
    foreach ($script in $map.Keys) {
        $text = Get-Content -LiteralPath $scriptFiles[$script] -Raw
        foreach ($file in $map[$script]) {
            $stem = if ($file -like 'SharePoint_OnPrem_*') { 'SharePoint_OnPrem_' } else { $file -replace '\.csv$', '' -replace '\.json\.txt$', '' }
            Assert-Case ($text.Contains($stem)) "$script does not mention $file."
        }
    }

    # 2. Empty DATA-LAST: required inputs block, optional ones do not, jobs without inputs are reported.
    $empty = Invoke-Validity
    Assert-Case ($empty.MaxSourceAgeHours -eq 336 -and $empty.SettingsSource -like 'Prepare template*') 'The Prepare template limit was not used without a local configuration.'
    Assert-Case (@($empty.Unmapped).Count -eq 0) 'Unmapped inputs with the real contracts.'
    Assert-Case ((Get-Row $empty 'M365-SecureScore-Inventory').Swi -eq 'Missing - blocks SWI') 'A missing required file does not block SWI.'
    Assert-Case ((Get-Row $empty 'AD-HealthCheck').Swi -eq 'Missing (optional)') 'A missing optional-only input blocks SWI.'
    Assert-Case ((Get-Row $empty 'Intune-Configurations-Inventory').Swi -eq 'Not collected (optional)') 'An absent Intune group is not optional.'
    Assert-Case ((Get-Row $empty 'M365-ActiveUsers-Inventory').Swi -like 'Missing* - blocks SWI') 'A native source without file or receipt does not block SWI.'
    $unused = @($empty.Rows | Where-Object Swi -eq 'Not used by SWI')
    Assert-Case ($unused.Count -gt 0 -and $null -eq $unused[0].DaysLeft) 'Jobs without SWI input are not reported.'

    # 3. File rule: global limit, per-file override, expiry, future timestamp, earliest expiry wins.
    Write-DataFile 'M365_Security_SecureScore.csv' 10
    $valid = Invoke-Validity
    $row = Get-Row $valid 'M365-SecureScore-Inventory'
    Assert-Case ($row.Swi -eq 'Valid' -and -not $row.Blocks -and [math]::Abs($row.DaysLeft - [math]::Round((336 - 10) / 24, 1)) -le 0.1) "File rule (336 h) is wrong: $($row.Swi) $($row.DaysLeft)"
    Assert-Case (@($empty.Rows | Where-Object { $_.Blocks -ne ($_.Swi -like '* - blocks SWI') }).Count -eq 0) 'Blocks does not match the blocking states.'
    Assert-Case ($row.Detail -like '*M365_Security_SecureScore.csv*DATA-LAST file time + 336 h*') 'The limiting input is not explained.'
    Write-DataFile 'M365_Security_SecureScore.csv' 340
    $expired = Get-Row (Invoke-Validity) 'M365-SecureScore-Inventory'
    Assert-Case ($expired.Swi -eq 'Expired - blocks SWI' -and $expired.Blocks) 'An expired required file does not block SWI.'
    Write-DataFile 'M365_Security_SecureScore.csv' 330
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-SecureScore-Inventory').Swi -eq 'Expires < 24 h') 'A file expiring within a day is not flagged.'
    Write-DataFile 'M365_Security_SecureScore.csv' -1
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-SecureScore-Inventory').Swi -eq 'Future timestamp - blocks SWI') 'A future file time was accepted.'
    Write-DataFile 'M365_Security_SecureScore.csv' 10 ''
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-SecureScore-Inventory').Swi -eq 'Empty - blocks SWI') 'An empty file was accepted.'
    $priceLimit = [int]((Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json).PreparedSourceAgeOverrides.'M365_License_Prices.csv')
    Write-DataFile 'M365_License_Prices.csv' 400
    $row = Get-Row (Invoke-Validity) 'M365-LicensePricing-Inventory'
    Assert-Case ($priceLimit -eq 744 -and $row.Swi -eq 'Valid' -and [math]::Abs($row.DaysLeft - [math]::Round((744 - 400) / 24, 1)) -le 0.1) 'The per-file override (744 h) was not applied.'
    Write-DataFile 'Intune_EndpointSecurity_DefenderAgents.csv' 20
    Write-DataFile 'Intune_EndpointSecurity_FirewallStatus.csv' 200
    $row = Get-Row (Invoke-Validity) 'Intune-EndpointSecurityHealth-Inventory'
    Assert-Case ($row.Detail -like 'Limiting input: Intune_EndpointSecurity_FirewallStatus.csv*' -and [math]::Abs($row.DaysLeft - [math]::Round((336 - 200) / 24, 1)) -le 0.1) 'The earliest required expiry does not decide.'

    # 4. Optional-only job: own limit (48 h for AD_HealthCheck), never blocks SWI.
    Write-DataFile 'AD_HealthCheck.csv' 47.5
    Assert-Case ((Get-Row (Invoke-Validity) 'AD-HealthCheck').Swi -eq 'Expires < 24 h') 'The optional 48 h limit was not applied.'
    Write-DataFile 'AD_HealthCheck.csv' 60
    $optionalRow = Get-Row (Invoke-Validity) 'AD-HealthCheck'
    Assert-Case ($optionalRow.Swi -eq 'Expired (optional)' -and -not $optionalRow.Blocks) 'An expired optional input blocks SWI.'

    # 5. Native receipt rule: start time of the file record, 48 h; DiscoveredApps 240 h with warning at 168 h.
    $activeUsers = @($registry.Producers | Where-Object Script -eq 'SmartM365-ActiveUsers-Inventory.ps1')[0]
    Write-DataFile 'M365_Users_Active.csv' 1
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-ActiveUsers-Inventory').Swi -eq 'Missing receipt - blocks SWI') 'A fresh native file without receipt does not block SWI.'
    Write-Receipt $activeUsers.Receipt 10 @($activeUsers.Files)
    $row = Get-Row (Invoke-Validity) 'M365-ActiveUsers-Inventory'
    Assert-Case ($row.Swi -eq 'Valid' -and [math]::Abs($row.DaysLeft - [math]::Round((48 - 10) / 24, 1)) -le 0.1 -and $row.Detail -like '*collector receipt start + 48 h*') "Native receipt rule is wrong: $($row.Swi) $($row.DaysLeft) $($row.Detail)"
    Write-Receipt $activeUsers.Receipt 10 @($activeUsers.Files) -FileStartAge @{ 'M365_Users_Active.csv' = 50 }
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-ActiveUsers-Inventory').Swi -eq 'Expired - blocks SWI') 'The file record start time is not used.'
    Write-Receipt $activeUsers.Receipt 10 @($activeUsers.Files) -Status 'Failed'
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-ActiveUsers-Inventory').Swi -eq 'Receipt not completed - blocks SWI') 'A failed receipt was accepted.'
    $apps = @($registry.Producers | Where-Object Script -eq 'SmartM365-Intune-DiscoveredApps-Inventory.ps1')[0]
    foreach ($file in $apps.Files) { Write-DataFile $file 1 }
    Write-Receipt $apps.Receipt 200 @($apps.Files)
    $row = Get-Row (Invoke-Validity) 'Intune-DiscoveredApps-Inventory'
    Assert-Case ($row.Swi -eq 'Aging (accepted)' -and [math]::Abs($row.DaysLeft - [math]::Round((240 - 200) / 24, 1)) -le 0.1) "DiscoveredApps freshness group is wrong: $($row.Swi) $($row.DaysLeft)"

    # 6. Intune operational group: complete receipt by completion time; a partial group blocks.
    $group = @($intune.groups | Where-Object name -eq 'EnrollmentFailures')[0]
    foreach ($file in $group.files.file) { Write-DataFile $file 1 }
    Write-Receipt $group.receipt 30 @($group.files.file) -CompletedAgeHours 20
    $row = Get-Row (Invoke-Validity) 'Intune-EnrollmentFailures-Inventory'
    Assert-Case ($row.Swi -eq 'Valid' -and [math]::Abs($row.DaysLeft - [math]::Round((336 - 20) / 24, 1)) -le 0.1) "Intune group rule is wrong: $($row.Swi) $($row.DaysLeft)"
    $managed = @($intune.groups | Where-Object name -eq 'ManagedApps')[0]
    Write-DataFile $managed.files[0].file 1
    Assert-Case ((Get-Row (Invoke-Validity) 'Intune-ManagedApps-Inventory').Swi -eq 'Incomplete group - blocks SWI') 'A partial Intune group was accepted.'

    # 7. Local Prepare configuration wins over the template; shared outputs name every job.
    $configPath = Join-Path $root 'Prepare.local.json'
    [IO.File]::WriteAllText($configPath, (@{ PreparedMaxSourceAgeHours = 100; PreparedSourceAgeOverrides = @{} } | ConvertTo-Json))
    Write-DataFile 'M365_Security_SecureScore.csv' 90
    $configured = Invoke-Validity -ConfigPath $configPath
    Assert-Case ($configured.MaxSourceAgeHours -eq 100 -and $configured.SettingsSource -eq 'Prepare configuration' -and (Get-Row $configured 'M365-SecureScore-Inventory').Swi -eq 'Expires < 24 h') 'The local Prepare configuration was not used.'
    $exo = Get-Row $configured 'EXO-Mailboxes-Inventory-Fast'
    Assert-Case ($exo.Detail -like '*Shared output of SmartM365-EXO-Mailboxes-Inventory.ps1*EXO-Mailboxes-Inventory*') 'A shared producer output is not explained.'

    # 8. DATA-LAST location and the GUI contract.
    Assert-Case ((Get-SmartM365OrchestratorLatestFolderPath -SharedDataFolderPath 'D:\W\Tenants\t\DATA-ALL\Orchestrator') -eq 'D:\W\Tenants\t\DATA-LAST') 'DATA-LAST is not beside DATA-ALL.'
    $refused = $false; try { $null = Get-SmartM365OrchestratorLatestFolderPath -SharedDataFolderPath 'D:\W\Other' } catch { $refused = $true }
    Assert-Case $refused 'An unexpected shared folder layout was accepted.'
    $gui = Get-Content -LiteralPath (Join-Path $orchestratorRoot 'SmartM365-Inventory-Orchestrator-GUI.ps1') -Raw
    Assert-Case (([regex]::Matches($gui, 'Header="SWI validity" Binding="\{Binding Swi\}"')).Count -eq 2 -and ([regex]::Matches($gui, 'Header="SWI days left"')).Count -eq 2) 'Jobs and Planning do not both show the SWI columns.'
    Assert-Case ($gui -match 'Header="Last success" Binding="\{Binding LastSuccess\}" Width="125"/>\s*<DataGridTextColumn Header="SWI validity"') 'The Jobs tab does not show the last success.'
    Assert-Case ($gui -match 'Get-SmartM365OrchestratorSwiValidityForRoot -SharedDataFolderPath \$root') 'The refresh worker does not compute the SWI validity.'
    Assert-Case (([regex]::Matches($gui, '<DataTrigger Binding="\{Binding SwiBlocks\}" Value="True"><Setter Property="Background" Value="#FDE7E9"/>')).Count -eq 2 -and ([regex]::Matches($gui, '<DataTrigger Binding="\{Binding SwiBlocks\}" Value="True"><Setter Property="Foreground" Value="#A4262C"/>')).Count -eq 2) 'Blocking jobs are not red in both grids.'
    Assert-Case (([regex]::Matches($gui, 'SwiBlocks = \[bool\]\$swi\.Blocks')).Count -eq 2) 'The rows of both grids do not carry SwiBlocks.'
    Assert-Case ($gui.Contains('Set-GuiMaintenanceBannerStyle -Level $(if (-not $control.Enabled -and $activationReady) { ''Normal'' }')) 'An inactive maintenance banner is not neutral.'
    "[{0}] PASS: {1} offline SWI validity checks. Synthetic DATA-LAST only." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $script:Cases
}
finally {
    $resolved = [IO.Path]::GetFullPath($root)
    if ($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-SwiValidity-Test-*' -and (Test-Path -LiteralPath $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpESgIHiWLm0AQ
# 3om+DqvZfggMFoqkRvOleaO8o27rtKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPT/XpAOrNv5tvssSefAbjDtzhc929U4lcqbbfEG+2nxMA0GCSqG
# SIb3DQEBAQUABIIBgK9WWh3UGcMYaCJDcSKKu9s1X0hfsJOOOR+Ib3pZ3W9y6psC
# jCMpVs+BeqehUHCjZEYaxWZ6PZGAaMCdfh4z2Rg1bE6/gyi9yZIHwpchf2R0W4uI
# vmqjy0YcEGLikzbsDlUgtb2j8kkU1NTLTmmql8t6oMQEIqCdVAQkK20hzBpbQy6A
# B5HXcqgxU/NJy3i5m3sf3eaJV4VGzhlklsK4Wobn4h9GmXCHRFeGNvP1QnSHXkSq
# G8O1BysDOrhXKEXT7geE7aNh8xPpQ4BLmPkYRvgzs85cMJJZ1sLawE6sSCIds7PO
# gWbVR+4Jm0nzqv4ONlIBCrt2YfO2xdyuyqrQtofw2QzMfFrY1lHy7um95laP2+AW
# EjXB1ZcMyjHfCyRUpBG7SMKDi5/uR7goVdv9HOKYfYv67IXTVgXRAoKVf2XUI9QK
# cdvw38+yTF+X1qlHSaOMg3jhylZXDZ2sJDIFodNUQKeXBWTozXA4CT4Y0tLTLSXd
# QxyOk9KStWRq+LCJXaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNzA4
# MzRaMC8GCSqGSIb3DQEJBDEiBCDnppErRXoGHEWQuVwo/pf+d0ztJnqcl5YxBWLI
# 7r1z/TANBgkqhkiG9w0BAQEFAASCAgCzgtwP9tZsTQb1tdd6lv7VG4ObSZq6s0no
# cxaZLgH9nAZBG/6KIJGzbD8xWxWgAK7nsWlAQ8XZOb5hgQwWEhTsiZFlpHZJd/WF
# y8H0kl1gBn7K4tl5Bj9L+kNrnZRPDt+LU29c97eE61ypz6ugirxDpLPzsUr1NKrS
# NP2d9Bei/hfVV4QX/RyB1Ff/IqyHAj2hpbR9c6I9tsD4p/c2YPY+++Zw+3XaTM/i
# LR3FTr8WIjwSezohxVrVXUk9X+LxJCPy51tjpR1Mt2iiEyCrRLvHHl5Df21revR/
# HAdHCve59XfUdUMOAp25eurLmvEV6f2W90QStmrIBTzc1ZxiA70Wf/UAtBAYMtws
# pSDO/iOtg9MFNhL6iEmA9trnvrLaQHkrCV1iwxXyZlz4g2TTiJB1V4RZVJbuk2gg
# egRsI7R8DINRw/Y9MhT2W1OP3RGXyVNE+QVciujhSJQWsGWfAk/W24AhKxUluPX9
# bxGX8EeXvoAQ3+EVa9uqD4TpLGgdYLhGj8BnE9RpGvFH9zqOnUdR2UmEWKM3nfJi
# M46rh2r88tRAra1ekMfhWcNb2Ju+MZQXvWl0I57b/f3Lm9PO8rcAWZiBa4GlSSaU
# 7F6hLDFcvH+C261kn2bscDvcL0JzI2hGk+AMGbxGLuDef9tk8i2KbLEJALYNK/Ct
# gRnZWp5MeQ==
# SIG # End signature block
