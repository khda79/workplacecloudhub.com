#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the SWI validity shown in the orchestrator GUI (Jobs and Planning). Real SWI
contracts, CMDB source registry and jobs template; synthetic DATA-LAST files and receipts only.
.VERSION
1.0.0
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
    Assert-Case ($row.Swi -eq 'Valid' -and [math]::Abs($row.DaysLeft - [math]::Round((336 - 10) / 24, 1)) -le 0.1) "File rule (336 h) is wrong: $($row.Swi) $($row.DaysLeft)"
    Assert-Case ($row.Detail -like '*M365_Security_SecureScore.csv*DATA-LAST file time + 336 h*') 'The limiting input is not explained.'
    Write-DataFile 'M365_Security_SecureScore.csv' 340
    Assert-Case ((Get-Row (Invoke-Validity) 'M365-SecureScore-Inventory').Swi -eq 'Expired - blocks SWI') 'An expired required file does not block SWI.'
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
    Assert-Case ((Get-Row (Invoke-Validity) 'AD-HealthCheck').Swi -eq 'Expired (optional)') 'An expired optional input blocks SWI.'

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
    "[{0}] PASS: {1} offline SWI validity checks. Synthetic DATA-LAST only." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $script:Cases
}
finally {
    $resolved = [IO.Path]::GetFullPath($root)
    if ($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-SwiValidity-Test-*' -and (Test-Path -LiteralPath $resolved)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
