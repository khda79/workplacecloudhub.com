#Requires -Version 7.0
<#
.SYNOPSIS
Checks the read-only Configuration tab of the Orchestrator GUI offline.
.DESCRIPTION
Verifies that every key of the three committed templates has a description and that no
description points to a vanished key, then reads a synthetic Config folder: transport naming,
lock respect without writing, effective values and sources, template differences, sensitive
masking, filtering, the background refresh contract and the absence of values in GUI logs.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$smartRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$folder = Join-Path $smartRoot 'SmartInventory/Orchestrator'
$configModule = (Resolve-Path (Join-Path $folder 'SmartM365.Orchestrator.ConfigView.psm1')).Path
$guiPath = Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
$descriptionsPath = Join-Path $smartRoot 'Config/SmartM365.ConfigDescriptions.json'
Import-Module $configModule -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
$script:Checks = 0
function Assert-Config {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

# 1. Description catalogue against the three committed templates.
$descriptions = Get-Content -LiteralPath $descriptionsPath -Raw -Encoding UTF8 | ConvertFrom-Json
$templates = @{
    'SmartM365.global.local.json' = 'Config/SmartM365.global.local.json.template'
    'Tenants/tenant.local.json' = 'Config/Tenants/tenant.local.json.template'
    'Tenants/tenant.mail.local.json' = 'Config/Tenants/tenant.mail.local.json.template'
}
Assert-Config (((@($descriptions.Files.PSObject.Properties.Name) | Sort-Object) -join ',') -eq ((@($templates.Keys) | Sort-Object) -join ',')) 'Descriptions cover exactly the three configuration files.'
function Test-PathPattern([string]$Pattern, [string]$Path) {
    $a = $Pattern.Split('.'); $b = $Path.Split('.')
    if ($a.Count -ne $b.Count) { return $false }
    for ($i = 0; $i -lt $a.Count; $i++) { if ($a[$i] -ne '*' -and $a[$i] -ine $b[$i]) { return $false } }
    return $true
}
foreach ($fileKey in $templates.Keys) {
    $template = Get-Content -LiteralPath (Join-Path $smartRoot $templates[$fileKey]) -Raw -Encoding UTF8 | ConvertFrom-Json
    $leaves = @(ConvertTo-SmartM365OrchestratorConfigEntries -Document $template | ForEach-Object Path)
    $containers = @($leaves | ForEach-Object { $parts = $_.Split('.'); for ($n = 1; $n -lt $parts.Count; $n++) { $parts[0..($n - 1)] -join '.' } } | Sort-Object -Unique)
    $known = @($leaves) + @($containers)
    $fileDescriptions = $descriptions.Files.$fileKey
    $names = @($fileDescriptions.PSObject.Properties.Name)
    foreach ($leaf in $leaves) {
        $direct = $names -contains $leaf -or @($names | Where-Object { $_.Contains('*') -and (Test-PathPattern $_ $leaf) }).Count -gt 0
        Assert-Config ($direct -and [string](Get-SmartM365OrchestratorConfigDescription -Descriptions $fileDescriptions -Path $leaf)) "Template key without description: $fileKey $leaf"
    }
    foreach ($name in $names) {
        $alive = if ($name.Contains('*')) { @($known | Where-Object { Test-PathPattern $name $_ }).Count -gt 0 -or @($leaves | Where-Object { $template -and $_ -eq ($name -replace '\.\*$', '') }).Count -gt 0 } else { $known -contains $name }
        Assert-Config $alive "Description points to a vanished key: $fileKey $name"
        Assert-Config ([string]$fileDescriptions.$name -match '^[\x20-\x7E]+$' -and ([string]$fileDescriptions.$name).Length -le 140) "Description must be short English ASCII: $fileKey $name"
    }
}
foreach ($pattern in @('Thumb', 'Thumbprint', '*Secret*', '*Password*', '*Token*')) { Assert-Config (@($descriptions.SensitiveKeyPatterns) -contains $pattern) "Sensitive pattern listed: $pattern" }
Assert-Config ((Test-SmartM365OrchestratorConfigSensitive -Patterns @($descriptions.SensitiveKeyPatterns) -Path 'Thumbprint') -and (Test-SmartM365OrchestratorConfigSensitive -Patterns @($descriptions.SensitiveKeyPatterns) -Path 'Graph.ClientSecretValue') -and (Test-SmartM365OrchestratorConfigSensitive -Patterns @($descriptions.SensitiveKeyPatterns) -Path 'TeamsAlertsWebhookUrl') -and -not (Test-SmartM365OrchestratorConfigSensitive -Patterns @($descriptions.SensitiveKeyPatterns) -Path 'OrgDomain')) 'Sensitive key detection.'

# 2. Synthetic configuration folder (fictional Contoso values only).
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiConfig-' + [guid]::NewGuid().ToString('N'))
$config = Join-Path $root 'Config'
$null = New-Item -ItemType Directory -Path (Join-Path $config 'Tenants')
foreach ($relative in @('SmartM365.global.local.json.template', 'Tenants/tenant.local.json.template', 'Tenants/tenant.mail.local.json.template', 'SmartM365.ConfigDescriptions.json')) {
    Copy-Item -LiteralPath (Join-Path $smartRoot ('Config/' + $relative)) -Destination (Join-Path $config $relative)
}
$utf8 = [Text.UTF8Encoding]::new($false)
function Write-Json([string]$Path, $Object) { [IO.File]::WriteAllText($Path, ($Object | ConvertTo-Json -Depth 20), $utf8) }
$globalDoc = Get-Content -LiteralPath (Join-Path $config 'SmartM365.global.local.json.template') -Raw | ConvertFrom-Json
$globalDoc.MailClientName = 'Contoso Global'
$globalDoc.PSObject.Properties.Remove('MailFooterProductName')
Write-Json (Join-Path $config 'SmartM365.global.local.json.txt') $globalDoc
$secret = 'C0FFEE00C0FFEE00C0FFEE00C0FFEE00C0FFEE00'
$tenantDoc = Get-Content -LiteralPath (Join-Path $config 'Tenants/tenant.local.json.template') -Raw | ConvertFrom-Json
$tenantDoc.ProfileKey = 'contoso'; $tenantDoc.Thumbprint = $secret; $tenantDoc.Thumb = $secret; $tenantDoc.From = 'tenant-sender@contoso.example'; $tenantDoc.Cc = ''
$tenantDoc.TeamsAlertsWebhookUrl = 'https://contoso.example/hook?sig=SECRETSIGNATURE123'
$tenantDoc.AdComputerNameCountryPrefixes = [pscustomobject]@{ FRP = 'France' }
$tenantDoc.PSObject.Properties.Remove('RemoteRoutingDomain')
$tenantDoc | Add-Member -NotePropertyName 'LegacyExtraKey' -NotePropertyValue 'extra'
Write-Json (Join-Path $config 'Tenants/contoso.local.json.txt') $tenantDoc
$mailDoc = Get-Content -LiteralPath (Join-Path $config 'Tenants/tenant.mail.local.json.template') -Raw | ConvertFrom-Json
$mailDoc.Scripts.'SmartM365-Licences-Inventory'.From = 'licensing@contoso.example'
$mailDoc.Scripts | Add-Member -NotePropertyName 'SmartM365-Retired-Script' -NotePropertyValue ([pscustomobject]@{ From = '__USE_GLOBAL__'; Cc = '__USE_GLOBAL__' })
Write-Json (Join-Path $config 'Tenants/contoso.mail.local.json') $mailDoc
function Get-Snapshot { @(Get-ChildItem -LiteralPath $root -Recurse -File | Sort-Object FullName | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.FullName, $_.Length, $(if ($_.Name -like '*.transport.lock') { 'lock' } else { (Get-FileHash -LiteralPath $_.FullName).Hash }), $_.LastWriteTimeUtc.Ticks }) }

try {
    $before = Get-Snapshot
    $view = Get-SmartM365OrchestratorConfigurationView -RootPath $root -ProfileKey 'Contoso'
    $after = Get-Snapshot
    Assert-Config (@(Compare-Object $before $after).Count -eq 0) 'Reading writes, renames or creates nothing (no lock or journal file).'
    $files = @{}; foreach ($file in @($view.Files)) { $files[$file.Kind] = $file }
    Assert-Config ($view.ProfileKey -eq 'contoso' -and @($view.Files).Count -eq 3 -and @($view.Files | Where-Object Status -eq 'OK').Count -eq 3) 'Three files read for the requested profile.'
    Assert-Config ($files['Mail'].Path -like '*contoso.mail.local.json' -and $files['Tenant'].Path -like '*contoso.local.json.txt') 'Preferred .json.txt and legacy .json names resolved.'
    function Get-Row([string]$Kind, [string]$Parameter) { @($files[$Kind].Rows | Where-Object Parameter -eq $Parameter)[0] }
    Assert-Config ((Get-Row 'Global' 'MailFooterProductName').Model -eq 'Missing from file' -and $files['Global'].MissingFromFile -eq 1) 'Template key missing from the global file.'
    Assert-Config ((Get-Row 'Tenant' 'LegacyExtraKey').Model -eq 'Not in template' -and (Get-Row 'Tenant' 'RemoteRoutingDomain').Model -eq 'Missing from file') 'Tenant key outside the template and template key missing from the tenant.'
    Assert-Config ((Get-Row 'Tenant' 'AdComputerNameCountryPrefixes.FRP').Model -eq 'OK' -and (Get-Row 'Tenant' 'AdComputerNameCountryPrefixes.FRP').Description -like 'Country mapped*' -and $null -eq (Get-Row 'Tenant' 'AdComputerNameCountryPrefixes')) 'Free map entries belong to the template.'
    Assert-Config ((Get-Row 'Mail' 'Scripts.SmartM365-Retired-Script.From').Model -eq 'Not in template' -and (Get-Row 'Mail' 'Scripts.SmartM365-Retired-Script.From').Description -like 'Sender mailbox*') 'Unknown mail script flagged with the generic description.'
    $clientName = Get-Row 'Tenant' 'MailClientName'
    Assert-Config ($clientName.RawValue -eq '__USE_GLOBAL__' -and $clientName.RawEffective -eq 'Contoso Global' -and $clientName.Source -eq 'global') '__USE_GLOBAL__ resolves from the global file.'
    $footer = Get-Row 'Tenant' 'MailFooterProductName'
    Assert-Config ($footer.RawEffective -eq 'workplacecloudhub.com' -and $footer.Source -eq 'default (global template)') 'Missing global key falls back to the global template default.'
    $url = Get-Row 'Tenant' 'MailPowerBiUrl'
    Assert-Config ($url.RawValue -eq '' -and $url.Source -eq 'global') 'A blank tenant value inherits like __USE_GLOBAL__.'
    $cc = Get-Row 'Tenant' 'Cc'
    Assert-Config ($cc.Source -eq 'default (script)' -and $cc.RawEffective -eq 'null') 'No global value and no template default: script default.'
    Assert-Config ((Get-Row 'Tenant' 'From').Source -eq 'tenant' -and (Get-Row 'Tenant' 'From').RawEffective -eq 'tenant-sender@contoso.example') 'Tenant value is its own source.'
    Assert-Config ((Get-Row 'Mail' 'Scripts.SmartM365-Licences-Inventory.From').Source -eq 'mail' -and (Get-Row 'Mail' 'Scripts.SmartM365-Teams-Inventory.From').Source -eq 'tenant' -and (Get-Row 'Mail' 'Scripts.SmartM365-Teams-Inventory.From').RawEffective -eq 'tenant-sender@contoso.example') 'Mail overrides and inherited tenant sender.'
    Assert-Config ((Get-Row 'Mail' 'Scripts.SmartM365-Teams-Inventory.Cc').Source -eq 'default (script)') 'Inherited Cc without tenant or global value.'
    Assert-Config ((Get-Row 'Tenant' 'EndpointProtectionPolicy.Requirements').RawValue -eq '[]' -and (Get-Row 'Global' 'MailBrandingEnabled').RawValue -eq 'true') 'Arrays and booleans are formatted.'

    # 3. Masking and filtering.
    $tenantRows = @(Select-SmartM365OrchestratorConfigRows -Rows @($files['Tenant'].Rows))
    $hidden = @($tenantRows | Where-Object Masked)
    Assert-Config ($hidden.Count -ge 5 -and @($tenantRows | Where-Object { $_.Value -eq $secret -or $_.Effective -eq $secret }).Count -eq 0 -and @($tenantRows | Where-Object { [string]$_.Value -like '*SECRETSIGNATURE*' }).Count -eq 0) 'Sensitive values hidden by default (thumbprints, webhook, ids).'
    Assert-Config (@($tenantRows | Where-Object Parameter -eq 'OrgDomain')[0].Masked -eq $false) 'Non-sensitive values stay visible.'
    Assert-Config (@(Select-SmartM365OrchestratorConfigRows -Rows @($files['Tenant'].Rows) -Filter 'C0FFEE').Count -eq 0) 'The filter never matches a hidden value.'
    $shown = @(Select-SmartM365OrchestratorConfigRows -Rows @($files['Tenant'].Rows) -ShowSensitive -Filter 'C0FFEE')
    Assert-Config ($shown.Count -eq 2 -and @($shown | Where-Object { $_.Value -eq $secret }).Count -eq 2) 'Show reveals values and allows filtering on them.'
    Assert-Config (@(Select-SmartM365OrchestratorConfigRows -Rows @($files['Tenant'].Rows) -Filter 'EndpointProtectionPolicy.Enabled').Count -eq 1) 'Filter on parameter names.'
    Assert-Config (@(Select-SmartM365OrchestratorConfigRows -Rows @($files['Global'].Rows) -Filter 'DATA-LAST').Count -ge 1) 'Filter on descriptions and visible values.'

    # 4. Transport lock respected, conflicts and missing files reported, still without writing.
    $lockPath = Join-Path $config 'Tenants/contoso.local.json.transport.lock'
    [IO.File]::WriteAllBytes($lockPath, [byte[]]@())
    $holder = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $busyBefore = Get-Snapshot
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $busy = Read-SmartM365OrchestratorConfigFile -Path (Join-Path $config 'Tenants/contoso.local.json') -LockTimeoutSeconds 1
        $watch.Stop()
        Assert-Config ($busy.Status -eq 'Busy' -and $null -eq $busy.Document -and $watch.ElapsedMilliseconds -ge 900) 'A held transport lock is waited for, then reported as Busy.'
        Assert-Config (@(Compare-Object $busyBefore (Get-Snapshot)).Count -eq 0) 'Waiting on the lock writes nothing.'
    }
    finally { $holder.Dispose() }
    $free = Read-SmartM365OrchestratorConfigFile -Path (Join-Path $config 'Tenants/contoso.local.json')
    Assert-Config ($free.Status -eq 'OK' -and (Get-Item -LiteralPath $lockPath).Length -eq 0) 'Released lock: file read, lock file untouched.'
    Remove-Item -LiteralPath $lockPath
    [IO.File]::WriteAllText((Join-Path $config 'Tenants/contoso.local.json'), '{"ProfileKey":"different"}', $utf8)
    $conflict = Read-SmartM365OrchestratorConfigFile -Path (Join-Path $config 'Tenants/contoso.local.json')
    Assert-Config ($conflict.Status -eq 'Error' -and $conflict.Error -like '*Conflicting JSON transport names*') 'Divergent .json and .json.txt copies are reported, not merged.'
    Remove-Item -LiteralPath (Join-Path $config 'Tenants/contoso.local.json')
    $missing = Get-SmartM365OrchestratorConfigurationView -RootPath $root -ProfileKey 'fabrikam'
    Assert-Config (@($missing.Files | Where-Object Kind -ne 'Global' | Where-Object Status -eq 'Missing').Count -eq 2 -and @(@($missing.Files | Where-Object Kind -eq 'Tenant')[0].Rows).Count -eq 0) 'Missing tenant files are reported.'
    Assert-Config (-not (Test-Path -LiteralPath (Join-Path $config 'Tenants/fabrikam.local.json.txt')) -and -not (Test-Path -LiteralPath (Join-Path $config 'Tenants/fabrikam.local.json.transport.lock'))) 'A missing file is never created.'
    $invalid = $null; try { $null = Get-SmartM365OrchestratorConfigurationView -RootPath $root -ProfileKey '..\evil' } catch { $invalid = $_.Exception.Message }
    Assert-Config ($invalid -like 'Invalid tenant profile name*') 'Profile names cannot escape the Tenants folder.'

    # 5. Static read-only guards.
    $moduleText = [IO.File]::ReadAllText($configModule) + [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.ConfigView.Functions.ps1'))
    Assert-Config ($moduleText -notmatch '(?i)Set-Content|Add-Content|Out-File|Remove-Item|New-Item|Move-Item|Copy-Item|WriteAll|Resolve-SmartM365JsonConfigurationPath|Move-SmartM365OwnedJsonFile|Write-SmartM365Json|FileAccess\]::(Read)?Write|FileMode\]::(Create|OpenOrCreate|Append|Truncate)') 'Configuration view code has no write path and avoids the migrating resolver.'
    $guiText = [IO.File]::ReadAllText($guiPath)
    $tabXaml = [regex]::Match($guiText, '(?s)<TabItem Header="Configuration">.*?\n            </TabItem>').Value
    Assert-Config ($tabXaml -and @([regex]::Matches($tabXaml, '<TextBox ')).Count -eq 1 -and $tabXaml -notmatch '(?i)Content="(Save|Apply|Edit|Publish)' -and @([regex]::Matches($tabXaml, '<Button ')).Count -eq 3) 'Tab offers only filter, Show, Refresh and Open folder; no edit control.'
    Assert-Config ($tabXaml -match 'ClipboardCopyMode="None"') 'Grids do not copy values to the clipboard.'

    # 6. GUI functions: masking, temporary Show and no value in logs.
    $tokens = $null; $errors = $null
    $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
    Assert-Config ($errors.Count -eq 0) 'GUI parses.'
    function Get-GuiFunction([string]$Name) { $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true) }
    foreach ($name in @('Update-GuiConfigurationView', 'Set-GuiConfiguration', 'Show-GuiConfigurationSensitive', 'Test-GuiConfigurationTabVisible')) { . ([scriptblock]::Create((Get-GuiFunction $name).Extent.Text)) }
    $script:Logged = [Collections.Generic.List[string]]::new()
    function Write-GuiActivity { param([string]$Message, [string]$Level = 'INFO', [switch]$DisplayOnly) $script:Logged.Add($Message) }
    function New-Fake { [pscustomobject]@{ ItemsSource = $null; Text = ''; Header = ''; Content = '' } }
    $script:Controls = @{ ConfigFilterBox = [pscustomobject]@{ Text = '' }; ConfigShowSensitiveButton = (New-Fake); ConfigSummaryText = (New-Fake) }
    foreach ($prefix in @('ConfigGlobal', 'ConfigTenant', 'ConfigMail')) { foreach ($suffix in @('Grid', 'StatusText', 'Tab')) { $script:Controls[$prefix + $suffix] = New-Fake } }
    $script:ConfigurationView = $null; $script:ConfigurationShowUntil = [datetime]::MinValue
    Set-GuiConfiguration -View $view
    Assert-Config (@($script:Controls.ConfigTenantGrid.ItemsSource | Where-Object { $_.Value -eq $secret }).Count -eq 0 -and $script:Controls.ConfigShowSensitiveButton.Content -eq 'Show') 'GUI masks by default.'
    Show-GuiConfigurationSensitive
    Assert-Config (@($script:Controls.ConfigTenantGrid.ItemsSource | Where-Object { $_.Value -eq $secret }).Count -eq 2 -and $script:Controls.ConfigShowSensitiveButton.Content -like 'Hide (*') 'Show reveals for a limited time.'
    Assert-Config (($script:ConfigurationShowUntil - [datetime]::UtcNow).TotalSeconds -le 30 -and ($script:ConfigurationShowUntil - [datetime]::UtcNow).TotalSeconds -gt 25) 'Show lasts 30 seconds.'
    $script:ConfigurationShowUntil = [datetime]::UtcNow.AddSeconds(-1); Update-GuiConfigurationView
    Assert-Config (@($script:Controls.ConfigTenantGrid.ItemsSource | Where-Object { $_.Value -eq $secret }).Count -eq 0) 'Values are hidden again when the window expires.'
    Set-GuiConfiguration -View $null -ErrorMessage 'Synthetic read failure'
    $allLogged = $script:Logged -join "`n"
    Assert-Config (-not $allLogged.Contains($secret) -and -not $allLogged.Contains('SECRETSIGNATURE') -and -not $allLogged.Contains('tenant-sender@contoso.example') -and $script:Logged.Count -ge 2) 'No configuration value reaches the GUI log.'
    Assert-Config ($guiText -match "elseif \(\[datetime\]::UtcNow -lt \`$script:ConfigurationShowUntil\) \{ \`$script:ConfigurationShowUntil = \[datetime\]::MinValue") 'Leaving the tab hides values again.'

    # 7. Background refresh contract.
    . ([scriptblock]::Create((Get-GuiFunction 'Get-GuiRefreshWork').Extent.Text))
    $shared = Join-Path $root 'Shared'; $null = New-Item -ItemType Directory -Path $shared
    $managementPath = Join-Path $root 'Management.psm1'; $insightsPath = Join-Path $root 'Insights.psm1'
    'function Read-SmartM365OrchestratorJson { param($Path) [pscustomobject]@{ExpectedOrchestratorServers=@()} }; Export-ModuleMember -Function *' | Set-Content -LiteralPath $managementPath
    'function Get-SmartM365OrchestratorOperations { param($SharedDataFolderPath,$ClusterDocument,$MailFolderPath,$MailHours) Start-Sleep -Milliseconds 700; [pscustomobject]@{Servers=@();Running=@();Pending=@();Incidents=@();Mails=@()} }; function Get-SmartM365OrchestratorRecentPipelineRuns { param($SharedDataFolderPath,$Count) @() }; Export-ModuleMember -Function *' | Set-Content -LiteralPath $insightsPath
    function Invoke-ConfigWorker([string]$ProfileName) {
        $context = [pscustomobject]@{ Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' }) }
        $inputData = [pscustomobject]@{ Mode = 'Live'; Root = $shared; MailFolderPath = ''; ManagementModulePath = $managementPath; InsightsModulePath = $insightsPath
            HistoryFrom = (Get-Date).Date; HistoryTo = (Get-Date); HistoryServer = ''; HistoryJob = ''; HistoryStatus = ''; OnlyFailures = $false; LogEventsHours = 0; LogEventsCache = @{}
            IncludeDashboard = $false; ConfigurationProfile = $ProfileName; ConfigRootPath = $root; ConfigViewModulePath = $configModule }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        Assert-Config (Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData) 'Worker started.'
        $watch.Stop(); Assert-Config ($watch.ElapsedMilliseconds -lt 600) "Configuration read must not block the caller ($($watch.ElapsedMilliseconds) ms)."
        $deadline = [datetime]::UtcNow.AddSeconds(30)
        while (-not $context.Worker.Handle.IsCompleted -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 30 }
        $received = Receive-SmartM365OrchestratorGuiWorker -Context $context
        Assert-Config (-not $received.Error -and @($received.Output).Count -eq 1) "Worker result: $($received.Error)"
        return $received.Output[0]
    }
    $withConfig = Invoke-ConfigWorker 'contoso'
    Assert-Config ($null -ne $withConfig.Configuration -and -not $withConfig.ConfigurationError -and @($withConfig.Configuration.Files).Count -eq 3) "Worker reads the configuration: $($withConfig.ConfigurationError)"
    Assert-Config ($null -eq (Invoke-ConfigWorker '').Configuration) 'No configuration read while the tab is hidden.'
    $failed = Invoke-ConfigWorker '..\evil'
    Assert-Config ($null -eq $failed.Configuration -and $failed.ConfigurationError -like 'Invalid tenant profile name*' -and $null -ne $failed.Operations) 'A configuration failure is isolated from live operations.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }

"[{0}] PASS: Orchestrator GUI configuration, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
