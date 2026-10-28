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
1.0.3
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
$globalDoc.PSObject.Properties.Remove('MailClientLogoMaxKB')
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
    foreach ($kind in @('Global', 'Tenant', 'Mail')) {
        $names = @($files[$kind].Rows | ForEach-Object { [string]$_.Parameter })
        $sorted = [string[]]$names.Clone(); [Array]::Sort($sorted, [StringComparer]::OrdinalIgnoreCase)
        Assert-Config ($names.Count -gt 1 -and (($names -join '|') -ceq ($sorted -join '|'))) "$kind rows are sorted alphabetically by parameter."
    }
    function Get-Row([string]$Kind, [string]$Parameter) { @($files[$Kind].Rows | Where-Object Parameter -eq $Parameter)[0] }
    Assert-Config ((Get-Row 'Global' 'MailClientLogoMaxKB').Model -eq 'Missing from file' -and $files['Global'].MissingFromFile -eq 1) 'Template key missing from the global file.'
    Assert-Config ((Get-Row 'Tenant' 'LegacyExtraKey').Model -eq 'Not in template' -and (Get-Row 'Tenant' 'RemoteRoutingDomain').Model -eq 'Missing from file') 'Tenant key outside the template and template key missing from the tenant.'
    Assert-Config ((Get-Row 'Tenant' 'AdComputerNameCountryPrefixes.FRP').Model -eq 'OK' -and (Get-Row 'Tenant' 'AdComputerNameCountryPrefixes.FRP').Description -like 'Country mapped*' -and $null -eq (Get-Row 'Tenant' 'AdComputerNameCountryPrefixes')) 'Free map entries belong to the template.'
    Assert-Config ((Get-Row 'Mail' 'Scripts.SmartM365-Retired-Script.From').Model -eq 'Not in template' -and (Get-Row 'Mail' 'Scripts.SmartM365-Retired-Script.From').Description -like 'Sender mailbox*') 'Unknown mail script flagged with the generic description.'
    $clientName = Get-Row 'Tenant' 'MailClientName'
    Assert-Config ($clientName.RawValue -eq '__USE_GLOBAL__' -and $clientName.RawEffective -eq 'Contoso Global' -and $clientName.Source -eq 'global') '__USE_GLOBAL__ resolves from the global file.'
    $footer = Get-Row 'Tenant' 'MailClientLogoMaxKB'
    Assert-Config ([string]$footer.RawEffective -eq '200' -and $footer.Source -eq 'default (global template)') 'Missing global key falls back to the global template default.'
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
    $visibleRows = @(Select-SmartM365OrchestratorConfigRows -Rows @($files['Tenant'].Rows) -HideNotInTemplate)
    Assert-Config (@($visibleRows | Where-Object Model -eq 'Not in template').Count -eq 0 -and $visibleRows.Count -eq $tenantRows.Count - @($tenantRows | Where-Object Model -eq 'Not in template').Count -and @($visibleRows | Where-Object Model -eq 'Missing from file').Count -ge 1) 'HideNotInTemplate must hide only the keys that are not in the template.'
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
    # The read-only files are the Settings sub-tab; Servers and versions (GUI 1.6.13) keeps its draft controls.
    $tabXaml = [regex]::Match($guiText, '(?s)<TabItem x:Name="ConfigurationSettingsTab" Header="Settings">.*?\n                        </TabItem>').Value
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDEXsIq8O+9qJrW
# 52QxcY5jB7sZNCdfo6U3dkPbOiXdYaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPI6WluTKco3uIyYHFdaUsawdl16/nh7M8y1/ifJPQxeMA0GCSqG
# SIb3DQEBAQUABIIBgE83e2ZAyZqXC9iNGdggWfGmRtdU/zLqF8CfVKZp5ycBlwTI
# wV9BtPBNnpAyu587MMvF5QfTrmdHrWxTWwg7jal4ZiurukS1HRHPKiRzL1ocJAgl
# jP5xNvT6nfIFCBEA60T1lRQY48v7z0QtuyyJEJgO5Ctzhm2aRqwY1lAysD3oXlOY
# VR8R2pKG4CJfri9p7ShPTUR5YrB5MZX0jjPEBmYD03HP2txyjV/8WhsQ+/Y1D5Wn
# 8YMW6WbjQlv0BpsXNgPobiGRY2gMcW115/ICuubSm1nQ0gwGRg2lAw5HHHdIoIjy
# uQ2OxhiIVoNvWBxpvH+aiQsNsibYD0ppHWIWjrUylU4roAyx7+TB/kcYMaD98NWW
# UCU4bwmG3eLizSD/SIhm3g6L/erYU2ACQjuu+DLZzOZvXGeeXP5xgRJGiHd1woYd
# q9SesWwyamHQU3GM+LjcKSj2JDlAaalkOtpJgPWPVNXEnpDI6jft2J9OdwkxlIAB
# rWqR39voWQlLQj94+aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMjU1
# MDVaMC8GCSqGSIb3DQEJBDEiBCDEWnCNUcbXkuUiB2DMJ+v36WglMzoT1QIA0CLN
# HrmQojANBgkqhkiG9w0BAQEFAASCAgCF8IrEKhPpi1TYiYl5V1H5iigbYGl5knBQ
# DOvJHQ2Vrl47NK9+bYlug8Oye8DzcrLhdU+pr9CmN7TFSslcczlZjw/3NIvHNSqZ
# TBhT6kZt/Jscr4o2DLUTHhfMcaQpWJLEECkDmk/yuKFLDKYUlgJ/R1+7z2HravE0
# NjROhgmMcZA2CiePAHhziWUt4t+KM0128ei3AudZs+MMXiGQZbEATD5qwGivPv4x
# W4xPXE7/Fw5GAlj8HVJoNdx5AsxjrtMws9IeUJKBV3Bg8cUyEbZQUXQtL+cySIG0
# UyipBIe+nm3QfJCQ+CXYcbJc6+9GnM5EXDsch/lE3lEIffJtDfsZDiqb5UOa/8Sv
# 9BaDROq89a9ToUbu8vy25uhAdDtwjEjsTLO3ffkJ+SHJJf3NI0y9KSVIGtN3jq20
# JJlys8SUwVebEJsyhkFg8gbjXZt8KEmOyQPuX1hIPROAXfnVmKqRzYno9KVYP0Ca
# kV9T/4+bWmQDKKb/l9IeoLWpeDq+LK+iaX81YocMsWYMS+dqevk1N4YA5qsSd23s
# fg3PRFJOlYvFyvzPzihgbRY1t/Gd3LeDJlLDBqr2F2sB9nmgFEtHy2mHzQPLvIj4
# hMha6fMJkz5KkfGc9YHyn5+qusW/X+w+E4Vc2rzUAUz3caB51cfTv6gwHBe6iaKf
# TyQ6uxPjsg==
# SIG # End signature block
