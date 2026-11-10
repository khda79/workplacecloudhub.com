# Version: 1.0.0
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-TeamsTest-' + [guid]::NewGuid().ToString('N'))
try {
    $folder = Join-Path $root 'Config/Tenants'
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $script:cfg = [pscustomobject]@{ ProfileKey='test'; SmartM365RootPath=$root }
    function Get-SmartM365EffectiveModuleGlobalConfig { return $script:cfg }
    function Resolve-SmartM365JsonConfigurationPath { param($Path) return $Path }
    function Read-SmartM365JsonDocument { param($Path) return [pscustomobject]@{Document=(Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -ErrorAction Stop)} }
    . (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365-TeamsPrivateConfig.ps1')
    if ($null -ne (Get-SmartM365PrivateTeamsWebhook -Name TeamsWebhookUrl)) { throw 'Missing file should permit legacy fallback.' }
    $file = Join-Path $folder 'test.teams.local.json'
    [IO.File]::WriteAllText($file, '{"TeamsAlertsWebhookUrl":"https://example.invalid/alert","TeamsInfosWebhookUrl":"https://example.invalid/info","TeamsWebhookUrl":"https://example.invalid/default"}')
    foreach ($case in @(@('TeamsAlertsWebhookUrl','alert'),@('TeamsInfosWebhookUrl','info'),@('TeamsWebhookUrl','default'))) {
        if ((Get-SmartM365PrivateTeamsWebhook -Name $case[0]) -ne ('https://example.invalid/' + $case[1])) { throw 'Private route did not resolve.' }
    }
    $script:cfg.ProfileKey = 'other'
    if ($null -ne (Get-SmartM365PrivateTeamsWebhook -Name TeamsWebhookUrl)) { throw 'Profiles must remain isolated.' }
    $script:cfg.ProfileKey = '../test'
    $rejected=$false; try { Get-SmartM365PrivateTeamsWebhook -Name TeamsWebhookUrl } catch { $rejected=$true }
    if (-not $rejected) { throw 'Invalid profile accepted.' }
    $script:cfg.ProfileKey = 'test'
    [IO.File]::WriteAllText($file, '{bad-json')
    $rejected=$false; try { Get-SmartM365PrivateTeamsWebhook -Name TeamsWebhookUrl } catch { $rejected=$true }
    if (-not $rejected) { throw 'Malformed file must fail closed.' }
    function Test-SmartM365InterruptedRun { return $false }
    function Get-ModuleLocalConfig { return @{} }
    function Get-ModuleLocalConfigValue { param($Config,$Name,$DefaultValue) if ($script:legacy.ContainsKey($Name)) { return $script:legacy[$Name] }; return $DefaultValue }
    function WriteLog { param($Message,$Level) }
    function Get-SmartM365WebRequestTimeout { param($Command,$TimeoutSeconds) return @{} }
    function Get-SmartM365NetworkSetting { param($Name) return 5 }
    function Invoke-RestMethod { param($Method,$Uri,$ContentType,$Body) $script:sentTo=$Uri }
    $script:legacy = @{EnableTeamsNotifications=$true;TeamsAlertsWebhookUrl='https://example.invalid/old-alert';TeamsWebhookUrl='https://example.invalid/old-default'}
    $modules = @('../Modules/SmartM365.Core/SmartM365.Core.psm1','../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')
    foreach ($module in $modules) {
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Resolve-Path (Join-Path $PSScriptRoot $module)).Path,[ref]$tokens,[ref]$errors)
        if ($errors.Count) { throw 'Module parsing failed.' }
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Send-SmartM365TeamsNotification'},$true)
        Invoke-Expression $node.Extent.Text
        [IO.File]::WriteAllText($file, '{"TeamsAlertsWebhookUrl":"https://example.invalid/new-alert","TeamsWebhookUrl":"https://example.invalid/new-default"}')
        $script:sentTo=''; $null=Send-SmartM365TeamsNotification -Title t -Message m -Channel Alerts -ThrowOnError
        if ($script:sentTo -ne 'https://example.invalid/new-alert') { throw 'Private channel precedence failed.' }
        $script:sentTo=''; $null=Send-SmartM365TeamsNotification -Title t -Message m -Channel Infos -ThrowOnError
        if ($script:sentTo -ne 'https://example.invalid/new-default') { throw 'Private generic fallback failed.' }
        Remove-Item -LiteralPath $file
        $script:sentTo=''; $null=Send-SmartM365TeamsNotification -Title t -Message m -Channel Alerts -ThrowOnError
        if ($script:sentTo -ne 'https://example.invalid/old-alert') { throw 'Legacy fallback failed.' }
        $script:sentTo=''; $null=Send-SmartM365TeamsNotification -Title t -Message m -WebhookUrl 'https://example.invalid/explicit' -ThrowOnError
        if ($script:sentTo -ne 'https://example.invalid/explicit') { throw 'Explicit URL precedence failed.' }
    }
    Write-Output 'Private Teams configuration checks passed.'
} finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
