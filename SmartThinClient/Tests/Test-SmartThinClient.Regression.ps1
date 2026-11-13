param([string]$SourcePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartThinClient-Shell.ps1'))
# Version: 1.1.4
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($SourcePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
# Load function declarations only. Never execute the application entry point.
foreach ($fn in $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] }) {
    . ([scriptblock]::Create($fn.Extent.Text))
}
$sourceRoot = Split-Path $SourcePath -Parent
$script:LogPath = $null; $script:CliMode = $false
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ThinClient-tests-' + [guid]::NewGuid())
New-Item -ItemType Directory $testRoot | Out-Null
$script:ScriptRoot = Join-Path $testRoot 'Source'
New-Item -ItemType Directory (Join-Path $script:ScriptRoot 'Profiles') -Force | Out-Null
Copy-Item -Path (Join-Path $sourceRoot 'Profiles\*.json.template') -Destination (Join-Path $script:ScriptRoot 'Profiles')
Copy-Item -LiteralPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Destination $script:ScriptRoot
$script:OutputRoot = $testRoot
$script:AppName = 'Smart ThinClient Shell'; $script:RunId = 'synthetic'
$script:RollbackRoot = $testRoot
$script:EvidenceJsonPath = Join-Path $testRoot 'evidence.json'; $script:EvidenceCsvPath = Join-Path $testRoot 'evidence.csv'
$Profile = 'Citrix'; $TargetUserMode = 'Auto'; $TargetUserName = ''; $DedicatedUserPassword = $null
$script:passed = 0; $script:failed = 0
function Check($Name, [scriptblock]$Body) {
    try { & $Body; $script:passed++; Write-Host "PASS $Name" }
    catch { $script:failed++; Write-Host "FAIL $Name : $_" }
}
function Assert($Condition, $Message) { if (-not $Condition) { throw $Message } }
function Reject([scriptblock]$Body) { $caught = $false; try { & $Body } catch { $caught = $true }; Assert $caught 'Expected rejection' }
# Every system-write primitive used below is replaced; unanticipated mutations fail.
function Start-Process { throw 'Unexpected process launch' }
function Add-Type { throw 'Unexpected UI or native code load' }
function Set-StringRegistryValue { throw 'Unexpected registry write' }
function Set-DwordRegistryValue { throw 'Unexpected registry write' }
function Restore-RegistryValue { throw 'Unexpected registry restore' }
function New-LocalUser { throw 'Unexpected user creation' }
function Set-AssignedAccess { throw 'Unexpected kiosk write' }
function Clear-AssignedAccess { throw 'Synthetic restore failure' }
function Invoke-CimMethod { throw 'Unexpected CIM call' }
function Get-ThinClientCustomShellPolicy { throw 'Unexpected local shell policy access' }
function Set-ThinClientCustomShellPolicy { throw 'Unexpected local shell policy write' }
function Initialize-ThinClientMachineRoot { throw 'Unexpected machine storage provisioning' }
function Protect-ThinClientMachineLauncher { throw 'Unexpected machine launcher ACL change' }
function Remove-ItemProperty { throw 'Unexpected startup registry removal' }
function Export-Evidence { param($Evidence,$Config) }
function Test-IsAdministrator { return $true }
function Get-ThinClientCurrentIdentity { [pscustomobject]@{Name='EXAMPLE\TEST-USER';Sid='S-1-5-21-1-2-3-1001';AccountType='Domain';Authenticated=$true} }
function Get-RegistryValueSnapshot { param($Path,$Name) [pscustomobject]@{Path=$Path;Name=$Name;ValueExists=$false;ValueKind='';Value=$null} }
$audit = [pscustomobject]@{EffectiveProfile='Citrix';CitrixWorkspacePath=$PSHOME;AvdClassicClientPath='';BrowserPath='';TargetUserName='TEST-USER';TargetUserSid='S-1-5-21-1-2-3-1001'}
Check 'Native configuration survives profile templates' {
    $merged = Merge-ProfileConfig -Config ([ordered]@{PreferredAccessMode='Native';UseWebShell=$false}) -EffectiveProfile Citrix
    Assert ($merged.PreferredAccessMode -eq 'Native' -and -not $merged.UseWebShell) 'Profile overrode explicit global configuration'
}
Check 'None restrictions write nothing' { Apply-ShellLimitations -Config ([ordered]@{EnableShellLimitations=$true;ShellRestrictionLevel='None'}) }
Check 'AutoLaunchMode None writes nothing' { Apply-AutoLaunch -Config ([ordered]@{AutoLaunchMode='None'}) -CommandLine 'synthetic' }
Check 'Rollback excludes untouched Winlogon values' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='RunKey'}) -Audit $audit -EffectiveProfile Citrix
    Assert (@($state.Registry | Where-Object {$_.Path -match 'Winlogon'}).Count -eq 0) 'Rollback includes unrelated system values'
}
Check 'Launch uses isolated file' {
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) return $audit }
    function Start-Process { param($FilePath,$ArgumentList,$WindowStyle) }
    $installed = New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $before = (Get-FileHash $installed).Hash
    $result = Invoke-ThinClientLaunchOnly -Config ([ordered]@{WebShellTitle='Different launch-only title'})
    Assert ($result.LauncherPath -ne $installed) 'Launch reused installed file'
    Assert ((Get-FileHash $installed).Hash -eq $before) 'Launch changed installed file'
}
Check 'Native process omits empty argument list and reports errors' {
    $launcher = New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $a = [Management.Automation.Language.Parser]::ParseFile($launcher,[ref]$tokens,[ref]$errors)
    Assert ($errors.Count -eq 0) 'Generated script parse error'
    $f = $a.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-WorkspaceProcess'},$true)
    . ([scriptblock]::Create($f.Extent.Text))
    function Start-Process { param($FilePath,[ValidateNotNullOrEmpty()][string[]]$ArgumentList) if ($PSBoundParameters.ContainsKey('ArgumentList')) { throw 'Empty args passed' } }
    Assert (Start-WorkspaceProcess -FilePath $PSHOME -Arguments '') 'No native success result'
}
Check 'Invalid access mode rejected before launcher generation' {
    Reject { New-LauncherScript -Config ([ordered]@{PreferredAccessMode="Web'; throw 'injected"}) -EffectiveProfile Citrix -Audit $audit }
}
Check 'All four generated profiles parse with quoted configuration' {
    foreach ($provider in @('Citrix','AVD','WebOnly','Hybrid')) {
        $launcher = New-LauncherScript -Config ([ordered]@{WebShellTitle="Example's title";CitrixWebUrl='https://example.invalid/?a=1&b=2'}) -EffectiveProfile $provider -Audit $audit
        [void][Management.Automation.Language.Parser]::ParseFile($launcher,[ref]$tokens,[ref]$errors)
        Assert ($errors.Count -eq 0) "Parse error in $provider"
    }
}
Check 'JSON string false cannot enable Apply' { Reject { Assert-ConfigValues -Config ([ordered]@{AllowApply='false'}) } }
Check 'CIM error code is not success' { Reject { Assert-CimSuccess -Response ([pscustomobject]@{ReturnValue=5}) } }
Check 'CIM success accepted' { Assert-CimSuccess -Response ([pscustomobject]@{ReturnValue=0}) }
Check 'Rollback captures existing launcher bytes' {
    $path = Join-Path $testRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
    [IO.File]::WriteAllBytes($path,[byte[]]@(1,2,3,4))
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    Assert ($state.LauncherExisted -and $state.LauncherContent -eq 'AQIDBA==') 'Previous launcher missing'
    Assert (@($state.Registry).Count -eq 0) 'No-write apply must not snapshot unrelated registry'
}
Check 'Foreign rollback rejected before any write' {
    $state = New-RollbackState -Config ([ordered]@{}) -Audit $audit -EffectiveProfile Citrix
    $state.ComputerName = 'OTHER-SYNTHETIC-COMPUTER'
    Reject { Assert-RollbackState -State ([pscustomobject]$state) }
}
Check 'Legacy rollback rejected before any write' { Reject { Assert-RollbackState -State ([pscustomobject]@{AppName=$script:AppName}) } }
Check 'Malformed late snapshot rejected before any restore' {
    $state = New-RollbackState -Config ([ordered]@{}) -Audit $audit -EffectiveProfile Citrix
    $state.Registry += [pscustomobject]@{Path='HKLM:\Unrelated';Name='Bad'}
    $RollbackPath = Join-Path $testRoot 'malformed.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    $ConfirmRestore = 'RESTORE'
    Reject { Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true}) }
}
Check 'Restore preserves original launcher bytes' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    $RollbackPath = Join-Path $testRoot 'valid.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    [IO.File]::WriteAllText($state.LauncherPath,'Changed')
    $ConfirmRestore = 'RESTORE'
    $result = Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true})
    Assert ($result.Status -eq 'Restored') 'Missing restored result'
    Assert ([Convert]::ToBase64String([IO.File]::ReadAllBytes($state.LauncherPath)) -eq $state.LauncherContent) 'Original bytes not restored'
}
Check 'Assigned Access restore failure propagates' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None';EnableAssignedAccess=$true}) -Audit $audit -EffectiveProfile Citrix
    $RollbackPath = Join-Path $testRoot 'kiosk-failure.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content $RollbackPath
    $ConfirmRestore = 'RESTORE'
    Reject { Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true}) }
}
Check 'Missing confirmation rejects apply' { $ConfirmApply=''; Reject { Assert-ApplyAllowed -Config ([ordered]@{AllowApply=$true}) } }
Check 'Dedicated user missing password rejected before mutation' {
    $target = [pscustomobject]@{Mode='DedicatedUser';LocalUserExists=$false}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{CreateDedicatedLocalUser=$true}) -Audit $audit -TargetUser $target }
}
Check 'Conflicting kiosk modes rejected' {
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;EnableShellLauncher=$true}) -Audit $audit -TargetUser $null }
}
Check 'Existing Assigned Access is refused before writes' {
    function Get-LocalGroupMember { param($SID) }
    function Get-AssignedAccess { [pscustomobject]@{UserName='OTHER-USER'} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    $kioskAudit=[pscustomobject]@{AssignedAccessCmdletAvailable=$true}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;AssignedAccessAppUserModelId='SYNTHETIC!App'}) -Audit $kioskAudit -TargetUser $target }
}
Check 'Empty Assigned Access preflight passes without writes' {
    function Get-LocalGroupMember { param($SID) }
    function Get-AssignedAccess { }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Assert-ApplyPreflight -Config ([ordered]@{EnableAssignedAccess=$true;AssignedAccessAppUserModelId='SYNTHETIC!App'}) -Audit ([pscustomobject]@{AssignedAccessCmdletAvailable=$true}) -TargetUser $target
}
Check 'Disabled Shell Launcher is refused' {
    function Get-LocalGroupMember { param($SID) }
    function Invoke-CimMethod { param($Namespace,$ClassName,$MethodName) [pscustomobject]@{ReturnValue=0;Enabled=$false} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableShellLauncher=$true}) -Audit ([pscustomobject]@{ShellLauncherClassAvailable=$true}) -TargetUser $target }
}
Check 'Existing custom shell is refused' {
    function Get-LocalGroupMember { param($SID) }
    function Invoke-CimMethod { param($Namespace,$ClassName,$MethodName) [pscustomobject]@{ReturnValue=0;Enabled=$true} }
    function Get-CimInstance { param($Namespace,$ClassName) [pscustomobject]@{Sid=$audit.TargetUserSid} }
    $target=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid}
    Reject { Assert-ApplyPreflight -Config ([ordered]@{EnableShellLauncher=$true}) -Audit ([pscustomobject]@{ShellLauncherClassAvailable=$true}) -TargetUser $target }
}
Check 'Apply saves rollback before account mutation and respects None' {
    $script:sequence = @()
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) $audit }
    function Get-TargetUserState { param($Config) [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true} }
    function Save-RollbackState { param($State) $script:sequence += 'save'; 'synthetic-rollback.json' }
    function New-DedicatedUserIfNeeded { param($Config,$TargetUser) $script:sequence += 'account'; $false }
    $ConfirmApply='APPLY'
    $result=Invoke-ThinClientApply -Config ([ordered]@{AllowApply=$true;AutoLaunchMode='None'})
    Assert ($result.Status -eq 'Applied') 'Missing Applied result'
    Assert (($script:sequence -join ',') -eq 'save,account,save') 'Account attempted before rollback'
}
Check 'Hybrid loads provider profile URLs and routes all three choices' {
    $paths = @()
    try {
        foreach ($provider in @('Citrix','AVD','WebOnly')) {
            $path = Join-Path $script:ScriptRoot ('Profiles\{0}.profile.json' -f $provider)
            $paths += $path
            $key = @{Citrix='CitrixWebUrl';AVD='AvdWebUrl';WebOnly='WebOnlyUrl'}[$provider]
            @{ $key = ('https://{0}.example.invalid' -f $provider.ToLowerInvariant()) } | ConvertTo-Json | Set-Content -LiteralPath $path
        }
        $config = Merge-ProfileConfig -Config ([ordered]@{UseWebShell=$false;HybridSelectionAtStartup=$false}) -EffectiveProfile Hybrid
        $urls = Get-WorkspaceWebUrls -Config $config -EffectiveProfile Hybrid
        foreach ($provider in @('Citrix','AVD','WebOnly')) {
            $expected = 'https://{0}.example.invalid' -f $provider.ToLowerInvariant()
            Assert ($urls[$provider] -eq $expected) "Audit URL missing for $provider"
            $config.HybridDefaultProvider = $provider
            $launcher = New-LauncherScript -Config $config -EffectiveProfile Hybrid -Audit $audit
            $openedUrls = New-Object 'System.Collections.Generic.List[string]'
            function Start-Process { param($FilePath,$ArgumentList) $openedUrls.Add($FilePath) }
            & $launcher
            Assert ($openedUrls.Count -eq 1 -and $openedUrls[0] -eq $expected) "Hybrid routed $provider to the wrong URL"
        }
    }
    finally { foreach ($path in $paths) { Remove-Item -LiteralPath $path -Force } }
}
Check 'Hybrid retains global URLs when profile values inherit' {
    $urls = Get-WorkspaceWebUrls -Config ([ordered]@{CitrixWebUrl='https://global.example.invalid';AvdWebUrl='https://avd.example.invalid';WebOnlyUrl='https://web.example.invalid'}) -EffectiveProfile Hybrid
    Assert ($urls.Citrix -eq 'https://global.example.invalid' -and $urls.AVD -eq 'https://avd.example.invalid' -and $urls.WebOnly -eq 'https://web.example.invalid') 'Global URLs lost'
}
Check 'Hybrid resolves legacy provider URL aliases' {
    function Get-ProfileConfig {
        param($ProfileName)
        switch ($ProfileName) {
            Citrix { [pscustomobject]@{StoreUrl='https://store.example.invalid'} }
            AVD { [pscustomobject]@{FeedUrl='https://feed.example.invalid'} }
            WebOnly { [pscustomobject]@{FallbackWebUrl='https://fallback.example.invalid'} }
        }
    }
    $urls = Get-WorkspaceWebUrls -Config ([ordered]@{}) -EffectiveProfile Hybrid
    Assert ($urls.Citrix -eq 'https://store.example.invalid' -and $urls.AVD -eq 'https://feed.example.invalid' -and $urls.WebOnly -eq 'https://fallback.example.invalid') 'Legacy aliases lost'
}
Check 'Modern Hybrid default wins over legacy Citrix preference' {
    $launcher = New-LauncherScript -Config ([ordered]@{UseWebShell=$false;HybridSelectionAtStartup=$false;HybridDefaultProvider='AVD';WebShellPreferredProvider='Citrix';AvdWebUrl='https://avd.example.invalid'}) -EffectiveProfile Hybrid -Audit $audit
    $openedUrls = New-Object 'System.Collections.Generic.List[string]'
    function Start-Process { param($FilePath,$ArgumentList) $openedUrls.Add($FilePath) }
    & $launcher
    Assert ($openedUrls.Count -eq 1 -and $openedUrls[0] -eq 'https://avd.example.invalid') 'Legacy preference overrode modern default'
}
Check 'Resolve config accepts legacy default only when modern key is absent' {
    $ConfigPath = Join-Path $testRoot 'legacy-config.json'
    '{"WebShellPreferredProvider":"AVD"}' | Set-Content -LiteralPath $ConfigPath
    $script:OutputRootOverride = ''; $TargetUserMode = 'Auto'; $TargetUserName = ''
    $config = Resolve-Config
    Assert ($config.HybridDefaultProvider -eq 'AVD') 'Legacy-only runtime lost preference'
    '{"WebShellPreferredProvider":"Citrix","HybridDefaultProvider":"WebOnly"}' | Set-Content -LiteralPath $ConfigPath
    $config = Resolve-Config
    Assert ($config.HybridDefaultProvider -eq 'WebOnly') 'Modern runtime default lost precedence'
}
Check 'External web readiness accepts configured portals without native clients' {
    $result = [pscustomobject]@{PreferredAccessMode='Web';UseWebShell=$false;BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true}
    $statuses = @(Get-ThinClientCliStatuses -Result $result)
    Assert (@($statuses | Where-Object {$_.Name -in @('Citrix Web','AVD Web') -and $_.State -eq 'READY'}).Count -eq 2) 'External web readiness rejected'
}
Check 'External web without browser is not ready' {
    $result = [pscustomobject]@{PreferredAccessMode='Web';UseWebShell=$false;BrowserInstalled=$false;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true}
    $statuses = @(Get-ThinClientCliStatuses -Result $result)
    Assert (@($statuses | Where-Object {$_.Name -in @('Citrix Web','AVD Web','Citrix','AVD') -and $_.State -eq 'READY'}).Count -eq 0) 'Missing browser marked ready'
}
Check 'Strict preflight rejects unloaded target hive before apply writes' {
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) $audit }
    function Get-TargetUserState { param($Config) [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Sid=$audit.TargetUserSid} }
    function Test-Path { param($LiteralPath) $false }
    function Save-RollbackState { throw 'Unexpected rollback write' }
    $ConfirmApply = 'APPLY'
    $caught = ''
    try { Invoke-ThinClientApply -Config ([ordered]@{AllowApply=$true;EnableShellLimitations=$true;ShellRestrictionLevel='Strict'}) } catch { $caught = $_.Exception.Message }
    Assert ($caught -like '*target user registry hive*loaded*') 'Apply did not stop at hive preflight'
}
Check 'Strict preflight rejects new dedicated user before creation' {
    $DedicatedUserPassword = New-Object Security.SecureString
    $caught = ''
    try { Assert-ApplyPreflight -Config ([ordered]@{CreateDedicatedLocalUser=$true;EnableShellLimitations=$true;ShellRestrictionLevel='Strict'}) -Audit $audit -TargetUser ([pscustomobject]@{Mode='DedicatedUser';LocalUserExists=$false;Sid=''}) } catch { $caught = $_.Exception.Message }
    Assert ($caught -like '*existing local target user*') 'New user accepted for Strict'
}
Check 'Strict captures and writes CMD policy only in target hive' {
    function Test-Path { param($LiteralPath) if ($LiteralPath -like 'Registry::*') { return $true }; Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath }
    $config = [ordered]@{AutoLaunchMode='None';EnableShellLimitations=$true;ShellRestrictionLevel='Strict'}
    $target = [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$true;Sid=$audit.TargetUserSid}
    Assert-ApplyPreflight -Config $config -Audit $audit -TargetUser $target
    $state = New-RollbackState -Config $config -Audit $audit -EffectiveProfile Citrix
    $cmd = @($state.Registry | Where-Object {$_.Name -eq 'DisableCMD'})
    $expected = 'Registry::HKEY_USERS\S-1-5-21-1-2-3-1001\Software\Policies\Microsoft\Windows\System'
    Assert ($cmd.Count -eq 1 -and $cmd[0].Path -eq $expected) 'Wrong CMD snapshot hive'
    $script:cmdPath = ''
    function Set-DwordRegistryValue { param($Path,$Name,$Value) if ($Name -eq 'DisableCMD') { $script:cmdPath = $Path } }
    Apply-ShellLimitations -Config $config -Audit $audit
    Assert ($script:cmdPath -eq $expected) 'Wrong CMD write hive'
    Assert-RollbackState -State ([pscustomobject]$state)
}
Check 'Restore rejects a snapshot for another user SID' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    $state.Registry = @([pscustomobject]@{Path='Registry::HKEY_USERS\S-1-5-21-1-2-3-9999\Software\Policies\Microsoft\Windows\System';Name='DisableCMD';ValueExists=$false;ValueKind='';Value=$null})
    Reject { Assert-RollbackState -State ([pscustomobject]$state) }
}
Check 'Restore refuses unloaded user hive before any registry restore' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='RunKey'}) -Audit $audit -EffectiveProfile Citrix
    $state.Registry += [pscustomobject]@{Path=(Get-TargetUserSystemPolicyPath -Sid $audit.TargetUserSid);Name='DisableCMD';ValueExists=$false;ValueKind='';Value=$null}
    $RollbackPath = Join-Path $testRoot 'unloaded-hive.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $RollbackPath
    function Test-Path { param($LiteralPath) if ($LiteralPath -like 'Registry::*') { return $false }; Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath }
    $ConfirmRestore = 'RESTORE'
    $caught = ''
    try { Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true}) } catch { $caught = $_.Exception.Message }
    Assert ($caught -like '*original target user registry hive*loaded*') 'Restore reached registry writes before checking hive'
}
Check 'Existing schema 2 machine CMD snapshot remains restorable' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    $state.Registry = @([pscustomobject]@{Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\System';Name='DisableCMD';ValueExists=$false;ValueKind='';Value=$null})
    Assert-RollbackState -State ([pscustomobject]$state)
}
Check 'CSV false preserves JSON and does not create or alter CSV' {
    $fn = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Export-Evidence' }
    . ([scriptblock]::Create($fn.Extent.Text))
    $script:EvidenceJsonPath = Join-Path $testRoot 'csv-disabled.json'
    $script:EvidenceCsvPath = Join-Path $testRoot 'csv-disabled.csv'
    $evidence = [pscustomobject]@{Status='Synthetic'}
    Export-Evidence -Evidence $evidence -Config ([ordered]@{WriteCsvEvidence=$false})
    Assert ((Test-Path -LiteralPath $script:EvidenceJsonPath) -and -not (Test-Path -LiteralPath $script:EvidenceCsvPath)) 'CSV false ignored'
    Set-Content -LiteralPath $script:EvidenceCsvPath -Value 'Existing evidence'
    $hash = (Get-FileHash -LiteralPath $script:EvidenceCsvPath).Hash
    Export-Evidence -Evidence $evidence -Config ([ordered]@{WriteCsvEvidence=$false})
    Assert ((Get-FileHash -LiteralPath $script:EvidenceCsvPath).Hash -eq $hash) 'Disabled CSV was altered'
}
Check 'CSV remains enabled by default and with explicit true' {
    $fn = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Export-Evidence' }
    . ([scriptblock]::Create($fn.Extent.Text))
    foreach ($config in @([ordered]@{}, [ordered]@{WriteCsvEvidence=$true})) {
        Export-Evidence -Evidence ([pscustomobject]@{Status='Synthetic'}) -Config $config
        Assert ((Import-Csv -LiteralPath $script:EvidenceCsvPath).Status -eq 'Synthetic') 'CSV default/true failed'
    }
}
Check 'CSV boolean rejects a string false' { Reject { Assert-ConfigValues -Config ([ordered]@{WriteCsvEvidence='false'}) } }
Check 'Loaded target hive restoration uses only its recorded CMD snapshot' {
    $state = New-RollbackState -Config ([ordered]@{AutoLaunchMode='None'}) -Audit $audit -EffectiveProfile Citrix
    $expected = Get-TargetUserSystemPolicyPath -Sid $audit.TargetUserSid
    $state.Registry = @([pscustomobject]@{Path=$expected;Name='DisableCMD';ValueExists=$true;ValueKind='DWord';Value=2})
    $RollbackPath = Join-Path $testRoot 'loaded-hive.json'
    [pscustomobject]$state | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $RollbackPath
    function Test-Path { param($LiteralPath) if ($LiteralPath -like 'Registry::*') { return $true }; Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath }
    $restoredSnapshots = New-Object 'System.Collections.Generic.List[object]'
    function Restore-RegistryValue { param($Snapshot) $restoredSnapshots.Add($Snapshot) }
    $ConfirmRestore = 'RESTORE'
    $result = Invoke-ThinClientRestore -Config ([ordered]@{AllowRestore=$true})
    Assert ($result.Status -eq 'Restored' -and $restoredSnapshots.Count -eq 1 -and $restoredSnapshots[0].Path -eq $expected -and $restoredSnapshots[0].Value -eq 2) 'Target CMD state not restored'
}
Check 'Audit uses Hybrid profile URLs and honors CSV false through the export call' {
    $fn = $ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Export-Evidence' }
    . ([scriptblock]::Create($fn.Extent.Text))
    function Get-CitrixWorkspaceSignal { param($Config) [pscustomobject]@{Installed=$false;Path=''} }
    function Get-AvdClientSignal { param($Config) [pscustomobject]@{Installed=$false;ClassicClientPath='';WindowsAppPackage='';WindowsAppVersion=''} }
    function Get-BrowserSignal { param($Config) [pscustomobject]@{Installed=$true;Path='synthetic-browser'} }
    function Get-KioskCapability { [pscustomobject]@{EditionId='Synthetic';AssignedAccessCmdletAvailable=$false;ShellLauncherClassAvailable=$false;SupportsProgressiveLockdown=$false} }
    function Get-TargetUserState { param($Config) [pscustomobject]@{Mode='ExistingUser';UserName='SYNTHETIC';LocalUserExists=$true;Enabled=$true;Sid=$audit.TargetUserSid} }
    function Get-CimInstance { param($ClassName) [pscustomobject]@{Caption='Synthetic OS';Version='0.0';BuildNumber='0'} }
    function Get-ProfileConfig { param($ProfileName) if ($ProfileName -eq 'Citrix') { [pscustomobject]@{CitrixWebUrl='https://citrix.example.invalid'} } }
    $script:EvidenceJsonPath = Join-Path $testRoot 'audit-no-csv.json'
    $script:EvidenceCsvPath = Join-Path $testRoot 'audit-no-csv.csv'
    $result = Invoke-ThinClientAudit -Config ([ordered]@{WriteCsvEvidence=$false;UseWebShell=$false}) -RequestedProfile Hybrid -RequestedAction Preview
    Assert ($result.CitrixWebUrlConfigured -and (Test-Path -LiteralPath $script:EvidenceJsonPath) -and -not (Test-Path -LiteralPath $script:EvidenceCsvPath)) 'Audit/export ignored effective configuration'
}
Check 'Fresh runtime inherits the Citrix and AVD demo defaults' {
    $previousRoot = $script:ScriptRoot
    try {
        $script:ScriptRoot = Join-Path $testRoot 'DemoFreshSource'
        New-Item -ItemType Directory (Join-Path $script:ScriptRoot 'Profiles') -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Destination $script:ScriptRoot
        Copy-Item -Path (Join-Path $sourceRoot 'Profiles\*.json.template') -Destination (Join-Path $script:ScriptRoot 'Profiles')
        $ConfigPath = ''; $script:OutputRootOverride = ''
        $config = Resolve-Config
        $runtime = Read-JsonFile -Path (Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.json')
        Assert ($runtime.CitrixWebUrl -eq 'https://workplacecloudhub.com/demo/smartthinclient/citrix/' -and $runtime.AvdWebUrl -eq 'https://workplacecloudhub.com/demo/smartthinclient/avd/') 'New runtime lacks demo defaults'
        foreach ($profileName in @('Citrix','AVD','Hybrid')) {
            $merged = Merge-ProfileConfig -Config $config -EffectiveProfile $profileName
            $urls = Get-WorkspaceWebUrls -Config $merged -EffectiveProfile $profileName
            Assert ($urls.Citrix -eq $runtime.CitrixWebUrl -and $urls.AVD -eq $runtime.AvdWebUrl) "Demo defaults lost in $profileName"
        }
        Assert ($config.WebOnlyUrl -eq 'https://workplacecloudhub.com/' -and $runtime.WebOnlyUrl -eq $config.WebOnlyUrl) 'WebOnly default missing'
    }
    finally { $script:ScriptRoot = $previousRoot }
}
Check 'Built-in demo defaults match the template without JSON sources' {
    function Initialize-LocalJsonFiles { param($DefaultConfigPath) }
    function Read-JsonFile { param($Path) return $null }
    $ConfigPath = ''; $script:OutputRootOverride = ''
    $config = Resolve-Config
    Assert ($config.CitrixWebUrl -eq 'https://workplacecloudhub.com/demo/smartthinclient/citrix/' -and $config.AvdWebUrl -eq 'https://workplacecloudhub.com/demo/smartthinclient/avd/') 'Built-in defaults disagree with template'
    Assert ($config.WebOnlyUrl -eq 'https://workplacecloudhub.com/') 'Built-in WebOnly default missing'
}
Check 'Existing custom runtime URLs survive template synchronization unchanged' {
    $runtimePath = Join-Path $testRoot 'custom-demo-config.json'
    $custom = Get-Content -Raw -LiteralPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') | ConvertFrom-Json
    $custom.WebOnlyUrl = 'https://custom-web.example.invalid'
    $custom.CitrixWebUrl = 'https://custom-citrix.example.invalid'
    $custom.AvdWebUrl = 'https://custom-avd.example.invalid'
    $custom | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $runtimePath -Encoding UTF8
    $before = (Get-FileHash -LiteralPath $runtimePath).Hash
    Merge-MissingJsonTemplateKeys -TemplatePath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -RuntimePath $runtimePath
    $ConfigPath = $runtimePath; $script:OutputRootOverride = ''
    $config = Resolve-Config
    Assert ($config.CitrixWebUrl -eq $custom.CitrixWebUrl -and $config.AvdWebUrl -eq $custom.AvdWebUrl) 'Custom runtime URL replaced'
    Assert ($config.WebOnlyUrl -eq $custom.WebOnlyUrl) 'Custom WebOnly URL replaced'
    Assert ((Get-FileHash -LiteralPath $runtimePath).Hash -eq $before) 'Complete custom runtime was rewritten'
}
Check 'Custom provider profile URLs and legacy aliases beat demo defaults' {
    function Get-ProfileConfig {
        param($ProfileName)
        switch ($ProfileName) {
            Citrix { [pscustomobject]@{CitrixWebUrl='https://profile-citrix.example.invalid'} }
            AVD { [pscustomobject]@{FeedUrl='https://profile-avd.example.invalid'} }
        }
    }
    $config = [ordered]@{CitrixWebUrl=(Get-DemoWorkspaceUrl Citrix);AvdWebUrl=(Get-DemoWorkspaceUrl AVD)}
    foreach ($profileName in @('Citrix','AVD','Hybrid')) {
        $merged = Merge-ProfileConfig -Config $config -EffectiveProfile $profileName
        $urls = Get-WorkspaceWebUrls -Config $merged -EffectiveProfile $profileName
        if ($profileName -in @('Citrix','Hybrid')) { Assert ($urls.Citrix -eq 'https://profile-citrix.example.invalid') 'Custom Citrix profile lost' }
        if ($profileName -in @('AVD','Hybrid')) { Assert ($urls.AVD -eq 'https://profile-avd.example.invalid') 'Custom AVD alias masked by demo default' }
    }
}
Check 'Main legacy URL aliases beat demo defaults without replacing custom primary URLs' {
    foreach ($alias in @('WorkspaceUrl','StoreUrl')) {
        $config = [ordered]@{CitrixWebUrl=(Get-DemoWorkspaceUrl Citrix);AvdWebUrl=(Get-DemoWorkspaceUrl AVD);FeedUrl='https://feed.example.invalid'}
        $config[$alias] = 'https://legacy-citrix.example.invalid'
        $urls = Get-WorkspaceWebUrls -Config $config -EffectiveProfile Citrix
        Assert ($urls.Citrix -eq $config[$alias] -and $urls.AVD -eq $config.FeedUrl) 'Legacy URL masked by demo'
        $config.CitrixWebUrl = 'https://primary-citrix.example.invalid'
        $config.AvdWebUrl = 'https://primary-avd.example.invalid'
        $urls = Get-WorkspaceWebUrls -Config $config -EffectiveProfile Citrix
        Assert ($urls.Citrix -eq $config.CitrixWebUrl -and $urls.AVD -eq $config.AvdWebUrl) 'Custom primary URL lost precedence'
    }
}
Check 'Explicit blank URL values remain blank on existing runtime' {
    $ConfigPath = Join-Path $testRoot 'blank-demo-config.json'
    '{"CitrixWebUrl":"","AvdWebUrl":"","WebOnlyUrl":""}' | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    $script:OutputRootOverride = ''
    $config = Resolve-Config
    $urls = Get-WorkspaceWebUrls -Config $config -EffectiveProfile Citrix
    Assert ([string]::IsNullOrEmpty($urls.Citrix) -and [string]::IsNullOrEmpty($urls.AVD)) 'Explicit blanks were silently replaced at startup'
    Assert ([string]::IsNullOrEmpty($config.WebOnlyUrl)) 'Explicit WebOnly blank replaced'
}
Check 'GUI Audit and Preview remain available with disabled mutation gates' {
    function Test-IsAdministrator { return $false }
    function Invoke-ThinClientAudit { param($Config,$RequestedProfile,$RequestedAction) [pscustomobject]@{Status=($RequestedAction+'Only');Profile=$RequestedProfile} }
    foreach($mode in @('Audit','Preview')) {
        $r=Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$false;AllowRestore=$false}) -RequestedAction $mode -SelectedProfile WebOnly
        Assert ($r.Status -eq ($mode+'Only') -and $r.Profile -eq 'WebOnly') 'Read-only GUI dispatch failed'
    }
}
Check 'GUI blocks Apply and Restore when JSON gates are false' {
    function Invoke-ThinClientApply { throw 'Mutation engine must not be reached' }
    function Invoke-ThinClientRestore { throw 'Mutation engine must not be reached' }
    foreach($mode in @('Apply','Restore')) {
        $message='';try {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$false;AllowRestore=$false}) -RequestedAction $mode -SelectedProfile WebOnly} catch {$message=$_.Exception.Message}
        Assert ($message -like '*is false in configuration*') 'Gate did not fail before engine dispatch'
    }
}
Check 'GUI requires administrator with enabled JSON gates' {
    function Test-IsAdministrator { return $false }
    foreach($mode in @('Apply','Restore')) {
        $message='';try {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$true;AllowRestore=$true}) -RequestedAction $mode -SelectedProfile WebOnly} catch {$message=$_.Exception.Message}
        Assert ($message -like '*elevated administrator*') 'Non-admin dispatch was not blocked'
    }
}
function Invoke-SyntheticGuiStatusHandler($Config,$Action,[bool]$Administrator) {
    function Test-IsAdministrator { return $Administrator }
    function Write-SmartThinClientLog { param($Level,$Message) }
    function Invoke-ThinClientGuiAction { throw 'Unexpected engine dispatch during status check' }
    $assignment=$ast.Find({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$runAction'},$true)
    if($null -eq $assignment){throw 'GUI handler not found'}
    $handler=& ([scriptblock]::Create($assignment.Right.Extent.Text))
    $profileCombo=[pscustomobject]@{Text='WebOnly'}
    $statusText=[pscustomobject]@{Text=''};$outputText=[pscustomobject]@{Text=''}
    $actionButtons=@([pscustomobject]@{IsEnabled=$true})
    $window=[pscustomobject]@{Container=[pscustomobject]@{Visibility='Visible'}}
    $window | Add-Member ScriptMethod FindName {param($Name) if($Name -ne 'AuditSummaryContainer'){throw 'Unexpected UI access'};return $this.Container}
    & $handler -RequestedAction $Action
    [pscustomobject]@{Status=$statusText.Text;Detail=$outputText.Text;ButtonsEnabled=$actionButtons[0].IsEnabled}
}
Check 'Actual GUI handler reports configuration-blocked Apply and Restore distinctly' {
    foreach($mode in @('Apply','Restore')) {
        $r=Invoke-SyntheticGuiStatusHandler -Config ([ordered]@{AllowApply=$false;AllowRestore=$false}) -Action $mode -Administrator $true
        Assert ($r.Status -ceq ($mode+' blocked by configuration') -and $r.Detail -like '*No changes were made*' -and $r.ButtonsEnabled) 'Configuration block displayed as an execution failure or left controls disabled'
    }
}
Check 'Actual GUI handler keeps other Apply and Restore failures distinct from configuration blocks' {
    foreach($mode in @('Apply','Restore')) {
        $r=Invoke-SyntheticGuiStatusHandler -Config ([ordered]@{AllowApply=$true;AllowRestore=$true}) -Action $mode -Administrator $false
        Assert ($r.Status -ceq ($mode+' did not complete') -and $r.Detail -like '*elevated administrator*' -and $r.ButtonsEnabled) 'Administrator refusal was mislabeled as a configuration block'
    }
}
Check 'GUI requires exact typed confirmation even with optional CLI phrases' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) $Config }
    foreach($mode in @('Apply','Restore')) {
        $message='';try {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$true;AllowRestore=$true;RequireConfirmationPhrase=$false}) -RequestedAction $mode -SelectedProfile WebOnly -Confirmation 'wrong'} catch {$message=$_.Exception.Message}
        Assert ($message -like '*exact confirmation phrase*') 'Missing GUI phrase gate'
    }
}
Check 'GUI cannot bypass an effective profile mutation gate' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) [ordered]@{AllowApply=$false;AllowRestore=$false} }
    Reject {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$true;AllowRestore=$true}) -RequestedAction Apply -SelectedProfile WebOnly -Confirmation 'APPLY'}
}
Check 'GUI Apply forwards selected profile phrase and secure password into existing engine' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) $Config }
    function Invoke-ThinClientApply { param($Config) [pscustomobject]@{Selected=$Profile;Phrase=$ConfirmApply;PasswordSet=($DedicatedUserPassword -is [securestring]);Status='MockApplied'} }
    $secret=ConvertTo-SecureString 'synthetic-fixture-only' -AsPlainText -Force
    try {
        $r=Invoke-ThinClientGuiAction -Config ([ordered]@{AllowApply=$true}) -RequestedAction Apply -SelectedProfile AVD -Confirmation 'APPLY' -Password $secret
        Assert ($r.Selected -eq 'AVD' -and $r.Phrase -eq 'APPLY' -and $r.PasswordSet -and $r.Status -eq 'MockApplied') 'Apply did not reuse engine inputs'
    } finally {$secret.Dispose()}
}
Check 'GUI Restore requires explicit reviewed rollback instead of latest fallback' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) $Config }
    $message='';try {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowRestore=$true}) -RequestedAction Restore -SelectedProfile WebOnly -Confirmation 'RESTORE'} catch {$message=$_.Exception.Message}
    Assert ($message -like '*explicit rollback*') 'Restore used implicit latest rollback'
}
Check 'GUI Restore refuses a rollback changed after review' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) $Config }
    $file=Join-Path $testRoot 'gui-rollback-changed.json';'{}' | Set-Content $file
    $message='';try {Invoke-ThinClientGuiAction -Config ([ordered]@{AllowRestore=$true}) -RequestedAction Restore -SelectedProfile WebOnly -Confirmation 'RESTORE' -SelectedRollbackPath $file -ExpectedRollbackHash ('0'*64)} catch {$message=$_.Exception.Message}
    Assert ($message -like '*changed after review*') 'Changed rollback was not refused'
}
Check 'GUI Restore forwards exact reviewed rollback and phrase into existing engine' {
    function Resolve-EffectiveProfile { param($RequestedProfile,$Config) $RequestedProfile }
    function Merge-ProfileConfig { param($Config,$EffectiveProfile) $Config }
    function Invoke-ThinClientRestore { param($Config) [pscustomobject]@{Path=$RollbackPath;Phrase=$ConfirmRestore;Status='MockRestored'} }
    $file=Join-Path $testRoot 'gui-rollback-reviewed.json';'{}' | Set-Content $file
    $r=Invoke-ThinClientGuiAction -Config ([ordered]@{AllowRestore=$true}) -RequestedAction Restore -SelectedProfile WebOnly -Confirmation 'RESTORE' -SelectedRollbackPath $file -ExpectedRollbackHash (Get-FileHash $file).Hash
    Assert ($r.Path -eq $file -and $r.Phrase -eq 'RESTORE' -and $r.Status -eq 'MockRestored') 'Restore engine inputs lost'
}
Check 'Preview plan describes intended configuration without generating a launcher' {
    $plan=Get-ThinClientPreviewPlan -Config ([ordered]@{AutoLaunchMode='None';EnableShellLimitations=$true;ShellRestrictionLevel='Basic';AllowApply=$false}) -Audit ([pscustomobject]@{EffectiveProfile='WebOnly';TargetUserName='SYNTHETIC'})
    Assert (-not $plan.ChangesApplied -and $plan.AutoLaunchMode -eq 'None' -and $plan.ShellRestrictions -eq 'Basic' -and $plan.EffectiveProfile -eq 'WebOnly' -and -not $plan.ApplyEnabled) 'Incorrect preview plan'
}


Check 'Visual preview accepts HTTP web targets and reports the existing WebOnly fallback' {
    $cfg=[ordered]@{CitrixWebUrl='';AvdWebUrl='';WebOnlyUrl='https://public.example.invalid/'}
    $target=Get-ThinClientVisualPreviewTarget -Config $cfg -Provider Citrix
    Assert ($target.IsValid -and $target.UsesWebOnlyFallback -and $target.Url -eq $cfg.WebOnlyUrl) 'Visual preview fallback disagrees with shell'
    $target=Get-ThinClientVisualPreviewTarget -Config $cfg -Provider WebOnly
    Assert ($target.IsValid -and -not $target.UsesWebOnlyFallback) 'WebOnly target incorrectly marked fallback'
}
Check 'Visual preview refuses missing URLs and external application schemes' {
    foreach($url in @('','file:///C:/Windows/System32/cmd.exe','ms-settings:display','javascript:alert(1)','https://user:secret@example.invalid/')) {
        $target=Get-ThinClientVisualPreviewTarget -Config ([ordered]@{WebOnlyUrl=$url}) -Provider WebOnly
        Assert (-not $target.IsValid) 'Unsafe/missing preview destination was accepted'
    }
}
Check 'Visual preview preserves Hybrid provider profile URL resolution' {
    function Get-ProfileConfig { param($ProfileName) if($ProfileName -eq 'Citrix'){[pscustomobject]@{CitrixWebUrl='https://provider.example.invalid/'}} }
    $target=Get-ThinClientVisualPreviewTarget -Config ([ordered]@{CitrixWebUrl='https://global.example.invalid/'}) -Provider Citrix -EffectiveProfile Hybrid
    Assert ($target.Url -eq 'https://provider.example.invalid/') 'Hybrid provider override lost in visual preview'
}
Check 'Visual preview and generated shell share one layout without a launch or mutation entrypoint' {
    [xml]$layout=Get-ThinClientShellXaml
    foreach($name in @('WorkspaceBrowser','CitrixButton','AvdButton','WebOnlyButton','RestartButton','ShutdownButton')) {
        Assert ($null -ne $layout.SelectSingleNode("//*[@Name='$name']")) 'Shared shell layout control missing'
    }
    $fn=$ast.EndBlock.Statements | Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Show-ThinClientVisualPreview'}
    $commands=@($fn.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst]},$true) | ForEach-Object {$_.GetCommandName()})
    foreach($name in @('Start-Process','Invoke-ThinClientLaunchOnly','New-LauncherScript','Invoke-ThinClientApply','Invoke-ThinClientRestore','shutdown.exe','Restart-Computer','Stop-Computer')) {
        Assert ($name -notin $commands) 'Visual preview has an operational entrypoint'
    }
}


foreach($profileCase in @('Citrix','AVD','WebOnly','Hybrid')) {
    Check ("Provider toolbar matches $profileCase profile") {
        $visible=Get-ThinClientProviderVisibility -EffectiveProfile $profileCase
        foreach($provider in @('Citrix','AVD','WebOnly')) { Assert ($visible[$provider] -eq ($profileCase -eq 'Hybrid' -or $profileCase -eq $provider)) 'Other profile provider remained visible' }
    }
}
Check 'Provider toolbar global hide still overrides every profile' {
    $v=Get-ThinClientProviderVisibility -EffectiveProfile Hybrid -ShowProviderButtons $false
    Assert (-not $v.Citrix -and -not $v.AVD -and -not $v.WebOnly) 'Global provider hide ignored'
}
Check 'Shared shell keeps navigation left and session controls right of the centered title' {
    [xml]$x=Get-ThinClientShellXaml
    $left=$x.SelectSingleNode("//*[@Name='NavigationButtons']");$right=$x.SelectSingleNode("//*[@Name='SessionButtons']");$title=$x.SelectSingleNode("//*[@Name='TitleText']")
    Assert ($left.GetAttribute('Grid.Column') -eq '0' -and $right.GetAttribute('Grid.Column') -eq '2' -and $title.GetAttribute('Grid.Column') -eq '1') 'Toolbar groups lost their columns'
    Assert (($left.ChildNodes.Name -join '|') -eq 'CitrixButton|AvdButton|WebOnlyButton|BackButton|RefreshButton|ExternalButton') 'Navigation controls moved to the wrong group'
    Assert (($right.ChildNodes.Name -join '|') -eq 'SettingsButton|AdminButton|SignOutButton|RestartButton|ShutdownButton') 'Session controls moved to the wrong group'
}
Check 'Custom workspace title is preserved as a literal in the generated shell' {
    $name="Example & Co <Workspace>'s name"
    $path=New-LauncherScript -Config ([ordered]@{WebShellTitle=$name}) -EffectiveProfile Citrix -Audit $audit
    $t=$null;$e=$null;$generated=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$t,[ref]$e)
    Assert ($e.Count -eq 0) 'Custom title broke launcher syntax'
    $assignment=$generated.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$WebShellTitle'},$true)
    $literal=$assignment.Right.Find({param($n)$n -is [Management.Automation.Language.StringConstantExpressionAst]},$true)
    Assert ($literal.Value -ceq $name) 'Custom title was escaped or interpreted incorrectly'
}
Check 'Audit summary accepts actual PSCustomObject evidence and false mutation gates are neutral' {
    $a=[pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$true;AllowApply=$false;AllowRestore=$false}
    $r=Get-ThinClientAuditSummary -Audit $a
    Assert ($r.State -eq 'Success' -and $r.Note -like '*Apply disabled*' -and $r.Note -like '*Restore disabled*') 'Safe gates incorrectly failed readiness signals'
}
Check 'Audit summary warns about a local target missing and fallback without claiming remote readiness' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Citrix';BrowserInstalled=$true;CitrixWebUrlConfigured=$false;WebOnlyUrlConfigured=$true;TargetUserExists=$false})
    Assert ($r.State -eq 'Warning' -and @($r.Checks|Where-Object State -eq Warning).Count -eq 2 -and $r.Note -like '*not tested*') 'Fallback/local-target warning lost'
}
Check 'Audit summary blocks missing browser or destination instead of treating unknown as passed' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{EffectiveProfile='WebOnly';TargetUserExists=$true})
    Assert ($r.State -eq 'Blocked') 'Missing evidence incorrectly passed'
}
Check 'Limited Settings refuses other URI schemes before opening UI or processes' {
    foreach($page in @('https://example.invalid/','file:///C:/Windows/System32/cmd.exe','javascript:alert(1)')) {
        $message='';try {Show-ThinClientLimitedSettings -SettingsPages @($page)}catch{$message=$_.Exception.Message}
        Assert ($message -like '*Only configured ms-settings*') 'Settings did not reject an external scheme before UI'
    }
}
Check 'Settings launch request rejects arbitrary URIs before loading native launch code' {
    foreach($uri in @('https://example.invalid/','file:///C:/Windows/System32/cmd.exe','ms-settings:workplace','ms-settings:display?unexpected=1')){Reject {New-ThinClientSettingsRequest -Page $uri}}
}
Check 'Native device settings import without initializing devices or changing settings' {
    . (Join-Path $sourceRoot 'SmartThinClient-DeviceSettings.ps1')
    Assert ([bool](Get-Command Show-SmartThinClientDeviceSettings -CommandType Function)) 'Native settings helper did not load'
    Reject {Show-SmartThinClientDeviceSettings -Page 'ms-settings:workplace'}
}
Check 'Native menu refuses an unexpected helper before creating its window' {
    $message='';try{Show-ThinClientLimitedSettings -SettingsPages @('ms-settings:display') -NativeMode -NativeSettingsHelper $SourcePath}catch{$message=$_.Exception.Message}
    Assert ($message -like '*native device settings helper is unavailable*') 'Unexpected helper reached native UI initialization'
}
Check 'Audit summary blocks an unavailable requested kiosk platform' {
    $a=[ordered]@{OSName='Windows';EffectiveProfile='WebOnly';ShellMode='ShellLauncher';ShellLauncherClassAvailable=$false;BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$true}
    $r=Get-ThinClientAuditSummary -Audit $a
    Assert ($r.State -eq 'Blocked' -and @($r.Checks|Where-Object {$_.Label -eq 'Windows' -and $_.State -eq 'Blocked'}).Count -eq 1) 'Unsupported platform was marked ready'
}
Check 'Generated shell and Preview use shared settings and provider code' {
    $fn=$ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'New-LauncherScript'}
    Assert ($fn.Extent.Text -like '*Show-ThinClientLimitedSettings*' -and $fn.Extent.Text -like '*Get-ThinClientProviderVisibility*') 'Generated shell did not share helper behavior'
    $fn=$ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Show-ThinClientVisualPreview'}
    Assert ($fn.Extent.Text -like '*Show-ThinClientLimitedSettings*' -and $fn.Extent.Text -like '*-PreviewMode*') 'Preview Settings not connected'
}
Check 'Audit summary requires a web destination when the embedded shell overrides native preference' {
    $a=[pscustomobject]@{OSName='Windows';EffectiveProfile='Citrix';UseWebShell=$true;PreferredAccessMode='Native';CitrixWorkspaceInstalled=$true;BrowserInstalled=$true;TargetUserExists=$true}
    Assert ((Get-ThinClientAuditSummary -Audit $a).State -eq 'Blocked') 'Native signal masked missing embedded-shell URL'
}
Check 'Audit summary recognizes native-only access without an installed browser' {
    $a=[pscustomobject]@{OSName='Windows';EffectiveProfile='Citrix';UseWebShell=$false;PreferredAccessMode='Native';CitrixWorkspaceInstalled=$true;BrowserInstalled=$false;TargetUserExists=$true}
    Assert ((Get-ThinClientAuditSummary -Audit $a).State -eq 'Success') 'Native-only client incorrectly required a web destination'
}
Check 'Citrix Web configuration never implies Citrix Workspace installation' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Citrix';BrowserInstalled=$true;CitrixWebUrlConfigured=$true;CitrixWorkspaceInstalled=$false;TargetUserExists=$true})
    Assert ($r.State -eq 'Success' -and @($r.Checks|Where-Object {$_.Label -eq 'Citrix Web' -and $_.State -eq 'Success' -and $_.Detail -eq 'URL configured'}).Count -eq 1) 'Web destination signal lost'
    Assert (@($r.Checks|Where-Object {$_.Label -eq 'Citrix Workspace' -and $_.State -eq 'Info' -and $_.Detail -eq 'Not detected'}).Count -eq 1) 'Missing native client was hidden or shown as installed'
}
Check 'Installed Citrix Workspace has its own successful card without masking a missing Web destination' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Citrix';UseWebShell=$true;BrowserInstalled=$true;CitrixWorkspaceInstalled=$true;TargetUserExists=$true})
    Assert ($r.State -eq 'Blocked' -and @($r.Checks|Where-Object {$_.Label -eq 'Citrix Workspace' -and $_.State -eq 'Success' -and $_.Detail -eq 'Detected'}).Count -eq 1) 'Installed native client masked missing Web destination'
}
Check 'AVD Web remains distinct from absent Windows App and legacy clients' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='AVD';BrowserInstalled=$true;AvdWebUrlConfigured=$true;AvdClientInstalled=$false;WindowsAppPackage='';AvdClassicClientPath='';TargetUserExists=$true})
    Assert ($r.State -eq 'Success' -and @($r.Checks|Where-Object {$_.Label -eq 'AVD Web' -and $_.State -eq 'Success'}).Count -eq 1) 'AVD Web signal lost'
    Assert (@($r.Checks|Where-Object {$_.Label -in @('Windows App','Remote Desktop (legacy)') -and $_.State -eq 'Info' -and $_.Detail -eq 'Not detected'}).Count -eq 2) 'AVD Web configuration implied native clients'
}
Check 'Windows App package is identified separately from legacy Remote Desktop' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='AVD';BrowserInstalled=$true;AvdWebUrlConfigured=$true;AvdClientInstalled=$true;WindowsAppPackage='MicrosoftCorporationII.Windows365';AvdClassicClientPath='';TargetUserExists=$true})
    Assert ($r.State -eq 'Success' -and @($r.Checks|Where-Object {$_.Label -eq 'Windows App' -and $_.State -eq 'Success' -and $_.Detail -eq 'Detected'}).Count -eq 1) 'Windows App package not identified'
    Assert (@($r.Checks|Where-Object {$_.Label -eq 'Remote Desktop (legacy)' -and $_.Detail -eq 'Not detected'}).Count -eq 1) 'Windows App was confused with legacy client'
}
Check 'Legacy AVD client is a warning and never implies Windows App presence' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='AVD';BrowserInstalled=$true;AvdWebUrlConfigured=$true;AvdClientInstalled=$true;WindowsAppPackage='';AvdClassicClientPath='C:\Synthetic\msrdcw.exe';TargetUserExists=$true})
    Assert ($r.State -eq 'Warning' -and @($r.Checks|Where-Object {$_.Label -eq 'Remote Desktop (legacy)' -and $_.State -eq 'Warning' -and $_.Detail -eq 'Detected (legacy)'}).Count -eq 1) 'Legacy installation was presented as a current client'
    Assert (@($r.Checks|Where-Object {$_.Label -eq 'Windows App' -and $_.Detail -eq 'Not detected'}).Count -eq 1) 'Legacy aggregate signal became a Windows App detection'
}
Check 'Both AVD native clients retain distinct installation signals' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='AVD';BrowserInstalled=$true;AvdWebUrlConfigured=$true;AvdClientInstalled=$true;WindowsAppPackage='MicrosoftCorporationII.Windows365';AvdClassicClientPath='C:\Synthetic\msrdcw.exe';TargetUserExists=$true})
    Assert (@($r.Checks|Where-Object {$_.Label -eq 'Windows App' -and $_.State -eq 'Success'}).Count -eq 1 -and @($r.Checks|Where-Object {$_.Label -eq 'Remote Desktop (legacy)' -and $_.State -eq 'Warning'}).Count -eq 1) 'Concurrent client signals were collapsed'
}
Check 'Required native client absence warns with Web fallback and blocks without it' {
    foreach($provider in @('Citrix','AVD')) {
        $a=[ordered]@{OSName='Windows';EffectiveProfile=$provider;UseWebShell=$false;PreferredAccessMode='Native';BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true;CitrixWorkspaceInstalled=$false;WindowsAppPackage='';AvdClassicClientPath='';TargetUserExists=$true}
        Assert ((Get-ThinClientAuditSummary -Audit $a).State -eq 'Warning') 'Missing required native client ignored despite Web fallback'
        $a.BrowserInstalled=$false
        Assert ((Get-ThinClientAuditSummary -Audit $a).State -eq 'Blocked') 'Missing native client and browser incorrectly passed'
    }
}
Check 'Unknown client details are never presented as installed or definitely absent' {
    foreach($provider in @('Citrix','AVD')) {
        $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile=$provider;BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true;AvdClientInstalled=$true;TargetUserExists=$true})
        $clientCards=@($r.Checks|Where-Object Label -in @('Citrix Workspace','Windows App','Remote Desktop (legacy)'))
        Assert ($clientCards.Count -gt 0 -and @($clientCards|Where-Object {$_.State -ne 'Info' -or $_.Detail -ne 'Not reported'}).Count -eq 0) 'Unknown client evidence inferred an installation state'
    }
}
Check 'Hybrid summary shows all five distinct provider checks' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Hybrid';BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true;CitrixWorkspaceInstalled=$false;WindowsAppPackage='';AvdClassicClientPath='';TargetUserExists=$true})
    $expected=@('Citrix Web','Citrix Workspace','AVD Web','Windows App','Remote Desktop (legacy)')
    Assert ($r.Checks.Count -eq 8 -and @($expected|Where-Object {$_ -notin $r.Checks.Label}).Count -eq 0) 'Hybrid lost provider/client cards'
}
Check 'Missing native clients offer only their current official download destinations' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Hybrid';BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true;CitrixWorkspaceInstalled=$false;WindowsAppPackage='';AvdClassicClientPath='';TargetUserExists=$true})
    $downloads=@($r.Checks|Where-Object DownloadClient)
    Assert ($downloads.Count -eq 2 -and 'CitrixWorkspace' -in $downloads.DownloadClient -and 'WindowsApp' -in $downloads.DownloadClient) 'Expected two current-client download prompts'
    foreach($item in $downloads){$target=Get-ThinClientClientDownload -Client $item.DownloadClient;Assert ($target.Url -match '^https://' -and $target.Label -like 'Download *') 'Download destination was not an HTTPS page'}
    Assert ((Get-ThinClientClientDownload -Client CitrixWorkspace).Url -eq 'https://www.citrix.com/downloads/workspace-app/windows/') 'Citrix download destination changed'
    Assert ((Get-ThinClientClientDownload -Client WindowsApp).Url -eq 'https://apps.microsoft.com/detail/9n1f85v9t8bn') 'Windows App download destination changed'
}
Check 'Detected clients suppress download prompts and WebOnly offers no native clients' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='Hybrid';BrowserInstalled=$true;CitrixWebUrlConfigured=$true;AvdWebUrlConfigured=$true;CitrixWorkspaceInstalled=$true;WindowsAppPackage='MicrosoftCorporationII.Windows365';AvdClassicClientPath='';TargetUserExists=$true})
    Assert (@($r.Checks|Where-Object DownloadClient).Count -eq 0) 'Detected client still requested installation'
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$true})
    Assert (@($r.Checks|Where-Object DownloadClient).Count -eq 0) 'WebOnly offered unrelated client downloads'
}
Check 'Legacy Remote Desktop does not suppress the Windows App download prompt' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='AVD';BrowserInstalled=$true;AvdWebUrlConfigured=$true;WindowsAppPackage='';AvdClassicClientPath='C:\Synthetic\msrdcw.exe';TargetUserExists=$true})
    $downloads=@($r.Checks|Where-Object DownloadClient)
    Assert ($downloads.Count -eq 1 -and $downloads[0].DownloadClient -eq 'WindowsApp') 'Legacy client masked current-client download'
}
Check 'Download destination selection refuses arbitrary URLs and legacy client identifiers' {
    foreach($client in @('https://example.invalid/','file:///C:/Windows/System32/cmd.exe','RemoteDesktop','')){Reject {Get-ThinClientClientDownload -Client $client}}
}
$settingsDefaultPages=@('ms-settings:network-status','ms-settings:display','ms-settings:sound','ms-settings:dateandtime')
Check 'Settings exposes four default pages without requiring new configuration keys' {
    $pages=@(Get-ThinClientVisibleSettingsPages -SettingsPages $settingsDefaultPages)
    Assert (($pages -join '|') -eq ($settingsDefaultPages -join '|')) 'Default Settings pages changed'
}
Check 'Fresh Settings defaults allow optional Bluetooth and Printers but show only the original four' {
    $savedRoot=$script:ScriptRoot
    try {
        $script:ScriptRoot=Join-Path $testRoot 'settings-template-free'
        New-Item -ItemType Directory $script:ScriptRoot -Force | Out-Null
        $ConfigPath=''
        $defaults=Resolve-Config
        $template=Get-Content -LiteralPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Raw | ConvertFrom-Json
        foreach($config in @($defaults,(Convert-ObjectToHashtable -InputObject $template))) {
            Assert ($config.WebShellSettingsButtons.Bluetooth -eq $false -and $config.WebShellSettingsButtons.Printers -eq $false) 'Optional pages enabled by default'
            $visible=@(Get-ThinClientVisibleSettingsPages -SettingsPages $config.WebShellAllowedSettingsPages -SettingsButtons $config.WebShellSettingsButtons)
            Assert (($visible -join '|') -eq ($settingsDefaultPages -join '|')) 'Default visible pages changed'
            $buttons=ConvertTo-ThinClientSettingsButtons -Value $config.WebShellSettingsButtons
            $buttons.Bluetooth=$true;$buttons.Printers=$true
            $visible=@(Get-ThinClientVisibleSettingsPages -SettingsPages $config.WebShellAllowedSettingsPages -SettingsButtons $buttons)
            Assert ($visible.Count -eq 6 -and 'ms-settings:bluetooth' -in $visible -and 'ms-settings:printers' -in $visible) 'Explicitly enabled optional pages missing'
        }
    } finally { $script:ScriptRoot=$savedRoot }
}
foreach($optionalCase in @(@('Bluetooth','ms-settings:bluetooth'),@('Printers','ms-settings:printers'))) {
    Check ('Settings independently hides optional '+$optionalCase[0]) {
        $buttons=[ordered]@{Bluetooth=$true;Printers=$true};$buttons[$optionalCase[0]]=$false
        $visible=@(Get-ThinClientVisibleSettingsPages -SettingsPages @($settingsDefaultPages+'ms-settings:bluetooth'+'ms-settings:printers') -SettingsButtons $buttons)
        Assert ($visible.Count -eq 5 -and $optionalCase[1] -notin $visible) 'Optional visibility hid the wrong page'
    }
}
Check 'Existing explicitly configured optional pages stay visible without visibility flags' {
    $visible=@(Get-ThinClientVisibleSettingsPages -SettingsPages @('ms-settings:bluetooth','ms-settings:printers'))
    Assert ($visible.Count -eq 2) 'Existing allowed-page choice was lost'
}
foreach($buttonCase in @(@('Network','ms-settings:network-status'),@('Display','ms-settings:display'),@('Sound','ms-settings:sound'),@('DateAndTime','ms-settings:dateandtime'))) {
    Check ('Settings independently hides '+$buttonCase[0]) {
        $flags=[ordered]@{};$flags[$buttonCase[0]]=$false
        $pages=@(Get-ThinClientVisibleSettingsPages -SettingsPages $settingsDefaultPages -SettingsButtons $flags)
        Assert ($pages.Count -eq 3 -and $buttonCase[1] -notin $pages) 'Button visibility affected the wrong page'
    }
}
Check 'All Settings pages can be hidden without restoring defaults' {
    $flags=[pscustomobject]@{Network=$false;Display=$false;Sound=$false;DateAndTime=$false}
    Assert (@(Get-ThinClientVisibleSettingsPages -SettingsPages $settingsDefaultPages -SettingsButtons $flags).Count -eq 0) 'False flags restored hidden defaults'
}
Check 'Settings flags do not add a page removed from the allowed list' {
    $pages=@(Get-ThinClientVisibleSettingsPages -SettingsPages @('ms-settings:display') -SettingsButtons ([pscustomobject]@{Network=$true;Display=$true;Sound=$true;DateAndTime=$true;Bluetooth=$true;Printers=$true}))
    Assert ($pages.Count -eq 1 -and $pages[0] -eq 'ms-settings:display') 'Existing allow-list choice was overwritten'
}
Check 'Settings visibility rejects string booleans and invalid root types' {
    foreach($value in @([pscustomobject]@{Network='false'},'false',@($true,$false))) { Reject {ConvertTo-ThinClientSettingsButtons -Value $value} }
}
Check 'Partial profile Settings flags inherit unspecified global visibility without changing global values' {
    function Get-ProfileConfig {param($ProfileName) [pscustomobject]@{WebShellSettingsButtons=[pscustomobject]@{Sound=$false;Bluetooth=$true}}}
    $cfg=[ordered]@{WebShellSettingsButtons=[pscustomobject]@{Network=$false;Display=$true;Bluetooth=$false;Printers=$false}}
    $merged=Merge-ProfileConfig -Config $cfg -EffectiveProfile Citrix
    Assert (-not $merged.WebShellSettingsButtons.Network -and -not $merged.WebShellSettingsButtons.Sound -and $merged.WebShellSettingsButtons.Display) 'Partial override lost inherited flags'
    Assert (-not $cfg.WebShellSettingsButtons.Network -and -not $cfg.WebShellSettingsButtons.PSObject.Properties['Sound']) 'Global Settings flags were mutated'
    Assert ($merged.WebShellSettingsButtons.Bluetooth -and -not $merged.WebShellSettingsButtons.Printers -and -not $cfg.WebShellSettingsButtons.Bluetooth) 'Optional profile flags lost inheritance or changed global state'
}
Check 'Generated launcher serializes boolean Settings flags and shares filtering logic' {
    $file=New-LauncherScript -Config ([ordered]@{WebShellSettingsButtons=[pscustomobject]@{Network=$false;Display=$true;Bluetooth=$true;Printers=$false};WebShellAllowedSettingsPages=@($settingsDefaultPages+'ms-settings:bluetooth'+'ms-settings:printers')}) -EffectiveProfile Citrix -Audit $audit
    $t=$null;$err=$null;$generated=[Management.Automation.Language.Parser]::ParseFile($file,[ref]$t,[ref]$err)
    Assert ($err.Count -eq 0) 'Visibility flags broke generated PowerShell'
    $text=Get-Content $file -Raw
    Assert ($text.Contains("'Network' = `$false") -and $text -like '*Get-ThinClientVisibleSettingsPages*') 'Generated shell lost visibility flags/filter'
    Assert ($text.Contains("'Bluetooth' = `$true") -and $text.Contains("'Printers' = `$false") -and $text.Contains("'ms-settings:bluetooth'") -and $text.Contains("'ms-settings:printers'")) 'Generated shell lost optional pages or boolean values'
}
function New-SyntheticTarget($Name='TEST-USER',$Mode='ExistingUser') { [pscustomobject]@{Mode=$Mode;UserName=$Name;LocalUserExists=$false;Enabled=$false;Sid=''} }
Check 'Current domain account resolves from its Windows token without changing local operational state' {
    $target=New-SyntheticTarget
    $identity=Resolve-ThinClientTargetIdentity -TargetUser $target
    Assert ($identity.Exists -and $identity.AccountType -eq 'Domain' -and $identity.Name -eq 'EXAMPLE\TEST-USER' -and $identity.Source -eq 'CurrentWindowsToken') 'Current domain identity not detected'
    Assert ($null -eq $identity.Enabled -and -not $target.LocalUserExists -and $target.Sid -eq '') 'Domain detection changed local guard state or invented directory enabled status'
}
Check 'Qualified current domain name matches case-insensitively' {
    $r=Resolve-ThinClientTargetIdentity -TargetUser (New-SyntheticTarget 'example\test-user')
    Assert ($r.Status -eq 'Detected' -and $r.AccountType -eq 'Domain') 'Qualified current identity not detected'
}
Check 'Another domain with the same short username is not confused with the current token' {
    $r=Resolve-ThinClientTargetIdentity -TargetUser (New-SyntheticTarget 'OTHER\TEST-USER')
    Assert ($null -eq $r.Exists -and $r.Status -eq 'NotVerified' -and $r.Sid -eq '') 'Another domain inherited the current SID'
}
Check 'Unresolved other account stays unknown instead of being reported nonexistent' {
    foreach($name in @('someone-else','other@example.invalid')) {
        $r=Resolve-ThinClientTargetIdentity -TargetUser (New-SyntheticTarget $name)
        Assert ($null -eq $r.Exists -and $r.Status -eq 'NotVerified') 'Unknown account was classified as missing or current'
    }
}
Check 'Dedicated local user requirement is not satisfied by a current domain account of the same name' {
    $r=Resolve-ThinClientTargetIdentity -TargetUser (New-SyntheticTarget 'TEST-USER' 'DedicatedUser')
    Assert (-not $r.Exists -and $r.AccountType -eq 'Local' -and $r.Source -eq 'LocalAccountLookup') 'Domain account satisfied dedicated-local requirement'
}
Check 'Verified local account keeps its original SID and enabled state' {
    $target=[pscustomobject]@{Mode='ExistingUser';UserName='LOCAL-USER';LocalUserExists=$true;Enabled=$true;Sid='S-1-5-21-4-5-6-1002'}
    $r=Resolve-ThinClientTargetIdentity -TargetUser $target
    Assert ($r.Exists -and $r.AccountType -eq 'Local' -and $r.Enabled -and $r.Sid -eq $target.Sid) 'Local identity changed'
}
Check 'Disabled local account is detected but remains a warning' {
    $target=[pscustomobject]@{Mode='ExistingUser';UserName='LOCAL-USER';LocalUserExists=$true;Enabled=$false;Sid='S-1-5-21-4-5-6-1002'}
    $r=Resolve-ThinClientTargetIdentity -TargetUser $target
    $s=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$r.Exists;TargetUserAccountType=$r.AccountType;TargetUserIdentityStatus=$r.Status})
    Assert ($r.Status -eq 'Disabled' -and $s.State -eq 'Warning') 'Disabled local account was marked ready'
}
Check 'Local account collision retains the operational local target rather than borrowing the domain token' {
    $target=[pscustomobject]@{Mode='ExistingUser';UserName='TEST-USER';LocalUserExists=$true;Enabled=$true;Sid='S-1-5-21-4-5-6-1002'}
    $r=Resolve-ThinClientTargetIdentity -TargetUser $target
    Assert ($r.AccountType -eq 'Local' -and $r.Sid -eq $target.Sid) 'Local/domain name collision replaced the local target'
}
Check 'Unauthenticated or SID-less token cannot confirm an identity' {
    foreach($token in @([pscustomobject]@{Name='EXAMPLE\TEST-USER';Sid='';AccountType='Domain';Authenticated=$true},[pscustomobject]@{Name='EXAMPLE\TEST-USER';Sid='S-1-5-21-1-2-3-1001';AccountType='Domain';Authenticated=$false})) {
        $r=Resolve-ThinClientTargetIdentity -TargetUser (New-SyntheticTarget) -CurrentIdentity $token
        Assert ($r.Status -eq 'NotVerified') 'Unverified token confirmed an account'
    }
}
Check 'Domain target card is successful without requiring a local account' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$true;TargetLocalUserExists=$false;TargetUserAccountType='Domain';TargetUserIdentityStatus='Detected'})
    $card=@($r.Checks|Where-Object Label -eq 'Target account')
    Assert ($r.State -eq 'Success' -and $card.Count -eq 1 -and $card[0].Detail -eq 'Domain account detected') 'Domain account retained the missing-local warning'
}
Check 'Unverified target card remains a warning even when other checks pass' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$null;TargetUserIdentityStatus='NotVerified'})
    Assert ($r.State -eq 'Warning') 'Unresolved identity incorrectly passed'
}
Check 'Domain identity detection does not relax Strict or kiosk local-user preflight guards' {
    $target=New-SyntheticTarget
    $identity=Resolve-ThinClientTargetIdentity -TargetUser $target
    foreach($cfg in @([ordered]@{EnableShellLimitations=$true;ShellRestrictionLevel='Strict'},[ordered]@{EnableAssignedAccess=$true;ShellMode='AssignedAccess'},[ordered]@{EnableShellLauncher=$true;ShellMode='ShellLauncher'})) {
        $message='';try{Assert-ApplyPreflight -Config $cfg -Audit ([pscustomobject]@{TargetUserExists=$identity.Exists;TargetUserResolvedSid=$identity.Sid;TargetUserSid=''}) -TargetUser $target}catch{$message=$_.Exception.Message}
        Assert ($message -like '*local*user*') 'Domain diagnostic identity bypassed local-only operational guards'
    }
}
Check 'System identity is not presented as an ordinary ready target' {
    $r=Get-ThinClientAuditSummary -Audit ([pscustomobject]@{OSName='Windows';EffectiveProfile='WebOnly';BrowserInstalled=$true;WebOnlyUrlConfigured=$true;TargetUserExists=$true;TargetUserIdentityStatus='Detected';TargetUserAccountType='System'})
    Assert ($r.State -eq 'Warning') 'System identity was marked ready for a user workspace'
}
Check 'APPLY is the default phrase and explicit custom phrases remain authoritative' {
    $template=Get-Content (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Raw|ConvertFrom-Json
    Assert ($template.ApplyConfirmationPhrase -ceq 'APPLY' -and $template.RequireConfirmationPhrase) 'Short default phrase or mandatory confirmation missing'
    $ConfirmApply='APPLY';Assert-ApplyAllowed -Config ([ordered]@{AllowApply=$true})
    $ConfirmApply='APPLY SMARTTHINCLIENT';Assert-ApplyAllowed -Config ([ordered]@{AllowApply=$true;ApplyConfirmationPhrase='APPLY SMARTTHINCLIENT'})
    $ConfirmApply='APPLY';Reject {Assert-ApplyAllowed -Config ([ordered]@{AllowApply=$true;ApplyConfirmationPhrase='CUSTOM PHRASE'})}
}
Check 'GUI action identifiers preserve earlier evidence and rollback names' {
    Start-ThinClientGuiRun -RequestedAction Apply
    $first=$script:RunId;$path=$script:EvidenceJsonPath
    Set-Content -LiteralPath $path -Value '{"Status":"Applied"}'
    Start-ThinClientGuiRun -RequestedAction Preview
    Assert ($script:RunId -ne $first -and $script:EvidenceJsonPath -ne $path -and (Get-Content $path -Raw) -like '*Applied*') 'Later action replaced earlier evidence'
}
Check 'Apply completion feedback identifies scope backup and next sign-in' {
    $r=Get-ThinClientOperationFeedback -Config ([ordered]@{ShellMode='CustomShell'}) -Result ([pscustomobject]@{Status='Applied';EffectiveProfile='Citrix';RollbackPath='C:\Synthetic\rollback.json';LauncherPath='C:\Synthetic\launcher.ps1';OutputRoot='C:\Synthetic'})
    Assert ($r.Title -eq 'Apply completed' -and $r.NextStep -like '*replace Explorer for all users*' -and $r.Details -like '*rollback.json*') 'Completion information missing'
    Reject {Get-ThinClientOperationFeedback -Config ([ordered]@{}) -Result ([pscustomobject]@{Status='PreviewOnly'})}
}
Check 'Completed operations offer sign-out only when a next sign-in is relevant' {
    $applied=[pscustomobject]@{Status='Applied'};$restored=[pscustomobject]@{Status='Restored'}
    foreach($cfg in @([ordered]@{ShellMode='CustomShell'},[ordered]@{ShellMode='Launcher';AutoLaunchMode='RunKey'})){
        $feedback=Get-ThinClientOperationFeedback -Result $applied -Config $cfg
        Assert ($feedback.SignOutRecommended -and $feedback.NextStep -like '*Save your open work*') 'Sign-out choice or save-work message missing'
    }
    Assert (-not (Get-ThinClientOperationFeedback -Result $applied -Config ([ordered]@{ShellMode='Launcher';AutoLaunchMode='None'})).SignOutRecommended) 'Launch-only completion offered an unnecessary sign-out'
    Assert ((Get-ThinClientOperationFeedback -Result $restored -Config ([ordered]@{})).SignOutRecommended) 'Restore omitted next-sign-in choice'
    Reject {Get-ThinClientOperationFeedback -Result ([pscustomobject]@{Status='PreviewOnly'}) -Config ([ordered]@{})}
}
Check 'Sign-out cancellation never launches a system process' {
    function Confirm-ThinClientSignOut {param($Owner) $false}
    function Start-Process {throw 'Cancellation must not launch shutdown'}
    Assert (-not (Invoke-ThinClientSignOut -Owner $null)) 'Cancelled sign-out reported a request'
}
Check 'Confirmed sign-out targets current session without forcing application closure' {
    function Confirm-ThinClientSignOut {param($Owner) $true}
    $script:signOutRequest=$null;$script:signOutExit=0
    function Start-Process {[CmdletBinding()]param($FilePath,$ArgumentList,$WindowStyle,[switch]$Wait,[switch]$PassThru) $script:signOutRequest=@{}+$PSBoundParameters;[pscustomobject]@{ExitCode=$script:signOutExit}}
    Assert (Invoke-ThinClientSignOut -Owner $null) 'Confirmed sign-out was not requested'
    Assert ($script:signOutRequest.FilePath -eq (Join-Path $env:SystemRoot 'System32\shutdown.exe') -and $script:signOutRequest.ArgumentList -ceq '/l' -and $script:signOutRequest.WindowStyle -eq 'Hidden' -and $script:signOutRequest.Wait -and $script:signOutRequest.PassThru) 'Wrong process, session scope or forced sign-out arguments'
    $script:signOutExit=5;Reject {Invoke-ThinClientSignOut -Owner $null}
}
Check 'Completion sign-out is already confirmed and never asks twice' {
    function Confirm-ThinClientSignOut {throw 'Unexpected second sign-out dialog'}
    $script:signOutRequest=$null
    function Start-Process {[CmdletBinding()]param($FilePath,$ArgumentList,$WindowStyle,[switch]$Wait,[switch]$PassThru) $script:signOutRequest=@{}+$PSBoundParameters;[pscustomobject]@{ExitCode=0}}
    Assert (Invoke-ThinClientSignOut -Confirmed) 'Explicit completion choice did not request sign-out'
    Assert ($script:signOutRequest.ArgumentList -ceq '/l') 'Completion forced application closure or targeted another session'
}
Check 'Completion button passes its explicit sign-out choice to the existing engine' {
    $feedback=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-ThinClientOperationFeedback'},$true)
    $event=$feedback.Find({param($n)$n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Expression.Extent.Text -eq '$signOut' -and $n.Member.Value -eq 'Add_Click'},$true)
    $choice=[pscustomobject]@{Confirmed=$false;Closed=$false};$dialog=[pscustomobject]@{}
    $dialog|Add-Member ScriptMethod Close {$choice.Closed=$true}
    function Invoke-ThinClientSignOut {param($Owner,[switch]$Confirmed) $choice.Confirmed=[bool]$Confirmed;return $true}
    & ($event.Arguments[0].ScriptBlock.GetScriptBlock())
    Assert ($choice.Confirmed -and $choice.Closed) 'Completion button still opened a second confirmation'
}
Check 'Restore defaults use RESTORE while an explicit custom phrase stays authoritative' {
    $template=Get-Content (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Raw|ConvertFrom-Json
    Assert ($template.RestoreConfirmationPhrase -ceq 'RESTORE') 'Template requires the old long phrase'
    $ConfirmRestore='RESTORE';Assert-RestoreAllowed -Config ([ordered]@{AllowRestore=$true})
    $custom=[ordered]@{AllowRestore=$true;RestoreConfirmationPhrase='RESTORE WINDOWS SHELL'}
    Reject {Assert-RestoreAllowed -Config $custom}
    $ConfirmRestore='RESTORE WINDOWS SHELL';Assert-RestoreAllowed -Config $custom
}
Check 'Restore completion feedback remains distinct from applying or launching' {
    $r=Get-ThinClientOperationFeedback -Config ([ordered]@{}) -Result ([pscustomobject]@{Status='Restored';RollbackPath='C:\Synthetic\selected.json'})
    Assert ($r.Title -eq 'Restore completed' -and $r.Message -like '*selected rollback*' -and $r.NextStep -like '*sign in*') 'Restore feedback was ambiguous'
}
Check 'Actual GUI Apply handler displays completion feedback instead of only JSON' {
    $script:feedbackShown=$null
    function Write-SmartThinClientLog {param($Level,$Message)}
    function Invoke-ThinClientGuiAction {param($Config,$RequestedAction,$SelectedProfile,$Confirmation,$Password) if($RequestedAction -eq 'Preview'){return $audit};[pscustomobject]@{Status='Applied';EffectiveProfile=$SelectedProfile;RollbackPath='synthetic-backup.json';OutputRoot=$testRoot}}
    function Show-ThinClientConfirmation {param($Owner,$RequestedAction,$Summary,$Phrase,$RequestPassword) [pscustomobject]@{Phrase=$Phrase;Password=$null}}
    function Show-ThinClientOperationFeedback {param($Owner,$Feedback) $script:feedbackShown=$Feedback}
    $Config=[ordered]@{AllowApply=$true;AllowRestore=$true};$profileCombo=[pscustomobject]@{Text='Citrix'}
    $statusText=[pscustomobject]@{Text=''};$outputText=[pscustomobject]@{Text=''};$actionButtons=@([pscustomobject]@{IsEnabled=$true})
    $window=[pscustomobject]@{Container=[pscustomobject]@{Visibility='Visible'}};$window|Add-Member ScriptMethod FindName {param($Name) $this.Container}
    $assignment=$ast.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$runAction'},$true)
    $handler=& ([scriptblock]::Create($assignment.Right.Extent.Text));& $handler -RequestedAction Apply
    Assert ($statusText.Text -eq 'Apply completed' -and $null -ne $script:feedbackShown -and $outputText.Text -like '*synthetic-backup.json*' -and $window.Container.Visibility -eq 'Collapsed' -and $actionButtons[0].IsEnabled) 'GUI did not show the completed operation and backup'
}
Check 'Admin panel refuses standard users before any process launch' {
    function Get-ThinClientAdminTokenState {[pscustomobject]@{IsAdministratorAccount=$false;IsElevated=$false}}
    $message='';try{Open-ThinClientAdminPanel -ScriptPath $SourcePath -ConfigPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json')}catch{$message=$_.Exception.Message}
    Assert ($message -like '*only to administrator accounts*') 'Standard account reached panel launch'
}
Check 'Filtered administrator deny-only SID is recognized without granting elevation' {
    function Add-Type {param($TypeDefinition) Microsoft.PowerShell.Utility\Add-Type -TypeDefinition $TypeDefinition}
    Initialize-ThinClientAdminTokenApi
    Assert ([SmartThinClient.AdminToken]::IsAdministratorGroup('S-1-5-32-544',16)) 'UAC deny-only administrator was hidden'
    Assert ([SmartThinClient.AdminToken]::IsAdministratorGroup('S-1-5-32-544',4)) 'Enabled administrator SID was ignored'
    Assert (-not [SmartThinClient.AdminToken]::IsAdministratorGroup('S-1-5-32-544',0)) 'Disabled SID was treated as enabled or deny-only'
    Assert (-not [SmartThinClient.AdminToken]::IsAdministratorGroup('S-1-5-32-545',16)) 'Standard user group was treated as administrator'
    Assert (-not [SmartThinClient.AdminToken]::IsAdministratorGroup('S-1-5-21-1-2-3-512',4)) 'Unexpanded domain group bypassed local administrator membership'
}
Check 'Admin panel requests elevation only for filtered administrator accounts' {
    $script:adminLaunch=$null
    function Get-ThinClientAdminTokenState {[pscustomobject]@{IsAdministratorAccount=$true;IsElevated=$false}}
    function Start-Process {[CmdletBinding()]param($FilePath,$ArgumentList,$WindowStyle,$Verb) $script:adminLaunch=@{}+$PSBoundParameters}
    Open-ThinClientAdminPanel -ScriptPath $SourcePath -ConfigPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Profile Citrix
    Assert ($script:adminLaunch.Verb -eq 'RunAs' -and $script:adminLaunch.WindowStyle -eq 'Hidden' -and $script:adminLaunch.ArgumentList -like '*-Action Audit*') 'Filtered admin was not elevated safely'
    function Get-ThinClientAdminTokenState {[pscustomobject]@{IsAdministratorAccount=$true;IsElevated=$true}}
    Open-ThinClientAdminPanel -ScriptPath $SourcePath -ConfigPath (Join-Path $sourceRoot 'SmartThinClient-Shell.config.template.json') -Profile AVD
    Assert (-not $script:adminLaunch.ContainsKey('Verb')) 'Elevated admin was prompted again'
}
Check 'Generated and Preview shells include a role-gated Admin entry' {
    [xml]$x=Get-ThinClientShellXaml
    $admin=$x.SelectSingleNode("//*[local-name()='Button' and @Name='AdminButton']")
    Assert ($admin.Visibility -eq 'Collapsed') 'Admin control was visible by default'
    $launcher=New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $text=Get-Content $launcher -Raw
    Assert ($text -like '*Get-ThinClientAdminTokenState*' -and $text -like '*Open-ThinClientAdminPanel*' -and $text.Contains('function Initialize-ThinClientAdminTokenApi') -and $text.Contains('[SmartThinClient.AdminToken]::HasAdministratorSid($identity.Token)')) 'Operational shell missed native deny-only membership detection or admin guards'
}
Check 'Operational Admin click retains workspace profile instead of automatic PowerShell profile path' {
    $path=New-LauncherScript -Config ([ordered]@{}) -EffectiveProfile Citrix -Audit $audit
    $t=$null;$e=$null;$generated=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$t,[ref]$e)
    $tag=$generated.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$adminButton.Tag'},$true)
    $event=$generated.Find({param($n)$n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Expression.Extent.Text -eq '$adminButton' -and $n.Member.Value -eq 'Add_Click'},$true)
    $adminButton=[pscustomobject]@{Tag=$null};$Profile='Citrix';$ControlPanelScript='synthetic-app.ps1';$ControlPanelConfig='synthetic-runtime.json'
    . ([scriptblock]::Create($tag.Extent.Text))
    $this=$adminButton;$Profile='C:\Users\TEST-USER\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1';$ControlPanelScript='unrelated.ps1';$ControlPanelConfig='unrelated.json';$script:adminRequest=$null
    function Open-ThinClientAdminPanel {param($ScriptPath,$ConfigPath,[ValidateSet('Citrix','AVD','WebOnly','Hybrid','Auto')]$Profile) $script:adminRequest=@{}+$PSBoundParameters}
    function Show-LauncherNotice {param($Message) throw $Message}
    & ($event.Arguments[0].ScriptBlock.GetScriptBlock())
    Assert ($script:adminRequest.Profile -ceq 'Citrix' -and $script:adminRequest.ScriptPath -ceq 'synthetic-app.ps1' -and $script:adminRequest.ConfigPath -ceq 'synthetic-runtime.json') 'Admin click used ambient callback variables instead of the captured control-panel request'
}
Check 'Custom-shell plan uses machine storage without applying changes' {
    $cfg=[ordered]@{ShellMode='CustomShell';EnableCustomShell=$true}
    $r=Get-ThinClientPreviewPlan -Config $cfg -Audit $audit
    Assert (-not $r.ChangesApplied -and $r.ShellScope -like 'All users*' -and $r.LauncherPath.StartsWith($env:ProgramData,[StringComparison]::OrdinalIgnoreCase)) 'Plan retained a user-profile launcher'
}
Check 'Primary Hybrid custom shell starts persistently without a cancellable provider picker' {
    $launcher=New-LauncherScript -Config ([ordered]@{ShellMode='CustomShell';EnableCustomShell=$true}) -EffectiveProfile Hybrid -Audit $audit
    $t=$null;$e=$null;$generated=[Management.Automation.Language.Parser]::ParseFile($launcher,[ref]$t,[ref]$e)
    if($e.Count){throw ($e|Out-String)}
    $entry=@($generated.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.IfStatementAst]})
    function Select-HybridProvider {throw 'Primary shell must not exit on provider-picker cancellation'}
    function Show-WebShell {param($InitialProvider) $script:primaryProvider=$InitialProvider}
    $Profile='Hybrid';$PrimaryCustomShell=$true;$UseWebShell=$true;$PreferredAccessMode='Web';$script:primaryProvider=''
    & ([scriptblock]::Create(($entry.Extent.Text -join "`r`n")))
    Assert ($script:primaryProvider -eq 'Hybrid') 'Primary Hybrid shell did not start'
}
Check 'Machine storage refuses user profiles and remote paths' {
    foreach($path in @($env:USERPROFILE,'\\server\share\SmartThinClient')){Reject {Get-ThinClientMachineRoot -Config ([ordered]@{MachineRoot=$path})}}
}
Check 'Custom-shell preflight requires explicit enablement supported OS persistent shell and confirmation' {
    function Get-ThinClientCustomShellPolicy {[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''}}
    $a=[pscustomobject]@{CustomShellPolicyAvailable=$true;BrowserInstalled=$true;WebOnlyUrlConfigured=$true};$u=[pscustomobject]@{Mode='ExistingUser';LocalUserExists=$false;Sid=''}
    foreach($cfg in @([ordered]@{ShellMode='CustomShell'},[ordered]@{ShellMode='CustomShell';EnableCustomShell=$true;UseWebShell=$false},[ordered]@{ShellMode='CustomShell';EnableCustomShell=$true;RequireConfirmationPhrase=$false})){Reject {Assert-ApplyPreflight -Config $cfg -Audit $a -TargetUser $u}}
    Reject {Assert-ApplyPreflight -Config ([ordered]@{ShellMode='CustomShell';EnableCustomShell=$true}) -Audit ([pscustomobject]@{CustomShellPolicyAvailable=$false}) -TargetUser $u}
    Reject {Assert-ApplyPreflight -Config ([ordered]@{ShellMode='CustomShell';EnableCustomShell=$true;PreferredAccessMode='Native'}) -Audit $a -TargetUser $u}
    Assert-ApplyPreflight -Config ([ordered]@{ShellMode='CustomShell';EnableCustomShell=$true}) -Audit $a -TargetUser $u
}
Check 'Custom-shell Apply saves rollback before policy change and restores report context' {
    $machine=Join-Path $testRoot 'MachineApply';New-Item -ItemType Directory "$machine\Launcher","$machine\Rollback" -Force|Out-Null
    $report=$script:OutputRoot;$priorRollback=$script:RollbackRoot;$script:machineSequence=@();$script:installedPolicy=$null
    function Invoke-ThinClientAudit {param($Config,$RequestedProfile,$RequestedAction) [pscustomobject]@{EffectiveProfile='Citrix';TargetUserName='TEST-USER';TargetUserSid='';CustomShellPolicyAvailable=$true;BrowserInstalled=$true;CitrixWebUrlConfigured=$true;CitrixWorkspacePath='';AvdClassicClientPath='';BrowserPath=''}}
    function Get-TargetUserState {param($Config) [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$false;Sid=''}}
    function Get-ThinClientMachineRoot {param($Config) $machine}
    function Initialize-ThinClientMachineRoot {param($Config) $machine}
    function Get-ThinClientCustomShellPolicy {[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''}}
    function Get-RegistryValueSnapshot {param($Path,$Name) [pscustomobject]@{Path=$Path;Name=$Name;ValueExists=$true;ValueKind='String';Value='old launcher'}}
    function Save-RollbackState {param($State) $script:machineSequence+='backup';$script:savedMachineState=$State;'synthetic-machine-rollback.json'}
    function Protect-ThinClientMachineLauncher {param($Path) $script:machineSequence+='acl'}
    function Set-ThinClientCustomShellPolicy {param($Desired,$Expected) $script:machineSequence+='policy';$script:installedPolicy=$Desired}
    function Remove-ItemProperty {param($LiteralPath,$Name,$ErrorAction) $script:machineSequence+='startup'}
    function Apply-AutoLaunch {throw 'RunKey must not be applied for a replacement shell'}
    $ConfirmApply='APPLY'
    $r=Invoke-ThinClientApply -Config ([ordered]@{AllowApply=$true;ShellMode='CustomShell';EnableCustomShell=$true;AutoLaunchMode='None'})
    Assert ($r.Status -eq 'Applied' -and $r.ShellScope -eq 'AllUsers' -and $r.OutputRoot -eq $machine -and $r.LauncherPath.StartsWith($machine)) 'Machine-scoped result missing'
    Assert (($script:machineSequence -join ',') -eq 'backup,backup,acl,policy,startup' -and $script:savedMachineState.SchemaVersion -eq 3) 'Policy or startup changed before rollback'
    Assert ($script:OutputRoot -eq $report -and $script:RollbackRoot -eq $priorRollback -and $script:installedPolicy.Value -like '*System32*PowerShell*') 'Report context or absolute shell executable missing'
}
Check 'Custom-shell Apply policy failure preserves startup and restores the report context' {
    $machine=Join-Path $testRoot 'MachineFailure';New-Item -ItemType Directory "$machine\Launcher","$machine\Rollback" -Force|Out-Null
    $report=$script:OutputRoot;$script:startupRemoved=$false
    function Invoke-ThinClientAudit {param($Config,$RequestedProfile,$RequestedAction) [pscustomobject]@{EffectiveProfile='Citrix';TargetUserName='TEST-USER';TargetUserSid='';CustomShellPolicyAvailable=$true;BrowserInstalled=$true;CitrixWebUrlConfigured=$true;CitrixWorkspacePath='';AvdClassicClientPath='';BrowserPath=''}}
    function Get-TargetUserState {param($Config) [pscustomobject]@{Mode='ExistingUser';LocalUserExists=$false;Sid=''}}
    function Get-ThinClientMachineRoot {param($Config) $machine}
    function Initialize-ThinClientMachineRoot {param($Config) $machine}
    function Get-ThinClientCustomShellPolicy {[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''}}
    function Save-RollbackState {param($State) 'synthetic-failure-rollback.json'}
    function Protect-ThinClientMachineLauncher {param($Path)}
    function Set-ThinClientCustomShellPolicy {param($Desired,$Expected) throw 'Synthetic policy save failure'}
    function Remove-ItemProperty {$script:startupRemoved=$true}
    $ConfirmApply='APPLY'
    $message='';try{Invoke-ThinClientApply -Config ([ordered]@{AllowApply=$true;ShellMode='CustomShell';EnableCustomShell=$true;AutoLaunchMode='None'})}catch{$message=$_.Exception.Message}
    Assert ($message -eq 'Synthetic policy save failure' -and -not $script:startupRemoved -and $script:OutputRoot -eq $report) 'Failure did not reach the policy boundary or changed startup/report context'
}
Check 'Custom-shell rollback restores policy and rejects unrelated current shell changes' {
    $machine=Join-Path $testRoot 'MachineRestore';New-Item -ItemType Directory "$machine\Launcher","$machine\Rollback" -Force|Out-Null
    function Get-ThinClientMachineRoot {param($Config) $machine}
    $cfg=[ordered]@{AllowRestore=$true;ShellMode='CustomShell';EnableCustomShell=$true;AutoLaunchMode='None'}
    $state=[pscustomobject]@{SchemaVersion=3;AppName=$script:AppName;ComputerName=$env:COMPUTERNAME;Config=$cfg;Registry=@();OutputRoot=$machine;LauncherPath=(Join-Path $machine 'Launcher\SmartThinClient-LaunchWorkspace.ps1');LauncherExisted=$false;LauncherContent='';TargetUserSid='';CustomShellPolicy=[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''};ManagedShellCommand=(Get-LauncherCommandLine -Root $machine -Machine)}
    $RollbackPath=Join-Path $machine 'Rollback\selected.json';$state|ConvertTo-Json -Depth 12|Set-Content $RollbackPath
    $ConfirmRestore='RESTORE';$script:restoredPolicy=$null;$script:foreignShell=$false
    function Get-ThinClientCustomShellPolicy {[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$true;ValueKind='String';Value=$(if($script:foreignShell){'unrelated.exe'}else{$state.ManagedShellCommand})}}
    function Set-ThinClientCustomShellPolicy {param($Desired,$Expected) $script:restoredPolicy=$Desired}
    $r=Invoke-ThinClientRestore -Config $cfg
    Assert ($r.Status -eq 'Restored' -and $null -ne $script:restoredPolicy -and -not $script:restoredPolicy.ValueExists) 'Original policy was not restored'
    $script:foreignShell=$true;$script:restoredPolicy=$null
    Reject {Invoke-ThinClientRestore -Config $cfg}
    Assert ($null -eq $script:restoredPolicy) 'Foreign policy was overwritten during Restore'
}
Check 'Custom-shell rollback rejects incomplete and mismatched policy snapshots' {
    $machine=Join-Path $testRoot 'MachineValidation'
    function Get-ThinClientMachineRoot {param($Config) $machine}
    $cfg=[ordered]@{ShellMode='CustomShell';EnableCustomShell=$true}
    $state=[pscustomobject]@{SchemaVersion=3;AppName=$script:AppName;ComputerName=$env:COMPUTERNAME;Config=$cfg;Registry=@();OutputRoot=$machine;LauncherPath=(Join-Path $machine 'Launcher\SmartThinClient-LaunchWorkspace.ps1');LauncherExisted=$false;LauncherContent='';TargetUserSid='';CustomShellPolicy=[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''};ManagedShellCommand=(Get-LauncherCommandLine -Root $machine -Machine)}
    Assert-RollbackState -State $state -Config $cfg
    $state.CustomShellPolicy.ValueExists='false';Reject {Assert-RollbackState -State $state -Config $cfg}
    $state.CustomShellPolicy.ValueExists=$false;$state.ManagedShellCommand='foreign.exe';Reject {Assert-RollbackState -State $state -Config $cfg}
    $state.PSObject.Properties.Remove('CustomShellPolicy');Reject {Assert-RollbackState -State $state -Config $cfg}
}
Check 'Shell version matches release header and is embedded independently of customer title' {
    $header=[regex]::Match((Get-Content $SourcePath -Raw),'(?m)^Version:\s*(\S+)').Groups[1].Value
    Assert ((Get-ThinClientVersion) -ceq $header) 'Runtime version drifted from release header'
    [xml]$x=Get-ThinClientShellXaml
    $label=$x.SelectSingleNode("//*[@Name='VersionText']")
    Assert ($label.Text -ceq ('Smart ThinClient v'+$header)) 'Shell version is missing'
    $path=New-LauncherScript -Config ([ordered]@{WebShellTitle='Customer workspace'}) -EffectiveProfile Citrix -Audit $audit
    Assert ((Get-Content $path -Raw).Contains('Smart ThinClient v'+$header)) 'Installed launcher omitted version'
}
foreach($powerAction in @('Restart','Shutdown')) {
    Check ("$powerAction cancellation leaves shell running and launches nothing") {
        $script:ShellExitRequested=$false;$owner=[object]::new()
        function Confirm-ThinClientPowerAction {param($Owner,$Action) Assert ($Owner -eq $owner -and $Action -eq $powerAction) 'Confirmation context lost';return $false}
        Assert (-not(Invoke-ThinClientPowerAction -Owner $owner -Action $powerAction)) 'Cancellation was accepted'
        Assert (-not $script:ShellExitRequested) 'Cancel allowed primary shell to exit'
    }
    Check ("$powerAction confirmation launches correct Windows command without force") {
        $script:ShellExitRequested=$false;$script:powerLaunch=$null
        function Confirm-ThinClientPowerAction {param($Owner,$Action) return $true}
        function Start-Process {[CmdletBinding()]param($FilePath,$ArgumentList,$WindowStyle,[switch]$Wait,[switch]$PassThru)
            Assert $script:ShellExitRequested 'Primary shell would block Windows session closure'
            $script:powerLaunch=@{}+$PSBoundParameters;[pscustomobject]@{ExitCode=0}
        }
        Assert (Invoke-ThinClientPowerAction -Action $powerAction) 'Confirmed action failed'
        $expected=if($powerAction -eq 'Restart'){'/r /t 0'}else{'/s /t 0'}
        Assert ($script:powerLaunch.FilePath -eq (Join-Path $env:SystemRoot 'System32\shutdown.exe') -and $script:powerLaunch.ArgumentList -ceq $expected) 'Incorrect or forced power command'
        Assert ($script:powerLaunch.WindowStyle -eq 'Hidden' -and $script:powerLaunch.Wait -and $script:powerLaunch.PassThru) 'Power command did not capture Windows result'
    }
    foreach($failureMode in @('Launch','ExitCode')) {
        Check ("$powerAction $failureMode failure restores shell close guard") {
            function Confirm-ThinClientPowerAction {param($Owner,$Action) return $true}
            function Start-Process {[CmdletBinding()]param($FilePath,$ArgumentList,$WindowStyle,[switch]$Wait,[switch]$PassThru)
                if($failureMode -eq 'Launch'){throw 'Synthetic Windows launch failure'}
                [pscustomobject]@{ExitCode=5}
            }
            foreach($prior in @($false,$true)) {
                $script:ShellExitRequested=$prior
                Reject {Invoke-ThinClientPowerAction -Action $powerAction}
                Assert ($script:ShellExitRequested -eq $prior) 'Failure lost previous shell closing state'
            }
        }
    }
}
function Invoke-SyntheticWindowsRestoreCase([scriptblock]$Body) {
    $root=Join-Path $testRoot ('Desktop-'+[guid]::NewGuid().ToString('N'))
    $machine=Join-Path $root 'Machine';$legacy=Join-Path $root 'Legacy';$report=Join-Path $root 'Reports'
    New-Item -ItemType Directory "$machine\Rollback","$machine\Launcher","$legacy\Rollback","$legacy\Launcher",$report -Force|Out-Null
    $config=[ordered]@{ShellMode='CustomShell';EnableCustomShell=$true;MachineRoot=$machine;AllowRestore=$true;RequireConfirmationPhrase=$true;RestoreConfirmationPhrase='RESTORE'}
    $runPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    $emptyRun=[pscustomobject]@{Path=$runPath;Name='SmartThinClientShell';KeyExists=$true;ValueExists=$false;ValueKind='';Value=$null}
    $legacyRun=[pscustomobject]@{Path=$runPath;Name='SmartThinClientShell';KeyExists=$true;ValueExists=$true;ValueKind='String';Value=(Get-LauncherCommandLine -Root $legacy)}
    $emptyPolicy=[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$false;ValueKind='';Value=''}
    $managed=Get-LauncherCommandLine -Root $machine -Machine
    $common=@{AppName='Smart ThinClient Shell';ComputerName=$env:COMPUTERNAME;LauncherExisted=$false;LauncherContent='';TargetUserSid='';EffectiveProfile='Citrix'}
    $old=[pscustomobject](@{}+$common+@{SchemaVersion=2;CreatedUtc='2020-01-01T00:00:00.0000000Z';OutputRoot=$legacy;LauncherPath="$legacy\Launcher\SmartThinClient-LaunchWorkspace.ps1";Config=[ordered]@{ShellMode='Launcher';AutoLaunchMode='RunKey'};Registry=@($emptyRun)})
    $new=[pscustomobject](@{}+$common+@{SchemaVersion=3;CreatedUtc='2020-01-02T00:00:00.0000000Z';OutputRoot=$machine;LauncherPath="$machine\Launcher\SmartThinClient-LaunchWorkspace.ps1";Config=$config;Registry=@($legacyRun);ManagedShellCommand=$managed;CustomShellPolicy=$emptyPolicy})
    $oldPath=Join-Path "$legacy\Rollback" ('SmartThinClient_Rollback_'+$env:COMPUTERNAME+'_old.json')
    $newPath=Join-Path "$machine\Rollback" ('SmartThinClient_Rollback_'+$env:COMPUTERNAME+'_new.json')
    $old|ConvertTo-Json -Depth 16|Set-Content $oldPath -Encoding UTF8;$new|ConvertTo-Json -Depth 16|Set-Content $newPath -Encoding UTF8
    'synthetic machine launcher'|Set-Content $new.LauncherPath;'synthetic legacy launcher'|Set-Content $old.LauncherPath
    $fixture=[pscustomobject]@{Config=$config;Machine=$machine;Legacy=$legacy;Old=$old;New=$new;OldPath=$oldPath;NewPath=$newPath;Report=$report;Current=[pscustomobject]@{Policy=[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$true;ValueKind='String';Value=$managed};Run=$emptyRun;Winlogon=[pscustomobject]@{ValueExists=$true;ValueKind='String';Value='explorer.exe'}};Mutations=[Collections.Generic.List[string]]::new();Exports=[Collections.Generic.List[object]]::new()}
    $saved=@($script:OutputRoot,$script:RollbackRoot,$script:EvidenceJsonPath,$script:EvidenceCsvPath)
    $script:OutputRoot=$report;$script:RollbackRoot=Join-Path $machine 'Rollback';$script:EvidenceJsonPath=Join-Path $report 'restore.json';$script:EvidenceCsvPath=Join-Path $report 'restore.csv'
    $RollbackPath='';$ConfirmRestore='RESTORE'
    function Test-IsAdministrator {return $true}
    function Get-ThinClientMachineRoot {param($Config) $fixture.Machine}
    function Get-ThinClientWindowsRestoreState {$fixture.Current}
    function Get-ThinClientCustomShellPolicy {$fixture.Current.Policy}
    function Set-ThinClientCustomShellPolicy {param($Desired,$Expected) Assert (Test-ThinClientRestoreValueEqual $fixture.Current.Policy $Expected) 'Policy CAS lost';$fixture.Current.Policy=$Desired;$fixture.Mutations.Add('Policy')}
    function Restore-RegistryValue {param($Snapshot) Assert ($Snapshot.Name -eq 'SmartThinClientShell') 'Unexpected registry mutation';$fixture.Current.Run=$Snapshot;$fixture.Mutations.Add('Run')}
    function Export-Evidence {param($Evidence,$Config) $fixture.Exports.Add([pscustomobject]@{Evidence=$Evidence;Path=$script:EvidenceJsonPath})}
    try {& $Body $fixture} finally {$script:OutputRoot=$saved[0];$script:RollbackRoot=$saved[1];$script:EvidenceJsonPath=$saved[2];$script:EvidenceCsvPath=$saved[3]}
}
Check 'Windows desktop review discovers CustomShell and previous Launcher without writes' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
        Assert ($p.Steps.Count -eq 2 -and $p.Steps[0].Path -eq $f.NewPath -and $p.Steps[1].Path -eq $f.OldPath -and -not $p.AlreadyWindows) 'Wrong reverse history order'
        Assert ($f.Mutations.Count -eq 0 -and (Test-Path $f.New.LauncherPath)) 'Review changed Windows or files'
        $text=Convert-ThinClientWindowsRestorePlanToText $p
        Assert ($text -like '*configurations to undo: 2*' -and $text -like '*All users*') 'Review did not explain complete Windows restore'
    }
}
Check 'Windows desktop restore uses both original engines and preserves per-step evidence' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
        $r=Invoke-ThinClientWindowsRestore -Config $f.Config -Plan $p
        Assert ($r.WindowsDesktopRestored -and $r.RestoredSnapshots.Count -eq 2 -and -not $f.Current.Policy.ValueExists -and -not $f.Current.Run.ValueExists) 'Only the latest Apply was restored'
        Assert (-not(Test-Path $f.New.LauncherPath) -and -not(Test-Path $f.Old.LauncherPath)) 'Recorded absent launchers were not restored'
        Assert (($f.Mutations -join ',') -ceq 'Policy,Run,Run' -and $script:OutputRoot -eq $f.Report) 'Wrong write ordering or leaked legacy output context'
        Assert ($f.Exports.Count -eq 3 -and @($f.Exports.Path|Select-Object -Unique).Count -eq 3) 'Steps overwrote the aggregate or one another'
    }
}
Check 'Windows desktop review follows repeated CustomShell snapshots before legacy startup' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $newer=$f.New|ConvertTo-Json -Depth 16|ConvertFrom-Json;$newer.CreatedUtc='2020-01-03T00:00:00.0000000Z';$newer.CustomShellPolicy=$f.Current.Policy;$newer.Registry=@($f.Current.Run);$newer.LauncherExisted=$true;$newer.LauncherContent=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('prior synthetic launcher'))
        $newer|ConvertTo-Json -Depth 16|Set-Content (Join-Path (Split-Path $f.NewPath) ('SmartThinClient_Rollback_'+$env:COMPUTERNAME+'_newer.json')) -Encoding UTF8
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
        Assert ($p.Steps.Count -eq 3 -and $p.Steps[1].Path -eq $f.NewPath -and $p.Steps[2].Path -eq $f.OldPath -and $f.Mutations.Count -eq 0) 'Repeated custom shell history did not reach Windows'
    }
}
Check 'Automatic Windows restore refuses missing legacy backup before any write' {
    Invoke-SyntheticWindowsRestoreCase {param($f) Remove-Item -LiteralPath $f.OldPath;Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config};Assert ($f.Mutations.Count -eq 0 -and (Test-Path $f.New.LauncherPath)) 'Missing earlier backup allowed partial restore'}
}
Check 'Automatic Windows restore refuses unrelated shell startup and kiosk configuration' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $prior=$f.Current.Policy.Value;$f.Current.Policy.Value='unrelated.exe';Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config};$f.Current.Policy.Value=$prior
        $prior=$f.Current.Run;$f.Current.Run=[pscustomobject]@{ValueExists=$true;ValueKind='String';Value='unrelated.exe'};Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config};$f.Current.Run=$prior
        $f.Config.ShellMode='AssignedAccess';Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config}
        Assert ($f.Mutations.Count -eq 0) 'Unrelated state reached mutations'
    }
}
Check 'Automatic Windows restore refuses foreign and ambiguous history' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $f.New.ComputerName='FOREIGN-PC';$f.New|ConvertTo-Json -Depth 16|Set-Content $f.NewPath -Encoding UTF8
        Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config}
        $f.New.ComputerName=$env:COMPUTERNAME;$f.New|ConvertTo-Json -Depth 16|Set-Content $f.NewPath -Encoding UTF8
        Copy-Item $f.NewPath (Join-Path (Split-Path $f.NewPath) ('SmartThinClient_Rollback_'+$env:COMPUTERNAME+'_duplicate.json'))
        Reject {Get-ThinClientWindowsRestorePlan -Config $f.Config}
        Assert ($f.Mutations.Count -eq 0) 'Invalid history reached writes'
    }
}
Check 'Windows restore pins reviewed files current launcher and policy before writing' {
    foreach($change in @('Backup','Launcher','Policy','Plan')) {
        Invoke-SyntheticWindowsRestoreCase {
            param($f)
            $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
            switch($change){Backup{Add-Content $f.OldPath ' '}Launcher{Add-Content $f.New.LauncherPath 'changed'}Policy{$f.Current.Policy.Value='other.exe'}Plan{$p.Steps[0].SHA256='FORGED'}}
            Reject {Invoke-ThinClientWindowsRestore -Config $f.Config -Plan $p}
            Assert ($f.Mutations.Count -eq 0) 'Changed review reached a restore step'
        }
    }
}
Check 'Windows restore verifies every locked snapshot before the first mutation' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
        function Get-ThinClientWindowsRestorePlan {param($Config) $p}
        Add-Content $f.OldPath ' '
        Reject {Invoke-ThinClientWindowsRestore -Config $f.Config -Plan $p}
        Assert ($f.Mutations.Count -eq 0 -and $f.Exports[-1].Evidence.Status -eq 'RestoreFailed' -and $f.Exports[-1].Evidence.CompletedSteps.Count -eq 0) 'Later hash mismatch allowed first step to run'
    }
}
Check 'Windows restore reports partial failure and never returns completed desktop status' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config;$original=(Get-Command Invoke-ThinClientRestore).ScriptBlock;$count=[pscustomobject]@{Value=0}
        function Invoke-ThinClientRestore {param($Config) $count.Value++;if($count.Value -eq 2){throw 'Synthetic second restore failure'};& $original -Config $Config}
        $message='';try{Invoke-ThinClientWindowsRestore -Config $f.Config -Plan $p}catch{$message=$_.Exception.Message}
        Assert ($message -like '*after 1 of 2 steps*' -and $f.Exports[-1].Evidence.Status -eq 'RestoreFailed' -and $f.Current.Run.ValueExists) 'Partial restore was reported as Windows desktop restored'
        Assert ($script:OutputRoot -eq $f.Report) 'Partial failure leaked output context'
    }
}
Check 'GUI Windows restore passes through configuration administrator and exact phrase gates' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        function Resolve-EffectiveProfile {param($RequestedProfile,$Config) 'Citrix'}
        function Merge-ProfileConfig {param($Config,$EffectiveProfile) $Config}
        $p=Get-ThinClientWindowsRestorePlan -Config $f.Config
        Reject {Invoke-ThinClientGuiAction -Config $f.Config -RequestedAction Restore -SelectedProfile Citrix -Confirmation wrong -WindowsRestorePlan $p}
        $f.Config.AllowRestore=$false;Reject {Invoke-ThinClientGuiAction -Config $f.Config -RequestedAction Restore -SelectedProfile Citrix -Confirmation 'RESTORE' -WindowsRestorePlan $p};$f.Config.AllowRestore=$true
        function Test-IsAdministrator {return $false}
        Reject {Invoke-ThinClientGuiAction -Config $f.Config -RequestedAction Restore -SelectedProfile Citrix -Confirmation 'RESTORE' -WindowsRestorePlan $p}
        Assert ($f.Mutations.Count -eq 0) 'New plan bypassed old GUI guards'
    }
}
Check 'Default GUI Restore reviews Windows desktop directly and cancellation changes nothing' {
    Invoke-SyntheticWindowsRestoreCase {
        param($f)
        function Resolve-EffectiveProfile {param($RequestedProfile,$Config) 'Citrix'}
        function Merge-ProfileConfig {param($Config,$EffectiveProfile) $Config}
        function Invoke-ThinClientGuiAction {throw 'Cancelled Windows restore dispatched a mutation'}
        $review=[pscustomobject]@{Shown=$false;Windows=$false}
        function Show-ThinClientConfirmation {param($Owner,$RequestedAction,$Summary,$Phrase,$RequestPassword,[switch]$WindowsDesktop) $review.Shown=$true;$review.Windows=[bool]$WindowsDesktop;return $null}
        $Config=$f.Config;$profileCombo=[pscustomobject]@{Text='Citrix'};$statusText=[pscustomobject]@{Text=''};$outputText=[pscustomobject]@{Text=''};$actionButtons=@([pscustomobject]@{IsEnabled=$true});$window=[pscustomobject]@{}
        $assignment=$ast.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$runAction'},$true)
        $handler=& ([scriptblock]::Create($assignment.Right.Extent.Text));& $handler -RequestedAction Restore
        Assert ($review.Shown -and $review.Windows -and $statusText.Text -eq 'Restore cancelled' -and $outputText.Text -like '*configurations to undo: 2*' -and $f.Mutations.Count -eq 0 -and $actionButtons[0].IsEnabled) 'Default Restore still required a picker or ignored cancellation'
    }
}
Check 'Windows desktop already configured produces a no-change review' {
    Invoke-SyntheticWindowsRestoreCase {param($f) $f.Current.Policy=$f.New.CustomShellPolicy;$p=Get-ThinClientWindowsRestorePlan -Config $f.Config;Assert ($p.AlreadyWindows -and $p.Steps.Count -eq 0 -and $f.Mutations.Count -eq 0) 'Already-restored Windows was scheduled for another rollback'}
}
Write-Host "RESULT passed=$script:passed failed=$script:failed"
if ($script:failed) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCMclm+8rFnlaLf
# VpyGaeg1ey2aV4wbOUoG0iX2WRC5/aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMyIqdlwMWhCxuOpv4G0NAvb4B5JOSekV6ZSu1Ow80GZMA0GCSqG
# SIb3DQEBAQUABIIBgHQ9Q5pKqrdPXx5FgYbDAX4wzG4gRKdb7t/07KK4QurJPD6E
# JoN6bYUrkPb5z1I3Krc8gAayffoIr0htFlG/40XiFmxeq/P/en00hjP6Hh/w+6sv
# /O+GP80ABy1dK96/mIERVTYVohyA6JfpbTmWCymbus+J6DpxKc+5IJY2Kwhdod2z
# 0lBjoYRa6jGFtTPJwMsblwDsy9BomRujegBv4pHayZ67wOfLY5+tKVWtsFvfImq9
# STRx/e78BJdhUyy3SisXJJCoPqyDiqNDsj1Y5CU6LzSirTqvDb/m8mnmT4QIKcNW
# O4pKC7iuYktHSR+bgIgGpRLTIQvzsv1vDOuj4SlkypcYpraXTfjrFt5m/7hoDzrW
# Lp/QP0cUQ1F+k7HttvlvD9U7yQAlgdCrIydO//+TLqyuNQsVAm8a2jhSSPPwZQLn
# JsR+7MnVeik0t0uWvrJHVOBmnyJBCK8OOLFF8hSpgKjTg8LofVaUY9PRZpy2cDzH
# +tLU8iFUhk117B8B9KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMyMTQ1
# NTdaMC8GCSqGSIb3DQEJBDEiBCA/m3zeovUD0fFQerEntdoz8b4qfJfitMCGeyXF
# nkLRAjANBgkqhkiG9w0BAQEFAASCAgBNoBSSHFMfZ1+rB8FT4+va5JoMqxyxvwyB
# 8KVuy4bhuX5LEUarLlRsIkSsMLTMVHaw4xaiG/gg2+CnSp96/Kt0IyPlHR0MbnUM
# A03ysIzBf24ZjzApluQyzbrnywyz0dl9MTbZ+DODFW0O+SGgXIKlbtu+HimBkKdH
# 1efcf3/Umzx0f4B18rqe+rilw2vleM0LfwmqBrIKxTD7bl76ZhvmgFQWrLYcrhid
# KVgCQvjLKJEQHg8q1CPgOGY5c3in3wutHu4f7vqDgf8rhcMOAVn1VaBKx+KEpPwh
# K7UvqUBgupX6jWyib+s6iGT91poYgB45/il7OlWnJXILZJpHxImyOtqvV0gPjLm3
# AF9QEjr2YXNo2hpiwObz5YRnbLdi9+ixACSL4VKhuoCmtwchcpCNFVSSEEyO5X4A
# TkO7J8KYcfiVvwSD8WFKhoVNfy22oFgQPptrBQ9A5oawQr57LKYbRiT+SL+OdEEa
# Tcfs55gR7jveBcGBo2PwL6T7wkPJqEbsQ7oMX0W1TCSJFyJ9HsNuLk7gFUWDp5PG
# XsdIYK5dt8gmjD2Oh3USYoii9t74mM1baRMGQHunlL/ex9sg2ISW3EXgqDYi5mM5
# Fw5NVdveK/h1FlVHqRlL7aDcLG8paGLOlH45UEFb1+kNhnPKxWu7glqflFYnUEC+
# b4E5GEUYTA==
# SIG # End signature block
