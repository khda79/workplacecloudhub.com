# Offline setup contract checks. No setup, installation or tenant access is performed.
# Version: 1.6
$ErrorActionPreference = 'Stop'
$setup = Join-Path $PSScriptRoot '../Setup'
$installer = Join-Path $setup 'Install-SmartM365-SmartInventoryPrerequisites.ps1'
$bootstrap = Join-Path $setup 'SmartM365-Create-AppRegistration.ps1'
foreach ($path in @($installer, $bootstrap)) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Parser errors in $path : $errors" }
    if ($path -eq $bootstrap) { $bootstrapAst = $ast }
}
$functionAst = $bootstrapAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SmartM365RequiredApiResource' }, $true)
if (-not $functionAst) { throw 'Permission function missing' }
. ([scriptblock]::Create($functionAst.Extent.Text))
function Get-GraphRoles([object[]]$Resources) { return @($Resources | Where-Object Name -eq 'Microsoft Graph' | ForEach-Object { $_.AppRoleValues }) }
$managed = 'DeviceManagementManagedDevices.ReadWrite.All'
$legacy = @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.ReadWrite.All')
$default = @(Get-GraphRoles @(Get-SmartM365RequiredApiResource))
$all = @(Get-GraphRoles @(Get-SmartM365RequiredApiResource -IncludeLegacyReadWrite))
$skip = @(Get-GraphRoles @(Get-SmartM365RequiredApiResource -SkipBroadReadWrite))
if ($managed -notin $default -or @($legacy | Where-Object { $_ -in $default }).Count) { throw 'Incorrect default Intune application roles' }
if (@($legacy | Where-Object { $_ -notin $all }).Count -or $managed -notin $all) { throw 'Legacy roles missing' }
if (@(($legacy + $managed) | Where-Object { $_ -in $skip }).Count) { throw 'Skip did not omit all three roles' }
if ('Mail.Send' -in $default) { throw 'Unscoped Graph Mail.Send must not be in the baseline' }
if (@(Get-SmartM365RequiredApiResource -SkipExchange | Where-Object Name -eq 'Office 365 Exchange Online').Count) { throw 'SkipExchange failed' }
$catalogue = Get-Content -Raw (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') | ConvertFrom-Json
foreach ($job in $catalogue.Jobs) {
    foreach ($role in @($job.RequiredGraphAppRoles)) {
        if ($role -and $role -notin $default) { throw "Catalogue job $($job.Name) requires role absent from bootstrap baseline: $role" }
    }
}

$rbacAst = $bootstrapAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Set-SmartM365MailSendAppRbac' }, $true)
if (-not $rbacAst) { throw 'App RBAC function missing' }
. ([scriptblock]::Create($rbacAst.Extent.Text))
$assertAst = $bootstrapAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Assert-SmartM365MailRbacTestResult' }, $true)
if (-not $assertAst) { throw 'App RBAC result validator missing' }
. ([scriptblock]::Create($assertAst.Extent.Text))
function Write-SmartM365SetupStatus { param($Message, $Level) }
$graphConnectAst = $bootstrapAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Connect-SmartM365GraphSetupSession' }, $true)
if (-not $graphConnectAst) { throw 'Graph setup connection function missing' }
. ([scriptblock]::Create($graphConnectAst.Extent.Text))
function Connect-MgGraph { param($Scopes,$NoWelcome,$ContextScope,$TenantId,$UseDeviceCode) Write-Output 'synthetic device prompt' }
function Get-MgContext { [pscustomobject]@{ TenantId='11111111-1111-1111-1111-111111111111'; AuthType='Delegated'; Account='admin@example.com' } }
$graphResult = @(Connect-SmartM365GraphSetupSession -RequestedTenantId '11111111-1111-1111-1111-111111111111' -UseDeviceCode)
if ($graphResult.Count -ne 1 -or $graphResult[0].TenantId -ne '11111111-1111-1111-1111-111111111111') { throw 'Graph setup did not return exactly one tenant context' }
$connectAst = $bootstrapAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Connect-SmartM365ExchangeOnlineSetupSession' }, $true)
if (-not $connectAst) { throw 'Exchange setup connection function missing' }
. ([scriptblock]::Create($connectAst.Extent.Text))
function Import-RequiredExchangeOnlineModule { }
function Connect-ExchangeOnline { param($UserPrincipalName,$Organization,$ShowBanner,$ShowProgress,$DisableWAM,[switch]$Device) $script:connectedOrganization = $Organization }
function Get-ConnectionInformation { [pscustomobject]@{ State='Connected'; TenantID=$script:mockExchangeTenant } }
$TenantId='11111111-1111-1111-1111-111111111111'
$script:mockExchangeTenant=$TenantId; $script:SmartM365ExchangeOnlineConnected=$false
Connect-SmartM365ExchangeOnlineSetupSession -UserPrincipalName 'admin@example.com'
if ($script:connectedOrganization -ne $TenantId -or -not $script:SmartM365ExchangeOnlineConnected) { throw 'Exchange tenant pinning failed' }
$script:mockExchangeTenant='22222222-2222-2222-2222-222222222222'; $script:SmartM365ExchangeOnlineConnected=$false; $caught=$false
try { Connect-SmartM365ExchangeOnlineSetupSession -UserPrincipalName 'admin@example.com' } catch { $caught=$true }
if (-not $caught -or -not $script:SmartM365ExchangeOnlineConnected) { throw 'Exchange mismatch was not rejected with cleanup enabled' }
function Get-Mailbox { param($Identity) [pscustomobject]@{ ExchangeGuid = if ($Identity -eq 'allowed@example.com' -or $script:sameMailbox) { '33333333-3333-3333-3333-333333333333' } else { '44444444-4444-4444-4444-444444444444' } } }
function Get-Group { param($Identity) [pscustomobject]@{ DistinguishedName = "CN=Allowed,OU=Test,DC=example,DC=com" } }
function Get-ServicePrincipal { @() }
function New-ServicePrincipal { param($AppId,$ObjectId,$DisplayName) $script:createdPointer++ }
function Get-ManagementScope { $script:mockScopes }
function New-ManagementScope { param($Name,$RecipientRestrictionFilter) $script:createdFilter = $RecipientRestrictionFilter }
function Get-ManagementRoleAssignment { param($RoleAssignee) @() }
function New-ManagementRoleAssignment { param($Name,$App,$Role,$CustomResourceScope) $script:createdRole = $Role; $script:createdScope = $CustomResourceScope }
function Test-ServicePrincipalAuthorization {
    param($Identity,$Resource)
    $inScope = $Resource -eq 'allowed@example.com' -or $script:forceOutside
    [pscustomobject]@{ GrantedPermissions = @('Mail.Send'); InScope = $inScope; RoleName = 'Application Mail.Send'; AllowedResourceScope = 'SmartM365-MailSend-11111111111111111111111111111111'; ScopeType = 'CustomRecipientScope' }
}
$script:createdPointer = 0; $script:createdFilter = ''; $script:createdRole = ''; $script:forceOutside = $false; $script:mockScopes = @()
$rbacParams = @{ AppId='11111111-1111-1111-1111-111111111111'; ServicePrincipalId='22222222-2222-2222-2222-222222222222'; GroupAddress='group@example.com'; SenderAddress='allowed@example.com'; OutsideAddress='outside@example.com' }
$script:sameMailbox = $true; $caught = $false
try { Set-SmartM365MailSendAppRbac @rbacParams } catch { $caught = $true }
if (-not $caught -or $script:createdPointer) { throw 'Mailbox aliases were not rejected before changes' }
$script:sameMailbox = $false
Set-SmartM365MailSendAppRbac @rbacParams
if ($script:createdPointer -ne 1 -or $script:createdFilter -ne "MemberOfGroup -eq 'CN=Allowed,OU=Test,DC=example,DC=com'" -or $script:createdRole -ne 'Application Mail.Send' -or -not $script:createdScope) { throw 'App RBAC creation contract failed' }
function Get-Mailbox {
    param($Identity,$ResultSize)
    if ($Identity) {return [pscustomobject]@{ExchangeGuid='33333333-3333-3333-3333-333333333333';PrimarySmtpAddress='allowed@example.com'}}
    return @([pscustomobject]@{ExchangeGuid='33333333-3333-3333-3333-333333333333';PrimarySmtpAddress='allowed@example.com'},[pscustomobject]@{ExchangeGuid='44444444-4444-4444-4444-444444444444';PrimarySmtpAddress='outside@example.com'})
}
$autoParams=@{AppId=$rbacParams.AppId;ServicePrincipalId=$rbacParams.ServicePrincipalId;GroupAddress=$rbacParams.GroupAddress;SenderAddress=$rbacParams.SenderAddress;AutoSelectOutsideAddress=$true}
Set-SmartM365MailSendAppRbac @autoParams
$script:forceOutside=$true;$caught=$false;try {Set-SmartM365MailSendAppRbac @autoParams} catch {$caught=$true};$script:forceOutside=$false
if (-not $caught) {throw 'Automatic selection accepted a missing out-of-scope mailbox.'}
function Get-Mailbox { param($Identity) [pscustomobject]@{ ExchangeGuid = if ($Identity -eq 'allowed@example.com' -or $script:sameMailbox) { '33333333-3333-3333-3333-333333333333' } else { '44444444-4444-4444-4444-444444444444' } } }
$script:mockScopes = @([pscustomobject]@{ Name='Shared-Scope'; RecipientFilter="MemberOfGroup -eq 'CN=Allowed,OU=Test,DC=example,DC=com'"; RecipientRoot=''; ScopeRestrictionType='RecipientScope'; Exclusive=$false })
$script:createdFilter = ''; $script:createdScope = ''
function Test-ServicePrincipalAuthorization { param($Identity,$Resource) [pscustomobject]@{ GrantedPermissions=@('Mail.Send'); InScope=($Resource -eq 'allowed@example.com'); RoleName='Application Mail.Send'; AllowedResourceScope='Shared-Scope'; ScopeType='CustomRecipientScope' } }
Set-SmartM365MailSendAppRbac @rbacParams
if ($script:createdFilter -or $script:createdScope -ne 'Shared-Scope') { throw 'Equivalent shared scope was not reused' }
$script:mockScopes = @()
function Test-ServicePrincipalAuthorization { param($Identity,$Resource) [pscustomobject]@{ GrantedPermissions=@('Mail.Send'); InScope=($Resource -eq 'allowed@example.com' -or $script:forceOutside); RoleName='Application Mail.Send'; AllowedResourceScope='SmartM365-MailSend-11111111111111111111111111111111'; ScopeType='CustomRecipientScope' } }
$script:forceOutside = $true
$caught = $false
try { Set-SmartM365MailSendAppRbac @rbacParams } catch { $caught = $true }
if (-not $caught) { throw 'App RBAC accepted an in-scope negative test' }
$script:forceOutside = $false; $script:createdPointer = 0; $script:createdRole = ''
Set-SmartM365MailSendAppRbac @rbacParams -WhatIf
if ($script:createdPointer -or $script:createdRole) { throw 'WhatIf changed App RBAC state' }
$scope = 'SmartM365-MailSend-11111111111111111111111111111111'
$row = [pscustomobject]@{ GrantedPermissions='Mail.Send'; InScope='False'; RoleName='Application Mail.Send'; AllowedResourceScope=$scope; ScopeType='CustomRecipientScope' }
Assert-SmartM365MailRbacTestResult -Results @($row) -ScopeName $scope -Expected $false -Address 'outside'
foreach ($results in @(@(), @([pscustomobject]@{ GrantedPermissions='Mail.Send'; InScope='Not Run'; RoleName='Application Mail.Send'; AllowedResourceScope=$scope; ScopeType='CustomRecipientScope' }), @([pscustomobject]@{ GrantedPermissions='Mail.Send'; InScope=$false; RoleName='Application Mail.Send'; AllowedResourceScope='wrong'; ScopeType='CustomRecipientScope' }))) {
    $caught = $false
    try { Assert-SmartM365MailRbacTestResult -Results $results -ScopeName $scope -Expected $false -Address 'outside' } catch { $caught = $true }
    if (-not $caught) { throw 'Indeterminate or wrong App RBAC result was accepted' }
}
$broad = [pscustomobject]@{ GrantedPermissions='Mail.Send'; InScope=$true; RoleName='Application Exchange Full Access'; AllowedResourceScope='Organization'; ScopeType='Organization' }
$caught = $false
try { Assert-SmartM365MailRbacTestResult -Results @($row,$broad) -ScopeName $scope -Expected $false -Address 'outside' } catch { $caught = $true }
if (-not $caught) { throw 'Additional broad mail permission was accepted' }

foreach ($name in @('Resolve-ApplicationPermission', 'Assert-SmartM365NoEntraMailSend')) {
    $node = $bootstrapAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) { throw "Missing $name" }
    . ([scriptblock]::Create($node.Extent.Text))
}
function Get-MgServicePrincipal {
    param($Filter, [switch]$All)
    if ($Filter -like '*00000003-0000-0000-c000-000000000000*') {
        [pscustomobject]@{ Id='graph'; AppId='00000003-0000-0000-c000-000000000000'; DisplayName='Graph'; AppRoles=@([pscustomobject]@{ Id='mail-role'; Value='Mail.Send'; AllowedMemberTypes=@('Application') }) }
    } else { [pscustomobject]@{ Id='application-sp' } }
}
function Get-MgServicePrincipalAppRoleAssignment {
    param($ServicePrincipalId,[switch]$All)
    if ($script:hasGrant) { [pscustomobject]@{ ResourceId='graph'; AppRoleId='mail-role' } }
}
$script:hasGrant = $false
$app = [pscustomobject]@{ AppId='11111111-1111-1111-1111-111111111111'; RequiredResourceAccess=@() }
Assert-SmartM365NoEntraMailSend $app
$script:hasGrant = $true; $caught = $false
try { Assert-SmartM365NoEntraMailSend $app } catch { $caught = $true }
if (-not $caught) { throw 'Existing Entra Mail.Send grant was not rejected' }
$script:hasGrant = $false
$app.RequiredResourceAccess = @([pscustomobject]@{ ResourceAppId='00000003-0000-0000-c000-000000000000'; ResourceAccess=@([pscustomobject]@{ Type='Role'; Id='mail-role' }) })
$caught = $false
try { Assert-SmartM365NoEntraMailSend $app } catch { $caught = $true }
if (-not $caught) { throw 'Existing Entra Mail.Send declaration was not rejected' }

# Stub module discovery and the separately tested host prerequisites in a temporary script.
$fixturePath=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SetupAudit-'+[guid]::NewGuid().ToString('N')+'.ps1')
$source=[IO.File]::ReadAllText((Resolve-Path $installer))
$parseTokens=$null;$parseErrors=$null
$installerAst=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$parseTokens,[ref]$parseErrors)
$replacements=@($installerAst.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Get-SmartM365PythonPrerequisite','Invoke-SmartM365WindowsPowerShellPrerequisites')},$true)|Sort-Object {$_.Extent.StartOffset} -Descending)
if($replacements.Count -ne 2){throw 'Prerequisite fixture functions not found.'}
foreach($definition in $replacements){$mock=if($definition.Name -eq 'Get-SmartM365PythonPrerequisite'){'function Get-SmartM365PythonPrerequisite { param([switch]$MachineOnly) [pscustomobject]@{Available=$true;Version="3.10.0";Path="synthetic"} }'}else{'function Invoke-SmartM365WindowsPowerShellPrerequisites { param($Mode,$ModuleScope,$Trust,$RepositoryName) }'};$source=$source.Remove($definition.Extent.StartOffset,$definition.Extent.EndOffset-$definition.Extent.StartOffset).Insert($definition.Extent.StartOffset,$mock)}
[IO.File]::WriteAllText($fixturePath,$source,[Text.UTF8Encoding]::new($false))
$quoted = $fixturePath.Replace("'", "''")
$wrapper = @'
function Get-Module {
    param([switch]$ListAvailable, [string]$Name)
    if ($Name -eq 'ActiveDirectory') { return }
    if ($Name -eq 'Microsoft.Graph.Policies') { throw 'Invalid Graph module requested' }
    if ($Name -eq 'Microsoft.Graph.Identity.SignIns' -and $env:SETUP_TEST_MISSING_SIGNINS -eq '1') { return }
    if ($Name -eq 'ImportExcel' -and $env:SETUP_TEST_MISSING -eq '1') { return }
    $version = if ($Name -eq 'Microsoft.Graph.Beta.Groups' -and $env:SETUP_TEST_MISMATCH -eq '1') { '2.0' } else { '1.0' }
    [pscustomobject]@{ Name = $Name; Version = [version]$version; Path = 'synthetic' }
}
function Get-PSRepository { throw 'Unexpected repository access' }
function Install-Module { throw 'Unexpected installation' }
function Import-Module { throw 'Unexpected import' }
& '__INSTALLER__' -AuditOnly -SkipMicrosoftGraphMetaModule 6>&1
'@.Replace('__INSTALLER__', $quoted)
$encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($wrapper))
try {
    foreach ($scenario in @('ok','missing','missing_signins','mismatch')) {
        $env:SETUP_TEST_MISSING = if ($scenario -eq 'missing') { '1' } else { '0' }
        $env:SETUP_TEST_MISSING_SIGNINS = if ($scenario -eq 'missing_signins') { '1' } else { '0' }
        $env:SETUP_TEST_MISMATCH = if ($scenario -eq 'mismatch') { '1' } else { '0' }
        $output = & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -EncodedCommand $encoded 2>&1
        $expected = if ($scenario -eq 'ok') { 0 } else { 1 }
        if ($LASTEXITCODE -ne $expected) { throw "Audit $scenario failed: $output" }
        $outputText = $output -join "`n"
        if ($scenario -eq 'missing' -and $outputText -notmatch 'MISSING ImportExcel') { throw 'Missing module not reported' }
        if ($scenario -eq 'missing_signins' -and $outputText -notmatch 'MISSING Microsoft.Graph.Identity.SignIns') { throw 'Missing Identity.SignIns not reported' }
        if ($scenario -eq 'ok' -and $outputText -notmatch 'AVAILABLE Microsoft.Graph.Identity.SignIns') { throw 'Identity.SignIns not audited' }
        if ($scenario -eq 'mismatch' -and $outputText -notmatch 'GRAPH VERSION MISMATCH Microsoft.Graph.Beta.Groups') { throw 'Mismatch not reported' }
    }
} finally {
    Remove-Item Env:SETUP_TEST_MISSING, Env:SETUP_TEST_MISSING_SIGNINS, Env:SETUP_TEST_MISMATCH -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $fixturePath -Force -ErrorAction SilentlyContinue
}
Write-Host 'Setup offline checks passed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByFvBYKa9dRoSc
# QAw9eYxuW0bt9b3G/MaW4zKjQyIXkKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIQQM7tyULEHdACXKIaaS/5CEWj4meopjo+2mMX3p/XLMA0GCSqG
# SIb3DQEBAQUABIIBgIKExxn5ZkwFDoMNWQdvyZYlM/+BRGdysaTFWaCCWElsVx5y
# x6gL7Tm7WESiKzbTFUgQL0PU2o39bq+JZ7JbmtND6+RCbcUNfqawvDGvgmDvA8VP
# 1wW7SKzXm9jBSjdI8GIBnGrKNl67Cc77Npuu/0QDgz2EPcjyGRB2DeZ5IFsY18qM
# IlfuRrsXvHhNAYC1gvLcGubCWDxzRZBCLFQWHySoUyyW6bkB0wzi9fDxPrViKaxH
# XUZdBJFvCfyh/8MBkC4MlygLVWrd1aeEKBZtF8CzwnAUeKTAPFl62TFCdYj1WRIs
# yZQsskB2tsiMIyQZu3n+w/RpOsnuRvej9BRapKjfclQV6J8Tunmk9qw92wY7b2PU
# ccwoc7ckZ3v52sXu2PEX3iumJKihEezwBoXoUX0tw2cadR2nndwTIdLHdDEBkCl3
# P4KqRdm7Z3SY55iCvSkgpXL2TijEJ136YRTqB9NOezLM1N1cy3ZEPfEvzAZzlS7E
# SPdOVB2R7GToOhTJqaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTMx
# MTNaMC8GCSqGSIb3DQEJBDEiBCBjnFiBjzozbVF8R8KakBjb4YwGr2Fn0ad4Gn8l
# Ju1EeTANBgkqhkiG9w0BAQEFAASCAgBe0bVvRvy1WpA8ZtppiFlPAEyEPq8Cw4r5
# vlqZFZKJfZi3PiJ1Pc1VkWXdQ7JSU24ucnbsYROYZwAnJhPGQBUWDbbp1pt34rcu
# u36IzaEgxhPDmO1jmPKejOWn1/06H2g4mre+gadVIcyfgVkQpLKwt5e6E2CjL0GQ
# xBj+U0FNGrgOvDznllSvtQ0vXcELXH8fN3vcF7iOCcQmchs0d9u9D1BilLYt0LIt
# ah35p9MUn4LV1vzOvi6u5lVThJ3797EL3y7r9g3467TfVzl12VhvXg4eQQuh9iEX
# S5dfhITEJAMKRfujmIQKpBTa0ImFHAyIjP3tbJpiEeEdHBI32sQfVb5uFqfDWSXx
# Im2E0jM0lCbE+xzOzgd4bzpIo8G1KtOxz379k+N5CmdfdHI2prupj3gF46tiPSC1
# 2FqTk5Ar2cMAy30h7syn6u3cfomKAJLzPkYQHBMHPnvRl6nr/SvaovQ+OMLB+Yyg
# e/FGqqCeswBYT61j/KEg538pkCVrGbYZCNTUpJNKohNMEnCDFMNVhNdaAO45YbsV
# 0YSVftK0fAKBLCzu0iuo1W9ZLYyEATxUZCifYKKSmP9dWI9fPiDDVGy+2LAy5vYT
# dtjTuWp7hq+kbji821Jow6GVme3Tp5rFGrb3iR3m1ublHpDnpCXGSINd4Y5bMtcQ
# CEy0rEDrxQ==
# SIG # End signature block
