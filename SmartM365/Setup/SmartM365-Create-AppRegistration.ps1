#Requires -Version 7.0

<#
.SYNOPSIS
    Creates or updates the Entra ID app registration used by SmartM365 automation.

.DESCRIPTION
    Creates an app registration, adds the Microsoft Graph application permissions
    required by the SmartM365 inventory/export scripts, optionally adds the
    Exchange Online app-only API permission, and uploads a certificate public key.

    This bootstrap script is intentionally interactive-only. It signs in with
    delegated Microsoft Graph setup scopes and must be run by an administrator
    allowed to create app registrations, update application permissions, and
    grant tenant-wide admin consent by default.

    Use -DisableGrantAdminConsent only when you want to add permissions without
    granting tenant-wide consent immediately.

    The script also creates or reuses a Teams team for SmartM365 shared exports,
    resolves its backing SharePoint site, and updates the selected tenant local profile
    with the SharePoint upload target.

    The script also creates or reuses a dedicated Exchange Online shared mailbox
    for SmartM365 report/error emails, creates a mail-enabled security group,
    adds the sender mailbox to that group, and scopes Microsoft Graph Mail.Send
    with an Exchange Online Application Access Policy.

.PARAMETER DisplayName
    Display name for the Entra ID application. Defaults to SmartM365 Automation.

.PARAMETER TenantId
    Optional tenant ID or verified domain to connect to.

.PARAMETER CertificateThumbprint
    Existing certificate thumbprint to upload from CurrentUser\My or LocalMachine\My.
    If omitted while updating an existing app, a matching local certificate
    already attached to the app is reused when possible. Otherwise, a new
    self-signed certificate is created in CurrentUser\My.
    In -WhatIf mode, provide an existing thumbprint to avoid local certificate changes.

.PARAMETER UpdateExisting
    Update the existing app registration with the same display name instead of
    failing when it already exists.

.PARAMETER RemoveAppRegistration
    Removes the app registration with the requested display name, removes
    matching application service principals, and clears app-only authentication
    fields from the selected tenant local profile. This does not remove the Teams
    workspace or local certificates. Use with -Confirm.

.PARAMETER DisableGrantAdminConsent
    Adds the requested API permissions without granting tenant-wide admin
    consent.

.PARAMETER UseDeviceCode
    Uses interactive device code authentication for Microsoft Graph setup
    sign-in. Useful when browser/WAM authentication is not available in the
    current terminal.

.PARAMETER ExchangeAdminUserPrincipalName
    Optional Exchange Online administrator account used for the interactive
    Exchange Online setup connection. If omitted, Exchange Online prompts for
    the account before Microsoft Graph connects.

.PARAMETER TeamDisplayName
    Display name of the Teams team used for SmartM365 shared exports. Defaults
    to SMART-M365.

.PARAMETER MailSenderAddress
    Primary SMTP address for the SmartM365 sender mailbox. If omitted, the
    script uses smartm365-reports@<default accepted domain>.

.PARAMETER MailSenderDisplayName
    Display name for the SmartM365 sender shared mailbox.

.PARAMETER MailSendSecurityGroupName
    Display name for the mail-enabled security group that scopes Mail.Send.

.PARAMETER MailSendSecurityGroupAlias
    Alias for the mail-enabled security group that scopes Mail.Send.

.PARAMETER DisableTeamsSetup
    Skips Teams team creation/reuse and SharePoint local configuration updates.

.PARAMETER LogPath
    Folder where the setup log and transcript are written. Defaults to
    Data\Tenants\<TenantKey>\LOG-ALL\Setup under the SmartM365 root, with
    fallback to Setup\Output\Tenants\<TenantKey>\LOG-ALL\Setup.

.EXAMPLE
    .\Setup\SmartM365-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Setup\SmartM365-Create-AppRegistration.ps1 -DisplayName SmartM365 -TenantId 00000000-0000-0000-0000-000000000000 -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UpdateExisting

.EXAMPLE
    .\Setup\SmartM365-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UseDeviceCode -WhatIf

.EXAMPLE
    .\Setup\SmartM365-Create-AppRegistration.ps1 -RemoveAppRegistration -Confirm
.REQUIREMENTS
    PowerShell 7+.
    Modules: Microsoft.Graph.Authentication; Microsoft.Graph.Applications; ExchangeOnlineManagement.
    Interactive delegated setup scopes: Application.ReadWrite.All; AppRoleAssignment.ReadWrite.All; Directory.Read.All; Group.ReadWrite.All; RoleManagement.ReadWrite.Directory; Sites.FullControl.All; Channel.Create; Channel.ReadBasic.All.
    Runtime app permissions created: see SmartM365-AppRegistration-Permissions.md and Get-RequiredApiResource.

.VERSION
1.2
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Tenant = 'test',
[Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName = 'SmartM365 Automation',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CertificateThumbprint,

    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$CertificateYears = 2,

    [Parameter()]
    [switch]$UpdateExisting,

    [Parameter()]
    [switch]$RemoveAppRegistration,

    [Parameter()]
    [switch]$DisableGrantAdminConsent,

    [Parameter()]
    [switch]$UseDeviceCode,

    [Parameter()]
    [switch]$SkipExchangeOnlinePermission,

    [Parameter()]
    [switch]$SkipBroadIntuneReadWritePermissions,

    [Parameter()]
    [string]$ExchangeAdminUserPrincipalName,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamDisplayName = 'SMART-M365',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamMailNickname = 'SMARTM365',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SharePointTargetFolderPath = 'SMART-M365/DATA',

    [Parameter()]
    [string]$MailSenderAddress,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSenderDisplayName = 'SmartM365 Reports',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSendSecurityGroupName = 'SMART-M365-MailSend-Allowed',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSendSecurityGroupAlias = 'smartm365-mailsend-allowed',

    [Parameter()]
    [switch]$DisableTeamsSetup,

    [Parameter()]
    [string]$LogPath = ''
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365RootPath = Find-SmartM365Root -StartPath $PSScriptRoot
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $defaultSetupLogPath = Join-Path -Path $script:SmartM365RootPath -ChildPath ("Data\Tenants\{0}\LOG-ALL\Setup" -f $Tenant)
    if (Test-SmartM365WritableDirectory -Path $defaultSetupLogPath) {
        $LogPath = $defaultSetupLogPath
    }
    else {
        $LogPath = Join-Path -Path $PSScriptRoot -ChildPath ("Output\Tenants\{0}\LOG-ALL\Setup" -f $Tenant)
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SmartM365SetupLogFile = $null
$script:SmartM365SetupTranscriptFile = $null
$script:SmartM365SetupTranscriptStarted = $false
$script:SmartM365ExchangeOnlineConnected = $false

function Format-SmartM365TimestampedLine {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    @([regex]::Split($Message, '\r?\n')) | ForEach-Object {
        "{0} [{1}] {2}" -f $timestamp, $Level, $_
    }
}

function Update-SmartM365TimestampedTranscriptFile {
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $updated = [System.IO.File]::ReadAllLines($Path) | ForEach-Object {
        if ($_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b') {
            $_
        }
        elseif ($_ -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
            "{0} {1}" -f $Matches[1], $Matches[2]
        }
        else {
            "{0} {1}" -f $timestamp, $_
        }
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($Path, [string[]]$updated, $utf8NoBom)
}

function Initialize-SmartM365SetupLogging {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $resolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedLogPath)) {
        New-Item -Path $resolvedLogPath -ItemType Directory -Force -ErrorAction Stop -Confirm:$false | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $baseName = 'SmartM365-Create-AppRegistration'
    $script:SmartM365SetupLogFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}.log" -f $baseName, $timestamp)
    $script:SmartM365SetupTranscriptFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}_Transcript.log" -f $baseName, $timestamp)

    Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value (Format-SmartM365TimestampedLine -Message 'Log started.' -Level 'INFO') -Confirm:$false

    try {
        Start-Transcript -Path $script:SmartM365SetupTranscriptFile -Append -ErrorAction Stop -Confirm:$false | Out-Null
        $script:SmartM365SetupTranscriptStarted = $true
    }
    catch {
        Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value (Format-SmartM365TimestampedLine -Message ("Transcript could not be started: {0}" -f $_.Exception.Message) -Level 'WARN') -Confirm:$false
    }
}

function Close-SmartM365SetupLogging {
    if ($script:SmartM365SetupLogFile) {
        Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value (Format-SmartM365TimestampedLine -Message 'Log finished.' -Level 'INFO') -Confirm:$false
        Write-Information ("[INFO] Log file: {0}" -f $script:SmartM365SetupLogFile) -InformationAction Continue
        if ($script:SmartM365SetupTranscriptFile) {
            Write-Information ("[INFO] Transcript file: {0}" -f $script:SmartM365SetupTranscriptFile) -InformationAction Continue
        }
    }

    if ($script:SmartM365SetupTranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch { Write-Debug ("Stop-Transcript failed: {0}" -f $_.Exception.Message) }
        try { Update-SmartM365TimestampedTranscriptFile -Path $script:SmartM365SetupTranscriptFile } catch { Write-Debug ("Transcript timestamp normalization failed: {0}" -f $_.Exception.Message) }
        $script:SmartM365SetupTranscriptStarted = $false
    }
}

function Write-SmartM365SetupStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][string]$Level = 'INFO'
    )

    if ($script:SmartM365SetupLogFile) {
        $logEntry = Format-SmartM365TimestampedLine -Message $Message -Level $Level
        try { Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value $logEntry -Confirm:$false } catch { Write-Debug ("Log write failed: {0}" -f $_.Exception.Message) }
    }

    switch ($Level) {
        'WARN' { Write-Warning $Message }
        'ERROR' { Write-Error $Message }
        default { Write-Information ("[{0}] {1}" -f $Level, $Message) -InformationAction Continue }
    }
}

function ConvertTo-ODataStringLiteral {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace("'", "''")
}

function Ensure-SmartM365SetupModule {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $availableModule = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $availableModule) {
        Write-SmartM365SetupStatus -Level WARN -Message ("Required setup module '{0}' is missing. Installing with Install-Module -Scope CurrentUser -Force -AllowClobber." -f $Name)
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        $availableModule = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $availableModule) {
            throw ("Required setup module '{0}' could not be resolved after installation." -f $Name)
        }
    }

    Import-Module $Name -ErrorAction Stop
    $loadedModule = Get-Module -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $loadedModule) { $loadedModule = $availableModule }
    Write-SmartM365SetupStatus -Message ("Setup module ready: {0} {1}; Path={2}" -f $loadedModule.Name, $loadedModule.Version, $loadedModule.Path) -Level OK
}

function Import-RequiredGraphModule {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications',
        'ExchangeOnlineManagement'
    )

    foreach ($moduleName in $requiredModules) {
        Ensure-SmartM365SetupModule -Name $moduleName
    }
}

function Import-RequiredExchangeOnlineModule {
    Ensure-SmartM365SetupModule -Name 'ExchangeOnlineManagement'
}

function Disconnect-SmartM365ExistingGraphSession {
    try {
        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $existingContext -and -not [string]::IsNullOrWhiteSpace($existingContext.Account)) {
            Write-SmartM365SetupStatus -Message ("Disconnecting existing Microsoft Graph session for {0}." -f $existingContext.Account)
        }
        else {
            Write-SmartM365SetupStatus -Message 'Clearing any existing Microsoft Graph session.'
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Could not disconnect existing Microsoft Graph session: {0}" -f $_.Exception.Message)
    }
}

function Disconnect-SmartM365ExistingExchangeOnlineSession {
    Import-RequiredExchangeOnlineModule

    try {
        $existingConnections = @()
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $existingConnections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
        }

        if ($existingConnections.Count -gt 0) {
            $connectionNames = @(
                $existingConnections |
                    ForEach-Object {
                        if (-not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName)) {
                            [string]$_.UserPrincipalName
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace([string]$_.Name)) {
                            [string]$_.Name
                        }
                    } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Select-Object -Unique
            )

            if ($connectionNames.Count -gt 0) {
                Write-SmartM365SetupStatus -Message ("Disconnecting existing Exchange Online session(s): {0}." -f ($connectionNames -join ', '))
            }
            else {
                Write-SmartM365SetupStatus -Message 'Disconnecting existing Exchange Online session(s).'
            }
        }
        else {
            Write-SmartM365SetupStatus -Message 'Clearing any existing Exchange Online session.'
        }

        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Could not disconnect existing Exchange Online session: {0}" -f $_.Exception.Message)
    }
}

function Connect-SmartM365GraphSetupSession {
    param(
        [string]$RequestedTenantId,
        [switch]$UseDeviceCode
    )

    $scopes = @(
        'Application.ReadWrite.All',
        'AppRoleAssignment.ReadWrite.All',
        'Channel.Create',
        'Channel.ReadBasic.All',
        'Directory.Read.All',
        'Group.ReadWrite.All',
        'RoleManagement.ReadWrite.Directory',
        'Sites.FullControl.All'
    )

    $connectParams = @{
        Scopes       = $scopes
        NoWelcome    = $true
        ContextScope = 'Process'
    }
    if (-not [string]::IsNullOrWhiteSpace($RequestedTenantId)) {
        $connectParams.TenantId = $RequestedTenantId
    }
    if ($UseDeviceCode) {
        $connectParams.UseDeviceCode = $true
    }

    Write-SmartM365SetupStatus -Message ("Connecting interactively to Microsoft Graph with setup scopes: {0}" -f ($scopes -join ', '))
    Connect-MgGraph @connectParams | Out-Null

    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace($context.TenantId)) {
        throw 'Microsoft Graph connection did not return a tenant context.'
    }
    if ($context.AuthType -ne 'Delegated' -or [string]::IsNullOrWhiteSpace($context.Account)) {
        throw 'This setup script must be run with an interactive delegated Microsoft Graph administrator session.'
    }

    Write-SmartM365SetupStatus -Message ("Connected as {0} in tenant {1}." -f $context.Account, $context.TenantId) -Level OK

    return $context
}

function Connect-SmartM365ExchangeOnlineSetupSession {
    param(
        [string]$UserPrincipalName
    )

    if ($script:SmartM365ExchangeOnlineConnected) {
        return
    }

    Import-RequiredExchangeOnlineModule
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Write-SmartM365SetupStatus -Message 'Connecting interactively to Exchange Online before Microsoft Graph.'
    }
    else {
        Write-SmartM365SetupStatus -Message ("Connecting interactively to Exchange Online as {0} before Microsoft Graph." -f $UserPrincipalName)
    }

    $connectParams = @{
        ShowBanner   = $false
        ShowProgress = $false
        ErrorAction  = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $connectParams.UserPrincipalName = $UserPrincipalName
    }
    if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) {
        $connectParams.DisableWAM = $true
    }

    try {
        Connect-ExchangeOnline @connectParams
    }
    catch {
        $message = [string]$_.Exception.Message
        if ($message -notmatch 'RuntimeBroker|Object reference not set|MSALTokenProvider|WAM') {
            throw
        }

        Write-SmartM365SetupStatus -Level WARN -Message 'Exchange Online WAM/broker authentication failed. Retrying with device code authentication.'
        $deviceParams = @{
            Device       = $true
            ShowBanner   = $false
            ShowProgress = $false
            ErrorAction  = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            $deviceParams.UserPrincipalName = $UserPrincipalName
        }
        Connect-ExchangeOnline @deviceParams
    }
    $script:SmartM365ExchangeOnlineConnected = $true
    Write-SmartM365SetupStatus -Message 'Connected to Exchange Online.' -Level OK
}

function Close-SmartM365ExchangeOnlineSetupSession {
    if (-not $script:SmartM365ExchangeOnlineConnected) {
        return
    }

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-SmartM365SetupStatus -Message 'Disconnected from Exchange Online.' -Level OK
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Exchange Online disconnect failed: {0}" -f $_.Exception.Message)
    }
    finally {
        $script:SmartM365ExchangeOnlineConnected = $false
    }
}

function Get-SmartM365RequiredApiResource {
    param(
        [switch]$SkipExchange,
        [switch]$SkipBroadReadWrite
    )

    $graphPermissions = @(
        'AuditLog.Read.All',
        'BackupRestore-Configuration.Read.All',
        'Directory.Read.All',
        'User.Read.All',
        'Device.Read.All',
        'Group.Read.All',
        'GroupMember.Read.All',
        'Reports.Read.All',
        'CallRecords.Read.All',
        'Sites.Read.All',
        'Team.ReadBasic.All',
        'TeamMember.Read.All',
        'Channel.ReadBasic.All',
        'ChannelMember.Read.All',
        'DeviceManagementApps.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'DeviceManagementScripts.Read.All',
        'DeviceManagementServiceConfig.Read.All',
        'Mail.Send',
        'Sites.Selected'
    )

    if (-not $SkipBroadReadWrite) {
        $graphPermissions += @(
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementConfiguration.ReadWrite.All',
            'DeviceManagementManagedDevices.ReadWrite.All'
        )
    }

    $resources = @(
        [pscustomobject]@{
            Name           = 'Microsoft Graph'
            ResourceAppId  = '00000003-0000-0000-c000-000000000000'
            AppRoleValues  = @($graphPermissions | Sort-Object -Unique)
        }
    )

    if (-not $SkipExchange) {
        $resources += [pscustomobject]@{
            Name           = 'Office 365 Exchange Online'
            ResourceAppId  = '00000002-0000-0ff1-ce00-000000000000'
            AppRoleValues  = @('Exchange.ManageAsApp')
        }
    }

    return $resources
}

function Get-OrCreate-ServicePrincipalByAppId {
    param(
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ResourceAppId
    )

    $filter = "appId eq '$ResourceAppId'"
    $servicePrincipal = @(Get-MgServicePrincipal -Filter $filter -All) | Select-Object -First 1

    if ($null -eq $servicePrincipal) {
        Write-SmartM365SetupStatus -Message ("Creating service principal for {0} ({1})." -f $ResourceName, $ResourceAppId)
        $servicePrincipal = New-MgServicePrincipal -AppId $ResourceAppId
    }

    return $servicePrincipal
}

function Resolve-ApplicationPermission {
    param(
        [Parameter(Mandatory)]$ResourceServicePrincipal,
        [Parameter(Mandatory)][string]$PermissionValue
    )

    $role = @($ResourceServicePrincipal.AppRoles) |
        Where-Object { $_.Value -eq $PermissionValue -and @($_.AllowedMemberTypes) -contains 'Application' } |
        Select-Object -First 1

    if ($null -eq $role) {
        throw ("Application permission '{0}' was not found on resource '{1}'." -f $PermissionValue, $ResourceServicePrincipal.DisplayName)
    }

    return $role
}

function Get-RequiredResourceAccessBlock {
    param([Parameter(Mandatory)]$RequiredApiResource)

    $blocks = foreach ($resource in $RequiredApiResource) {
        $resourceSp = Get-OrCreate-ServicePrincipalByAppId -ResourceName $resource.Name -ResourceAppId $resource.ResourceAppId
        $access = foreach ($permission in $resource.AppRoleValues) {
            $role = Resolve-ApplicationPermission -ResourceServicePrincipal $resourceSp -PermissionValue $permission
            @{
                id   = $role.Id
                type = 'Role'
            }
        }

        @{
            resourceAppId  = $resource.ResourceAppId
            resourceAccess = @($access)
        }
    }

    return @($blocks)
}

function Merge-RequiredResourceAccess {
    param(
        [AllowNull()]$ExistingAccess,
        [Parameter(Mandatory)][array]$RequiredAccess
    )

    $merged = @{}

    foreach ($block in @($ExistingAccess)) {
        if ($null -eq $block -or [string]::IsNullOrWhiteSpace($block.ResourceAppId)) { continue }
        if (-not $merged.ContainsKey($block.ResourceAppId)) {
            $merged[$block.ResourceAppId] = New-Object 'System.Collections.Generic.HashSet[string]'
        }
        foreach ($access in @($block.ResourceAccess)) {
            if ($null -ne $access.Id) {
                [void]$merged[$block.ResourceAppId].Add(("{0}|{1}" -f $access.Id, $access.Type))
            }
        }
    }

    foreach ($block in @($RequiredAccess)) {
        if (-not $merged.ContainsKey($block.resourceAppId)) {
            $merged[$block.resourceAppId] = New-Object 'System.Collections.Generic.HashSet[string]'
        }
        foreach ($access in @($block.resourceAccess)) {
            [void]$merged[$block.resourceAppId].Add(("{0}|{1}" -f $access.id, $access.type))
        }
    }

    return @(
        foreach ($resourceAppId in ($merged.Keys | Sort-Object)) {
            @{
                resourceAppId  = $resourceAppId
                resourceAccess = @(
                    foreach ($entry in ($merged[$resourceAppId] | Sort-Object)) {
                        $parts = $entry -split '\|', 2
                        @{
                            id   = $parts[0]
                            type = $parts[1]
                        }
                    }
                )
            }
        }
    )
}

function Get-SmartM365RequiredResourceAccessWithoutPermission {
    param(
        [Parameter(Mandatory)][array]$RequiredAccess,
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ResourceAppId,
        [Parameter(Mandatory)][string[]]$PermissionValues
    )

    $resourceSp = Get-OrCreate-ServicePrincipalByAppId -ResourceName $ResourceName -ResourceAppId $ResourceAppId
    $roleIdsToRemove = @(
        foreach ($permissionValue in $PermissionValues) {
            (Resolve-ApplicationPermission -ResourceServicePrincipal $resourceSp -PermissionValue $permissionValue).Id
        }
    )

    $filtered = foreach ($block in @($RequiredAccess)) {
        $blockResourceAppId = if ($null -ne $block.resourceAppId) { $block.resourceAppId } else { $block.ResourceAppId }
        $blockResourceAccess = if ($null -ne $block.resourceAccess) { $block.resourceAccess } else { $block.ResourceAccess }

        $resourceAccess = @($blockResourceAccess | Where-Object {
            -not ($blockResourceAppId -eq $ResourceAppId -and $roleIdsToRemove -contains $_.id)
        })

        if ($resourceAccess.Count -gt 0) {
            @{
                resourceAppId  = $blockResourceAppId
                resourceAccess = $resourceAccess
            }
        }
    }

    return @($filtered)
}

function ConvertTo-SmartM365RequiredResourceAccessPayload {
    param(
        [AllowNull()]$RequiredAccess
    )

    $payload = foreach ($block in @($RequiredAccess)) {
        if ($null -eq $block) {
            continue
        }

        $resourceAppId = if ($null -ne $block.resourceAppId) { $block.resourceAppId } else { $block.ResourceAppId }
        if ([string]::IsNullOrWhiteSpace([string]$resourceAppId)) {
            continue
        }

        $rawResourceAccess = if ($null -ne $block.resourceAccess) { $block.resourceAccess } else { $block.ResourceAccess }
        $resourceAccess = @(
            foreach ($access in @($rawResourceAccess)) {
                if ($null -eq $access) {
                    continue
                }

                $id = if ($null -ne $access.id) { $access.id } else { $access.Id }
                $type = if ($null -ne $access.type) { $access.type } else { $access.Type }

                if ($null -eq $id -or [string]::IsNullOrWhiteSpace([string]$type)) {
                    continue
                }

                @{
                    id   = $id
                    type = [string]$type
                }
            }
        )

        if ($resourceAccess.Count -eq 0) {
            continue
        }

        @{
            resourceAppId  = [string]$resourceAppId
            resourceAccess = $resourceAccess
        }
    }

    return @($payload)
}

function Get-SmartM365Certificate {
    param(
        [string]$Thumbprint,
        [Parameter(Mandatory)][string]$AppDisplayName,
        [Parameter(Mandatory)][int]$ValidityYears,
        [AllowNull()]$Application
    )

    if (-not [string]::IsNullOrWhiteSpace($Thumbprint)) {
        $normalizedThumbprint = $Thumbprint.Replace(' ', '').ToUpperInvariant()
        $cert = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') |
            ForEach-Object { Get-ChildItem -Path $_ -ErrorAction SilentlyContinue } |
            Where-Object { $_.Thumbprint -eq $normalizedThumbprint } |
            Select-Object -First 1

        if ($null -eq $cert) {
            throw "Certificate thumbprint '$Thumbprint' was not found in CurrentUser\My or LocalMachine\My."
        }
        if (-not $cert.HasPrivateKey) {
            throw "Certificate '$Thumbprint' was found but does not contain a private key."
        }
        return $cert
    }

    if ($null -ne $Application) {
        $appKeyIds = @(
            @($Application.KeyCredentials) | ForEach-Object {
            if ($null -eq $_.CustomKeyIdentifier) {
                $null
            }
            elseif ($_.CustomKeyIdentifier -is [byte[]]) {
                [System.Convert]::ToBase64String($_.CustomKeyIdentifier)
            }
            else {
                [string]$_.CustomKeyIdentifier
            }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($appKeyIds.Count -gt 0) {
            $localCertificate = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') |
                ForEach-Object { Get-ChildItem -Path $_ -ErrorAction SilentlyContinue } |
                Where-Object {
                    $_.HasPrivateKey -and
                    $_.NotAfter -gt (Get-Date) -and
                    ($appKeyIds -contains [System.Convert]::ToBase64String($_.GetCertHash()))
                } |
                Sort-Object NotAfter -Descending |
                Select-Object -First 1

            if ($null -ne $localCertificate) {
                Write-SmartM365SetupStatus -Message ("Reusing existing local certificate already attached to the app registration: {0}" -f $localCertificate.Thumbprint) -Level OK
                return $localCertificate
            }
        }
    }

    if ($WhatIfPreference) {
        throw 'WhatIf mode without -CertificateThumbprint would require creating a local self-signed certificate. Provide an existing -CertificateThumbprint for a side-effect-free dry-run.'
    }

    $subject = "CN=$AppDisplayName"
    Write-SmartM365SetupStatus -Message ("Creating self-signed certificate '{0}' in CurrentUser\My." -f $subject)
    return New-SelfSignedCertificate `
        -Subject $subject `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeySpec Signature `
        -KeyExportPolicy Exportable `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($ValidityYears)
}

function Get-KeyCredentialFromCertificate {
    param([Parameter(Mandatory)]$Certificate)

    $keyCredential = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]::new()
    $keyCredential.CustomKeyIdentifier = $Certificate.GetCertHash()
    $keyCredential.DisplayName = $Certificate.Subject
    $keyCredential.EndDateTime = $Certificate.NotAfter
    $keyCredential.Key = $Certificate.RawData
    $keyCredential.StartDateTime = $Certificate.NotBefore
    $keyCredential.Type = 'AsymmetricX509Cert'
    $keyCredential.Usage = 'Verify'

    return $keyCredential
}

function Add-CertificateToApplication {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)]$Certificate
    )

    $existingKeys = @($Application.KeyCredentials)
    $certificateKeyId = [System.Convert]::ToBase64String($Certificate.GetCertHash())
    $alreadyUploaded = $existingKeys | Where-Object {
        if ($null -eq $_.CustomKeyIdentifier) {
            $false
        }
        elseif ($_.CustomKeyIdentifier -is [byte[]]) {
            [System.Convert]::ToBase64String($_.CustomKeyIdentifier) -eq $certificateKeyId
        }
        else {
            [string]$_.CustomKeyIdentifier -eq $certificateKeyId
        }
    }

    if ($alreadyUploaded) {
        Write-SmartM365SetupStatus -Message 'Certificate public key is already present on the app registration.' -Level OK
        return @($existingKeys)
    }

    return @($existingKeys + (Get-KeyCredentialFromCertificate -Certificate $Certificate))
}

function Test-ApplicationHasCertificate {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)]$Certificate
    )

    $certificateKeyId = [System.Convert]::ToBase64String($Certificate.GetCertHash())
    $matchingKey = @($Application.KeyCredentials) | Where-Object {
        if ($null -eq $_.CustomKeyIdentifier) {
            $false
        }
        elseif ($_.CustomKeyIdentifier -is [byte[]]) {
            [System.Convert]::ToBase64String($_.CustomKeyIdentifier) -eq $certificateKeyId
        }
        else {
            [string]$_.CustomKeyIdentifier -eq $certificateKeyId
        }
    } | Select-Object -First 1

    return ($null -ne $matchingKey)
}

function Grant-SmartM365ApplicationPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$ApplicationServicePrincipal,
        [Parameter(Mandatory)]$RequiredApiResource
    )

    $existingAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ApplicationServicePrincipal.Id -All)

    foreach ($resource in $RequiredApiResource) {
        $resourceSp = Get-OrCreate-ServicePrincipalByAppId -ResourceName $resource.Name -ResourceAppId $resource.ResourceAppId

        foreach ($permission in $resource.AppRoleValues) {
            $role = Resolve-ApplicationPermission -ResourceServicePrincipal $resourceSp -PermissionValue $permission
            $existing = $existingAssignments | Where-Object { $_.ResourceId -eq $resourceSp.Id -and $_.AppRoleId -eq $role.Id } | Select-Object -First 1

            if ($existing) {
                Write-SmartM365SetupStatus -Message ("Admin consent already granted: {0} / {1}" -f $resource.Name, $permission) -Level OK
                continue
            }

            if ($PSCmdlet.ShouldProcess(("{0} / {1}" -f $resource.Name, $permission), 'Grant admin consent')) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ApplicationServicePrincipal.Id -BodyParameter @{
                    principalId = $ApplicationServicePrincipal.Id
                    resourceId  = $resourceSp.Id
                    appRoleId   = $role.Id
                } | Out-Null
                Write-SmartM365SetupStatus -Message ("Granted admin consent: {0} / {1}" -f $resource.Name, $permission) -Level OK
            }
        }
    }
}

function Revoke-SmartM365ApplicationPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$ApplicationServicePrincipal,
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ResourceAppId,
        [Parameter(Mandatory)][string[]]$PermissionValues
    )

    $resourceSp = Get-OrCreate-ServicePrincipalByAppId -ResourceName $ResourceName -ResourceAppId $ResourceAppId
    $existingAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ApplicationServicePrincipal.Id -All)

    foreach ($permission in $PermissionValues) {
        $role = Resolve-ApplicationPermission -ResourceServicePrincipal $resourceSp -PermissionValue $permission
        $assignmentsToRemove = @($existingAssignments | Where-Object { $_.ResourceId -eq $resourceSp.Id -and $_.AppRoleId -eq $role.Id })

        if ($assignmentsToRemove.Count -eq 0) {
            Write-SmartM365SetupStatus -Message ("Broad admin consent not present: {0} / {1}" -f $ResourceName, $permission) -Level OK
            continue
        }

        foreach ($assignment in $assignmentsToRemove) {
            if ($PSCmdlet.ShouldProcess(("{0} / {1}" -f $ResourceName, $permission), 'Revoke broad admin consent')) {
                Remove-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $ApplicationServicePrincipal.Id `
                    -AppRoleAssignmentId $assignment.Id `
                    -ErrorAction Stop
                Write-SmartM365SetupStatus -Message ("Revoked broad admin consent: {0} / {1}" -f $ResourceName, $permission) -Level OK
            }
        }
    }
}

function Grant-SmartM365ExchangeOnlineDirectoryRole {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$ApplicationServicePrincipal,
        [Parameter()][ValidateNotNullOrEmpty()][string]$RoleDisplayName = 'Global Reader'
    )

    $escapedRoleDisplayName = ConvertTo-ODataStringLiteral -Value $RoleDisplayName
    $encodedFilter = [System.Uri]::EscapeDataString("displayName eq '$escapedRoleDisplayName'")
    $roleDefinitionsResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$filter=$encodedFilter" `
        -OutputType PSObject `
        -ErrorAction Stop
    $roleDefinition = @($roleDefinitionsResponse.value) | Select-Object -First 1

    if ($null -eq $roleDefinition -or [string]::IsNullOrWhiteSpace([string]$roleDefinition.id)) {
        throw ("Microsoft Entra directory role '{0}' was not found." -f $RoleDisplayName)
    }

    $encodedAssignmentFilter = [System.Uri]::EscapeDataString("principalId eq '$($ApplicationServicePrincipal.Id)' and roleDefinitionId eq '$($roleDefinition.id)'")
    $existingAssignmentsResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=$encodedAssignmentFilter" `
        -OutputType PSObject `
        -ErrorAction Stop
    $existingAssignment = @($existingAssignmentsResponse.value) | Where-Object {
        $_.directoryScopeId -eq '/' -or [string]::IsNullOrWhiteSpace([string]$_.directoryScopeId)
    } | Select-Object -First 1

    if ($existingAssignment) {
        Write-SmartM365SetupStatus -Message ("Microsoft Entra role already assigned to app service principal: {0}." -f $RoleDisplayName) -Level OK
        return $RoleDisplayName
    }

    $body = @{
        principalId      = $ApplicationServicePrincipal.Id
        roleDefinitionId = $roleDefinition.id
        directoryScopeId = '/'
    }

    if ($PSCmdlet.ShouldProcess(("Service principal {0}" -f $ApplicationServicePrincipal.Id), ("Assign Microsoft Entra role '{0}'" -f $RoleDisplayName))) {
        $null = Invoke-MgGraphRequest `
            -Method POST `
            -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' `
            -Body ($body | ConvertTo-Json -Depth 5) `
            -ContentType 'application/json' `
            -OutputType PSObject `
            -ErrorAction Stop
        Write-SmartM365SetupStatus -Message ("Assigned Microsoft Entra role '{0}' to app service principal." -f $RoleDisplayName) -Level OK
    }

    return $RoleDisplayName
}

function Revoke-SmartM365ExchangeOnlineDirectoryRole {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$ApplicationServicePrincipal,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$RoleDisplayName
    )

    $escapedRoleDisplayName = ConvertTo-ODataStringLiteral -Value $RoleDisplayName
    $encodedFilter = [System.Uri]::EscapeDataString("displayName eq '$escapedRoleDisplayName'")
    $roleDefinitionsResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$filter=$encodedFilter" `
        -OutputType PSObject `
        -ErrorAction Stop
    $roleDefinition = @($roleDefinitionsResponse.value) | Select-Object -First 1

    if ($null -eq $roleDefinition -or [string]::IsNullOrWhiteSpace([string]$roleDefinition.id)) {
        Write-SmartM365SetupStatus -Level WARN -Message ("Microsoft Entra directory role '{0}' was not found; nothing to revoke." -f $RoleDisplayName)
        return
    }

    $encodedAssignmentFilter = [System.Uri]::EscapeDataString("principalId eq '$($ApplicationServicePrincipal.Id)' and roleDefinitionId eq '$($roleDefinition.id)'")
    $existingAssignmentsResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=$encodedAssignmentFilter" `
        -OutputType PSObject `
        -ErrorAction Stop
    $existingAssignments = @(
        @($existingAssignmentsResponse.value) | Where-Object {
            $_.directoryScopeId -eq '/' -or [string]::IsNullOrWhiteSpace([string]$_.directoryScopeId)
        }
    )

    if ($existingAssignments.Count -eq 0) {
        Write-SmartM365SetupStatus -Message ("Microsoft Entra role not present on app service principal: {0}." -f $RoleDisplayName) -Level OK
        return
    }

    foreach ($assignment in $existingAssignments) {
        $encodedAssignmentId = [System.Uri]::EscapeDataString([string]$assignment.id)
        if ($PSCmdlet.ShouldProcess(("Service principal {0}" -f $ApplicationServicePrincipal.Id), ("Remove Microsoft Entra role '{0}'" -f $RoleDisplayName))) {
            Invoke-MgGraphRequest `
                -Method DELETE `
                -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments/$encodedAssignmentId" `
                -ErrorAction Stop | Out-Null
            Write-SmartM365SetupStatus -Message ("Removed Microsoft Entra role '{0}' from app service principal." -f $RoleDisplayName) -Level OK
        }
    }
}

function Test-SmartM365GraphNotFound {
    param([Parameter(Mandatory)]$ErrorRecord)

    $message = [string]$ErrorRecord.Exception.Message
    return ($message -match '\b404\b' -or $message -match 'NotFound' -or $message -match 'Not Found')
}

function Get-SmartM365SetupUser {
    param([Parameter(Mandatory)][string]$UserPrincipalName)

    $encodedUser = [System.Uri]::EscapeDataString($UserPrincipalName)
    return Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/users/${encodedUser}?`$select=id,userPrincipalName,displayName" `
        -OutputType PSObject `
        -ErrorAction Stop
}

function Get-SmartM365GroupByDisplayName {
    param([Parameter(Mandatory)][string]$GroupDisplayName)

    $escapedName = ConvertTo-ODataStringLiteral -Value $GroupDisplayName
    $response = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escapedName'&`$select=id,displayName,mailNickname,resourceProvisioningOptions,webUrl" `
        -OutputType PSObject `
        -ErrorAction Stop

    return @($response.value)
}

function Invoke-SmartM365TeamGroupCreation {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][string]$MailNickname,
        [Parameter(Mandatory)][string]$OwnerUserId
    )

    $ownerUrl = "https://graph.microsoft.com/v1.0/users/$OwnerUserId"
    $body = @{
        description          = 'SmartM365 shared workspace for automation exports.'
        displayName          = $GroupDisplayName
        groupTypes           = @('Unified')
        mailEnabled          = $true
        mailNickname         = $MailNickname
        securityEnabled      = $false
        visibility           = 'Private'
        'owners@odata.bind'  = @($ownerUrl)
        'members@odata.bind' = @($ownerUrl)
    }

    Write-SmartM365SetupStatus -Message ("Creating Microsoft 365 group for Teams team '{0}'." -f $GroupDisplayName)
    if ($PSCmdlet.ShouldProcess($GroupDisplayName, 'Create Microsoft 365 group for Teams')) {
        return Invoke-MgGraphRequest `
            -Method POST `
            -Uri 'https://graph.microsoft.com/v1.0/groups' `
            -Body ($body | ConvertTo-Json -Depth 8) `
            -ContentType 'application/json' `
            -OutputType PSObject `
            -ErrorAction Stop
    }
}

function Test-SmartM365TeamPresent {
    param([Parameter(Mandatory)][string]$GroupId)

    try {
        $null = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/teams/$GroupId" `
            -OutputType PSObject `
            -ErrorAction Stop
        return $true
    }
    catch {
        if (Test-SmartM365GraphNotFound -ErrorRecord $_) {
            return $false
        }
        throw
    }
}

function Invoke-SmartM365TeamCreationFromGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$GroupId)

    $body = @{
        memberSettings = @{
            allowCreatePrivateChannels = $false
            allowCreateUpdateChannels  = $true
        }
        messagingSettings = @{
            allowUserDeleteMessages = $true
            allowUserEditMessages   = $true
        }
        funSettings = @{
            allowCustomMemes = $false
            allowGiphy       = $true
            giphyContentRating = 'moderate'
        }
    }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            Write-SmartM365SetupStatus -Message ("Creating Teams team from group {0} (attempt {1}/6)." -f $GroupId, $attempt)
            if ($PSCmdlet.ShouldProcess($GroupId, 'Create Teams team from Microsoft 365 group')) {
                $null = Invoke-MgGraphRequest `
                    -Method PUT `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/team" `
                    -Body ($body | ConvertTo-Json -Depth 8) `
                    -ContentType 'application/json' `
                    -OutputType PSObject `
                    -ErrorAction Stop
            }
            Write-SmartM365SetupStatus -Message 'Teams team created.' -Level OK
            return
        }
        catch {
            if ($attempt -ge 6) {
                throw
            }
            Write-SmartM365SetupStatus -Level WARN -Message ("Teams creation is not ready yet: {0}. Retrying in 10 seconds." -f $_.Exception.Message)
            Start-Sleep -Seconds 10
        }
    }
}

function Get-SmartM365TeamChannelByDisplayName {
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelDisplayName
    )

    $response = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/channels" `
        -OutputType PSObject `
        -ErrorAction Stop

    return @($response.value) |
        Where-Object { $_.displayName -eq $ChannelDisplayName } |
        Select-Object -First 1
}

function Get-OrCreate-SmartM365TeamChannel {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelDisplayName,
        [Parameter(Mandatory)][string]$Description
    )

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            $existingChannel = Get-SmartM365TeamChannelByDisplayName -TeamId $TeamId -ChannelDisplayName $ChannelDisplayName
            if ($existingChannel) {
                Write-SmartM365SetupStatus -Message ("Teams channel already exists: {0}." -f $ChannelDisplayName) -Level OK
                return $existingChannel
            }

            $body = @{
                displayName    = $ChannelDisplayName
                description    = $Description
                membershipType = 'standard'
            }

            if ($PSCmdlet.ShouldProcess(("{0} / {1}" -f $TeamId, $ChannelDisplayName), 'Create Teams standard channel')) {
                $channel = Invoke-MgGraphRequest `
                    -Method POST `
                    -Uri "https://graph.microsoft.com/v1.0/teams/$TeamId/channels" `
                    -Body ($body | ConvertTo-Json -Depth 5) `
                    -ContentType 'application/json' `
                    -OutputType PSObject `
                    -ErrorAction Stop
                Write-SmartM365SetupStatus -Message ("Created Teams channel '{0}'." -f $ChannelDisplayName) -Level OK
                return $channel
            }

            return $null
        }
        catch {
            if ($attempt -ge 6) {
                throw
            }
            Write-SmartM365SetupStatus -Level WARN -Message ("Teams channel '{0}' is not ready yet: {1}. Retrying in 10 seconds." -f $ChannelDisplayName, $_.Exception.Message)
            Start-Sleep -Seconds 10
        }
    }
}

function Wait-SmartM365TeamSharePointSite {
    param([Parameter(Mandatory)][string]$GroupId)

    for ($attempt = 1; $attempt -le 30; $attempt++) {
        try {
            $site = Invoke-MgGraphRequest `
                -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/sites/root?`$select=id,webUrl,displayName" `
                -OutputType PSObject `
                -ErrorAction Stop

            if ($site -and -not [string]::IsNullOrWhiteSpace($site.webUrl)) {
                return $site
            }
        }
        catch {
            if (-not (Test-SmartM365GraphNotFound -ErrorRecord $_)) {
                Write-SmartM365SetupStatus -Level WARN -Message ("SharePoint site lookup failed: {0}" -f $_.Exception.Message)
            }
        }

        Write-SmartM365SetupStatus -Message ("Waiting for Teams SharePoint site provisioning (attempt {0}/30)." -f $attempt)
        Start-Sleep -Seconds 10
    }

    throw "The SharePoint site for group '$GroupId' was not available after 5 minutes."
}

function Get-SmartM365GroupDriveName {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$DefaultLibraryDisplayName
    )

    try {
        $drive = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/drive?`$select=id,name,webUrl" `
            -OutputType PSObject `
            -ErrorAction Stop

        if ($drive -and -not [string]::IsNullOrWhiteSpace($drive.name)) {
            return $drive.name
        }
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Could not resolve default document library name, using '{0}': {1}" -f $DefaultLibraryDisplayName, $_.Exception.Message)
    }

    return $DefaultLibraryDisplayName
}

function Get-OrCreate-SmartM365TeamsWorkspace {
    param(
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][string]$MailNickname,
        [Parameter(Mandatory)][string]$OwnerUserPrincipalName,
        [Parameter(Mandatory)][string]$TargetFolderPath,
        [Parameter(Mandatory)][string]$DefaultLibraryDisplayName
    )

    $setupUser = Get-SmartM365SetupUser -UserPrincipalName $OwnerUserPrincipalName
    $groups = @(Get-SmartM365GroupByDisplayName -GroupDisplayName $GroupDisplayName)

    if ($groups.Count -gt 1) {
        $matchingNickname = @($groups | Where-Object { $_.mailNickname -eq $MailNickname }) | Select-Object -First 1
        if ($null -eq $matchingNickname) {
            throw "Multiple Microsoft 365 groups named '$GroupDisplayName' were found. Rename duplicates or specify a unique -TeamDisplayName."
        }
        $group = $matchingNickname
    }
    elseif ($groups.Count -eq 1) {
        $group = $groups[0]
        Write-SmartM365SetupStatus -Message ("Reusing existing Microsoft 365 group '{0}' ({1})." -f $group.displayName, $group.id) -Level OK
    }
    else {
        $group = Invoke-SmartM365TeamGroupCreation -GroupDisplayName $GroupDisplayName -MailNickname $MailNickname -OwnerUserId $setupUser.id
        Write-SmartM365SetupStatus -Message ("Created Microsoft 365 group '{0}' ({1})." -f $group.displayName, $group.id) -Level OK
    }

    if (Test-SmartM365TeamPresent -GroupId $group.id) {
        Write-SmartM365SetupStatus -Message ("Teams team already exists for '{0}'." -f $GroupDisplayName) -Level OK
    }
    else {
        Invoke-SmartM365TeamCreationFromGroup -GroupId $group.id
    }

    $alertsChannel = Get-OrCreate-SmartM365TeamChannel `
        -TeamId $group.id `
        -ChannelDisplayName 'Alerts' `
        -Description 'SmartM365 script error and failure notifications.'
    $infosChannel = Get-OrCreate-SmartM365TeamChannel `
        -TeamId $group.id `
        -ChannelDisplayName 'Infos' `
        -Description 'SmartM365 successful completion and informational notifications.'

    $site = Wait-SmartM365TeamSharePointSite -GroupId $group.id
    $siteUri = [System.Uri]$site.webUrl
    $libraryDisplayName = Get-SmartM365GroupDriveName -GroupId $group.id -DefaultLibraryDisplayName $DefaultLibraryDisplayName

    return [pscustomobject]@{
        TeamId                       = $group.id
        TeamDisplayName              = $GroupDisplayName
        AlertsChannelId              = if ($alertsChannel) { $alertsChannel.id } else { $null }
        AlertsChannelDisplayName     = 'Alerts'
        InfosChannelId               = if ($infosChannel) { $infosChannel.id } else { $null }
        InfosChannelDisplayName      = 'Infos'
        SharePointSiteId             = $site.id
        SharePointSiteHostname       = $siteUri.Host
        SharePointSitePath           = $siteUri.AbsolutePath
        SharePointLibraryDisplayName = $libraryDisplayName
        SharePointTargetFolderPath   = $TargetFolderPath
    }
}

function Get-SmartM365ApplicationIdFromSitePermission {
    param([Parameter(Mandatory)]$Permission)

    $applicationIds = @()

    $grantedToIdentitiesV2 = if ($null -ne $Permission.PSObject.Properties['grantedToIdentitiesV2']) { $Permission.grantedToIdentitiesV2 } else { @() }
    $grantedToIdentities = if ($null -ne $Permission.PSObject.Properties['grantedToIdentities']) { $Permission.grantedToIdentities } else { @() }
    $grantedToV2 = if ($null -ne $Permission.PSObject.Properties['grantedToV2']) { $Permission.grantedToV2 } else { $null }
    $grantedTo = if ($null -ne $Permission.PSObject.Properties['grantedTo']) { $Permission.grantedTo } else { $null }

    foreach ($identity in @($grantedToIdentitiesV2)) {
        if ($identity.application -and -not [string]::IsNullOrWhiteSpace([string]$identity.application.id)) {
            $applicationIds += [string]$identity.application.id
        }
    }

    foreach ($identity in @($grantedToIdentities)) {
        if ($identity.application -and -not [string]::IsNullOrWhiteSpace([string]$identity.application.id)) {
            $applicationIds += [string]$identity.application.id
        }
    }

    if ($grantedToV2 -and $grantedToV2.application -and -not [string]::IsNullOrWhiteSpace([string]$grantedToV2.application.id)) {
        $applicationIds += [string]$grantedToV2.application.id
    }

    if ($grantedTo -and $grantedTo.application -and -not [string]::IsNullOrWhiteSpace([string]$grantedTo.application.id)) {
        $applicationIds += [string]$grantedTo.application.id
    }

    return @($applicationIds | Sort-Object -Unique)
}

function Grant-SmartM365SelectedSitePermission {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SiteId,
        [Parameter(Mandatory)][string]$ApplicationAppId,
        [Parameter(Mandatory)][string]$ApplicationDisplayName,
        [Parameter()][ValidateSet('read', 'write', 'manage', 'fullcontrol')][string]$Role = 'write'
    )

    $encodedSiteId = [System.Uri]::EscapeDataString($SiteId)
    $permissionsResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/permissions" `
        -OutputType PSObject `
        -ErrorAction Stop

    $existingPermission = @($permissionsResponse.value) | Where-Object {
        @(Get-SmartM365ApplicationIdFromSitePermission -Permission $_) -contains $ApplicationAppId
    } | Select-Object -First 1

    if ($existingPermission) {
        $roles = @($existingPermission.roles)
        if ($roles -contains $Role -or $roles -contains 'fullcontrol' -or ($Role -eq 'write' -and $roles -contains 'manage')) {
            Write-SmartM365SetupStatus -Message ("Sites.Selected permission already grants '{0}' or higher to app '{1}' on site '{2}'." -f $Role, $ApplicationDisplayName, $SiteId) -Level OK
            return
        }

        Write-SmartM365SetupStatus -Level WARN -Message ("App '{0}' already has selected-site permission on '{1}', but roles are '{2}'. Create/update manually if a higher role is required." -f $ApplicationDisplayName, $SiteId, ($roles -join ', '))
        return
    }

    $body = @{
        roles               = @($Role)
        grantedToIdentities = @(
            @{
                application = @{
                    id          = $ApplicationAppId
                    displayName = $ApplicationDisplayName
                }
            }
        )
    }

    if ($PSCmdlet.ShouldProcess($SiteId, ("Grant Sites.Selected '{0}' permission to app '{1}'" -f $Role, $ApplicationDisplayName))) {
        $null = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/permissions" `
            -Body ($body | ConvertTo-Json -Depth 8) `
            -ContentType 'application/json' `
            -OutputType PSObject `
            -ErrorAction Stop
        Write-SmartM365SetupStatus -Message ("Granted Sites.Selected '{0}' permission to app '{1}' on site '{2}'." -f $Role, $ApplicationDisplayName, $SiteId) -Level OK
    }
}

function Get-SmartM365ExchangeDefaultAcceptedDomain {
    $defaultDomain = @(Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.Default -eq $true }) | Select-Object -First 1
    if ($null -eq $defaultDomain) {
        $defaultDomain = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1)
    }

    if ($null -eq $defaultDomain -or [string]::IsNullOrWhiteSpace([string]$defaultDomain.DomainName)) {
        throw 'Could not resolve an Exchange Online accepted domain for SmartM365 mail setup.'
    }

    return [string]$defaultDomain.DomainName
}

function Resolve-SmartM365MailAddress {
    param(
        [string]$ConfiguredAddress,
        [Parameter(Mandatory)][string]$DefaultAlias
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredAddress)) {
        return $ConfiguredAddress.Trim()
    }

    $domain = Get-SmartM365ExchangeDefaultAcceptedDomain
    return ('{0}@{1}' -f $DefaultAlias, $domain)
}

function Get-OrCreate-SmartM365SenderMailbox {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SenderAddress,
        [Parameter(Mandatory)][string]$SenderDisplayName
    )

    $mailbox = Get-Mailbox -Identity $SenderAddress -ErrorAction SilentlyContinue
    if ($null -ne $mailbox) {
        if ($mailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            Write-SmartM365SetupStatus -Level WARN -Message ("Reusing existing sender mailbox '{0}', but it is '{1}' rather than a SharedMailbox." -f $SenderAddress, $mailbox.RecipientTypeDetails)
        }
        else {
            Write-SmartM365SetupStatus -Message ("Reusing existing SmartM365 sender shared mailbox '{0}'." -f $SenderAddress) -Level OK
        }
        return $mailbox
    }

    $alias = ($SenderAddress -split '@', 2)[0]
    if ($PSCmdlet.ShouldProcess($SenderAddress, 'Create SmartM365 sender shared mailbox')) {
        $mailbox = New-Mailbox `
            -Shared `
            -Name $SenderDisplayName `
            -DisplayName $SenderDisplayName `
            -Alias $alias `
            -PrimarySmtpAddress $SenderAddress `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Created SmartM365 sender shared mailbox '{0}'." -f $SenderAddress) -Level OK
    }

    return $mailbox
}

function Get-OrCreate-SmartM365MailSendSecurityGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$GroupName,
        [Parameter(Mandatory)][string]$GroupAlias,
        [Parameter(Mandatory)][string]$GroupAddress
    )

    $group = Get-DistributionGroup -Identity $GroupAddress -ErrorAction SilentlyContinue
    if ($null -eq $group) {
        $group = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Where-Object { $_.Alias -eq $GroupAlias -or $_.DisplayName -eq $GroupName }) | Select-Object -First 1
    }

    if ($null -ne $group) {
        Write-SmartM365SetupStatus -Message ("Reusing existing Mail.Send scope group '{0}' ({1})." -f $group.DisplayName, $group.PrimarySmtpAddress) -Level OK
        return $group
    }

    if ($PSCmdlet.ShouldProcess($GroupAddress, 'Create SmartM365 Mail.Send mail-enabled security group')) {
        $group = New-DistributionGroup `
            -Name $GroupName `
            -Alias $GroupAlias `
            -PrimarySmtpAddress $GroupAddress `
            -Type Security `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Created Mail.Send scope group '{0}' ({1})." -f $GroupName, $GroupAddress) -Level OK
    }

    return $group
}

function Add-SmartM365SenderMailboxToScopeGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][string]$SenderAddress
    )

    $members = @(Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited -ErrorAction Stop)
    $existingMember = $members | Where-Object {
        ([string]$_.PrimarySmtpAddress).Equals($SenderAddress, [System.StringComparison]::OrdinalIgnoreCase) -or
        ([string]$_.WindowsEmailAddress).Equals($SenderAddress, [System.StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1

    if ($null -ne $existingMember) {
        Write-SmartM365SetupStatus -Message ("Sender mailbox '{0}' is already a member of '{1}'." -f $SenderAddress, $Group.DisplayName) -Level OK
        return
    }

    if ($PSCmdlet.ShouldProcess(("{0} -> {1}" -f $SenderAddress, $Group.DisplayName), 'Add SmartM365 sender mailbox to Mail.Send scope group')) {
        Add-DistributionGroupMember `
            -Identity $Group.Identity `
            -Member $SenderAddress `
            -BypassSecurityGroupManagerCheck `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Added sender mailbox '{0}' to Mail.Send scope group '{1}'." -f $SenderAddress, $Group.DisplayName) -Level OK
    }
}

function Set-SmartM365MailSendApplicationAccessPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][string]$GroupAddress
    )

    $scopeIdentifiers = @(
        $GroupAddress,
        [string]$Group.Identity,
        [string]$Group.Name,
        [string]$Group.DisplayName,
        [string]$Group.Alias,
        [string]$Group.PrimarySmtpAddress,
        [string]$Group.ExternalDirectoryObjectId,
        [string]$Group.Guid
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

    $policies = @(Get-SmartM365ApplicationAccessPolicy)
    $matchingPolicy = $policies | Where-Object {
        $policy = $_
        @($_.AppId) -contains $AppId -and
        $_.AccessRight -eq 'RestrictAccess' -and
        ($scopeIdentifiers | Where-Object { $_.Equals([string]$policy.ScopeName, [System.StringComparison]::OrdinalIgnoreCase) })
    } | Select-Object -First 1

    if ($null -ne $matchingPolicy) {
        Write-SmartM365SetupStatus -Message ("Application Access Policy already scopes Mail.Send for app {0} to '{1}'." -f $AppId, $matchingPolicy.ScopeName) -Level OK
        return $matchingPolicy
    }

    $otherRestrictPolicy = $policies | Where-Object {
        @($_.AppId) -contains $AppId -and $_.AccessRight -eq 'RestrictAccess'
    } | Select-Object -First 1
    if ($null -ne $otherRestrictPolicy) {
        Write-SmartM365SetupStatus -Level WARN -Message ("App {0} already has an Application Access Policy scoped to '{1}'. A second policy will be added for '{2}'." -f $AppId, $otherRestrictPolicy.ScopeName, $GroupAddress)
    }

    if ($PSCmdlet.ShouldProcess($AppId, ("Create Application Access Policy scoped to {0}" -f $GroupAddress))) {
        $policy = New-ApplicationAccessPolicy `
            -AccessRight RestrictAccess `
            -AppId $AppId `
            -PolicyScopeGroupId $GroupAddress `
            -Description 'Restrict SmartM365 Graph Mail.Send to approved SmartM365 sender mailboxes.' `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Created Application Access Policy for app {0} scoped to '{1}'." -f $AppId, $GroupAddress) -Level OK
        return $policy
    }
}

function Get-SmartM365ApplicationAccessPolicy {
    try {
        return @(Get-ApplicationAccessPolicy -ErrorAction Stop)
    }
    catch {
        $message = [string]$_.Exception.Message
        if ($message -match 'introuvable|not found|not be found|ObjectNotFound|Cannot find|couldn''t be found') {
            Write-SmartM365SetupStatus -Level WARN -Message 'No existing Exchange Online Application Access Policy was returned; continuing as if none exist.'
            return @()
        }

        throw
    }
}

function Remove-SmartM365MailSendApplicationAccessPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$AppId
    )

    $policies = @(Get-SmartM365ApplicationAccessPolicy | Where-Object { @($_.AppId) -contains $AppId })
    if ($policies.Count -eq 0) {
        Write-SmartM365SetupStatus -Message ("No Application Access Policy found for app {0}." -f $AppId) -Level OK
        return
    }

    foreach ($policy in $policies) {
        if ($PSCmdlet.ShouldProcess($policy.Identity, ("Remove Application Access Policy for app {0}" -f $AppId))) {
            Remove-ApplicationAccessPolicy -Identity $policy.Identity -Confirm:$false -ErrorAction Stop
            Write-SmartM365SetupStatus -Message ("Removed Application Access Policy '{0}' for app {1}." -f $policy.Identity, $AppId) -Level OK
        }
    }
}

function Test-SmartM365MailSendApplicationAccessPolicy {
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$SenderAddress
    )

    if (-not (Get-Command -Name Test-ApplicationAccessPolicy -ErrorAction SilentlyContinue)) {
        Write-SmartM365SetupStatus -Level WARN -Message 'Test-ApplicationAccessPolicy is not available in this Exchange Online session; skipping Mail.Send scope test.'
        return
    }

    try {
        $testResult = Test-ApplicationAccessPolicy -Identity $SenderAddress -AppId $AppId -ErrorAction Stop
        Write-SmartM365SetupStatus -Message ("Application Access Policy test for '{0}' returned: {1}." -f $SenderAddress, $testResult.AccessCheckResult) -Level OK
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Application Access Policy test failed for '{0}': {1}" -f $SenderAddress, $_.Exception.Message)
    }
}

function Set-SmartM365MailLocalConfig {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConfigPath,
        [Parameter(Mandatory)][string]$SenderAddress,
        [Parameter(Mandatory)][string]$ScopeGroupAddress
    )

    if (Test-Path -LiteralPath $ConfigPath) {
        $config = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    else {
        $config = [pscustomobject]@{}
    }

    $values = @{
        From                    = $SenderAddress
        SmtpServer              = ''
        MailSendAccessPolicyGroup = $ScopeGroupAddress
    }

    foreach ($propertyName in $values.Keys) {
        if ($null -eq $config.PSObject.Properties[$propertyName]) {
            $config | Add-Member -NotePropertyName $propertyName -NotePropertyValue $values[$propertyName]
        }
        else {
            $config.$propertyName = $values[$propertyName]
        }
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Update SmartM365 mail local configuration')) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Updated mail settings in {0}." -f $ConfigPath) -Level OK
    }
}

function Set-SmartM365ExchangeMailSendSetup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string]$AdminUserPrincipalName,
        [Parameter(Mandatory)][string]$ConfigPath,
        [string]$SenderAddress,
        [Parameter(Mandatory)][string]$SenderDisplayName,
        [Parameter(Mandatory)][string]$ScopeGroupName,
        [Parameter(Mandatory)][string]$ScopeGroupAlias
    )

    Connect-SmartM365ExchangeOnlineSetupSession -UserPrincipalName $AdminUserPrincipalName
    $effectiveSenderAddress = Resolve-SmartM365MailAddress -ConfiguredAddress $SenderAddress -DefaultAlias 'smartm365-reports'
    $domain = ($effectiveSenderAddress -split '@', 2)[1]
    $scopeGroupAddress = ('{0}@{1}' -f $ScopeGroupAlias, $domain)

    $mailbox = Get-OrCreate-SmartM365SenderMailbox -SenderAddress $effectiveSenderAddress -SenderDisplayName $SenderDisplayName
    if ($null -eq $mailbox) {
        Write-SmartM365SetupStatus -Level WARN -Message 'SmartM365 sender mailbox was not created because WhatIf was used.'
    }

    $group = Get-OrCreate-SmartM365MailSendSecurityGroup -GroupName $ScopeGroupName -GroupAlias $ScopeGroupAlias -GroupAddress $scopeGroupAddress
    if ($null -eq $group) {
        Write-SmartM365SetupStatus -Level WARN -Message 'SmartM365 Mail.Send scope group was not created because WhatIf was used.'
    }
    else {
        Add-SmartM365SenderMailboxToScopeGroup -Group $group -SenderAddress $effectiveSenderAddress
        Set-SmartM365MailSendApplicationAccessPolicy -AppId $Application.AppId -Group $group -GroupAddress $scopeGroupAddress | Out-Null
        Test-SmartM365MailSendApplicationAccessPolicy -AppId $Application.AppId -SenderAddress $effectiveSenderAddress
    }

    Set-SmartM365MailLocalConfig -ConfigPath $ConfigPath -SenderAddress $effectiveSenderAddress -ScopeGroupAddress $scopeGroupAddress

    return [pscustomobject]@{
        SenderAddress    = $effectiveSenderAddress
        ScopeGroupName   = $ScopeGroupName
        ScopeGroupAddress = $scopeGroupAddress
    }
}

function Set-SmartM365SharePointLocalConfig {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$WorkspaceInfo,
        [Parameter(Mandatory)][string]$ConfigPath
    )

    if (Test-Path -LiteralPath $ConfigPath) {
        $config = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    else {
        $config = [pscustomobject]@{}
    }

    foreach ($propertyName in @('SharePointSiteHostname', 'SharePointSitePath', 'SharePointLibraryDisplayName', 'SharePointTargetFolderPath')) {
        if ($null -eq $config.PSObject.Properties[$propertyName]) {
            $config | Add-Member -NotePropertyName $propertyName -NotePropertyValue $WorkspaceInfo.$propertyName
        }
        else {
            $config.$propertyName = $WorkspaceInfo.$propertyName
        }
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Update SharePoint local configuration')) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
    }
}

function Set-SmartM365AuthLocalConfig {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ConfigPath
    )

    if (Test-Path -LiteralPath $ConfigPath) {
        $config = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    else {
        $config = [pscustomobject]@{}
    }

    $values = @{
        AppId      = $Application.AppId
        TenantId   = $TenantId
        Thumb      = $Certificate.Thumbprint
        Thumbprint = $Certificate.Thumbprint
    }

    foreach ($propertyName in $values.Keys) {
        if ($null -eq $config.PSObject.Properties[$propertyName]) {
            $config | Add-Member -NotePropertyName $propertyName -NotePropertyValue $values[$propertyName]
        }
        else {
            $config.$propertyName = $values[$propertyName]
        }
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Update SmartM365 app authentication local configuration')) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
    }
}

function Clear-SmartM365AuthLocalConfig {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConfigPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return
    }

    $config = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $removedProperties = @()

    foreach ($propertyName in @('AppId', 'TenantId', 'Thumb', 'Thumbprint')) {
        if ($null -ne $config.PSObject.Properties[$propertyName]) {
            $config.PSObject.Properties.Remove($propertyName)
            $removedProperties += $propertyName
        }
    }

    if ($removedProperties.Count -eq 0) {
        Write-SmartM365SetupStatus -Message ("No app authentication settings to clear in {0}." -f $ConfigPath) -Level OK
        return
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, ("Clear SmartM365 app authentication settings: {0}" -f ($removedProperties -join ', ')))) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Cleared app authentication settings in {0}: {1}." -f $ConfigPath, ($removedProperties -join ', ')) -Level OK
    }
}

function Remove-SmartM365AppRegistration {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string]$ConfigPath
    )

    $servicePrincipals = @(Get-MgServicePrincipal -Filter "appId eq '$($Application.AppId)'" -All)

    foreach ($servicePrincipal in $servicePrincipals) {
        if ($PSCmdlet.ShouldProcess(("Service principal {0} ({1})" -f $servicePrincipal.DisplayName, $servicePrincipal.Id), 'Remove SmartM365 application service principal')) {
            Remove-MgServicePrincipal -ServicePrincipalId $servicePrincipal.Id -Confirm:$false
            Write-SmartM365SetupStatus -Message ("Removed application service principal '{0}' ({1})." -f $servicePrincipal.DisplayName, $servicePrincipal.Id) -Level OK
        }
    }

    if ($PSCmdlet.ShouldProcess(("App registration {0} ({1})" -f $Application.DisplayName, $Application.Id), 'Remove SmartM365 app registration')) {
        Remove-MgApplication -ApplicationId $Application.Id -Confirm:$false
        Write-SmartM365SetupStatus -Message ("Removed app registration '{0}' ({1})." -f $Application.DisplayName, $Application.Id) -Level OK
    }

    Clear-SmartM365AuthLocalConfig -ConfigPath $ConfigPath
    Write-SmartM365SetupStatus -Level WARN -Message 'Teams workspace, SharePoint files, and local certificates were not removed.'
}

function Get-SmartM365LocalConfigCertificateThumbprint {
    param(
        [Parameter(Mandatory)][string]$ConfigPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return $null
    }

    try {
        $config = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Could not read local config for certificate reuse: {0}" -f $_.Exception.Message)
        return $null
    }

    foreach ($propertyName in @('Thumbprint', 'Thumb')) {
        if ($null -eq $config.PSObject.Properties[$propertyName]) {
            continue
        }

        $configuredThumbprint = [string]$config.$propertyName
        if ([string]::IsNullOrWhiteSpace($configuredThumbprint)) {
            continue
        }

        $normalizedThumbprint = $configuredThumbprint.Replace(' ', '').ToUpperInvariant()
        $certificate = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') |
            ForEach-Object { Get-ChildItem -Path $_ -ErrorAction SilentlyContinue } |
            Where-Object { $_.Thumbprint -eq $normalizedThumbprint } |
            Select-Object -First 1

        if ($null -eq $certificate) {
            Write-SmartM365SetupStatus -Level WARN -Message ("Local config certificate thumbprint '{0}' was not found in CurrentUser\My or LocalMachine\My; falling back to app key lookup." -f $normalizedThumbprint)
            return $null
        }

        if (-not $certificate.HasPrivateKey) {
            Write-SmartM365SetupStatus -Level WARN -Message ("Local config certificate thumbprint '{0}' exists but has no private key; falling back to app key lookup." -f $normalizedThumbprint)
            return $null
        }

        Write-SmartM365SetupStatus -Message ("Reusing certificate thumbprint from local tenant config: {0}" -f $normalizedThumbprint) -Level OK
        return $normalizedThumbprint
    }

    return $null
}

Initialize-SmartM365SetupLogging -Path $LogPath

try {
Import-RequiredGraphModule
Disconnect-SmartM365ExistingGraphSession
Disconnect-SmartM365ExistingExchangeOnlineSession
Connect-SmartM365ExchangeOnlineSetupSession -UserPrincipalName $ExchangeAdminUserPrincipalName
$graphContext = Connect-SmartM365GraphSetupSession -RequestedTenantId $TenantId -UseDeviceCode:$UseDeviceCode
$effectiveTenantId = $graphContext.TenantId
$localConfigPath = Join-Path -Path $script:SmartM365RootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $Tenant)

$escapedDisplayName = ConvertTo-ODataStringLiteral -Value $DisplayName
$existingApps = @(Get-MgApplication -Filter "displayName eq '$escapedDisplayName'" -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials' -All)
$application = $null

if ($existingApps.Count -gt 1) {
    throw "Multiple app registrations named '$DisplayName' were found. Rename duplicates or use a unique -DisplayName."
}

if ($RemoveAppRegistration) {
    if ($existingApps.Count -eq 0) {
        Write-SmartM365SetupStatus -Level WARN -Message ("No app registration named '{0}' was found. Nothing to remove." -f $DisplayName)
        return
    }

    try {
        Connect-SmartM365ExchangeOnlineSetupSession -UserPrincipalName $graphContext.Account
        Remove-SmartM365MailSendApplicationAccessPolicy -AppId $existingApps[0].AppId
    }
    catch {
        Write-SmartM365SetupStatus -Level WARN -Message ("Exchange Online Mail.Send policy cleanup failed and app removal will continue: {0}" -f $_.Exception.Message)
    }

    Remove-SmartM365AppRegistration -Application $existingApps[0] -ConfigPath $localConfigPath
    return
}

$requiredApiResources = Get-SmartM365RequiredApiResource `
    -SkipExchange:$SkipExchangeOnlinePermission `
    -SkipBroadReadWrite:$SkipBroadIntuneReadWritePermissions

$obsoleteBroadSharePointPermissions = @('Files.ReadWrite.All', 'Sites.ReadWrite.All')
$exchangeOnlineDirectoryRoleName = 'Global Reader'
$obsoleteExchangeOnlineDirectoryRoleNames = @('Exchange Administrator')

if (-not $SkipBroadIntuneReadWritePermissions) {
    Write-SmartM365SetupStatus -Level WARN -Message 'Including broad Intune ReadWrite application permissions because one current Autopatch report script requests them. Use -SkipBroadIntuneReadWritePermissions after hardening those scripts.'
}

$requiredResourceAccess = Get-RequiredResourceAccessBlock -RequiredApiResource $requiredApiResources

$existingApplicationForCertificate = if ($existingApps.Count -eq 1) { $existingApps[0] } else { $null }
$effectiveCertificateThumbprint = if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    $CertificateThumbprint
}
else {
    Get-SmartM365LocalConfigCertificateThumbprint -ConfigPath $localConfigPath
}
$certificate = Get-SmartM365Certificate -Thumbprint $effectiveCertificateThumbprint -AppDisplayName $DisplayName -ValidityYears $CertificateYears -Application $existingApplicationForCertificate

if ($existingApps.Count -eq 1) {
    if (-not $UpdateExisting) {
        throw "App registration '$DisplayName' already exists. Re-run with -UpdateExisting to merge SmartM365 permissions and certificate."
    }

    $application = $existingApps[0]
    $mergedRequiredResourceAccess = Merge-RequiredResourceAccess -ExistingAccess $application.RequiredResourceAccess -RequiredAccess $requiredResourceAccess
    $mergedRequiredResourceAccess = Get-SmartM365RequiredResourceAccessWithoutPermission `
        -RequiredAccess $mergedRequiredResourceAccess `
        -ResourceName 'Microsoft Graph' `
        -ResourceAppId '00000003-0000-0000-c000-000000000000' `
        -PermissionValues $obsoleteBroadSharePointPermissions
    $mergedRequiredResourceAccess = ConvertTo-SmartM365RequiredResourceAccessPayload -RequiredAccess $mergedRequiredResourceAccess
    $mergedKeyCredentials = Add-CertificateToApplication -Application $application -Certificate $certificate

    if ($PSCmdlet.ShouldProcess($DisplayName, 'Update app registration permissions and certificate')) {
        Update-MgApplication `
            -ApplicationId $application.Id `
            -RequiredResourceAccess $mergedRequiredResourceAccess `
            -KeyCredentials $mergedKeyCredentials `
            -ErrorAction Stop
        $application = Get-MgApplication -ApplicationId $application.Id -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials'
        Write-SmartM365SetupStatus -Message ("Updated app registration '{0}'." -f $DisplayName) -Level OK
    }
}
else {
    if ($PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) {
        $application = New-MgApplication `
            -DisplayName $DisplayName `
            -SignInAudience 'AzureADMyOrg' `
            -RequiredResourceAccess $requiredResourceAccess `
            -KeyCredentials @(Get-KeyCredentialFromCertificate -Certificate $certificate) `
            -ErrorAction Stop
        Write-SmartM365SetupStatus -Message ("Created app registration '{0}'." -f $DisplayName) -Level OK
    }
}

if ($null -eq $application) {
    Write-SmartM365SetupStatus -Level WARN -Message 'No app registration changes were applied because WhatIf was used.'
    return
}

$application = Get-MgApplication -ApplicationId $application.Id -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials'
if (-not (Test-ApplicationHasCertificate -Application $application -Certificate $certificate)) {
    throw ("Certificate public key was not found on app registration '{0}' after create/update. Refusing to write unusable app-only configuration." -f $DisplayName)
}
Write-SmartM365SetupStatus -Message 'Certificate public key verified on the app registration.' -Level OK

Set-SmartM365AuthLocalConfig -Application $application -Certificate $certificate -TenantId $effectiveTenantId -ConfigPath $localConfigPath
if (-not $WhatIfPreference) {
    Write-SmartM365SetupStatus -Message ("Updated app authentication settings in {0}." -f $localConfigPath) -Level OK
}

$appServicePrincipal = @(Get-MgServicePrincipal -Filter "appId eq '$($application.AppId)'" -All) | Select-Object -First 1
if ($null -eq $appServicePrincipal) {
    if ($PSCmdlet.ShouldProcess($DisplayName, 'Create application service principal')) {
        $appServicePrincipal = New-MgServicePrincipal -AppId $application.AppId
        Write-SmartM365SetupStatus -Message 'Created application service principal.' -Level OK
    }
}

if (-not $DisableGrantAdminConsent -and $null -ne $appServicePrincipal) {
    Grant-SmartM365ApplicationPermission -ApplicationServicePrincipal $appServicePrincipal -RequiredApiResource $requiredApiResources
    Revoke-SmartM365ApplicationPermission `
        -ApplicationServicePrincipal $appServicePrincipal `
        -ResourceName 'Microsoft Graph' `
        -ResourceAppId '00000003-0000-0000-c000-000000000000' `
        -PermissionValues $obsoleteBroadSharePointPermissions
}
elseif ($DisableGrantAdminConsent) {
    Write-SmartM365SetupStatus -Level WARN -Message 'Admin consent was not granted because -DisableGrantAdminConsent was used. Use the admin consent URL printed below if needed.'
}

$exchangeOnlineDirectoryRole = $null
if (-not $SkipExchangeOnlinePermission -and $null -ne $appServicePrincipal) {
    $exchangeOnlineDirectoryRole = Grant-SmartM365ExchangeOnlineDirectoryRole `
        -ApplicationServicePrincipal $appServicePrincipal `
        -RoleDisplayName $exchangeOnlineDirectoryRoleName
    foreach ($obsoleteRoleName in $obsoleteExchangeOnlineDirectoryRoleNames) {
        Revoke-SmartM365ExchangeOnlineDirectoryRole `
            -ApplicationServicePrincipal $appServicePrincipal `
            -RoleDisplayName $obsoleteRoleName
    }
}

$mailSendSetup = Set-SmartM365ExchangeMailSendSetup `
    -Application $application `
    -AdminUserPrincipalName $graphContext.Account `
    -ConfigPath $localConfigPath `
    -SenderAddress $MailSenderAddress `
    -SenderDisplayName $MailSenderDisplayName `
    -ScopeGroupName $MailSendSecurityGroupName `
    -ScopeGroupAlias $MailSendSecurityGroupAlias

$teamsWorkspace = $null
if (-not $DisableTeamsSetup) {
    $teamsWorkspace = Get-OrCreate-SmartM365TeamsWorkspace `
        -GroupDisplayName $TeamDisplayName `
        -MailNickname $TeamMailNickname `
        -OwnerUserPrincipalName $graphContext.Account `
        -TargetFolderPath $SharePointTargetFolderPath `
        -DefaultLibraryDisplayName 'Documents'

    Set-SmartM365SharePointLocalConfig -WorkspaceInfo $teamsWorkspace -ConfigPath $localConfigPath
    if (-not $WhatIfPreference) {
        Write-SmartM365SetupStatus -Message ("Updated SharePoint settings in {0}." -f $localConfigPath) -Level OK
    }

    Grant-SmartM365SelectedSitePermission `
        -SiteId $teamsWorkspace.SharePointSiteId `
        -ApplicationAppId $application.AppId `
        -ApplicationDisplayName $DisplayName `
        -Role 'write'
}
else {
    Write-SmartM365SetupStatus -Level WARN -Message 'Teams workspace creation and SharePoint local configuration update were skipped because -DisableTeamsSetup was used.'
}

$adminConsentUrl = "https://login.microsoftonline.com/$effectiveTenantId/adminconsent?client_id=$($application.AppId)"

Write-Output ''
Write-SmartM365SetupStatus -Message 'SmartM365 app registration summary' -Level OK
Write-Output ("DisplayName : {0}" -f $DisplayName)
Write-Output ("TenantId    : {0}" -f $effectiveTenantId)
Write-Output ("AppId       : {0}" -f $application.AppId)
Write-Output ("ObjectId    : {0}" -f $application.Id)
Write-Output ("SpObjectId  : {0}" -f $(if ($appServicePrincipal) { $appServicePrincipal.Id } else { '<not created>' }))
Write-Output ("Thumbprint  : {0}" -f $certificate.Thumbprint)
Write-Output ("Consent URL : {0}" -f $adminConsentUrl)
if ($teamsWorkspace) {
    Write-Output ("Team        : {0} ({1})" -f $teamsWorkspace.TeamDisplayName, $teamsWorkspace.TeamId)
    Write-Output ("Alerts Ch   : {0} ({1})" -f $teamsWorkspace.AlertsChannelDisplayName, $(if ($teamsWorkspace.AlertsChannelId) { $teamsWorkspace.AlertsChannelId } else { '<not created>' }))
    Write-Output ("Infos Ch    : {0} ({1})" -f $teamsWorkspace.InfosChannelDisplayName, $(if ($teamsWorkspace.InfosChannelId) { $teamsWorkspace.InfosChannelId } else { '<not created>' }))
    Write-Output ("SP Host     : {0}" -f $teamsWorkspace.SharePointSiteHostname)
    Write-Output ("SP Path     : {0}" -f $teamsWorkspace.SharePointSitePath)
    Write-Output ("SP Library  : {0}" -f $teamsWorkspace.SharePointLibraryDisplayName)
    Write-Output ("SP Folder   : {0}" -f $teamsWorkspace.SharePointTargetFolderPath)
}
if ($mailSendSetup) {
    Write-Output ("Mail From   : {0}" -f $mailSendSetup.SenderAddress)
    Write-Output ("Mail Scope  : {0} ({1})" -f $mailSendSetup.ScopeGroupName, $mailSendSetup.ScopeGroupAddress)
}
if ($exchangeOnlineDirectoryRole) {
    Write-Output ("EXO Role    : {0}" -f $exchangeOnlineDirectoryRole)
}
Write-Output ''
Write-Output ("Tenant local configuration updated: {0}" -f $localConfigPath)
Write-Output 'Core app values:'
Write-Output (@"
{
  "AppId": "$($application.AppId)",
  "TenantId": "$effectiveTenantId",
  "Thumb": "$($certificate.Thumbprint)",
  "Thumbprint": "$($certificate.Thumbprint)"
}
"@)
if ($teamsWorkspace) {
    Write-Output ''
    Write-Output 'SharePoint settings written to tenant local configuration:'
    Write-Output (@"
{
  "SharePointSiteHostname": "$($teamsWorkspace.SharePointSiteHostname)",
  "SharePointSitePath": "$($teamsWorkspace.SharePointSitePath)",
  "SharePointLibraryDisplayName": "$($teamsWorkspace.SharePointLibraryDisplayName)",
  "SharePointTargetFolderPath": "$($teamsWorkspace.SharePointTargetFolderPath)"
}
"@)
}
Write-Output ''
if ($mailSendSetup) {
    Write-Output 'Mail settings written to tenant local configuration:'
    Write-Output (@"
{
  "From": "$($mailSendSetup.SenderAddress)",
  "SmtpServer": "",
  "MailSendAccessPolicyGroup": "$($mailSendSetup.ScopeGroupAddress)"
}
"@)
    Write-Output ''
}
if ($SkipExchangeOnlinePermission) {
    Write-SmartM365SetupStatus -Level WARN -Message 'Exchange Online app-only permission and directory role assignment were skipped because -SkipExchangeOnlinePermission was used.'
}
}
finally {
    Close-SmartM365ExchangeOnlineSetupSession
    Close-SmartM365SetupLogging
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+9UdrSQRkrg4h
# u0nqRCA6lOUsM+lOe+5/GA972JxCUKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJoKbQoAJQaWSUJGmIDndvRsjpO9ax4UZbHpBZTKgtgBMA0GCSqG
# SIb3DQEBAQUABIIBgDZzpKk8+DAF6O8LEg0wd5QRXPcj9w/9RmCvksFuItbtWKAe
# XlhyKinQqTlVdYC0yPoizV23FaZCZN5Czucs2pPPKzenbFYv6IPAnq7bcoahUFxo
# 0QWLAC2svosA0DB7TiJbmKP4+N27J7YToGKgZwg3KeKqAFZAPxhalJJpStRvCSo9
# Xutn321LJPKJxCLH2PlgcVOaF8U20LQCKkDZg5H2uNH61EZVzc6oLAyFBKCwQANM
# aQHmgmI4/8Vb3b1giTB6G9LPQFsF0kwsiH+ibldfFO5w4FTan950hGUJSY7eL71H
# ukf1lJ7mFvjr5G6UrMciwg0ydSDHZCr08VCRsHK+8sNR19BRVv9UCfHp7H4En0n0
# bpIEhSnU44NwshfowiO6ZKddZ8Dr7O4WrG7l4HsW52v/K8i9Xx54NKfSW/UVZdpg
# SAhCIXqnf41o84o79yjpOMJ48o1+zNdUpqZLh5CmFkryWy2EYvrFx5ciwEA8tQ9a
# yqpbzYtXUYQpmOXmfqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgyMDMw
# NTBaMC8GCSqGSIb3DQEJBDEiBCByDaU7HyHlKjnwVIRXarZ/FhQ9VXXe/DdTBhtF
# EGY8BjANBgkqhkiG9w0BAQEFAASCAgCrl/9g9TFVhcCrmtjkxWPzP/tk4YXURdt6
# 2pP8tzScB1Z0shmKxMamMnNs7LSO1GUpU9HLZeznYq4LiHbhmQjJVSYzRH17zaE6
# I1zxIrl0+RC9Yr0r3NfJ10SW9YrUSg4HGeYqjDXCn3xqxgOJ8+F2UV3b0Im4uI++
# RJ/ibvjPXK2P00EmlkEIbn+OFgd75dkac1VbDrsrVt6CijeRIr6nzSt0xzQpyYy9
# nwxY8/n/McXgEP9YwYcmGiKLNTKjZmG9ywDqDnTQgndF+vpDkHvYgz0uaRzNfYKO
# gjc0tAGyDgnM9FMFtdBBGC97vXrtJUIlVuesm3CHkh0neOOBVHGIkVIhvUxH484Z
# /qUBombPbC6NbBYfMHWEUuk8Tz04sP68lpngnwUOxzUTXwfHLv8qBEAy4j4n44E7
# 17K29PinqOIdKdRuB9O6pSAajcK0lRmISO1VQ0ORd8TAID4omBXtpZvnAy/WHR6u
# kTpDR3ikJ1i5MRJ9h2kT66GUvgBfnzOkJ0pps7vHdWXne8rMQ7T3kzGk+Ahs3EmM
# EV5C7p/I7OdpVx3IgIiaACMRB0hykrLLG+z4yHJOQ+Hqxy/akqf2PVas3khWAY08
# n2FLFE1quMhhxVD+CdThGWzMZhkF92yVT7QPXO4uy5RD+EEELBVPnQrGeg5rsmnY
# WApQCVGJ8w==
# SIG # End signature block
