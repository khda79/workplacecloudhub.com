#Requires -Version 7.0

<#
.SYNOPSIS
    Creates or updates the Entra ID app registration used by SmartAzure automation.

.DESCRIPTION
    Creates an app registration, adds the Microsoft Graph application permissions
    required by the shared SmartAzure mail and SharePoint features, and uploads
    a certificate public key.

    This bootstrap script is intentionally interactive-only. It signs in with
    delegated Microsoft Graph setup scopes and must be run by an administrator
    allowed to create app registrations, update application permissions, and
    grant tenant-wide admin consent by default.

    Use -DisableGrantAdminConsent only when you want to add permissions without
    granting tenant-wide consent immediately.

    The script also creates or reuses a Teams team for SmartAzure shared exports,
    resolves its backing SharePoint site, and updates the selected tenant local profile
    with the SharePoint upload target.

    The script also creates or reuses a dedicated Exchange Online shared mailbox
    for SmartAzure report/error emails, creates a mail-enabled security group,
    adds the sender mailbox to that group, and scopes Microsoft Graph Mail.Send
    with an Exchange Online Application Access Policy.

.PARAMETER DisplayName
    Display name for the Entra ID application. Defaults to SmartAzure Automation.

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
    Mail.Send scope setup connection.

.PARAMETER DisableMailSendScopeSetup
    Skips shared mailbox, mail-enabled security group, and Exchange Online
    Application Access Policy setup for Microsoft Graph Mail.Send.

.PARAMETER TeamDisplayName
    Display name of the Teams team used for SmartAzure shared exports. Defaults
    to SMART-AZURE.

.PARAMETER MailSenderAddress
    Primary SMTP address for the SmartAzure sender mailbox. If omitted, the
    script uses smartazure-reports@<default accepted domain>.

.PARAMETER MailSenderDisplayName
    Display name for the SmartAzure sender shared mailbox.

.PARAMETER MailSendSecurityGroupName
    Display name for the mail-enabled security group that scopes Mail.Send.

.PARAMETER MailSendSecurityGroupAlias
    Alias for the mail-enabled security group that scopes Mail.Send.

.PARAMETER DisableTeamsSetup
    Skips Teams team creation/reuse and SharePoint local configuration updates.

.PARAMETER LogPath
    Folder where the setup log and transcript are written. Defaults to
    Data\Tenants\<TenantKey>\LOG-ALL\Setup under the SmartAzure root, with
    fallback to Setup\Output\Tenants\<TenantKey>\LOG-ALL\Setup.

.EXAMPLE
    .\Setup\SmartAzure-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Setup\SmartAzure-Create-AppRegistration.ps1 -DisplayName SmartAzure -TenantId 00000000-0000-0000-0000-000000000000 -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UpdateExisting

.EXAMPLE
    .\Setup\SmartAzure-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UseDeviceCode -WhatIf

.EXAMPLE
    .\Setup\SmartAzure-Create-AppRegistration.ps1 -RemoveAppRegistration -Confirm
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Tenant = 'test',
[Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName = 'SmartAzure Automation',

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
    [string]$ExchangeAdminUserPrincipalName,

    [Parameter()]
    [switch]$DisableMailSendScopeSetup,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamDisplayName = 'SMART-AZURE',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamMailNickname = 'SMARTAZURE',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SharePointTargetFolderPath = 'SMART-AZURE/CSV',

    [Parameter()]
    [string]$MailSenderAddress,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSenderDisplayName = 'SmartAzure Reports',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSendSecurityGroupName = 'SMART-AZURE-MailSend-Allowed',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$MailSendSecurityGroupAlias = 'smartazure-mailsend-allowed',

    [Parameter()]
    [switch]$DisableTeamsSetup,

    [Parameter()]
    [string]$LogPath = ''
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartAzure-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartAzure-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzure-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAzureRootPath = Find-SmartAzureRoot -StartPath $PSScriptRoot
Initialize-SmartAzureTenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $defaultSetupLogPath = Join-Path -Path $script:SmartAzureRootPath -ChildPath ("Data\Tenants\{0}\LOG-ALL\Setup" -f $Tenant)
    if (Test-SmartAzureWritableDirectory -Path $defaultSetupLogPath) {
        $LogPath = $defaultSetupLogPath
    }
    else {
        $LogPath = Join-Path -Path $PSScriptRoot -ChildPath ("Output\Tenants\{0}\LOG-ALL\Setup" -f $Tenant)
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SmartAzureSetupLogFile = $null
$script:SmartAzureSetupTranscriptFile = $null
$script:SmartAzureSetupTranscriptStarted = $false
$script:SmartAzureExchangeOnlineConnected = $false

function Format-SmartAzureTimestampedLine {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    @([regex]::Split($Message, '\r?\n')) | ForEach-Object {
        "{0} [{1}] {2}" -f $timestamp, $Level, $_
    }
}

function Update-SmartAzureTimestampedTranscriptFile {
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

function Initialize-SmartAzureSetupLogging {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $resolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedLogPath)) {
        New-Item -Path $resolvedLogPath -ItemType Directory -Force -ErrorAction Stop -Confirm:$false | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $baseName = 'SmartAzure-Create-AppRegistration'
    $script:SmartAzureSetupLogFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}.log" -f $baseName, $timestamp)
    $script:SmartAzureSetupTranscriptFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}_Transcript.log" -f $baseName, $timestamp)

    Add-Content -LiteralPath $script:SmartAzureSetupLogFile -Encoding UTF8 -Value (Format-SmartAzureTimestampedLine -Message 'Log started.' -Level 'INFO') -Confirm:$false

    try {
        Start-Transcript -Path $script:SmartAzureSetupTranscriptFile -Append -ErrorAction Stop -Confirm:$false | Out-Null
        $script:SmartAzureSetupTranscriptStarted = $true
    }
    catch {
        Add-Content -LiteralPath $script:SmartAzureSetupLogFile -Encoding UTF8 -Value (Format-SmartAzureTimestampedLine -Message ("Transcript could not be started: {0}" -f $_.Exception.Message) -Level 'WARN') -Confirm:$false
    }
}

function Close-SmartAzureSetupLogging {
    if ($script:SmartAzureSetupLogFile) {
        Add-Content -LiteralPath $script:SmartAzureSetupLogFile -Encoding UTF8 -Value (Format-SmartAzureTimestampedLine -Message 'Log finished.' -Level 'INFO') -Confirm:$false
        Write-Information ("[INFO] Log file: {0}" -f $script:SmartAzureSetupLogFile) -InformationAction Continue
        if ($script:SmartAzureSetupTranscriptFile) {
            Write-Information ("[INFO] Transcript file: {0}" -f $script:SmartAzureSetupTranscriptFile) -InformationAction Continue
        }
    }

    if ($script:SmartAzureSetupTranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch { Write-Debug ("Stop-Transcript failed: {0}" -f $_.Exception.Message) }
        try { Update-SmartAzureTimestampedTranscriptFile -Path $script:SmartAzureSetupTranscriptFile } catch { Write-Debug ("Transcript timestamp normalization failed: {0}" -f $_.Exception.Message) }
        $script:SmartAzureSetupTranscriptStarted = $false
    }
}

function Write-SmartAzureSetupStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][string]$Level = 'INFO'
    )

    if ($script:SmartAzureSetupLogFile) {
        $logEntry = Format-SmartAzureTimestampedLine -Message $Message -Level $Level
        try { Add-Content -LiteralPath $script:SmartAzureSetupLogFile -Encoding UTF8 -Value $logEntry -Confirm:$false } catch { Write-Debug ("Log write failed: {0}" -f $_.Exception.Message) }
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

function Import-RequiredGraphModule {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications'
    )

    foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            throw "Required module '$moduleName' is not installed. Install it with: Install-Module $moduleName -Scope CurrentUser"
        }
        Import-Module $moduleName -ErrorAction Stop
    }
}

function Import-RequiredExchangeOnlineModule {
    $moduleName = 'ExchangeOnlineManagement'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw "Required module '$moduleName' is not installed. Install it with: Install-Module $moduleName -Scope CurrentUser"
    }

    Import-Module $moduleName -ErrorAction Stop
}

function Disconnect-SmartAzureExistingGraphSession {
    try {
        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $existingContext -and -not [string]::IsNullOrWhiteSpace($existingContext.Account)) {
            Write-SmartAzureSetupStatus -Message ("Disconnecting existing Microsoft Graph session for {0}." -f $existingContext.Account)
        }
        else {
            Write-SmartAzureSetupStatus -Message 'Clearing any existing Microsoft Graph session.'
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-SmartAzureSetupStatus -Level WARN -Message ("Could not disconnect existing Microsoft Graph session: {0}" -f $_.Exception.Message)
    }
}

function Disconnect-SmartAzureExistingExchangeOnlineSession {
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
                Write-SmartAzureSetupStatus -Message ("Disconnecting existing Exchange Online session(s): {0}." -f ($connectionNames -join ', '))
            }
            else {
                Write-SmartAzureSetupStatus -Message 'Disconnecting existing Exchange Online session(s).'
            }
        }
        else {
            Write-SmartAzureSetupStatus -Message 'Clearing any existing Exchange Online session.'
        }

        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-SmartAzureSetupStatus -Level WARN -Message ("Could not disconnect existing Exchange Online session: {0}" -f $_.Exception.Message)
    }
}

function Connect-SmartAzureGraphSetupSession {
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

    Write-SmartAzureSetupStatus -Message ("Connecting interactively to Microsoft Graph with setup scopes: {0}" -f ($scopes -join ', '))
    Connect-MgGraph @connectParams | Out-Null

    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace($context.TenantId)) {
        throw 'Microsoft Graph connection did not return a tenant context.'
    }
    if ($context.AuthType -ne 'Delegated' -or [string]::IsNullOrWhiteSpace($context.Account)) {
        throw 'This setup script must be run with an interactive delegated Microsoft Graph administrator session.'
    }

    Write-SmartAzureSetupStatus -Message ("Connected as {0} in tenant {1}." -f $context.Account, $context.TenantId) -Level OK

    return $context
}

function Connect-SmartAzureExchangeOnlineSetupSession {
    param(
        [string]$UserPrincipalName
    )

    if ($script:SmartAzureExchangeOnlineConnected) {
        return
    }

    Import-RequiredExchangeOnlineModule
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Write-SmartAzureSetupStatus -Message 'Connecting interactively to Exchange Online before Microsoft Graph.'
    }
    else {
        Write-SmartAzureSetupStatus -Message ("Connecting interactively to Exchange Online as {0} before Microsoft Graph." -f $UserPrincipalName)
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

        Write-SmartAzureSetupStatus -Level WARN -Message 'Exchange Online WAM/broker authentication failed. Retrying with device code authentication.'
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
    $script:SmartAzureExchangeOnlineConnected = $true
    Write-SmartAzureSetupStatus -Message 'Connected to Exchange Online.' -Level OK
}

function Close-SmartAzureExchangeOnlineSetupSession {
    if (-not $script:SmartAzureExchangeOnlineConnected) {
        return
    }

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-SmartAzureSetupStatus -Message 'Disconnected from Exchange Online.' -Level OK
    }
    catch {
        Write-SmartAzureSetupStatus -Level WARN -Message ("Exchange Online disconnect failed: {0}" -f $_.Exception.Message)
    }
    finally {
        $script:SmartAzureExchangeOnlineConnected = $false
    }
}

function Get-SmartAzureRequiredApiResource {
    $graphPermissions = @(
        'Mail.Send',
        'Sites.Selected'
    )

    return @(
        [pscustomobject]@{
            Name           = 'Microsoft Graph'
            ResourceAppId  = '00000003-0000-0000-c000-000000000000'
            AppRoleValues  = @($graphPermissions | Sort-Object -Unique)
        }
    )
}

function Get-OrCreate-ServicePrincipalByAppId {
    param(
        [Parameter(Mandatory)][string]$ResourceName,
        [Parameter(Mandatory)][string]$ResourceAppId
    )

    $filter = "appId eq '$ResourceAppId'"
    $servicePrincipal = @(Get-MgServicePrincipal -Filter $filter -All) | Select-Object -First 1

    if ($null -eq $servicePrincipal) {
        Write-SmartAzureSetupStatus -Message ("Creating service principal for {0} ({1})." -f $ResourceName, $ResourceAppId)
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

function Get-SmartAzureRequiredResourceAccessWithoutPermission {
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

function ConvertTo-SmartAzureRequiredResourceAccessPayload {
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

function Get-SmartAzureCertificate {
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
                Write-SmartAzureSetupStatus -Message ("Reusing existing local certificate already attached to the app registration: {0}" -f $localCertificate.Thumbprint) -Level OK
                return $localCertificate
            }
        }
    }

    if ($WhatIfPreference) {
        throw 'WhatIf mode without -CertificateThumbprint would require creating a local self-signed certificate. Provide an existing -CertificateThumbprint for a side-effect-free dry-run.'
    }

    $subject = "CN=$AppDisplayName"
    Write-SmartAzureSetupStatus -Message ("Creating self-signed certificate '{0}' in CurrentUser\My." -f $subject)
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
        Write-SmartAzureSetupStatus -Message 'Certificate public key is already present on the app registration.' -Level OK
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

function Grant-SmartAzureApplicationPermission {
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
                Write-SmartAzureSetupStatus -Message ("Admin consent already granted: {0} / {1}" -f $resource.Name, $permission) -Level OK
                continue
            }

            if ($PSCmdlet.ShouldProcess(("{0} / {1}" -f $resource.Name, $permission), 'Grant admin consent')) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ApplicationServicePrincipal.Id -BodyParameter @{
                    principalId = $ApplicationServicePrincipal.Id
                    resourceId  = $resourceSp.Id
                    appRoleId   = $role.Id
                } | Out-Null
                Write-SmartAzureSetupStatus -Message ("Granted admin consent: {0} / {1}" -f $resource.Name, $permission) -Level OK
            }
        }
    }
}

function Revoke-SmartAzureApplicationPermission {
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
            Write-SmartAzureSetupStatus -Message ("Broad admin consent not present: {0} / {1}" -f $ResourceName, $permission) -Level OK
            continue
        }

        foreach ($assignment in $assignmentsToRemove) {
            if ($PSCmdlet.ShouldProcess(("{0} / {1}" -f $ResourceName, $permission), 'Revoke broad admin consent')) {
                Remove-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $ApplicationServicePrincipal.Id `
                    -AppRoleAssignmentId $assignment.Id `
                    -ErrorAction Stop
                Write-SmartAzureSetupStatus -Message ("Revoked broad admin consent: {0} / {1}" -f $ResourceName, $permission) -Level OK
            }
        }
    }
}

function Test-SmartAzureGraphNotFound {
    param([Parameter(Mandatory)]$ErrorRecord)

    $message = [string]$ErrorRecord.Exception.Message
    return ($message -match '\b404\b' -or $message -match 'NotFound' -or $message -match 'Not Found')
}

function Get-SmartAzureSetupUser {
    param([Parameter(Mandatory)][string]$UserPrincipalName)

    $encodedUser = [System.Uri]::EscapeDataString($UserPrincipalName)
    return Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/users/${encodedUser}?`$select=id,userPrincipalName,displayName" `
        -OutputType PSObject `
        -ErrorAction Stop
}

function Get-SmartAzureGroupByDisplayName {
    param([Parameter(Mandatory)][string]$GroupDisplayName)

    $escapedName = ConvertTo-ODataStringLiteral -Value $GroupDisplayName
    $response = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escapedName'&`$select=id,displayName,mailNickname,resourceProvisioningOptions,webUrl" `
        -OutputType PSObject `
        -ErrorAction Stop

    return @($response.value)
}

function Invoke-SmartAzureTeamGroupCreation {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][string]$MailNickname,
        [Parameter(Mandatory)][string]$OwnerUserId
    )

    $ownerUrl = "https://graph.microsoft.com/v1.0/users/$OwnerUserId"
    $body = @{
        description          = 'SmartAzure shared workspace for automation exports.'
        displayName          = $GroupDisplayName
        groupTypes           = @('Unified')
        mailEnabled          = $true
        mailNickname         = $MailNickname
        securityEnabled      = $false
        visibility           = 'Private'
        'owners@odata.bind'  = @($ownerUrl)
        'members@odata.bind' = @($ownerUrl)
    }

    Write-SmartAzureSetupStatus -Message ("Creating Microsoft 365 group for Teams team '{0}'." -f $GroupDisplayName)
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

function Test-SmartAzureTeamPresent {
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
        if (Test-SmartAzureGraphNotFound -ErrorRecord $_) {
            return $false
        }
        throw
    }
}

function Invoke-SmartAzureTeamCreationFromGroup {
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
            Write-SmartAzureSetupStatus -Message ("Creating Teams team from group {0} (attempt {1}/6)." -f $GroupId, $attempt)
            if ($PSCmdlet.ShouldProcess($GroupId, 'Create Teams team from Microsoft 365 group')) {
                $null = Invoke-MgGraphRequest `
                    -Method PUT `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/team" `
                    -Body ($body | ConvertTo-Json -Depth 8) `
                    -ContentType 'application/json' `
                    -OutputType PSObject `
                    -ErrorAction Stop
            }
            Write-SmartAzureSetupStatus -Message 'Teams team created.' -Level OK
            return
        }
        catch {
            if ($attempt -ge 6) {
                throw
            }
            Write-SmartAzureSetupStatus -Level WARN -Message ("Teams creation is not ready yet: {0}. Retrying in 10 seconds." -f $_.Exception.Message)
            Start-Sleep -Seconds 10
        }
    }
}

function Get-SmartAzureTeamChannelByDisplayName {
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

function Get-OrCreate-SmartAzureTeamChannel {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][string]$ChannelDisplayName,
        [Parameter(Mandatory)][string]$Description
    )

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            $existingChannel = Get-SmartAzureTeamChannelByDisplayName -TeamId $TeamId -ChannelDisplayName $ChannelDisplayName
            if ($existingChannel) {
                Write-SmartAzureSetupStatus -Message ("Teams channel already exists: {0}." -f $ChannelDisplayName) -Level OK
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
                Write-SmartAzureSetupStatus -Message ("Created Teams channel '{0}'." -f $ChannelDisplayName) -Level OK
                return $channel
            }

            return $null
        }
        catch {
            if ($attempt -ge 6) {
                throw
            }
            Write-SmartAzureSetupStatus -Level WARN -Message ("Teams channel '{0}' is not ready yet: {1}. Retrying in 10 seconds." -f $ChannelDisplayName, $_.Exception.Message)
            Start-Sleep -Seconds 10
        }
    }
}

function Wait-SmartAzureTeamSharePointSite {
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
            if (-not (Test-SmartAzureGraphNotFound -ErrorRecord $_)) {
                Write-SmartAzureSetupStatus -Level WARN -Message ("SharePoint site lookup failed: {0}" -f $_.Exception.Message)
            }
        }

        Write-SmartAzureSetupStatus -Message ("Waiting for Teams SharePoint site provisioning (attempt {0}/30)." -f $attempt)
        Start-Sleep -Seconds 10
    }

    throw "The SharePoint site for group '$GroupId' was not available after 5 minutes."
}

function Get-SmartAzureGroupDriveName {
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
        Write-SmartAzureSetupStatus -Level WARN -Message ("Could not resolve default document library name, using '{0}': {1}" -f $DefaultLibraryDisplayName, $_.Exception.Message)
    }

    return $DefaultLibraryDisplayName
}

function Get-OrCreate-SmartAzureTeamsWorkspace {
    param(
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][string]$MailNickname,
        [Parameter(Mandatory)][string]$OwnerUserPrincipalName,
        [Parameter(Mandatory)][string]$TargetFolderPath,
        [Parameter(Mandatory)][string]$DefaultLibraryDisplayName
    )

    $setupUser = Get-SmartAzureSetupUser -UserPrincipalName $OwnerUserPrincipalName
    $groups = @(Get-SmartAzureGroupByDisplayName -GroupDisplayName $GroupDisplayName)

    if ($groups.Count -gt 1) {
        $matchingNickname = @($groups | Where-Object { $_.mailNickname -eq $MailNickname }) | Select-Object -First 1
        if ($null -eq $matchingNickname) {
            throw "Multiple Microsoft 365 groups named '$GroupDisplayName' were found. Rename duplicates or specify a unique -TeamDisplayName."
        }
        $group = $matchingNickname
    }
    elseif ($groups.Count -eq 1) {
        $group = $groups[0]
        Write-SmartAzureSetupStatus -Message ("Reusing existing Microsoft 365 group '{0}' ({1})." -f $group.displayName, $group.id) -Level OK
    }
    else {
        $group = Invoke-SmartAzureTeamGroupCreation -GroupDisplayName $GroupDisplayName -MailNickname $MailNickname -OwnerUserId $setupUser.id
        Write-SmartAzureSetupStatus -Message ("Created Microsoft 365 group '{0}' ({1})." -f $group.displayName, $group.id) -Level OK
    }

    if (Test-SmartAzureTeamPresent -GroupId $group.id) {
        Write-SmartAzureSetupStatus -Message ("Teams team already exists for '{0}'." -f $GroupDisplayName) -Level OK
    }
    else {
        Invoke-SmartAzureTeamCreationFromGroup -GroupId $group.id
    }

    $alertsChannel = Get-OrCreate-SmartAzureTeamChannel `
        -TeamId $group.id `
        -ChannelDisplayName 'Alerts' `
        -Description 'SmartAzure script error and failure notifications.'
    $infosChannel = Get-OrCreate-SmartAzureTeamChannel `
        -TeamId $group.id `
        -ChannelDisplayName 'Infos' `
        -Description 'SmartAzure successful completion and informational notifications.'

    $site = Wait-SmartAzureTeamSharePointSite -GroupId $group.id
    $siteUri = [System.Uri]$site.webUrl
    $libraryDisplayName = Get-SmartAzureGroupDriveName -GroupId $group.id -DefaultLibraryDisplayName $DefaultLibraryDisplayName

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

function Get-SmartAzureApplicationIdFromSitePermission {
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

function Grant-SmartAzureSelectedSitePermission {
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
        @(Get-SmartAzureApplicationIdFromSitePermission -Permission $_) -contains $ApplicationAppId
    } | Select-Object -First 1

    if ($existingPermission) {
        $roles = @($existingPermission.roles)
        if ($roles -contains $Role -or $roles -contains 'fullcontrol' -or ($Role -eq 'write' -and $roles -contains 'manage')) {
            Write-SmartAzureSetupStatus -Message ("Sites.Selected permission already grants '{0}' or higher to app '{1}' on site '{2}'." -f $Role, $ApplicationDisplayName, $SiteId) -Level OK
            return
        }

        Write-SmartAzureSetupStatus -Level WARN -Message ("App '{0}' already has selected-site permission on '{1}', but roles are '{2}'. Create/update manually if a higher role is required." -f $ApplicationDisplayName, $SiteId, ($roles -join ', '))
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
        Write-SmartAzureSetupStatus -Message ("Granted Sites.Selected '{0}' permission to app '{1}' on site '{2}'." -f $Role, $ApplicationDisplayName, $SiteId) -Level OK
    }
}

function Get-SmartAzureExchangeDefaultAcceptedDomain {
    $defaultDomain = @(Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.Default -eq $true }) | Select-Object -First 1
    if ($null -eq $defaultDomain) {
        $defaultDomain = @(Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1)
    }

    if ($null -eq $defaultDomain -or [string]::IsNullOrWhiteSpace([string]$defaultDomain.DomainName)) {
        throw 'Could not resolve an Exchange Online accepted domain for SmartAzure mail setup.'
    }

    return [string]$defaultDomain.DomainName
}

function Resolve-SmartAzureMailAddress {
    param(
        [string]$ConfiguredAddress,
        [Parameter(Mandatory)][string]$DefaultAlias
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredAddress)) {
        return $ConfiguredAddress.Trim()
    }

    $domain = Get-SmartAzureExchangeDefaultAcceptedDomain
    return ('{0}@{1}' -f $DefaultAlias, $domain)
}

function Get-OrCreate-SmartAzureSenderMailbox {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SenderAddress,
        [Parameter(Mandatory)][string]$SenderDisplayName
    )

    $mailbox = Get-Mailbox -Identity $SenderAddress -ErrorAction SilentlyContinue
    if ($null -ne $mailbox) {
        if ($mailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            Write-SmartAzureSetupStatus -Level WARN -Message ("Reusing existing sender mailbox '{0}', but it is '{1}' rather than a SharedMailbox." -f $SenderAddress, $mailbox.RecipientTypeDetails)
        }
        else {
            Write-SmartAzureSetupStatus -Message ("Reusing existing SmartAzure sender shared mailbox '{0}'." -f $SenderAddress) -Level OK
        }
        return $mailbox
    }

    $alias = ($SenderAddress -split '@', 2)[0]
    if ($PSCmdlet.ShouldProcess($SenderAddress, 'Create SmartAzure sender shared mailbox')) {
        $mailbox = New-Mailbox `
            -Shared `
            -Name $SenderDisplayName `
            -DisplayName $SenderDisplayName `
            -Alias $alias `
            -PrimarySmtpAddress $SenderAddress `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Created SmartAzure sender shared mailbox '{0}'." -f $SenderAddress) -Level OK
    }

    return $mailbox
}

function Get-OrCreate-SmartAzureMailSendSecurityGroup {
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
        Write-SmartAzureSetupStatus -Message ("Reusing existing Mail.Send scope group '{0}' ({1})." -f $group.DisplayName, $group.PrimarySmtpAddress) -Level OK
        return $group
    }

    if ($PSCmdlet.ShouldProcess($GroupAddress, 'Create SmartAzure Mail.Send mail-enabled security group')) {
        $group = New-DistributionGroup `
            -Name $GroupName `
            -Alias $GroupAlias `
            -PrimarySmtpAddress $GroupAddress `
            -Type Security `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Created Mail.Send scope group '{0}' ({1})." -f $GroupName, $GroupAddress) -Level OK
    }

    return $group
}

function Add-SmartAzureSenderMailboxToScopeGroup {
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
        Write-SmartAzureSetupStatus -Message ("Sender mailbox '{0}' is already a member of '{1}'." -f $SenderAddress, $Group.DisplayName) -Level OK
        return
    }

    if ($PSCmdlet.ShouldProcess(("{0} -> {1}" -f $SenderAddress, $Group.DisplayName), 'Add SmartAzure sender mailbox to Mail.Send scope group')) {
        Add-DistributionGroupMember `
            -Identity $Group.Identity `
            -Member $SenderAddress `
            -BypassSecurityGroupManagerCheck `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Added sender mailbox '{0}' to Mail.Send scope group '{1}'." -f $SenderAddress, $Group.DisplayName) -Level OK
    }
}

function Set-SmartAzureMailSendApplicationAccessPolicy {
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

    $policies = @(Get-SmartAzureApplicationAccessPolicy)
    $matchingPolicy = $policies | Where-Object {
        $policy = $_
        @($_.AppId) -contains $AppId -and
        $_.AccessRight -eq 'RestrictAccess' -and
        ($scopeIdentifiers | Where-Object { $_.Equals([string]$policy.ScopeName, [System.StringComparison]::OrdinalIgnoreCase) })
    } | Select-Object -First 1

    if ($null -ne $matchingPolicy) {
        Write-SmartAzureSetupStatus -Message ("Application Access Policy already scopes Mail.Send for app {0} to '{1}'." -f $AppId, $matchingPolicy.ScopeName) -Level OK
        return $matchingPolicy
    }

    $otherRestrictPolicy = $policies | Where-Object {
        @($_.AppId) -contains $AppId -and $_.AccessRight -eq 'RestrictAccess'
    } | Select-Object -First 1
    if ($null -ne $otherRestrictPolicy) {
        Write-SmartAzureSetupStatus -Level WARN -Message ("App {0} already has an Application Access Policy scoped to '{1}'. A second policy will be added for '{2}'." -f $AppId, $otherRestrictPolicy.ScopeName, $GroupAddress)
    }

    if ($PSCmdlet.ShouldProcess($AppId, ("Create Application Access Policy scoped to {0}" -f $GroupAddress))) {
        $policy = New-ApplicationAccessPolicy `
            -AccessRight RestrictAccess `
            -AppId $AppId `
            -PolicyScopeGroupId $GroupAddress `
            -Description 'Restrict SmartAzure Graph Mail.Send to approved SmartAzure sender mailboxes.' `
            -ErrorAction Stop `
            -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Created Application Access Policy for app {0} scoped to '{1}'." -f $AppId, $GroupAddress) -Level OK
        return $policy
    }
}

function Get-SmartAzureApplicationAccessPolicy {
    try {
        return @(Get-ApplicationAccessPolicy -ErrorAction Stop)
    }
    catch {
        $message = [string]$_.Exception.Message
        if ($message -match 'introuvable|not found|not be found|ObjectNotFound|Cannot find|couldn''t be found') {
            Write-SmartAzureSetupStatus -Level WARN -Message 'No existing Exchange Online Application Access Policy was returned; continuing as if none exist.'
            return @()
        }

        throw
    }
}

function Remove-SmartAzureMailSendApplicationAccessPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$AppId
    )

    $policies = @(Get-SmartAzureApplicationAccessPolicy | Where-Object { @($_.AppId) -contains $AppId })
    if ($policies.Count -eq 0) {
        Write-SmartAzureSetupStatus -Message ("No Application Access Policy found for app {0}." -f $AppId) -Level OK
        return
    }

    foreach ($policy in $policies) {
        if ($PSCmdlet.ShouldProcess($policy.Identity, ("Remove Application Access Policy for app {0}" -f $AppId))) {
            Remove-ApplicationAccessPolicy -Identity $policy.Identity -Confirm:$false -ErrorAction Stop
            Write-SmartAzureSetupStatus -Message ("Removed Application Access Policy '{0}' for app {1}." -f $policy.Identity, $AppId) -Level OK
        }
    }
}

function Test-SmartAzureMailSendApplicationAccessPolicy {
    param(
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$SenderAddress
    )

    if (-not (Get-Command -Name Test-ApplicationAccessPolicy -ErrorAction SilentlyContinue)) {
        Write-SmartAzureSetupStatus -Level WARN -Message 'Test-ApplicationAccessPolicy is not available in this Exchange Online session; skipping Mail.Send scope test.'
        return
    }

    try {
        $testResult = Test-ApplicationAccessPolicy -Identity $SenderAddress -AppId $AppId -ErrorAction Stop
        Write-SmartAzureSetupStatus -Message ("Application Access Policy test for '{0}' returned: {1}." -f $SenderAddress, $testResult.AccessCheckResult) -Level OK
    }
    catch {
        Write-SmartAzureSetupStatus -Level WARN -Message ("Application Access Policy test failed for '{0}': {1}" -f $SenderAddress, $_.Exception.Message)
    }
}

function Set-SmartAzureMailLocalConfig {
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

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Update SmartAzure mail local configuration')) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Updated mail settings in {0}." -f $ConfigPath) -Level OK
    }
}

function Set-SmartAzureExchangeMailSendSetup {
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

    Connect-SmartAzureExchangeOnlineSetupSession -UserPrincipalName $AdminUserPrincipalName
    $effectiveSenderAddress = Resolve-SmartAzureMailAddress -ConfiguredAddress $SenderAddress -DefaultAlias 'smartazure-reports'
    $domain = ($effectiveSenderAddress -split '@', 2)[1]
    $scopeGroupAddress = ('{0}@{1}' -f $ScopeGroupAlias, $domain)

    $mailbox = Get-OrCreate-SmartAzureSenderMailbox -SenderAddress $effectiveSenderAddress -SenderDisplayName $SenderDisplayName
    if ($null -eq $mailbox) {
        Write-SmartAzureSetupStatus -Level WARN -Message 'SmartAzure sender mailbox was not created because WhatIf was used.'
    }

    $group = Get-OrCreate-SmartAzureMailSendSecurityGroup -GroupName $ScopeGroupName -GroupAlias $ScopeGroupAlias -GroupAddress $scopeGroupAddress
    if ($null -eq $group) {
        Write-SmartAzureSetupStatus -Level WARN -Message 'SmartAzure Mail.Send scope group was not created because WhatIf was used.'
    }
    else {
        Add-SmartAzureSenderMailboxToScopeGroup -Group $group -SenderAddress $effectiveSenderAddress
        Set-SmartAzureMailSendApplicationAccessPolicy -AppId $Application.AppId -Group $group -GroupAddress $scopeGroupAddress | Out-Null
        Test-SmartAzureMailSendApplicationAccessPolicy -AppId $Application.AppId -SenderAddress $effectiveSenderAddress
    }

    Set-SmartAzureMailLocalConfig -ConfigPath $ConfigPath -SenderAddress $effectiveSenderAddress -ScopeGroupAddress $scopeGroupAddress

    return [pscustomobject]@{
        SenderAddress    = $effectiveSenderAddress
        ScopeGroupName   = $ScopeGroupName
        ScopeGroupAddress = $scopeGroupAddress
    }
}

function Set-SmartAzureSharePointLocalConfig {
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

function Set-SmartAzureAuthLocalConfig {
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

    if ($PSCmdlet.ShouldProcess($ConfigPath, 'Update SmartAzure app authentication local configuration')) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
    }
}

function Clear-SmartAzureAuthLocalConfig {
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
        Write-SmartAzureSetupStatus -Message ("No app authentication settings to clear in {0}." -f $ConfigPath) -Level OK
        return
    }

    if ($PSCmdlet.ShouldProcess($ConfigPath, ("Clear SmartAzure app authentication settings: {0}" -f ($removedProperties -join ', ')))) {
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Cleared app authentication settings in {0}: {1}." -f $ConfigPath, ($removedProperties -join ', ')) -Level OK
    }
}

function Remove-SmartAzureAppRegistration {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string]$ConfigPath
    )

    $servicePrincipals = @(Get-MgServicePrincipal -Filter "appId eq '$($Application.AppId)'" -All)

    foreach ($servicePrincipal in $servicePrincipals) {
        if ($PSCmdlet.ShouldProcess(("Service principal {0} ({1})" -f $servicePrincipal.DisplayName, $servicePrincipal.Id), 'Remove SmartAzure application service principal')) {
            Remove-MgServicePrincipal -ServicePrincipalId $servicePrincipal.Id -Confirm:$false
            Write-SmartAzureSetupStatus -Message ("Removed application service principal '{0}' ({1})." -f $servicePrincipal.DisplayName, $servicePrincipal.Id) -Level OK
        }
    }

    if ($PSCmdlet.ShouldProcess(("App registration {0} ({1})" -f $Application.DisplayName, $Application.Id), 'Remove SmartAzure app registration')) {
        Remove-MgApplication -ApplicationId $Application.Id -Confirm:$false
        Write-SmartAzureSetupStatus -Message ("Removed app registration '{0}' ({1})." -f $Application.DisplayName, $Application.Id) -Level OK
    }

    Clear-SmartAzureAuthLocalConfig -ConfigPath $ConfigPath
    Write-SmartAzureSetupStatus -Level WARN -Message 'Teams workspace, SharePoint files, and local certificates were not removed.'
}

function Get-SmartAzureLocalConfigCertificateThumbprint {
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
        Write-SmartAzureSetupStatus -Level WARN -Message ("Could not read local config for certificate reuse: {0}" -f $_.Exception.Message)
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
            Write-SmartAzureSetupStatus -Level WARN -Message ("Local config certificate thumbprint '{0}' was not found in CurrentUser\My or LocalMachine\My; falling back to app key lookup." -f $normalizedThumbprint)
            return $null
        }

        if (-not $certificate.HasPrivateKey) {
            Write-SmartAzureSetupStatus -Level WARN -Message ("Local config certificate thumbprint '{0}' exists but has no private key; falling back to app key lookup." -f $normalizedThumbprint)
            return $null
        }

        Write-SmartAzureSetupStatus -Message ("Reusing certificate thumbprint from local tenant config: {0}" -f $normalizedThumbprint) -Level OK
        return $normalizedThumbprint
    }

    return $null
}

Initialize-SmartAzureSetupLogging -Path $LogPath

try {
Import-RequiredGraphModule
Disconnect-SmartAzureExistingGraphSession
Disconnect-SmartAzureExistingExchangeOnlineSession
if (-not $DisableMailSendScopeSetup) {
    Connect-SmartAzureExchangeOnlineSetupSession -UserPrincipalName $ExchangeAdminUserPrincipalName
}
$graphContext = Connect-SmartAzureGraphSetupSession -RequestedTenantId $TenantId -UseDeviceCode:$UseDeviceCode
$effectiveTenantId = $graphContext.TenantId
$localConfigPath = Join-Path -Path $script:SmartAzureRootPath -ChildPath ("Config\Tenants\{0}.local.json" -f $Tenant)

$escapedDisplayName = ConvertTo-ODataStringLiteral -Value $DisplayName
$existingApps = @(Get-MgApplication -Filter "displayName eq '$escapedDisplayName'" -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials' -All)
$application = $null

if ($existingApps.Count -gt 1) {
    throw "Multiple app registrations named '$DisplayName' were found. Rename duplicates or use a unique -DisplayName."
}

if ($RemoveAppRegistration) {
    if ($existingApps.Count -eq 0) {
        Write-SmartAzureSetupStatus -Level WARN -Message ("No app registration named '{0}' was found. Nothing to remove." -f $DisplayName)
        return
    }

    if (-not $DisableMailSendScopeSetup) {
        try {
            Connect-SmartAzureExchangeOnlineSetupSession -UserPrincipalName $graphContext.Account
            Remove-SmartAzureMailSendApplicationAccessPolicy -AppId $existingApps[0].AppId
        }
        catch {
            Write-SmartAzureSetupStatus -Level WARN -Message ("Exchange Online Mail.Send policy cleanup failed and app removal will continue: {0}" -f $_.Exception.Message)
        }
    }
    else {
        Write-SmartAzureSetupStatus -Level WARN -Message 'Exchange Online Mail.Send policy cleanup was skipped because -DisableMailSendScopeSetup was used.'
    }

    Remove-SmartAzureAppRegistration -Application $existingApps[0] -ConfigPath $localConfigPath
    return
}

$requiredApiResources = Get-SmartAzureRequiredApiResource
$obsoleteBroadSharePointPermissions = @('Files.ReadWrite.All', 'Sites.ReadWrite.All')

$requiredResourceAccess = Get-RequiredResourceAccessBlock -RequiredApiResource $requiredApiResources

$existingApplicationForCertificate = if ($existingApps.Count -eq 1) { $existingApps[0] } else { $null }
$effectiveCertificateThumbprint = if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    $CertificateThumbprint
}
else {
    Get-SmartAzureLocalConfigCertificateThumbprint -ConfigPath $localConfigPath
}
$certificate = Get-SmartAzureCertificate -Thumbprint $effectiveCertificateThumbprint -AppDisplayName $DisplayName -ValidityYears $CertificateYears -Application $existingApplicationForCertificate

if ($existingApps.Count -eq 1) {
    if (-not $UpdateExisting) {
        throw "App registration '$DisplayName' already exists. Re-run with -UpdateExisting to merge SmartAzure permissions and certificate."
    }

    $application = $existingApps[0]
    $mergedRequiredResourceAccess = Merge-RequiredResourceAccess -ExistingAccess $application.RequiredResourceAccess -RequiredAccess $requiredResourceAccess
    $mergedRequiredResourceAccess = Get-SmartAzureRequiredResourceAccessWithoutPermission `
        -RequiredAccess $mergedRequiredResourceAccess `
        -ResourceName 'Microsoft Graph' `
        -ResourceAppId '00000003-0000-0000-c000-000000000000' `
        -PermissionValues $obsoleteBroadSharePointPermissions
    $mergedRequiredResourceAccess = ConvertTo-SmartAzureRequiredResourceAccessPayload -RequiredAccess $mergedRequiredResourceAccess
    $mergedKeyCredentials = Add-CertificateToApplication -Application $application -Certificate $certificate

    if ($PSCmdlet.ShouldProcess($DisplayName, 'Update app registration permissions and certificate')) {
        Update-MgApplication `
            -ApplicationId $application.Id `
            -RequiredResourceAccess $mergedRequiredResourceAccess `
            -KeyCredentials $mergedKeyCredentials `
            -ErrorAction Stop
        $application = Get-MgApplication -ApplicationId $application.Id -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials'
        Write-SmartAzureSetupStatus -Message ("Updated app registration '{0}'." -f $DisplayName) -Level OK
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
        Write-SmartAzureSetupStatus -Message ("Created app registration '{0}'." -f $DisplayName) -Level OK
    }
}

if ($null -eq $application) {
    Write-SmartAzureSetupStatus -Level WARN -Message 'No app registration changes were applied because WhatIf was used.'
    return
}

$application = Get-MgApplication -ApplicationId $application.Id -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials'
if (-not (Test-ApplicationHasCertificate -Application $application -Certificate $certificate)) {
    throw ("Certificate public key was not found on app registration '{0}' after create/update. Refusing to write unusable app-only configuration." -f $DisplayName)
}
Write-SmartAzureSetupStatus -Message 'Certificate public key verified on the app registration.' -Level OK

Set-SmartAzureAuthLocalConfig -Application $application -Certificate $certificate -TenantId $effectiveTenantId -ConfigPath $localConfigPath
if (-not $WhatIfPreference) {
    Write-SmartAzureSetupStatus -Message ("Updated app authentication settings in {0}." -f $localConfigPath) -Level OK
}

$appServicePrincipal = @(Get-MgServicePrincipal -Filter "appId eq '$($application.AppId)'" -All) | Select-Object -First 1
if ($null -eq $appServicePrincipal) {
    if ($PSCmdlet.ShouldProcess($DisplayName, 'Create application service principal')) {
        $appServicePrincipal = New-MgServicePrincipal -AppId $application.AppId
        Write-SmartAzureSetupStatus -Message 'Created application service principal.' -Level OK
    }
}

if (-not $DisableGrantAdminConsent -and $null -ne $appServicePrincipal) {
    Grant-SmartAzureApplicationPermission -ApplicationServicePrincipal $appServicePrincipal -RequiredApiResource $requiredApiResources
    Revoke-SmartAzureApplicationPermission `
        -ApplicationServicePrincipal $appServicePrincipal `
        -ResourceName 'Microsoft Graph' `
        -ResourceAppId '00000003-0000-0000-c000-000000000000' `
        -PermissionValues $obsoleteBroadSharePointPermissions
}
elseif ($DisableGrantAdminConsent) {
    Write-SmartAzureSetupStatus -Level WARN -Message 'Admin consent was not granted because -DisableGrantAdminConsent was used. Use the admin consent URL printed below if needed.'
}

$mailSendSetup = $null
if (-not $DisableMailSendScopeSetup) {
    $mailSendSetup = Set-SmartAzureExchangeMailSendSetup `
        -Application $application `
        -AdminUserPrincipalName $graphContext.Account `
        -ConfigPath $localConfigPath `
        -SenderAddress $MailSenderAddress `
        -SenderDisplayName $MailSenderDisplayName `
        -ScopeGroupName $MailSendSecurityGroupName `
        -ScopeGroupAlias $MailSendSecurityGroupAlias
}
else {
    Write-SmartAzureSetupStatus -Level WARN -Message 'Mail.Send shared mailbox, group, and Application Access Policy setup were skipped because -DisableMailSendScopeSetup was used.'
}

$teamsWorkspace = $null
if (-not $DisableTeamsSetup) {
    $teamsWorkspace = Get-OrCreate-SmartAzureTeamsWorkspace `
        -GroupDisplayName $TeamDisplayName `
        -MailNickname $TeamMailNickname `
        -OwnerUserPrincipalName $graphContext.Account `
        -TargetFolderPath $SharePointTargetFolderPath `
        -DefaultLibraryDisplayName 'Documents'

    Set-SmartAzureSharePointLocalConfig -WorkspaceInfo $teamsWorkspace -ConfigPath $localConfigPath
    if (-not $WhatIfPreference) {
        Write-SmartAzureSetupStatus -Message ("Updated SharePoint settings in {0}." -f $localConfigPath) -Level OK
    }

    Grant-SmartAzureSelectedSitePermission `
        -SiteId $teamsWorkspace.SharePointSiteId `
        -ApplicationAppId $application.AppId `
        -ApplicationDisplayName $DisplayName `
        -Role 'write'
}
else {
    Write-SmartAzureSetupStatus -Level WARN -Message 'Teams workspace creation and SharePoint local configuration update were skipped because -DisableTeamsSetup was used.'
}

$adminConsentUrl = "https://login.microsoftonline.com/$effectiveTenantId/adminconsent?client_id=$($application.AppId)"

Write-Output ''
Write-SmartAzureSetupStatus -Message 'SmartAzure app registration summary' -Level OK
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
}
finally {
    Close-SmartAzureExchangeOnlineSetupSession
    Close-SmartAzureSetupLogging
}
