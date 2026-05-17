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
    resolves its backing SharePoint site, and updates SmartM365.global.local.json
    with the SharePoint upload target.

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

.PARAMETER DisableGrantAdminConsent
    Adds the requested API permissions without granting tenant-wide admin
    consent.

.PARAMETER UseDeviceCode
    Uses interactive device code authentication for Microsoft Graph setup
    sign-in. Useful when browser/WAM authentication is not available in the
    current terminal.

.PARAMETER TeamDisplayName
    Display name of the Teams team used for SmartM365 shared exports. Defaults
    to SMART-M365.

.PARAMETER DisableTeamsSetup
    Skips Teams team creation/reuse and SharePoint local configuration updates.

.PARAMETER LogPath
    Folder where the setup log and transcript are written. Defaults to
    C:\Temp\WORKPLACE.

.EXAMPLE
    .\SmartM365-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\SmartM365-Create-AppRegistration.ps1 -DisplayName SmartM365 -TenantId 00000000-0000-0000-0000-000000000000 -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UpdateExisting

.EXAMPLE
    .\SmartM365-Create-AppRegistration.ps1 -TenantId contoso.onmicrosoft.com -CertificateThumbprint 00112233445566778899AABBCCDDEEFF00112233 -UseDeviceCode -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
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
    [switch]$DisableGrantAdminConsent,

    [Parameter()]
    [switch]$UseDeviceCode,

    [Parameter()]
    [switch]$SkipExchangeOnlinePermission,

    [Parameter()]
    [switch]$SkipBroadIntuneReadWritePermissions,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamDisplayName = 'SMART-M365',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TeamMailNickname = 'SMARTM365',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SharePointTargetFolderPath = 'SMART-M365/CSV',

    [Parameter()]
    [switch]$DisableTeamsSetup,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = 'C:\Temp\WORKPLACE'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SmartM365SetupLogFile = $null
$script:SmartM365SetupTranscriptFile = $null
$script:SmartM365SetupTranscriptStarted = $false

function Initialize-SmartM365SetupLogging {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $resolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedLogPath)) {
        New-Item -Path $resolvedLogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $baseName = 'SmartM365-Create-AppRegistration'
    $script:SmartM365SetupLogFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}.log" -f $baseName, $timestamp)
    $script:SmartM365SetupTranscriptFile = Join-Path -Path $resolvedLogPath -ChildPath ("{0}-{1}_Transcript.log" -f $baseName, $timestamp)

    Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value ("[{0}] [INFO] Log started." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

    try {
        Start-Transcript -Path $script:SmartM365SetupTranscriptFile -Append -ErrorAction Stop | Out-Null
        $script:SmartM365SetupTranscriptStarted = $true
    }
    catch {
        Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value ("[{0}] [WARN] Transcript could not be started: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $_.Exception.Message)
    }
}

function Close-SmartM365SetupLogging {
    if ($script:SmartM365SetupLogFile) {
        Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value ("[{0}] [INFO] Log finished." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        Write-Information ("[INFO] Log file: {0}" -f $script:SmartM365SetupLogFile) -InformationAction Continue
        if ($script:SmartM365SetupTranscriptFile) {
            Write-Information ("[INFO] Transcript file: {0}" -f $script:SmartM365SetupTranscriptFile) -InformationAction Continue
        }
    }

    if ($script:SmartM365SetupTranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch { Write-Debug ("Stop-Transcript failed: {0}" -f $_.Exception.Message) }
        $script:SmartM365SetupTranscriptStarted = $false
    }
}

function Write-SmartM365SetupStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][string]$Level = 'INFO'
    )

    if ($script:SmartM365SetupLogFile) {
        $logEntry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        try { Add-Content -LiteralPath $script:SmartM365SetupLogFile -Encoding UTF8 -Value $logEntry } catch { Write-Debug ("Log write failed: {0}" -f $_.Exception.Message) }
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

function Connect-SmartM365GraphSetupSession {
    param(
        [string]$RequestedTenantId,
        [switch]$UseDeviceCode
    )

    $scopes = @(
        'Application.ReadWrite.All',
        'AppRoleAssignment.ReadWrite.All',
        'Directory.Read.All',
        'Group.ReadWrite.All',
        'Sites.Read.All'
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

function Get-SmartM365RequiredApiResource {
    param(
        [switch]$SkipExchange,
        [switch]$SkipBroadReadWrite
    )

    $graphPermissions = @(
        'Directory.Read.All',
        'User.Read.All',
        'Device.Read.All',
        'GroupMember.Read.All',
        'DeviceManagementApps.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'DeviceManagementScripts.Read.All',
        'DeviceManagementServiceConfig.Read.All',
        'Files.ReadWrite.All',
        'Mail.Send',
        'Sites.ReadWrite.All'
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

    return @{
        customKeyIdentifier = [System.Convert]::ToBase64String($Certificate.GetCertHash())
        displayName         = $Certificate.Subject
        endDateTime         = $Certificate.NotAfter
        key                 = $Certificate.RawData
        startDateTime       = $Certificate.NotBefore
        type                = 'AsymmetricX509Cert'
        usage               = 'Verify'
    }
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

    $site = Wait-SmartM365TeamSharePointSite -GroupId $group.id
    $siteUri = [System.Uri]$site.webUrl
    $libraryDisplayName = Get-SmartM365GroupDriveName -GroupId $group.id -DefaultLibraryDisplayName $DefaultLibraryDisplayName

    return [pscustomobject]@{
        TeamId                       = $group.id
        TeamDisplayName              = $GroupDisplayName
        SharePointSiteHostname       = $siteUri.Host
        SharePointSitePath           = $siteUri.AbsolutePath
        SharePointLibraryDisplayName = $libraryDisplayName
        SharePointTargetFolderPath   = $TargetFolderPath
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
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
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
        $config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    }
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

        Write-SmartM365SetupStatus -Message ("Reusing certificate thumbprint from SmartM365.global.local.json: {0}" -f $normalizedThumbprint) -Level OK
        return $normalizedThumbprint
    }

    return $null
}

Initialize-SmartM365SetupLogging -Path $LogPath

try {
Import-RequiredGraphModule
$graphContext = Connect-SmartM365GraphSetupSession -RequestedTenantId $TenantId -UseDeviceCode:$UseDeviceCode
$effectiveTenantId = $graphContext.TenantId
$localConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.global.local.json'

$requiredApiResources = Get-SmartM365RequiredApiResource `
    -SkipExchange:$SkipExchangeOnlinePermission `
    -SkipBroadReadWrite:$SkipBroadIntuneReadWritePermissions

if (-not $SkipBroadIntuneReadWritePermissions) {
    Write-SmartM365SetupStatus -Level WARN -Message 'Including broad Intune ReadWrite application permissions because one current Autopatch report script requests them. Use -SkipBroadIntuneReadWritePermissions after hardening those scripts.'
}

$requiredResourceAccess = Get-RequiredResourceAccessBlock -RequiredApiResource $requiredApiResources

$escapedDisplayName = ConvertTo-ODataStringLiteral -Value $DisplayName
$existingApps = @(Get-MgApplication -Filter "displayName eq '$escapedDisplayName'" -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials' -All)
$application = $null

if ($existingApps.Count -gt 1) {
    throw "Multiple app registrations named '$DisplayName' were found. Rename duplicates or use a unique -DisplayName."
}

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
    $mergedKeyCredentials = Add-CertificateToApplication -Application $application -Certificate $certificate

    if ($PSCmdlet.ShouldProcess($DisplayName, 'Update app registration permissions and certificate')) {
        Update-MgApplication -ApplicationId $application.Id -BodyParameter @{
            requiredResourceAccess = $mergedRequiredResourceAccess
            keyCredentials         = $mergedKeyCredentials
        }
        $application = Get-MgApplication -ApplicationId $application.Id -Property 'id,appId,displayName,requiredResourceAccess,keyCredentials'
        Write-SmartM365SetupStatus -Message ("Updated app registration '{0}'." -f $DisplayName) -Level OK
    }
}
else {
    $body = @{
        displayName            = $DisplayName
        signInAudience         = 'AzureADMyOrg'
        requiredResourceAccess = $requiredResourceAccess
        keyCredentials         = @(Get-KeyCredentialFromCertificate -Certificate $certificate)
    }

    if ($PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) {
        $application = New-MgApplication -BodyParameter $body
        Write-SmartM365SetupStatus -Message ("Created app registration '{0}'." -f $DisplayName) -Level OK
    }
}

if ($null -eq $application) {
    Write-SmartM365SetupStatus -Level WARN -Message 'No app registration changes were applied because WhatIf was used.'
    return
}

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
}
elseif ($DisableGrantAdminConsent) {
    Write-SmartM365SetupStatus -Level WARN -Message 'Admin consent was not granted because -DisableGrantAdminConsent was used. Use the admin consent URL printed below if needed.'
}

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
    Write-Output ("SP Host     : {0}" -f $teamsWorkspace.SharePointSiteHostname)
    Write-Output ("SP Path     : {0}" -f $teamsWorkspace.SharePointSitePath)
    Write-Output ("SP Library  : {0}" -f $teamsWorkspace.SharePointLibraryDisplayName)
    Write-Output ("SP Folder   : {0}" -f $teamsWorkspace.SharePointTargetFolderPath)
}
Write-Output ''
Write-Output 'Add these values to SmartM365.global.local.json:'
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
    Write-Output 'SharePoint settings written to SmartM365.global.local.json:'
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
Write-SmartM365SetupStatus -Level WARN -Message 'Exchange Online still needs RBAC for the application service principal, for example the relevant read-only Exchange role group or RBAC for Applications assignment.'
}
finally {
    Close-SmartM365SetupLogging
}
