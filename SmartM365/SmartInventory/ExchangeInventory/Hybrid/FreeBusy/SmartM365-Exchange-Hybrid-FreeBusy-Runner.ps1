<#
.SYNOPSIS
    Child-process runner of the Exchange hybrid free/busy inventory.

.DESCRIPTION
    Started by SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1 from its protected work folder, in its
    own Windows PowerShell process so a hung run can be stopped and Free/Busy Checker's own exit calls
    end this process only. It:
    - checks the ActiveDirectory module (Windows feature, never installed here);
    - makes ExchangeOnlineManagement available (installed for the current user from PSGallery when missing);
    - opens the Exchange Online session with prefix EO using the tenant app and its certificate only,
      and proves it is that app-only certificate session; it never signs in interactively;
    - records which Exchange Online cmdlets Free/Busy Checker calls are not exposed to the app;
    - verifies the Free/Busy Checker copy (SHA256 and Microsoft signature) and runs it with
      -SkipVersionCheck only.
    Progress and results are written to the status file passed by the collector.

.VERSION
1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','StatusPath',Justification='Read by Save-FreeBusyRunnerStatus from the script scope.')]
param(
    [Parameter(Mandatory)][string]$FreeBusyCheckerPath,
    [Parameter(Mandatory)][string]$ExpectedSha256,
    [Parameter(Mandatory)][string]$StatusPath,
    [string]$MinimumExchangeOnlineModuleVersion = '3.0.0'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$RunnerVersion = '1.0.0'
# Exchange Online cmdlets called by Free/Busy Checker with the EO prefix.
$FreeBusyExchangeOnlineCmdlets = @('Get-EOMailbox', 'Get-EOOrganizationRelationship', 'Get-EOFederatedOrganizationIdentifier', 'Get-EOSharingPolicy', 'Test-EOOrganizationRelationship', 'Get-EOIntraOrganizationConnector', 'Get-EOIntraOrganizationConfiguration', 'Get-EOOnPremisesOrganization', 'Get-EOAuthServer', 'Test-EOOAuthConnectivity')

$status = [ordered]@{
    RunnerVersion               = $RunnerVersion
    Stage                       = 'Started'
    RunStatus                   = $null
    Message                     = $null
    ExchangeOnline              = 'NotConnected'
    ExchangeOnlineMessage       = $null
    NotAuthorizedCmdlets        = @()
    ExchangeOnlineModuleVersion = $null
    UpdatedUtc                  = $null
}

function Save-FreeBusyRunnerStatus {
    $status.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    $temporaryPath = "$StatusPath.tmp"
    [System.IO.File]::WriteAllText($temporaryPath, ($status | ConvertTo-Json -Depth 4), [System.Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporaryPath -Destination $StatusPath -Force
}

function Test-FreeBusyCheckerIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256
    )

    $result = [pscustomobject]@{ IsValid = $false; Reason = $null; Sha256 = $null; SignatureStatus = $null; Signer = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "Free/Busy Checker script not found: $Path"
        return $result
    }

    $result.Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($result.Sha256 -ne $ExpectedSha256.ToUpperInvariant()) {
        $result.Reason = "SHA256 mismatch: expected $($ExpectedSha256.ToUpperInvariant()), found $($result.Sha256)."
        return $result
    }

    $signature = Get-AuthenticodeSignature -FilePath $Path
    $result.SignatureStatus = [string]$signature.Status
    if ($signature.SignerCertificate) { $result.Signer = [string]$signature.SignerCertificate.Subject }
    if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) {
        $result.Reason = "Authenticode signature is not valid (status $($signature.Status))."
        return $result
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $issuer = [string]$signature.SignerCertificate.Issuer
    if ($subject -notmatch '(^|,\s*)CN=Microsoft Corporation(,|$)' -or $subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or $issuer -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        $result.Reason = "Authenticode signer is not Microsoft Corporation: $subject"
        return $result
    }

    $result.IsValid = $true
    return $result
}

function Install-FreeBusyExchangeOnlineModule {
    param([Parameter(Mandatory)][version]$MinimumVersion)

    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if ($module -and $module.Version -ge $MinimumVersion) { return $module }
    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) { throw 'Install-Module is unavailable (PowerShellGet is missing).' }
    $previousProtocol = [Net.ServicePointManager]::SecurityProtocol
    try {
        [Net.ServicePointManager]::SecurityProtocol = $previousProtocol -bor [Net.SecurityProtocolType]::Tls12
        $nuget = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $nuget -or $nuget.Version -lt [version]'2.8.5.201') {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }
        Install-Module -Name ExchangeOnlineManagement -MinimumVersion $MinimumVersion -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
    }
    finally {
        [Net.ServicePointManager]::SecurityProtocol = $previousProtocol
    }
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module -or $module.Version -lt $MinimumVersion) { throw "ExchangeOnlineManagement $MinimumVersion or later is still not available after installation." }
    return $module
}

function Connect-FreeBusyExchangeOnline {
    param(
        [AllowEmptyString()][string]$AppId,
        [AllowEmptyString()][string]$CertificateThumbprint,
        [AllowEmptyString()][string]$Organization
    )

    if ([string]::IsNullOrWhiteSpace($AppId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or [string]::IsNullOrWhiteSpace($Organization)) {
        return [pscustomobject]@{ Connected = $false; Message = 'The tenant app settings (AppId, certificate thumbprint, organization) are not configured; no Exchange Online connection was attempted.' }
    }
    try {
        # App-only certificate authentication: this form never shows a sign-in window.
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -Prefix EO -ShowBanner:$false -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{ Connected = $false; Message = "Exchange Online app-only connection failed: $($_.Exception.Message)" }
    }
    $connection = @(Get-ConnectionInformation | Where-Object { $_.ModulePrefix -eq 'EO' -and $_.State -eq 'Connected' })
    if ($connection.Count -ne 1 -or [string]$connection[0].AppId -ne $AppId -or -not [bool]$connection[0].CertificateAuthentication) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Connected = $false; Message = 'The Exchange Online session is not the tenant app certificate session; it was closed and not used.' }
    }
    return [pscustomobject]@{ Connected = $true; Message = $null }
}

$connected = $false
try {
    Save-FreeBusyRunnerStatus
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        $status.RunStatus = 'ActiveDirectoryModuleMissing'
        $status.Message = 'The ActiveDirectory module is missing (Windows feature RSAT-AD-PowerShell). An administrator installs it with: Install-WindowsFeature RSAT-AD-PowerShell'
        return
    }

    $status.Stage = 'ExchangeOnlineModule'
    Save-FreeBusyRunnerStatus
    try {
        $module = Install-FreeBusyExchangeOnlineModule -MinimumVersion ([version]$MinimumExchangeOnlineModuleVersion)
        Import-Module -Name ExchangeOnlineManagement -MinimumVersion $MinimumExchangeOnlineModuleVersion -ErrorAction Stop
        foreach ($commandName in @('Connect-ExchangeOnline', 'Get-ConnectionInformation', 'Disconnect-ExchangeOnline')) {
            if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) { throw "ExchangeOnlineManagement command is unavailable: $commandName" }
        }
        $status.ExchangeOnlineModuleVersion = [string]$module.Version
    }
    catch {
        $status.RunStatus = 'ModuleInstallFailed'
        $status.Message = "ExchangeOnlineManagement could not be installed or loaded for the current user: $($_.Exception.Message)"
        return
    }

    $status.Stage = 'ExchangeOnlineConnection'
    Save-FreeBusyRunnerStatus
    # Identifiers come from the collector through the environment, not the command line.
    $connection = Connect-FreeBusyExchangeOnline -AppId ([string]$env:SMARTM365_FREEBUSY_APP_ID) -CertificateThumbprint ([string]$env:SMARTM365_FREEBUSY_CERT_THUMBPRINT) -Organization ([string]$env:SMARTM365_FREEBUSY_ORGANIZATION)
    if ($connection.Connected) {
        $connected = $true
        $missing = @($FreeBusyExchangeOnlineCmdlets | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
        $status.NotAuthorizedCmdlets = $missing
        if ($missing -contains 'Get-EOMailbox') {
            # Free/Busy Checker tests the session with Get-EOMailbox and skips Exchange Online without it.
            $status.ExchangeOnline = 'NotConnected'
            $status.ExchangeOnlineMessage = 'Get-EOMailbox is not available to the app: Free/Busy Checker cannot use the Exchange Online session.'
        }
        elseif ($missing.Count -gt 0) {
            $status.ExchangeOnline = 'NotAuthorizedPartial'
            $status.ExchangeOnlineMessage = 'Exchange Online cmdlets not available to the app: ' + ($missing -join ', ')
        }
        else {
            $status.ExchangeOnline = 'Connected'
        }
    }
    else {
        $status.ExchangeOnline = 'NotConnected'
        $status.ExchangeOnlineMessage = $connection.Message
    }

    # Checked again immediately before the start, so the file that runs is the one just verified.
    $integrity = Test-FreeBusyCheckerIntegrity -Path $FreeBusyCheckerPath -ExpectedSha256 $ExpectedSha256
    if (-not $integrity.IsValid) {
        $status.RunStatus = 'IntegrityFailed'
        $status.Message = "Free/Busy Checker was not run: $($integrity.Reason)"
        return
    }
    $status.Stage = 'FreeBusyCheckerStarted'
    Save-FreeBusyRunnerStatus
    Set-Location -LiteralPath (Split-Path -Path $FreeBusyCheckerPath -Parent)
    # -SkipVersionCheck alone: any other parameter selects a parameter set that lets the script update itself.
    & $FreeBusyCheckerPath -SkipVersionCheck
    $status.Stage = 'FreeBusyCheckerReturned'
}
catch {
    $status.RunStatus = 'RunnerFailed'
    $status.Message = "Free/Busy Checker runner failed: $($_.Exception.Message)"
}
finally {
    if ($connected) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
    Save-FreeBusyRunnerStatus
}
