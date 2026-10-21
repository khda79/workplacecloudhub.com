<#
.SYNOPSIS
    Child-process runner of the Exchange hybrid free/busy inventory.

.DESCRIPTION
    Started by SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1 from its protected work folder, in its
    own Windows PowerShell process so a hung run can be stopped and Free/Busy Checker's own exit calls
    end this process only. It:
    - checks the ActiveDirectory module (Windows feature, never installed here);
    - loads the Exchange Management Shell (local Exchange server session, cmdlets without prefix),
      which Free/Busy Checker expects and does not load itself;
    - makes ExchangeOnlineManagement available (installed for the current user from PSGallery when missing);
    - opens the Exchange Online session with prefix EO using the tenant app and its certificate only,
      and proves it is that app-only certificate session; it never signs in interactively;
    - records which Exchange Online cmdlets Free/Busy Checker calls are not exposed to the app;
    - verifies the Free/Busy Checker copy (SHA256 and Microsoft signature) and runs it with
      -SkipVersionCheck only.
    Progress and results are written to the status file passed by the collector.

.VERSION
1.0.3
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
$RunnerVersion = '1.0.3'
# Exchange on-premises cmdlets Free/Busy Checker needs from the Exchange Management Shell.
$FreeBusyExchangeShellCmdlets = @('Get-RemoteMailbox', 'Get-OrganizationRelationship', 'Test-OAuthConnectivity')
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

function Import-FreeBusyExchangeShell {
    param([Parameter(Mandatory)][string[]]$RequiredCmdlets)

    # Free/Busy Checker calls the Exchange on-premises cmdlets directly and never loads the shell itself.
    # RemoteExchange.ps1 and Connect-ExchangeServer are Exchange's own supported loader (local server
    # session through IIS); they are written for default settings, kept to this function scope.
    Set-StrictMode -Off
    $ErrorActionPreference = 'Continue'
    if (-not (Get-Command -Name $RequiredCmdlets[0] -ErrorAction SilentlyContinue)) {
        if ([string]::IsNullOrWhiteSpace($env:ExchangeInstallPath)) {
            return [pscustomobject]@{ Loaded = $false; Message = 'ExchangeInstallPath is not set: the Exchange Management Tools are not installed on this server.' }
        }
        $remoteExchange = Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1'
        if (-not (Test-Path -LiteralPath $remoteExchange -PathType Leaf)) {
            return [pscustomobject]@{ Loaded = $false; Message = "RemoteExchange.ps1 was not found: $remoteExchange" }
        }
        try {
            . $remoteExchange
            Connect-ExchangeServer -auto -ClientApplication:ManagementShell
        }
        catch {
            return [pscustomobject]@{ Loaded = $false; Message = "The Exchange Management Shell could not be loaded: $($_.Exception.Message)" }
        }
    }
    $missing = @($RequiredCmdlets | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        return [pscustomobject]@{ Loaded = $false; Message = 'Exchange Management Shell cmdlets are unavailable: ' + ($missing -join ', ') }
    }
    return [pscustomobject]@{ Loaded = $true; Message = $null }
}
function Invoke-FreeBusyCheckerScript {
    param([Parameter(Mandatory)][string]$Path)

    # Free/Busy Checker expects default session settings: strict mode off (with -SkipVersionCheck it reads
    # $versionsUrl, which it never sets) and non-terminating errors written to its report, not thrown.
    # Both settings apply to this function scope only; the runner keeps its own.
    Set-StrictMode -Off
    $ErrorActionPreference = 'Continue'
    # -SkipVersionCheck alone: any other parameter selects a parameter set that lets the script update itself.
    & $Path -SkipVersionCheck
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

    # On-premises cmdlets first, without prefix; the Exchange Online session that follows uses the EO prefix.
    $status.Stage = 'ExchangeShell'
    Save-FreeBusyRunnerStatus
    $shell = Import-FreeBusyExchangeShell -RequiredCmdlets $FreeBusyExchangeShellCmdlets
    if (-not $shell.Loaded) {
        $status.RunStatus = 'ExchangeShellNotLoaded'
        $status.Message = $shell.Message
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

    # The Exchange Online session must not hide the on-premises cmdlets.
    $hidden = @($FreeBusyExchangeShellCmdlets | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
    if ($hidden.Count -gt 0) {
        $status.RunStatus = 'ExchangeShellNotLoaded'
        $status.Message = 'Exchange Management Shell cmdlets are unavailable after the Exchange Online connection: ' + ($hidden -join ', ')
        return
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
    Invoke-FreeBusyCheckerScript -Path $FreeBusyCheckerPath
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDTLWhxEFFoEx15
# ajyBZB/vtZfuGyelgX0sXbcAAS201aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBD1kt/iGgMvZRjEaQR+KgpH0MUEDGR9SqAZGQwKv7WDMA0GCSqG
# SIb3DQEBAQUABIIBgDcoO6nobPY4Der80r6iUYCd/U3rVhS9fnPG2J7azEtnRc37
# CyeCJKvavz4vpJw2oVpZLQLDym1ZX8oISfze/Gj9JdSkUEpGJfwYKQ86bmJ/V5ZK
# qHtS4zyGRg8RV74xY7gH5TGQrCWCmCNYNiL1tWE4ztPm8gVfQpkle9KIrqrHt3fl
# F47RYxOI++6p2xUPg7y3nE5jWmKBdcriirKxHLKs4N6JphFWfWPdA45wGY8g3SLl
# zb7O26RHLOz11w4L+7wzVCIo2R+9I70EjrTd9H+BkG+qZymYcFhjD0U6y08OvZ+i
# liWSFH7/QIrSGPXyfeqPG2e/gnFsYFbVnO65yG/cHZFhrbWSKSsKnZfXhH0tlovg
# czjrIC6M82zY6ZmPXGM1XkcqUKGmLuwkeap0zwR1abKQGtv7X+VBj3FCUZs2gT4g
# r7kjch9vmFhGEF2gjIzxqp8DwYjx2aQtpFdNDqh2PFBTS1bDqHbyEoBUXzg07RcU
# gwVwzc8IsTmNBoKg+qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEyMTQy
# MzJaMC8GCSqGSIb3DQEJBDEiBCDDbd70dzu9Rf5Wx7FKsLKELNw/tQWiWhEnAiwv
# JqwwuzANBgkqhkiG9w0BAQEFAASCAgBDl6QS5cqPwUQ8SmdU8LT8gnbFH8X4phHx
# 5H2urHeh9fidkmPUp9DK3bmS3nsA3KHK04SM28LyRULcreXxUSTAG6L4HtxP4SdJ
# Z/F4P8GuQHjoQJ/4y5TP0HezuIm72XXEPwSnxCto4LPoM1JlORA0H4oagjuqxRSX
# zs9zZdgtIUplBAiYouu4/TivZF3kz1uI8xoVy3jQzcS56aWXQ11P7r6nLRU3W1y4
# voKubnX2q4v/JJIRnKqbe3tllY4cfohzWX+qAcvFljarHc3nszbkJDLU9SD8hTha
# HSs15V8CAJkYhpcdwECGDy8OwXhxy9d44rhCLLs81WIAc+9PUwOWo4hJMQ+2tL17
# KBW1+7xa3np71bRD0E/Op/3CN3tIbVvEcED0p6yrlbd0GwrHJ8GMhTxu8J6S6y0f
# cC0aXxvnwYb1R4GFOlax+K3dsT03tA25BH/QPKd+lrg1Jt8VEq25Qik27uqmvS0i
# G83se3vkLxVH44m4MLlr0p9re/UJsn87hd8GEPZg2bxdy8dvj2SaCEE0zb5jAUgI
# Lz3bLc7hfWKlINvjTEFsVmfr2r098KLxEJjS17V2QrU4jLgi7xJeiDUos+s26rF+
# tYqOtRV3pn8o25kP0aDJAPddPT6fZG//Gu1CoDDDkNDwbbQupI4ei0qj6aUJFqZw
# xkr4L8A+rA==
# SIG # End signature block
