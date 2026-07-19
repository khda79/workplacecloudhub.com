<#
.SYNOPSIS
Collects Microsoft Graph evidence for Smart Exchange Migration Readiness in an isolated process.

.VERSION
1.11.8
#>
#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$TenantId = '',
    [string]$ScopesPath = '',
    [string]$InputPath = '',
    [string]$OutputPath = '',
    [string]$ErrorPath = '',
    [string]$ProgressPath = '',
    [string]$LogPath = '',
    [ValidateSet('CurrentUser','Process')][string]$ContextScope = 'CurrentUser',
    [switch]$ForceAuthentication,
    [switch]$InspectContext,
    [switch]$ValidateOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$workerStartedAt = Get-Date
$workerRunId = if ([string]::IsNullOrWhiteSpace($LogPath)) { '-' } else { Split-Path (Split-Path -Parent $LogPath) -Leaf }

function Write-GraphWorkerLog {
    param([ValidateSet('DEBUG','INFO','WARN','ERROR','SUCCESS')][string]$Level,[string]$Message)
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    try {
        $directory = Split-Path -Parent $LogPath
        if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) { [void](New-Item -Path $directory -ItemType Directory -Force) }
        $elapsedMs = [math]::Round(((Get-Date) - $workerStartedAt).TotalMilliseconds)
        $line = "{0} [{1}] [PID={2}] [Run={3}] [Component=GraphWorker] [ElapsedMs={4}] {5}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),$Level,$PID,$workerRunId,$elapsedMs,($Message -replace '[\r\n]+',' | ')
        [IO.File]::AppendAllText($LogPath,"$line`r`n",[Text.UTF8Encoding]::new($false))
    }
    catch { $null = $_ }
}

function Format-GraphWorkerError {
    param([Parameter(Mandatory)]$ErrorRecord)
    $parts = [System.Collections.Generic.List[string]]::new()
    [void]$parts.Add("ExceptionType=$($ErrorRecord.Exception.GetType().FullName)")
    [void]$parts.Add("Message=$($ErrorRecord.Exception.Message)")
    if($ErrorRecord.FullyQualifiedErrorId){[void]$parts.Add("FullyQualifiedErrorId=$($ErrorRecord.FullyQualifiedErrorId)")}
    if($ErrorRecord.CategoryInfo){[void]$parts.Add("CategoryInfo=$($ErrorRecord.CategoryInfo)")}
    if($ErrorRecord.InvocationInfo){[void]$parts.Add("Command=$($ErrorRecord.InvocationInfo.MyCommand.Name); Script=$($ErrorRecord.InvocationInfo.ScriptName); Line=$($ErrorRecord.InvocationInfo.ScriptLineNumber); SourceLine=$(if($ErrorRecord.InvocationInfo.Line){$ErrorRecord.InvocationInfo.Line.Trim()}else{''})")}
    if($ErrorRecord.ScriptStackTrace){[void]$parts.Add("ScriptStackTrace=$($ErrorRecord.ScriptStackTrace)")}
    $inner=$ErrorRecord.Exception.InnerException;$depth=0
    while($inner -and $depth -lt 5){$depth++;[void]$parts.Add("InnerException${depth}=$($inner.GetType().FullName): $($inner.Message)");$inner=$inner.InnerException}
    return ($parts -join ' | ')
}

function Write-WorkerTextFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowEmptyString()][string]$Value = ''
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    [IO.File]::WriteAllText($Path, $Value, [Text.UTF8Encoding]::new($false))
}

try {
    Write-GraphWorkerLog -Level INFO -Message "Worker starting. Version=1.11.8; PowerShell=$($PSVersionTable.PSVersion); Computer=$env:COMPUTERNAME; TenantId=$TenantId; RuntimeInput=$InputPath; ContextScope=$ContextScope; ForceAuthentication=$ForceAuthentication; InspectContext=$InspectContext."
    $moduleTimer=[Diagnostics.Stopwatch]::StartNew()
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if ($InspectContext) {
        if ([string]::IsNullOrWhiteSpace($OutputPath) -or [string]::IsNullOrWhiteSpace($ErrorPath)) { throw 'Graph context inspection output and error paths are required.' }
        $existingContext = Get-MgContext
        [pscustomobject][ordered]@{
            Available = $null -ne $existingContext
            Usable = $null -ne $existingContext
            Account = if ($existingContext) { [string]$existingContext.Account } else { '' }
            TenantId = if ($existingContext) { [string]$existingContext.TenantId } else { '' }
            Organization = ''
            AuthType = if ($existingContext) { [string]$existingContext.AuthType } else { '' }
            ContextScope = if ($existingContext) { [string]$existingContext.ContextScope } else { '' }
            ClientId = if ($existingContext) { [string]$existingContext.ClientId } else { '' }
            Scopes = if ($existingContext) { @($existingContext.Scopes) } else { @() }
            Error = ''
        } | Export-Clixml -LiteralPath $OutputPath -Depth 5 -Force
        Write-GraphWorkerLog -Level INFO -Message "Context inspection complete. Available=$($null -ne $existingContext); Account=$(if($existingContext){$existingContext.Account}else{''}); Tenant=$(if($existingContext){$existingContext.TenantId}else{''})."
        exit 0
    }

    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    $moduleTimer.Stop()
    Write-GraphWorkerLog -Level SUCCESS -Message "Graph modules imported. DurationMs=$($moduleTimer.ElapsedMilliseconds); AuthenticationVersion=$((Get-Module Microsoft.Graph.Authentication).Version); UsersVersion=$((Get-Module Microsoft.Graph.Users).Version); DirectoryManagementVersion=$((Get-Module Microsoft.Graph.Identity.DirectoryManagement).Version)."

    if ($ValidateOnly) {
        foreach ($commandName in @('Connect-MgGraph', 'Get-MgUser', 'Get-MgUserLicenseDetail', 'Get-MgSubscribedSku', 'Get-MgOrganization')) {
            if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) { throw "Required Microsoft Graph command is unavailable: $commandName" }
        }
        'VALIDATION_OK SmartM365 Exchange Migration Readiness Graph worker v1.11.8'
        exit 0
    }

    foreach ($requiredPath in @($ScopesPath, $InputPath, $OutputPath, $ErrorPath, $ProgressPath)) {
        if ([string]::IsNullOrWhiteSpace($requiredPath)) { throw 'Graph worker runtime paths are required.' }
    }
    $scopes = @(Import-Clixml -LiteralPath $ScopesPath)
    $emailAddresses = @(Import-Clixml -LiteralPath $InputPath)
    Write-GraphWorkerLog -Level INFO -Message "Runtime input loaded. ScopeCount=$($scopes.Count); MailboxCount=$($emailAddresses.Count); Scopes=$($scopes -join ',')."
    $existingContext = Get-MgContext
    if ($ForceAuthentication -and $existingContext -and (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-GraphWorkerLog -Level INFO -Message "Forcing new authentication; disconnecting existing account=$($existingContext.Account), tenant=$($existingContext.TenantId)."
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        $existingContext = $null
    }
    $requiredScopeSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($scope in $scopes) { [void]$requiredScopeSet.Add([string]$scope) }
    $missingScopes = if ($existingContext) { @($requiredScopeSet | Where-Object { @($existingContext.Scopes) -notcontains $_ }) } else { @($requiredScopeSet) }
    $tenantCompatible = $existingContext -and (-not $TenantId -or [string]$existingContext.TenantId -eq $TenantId)
    $contextCompatible = $tenantCompatible -and $missingScopes.Count -eq 0
    $authenticationTimer=[Diagnostics.Stopwatch]::StartNew()
    if ($contextCompatible -and -not $ForceAuthentication) {
        $context = $existingContext
        Write-GraphWorkerLog -Level SUCCESS -Message "Reusing existing delegated Graph context. Account=$($context.Account); Tenant=$($context.TenantId); ContextScope=$($context.ContextScope)."
    }
    else {
        if ($existingContext -and (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue)) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
        $connectParameters = @{ ContextScope=$ContextScope; NoWelcome=$true; ErrorAction='Stop'; Scopes=$scopes }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParameters.TenantId = $TenantId }
        Write-GraphWorkerLog -Level INFO -Message "Interactive Microsoft Graph authentication starting. MissingScopes=$($missingScopes -join ','); TenantCompatible=$tenantCompatible."
        Connect-MgGraph @connectParameters | Out-Null
        $context = Get-MgContext
    }
    $authenticationTimer.Stop()
    if (-not $context) {
        throw 'Microsoft Graph authentication did not return a context.'
    }
    if ($TenantId -and [string]$context.TenantId -ne $TenantId) {
        throw "Microsoft Graph connected to tenant '$($context.TenantId)' instead of configured tenant '$TenantId'."
    }
    Write-GraphWorkerLog -Level SUCCESS -Message "Authentication completed. DurationMs=$($authenticationTimer.ElapsedMilliseconds); ConnectedTenant=$($context.TenantId); Account=$($context.Account); AuthType=$($context.AuthType)."

    $evidence = [System.Collections.Generic.List[object]]::new()
    $uniqueEmails = @($emailAddresses | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Sort-Object -Unique)
    $current = 0
    foreach ($emailAddress in $uniqueEmails) {
        $current++
        Write-WorkerTextFile -Path $ProgressPath -Value ("{0}|{1}|{2}" -f $current, $uniqueEmails.Count, $emailAddress)
        $users = @()
        $licenseDetails = @()
        $queryError = ''
        $mailboxTimer=[Diagnostics.Stopwatch]::StartNew()
        Write-GraphWorkerLog -Level DEBUG -Message "Mailbox query starting. Index=$current/$($uniqueEmails.Count); EmailAddress=$emailAddress."
        try {
            $escaped = $emailAddress.Replace("'", "''")
            $graphUsers = @(Get-MgUser -Filter "userPrincipalName eq '$escaped' or mail eq '$escaped'" -Property @(
                'id', 'displayName', 'userPrincipalName', 'mail', 'proxyAddresses', 'accountEnabled',
                'onPremisesSyncEnabled', 'onPremisesImmutableId', 'onPremisesLastSyncDateTime',
                'onPremisesProvisioningErrors', 'usageLocation', 'assignedLicenses', 'assignedPlans'
            ) -All -ErrorAction Stop)
            $users = @($graphUsers | ForEach-Object {
                [pscustomobject][ordered]@{
                    Id = [string]$_.Id
                    DisplayName = [string]$_.DisplayName
                    UserPrincipalName = [string]$_.UserPrincipalName
                    Mail = [string]$_.Mail
                    ProxyAddresses = @($_.ProxyAddresses | ForEach-Object { [string]$_ })
                    AccountEnabled = [bool]$_.AccountEnabled
                    OnPremisesSyncEnabled = [bool]$_.OnPremisesSyncEnabled
                    OnPremisesImmutableId = [string]$_.OnPremisesImmutableId
                    OnPremisesLastSyncDateTime = [string]$_.OnPremisesLastSyncDateTime
                    OnPremisesProvisioningErrors = @($_.OnPremisesProvisioningErrors)
                    UsageLocation = [string]$_.UsageLocation
                    AssignedLicenses = @($_.AssignedLicenses | ForEach-Object { [string]$_.SkuId })
                    AssignedPlans = @($_.AssignedPlans | ForEach-Object { [string]$_.ServicePlanId })
                }
            })
            if ($graphUsers.Count -eq 1) {
                $licenseDetails = @(Get-MgUserLicenseDetail -UserId $graphUsers[0].Id -ErrorAction Stop | ForEach-Object {
                    [pscustomobject][ordered]@{
                        SkuId = [string]$_.SkuId
                        SkuPartNumber = [string]$_.SkuPartNumber
                        ServicePlans = @($_.ServicePlans | ForEach-Object {
                            [pscustomobject][ordered]@{
                                ServicePlanName = [string]$_.ServicePlanName
                                ProvisioningStatus = [string]$_.ProvisioningStatus
                                AppliesTo = [string]$_.AppliesTo
                            }
                        })
                    }
                })
            }
        }
        catch {
            $queryError = $_.Exception.Message
            Write-GraphWorkerLog -Level ERROR -Message "Mailbox query failed. Index=$current/$($uniqueEmails.Count); EmailAddress=$emailAddress; $(Format-GraphWorkerError $_)"
        }
        finally {
            $mailboxTimer.Stop()
            Write-GraphWorkerLog -Level $(if($queryError){'WARN'}else{'SUCCESS'}) -Message "Mailbox query ended. Index=$current/$($uniqueEmails.Count); EmailAddress=$emailAddress; DurationMs=$($mailboxTimer.ElapsedMilliseconds); UserCount=$($users.Count); LicenseDetailCount=$($licenseDetails.Count); Error=$queryError."
        }

        [void]$evidence.Add([pscustomobject][ordered]@{
            EmailAddress = $emailAddress
            SourceTimestamp = Get-Date
            Users = $users
            LicenseDetails = $licenseDetails
            QueryError = $queryError
        })
    }

    $subscribedSkus = @()
    Write-WorkerTextFile -Path $ProgressPath -Value '0|0|Tenant synchronization health'
    $organization = $null
    $organizationError = ''
    try {
        $organizationTimer=[Diagnostics.Stopwatch]::StartNew()
        Write-GraphWorkerLog -Level INFO -Message 'Tenant organization query starting.'
        $organizationRow = @(Get-MgOrganization -Property @(
            'id', 'displayName', 'onPremisesSyncEnabled', 'onPremisesLastSyncDateTime', 'verifiedDomains'
        ) -ErrorAction Stop | Select-Object -First 1)
        if ($organizationRow.Count -eq 1) {
            $organization = [pscustomobject][ordered]@{
                Id = [string]$organizationRow[0].Id
                DisplayName = [string]$organizationRow[0].DisplayName
                OnPremisesSyncEnabled = $organizationRow[0].OnPremisesSyncEnabled
                OnPremisesLastSyncDateTime = [string]$organizationRow[0].OnPremisesLastSyncDateTime
                VerifiedDomains = @($organizationRow[0].VerifiedDomains | ForEach-Object {
                    [pscustomobject][ordered]@{
                        Name = [string]$_.Name
                        IsDefault = [bool]$_.IsDefault
                        IsInitial = [bool]$_.IsInitial
                        Type = [string]$_.Type
                        Capabilities = [string]$_.Capabilities
                    }
                })
                CollectedAt = Get-Date
            }
        }
        else { $organizationError = 'Microsoft Graph returned no organization object.' }
    }
    catch { $organizationError = $_.Exception.Message; Write-GraphWorkerLog -Level ERROR -Message "Tenant organization query failed. $(Format-GraphWorkerError $_)" }
    finally { if($organizationTimer){$organizationTimer.Stop();Write-GraphWorkerLog -Level $(if($organizationError){'WARN'}else{'SUCCESS'}) -Message "Tenant organization query ended. DurationMs=$($organizationTimer.ElapsedMilliseconds); Available=$($null -ne $organization); Error=$organizationError."} }

    $subscribedSkuError = ''
    try {
        $skuTimer=[Diagnostics.Stopwatch]::StartNew()
        Write-GraphWorkerLog -Level INFO -Message 'Subscribed SKU query starting.'
        $subscribedSkus = @(Get-MgSubscribedSku -All -ErrorAction Stop | ForEach-Object {
            [pscustomobject][ordered]@{
                SkuId = [string]$_.SkuId
                SkuPartNumber = [string]$_.SkuPartNumber
                Enabled = [int]$_.PrepaidUnits.Enabled
                Consumed = [int]$_.ConsumedUnits
                ServicePlans = @($_.ServicePlans | ForEach-Object {
                    [pscustomobject][ordered]@{
                        ServicePlanId = [string]$_.ServicePlanId
                        ServicePlanName = [string]$_.ServicePlanName
                        ProvisioningStatus = [string]$_.ProvisioningStatus
                        AppliesTo = [string]$_.AppliesTo
                    }
                })
            }
        })
    }
    catch {
        $subscribedSkuError = $_.Exception.Message
        Write-GraphWorkerLog -Level ERROR -Message "Subscribed SKU query failed. $(Format-GraphWorkerError $_)"
    }
    finally { if($skuTimer){$skuTimer.Stop();Write-GraphWorkerLog -Level $(if($subscribedSkuError){'WARN'}else{'SUCCESS'}) -Message "Subscribed SKU query ended. DurationMs=$($skuTimer.ElapsedMilliseconds); SkuCount=$($subscribedSkus.Count); Error=$subscribedSkuError."} }

    [pscustomobject][ordered]@{
        TenantId = [string]$context.TenantId
        Account = [string]$context.Account
        AuthType = [string]$context.AuthType
        ContextScope = [string]$context.ContextScope
        Scopes = @($context.Scopes)
        CollectedAt = Get-Date
        Evidence = @($evidence)
        SubscribedSkus = $subscribedSkus
        Organization = $organization
        OrganizationError = $organizationError
        SubscribedSkuError = $subscribedSkuError
    } | Export-Clixml -LiteralPath $OutputPath -Depth 8 -Force
    Write-GraphWorkerLog -Level SUCCESS -Message "Evidence exported. MailboxEvidenceCount=$($evidence.Count); SkuCount=$($subscribedSkus.Count); OrganizationAvailable=$($null -ne $organization); OutputPath=$OutputPath."
}
catch {
    $message = $_.Exception.Message
    if ($_.Exception.InnerException) {
        $message += " | InnerException: $($_.Exception.InnerException.Message)"
    }
    $diagnostic = Format-GraphWorkerError $_
    Write-GraphWorkerLog -Level ERROR -Message "Fatal worker error. $diagnostic"
    Write-WorkerTextFile -Path $ErrorPath -Value "$message`r`n$diagnostic"
    exit 1
}
finally {
    if ($ContextScope -eq 'Process' -and (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    Write-GraphWorkerLog -Level INFO -Message "Worker exiting. TotalDurationMs=$([math]::Round(((Get-Date)-$workerStartedAt).TotalMilliseconds))."
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUJViQ0RDhVe0I
# rJMNO60JcUsqYIB6YL8fbwaBqQCspKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICYmykzyuODfjHGaXXbc/+En8HhS6oemChx4K3G5FVwcMA0GCSqG
# SIb3DQEBAQUABIIBgGA3AIr4EhakJAd7fi5I9qM7WWwiPULZn8QUcLMjiSFxtfKr
# OKC83R5ASeQ7M56gXnyNYspEJcXtsLzF9hm5ZAdkpK6KTdI/KoXSOrs/BGSdWZ5T
# 1n3hW3xcgdNLPH+3kPuvljY1jNXxZtV40BAb7GCqc3MWVJaE9wtGm+vWf41BkZrw
# P3Pb3MDR9WZ0mK5HNkah4Oaqn8bjoZ8M/mBVYpgkFSi8gFEWYiJ8ZFODIdfkFMVU
# ERKTfKA6q0LbX6DO/mGueS5UT5fmwO0gkU++WEkW8vmmejjSZ8GVZvjt+jSlEkmm
# EgRimxsg0lgBHSqeygQH8sH26nxtvS4Igy+O2vRtEwHl1HWsDY3jEw2xTh2xA2vd
# gt5i3G3tYw0VQW86yo5ht6peZVjW+mblonVYNTlIAJ+ZXFAF0fv2nJoNobr7owrx
# /aGiizHvuNgirWqtpeat/+3anhHMRD98fxzhp0YQfmPOMmU+P/hWW/CdRI6rgdy7
# 0/VTedgkYnFJ+oQutqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxNzA5
# MDZaMC8GCSqGSIb3DQEJBDEiBCAEvR4gwpLYIiweLOyFP5YUMxIKw/ztqVcPzzin
# utOJrDANBgkqhkiG9w0BAQEFAASCAgC3S1AQl7nbTh6pmJI0JPfyxcr7D3g/MMCh
# X5dpdg3SHyBpzhM3rAyfXrzwkzYC1yDKEqAbugqmYhackJyq3nNrzZlQ7xhkwJHD
# 3pMeJH6tZJwI6yn6vPR7pPVJVb/hizbb7qjldzCF1BZLk0zQRpCRuVAWpyWG5Qpt
# SkpSad4+CUkaVSEcbhyAyYmgCGMyKSDyYfUwf6q8wUe0bV1+dOi61V+qtEDx02+o
# Q+U8O8RAHzJaDzC5mK712EJcHl/0DXXlbYYBLFTbP8qea9LlBwHyINrKQtEPbixu
# OBnIerxnegnSIsG/aJ5VHOjo+pO1iDDFnzFm/lHPn/yAw+gaHSbizkizrFjaVJom
# bX3hm/fDB8f1fjVG18pLgdz3rbd01RO4bS6r/kj9JAQEfXEMc20v/IxhXwdMrp1o
# MOnV53usFwasLIIVRt0BDvDssM7yshV+A+v70uI1ttx97P7RElQsgXwvhNfgKIPT
# 4Caa8E3Dywk+LkE1hUxvqRw9I1GyWQkD+WlyG5i3U8/JZIxf/RVaGpgiSHnF+A4L
# /Y/HO8OWw2pIs9ggZGUNjCx6eB3yEjx8nOA4Xh18R5d4mxqPaeZvhQWvEIDCWR8p
# k0sVJEvQ2x3II5S/5bOXC/KzoXrWsDrfi/n6Xv9AQZp0I9DDzAOgvWBQQzkY4Zvw
# 9nddmoBmdw==
# SIG # End signature block
