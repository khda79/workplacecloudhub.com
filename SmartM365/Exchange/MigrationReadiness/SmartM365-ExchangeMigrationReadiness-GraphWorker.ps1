<#
.SYNOPSIS
Collects Microsoft Graph evidence for Smart Exchange Migration Readiness in an isolated process.

.VERSION
1.11.17
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
        [Parameter(Mandatory)][AllowEmptyString()][string]$Path,
        [AllowEmptyString()][string]$Value = ''
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    [IO.File]::WriteAllText($Path, $Value, [Text.UTF8Encoding]::new($false))
}

try {
    Write-GraphWorkerLog -Level INFO -Message "Worker starting. Version=1.11.17; PowerShell=$($PSVersionTable.PSVersion); Computer=$env:COMPUTERNAME; TenantId=$TenantId; RuntimeInput=$InputPath; ContextScope=$ContextScope; ForceAuthentication=$ForceAuthentication; InspectContext=$InspectContext."
    $moduleTimer=[Diagnostics.Stopwatch]::StartNew()
    $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement')
    if ($InspectContext) { $requiredModules = @('Microsoft.Graph.Authentication') }
    $commonVersions = @((Get-Module -ListAvailable -Name $requiredModules[0]).Version | Sort-Object -Descending -Unique)
    foreach ($moduleName in $requiredModules | Select-Object -Skip 1) {
        $installedVersions = @((Get-Module -ListAvailable -Name $moduleName).Version)
        $commonVersions = @($commonVersions | Where-Object { $_ -in $installedVersions })
    }
    if ($commonVersions.Count -eq 0) { throw 'Install Microsoft.Graph.Authentication, Microsoft.Graph.Users and Microsoft.Graph.Identity.DirectoryManagement at the same version in PowerShell 7.' }
    $graphModuleVersion = $commonVersions[0]
    Import-Module Microsoft.Graph.Authentication -RequiredVersion $graphModuleVersion -ErrorAction Stop

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

    Import-Module Microsoft.Graph.Users -RequiredVersion $graphModuleVersion -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion $graphModuleVersion -ErrorAction Stop
    $moduleTimer.Stop()
    Write-GraphWorkerLog -Level SUCCESS -Message "Graph modules imported. DurationMs=$($moduleTimer.ElapsedMilliseconds); AuthenticationVersion=$((Get-Module Microsoft.Graph.Authentication).Version); UsersVersion=$((Get-Module Microsoft.Graph.Users).Version); DirectoryManagementVersion=$((Get-Module Microsoft.Graph.Identity.DirectoryManagement).Version)."

    if ($ValidateOnly) {
        foreach ($commandName in @('Connect-MgGraph', 'Get-MgUser', 'Get-MgUserLicenseDetail', 'Get-MgSubscribedSku', 'Get-MgOrganization')) {
            if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) { throw "Required Microsoft Graph command is unavailable: $commandName" }
        }
        'VALIDATION_OK SmartM365 Exchange Migration Readiness Graph worker v1.11.17'
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
    if ($ValidateOnly) { [Console]::Error.WriteLine($diagnostic) }
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDLl6036AbtU8AB
# ozAW7ytS323JrtVh68eMxKGaxbZ+mqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICoZ5jSyLmp0ycmSoVSXVxigCu8Kc1jF4wNvDn+R3my6MA0GCSqG
# SIb3DQEBAQUABIIBgGbSS5h9LyW0jV0y6uvhptid3z1OOPncpa792jlUOFopmvV2
# ryaQfQeUrTOdNiezLS1gaagZ1CdPOSskQiSU6bB+JlVU5OwikNwb1nJSAW7aPjXq
# Hxo6M2JKHZgsCpL+ODxxANaKz+gp+YrQGwENIe8JXQoTPsKonyHBtDHNJjUn89Xe
# 6DG9w/3jDnEFB1kZEvbCZMyanf1v7ZC5eIWWQsmZTZdwBudambOp2LW4XfUtGgOv
# DL/K9qpZqUzBc/e4P/hyicdxVVHojKyYcFcU2wI3jPGDJyc2/BL8A0QhSWuPT7B3
# OwFkgDNPdDTtSxkZqih9/O2Mz/MH5/ttvps0Oc44EXNJlG6/DiyWsZDcYWpZOPjj
# Cpoy4IYGhXtfIAdGFbzY8pjtXRig1HP0knUyRH1cVzLo8v/x1eh7INNrw/Qyob5Y
# FLiIiOVgHmHIDeTgDgrmoQPfQRS1wmi1Kue5TORczQQiDeSHEeTbmoPhTY7uR10G
# xtsq/bqQObRqP9hw7aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDgyMTQ5
# MjhaMC8GCSqGSIb3DQEJBDEiBCCxzGGoK7FoCqefvAFem4Fqx7hf9PaT9T/qBHFQ
# ADd9gzANBgkqhkiG9w0BAQEFAASCAgAmSMDdCndkWPavTjV59xLtJjPLvEMBGVZ+
# r3jNJGsQSHyoue8A02wSd7dasxKVdBHP9oWPG3JjfWfC672KXehAx3sNQtjKsBC2
# z5mRNPBJ6e5IHn18o7eeDWur1nThvYe2qc/BVt5mw9CN/DUPIyBU8U64vgG3mCfj
# yd4+gSOuJ04oQJQsoOB9mABk4tUAHEO6qiedUT1u1EyP+n5VCdi/s6r9mvAVzs6Y
# 0ra9bBBIxdSJCka156wCYAhAGFLeYLNLUDJqFANPLK/yGldo2ZP2O/VK4JHUiezE
# LXaVC3xpyAVxeCi8bD4i4OXqETZ+dd7yrCG4AJ6VA/Whcq98hVG1tN3Z1R0UpMDt
# s9eQRMekT+ckfsK3TWUIteXqSOvpN37fMQAQdCkAfnpMw0GvuNLWZ+paLJVS2DPt
# WdLPW9av1O2FEhWvT3z+MTE5/2uDNZRJZCRCGBHmER8Vao8vw2V7N3Zc0uR5bXI/
# mwr/Nspq6p/qT4edjoxFKwTX7xFDP2H4PfQs7avaIe985BBeUYGJlsGRjqn8Cvbp
# Js1jo2DlUlnj+Vzgz13gfQXJwyRMR6s+w+I0jKCEEMvAZ73Yjso4Fl0GXNXCEPfB
# XOO8FtNevZHWE9btQq1AEdafOQqbh5sl2Ja/XYQBgkRdLZU0iUQyRBfYF9HuVxFG
# eT+HBjqsJA==
# SIG # End signature block
