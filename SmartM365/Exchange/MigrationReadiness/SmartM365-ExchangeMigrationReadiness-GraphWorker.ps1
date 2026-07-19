<#
.SYNOPSIS
Collects Microsoft Graph evidence for Smart Exchange Migration Readiness in an isolated process.

.VERSION
1.11.14
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
    Write-GraphWorkerLog -Level INFO -Message "Worker starting. Version=1.11.14; PowerShell=$($PSVersionTable.PSVersion); Computer=$env:COMPUTERNAME; TenantId=$TenantId; RuntimeInput=$InputPath; ContextScope=$ContextScope; ForceAuthentication=$ForceAuthentication; InspectContext=$InspectContext."
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
        'VALIDATION_OK SmartM365 Exchange Migration Readiness Graph worker v1.11.14'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOeq4dJk2h/NNw
# SJd8WHTaG7cMHMFbrJWMFZSH/Y3MqKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBpnEmVqzuxfZXeGLK0nPuv
# vhsi3QcH3RKmhgE4xgEtODANBgkqhkiG9w0BAQEFAASCAYB3/gutHaLuE5eIS7hR
# 6Eb1IC5IZRAbXXmBJRxvnKkvItI6xJttI/sWcIrIY5yWzSW73zcQIeXYB+hdl2lv
# yX7Jiw/1Vtr1Jv9xDvqz8TKkr6WEGkUG6jqJrWE1iZJa2rRCItzZ13xOiBPh2Yjj
# gV0rhe4md4EKKkkiMecGGM0RCZ8nkCJRBuPl8iR1VhY+m3T6swGR4FOs/Lt8oo0x
# 7wRmQpHZEUoASEyNhMxGxqn8zZpkKC9w7iIkBCUTl5dlVOaVR0Q/pmghbeMmTpRI
# gFsCZXT3IsuxWJfjMFkqcOhjeRqA5f3Ygu7ZHW9fIwzukSIVjjFgOUzQvhausumO
# CqyC8oWgQcXw7TZRV1NAlGkuqU1c2wV92xFF7WBdusk8d4bxVcmNFlbeFj4hw5KG
# cR+JtqD4sHRkJA1ZDnO8RKB6GiRnx7P1Dl93UwwBCgNkM2EpLHmpZ8EQFD4dQhw4
# 3vQWr2MfO0kNcjADxrCSw1tiq6mKJe4SuIpisIJ9ETud7IA=
# SIG # End signature block
