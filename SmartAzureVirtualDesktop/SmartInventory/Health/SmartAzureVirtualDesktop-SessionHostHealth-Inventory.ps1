<#
.SYNOPSIS
Exports Azure Virtual Desktop session host health and capacity signals.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible Azure Virtual Desktop host pools and session
hosts, and exports CSV files for session host health, user sessions, and host pool capacity
summaries.

The script is read-only. It does not modify Azure resources or user sessions.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
Requires: Az.Accounts
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$TenantId,
    [string[]]$SubscriptionId,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [switch]$Connect,
    [switch]$UseDeviceCode,
    [int]$StaleHeartbeatHours = 24,
    [string]$AvdApiVersion = '2024-04-03',
    [string]$ResourceApiVersion = '2021-04-01'
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzureVirtualDesktop-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzureVirtualDesktop-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAvdEffectiveConfig = Initialize-SmartAvdTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAvdScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '' }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') { $TenantId = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue '' }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Health'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-SessionHostHealth-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-HealthIssue {
    [CmdletBinding()]
    param(
        [string]$Status,
        [AllowNull()]$LastHeartBeat,
        [AllowNull()]$AllowNewSession,
        [int]$StaleHours
    )

    $issues = @()
    if ($Status -and $Status -notin @('Available', 'NeedsAssistance')) { $issues += "Status=$Status" }
    if ($Status -eq 'NeedsAssistance') { $issues += 'NeedsAssistance' }
    if ($AllowNewSession -eq $false) { $issues += 'DrainMode' }
    if ($LastHeartBeat) {
        try {
            $ageHours = ((Get-Date).ToUniversalTime() - ([datetime]$LastHeartBeat).ToUniversalTime()).TotalHours
            if ($ageHours -gt $StaleHours) { $issues += ("StaleHeartbeat>{0}h" -f $StaleHours) }
        }
        catch {}
    }
    return ($issues -join ';')
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop session host health inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $sessionHostRows = @()
    $userSessionRows = @()
    $capacityRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $hostPools = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType 'Microsoft.DesktopVirtualization/hostPools' -ApiVersion $ResourceApiVersion)

        foreach ($hostPool in $hostPools) {
            $hostPoolDetails = Get-SmartAvdResourceById -ResourceId $hostPool.id -ApiVersion $AvdApiVersion -Operation 'Host pool details'
            if ($null -eq $hostPoolDetails) { $hostPoolDetails = $hostPool }
            $hostPoolProperties = Get-SmartAvdObjectPropertyValue -InputObject $hostPoolDetails -PropertyName @('properties')
            $hostPoolName = Get-SmartAvdResourceNameFromId -ResourceId $hostPool.id
            $resourceGroupName = Get-SmartAvdResourceGroupNameFromId -ResourceId $hostPool.id
            $maxSessionLimit = [int](Get-SmartAvdObjectPropertyValue -InputObject $hostPoolProperties -PropertyName @('maxSessionLimit'))
            $hostPoolType = Get-SmartAvdObjectPropertyValue -InputObject $hostPoolProperties -PropertyName @('hostPoolType')

            $hosts = @(Get-SmartAvdChildResources -ParentResourceId $hostPool.id -ChildPath 'sessionHosts' -ApiVersion $AvdApiVersion -Operation "Session hosts for $hostPoolName")
            $poolSessionCount = 0
            $availableHostCount = 0
            $drainHostCount = 0
            $unhealthyHostCount = 0
            $staleHeartbeatCount = 0

            foreach ($sessionHost in $hosts) {
                $properties = Get-SmartAvdObjectPropertyValue -InputObject $sessionHost -PropertyName @('properties')
                $sessionHostName = Get-SmartAvdResourceNameFromId -ResourceId $sessionHost.id
                $status = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('status')
                $allowNewSession = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('allowNewSession')
                $lastHeartBeat = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('lastHeartBeat')
                $sessions = [int](Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('sessions'))
                $healthIssue = Get-HealthIssue -Status $status -LastHeartBeat $lastHeartBeat -AllowNewSession $allowNewSession -StaleHours $StaleHeartbeatHours
                $heartbeatAgeHours = $null
                if ($lastHeartBeat) {
                    try { $heartbeatAgeHours = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]$lastHeartBeat).ToUniversalTime()).TotalHours, 2) } catch {}
                }

                $poolSessionCount += $sessions
                if ($status -eq 'Available') { $availableHostCount++ }
                if ($allowNewSession -eq $false) { $drainHostCount++ }
                if (-not [string]::IsNullOrWhiteSpace($healthIssue)) { $unhealthyHostCount++ }
                if ($heartbeatAgeHours -ne $null -and $heartbeatAgeHours -gt $StaleHeartbeatHours) { $staleHeartbeatCount++ }

                $sessionHostRows += [pscustomobject]@{
                    RunId             = $runId
                    SubscriptionId    = $subscription.Id
                    SubscriptionName  = $subscription.Name
                    ResourceGroupName = $resourceGroupName
                    HostPoolName      = $hostPoolName
                    HostPoolType      = $hostPoolType
                    SessionHostName   = $sessionHostName
                    Status            = $status
                    HealthIssue       = $healthIssue
                    UpdateState       = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('updateState')
                    AllowNewSession   = ConvertTo-SmartAvdBoolString -Value $allowNewSession
                    Sessions          = $sessions
                    AssignedUser      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('assignedUser')
                    LastHeartBeat     = $lastHeartBeat
                    HeartbeatAgeHours = $heartbeatAgeHours
                    AgentVersion      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('agentVersion')
                    OSVersion         = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('osVersion')
                    StatusTimestamp   = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('statusTimestamp')
                    VmResourceId      = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('resourceId')
                    ResourceId        = $sessionHost.id
                }

                $encodedSessionHostName = [uri]::EscapeDataString($sessionHostName)
                $userSessions = @(Get-SmartAvdChildResources -ParentResourceId $hostPool.id -ChildPath ("sessionHosts/{0}/userSessions" -f $encodedSessionHostName) -ApiVersion $AvdApiVersion -Operation "User sessions for $sessionHostName")
                foreach ($userSession in $userSessions) {
                    $userSessionProperties = Get-SmartAvdObjectPropertyValue -InputObject $userSession -PropertyName @('properties')
                    $userSessionRows += [pscustomobject]@{
                        RunId             = $runId
                        SubscriptionId    = $subscription.Id
                        SubscriptionName  = $subscription.Name
                        ResourceGroupName = $resourceGroupName
                        HostPoolName      = $hostPoolName
                        SessionHostName   = $sessionHostName
                        UserSessionId     = Get-SmartAvdResourceNameFromId -ResourceId $userSession.id
                        UserPrincipalName = Get-SmartAvdObjectPropertyValue -InputObject $userSessionProperties -PropertyName @('userPrincipalName')
                        ApplicationType   = Get-SmartAvdObjectPropertyValue -InputObject $userSessionProperties -PropertyName @('applicationType')
                        SessionState      = Get-SmartAvdObjectPropertyValue -InputObject $userSessionProperties -PropertyName @('sessionState')
                        CreateTime        = Get-SmartAvdObjectPropertyValue -InputObject $userSessionProperties -PropertyName @('createTime')
                        ResourceId        = $userSession.id
                    }
                }
            }

            $capacity = if ($maxSessionLimit -gt 0) { $hosts.Count * $maxSessionLimit } else { 0 }
            $capacityRows += [pscustomobject]@{
                RunId                 = $runId
                SubscriptionId        = $subscription.Id
                SubscriptionName      = $subscription.Name
                ResourceGroupName     = $resourceGroupName
                HostPoolName          = $hostPoolName
                HostPoolType          = $hostPoolType
                MaxSessionLimit       = $maxSessionLimit
                SessionHostCount      = $hosts.Count
                AvailableHostCount    = $availableHostCount
                DrainHostCount        = $drainHostCount
                UnhealthyHostCount    = $unhealthyHostCount
                StaleHeartbeatCount   = $staleHeartbeatCount
                CurrentSessionCount   = $poolSessionCount
                TheoreticalCapacity   = $capacity
                CapacityUsedPercent   = if ($capacity -gt 0) { [math]::Round(($poolSessionCount / $capacity) * 100, 2) } else { $null }
                LoadBalancerType      = Get-SmartAvdObjectPropertyValue -InputObject $hostPoolProperties -PropertyName @('loadBalancerType')
                StartVMOnConnect      = ConvertTo-SmartAvdBoolString -Value (Get-SmartAvdObjectPropertyValue -InputObject $hostPoolProperties -PropertyName @('startVMOnConnect'))
                ResourceId            = $hostPool.id
            }
        }
    }

    Export-SmartAvdCsv -Name 'AVD_SessionHostHealth' -Rows $sessionHostRows
    Export-SmartAvdCsv -Name 'AVD_UserSessions' -Rows $userSessionRows
    Export-SmartAvdCsv -Name 'AVD_HostPoolCapacitySummary' -Rows $capacityRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed session host health inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJu7WXVhFKQwav
# +/di0t2i5zHtm+LdJsh/Sd52EsF5v6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAQc/9ERiGv0YM/hiOO
# tWgnO/d9kH5Fsc2nS0ZZuiDktDANBgkqhkiG9w0BAQEFAASCAYCLV+hcWn8/pH5H
# lCY6MVMAiiyOkfHN0U8wReX6a/vlVqhU9aAfKcjU2p98YKVtV3vrNHrDzBAFU4BP
# +IxR/DMo4aUfGLjKHQQYLmBekOWVECh10ZKm4lw0sAtZg9jGyfDlZlxc8KowVnNr
# Oqa3LD+dtfy9kudD7L1LvSEV2G9zQOf5/ahhO7jHjAzMjuuZP7AnsXxew6xSCiWm
# lgB0HShmXAcFoKQjWej5xwJO+oPrffR0C6Ubb7OQ7cFwJa6RKs9ZgRB8xGlT9ukC
# gRgaWa9tmX/sX5mNPXuU7ZD6UWojEZcTxEDfExQ3d4okJHlMNXB/IZSc0D6jH7Cd
# G8tVBx/5Wc54AxrH7Bb9xtq0sbpgekJBvAd5W/Na0N700T4Iu/6HSW8cDMV7pCpR
# M5OyhsJIYihac5BHE1Mv/usvt3UU2Ufnxtnh1oqt90QiP/60fpA7yMo7ssRds8Fh
# ZGDJzvRP5DK5JacyjwTxkRa04HG12qpfaZXvSg/hi184UQR2LkQ=
# SIG # End signature block
