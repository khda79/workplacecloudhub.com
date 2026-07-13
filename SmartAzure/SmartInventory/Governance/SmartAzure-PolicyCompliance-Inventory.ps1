<#
.SYNOPSIS
Exports Azure Policy governance and compliance inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
policy assignments, policy definitions, initiative definitions, policy exemptions, policy
state details, and policy state summaries.

The script is read-only. It uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

.PARAMETER Tenant
Local tenant profile key used to isolate SmartAzure output folders. Defaults to test.

.PARAMETER TenantId
Optional tenant ID or domain used by Connect-AzAccount. If omitted, the SmartAzure tenant
profile AzureTenantId, TenantId, or OrgDomain is used when available.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. When omitted, all visible subscriptions are included.

.PARAMETER OutputRoot
Historical output root. A run-specific folder is created below this path.

.PARAMETER LatestOutputRoot
Latest CSV output root. Stable non-timestamped copies are written here.

.PARAMETER Connect
Forces an Azure sign-in before inventory.

.PARAMETER UseDeviceCode
Uses device code authentication for Connect-AzAccount.

.PARAMETER IncludeCompliantPolicyStates
Exports compliant policy state rows too. By default, policy state detail export is limited to
non-compliant states to keep CSV size practical.

.PARAMETER PolicyStateTop
Maximum number of policy state rows requested per page. Default is 5000.

.PARAMETER SkipPolicyStates
Skips PolicyInsights policy state detail export.

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
    [switch]$IncludeCompliantPolicyStates,
    [int]$PolicyStateTop = 5000,
    [switch]$SkipPolicyStates
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartAzure-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartAzure-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartAzureEffectiveConfig = Initialize-SmartAzureTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartAzureScriptLocalConfig -ScriptPath $PSCommandPath
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'AzureTenantId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue ''
    }
    if ([string]::IsNullOrWhiteSpace($TenantId) -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
        $TenantId = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue ''
    }
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue ''
}
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) {
    $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue ''
}

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Policy'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-PolicyCompliance-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-ResourceNameFromId {
    [CmdletBinding()]
    param([string]$ResourceId)

    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return '' }
    return ($ResourceId.TrimEnd('/') -split '/')[-1]
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Policy compliance inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }

    if ($subscriptions.Count -eq 0) {
        throw "No Azure subscriptions were found for the current identity and filters."
    }

    $assignmentRows = @()
    $policyDefinitionRows = @()
    $initiativeDefinitionRows = @()
    $exemptionRows = @()
    $policyStateRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $subscriptionScope = "/subscriptions/$encodedSubscriptionId"

        $assignments = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Authorization/policyAssignments?api-version=2022-06-01" -f $encodedSubscriptionId) -Operation "Policy assignments for $($subscription.Name)"
        foreach ($assignment in $assignments) {
            $properties = Get-ObjectPropertyValue -InputObject $assignment -PropertyName @('properties')
            $assignmentRows += [pscustomobject]@{
                RunId                 = $runId
                SubscriptionId        = $subscription.Id
                SubscriptionName      = $subscription.Name
                PolicyAssignmentName  = $assignment.name
                DisplayName           = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                Description           = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('description')
                Scope                 = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('scope')
                NotScopesJson         = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('notScopes'))
                EnforcementMode       = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('enforcementMode')
                PolicyDefinitionId    = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyDefinitionId')
                DefinitionName        = Get-ResourceNameFromId -ResourceId (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyDefinitionId'))
                ParametersJson        = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('parameters'))
                NonComplianceMessagesJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('nonComplianceMessages'))
                MetadataJson          = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata'))
                IdentityType          = if ($assignment.identity) { $assignment.identity.type } else { '' }
                Location              = $assignment.location
                ResourceId            = $assignment.id
            }
        }

        $policyDefinitions = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Authorization/policyDefinitions?api-version=2021-06-01" -f $encodedSubscriptionId) -Operation "Policy definitions for $($subscription.Name)"
        foreach ($definition in $policyDefinitions) {
            $properties = Get-ObjectPropertyValue -InputObject $definition -PropertyName @('properties')
            $policyDefinitionRows += [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                Name               = $definition.name
                DisplayName        = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                PolicyType         = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyType')
                Mode               = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('mode')
                Category           = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata')) -PropertyName @('category')
                Version            = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata')) -PropertyName @('version')
                ParametersJson     = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('parameters'))
                PolicyRuleJson     = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyRule'))
                MetadataJson       = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata'))
                ResourceId         = $definition.id
            }
        }

        $initiatives = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Authorization/policySetDefinitions?api-version=2021-06-01" -f $encodedSubscriptionId) -Operation "Policy set definitions for $($subscription.Name)"
        foreach ($initiative in $initiatives) {
            $properties = Get-ObjectPropertyValue -InputObject $initiative -PropertyName @('properties')
            $policyDefinitionsInSet = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyDefinitions')
            $initiativeDefinitionRows += [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                Name               = $initiative.name
                DisplayName        = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                PolicyType         = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyType')
                Category           = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata')) -PropertyName @('category')
                Version            = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata')) -PropertyName @('version')
                PolicyDefinitionCount = @($policyDefinitionsInSet).Count
                ParametersJson     = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('parameters'))
                PolicyDefinitionsJson = ConvertTo-CompactJson -Value $policyDefinitionsInSet
                MetadataJson       = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata'))
                ResourceId         = $initiative.id
            }
        }

        $exemptions = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Authorization/policyExemptions?api-version=2022-07-01-preview" -f $encodedSubscriptionId) -Operation "Policy exemptions for $($subscription.Name)"
        foreach ($exemption in $exemptions) {
            $properties = Get-ObjectPropertyValue -InputObject $exemption -PropertyName @('properties')
            $expiresOn = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('expiresOn')
            $exemptionRows += [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                Name               = $exemption.name
                DisplayName        = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                Description        = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('description')
                ExemptionCategory  = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('exemptionCategory')
                PolicyAssignmentId = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyAssignmentId')
                PolicyDefinitionReferenceIdsJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('policyDefinitionReferenceIds'))
                ExpiresOn          = $expiresOn
                IsExpired          = if ($expiresOn) { ([datetime]$expiresOn) -lt (Get-Date) } else { $false }
                MetadataJson       = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata'))
                ResourceId         = $exemption.id
            }
        }

        $subscriptionPolicyStateRows = @()
        if (-not $SkipPolicyStates) {
            $queryParameters = @("api-version=2019-10-01", "`$top=$PolicyStateTop")
            if (-not $IncludeCompliantPolicyStates) {
                $queryParameters += ("`$filter={0}" -f [uri]::EscapeDataString("ComplianceState ne 'Compliant'"))
            }
            $policyStatesUri = "https://management.azure.com{0}/providers/Microsoft.PolicyInsights/policyStates/latest/queryResults?{1}" -f $subscriptionScope, ($queryParameters -join '&')
            $subscriptionPolicyStateRows = @(Invoke-SmartAzureArmGetPaged -Uri $policyStatesUri -Operation "Policy states for $($subscription.Name)")
            foreach ($state in $subscriptionPolicyStateRows) {
                $policyStateRows += [pscustomobject]@{
                    RunId                    = $runId
                    SubscriptionId           = $subscription.Id
                    SubscriptionName         = $subscription.Name
                    Timestamp                = Get-ObjectPropertyValue -InputObject $state -PropertyName @('timestamp')
                    ComplianceState          = Get-ObjectPropertyValue -InputObject $state -PropertyName @('complianceState')
                    PolicyAssignmentId       = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyAssignmentId')
                    PolicyAssignmentName     = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyAssignmentName')
                    PolicyAssignmentScope    = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyAssignmentScope')
                    PolicyDefinitionId       = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyDefinitionId')
                    PolicyDefinitionName     = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyDefinitionName')
                    PolicyDefinitionAction   = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyDefinitionAction')
                    PolicySetDefinitionId    = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policySetDefinitionId')
                    PolicySetDefinitionName  = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policySetDefinitionName')
                    PolicyDefinitionReferenceId = Get-ObjectPropertyValue -InputObject $state -PropertyName @('policyDefinitionReferenceId')
                    ResourceId               = Get-ObjectPropertyValue -InputObject $state -PropertyName @('resourceId')
                    ResourceType             = Get-ObjectPropertyValue -InputObject $state -PropertyName @('resourceType')
                    ResourceLocation         = Get-ObjectPropertyValue -InputObject $state -PropertyName @('resourceLocation')
                    ResourceGroup            = Get-ObjectPropertyValue -InputObject $state -PropertyName @('resourceGroup')
                    IsCompliant              = (Get-ObjectPropertyValue -InputObject $state -PropertyName @('isCompliant'))
                    ComponentsJson           = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $state -PropertyName @('components'))
                }
            }
        }

        $nonCompliantStateCount = @($subscriptionPolicyStateRows | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('complianceState')) -eq 'NonCompliant' }).Count
        $compliantStateCount = @($subscriptionPolicyStateRows | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('complianceState')) -eq 'Compliant' }).Count
        $conflictStateCount = @($subscriptionPolicyStateRows | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('complianceState')) -eq 'Conflict' }).Count
        $exemptStateCount = @($subscriptionPolicyStateRows | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('complianceState')) -eq 'Exempt' }).Count

        $summaryRows += [pscustomobject]@{
            RunId                    = $runId
            SubscriptionId           = $subscription.Id
            SubscriptionName         = $subscription.Name
            TenantId                 = $subscription.TenantId
            AssignmentCount          = $assignments.Count
            DefinitionCount          = $policyDefinitions.Count
            InitiativeDefinitionCount = $initiatives.Count
            ExemptionCount           = $exemptions.Count
            PolicyStateRowsExported  = $subscriptionPolicyStateRows.Count
            NonCompliantStateCount   = $nonCompliantStateCount
            CompliantStateCount      = $compliantStateCount
            ConflictStateCount       = $conflictStateCount
            ExemptStateCount         = $exemptStateCount
            IncludeCompliantPolicyStates = [bool]$IncludeCompliantPolicyStates
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Policy_Assignments' -Rows $assignmentRows
    Export-SmartAzureCsv -Name 'Azure_Policy_Definitions' -Rows $policyDefinitionRows
    Export-SmartAzureCsv -Name 'Azure_Policy_InitiativeDefinitions' -Rows $initiativeDefinitionRows
    Export-SmartAzureCsv -Name 'Azure_Policy_Exemptions' -Rows $exemptionRows
    Export-SmartAzureCsv -Name 'Azure_Policy_States' -Rows $policyStateRows
    Export-SmartAzureCsv -Name 'Azure_Policy_ComplianceSummary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Policy compliance inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpN9KQ/jEK1Lv/
# djowaPE4otHJ5cCmmk1vewElCvWqEqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICQ95GW3y24ba3x8t8O5QvxEAOIa2hBVzTcfIQ91sM39MA0GCSqG
# SIb3DQEBAQUABIIBgBNHC7RrIS9sFtAH6sPLMHJEpO2SF7NuchC5AvKU9BOZ+8qw
# JYNHvx4H6wqiUN2AfPhwZk7Rnz7nYE42kRKDc5uJngD27Y6Y3S/jb1xXwniORigu
# ROxqAt0+UAos6mdOLfni+AcwTs0ZJmVQLIS6TxgJgyL+BUuwOHezUGqCM8UY9TbX
# rfM4wRJM2ANB3ALcuTJNI7t8HuvQ9t6i3v61qQt/jryvYYCyWTx+01gQ+Ua34iMB
# RYXB1nzFdeh6CIjTbMx8s6N0K2iQAizMFjwKwCUTvYgfA+2NnShlcK4LWHtBMN77
# JTS6ATvOPMif2m1mu3i6kK3GTF4wOeB9ROYhlDSqNmA3g3Eb6tDTb3i/aQb3SDw4
# o+mGtCwnSVuBdXvp6mCZi7efDc5/VlorEsQ684YgOCe1UNuhcpvHSAccoZjyKNSE
# aiGgO/YnZwX6F/4mT7XonDB1V9az2YvGxgsvhGhvcnAQjrnTngmBa3FfyadpEHzt
# 66PN3nooGXktzCfZ2KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MTFaMC8GCSqGSIb3DQEJBDEiBCDbqrmfixKiaUuVLqu5S/keZr6XE3PQPmY+LFbs
# E5rR+DANBgkqhkiG9w0BAQEFAASCAgCqj3wT0GtV7M7+Y7AO2ysgmsqqRG28MFd0
# V0BwTS8RM0AgFSRxao/AbY8+BDHSmHEX9+b4edlKwBL876148ShpOmc1JOtObo8H
# ZM4pLfTjPJ4IJlnAeSL4MpG40RsukZ/fCeDDWQDyFG/+om9fJd7vKkGh5ntMoaPW
# PCQmVEMFtLQug71kqfhbTl0M0kvsMNISHPCfmWmdvaGPig8G6Hexk4vQDqd9D43v
# gC5tVuIoeXHOmT4d+yLmLFM56EI36KVdy3ZNXiSHhzemtCv6KopYUOkOGeLyagca
# kqVup+dM3Vqb4rhXVBnKXdqGtWwkDMXVcENcdUENOV2dkd/kUMbBc+MzPSZafqJo
# pscm+VaxrI7glTJW43IMaK9d1NwLuQ8hWhV6Up520ncZ4yau7bx9Ep8lSRnsBOXy
# hBOFWtV0+KMlvujpKaYztDb0GprGhE1xWvEqbdA67cbqgUfcjXUkxf3rcOY1LJi7
# VrPJLl0a0zIm5T98wZNMGnVVWFFEpceMyzJX6QhT0kusaZYNvsSlq2CwE0XFtJEK
# CmNaCb3dXKlrT+O9ASwVCECUH/0SRaobq9DT7xx+2MEj1RJMHTX1zXgK6CgfhwID
# Wam5OaHzB1OC0GeySfBuynxe8LkZ7k7gj4VPRVcSrVdW9FLuT/i56DIFwlbkg/5z
# gqHU7T07Ig==
# SIG # End signature block
