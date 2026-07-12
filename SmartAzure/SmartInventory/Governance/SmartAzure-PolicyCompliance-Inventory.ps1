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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpN9KQ/jEK1Lv/
# djowaPE4otHJ5cCmmk1vewElCvWqEqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAkPeRlt8tuG2t8fLfDuUL8RADiGtoQVc03HyEPdbDN/TANBgkqhkiG9w0B
# AQEFAASCAYBOTggqfpInxDlw6LRecZ0tvw2CS4TxnndPeWKALJnt6gWSIzFltk6q
# ivd2nTwUK2mCiliXieatM24LfUOCsXpLRTfAYz+a9a++cxpHnMXf0BdiLIES3yjx
# YybYDbA1Ns5fPFubn93e/h7+2jRZSm6TbEH+vvKGeFI3Nnqw2GPxVddziwKrQvAV
# 9RddXFCKHSAyfZXv/c6zpAJe2IztTFDCdE/O8hx8gZw0o6MzpC4veNE/fr/9/PTV
# wkKi7v+F7sXBUHvdzjwv9PRgvJv1H8PZ+98U4eQestGAGnBzEjRW32pBfBDXx/Lr
# mK2FqmVk/V831Hf9H/Mygf2dllmnOXQMvt1y5Bwc5t4v+O7yEtRD4Uc3N07zte1p
# Q1SjFeKYmGgSqdulSQAhmHUXdLrV/cNZUGrP0bvu8WIJf/NFgWPJPVMmwFkIJQlC
# LZyrcYD6CQKTXSF7t7TMGKuDuEflH3vRZwBNXYxjPIJyZ5rEkuUWP6Mmr3xNZz31
# FEbXYSNYVxc=
# SIG # End signature block
