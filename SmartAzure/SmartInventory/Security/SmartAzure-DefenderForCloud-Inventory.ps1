<#
.SYNOPSIS
Exports Microsoft Defender for Cloud security posture inventory.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for
Defender plans, secure score, secure score controls, security recommendations, security
contacts, auto-provisioning settings, regulatory compliance data, and subscription summaries.

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

.PARAMETER IncludeHealthyAssessments
Exports healthy Defender assessments too. By default, healthy assessments are skipped to keep
the recommendations CSV focused.

.PARAMETER SkipAssessments
Skips Defender assessment export.

.PARAMETER SkipRegulatoryCompliance
Skips regulatory compliance standard/control/assessment export.

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
    [switch]$IncludeHealthyAssessments,
    [switch]$SkipAssessments,
    [switch]$SkipRegulatoryCompliance,
    [string]$PricingApiVersion = '2024-01-01',
    [string]$SecureScoreApiVersion = '2020-01-01',
    [string]$AssessmentApiVersion = '2020-01-01',
    [string]$AutoProvisioningApiVersion = '2017-08-01-preview',
    [string]$SecurityContactApiVersion = '2020-01-01-preview',
    [string]$RegulatoryComplianceApiVersion = '2019-01-01-preview'
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
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Security'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-DefenderForCloud-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Defender for Cloud inventory. RunId=$runId"
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

    $planRows = @()
    $secureScoreRows = @()
    $secureScoreControlRows = @()
    $assessmentRows = @()
    $autoProvisioningRows = @()
    $securityContactRows = @()
    $regulatoryStandardRows = @()
    $regulatoryControlRows = @()
    $regulatoryAssessmentRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null

        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $subscriptionScope = "/subscriptions/$encodedSubscriptionId"

        $plans = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/pricings?api-version={1}" -f $encodedSubscriptionId, $PricingApiVersion) -Operation "Defender plans for $($subscription.Name)"
        foreach ($plan in $plans) {
            $properties = Get-ObjectPropertyValue -InputObject $plan -PropertyName @('properties')
            $planRows += [pscustomobject]@{
                RunId                  = $runId
                SubscriptionId         = $subscription.Id
                SubscriptionName       = $subscription.Name
                PlanName               = $plan.name
                PricingTier            = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('pricingTier')
                SubPlan                = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('subPlan')
                FreeTrialRemainingTime = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('freeTrialRemainingTime')
                EnablementTime         = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('enablementTime')
                Deprecated             = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('deprecated')
                ExtensionsJson         = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('extensions'))
                PropertiesJson         = ConvertTo-CompactJson -Value $properties
                ResourceId             = $plan.id
            }
        }

        $secureScores = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/secureScores?api-version={1}" -f $encodedSubscriptionId, $SecureScoreApiVersion) -Operation "Secure scores for $($subscription.Name)"
        foreach ($secureScore in $secureScores) {
            $properties = Get-ObjectPropertyValue -InputObject $secureScore -PropertyName @('properties')
            $score = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('score')
            $secureScoreRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                Name             = $secureScore.name
                DisplayName      = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                CurrentScore     = Get-ObjectPropertyValue -InputObject $score -PropertyName @('current')
                MaxScore         = Get-ObjectPropertyValue -InputObject $score -PropertyName @('max')
                Percentage       = Get-ObjectPropertyValue -InputObject $score -PropertyName @('percentage')
                Weight           = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('weight')
                ResourceId       = $secureScore.id
            }
        }

        $secureScoreControls = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/secureScoreControls?api-version={1}" -f $encodedSubscriptionId, $SecureScoreApiVersion) -Operation "Secure score controls for $($subscription.Name)"
        foreach ($control in $secureScoreControls) {
            $properties = Get-ObjectPropertyValue -InputObject $control -PropertyName @('properties')
            $score = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('score')
            $secureScoreControlRows += [pscustomobject]@{
                RunId                      = $runId
                SubscriptionId             = $subscription.Id
                SubscriptionName           = $subscription.Name
                ControlName                = $control.name
                DisplayName                = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                CurrentScore               = Get-ObjectPropertyValue -InputObject $score -PropertyName @('current')
                MaxScore                   = Get-ObjectPropertyValue -InputObject $score -PropertyName @('max')
                Percentage                 = Get-ObjectPropertyValue -InputObject $score -PropertyName @('percentage')
                HealthyResourceCount       = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('healthyResourceCount')
                UnhealthyResourceCount     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('unhealthyResourceCount')
                NotApplicableResourceCount = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('notApplicableResourceCount')
                Weight                     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('weight')
                AssessmentDefinitionsJson  = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('assessmentDefinitions'))
                ResourceId                 = $control.id
            }
        }

        $assessments = @()
        if (-not $SkipAssessments) {
            $assessments = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/assessments?api-version={1}" -f $encodedSubscriptionId, $AssessmentApiVersion) -Operation "Security assessments for $($subscription.Name)"
            foreach ($assessment in $assessments) {
                $properties = Get-ObjectPropertyValue -InputObject $assessment -PropertyName @('properties')
                $status = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('status')
                $metadata = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('metadata')
                $statusCode = Get-ObjectPropertyValue -InputObject $status -PropertyName @('code')
                if (-not $IncludeHealthyAssessments -and $statusCode -eq 'Healthy') { continue }

                $resourceDetails = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('resourceDetails')
                $assessmentRows += [pscustomobject]@{
                    RunId             = $runId
                    SubscriptionId    = $subscription.Id
                    SubscriptionName  = $subscription.Name
                    AssessmentName    = $assessment.name
                    DisplayName       = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                    StatusCode        = $statusCode
                    StatusCause       = Get-ObjectPropertyValue -InputObject $status -PropertyName @('cause')
                    StatusDescription = Get-ObjectPropertyValue -InputObject $status -PropertyName @('description')
                    Severity          = Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('severity')
                    CategoriesJson    = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('categories'))
                    Remediation       = Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('remediationDescription')
                    ThreatsJson       = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('threats'))
                    UserImpact        = Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('userImpact')
                    ImplementationEffort = Get-ObjectPropertyValue -InputObject $metadata -PropertyName @('implementationEffort')
                    ResourceSource    = Get-ObjectPropertyValue -InputObject $resourceDetails -PropertyName @('source')
                    AssessedResourceId = Get-ObjectPropertyValue -InputObject $resourceDetails -PropertyName @('id', 'resourceId')
                    AdditionalDataJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('additionalData'))
                    ResourceId        = $assessment.id
                }
            }
        }

        $autoProvisioningSettings = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/autoProvisioningSettings?api-version={1}" -f $encodedSubscriptionId, $AutoProvisioningApiVersion) -Operation "Auto-provisioning settings for $($subscription.Name)"
        foreach ($setting in $autoProvisioningSettings) {
            $properties = Get-ObjectPropertyValue -InputObject $setting -PropertyName @('properties')
            $autoProvisioningRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                Name             = $setting.name
                AutoProvision    = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('autoProvision')
                ResourceId       = $setting.id
                PropertiesJson   = ConvertTo-CompactJson -Value $properties
            }
        }

        $securityContacts = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Security/securityContacts?api-version={1}" -f $encodedSubscriptionId, $SecurityContactApiVersion) -Operation "Security contacts for $($subscription.Name)"
        foreach ($contact in $securityContacts) {
            $properties = Get-ObjectPropertyValue -InputObject $contact -PropertyName @('properties')
            $securityContactRows += [pscustomobject]@{
                RunId              = $runId
                SubscriptionId     = $subscription.Id
                SubscriptionName   = $subscription.Name
                Name               = $contact.name
                Emails             = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('emails')
                Phone              = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('phone')
                AlertNotifications = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('alertNotifications')
                AlertsToAdmins     = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('alertsToAdmins')
                NotificationsByRoleJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('notificationsByRole'))
                ResourceId         = $contact.id
            }
        }

        $regulatoryAssessments = @()
        if (-not $SkipRegulatoryCompliance) {
            $standards = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com{0}/providers/Microsoft.Security/regulatoryComplianceStandards?api-version={1}" -f $subscriptionScope, $RegulatoryComplianceApiVersion) -Operation "Regulatory compliance standards for $($subscription.Name)"
            foreach ($standard in $standards) {
                $properties = Get-ObjectPropertyValue -InputObject $standard -PropertyName @('properties')
                $regulatoryStandardRows += [pscustomobject]@{
                    RunId            = $runId
                    SubscriptionId   = $subscription.Id
                    SubscriptionName = $subscription.Name
                    StandardName     = $standard.name
                    DisplayName      = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('displayName')
                    State            = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('state')
                    PassedControls   = Get-SmartAzureNestedPropertyValue -InputObject $properties -Path @('passedControls')
                    FailedControls   = Get-SmartAzureNestedPropertyValue -InputObject $properties -Path @('failedControls')
                    SkippedControls  = Get-SmartAzureNestedPropertyValue -InputObject $properties -Path @('skippedControls')
                    UnsupportedControls = Get-SmartAzureNestedPropertyValue -InputObject $properties -Path @('unsupportedControls')
                    ResourceId       = $standard.id
                    PropertiesJson   = ConvertTo-CompactJson -Value $properties
                }

                $encodedStandardName = [uri]::EscapeDataString($standard.name)
                $controls = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com{0}/providers/Microsoft.Security/regulatoryComplianceStandards/{1}/regulatoryComplianceControls?api-version={2}" -f $subscriptionScope, $encodedStandardName, $RegulatoryComplianceApiVersion) -Operation "Regulatory controls $($standard.name) for $($subscription.Name)"
                foreach ($regulatoryControl in $controls) {
                    $controlProperties = Get-ObjectPropertyValue -InputObject $regulatoryControl -PropertyName @('properties')
                    $regulatoryControlRows += [pscustomobject]@{
                        RunId            = $runId
                        SubscriptionId   = $subscription.Id
                        SubscriptionName = $subscription.Name
                        StandardName     = $standard.name
                        ControlName      = $regulatoryControl.name
                        DisplayName      = Get-ObjectPropertyValue -InputObject $controlProperties -PropertyName @('displayName')
                        State            = Get-ObjectPropertyValue -InputObject $controlProperties -PropertyName @('state')
                        PassedAssessments = Get-ObjectPropertyValue -InputObject $controlProperties -PropertyName @('passedAssessments')
                        FailedAssessments = Get-ObjectPropertyValue -InputObject $controlProperties -PropertyName @('failedAssessments')
                        SkippedAssessments = Get-ObjectPropertyValue -InputObject $controlProperties -PropertyName @('skippedAssessments')
                        ResourceId       = $regulatoryControl.id
                        PropertiesJson   = ConvertTo-CompactJson -Value $controlProperties
                    }

                    $encodedControlName = [uri]::EscapeDataString($regulatoryControl.name)
                    $controlAssessments = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com{0}/providers/Microsoft.Security/regulatoryComplianceStandards/{1}/regulatoryComplianceControls/{2}/regulatoryComplianceAssessments?api-version={3}" -f $subscriptionScope, $encodedStandardName, $encodedControlName, $RegulatoryComplianceApiVersion) -Operation "Regulatory assessments $($standard.name)/$($regulatoryControl.name) for $($subscription.Name)"
                    foreach ($regulatoryAssessment in $controlAssessments) {
                        $assessmentProperties = Get-ObjectPropertyValue -InputObject $regulatoryAssessment -PropertyName @('properties')
                        $regulatoryAssessments += $regulatoryAssessment
                        $regulatoryAssessmentRows += [pscustomobject]@{
                            RunId            = $runId
                            SubscriptionId   = $subscription.Id
                            SubscriptionName = $subscription.Name
                            StandardName     = $standard.name
                            ControlName      = $regulatoryControl.name
                            AssessmentName   = $regulatoryAssessment.name
                            DisplayName      = Get-ObjectPropertyValue -InputObject $assessmentProperties -PropertyName @('displayName')
                            State            = Get-ObjectPropertyValue -InputObject $assessmentProperties -PropertyName @('state')
                            Description      = Get-ObjectPropertyValue -InputObject $assessmentProperties -PropertyName @('description')
                            ResourceId       = $regulatoryAssessment.id
                            PropertiesJson   = ConvertTo-CompactJson -Value $assessmentProperties
                        }
                    }
                }
            }
        }

        $unhealthyAssessments = @($assessments | Where-Object {
                (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('status')).code -eq 'Unhealthy'
            })
        $highSeverityAssessments = @($assessments | Where-Object {
                (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('metadata')).severity -eq 'High'
            })
        $standardTierPlans = @($plans | Where-Object {
                (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('pricingTier')) -eq 'Standard'
            })
        $latestSecureScore = @($secureScoreRows | Where-Object { $_.SubscriptionId -eq $subscription.Id } | Select-Object -First 1)

        $summaryRows += [pscustomobject]@{
            RunId                         = $runId
            SubscriptionId                = $subscription.Id
            SubscriptionName              = $subscription.Name
            TenantId                      = $subscription.TenantId
            DefenderPlanCount             = $plans.Count
            StandardTierPlanCount         = $standardTierPlans.Count
            SecureScorePercentage         = if ($latestSecureScore.Count -gt 0) { $latestSecureScore[0].Percentage } else { $null }
            SecureScoreControlCount       = $secureScoreControls.Count
            AssessmentCount               = $assessments.Count
            ExportedAssessmentCount       = @($assessmentRows | Where-Object { $_.SubscriptionId -eq $subscription.Id }).Count
            UnhealthyAssessmentCount      = $unhealthyAssessments.Count
            HighSeverityAssessmentCount   = $highSeverityAssessments.Count
            AutoProvisioningSettingCount  = $autoProvisioningSettings.Count
            SecurityContactCount          = $securityContacts.Count
            RegulatoryAssessmentCount     = $regulatoryAssessments.Count
            IncludeHealthyAssessments     = [bool]$IncludeHealthyAssessments
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Defender_Plans' -Rows $planRows
    Export-SmartAzureCsv -Name 'Azure_Defender_SecureScores' -Rows $secureScoreRows
    Export-SmartAzureCsv -Name 'Azure_Defender_SecureScoreControls' -Rows $secureScoreControlRows
    Export-SmartAzureCsv -Name 'Azure_Defender_Assessments' -Rows $assessmentRows
    Export-SmartAzureCsv -Name 'Azure_Defender_AutoProvisioning' -Rows $autoProvisioningRows
    Export-SmartAzureCsv -Name 'Azure_Defender_SecurityContacts' -Rows $securityContactRows
    Export-SmartAzureCsv -Name 'Azure_Defender_RegulatoryStandards' -Rows $regulatoryStandardRows
    Export-SmartAzureCsv -Name 'Azure_Defender_RegulatoryControls' -Rows $regulatoryControlRows
    Export-SmartAzureCsv -Name 'Azure_Defender_RegulatoryAssessments' -Rows $regulatoryAssessmentRows
    Export-SmartAzureCsv -Name 'Azure_Defender_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Defender for Cloud inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAoBIfmRAevV14g
# g6l6XLsFO3joIlLSiWh7VZzRH+4Z3KCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAcE3KhwSIaqAxMJTT8+dCR7Y8G5TVN3NPPFeHKlAgiGjANBgkqhkiG9w0B
# AQEFAASCAYAQskYVvoOUgv9w0IXXWSWZagmqsYuwrNGTuDok7C3ngOb6LHEmfIhy
# 43iUqc2qyofQiYhTKJgq6KRr8kUok3PDpP7cj8UMlnBXeDTkNAB4OU7okNvniZbQ
# sFAi0hreOK4Ij9L88l04MS7fokNGUg74AMDfRpmHEFt6C5/5akTy1Y3HIjmTZhT8
# G/hHMWQzBw3c+Qrg74k1iY61yr71efI659XQJl/rV29DgM5k3BxWH/YXtV2RRbV5
# RTKe5YVpyj9z3qZ+cbpwxG6jDizTsrw/XXfgrP3fR+pzR9wM6J6FH95xWUykyg2u
# pYQc0V7ylp09ZTiVASoxobcyyyr8GBK254CEiC8E41wSSz6g/A+BMb56Spe6dBpb
# ZIGu41djdJ6Z+0A+b212/6l9wNserpsd1O5PuhJrK7LCtn3enZnkmvnhCt1KNiJz
# xtyM6SbV1eDBO0F2aR2IO/NL8QY0nyagG9vZyhoN2zFPn/ItghoadlYksdNFjj+E
# 8RxUshf19VQ=
# SIG # End signature block
