<#
.SYNOPSIS
Exports Azure Monitor diagnostic settings for Azure Virtual Desktop resources.

.DESCRIPTION
Connects with Az PowerShell, enumerates Azure Virtual Desktop host pools, workspaces,
application groups, and scaling plans, then exports diagnostic settings and resources
without diagnostic settings.

The script is read-only. It does not create or update diagnostic settings.

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
    [string]$DiagnosticSettingsApiVersion = '2021-05-01-preview',
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
$resolvedOutputRoots = Resolve-SmartAvdOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'AzureVirtualDesktop\Diagnostics'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzureVirtualDesktop-Diagnostics-Inventory_{0}.log" -f $runId)

Import-SmartAvdCoreModule -StartPath $PSScriptRoot
Set-SmartAvdCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAvdScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

function Get-AvdResourceTypeLabel {
    [CmdletBinding()]
    param([string]$ResourceType)

    switch -Regex ($ResourceType) {
        'hostPools$' { return 'HostPool' }
        'workspaces$' { return 'Workspace' }
        'applicationGroups$' { return 'ApplicationGroup' }
        'scalingPlans$' { return 'ScalingPlan' }
        default { return $ResourceType }
    }
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAvdLog -Message "Starting SmartAzureVirtualDesktop diagnostics inventory. RunId=$runId"
    Import-SmartAvdRequiredModule -Name Az.Accounts
    $context = Connect-SmartAvdCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw 'No Azure subscriptions were found for the current identity and filters.' }

    $diagnosticRows = @()
    $missingRows = @()
    $summaryRows = @()
    $avdResourceTypes = @(
        'Microsoft.DesktopVirtualization/hostPools',
        'Microsoft.DesktopVirtualization/workspaces',
        'Microsoft.DesktopVirtualization/applicationGroups',
        'Microsoft.DesktopVirtualization/scalingPlans'
    )

    foreach ($subscription in $subscriptions) {
        Write-SmartAvdLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $subscriptionResourceCount = 0
        $subscriptionConfiguredCount = 0

        foreach ($resourceType in $avdResourceTypes) {
            $resources = @(Get-SmartAvdSubscriptionResources -SubscriptionId $subscription.Id -ResourceType $resourceType -ApiVersion $ResourceApiVersion)
            foreach ($resource in $resources) {
                $subscriptionResourceCount++
                $typeLabel = Get-AvdResourceTypeLabel -ResourceType $resourceType
                $resourceName = Get-SmartAvdResourceNameFromId -ResourceId $resource.id
                $uri = "https://management.azure.com{0}/providers/microsoft.insights/diagnosticSettings?api-version={1}" -f $resource.id, $DiagnosticSettingsApiVersion
                $settings = @(Invoke-SmartAvdArmGetPaged -Uri $uri -Operation "Diagnostic settings for $resourceName")

                if ($settings.Count -eq 0) {
                    $missingRows += [pscustomobject]@{
                        RunId             = $runId
                        SubscriptionId    = $subscription.Id
                        SubscriptionName  = $subscription.Name
                        ResourceGroupName = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                        ResourceType      = $typeLabel
                        ResourceName      = $resourceName
                        Location          = $resource.location
                        Finding           = 'NoDiagnosticSettings'
                        ResourceId        = $resource.id
                    }
                    continue
                }

                $subscriptionConfiguredCount++
                foreach ($setting in $settings) {
                    $properties = Get-SmartAvdObjectPropertyValue -InputObject $setting -PropertyName @('properties')
                    $logs = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('logs'))
                    $metrics = @(Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('metrics'))
                    $enabledLogs = @($logs | Where-Object { (Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('enabled')) -eq $true })
                    $enabledMetrics = @($metrics | Where-Object { (Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('enabled')) -eq $true })

                    $diagnosticRows += [pscustomobject]@{
                        RunId                      = $runId
                        SubscriptionId             = $subscription.Id
                        SubscriptionName           = $subscription.Name
                        ResourceGroupName          = Get-SmartAvdResourceGroupNameFromId -ResourceId $resource.id
                        ResourceType               = $typeLabel
                        ResourceName               = $resourceName
                        DiagnosticSettingName      = Get-SmartAvdResourceNameFromId -ResourceId $setting.id
                        EnabledLogCount            = $enabledLogs.Count
                        EnabledMetricCount         = $enabledMetrics.Count
                        EnabledLogCategories       = (@($enabledLogs | ForEach-Object { Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('category', 'categoryGroup') }) -join ';')
                        EnabledMetricCategories    = (@($enabledMetrics | ForEach-Object { Get-SmartAvdObjectPropertyValue -InputObject $_ -PropertyName @('category') }) -join ';')
                        WorkspaceId                = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('workspaceId')
                        StorageAccountId           = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('storageAccountId')
                        EventHubAuthorizationRuleId = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('eventHubAuthorizationRuleId')
                        EventHubName               = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('eventHubName')
                        MarketplacePartnerId       = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('marketplacePartnerId')
                        LogAnalyticsDestinationType = Get-SmartAvdObjectPropertyValue -InputObject $properties -PropertyName @('logAnalyticsDestinationType')
                        LogsJson                   = ConvertTo-SmartAvdCompactJson -Value $logs
                        MetricsJson                = ConvertTo-SmartAvdCompactJson -Value $metrics
                        ResourceId                 = $resource.id
                        DiagnosticSettingId        = $setting.id
                    }
                }
            }
        }

        $summaryRows += [pscustomobject]@{
            RunId                         = $runId
            SubscriptionId                = $subscription.Id
            SubscriptionName              = $subscription.Name
            TenantId                      = $subscription.TenantId
            AvdResourceCount              = $subscriptionResourceCount
            ResourceWithDiagnosticsCount  = $subscriptionConfiguredCount
            ResourceWithoutDiagnosticsCount = $subscriptionResourceCount - $subscriptionConfiguredCount
            DiagnosticsCoveragePercent    = if ($subscriptionResourceCount -gt 0) { [math]::Round(($subscriptionConfiguredCount / $subscriptionResourceCount) * 100, 2) } else { $null }
        }
    }

    Export-SmartAvdCsv -Name 'AVD_DiagnosticSettings' -Rows $diagnosticRows
    Export-SmartAvdCsv -Name 'AVD_DiagnosticsMissing' -Rows $missingRows
    Export-SmartAvdCsv -Name 'AVD_DiagnosticsSummary' -Rows $summaryRows

    Write-SmartAvdLog -Level SUCCESS -Message ("Completed diagnostics inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    if (Get-Command -Name Write-SmartAvdLog -ErrorAction SilentlyContinue) { Write-SmartAvdLog -Level ERROR -Message $_.Exception.Message }
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBRHoz2iGlE8Eqi
# RK+7dLDKD9KwBMRFmqlOJ5fhYlXzEaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCANnVGRuqlbcCcNXA6Y1Ge+5se8S8XFDB0wR7ynTiugsjANBgkqhkiG9w0B
# AQEFAASCAYBx7H1qoHCtkmGG+dCgkt/vlfqe+vQY8vz/Muc50//0qJ+tfBvs3ZtG
# E5Xcj5KLRaavzSjsfcUQsgaPMUy1JYVooe4kV9QrU9lxzNgkfMTscptPXWmyYD7D
# rq2TSA+6HQOqB2yFXBGLY9POOum4ovr3MDdPDFxyqvgvBj33BGmASemv8KhkcxsN
# vIDx8j6KhhkDso1pgMY+88T2H6toblGccGwJF52ykMAU1FlHl4fviv6UYBgRxWt0
# VJdW0srMF/7VwkKSgfIyuPUdr2wJkbcjFPybmYiwFICLruAquiuyLSBqEiljM6Uv
# FCCJFVgdf7JU9maq4V8FbhYudnlzwJHf/e1QvbjXviLHj3yBUynNdJbtDQmOCydv
# 6p7SwqKjw/l00TzPA00S7tzaHOKJmQJ84lg4QGybMXjfXPeTKEEmEmdmouCy47Fh
# XFEQIth9RQyKfvYg8eP8WuLpHHmR2SDsBn67Vg2Cy2pzDpCcltx7+U1HEMzZIj35
# w+BsLp08Rfw=
# SIG # End signature block
