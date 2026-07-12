<#
.SYNOPSIS
Exports Azure Advisor recommendations across all categories.

.DESCRIPTION
Connects with Az PowerShell, enumerates visible subscriptions, and exports CSV files for Azure
Advisor recommendations and per-subscription/category summaries.

The script is read-only and uses Azure Resource Manager REST calls through Invoke-AzRestMethod.

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
    [string]$AdvisorApiVersion = '2023-01-01'
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
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$resolvedOutputRoots = Resolve-SmartAzureOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Azure\Advisor'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartAzure-Advisor-Inventory_{0}.log" -f $runId)

Import-SmartAzureCoreModule -StartPath $PSScriptRoot
Set-SmartAzureCoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestOutputRoot -LogPath $logPath -RetentionMaxCsv ([int](Get-SmartAzureScriptConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30))

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null

    Write-SmartAzureLog -Message "Starting SmartAzure Advisor inventory. RunId=$runId"
    Import-RequiredModule -Name Az.Accounts

    $context = Connect-SmartAzureCloudSession -TenantId $TenantId -Connect:$Connect -UseDeviceCode:$UseDeviceCode

    $subscriptions = @(Get-AzSubscription)
    if ($SubscriptionId -and $SubscriptionId.Count -gt 0) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $SubscriptionId) { [void]$wanted.Add($id) }
        $subscriptions = @($subscriptions | Where-Object { $wanted.Contains($_.Id) })
    }
    if ($subscriptions.Count -eq 0) { throw "No Azure subscriptions were found for the current identity and filters." }

    $recommendationRows = @()
    $summaryRows = @()

    foreach ($subscription in $subscriptions) {
        Write-SmartAzureLog -Message ("Processing subscription: {0} ({1})" -f $subscription.Name, $subscription.Id)
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
        $encodedSubscriptionId = [uri]::EscapeDataString($subscription.Id)
        $recommendations = Invoke-SmartAzureArmGetPaged -Uri ("https://management.azure.com/subscriptions/{0}/providers/Microsoft.Advisor/recommendations?api-version={1}" -f $encodedSubscriptionId, $AdvisorApiVersion) -Operation "Advisor recommendations for $($subscription.Name)"

        foreach ($recommendation in $recommendations) {
            $properties = Get-ObjectPropertyValue -InputObject $recommendation -PropertyName @('properties')
            $resourceMetadata = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('resourceMetadata')
            $shortDescription = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('shortDescription')
            $recommendationRows += [pscustomobject]@{
                RunId                 = $runId
                SubscriptionId        = $subscription.Id
                SubscriptionName      = $subscription.Name
                Category              = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('category')
                Impact                = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('impact')
                ImpactedField         = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('impactedField')
                ImpactedValue         = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('impactedValue')
                RecommendationTypeId  = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('recommendationTypeId')
                Problem               = Get-ObjectPropertyValue -InputObject $shortDescription -PropertyName @('problem')
                Solution              = Get-ObjectPropertyValue -InputObject $shortDescription -PropertyName @('solution')
                ResourceId            = Get-ObjectPropertyValue -InputObject $resourceMetadata -PropertyName @('resourceId')
                ResourceType          = Get-ObjectPropertyValue -InputObject $resourceMetadata -PropertyName @('source')
                LastUpdated           = Get-ObjectPropertyValue -InputObject $properties -PropertyName @('lastUpdated')
                SuppressionIdsJson    = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('suppressionIds'))
                ExtendedPropertiesJson = ConvertTo-CompactJson -Value (Get-ObjectPropertyValue -InputObject $properties -PropertyName @('extendedProperties'))
                AdvisorResourceId     = $recommendation.id
            }
        }

        foreach ($category in @('HighAvailability', 'Security', 'Performance', 'Cost', 'OperationalExcellence')) {
            $categoryRows = @($recommendations | Where-Object {
                    (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('category')) -eq $category
                })
            $summaryRows += [pscustomobject]@{
                RunId            = $runId
                SubscriptionId   = $subscription.Id
                SubscriptionName = $subscription.Name
                TenantId         = $subscription.TenantId
                Category         = $category
                RecommendationCount = $categoryRows.Count
                HighImpactCount  = @($categoryRows | Where-Object { (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('impact')) -eq 'High' }).Count
                MediumImpactCount = @($categoryRows | Where-Object { (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('impact')) -eq 'Medium' }).Count
                LowImpactCount   = @($categoryRows | Where-Object { (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $_ -PropertyName @('properties')) -PropertyName @('impact')) -eq 'Low' }).Count
            }
        }
    }

    Export-SmartAzureCsv -Name 'Azure_Advisor_Recommendations' -Rows $recommendationRows
    Export-SmartAzureCsv -Name 'Azure_Advisor_Summary' -Rows $summaryRows

    Write-SmartAzureLog -Level SUCCESS -Message ("Completed Advisor inventory. Output={0}; Latest={1}" -f $runOutputRoot, $LatestOutputRoot)
}
catch {
    Write-SmartAzureLog -Level ERROR -Message $_.Exception.Message
    Send-SmartAzureScriptFailureNotification -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) -ErrorRecord $_ -RunId $runId -LogPath $logPath -Config $ScriptLocalConfig
    throw
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH7aQwt+Qc9Bqd
# KW3eHF8Wj95NQnK3xnGYhfc1/n+iM6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCc/mqUgz/jN9UAkglst7+bLGt19pXafIS9+jkTN3KOljANBgkqhkiG9w0B
# AQEFAASCAYAqz5cpsp2QbjZwlSBy8mwgw7diW7IsfELuT/dbdoBQmoPUbwU3hFqf
# ucBIe3FXD4Yh4IMR3i2mIWoWsTsDGedya1dLApVS+viqvqAQuU+MlbD0QshE5pki
# GsjpSxkBF9BdenPmN9g2qEJOpba7d80/tkt6g3H8z7ijvd4mw/8Plbwhb0HcfLbC
# hsAFoAXY94+Gt1XPjTuMzdKv+crckmF5iayAJ6l7+7D6PJnQU+Tm3jRIeWBuPku/
# hgRJgOxyVwz2Dh8VF+9fY/hQhH+VE4Fl3duhS/Jl6AiGg3eHEep43POHcB7YjGC7
# 2LssoeT4ltDKgua5h4dUZvQDcEeNgibr2RnuVB9vqhM8UG1/HGkM616B47oNcN6t
# 7IOC4hm9if18ndsmmvH8BNqPIULIVSq9BXiN1coqyN5KiUX0kLf/JaQAH+Khpl0b
# aUF+XFKz20IVlbba72U6zG+K6hK0cTCv+XS6+j6UblSvl1kE59l4zQuVu8Z95K4H
# TUi6uDpiORQ=
# SIG # End signature block
