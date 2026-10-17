#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks for the per-SKU license decisions: disabled accounts on every paid SKU, Copilot without licensed activity,
  Exchange retention guard, resource accounts, dependency guard, exclusions and SKU capacity.
.VERSION
1.0.0
.NOTES
  Synthetic data only. Uses the real SmartM365.Core module; no tenant, Graph, mail or SharePoint access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline mocks keep the production command signatures.')]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$licensingFolder = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$inventoryPath = Join-Path $licensingFolder 'SmartM365-Licences-Inventory.ps1'
$rulesPath = Join-Path $licensingFolder 'SmartM365-LicenseDependencies.json.txt'
$exclusionsPath = Join-Path $licensingFolder 'SmartM365-LicenseExclusions.json.txt'
$templatePath = Join-Path $licensingFolder 'SmartM365-Licences-Inventory.local.json.template'
$coreManifest = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$names = @(
  'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource',
  'Import-LicensesSourceCsv', 'Read-LicensesIndexedSource', 'Get-LicensesSuiteProduct', 'Get-LicensesFocusedSummaryRows', 'Get-LicensesOptionalProperty', 'Get-LicensesDecisionSettings', 'Read-LicensesDependencyRules',
  'Get-LicensesAccountState', 'Get-LicensesMailboxGuard', 'Get-LicensesCopilotUsageState', 'Get-LicensesCopilotActivity',
  'Get-LicensesSkuDecisions', 'Publish-LicensesSkuReportCsv'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'License SKU decision functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }
Import-Module $coreManifest -MinimumVersion '1.0.90' -Force | Out-Null

$script:checks = 0
function Check {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) { throw $Message }
  $script:checks++
}
function Get-ScriptLocalConfigValue {
  param($Config, [string]$Name, $DefaultValue)
  if ($Config -and $Config.ContainsKey($Name)) { return $Config[$Name] }
  return $DefaultValue
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-LicenseSkuDecisions-' + [guid]::NewGuid().ToString('N'))
$dataLast = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory -Path $dataLast -Force | Out-Null
$now = [datetime]::UtcNow
$day = { param([int]$Offset) $now.Date.AddDays($Offset).ToString('yyyy-MM-dd') }
$stamp = { param([int]$Offset) $now.AddDays($Offset).ToString('o') }

function Write-FixtureCsv {
  param([string]$Name, [object[]]$Rows)
  $path = Join-Path $dataLast $Name
  $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM
  return $path
}
function Write-ForeignReceipt {
  param([string]$Producer, [string]$FileName)
  $path = Join-Path $dataLast $FileName
  @{ Status='Completed'; IsPartialInventory=$false; ConsumerScopeQualified=$true
     Files=@(@{ File=$FileName; Status='Success'; IsPartialInventory=$false; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }) } |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $dataLast "SmartInventory_$Producer.current.json.txt")
}
function New-Proof {
  # Same shape as the accessor proof built by Publish-LicensesRunReport.
  $hashes = [System.Collections.Hashtable]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($name in @('M365_Licenses_Users.csv','M365_Licenses_AssignmentPaths.csv','M365_Licenses_ServicePlans_Catalog.csv')) {
    $path = [IO.Path]::GetFullPath((Join-Path $dataLast $name))
    $hashes[$path] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
  }
  return [pscustomobject]@{ RunId='run-synthetic'; Hashes=$hashes }
}
function Write-CopilotReport {
  param([object[]]$Rows)
  Write-FixtureCsv 'M365_Copilot_UserUsage.csv' $Rows | Out-Null
  Write-ForeignReceipt 'SmartM365-CopilotUsage-Inventory' 'M365_Copilot_UserUsage.csv'
}
function New-CopilotRow {
  param([string]$Upn, [string]$Licensed = '', [string]$Chat = '', [string]$Period = 'D180', [string]$Refresh)
  [pscustomobject]@{ TenantKey='prod'; ReportPeriod=$Period; ReportRefreshDate=$Refresh; UserPrincipalName=$Upn; Microsoft365CopilotLastActivityDate=$Licensed
    TeamsCopilotLastActivityDate=''; WordCopilotLastActivityDate=''; ExcelCopilotLastActivityDate=''; PowerPointCopilotLastActivityDate=''; OutlookCopilotLastActivityDate=''
    OneNoteCopilotLastActivityDate=''; LoopCopilotLastActivityDate=''; CopilotChatLastActivityDate=$Chat; CopilotChatWorkLastActivityDate=''; CopilotChatWebLastActivityDate=''
    EdgeLastActivityDate=''; CopilotAgentLastActivityDate='' }
}

try {
  # Configuration and rules: committed template, versioned exclusions and dependency file are the defaults.
  $template = Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json -AsHashtable
  Check (-not $template.ContainsKey('LicenseDecisionExcludedSkus') -and @($template['LicenseDecisionAdditionalExcludedSkus']).Count -eq 0) 'Template holds only an empty list of additional local exclusions.'
  $settings = Get-LicensesDecisionSettings -Config $template -ExclusionsPath $exclusionsPath
  Check ($settings.CopilotInactivityDays -eq 60 -and $settings.CopilotAssignmentGraceDays -eq 30 -and $settings.MailboxReviewSizeGB -eq 50) 'Template carries the decision defaults.'
  Check ($settings.ExcludedSkus.ContainsKey('POWER_BI_STANDARD') -and $settings.ExcludedSkus.ContainsKey('FLOW_FREE') -and $settings.ExcludedSkus.Count -eq 14) 'Versioned file lists free, trial, viral and capacity SKU exclusions.'
  $local = Get-LicensesDecisionSettings -Config @{ LicenseDecisionAdditionalExcludedSkus=@(@{ SkuPartNumber='visioclient'; Reason='Local test' }, @{ SkuPartNumber='FLOW_FREE'; Reason='Local override attempt' }) } -ExclusionsPath $exclusionsPath
  Check ($local.ExcludedSkus['VISIOCLIENT'] -eq 'Local: Local test' -and $local.ExcludedSkus['FLOW_FREE'] -notlike 'Local*' -and $local.ExcludedSkus.Count -eq 15) 'Local configuration adds exclusions and never replaces a versioned one.'
  foreach ($case in @(@{ Name='missing'; Text=$null }, @{ Name='invalid JSON'; Text='{' }, @{ Name='entry without reason'; Text='{"Version":"1.0","Exclusions":[{"SkuPartNumber":"X"}]}' }, @{ Name='empty'; Text='{"Version":"1.0","Exclusions":[]}' })) {
    $badExclusions = Join-Path $root ('exclusions-' + [guid]::NewGuid().ToString('N') + '.json')
    if ($null -ne $case.Text) { Set-Content -LiteralPath $badExclusions -Value $case.Text }
    $threw = $false; try { [void](Get-LicensesDecisionSettings -Config $template -ExclusionsPath $badExclusions) } catch { $threw = $true }
    Check $threw "Versioned exclusions $($case.Name) are refused."
  }
  $rules = @(Read-LicensesDependencyRules -Path $rulesPath)
  Check ($rules.Count -ge 4 -and @($rules | Where-Object { 'MICROSOFT_365_COPILOT' -in $_.Dependents -and 'SPE_E3' -in $_.Qualifying }).Count -eq 1) 'Copilot dependency is dated and lists E3.'
  Check (@($rules | Where-Object { 'DYN365_PROJECT_OPERATIONS_ATTACH' -in $_.Dependents -and 'DYN365_FINANCE' -in $_.Qualifying -and $_.SourceDate }).Count -eq 1) 'Dynamics attach dependency is dated.'
  $badRules = Join-Path $root 'bad-rules.json'
  '{"Dependencies":[{"Dependent":["X"],"QualifyingAnyOf":["Y"]}]}' | Set-Content -LiteralPath $badRules
  $threw = $false; try { [void](Read-LicensesDependencyRules -Path $badRules) } catch { $threw = $true }
  Check $threw 'A dependency without a dated source is rejected.'
  $threw = $false; try { [void](Get-LicensesDecisionSettings -Config @{ LicenseDecisionAdditionalExcludedSkus=@(@{ Reason='x' }) } -ExclusionsPath $exclusionsPath) } catch { $threw = $true }
  Check $threw 'A local exclusion without SkuPartNumber is rejected.'

  # Tenant catalog: Exchange plans are detected by service plan, never by SKU name.
  $skus = @(
    ,@('sku-e3','SPE_E3',@('EXCHANGE_S_ENTERPRISE','OFFICESUBSCRIPTION'),20,18)
    ,@('sku-cop','Microsoft_365_Copilot',@('M365_COPILOT_APPS'),10,8)
    ,@('sku-eoa','EXCHANGEARCHIVE_ADDON',@('EXCHANGE_S_ARCHIVE_ADDON'),5,3)
    ,@('sku-vis','VISIOCLIENT',@('VISIO_CLIENT_SUBSCRIPTION','EXCHANGE_S_FOUNDATION'),6,6)
    ,@('sku-pbi','POWER_BI_STANDARD',@('BI_AZURE_P0'),1000000,4)
    ,@('sku-fin','DYN365_FINANCE',@('D365_FINANCE'),3,2)
    ,@('sku-poa','DYN365_PROJECT_OPERATIONS_ATTACH',@('D365_PROJECT_OPERATIONS_ATTACH'),3,1)
    ,@('sku-vu','MCOEV_VIRTUALUSER',@('MCOEV_VIRTUALUSER'),2,1)
    ,@('c2fe850d-fbbb-4858-b67d-bd0c6e746da3',"O365_w/o$([char]0xA0)Teams$([char]0xA0)Bundle_M3",@('EXCHANGE_S_ENTERPRISE'),4,1)
  )
  $tenantRows = @(foreach ($sku in $skus) {
    [pscustomobject]@{ Id=$sku[0]; TenantSkuPartNumber=$sku[1]; TenantSkuDisplayName="$($sku[1]) display"; TenantPrepaidEnabled=$sku[3]; TenantConsumedUnits=$sku[4]; TenantPrepaidWarning=0; TenantPrepaidSuspended=0; TenantPrepaidLockedOut=0 }
  })
  $partBySku = @{}; foreach ($sku in $skus) { $partBySku[$sku[0]] = $sku[1] }
  Write-FixtureCsv 'M365_Licenses_ServicePlans_Catalog.csv' @(foreach ($sku in $skus) { foreach ($plan in $sku[2]) { [pscustomobject]@{ TenantKey='prod'; SkuId=$sku[0]; PlanName=$plan } } }) | Out-Null

  # user, cloud/sync immutable id, Entra enabled, AD enabled ('' = no AD row), SKUs, assignment age in days ('' = unknown)
  $users = @(
    ,@('d1','', 'False','', @('sku-e3','sku-vis'), -200)
    ,@('d2','', 'False','', @('sku-eoa'), -200)
    ,@('d3','', 'False','', @('sku-eoa'), -200)
    ,@('d4','', 'False','', @('sku-eoa'), -200)
    ,@('d5','', 'False','', @('sku-pbi'), -200)
    ,@('d6','', 'False','', @('sku-e3','sku-cop'), -200)
    ,@('s1','im-s1','False','True', @('sku-vis'), -200)
    ,@('s2','im-s2','False','False', @('sku-vis'), -200)
    ,@('s3','im-s3','False','', @('sku-vis'), -200)
    ,@('s4','im-s4','True','False', @('sku-vis'), -200)
    ,@('r1','', 'False','', @('sku-e3'), -200)
    ,@('r2','', 'False','', @('sku-vu'), -200)
    ,@('y1','', 'False','', @('sku-fin','sku-poa'), -200)
    ,@('c1','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('c2','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('c3','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('c4','', 'True','', @('sku-e3','sku-cop'), -10)
    ,@('c5','', 'True','', @('sku-e3','sku-cop'), '')
    ,@('c6','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('c7','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('e1','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('e2','', 'True','', @('sku-e3','sku-cop'), -200)
    ,@('a1','', 'True','', @('sku-fin'), -200)
    ,@('x1','', 'False','', @('c2fe850d-fbbb-4858-b67d-bd0c6e746da3'), -200)
  )
  Write-FixtureCsv 'M365_Licenses_Users.csv' @(foreach ($user in $users) { foreach ($skuId in $user[4]) {
    [pscustomobject]@{ TenantKey='prod'; UserId=$user[0]; 'User principal name'="$($user[0])@example.invalid"; 'Display name'="User $($user[0])"; SkuId=$skuId; SkuPartNumber=$partBySku[$skuId]; Source='Group'; GroupsAssigningSku='License group' }
  } }) | Out-Null
  Write-FixtureCsv 'M365_Licenses_AssignmentPaths.csv' @(foreach ($user in $users) { foreach ($skuId in $user[4]) {
    [pscustomobject]@{ TenantKey='prod'; UserId=$user[0]; SkuId=$skuId; AssignmentRoute='Group'; LastUpdatedDateTime=$(if ($user[5] -is [int]) { & $stamp $user[5] } else { '' }) }
  } }) | Out-Null
  Write-FixtureCsv 'M365_Users_Active.csv' @(foreach ($user in $users) {
    [pscustomobject]@{ TenantKey='prod'; 'Object Id'=$user[0]; 'User principal name'="$($user[0])@example.invalid"; AccountEnabled=$user[2]; OnPremisesImmutableId=$user[1] }
  }) | Out-Null
  Write-ForeignReceipt 'SmartM365-ActiveUsers-Inventory' 'M365_Users_Active.csv'
  Write-FixtureCsv 'AD_Users_AllDomains.csv' @(foreach ($user in @($users | Where-Object { $_[1] -and $_[3] })) {
    [pscustomobject]@{ TenantKey='prod'; DomainName='corp.example.invalid'; ImmutableId_AD=$user[1]; UserPrincipalName="$($user[0])@example.invalid"; Enabled=$user[3] }
  }) | Out-Null
  Write-ForeignReceipt 'SmartM365-ActiveDirectory-Inventory' 'AD_Users_AllDomains.csv'
  $mailbox = { param($Id, $Type = 'UserMailbox', $Size = '1', $Archive = 'None', $Litigation = 'False', $Retention = 'False')
    [pscustomobject]@{ TenantKey='prod'; ExternalDirectoryObjectId=$Id; RecipientTypeDetails=$Type; NativeIdentityStatus='Observed'; TotalItemSizeGB=$Size; ArchiveStatus=$Archive; LitigationHoldEnabled=$Litigation; RetentionHoldEnabled=$Retention } }
  Write-FixtureCsv 'Exchange_EXO_Mailboxes_AllDomains.csv' @(
    & $mailbox 'x1'; & $mailbox 'd1'; & $mailbox 'd2' -Litigation 'True'; & $mailbox 'd3' -Size '60'; & $mailbox 'd6'; & $mailbox 'r1' -Type 'RoomMailbox'
  ) | Out-Null
  Write-ForeignReceipt 'SmartM365-EXO-Mailboxes-Inventory' 'Exchange_EXO_Mailboxes_AllDomains.csv'
  Write-FixtureCsv 'M365_Teams_PhoneAssignments.csv' @([pscustomobject]@{ TenantKey='prod'; AssignedPstnTargetId='nobody'; TargetType='User' }) | Out-Null
  Write-ForeignReceipt 'SmartM365-TeamsPhonePstnUsage-Inventory' 'M365_Teams_PhoneAssignments.csv'
  $refresh = & $day -2
  Write-CopilotReport @(
    New-CopilotRow 'c1@example.invalid' -Licensed (& $day -5) -Refresh $refresh
    New-CopilotRow 'c2@example.invalid' -Refresh $refresh
    New-CopilotRow 'c3@example.invalid' -Chat (& $day -5) -Refresh $refresh
    New-CopilotRow 'c4@example.invalid' -Refresh $refresh
    New-CopilotRow 'c5@example.invalid' -Refresh $refresh
    New-CopilotRow 'c7@example.invalid' -Licensed (& $day -72) -Refresh $refresh
    New-CopilotRow 'e1@example.invalid' -Licensed (& $day -3) -Refresh $refresh
    New-CopilotRow 'e2@example.invalid' -Licensed (& $day -3) -Refresh $refresh
    New-CopilotRow 'd6@example.invalid' -Refresh $refresh
  )
  $recovery = @([pscustomobject]@{ License='Microsoft 365 E3'; UserId='x1' }, [pscustomobject]@{ License='Microsoft 365 E3'; UserId='d1' }, [pscustomobject]@{ License='Microsoft 365 E3'; UserId='e2' })
  $downgrade = @([pscustomobject]@{ UserId='e1' })

  $run = { Get-LicensesSkuDecisions -Folder $dataLast -TenantKey 'prod' -TenantRows $tenantRows -CurrentRunPublication (New-Proof) -RecoveryDetails $recovery -DowngradeDetails $downgrade -Settings $settings -DependencyRules $rules -AsOfUtc $now }
  $decisions = & $run
  $row = { param($UserId, $Part) $found = @($decisions.Candidates | Where-Object { $_.UserId -eq $UserId -and $_.SkuPartNumber -eq $Part }); if ($found.Count -gt 1) { throw "duplicate row $UserId $Part" }; if ($found.Count -eq 1) { $found[0] } else { $null } }

  # Disabled accounts on every paid SKU, with Entra + AD state.
  $r = & $row 'd1' 'VISIOCLIENT'
  Check ($r.CandidateType -eq 'Disabled account' -and $r.Recommendation -eq 'Recommended' -and $r.MailboxGuard -eq 'Not applicable') 'Disabled cloud account frees a non-suite SKU; Exchange Foundation is not a mailbox plan.'
  $r = & $row 'd1' 'SPE_E3'
  Check ($r.Recommendation -eq 'Referenced' -and $r.CountedIn -eq 'ReportCandidates:Recovery') 'Suite already in ReportCandidates is referenced, never recounted.'
  Check ((& $row 'd2' 'EXCHANGEARCHIVE_ADDON').Recommendation -eq 'Blocked by dependency') 'Litigation hold blocks an Exchange add-on removal.'
  Check ((& $row 'd3' 'EXCHANGEARCHIVE_ADDON').Reason -like '*mailbox above 50 GB*') 'Mailbox above the review size blocks the removal.'
  $r = & $row 'd4' 'EXCHANGEARCHIVE_ADDON'
  Check ($r.Recommendation -eq 'Review' -and $r.MailboxGuard -like 'Unknown:*') 'Unknown mailbox state fails closed to Review.'
  Check ($null -eq (& $row 'd5' 'POWER_BI_STANDARD')) 'Excluded SKUs produce no candidate.'
  $r = & $row 's1' 'VISIOCLIENT'
  Check ($r.Recommendation -eq 'Review' -and $r.Reason -like '*Entra disabled but AD enabled*') 'Entra/AD conflict is a review.'
  $r = & $row 's2' 'VISIOCLIENT'
  Check ($r.Recommendation -eq 'Recommended' -and $r.AdAccountEnabled -eq 'false') 'Disabled in Entra and AD is recommended.'
  Check ((& $row 's3' 'VISIOCLIENT').Recommendation -eq 'Review') 'Synchronized account missing from AD is a review.'
  Check ((& $row 's4' 'VISIOCLIENT').Reason -like '*AD disabled but Entra enabled*') 'AD-only disablement is a review.'
  # Resource accounts are never plain disabled-account candidates.
  $r = & $row 'r1' 'SPE_E3'
  Check ($r.Recommendation -eq 'Review' -and $r.ResourceAccount -eq 'Yes' -and $r.Reason -like '*RoomMailbox*') 'Room mailbox account is reviewed with its reason.'
  Check ((& $row 'r2' 'MCOEV_VIRTUALUSER').Recommendation -eq 'Review') 'Teams Phone resource account license is reviewed.'
  # Dependency guard.
  Check ((& $row 'y1' 'DYN365_FINANCE').DependencyStatus -eq 'Dependents removed together') 'Dynamics base and attach on a disabled account are removed together.'
  $r = & $row 'd6' 'SPE_E3'
  Check ($r.Recommendation -eq 'Recommended' -and $r.DependencyStatus -eq 'Dependents removed together' -and $r.DependentSkus -eq 'MICROSOFT_365_COPILOT') 'Copilot base removal is valued only with its dependent.'
  Check ((& $row 'd6' 'Microsoft_365_Copilot').CandidateType -eq 'Disabled account') 'Copilot on a disabled account is a disabled-account candidate.'
  $r = & $row 'e2' 'SPE_E3'
  Check ($r.Recommendation -eq 'Referenced' -and $r.DependencyStatus -eq 'Blocked by dependency' -and $r.DependentSkus -eq 'MICROSOFT_365_COPILOT') 'Referenced recovery keeping active Copilot is flagged by the dependency guard.'
  Check ((& $row 'e1' 'SPE_E3').DependencyStatus -eq 'Clear') 'E3 to F3 review keeps a qualifying F3 for Copilot.'
  Check ($null -eq (& $row 'a1' 'DYN365_FINANCE')) 'Active Dynamics users are a signal only, never a candidate.'
  # EEA suites are recognized by SkuId GUID, whatever their part number spelling.
  $r = @($decisions.Candidates | Where-Object { $_.UserId -eq 'x1' })
  Check ($r.Count -eq 1 -and $r[0].Recommendation -eq 'Referenced' -and $r[0].CountedIn -eq 'ReportCandidates:Recovery') 'EEA E3 already in ReportCandidates is referenced by GUID.'
  Check ((Get-LicensesSuiteProduct -SkuId 'C2FE850D-FBBB-4858-B67D-BD0C6E746DA3' -SkuPartNumber 'renamed') -eq 'Microsoft 365 E3') 'EEA E3 GUID maps to E3.'
  Check ((Get-LicensesSuiteProduct -SkuId '3271cf8e-2be5-4a09-a549-70fd05baaa17' -SkuPartNumber '') -eq 'Microsoft 365 E5' -and (Get-LicensesSuiteProduct -SkuId 'f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87' -SkuPartNumber '') -eq 'Microsoft 365 F3' -and (Get-LicensesSuiteProduct -SkuId '0666269f-b167-4c5b-a76f-fc574f2b1118' -SkuPartNumber '') -eq 'Microsoft 365 F1') 'EEA E5, F3 and F1 GUIDs map to their suites.'
  Check ((Get-LicensesSuiteProduct -SkuId 'a23dbafb-3396-48b3-ad9c-a304fe206043' -SkuPartNumber 'Microsoft_365_E3_EEA_(no_Teams)_Unattended_License') -eq '' -and (Get-LicensesSuiteProduct -SkuId 'x' -SkuPartNumber 'spe_e3') -eq 'Microsoft 365 E3') 'Unattended licenses stay out; part numbers still match case-insensitively.'
  $summary = @(Get-LicensesFocusedSummaryRows -TenantRows $tenantRows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Check ($summary.Enabled -eq 24 -and $summary.Consumed -eq 19) 'Summary capacity adds the EEA E3 subscription to E3.'

  # Copilot without licensed activity.
  Check ($null -eq (& $row 'c1' 'Microsoft_365_Copilot')) 'Recent licensed Copilot activity keeps the license.'
  $r = & $row 'c2' 'Microsoft_365_Copilot'
  Check ($r.CandidateType -eq 'Copilot without activity' -and $r.Recommendation -eq 'Recommended' -and $r.ActivityWindowEnd -eq $refresh) 'No licensed Copilot activity is recommended; the window ends at the report refresh date.'
  $r = & $row 'c3' 'Microsoft_365_Copilot'
  Check ($r.Recommendation -eq 'Review' -and $r.Reason -like 'Copilot Chat only*') 'Copilot Chat only is a review.'
  Check ($null -eq (& $row 'c4' 'Microsoft_365_Copilot')) 'Recent assignment is in the grace period.'
  Check ((& $row 'c5' 'Microsoft_365_Copilot').Reason -like '*Assignment date unknown*') 'Unknown assignment date is a review.'
  Check ($null -eq (& $row 'c6' 'Microsoft_365_Copilot')) 'A user absent from the report is not measured.'
  Check ((& $row 'c7' 'Microsoft_365_Copilot').Recommendation -eq 'Recommended') 'Licensed activity older than the window is a candidate.'
  Check ($null -eq (& $row 'c2' 'SPE_E3')) 'Enabled accounts produce no suite row outside ReportCandidates.'
  # Capacity.
  $capacity = @{}; foreach ($item in $decisions.Capacity) { $capacity[$item.SkuPartNumber] = $item }
  Check ($capacity.Count -eq $skus.Count) 'Every tenant SKU has a capacity row.'
  Check ($capacity['POWER_BI_STANDARD'].Excluded -and $capacity['POWER_BI_STANDARD'].ExclusionReason) 'Excluded SKU keeps its exclusion reason.'
  Check ($capacity['SPE_E3'].AvailableUnits -eq 2 -and $capacity['SPE_E3'].UsageEvidence -eq 'Not applicable') 'Capacity counts unassigned units.'
  Check ($capacity['Microsoft_365_Copilot'].UsageEvidence -eq 'Measured' -and $capacity['Microsoft_365_Copilot'].CopilotGraceUsers -eq 1) 'Copilot usage is measured with grace users counted.'
  Check ($capacity['DYN365_FINANCE'].UsageEvidence -eq 'Not measured' -and $capacity['DYN365_FINANCE'].ActiveUsersUsageNotMeasured -eq 1) 'Dynamics active users are not measured.'
  Check ($capacity['VISIOCLIENT'].DisabledAccountUsers -eq 2 -and $capacity['VISIOCLIENT'].RecommendedCandidates -eq 2 -and $capacity['VISIOCLIENT'].ReviewCandidates -eq 3) 'Capacity counts by recommendation.'

  # Publication: snapshot identity on every row, no pending file left.
  $paths = @(Publish-LicensesSkuReportCsv -Folder $dataLast -SnapshotId 'snap-1' -TenantKey 'prod' -ReceiptRunId 'run-1' -Decisions $decisions)
  Check ($paths.Count -eq 2) 'Two decision CSV files are published.'
  $published = @(Import-Csv -LiteralPath $paths[0])
  Check ($published.Count -eq @($decisions.Candidates).Count -and @($published | Where-Object { $_.SnapshotId -ne 'snap-1' -or $_.ReceiptRunId -ne 'run-1' -or $_.TenantKey -ne 'prod' }).Count -eq 0) 'Candidate rows carry snapshot, tenant and receipt run.'
  Check (@(Import-Csv -LiteralPath $paths[1]).Count -eq $skus.Count) 'Capacity CSV holds one row per SKU.'
  Check (@(Get-ChildItem -LiteralPath $dataLast -Filter '*.pending').Count -eq 0) 'No pending file remains.'
  $empty = [pscustomobject]@{ Candidates=@(); Capacity=@() }
  $emptyPaths = @(Publish-LicensesSkuReportCsv -Folder $dataLast -SnapshotId 'snap-2' -TenantKey 'prod' -ReceiptRunId 'run-2' -Decisions $empty)
  Check ((Get-Content -LiteralPath $emptyPaths[0] -TotalCount 1) -like '"SnapshotId","TenantKey","ReceiptRunId","CandidateType"*') 'An empty decision CSV keeps its header.'

  # Copilot report not measured: anonymized names, short period, stale refresh.
  foreach ($case in @(
      @{ Name='anonymized'; Rows=@(New-CopilotRow '6A1F0C9B2E7D4A3C' -Refresh $refresh) }
      @{ Name='short period'; Rows=@(New-CopilotRow 'c2@example.invalid' -Period 'D30' -Refresh $refresh) }
      @{ Name='stale'; Rows=@(New-CopilotRow 'c2@example.invalid' -Refresh (& $day -30)) })) {
    Write-CopilotReport $case.Rows
    $decisions = & $run
    Check (@($decisions.Candidates | Where-Object CandidateType -eq 'Copilot without activity').Count -eq 0 -and -not $decisions.CopilotMeasured) "Copilot report $($case.Name) is not measured and produces no candidate."
    Check (@($decisions.Capacity | Where-Object SkuPartNumber -eq 'Microsoft_365_Copilot')[0].UsageEvidence -eq 'Not measured') "Copilot capacity is Not measured when the report is $($case.Name)."
    Check ($null -ne (& $row 'd6' 'Microsoft_365_Copilot')) "Disabled-account candidates do not depend on Copilot usage ($($case.Name))."
  }

  # Missing EXO source: every Exchange-plan SKU on a disabled account fails closed.
  Remove-Item -LiteralPath (Join-Path $dataLast 'Exchange_EXO_Mailboxes_AllDomains.csv')
  $decisions = & $run
  Check ((& $row 'd1' 'VISIOCLIENT').Recommendation -eq 'Recommended') 'Non-Exchange SKU stays recommended without EXO.'
  Check ((& $row 'd6' 'SPE_E3').Recommendation -eq 'Review') 'Exchange-plan SKU is a review without EXO.'
  # Foreign tenant in the license source is refused.
  Write-FixtureCsv 'M365_Licenses_ServicePlans_Catalog.csv' @([pscustomobject]@{ TenantKey='other'; SkuId='sku-e3'; PlanName='X' }) | Out-Null
  $threw = $false; try { [void](& $run) } catch { $threw = $true }
  Check $threw 'Another tenant in the license sources is refused.'
  # A license file not published by the current run is refused.
  $proof = New-Proof
  Write-FixtureCsv 'M365_Licenses_ServicePlans_Catalog.csv' @([pscustomobject]@{ TenantKey='prod'; SkuId='sku-e3'; PlanName='EXCHANGE_S_ENTERPRISE' }) | Out-Null
  $threw = $false; try { [void](Get-LicensesSkuDecisions -Folder $dataLast -TenantKey 'prod' -TenantRows $tenantRows -CurrentRunPublication $proof -Settings $settings -DependencyRules $rules -AsOfUtc $now) } catch { $threw = $_.Exception.Message -like '*current-run publication*' }
  Check $threw 'License file changed after the current-run publication is refused.'

  "PASS: License SKU decision offline checks ($script:checks checks)."
}
finally {
  Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
