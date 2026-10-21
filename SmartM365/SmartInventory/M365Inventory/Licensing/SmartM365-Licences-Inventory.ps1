<#
.SYNOPSIS
  Export M365 license assignments with Users/Tenant/Groups plus compact user service-plan state codes and a service-plan catalog.
  WeeklyHistory retains the detailed IsEnabled and PlanStatus representation.
  Detects Direct vs Group via user.LicenseAssignmentStates.assignedByGroup.
  Maps SKU & Service Plan friendly names from the Microsoft CSV (default: script folder).
  Every full run publishes the license report (snapshot + M365_Licenses_Report*.csv) as optional outputs of its source receipt.
  DATA-LAST changes only when a run completes: CSVs and report snapshot are staged and promoted with the receipt (Core 1.0.95).
  A user whose license details return 404 is excluded only when GET /users/{id} confirms the deletion; the count is
  bounded (LicenseDeletedUsersMaxCount, default 100, and LicenseDeletedUsersMaxPercent, default 1) and recorded
  as the receipt qualification UsersDeletedDuringCollection. Any other user error still stops publication.
  The email, including SendLicenseSummaryEmailOnly, renders that published report after checking it against the receipt; it never recomputes.
  ForceAdCsvAnalysis (full runs only) uses a fresh, structurally valid AD CSV for the report when only its collector receipt is rejected.
  ForceLicenseSummaryEmail sends the report again even when it was already sent on the current Europe/Paris day.
  The report also lists target-suite users awaiting license capacity and other assignment errors from qualified existing CSVs.
  Qualified PartialAccepted AD coverage is usable but remains provisional; available allowlisted domains remain included.
  Per-SKU decisions (disabled accounts on every paid SKU, Copilot without licensed activity, unassigned capacity, dependency guard) are published
  as optional files M365_Licenses_ReportSkuCandidates.csv and M365_Licenses_ReportSkuCapacity.csv with the report snapshot identity;
  ReportSkuCapacity names the report product family of each SKU (ReportProduct) so consumers never re-implement the SKU rules.
.VERSION
 1.58
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Identity.DirectoryManagement; Microsoft.Graph.Users; Microsoft.Graph.Groups; ImportExcel for the report attachment.
    Minimum Graph application permissions: Directory.Read.All; User.Read.All; Group.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
  Author: https://github.com/khda79/workplacecloudhub.com
     Version : 1.51
  PowerShell: PowerShell 7+
  Minimum application permissions: Directory.Read.All, User.Read.All, Group.Read.All
  Requires: Microsoft.Graph.Authentication
            Microsoft.Graph.Identity.DirectoryManagement
            Microsoft.Graph.Users
            Microsoft.Graph.Groups
            SmartM365.Core.psd1
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
  [switch]$Connect,
  [int]$TopUsers = 0,
  [switch]$FastSample,
  [switch]$ServicePlans,
  [string]$SkuNameCsvPath = $(Join-Path $PSScriptRoot 'Product names and service plan identifiers for licensing.csv'),
  [switch]$RequireSkuNameCsv,
  [switch]$InteractiveAuth,
  [switch]$SendLicenseSummaryEmailOnly,
  [switch]$ForceLicenseSummaryEmail,
  [switch]$ForceAdCsvAnalysis,
    [int]$MaxItems = 0
)
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

# Defaults: enable RequireSkuNameCsv and ServicePlans unless explicitly set
if (-not $PSBoundParameters.ContainsKey('RequireSkuNameCsv')) { $RequireSkuNameCsv = $true }
if (-not $PSBoundParameters.ContainsKey('ServicePlans'))      { $ServicePlans      = $true }

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

# Avoid PS function-capacity issues
$MaximumFunctionCount = 32768

# ==========================================================
# App-only authentication parameters (same app as other scripts)
# ==========================================================
function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)); $configPath = Resolve-SmartM365JsonConfigurationPath -Path $configPath
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = (Get-SmartM365JsonTemplateName -Path $configPath)
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) {
                $message = @(
                    "Local configuration file not found: $configPath",
                    "Template to copy is also missing: $templatePath",
                    'Create the .local.json file from a safe template, then run the script again.'
                ) -join [Environment]::NewLine
                throw $message
            }

            Write-SmartM365JsonBytesAtomically -Path $configPath -Bytes ([IO.File]::ReadAllBytes($templatePath)) -ExpectedSHA256 'ABSENT' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration template must be an object.'}} | Out-Null
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($ScriptRoot) { $ScriptRoot } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'; $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}
function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'; $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

$ScriptLocalConfig = Get-ScriptLocalConfig


$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.98' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

function Get-PrimarySmtpAddress {
    param([Parameter(Mandatory=$true)]$User)

    # 1) Prefer the Exchange primary SMTP from proxyAddresses (uppercase 'SMTP:')
    if ($User.ProxyAddresses) {
        $pri = $User.ProxyAddresses |
            Where-Object { $_ -cmatch '^SMTP:' } |
            Select-Object -First 1

        if ($pri) {
            return (($pri -replace '^SMTP:', '').Trim())
        }
    }

    # 2) Fallback to Mail if present (can be blank or not primary, but better than secondary alias)
    if (-not [string]::IsNullOrWhiteSpace([string]$User.Mail)) {
        return ([string]$User.Mail).Trim()
    }

    # 3) Last resort: UPN
    if (-not [string]::IsNullOrWhiteSpace([string]$User.UserPrincipalName)) {
        return ([string]$User.UserPrincipalName).Trim()
    }

    return $null
}

# ---------------- Mapping loader (SKU + Plans) ----------------
# CSV columns supported:
#   Product_Display_Name, String_Id (SkuPartNumber), GUID (SkuId),
#   Service_Plan_Name, Service_Plan_Id, Service_Plans_Included_Friendly_Names
function Load-SkuNameMap {
  param([string]$Path)

  $mapSkuByPart = @{}  # String_Id (UPPER) -> Product_Display_Name
  $mapSkuById   = @{}  # [Guid] Sku GUID   -> Product_Display_Name
  $mapSvcByName = @{}  # unused for this CSV (kept for compatibility)
  $mapSvcById   = @{}  # [Guid] Service_Plan_Id -> FriendlyName (rebuilt)

  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
    Write-Warning "SkuNameCsvPath not found: $Path"
    return @($mapSkuByPart,$mapSkuById,$mapSvcByName,$mapSvcById)
  }

  try {
    $rows = Import-Csv -Path $Path

    # Temp stores to align plan IDs with friendly names per product
    $planIdsBySkuId       = @{} # [Guid Sku] -> List[Guid] of Service_Plan_Id (unique, ordered)
    $friendlyListBySkuId  = @{} # [Guid Sku] -> List[string] of friendly names

    $nSkuP=0; $nSkuI=0; $nSvcI=0

    foreach ($r in $rows) {
      # ---- SKU mapping ----
      $prodName = $r.Product_Display_Name
      $skuPN    = $r.String_Id
      $skuIdTxt = $r.GUID

      if ($prodName -and $skuPN) {
        $key = $skuPN.ToUpper()
        if (-not $mapSkuByPart.ContainsKey($key)) { $mapSkuByPart[$key] = $prodName; $nSkuP++ }
      }
      $skuGuid = $null
      if ($prodName -and $skuIdTxt) {
        try { $skuGuid = [Guid]$skuIdTxt } catch { $skuGuid = $null }
        if ($skuGuid -and -not $mapSkuById.ContainsKey($skuGuid)) { $mapSkuById[$skuGuid] = $prodName; $nSkuI++ }
      }

      # ---- Plan IDs (one per line) ----
      $planIdTxt = $r.Service_Plan_Id
      if ($skuGuid -and $planIdTxt) {
        try { $planGuid = [Guid]$planIdTxt } catch { $planGuid = $null }
        if ($planGuid) {
          if (-not $planIdsBySkuId.ContainsKey($skuGuid)) {
            $planIdsBySkuId[$skuGuid] = New-Object System.Collections.Generic.List[System.Guid]
          }
          if (-not $planIdsBySkuId[$skuGuid].Contains($planGuid)) {
            $planIdsBySkuId[$skuGuid].Add($planGuid) | Out-Null
          }
        }
      }

      # ---- Friendly list per product (aggregated) ----
      $agg = $r.Service_Plans_Included_Friendly_Names
      if ($skuGuid -and $agg -and -not $friendlyListBySkuId.ContainsKey($skuGuid)) {
        $parts = @()
        foreach ($p in ($agg -split '[;,\|]')) {
          $t = ($p -as [string]).Trim()
          if ($t) { $parts += $t }
        }
        $friendlyListBySkuId[$skuGuid] = $parts
      }
    }

    # ---- Build PlanId -> FriendlyName by index alignment when counts match ----
    foreach ($skuKey in $planIdsBySkuId.Keys) {
      $ids = $planIdsBySkuId[$skuKey]
      $friendly = $friendlyListBySkuId[$skuKey]
      if ($ids -and $friendly -and $ids.Count -eq $friendly.Count) {
        for ($i=0; $i -lt $ids.Count; $i++) {
          $planIdToMap = $ids[$i]
          $fn = $friendly[$i]
          if ($planIdToMap -and $fn -and -not $mapSvcById.ContainsKey($planIdToMap)) {
            $mapSvcById[$planIdToMap] = $fn
            $nSvcI++
          }
        }
      }
    }

    Write-Host ("SKU/Service mapping loaded: SKU(byPart)={0}, SKU(byId)={1}, Svc(byId)={2}" -f $nSkuP,$nSkuI,$nSvcI)
  } catch {
    Write-Warning "Failed to parse CSV mapping: $($_.Exception.Message)"
  }

  return @($mapSkuByPart,$mapSkuById,$mapSvcByName,$mapSvcById)
}

function Get-SkuDisplayName {
  param([Guid]$SkuId,[string]$SkuPartNumber,[hashtable]$MapByPart,[hashtable]$MapById)
  if ($SkuId -and $MapById -and $MapById.ContainsKey($SkuId)) { return $MapById[$SkuId] }
  if ($SkuPartNumber) { $key=$SkuPartNumber.ToUpper(); if ($MapByPart -and $MapByPart.ContainsKey($key)) { return $MapByPart[$key] } }
  return $SkuPartNumber
}

function Get-ServiceFriendly {
  param([string]$PlanName,[Guid]$PlanId,[hashtable]$ByName,[hashtable]$ById)
  if ($PlanId -and $ById -and $ById.ContainsKey($PlanId)) { return $ById[$PlanId] }
  return $PlanName
}

# ---------------- Robust helpers ----------------
function Invoke-GraphWithRetry {
  [CmdletBinding()]
  param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[int]$MaxRetries=6,[int]$BaseDelaySeconds=2)
  $attempt=0
  while ($true) {
    try { return & $ScriptBlock } catch {
      $attempt++; $msg=$_.Exception.Message
      # Get-Mg* errors carry the status in ResponseStatusCode or only in the message (InternalServerError).
      $statusCode=Get-SmartM365GraphErrorStatusCode -ErrorRecord $_
      $isTransient=$statusCode -notin @(401,403) -and ($statusCode -in @(408,409,429,500,502,503,504) -or $msg -match '(?i)throttl|TooManyRequests|temporarily unavailable|timeout|timed out|An error occurred while sending the request|connection (?:was )?(?:reset|closed|aborted|forcibly closed)')
      if (-not $isTransient) { throw }
      if ($attempt -ge $MaxRetries) { throw "Max retry attempts reached: $msg" }
      $retryAfter=$null; try {
        if ($_.Exception.Response -and $_.Exception.Response.Headers) { $retryAfter = @($_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1)[0] }
        elseif ($_.Exception.Data['Retry-After']) { $retryAfter = $_.Exception.Data['Retry-After'] }
      } catch {}
      $retrySeconds=0
      if ($retryAfter -and [int]::TryParse([string]$retryAfter,[ref]$retrySeconds) -and $retrySeconds -gt 0) { $delay=[Math]::Min(300,$retrySeconds) } else { $delay=[math]::Min(60, ($BaseDelaySeconds * [math]::Pow(2, $attempt))) }
      Write-Warning ("Graph throttled/unavailable (attempt {0}/{1}). Sleeping {2}s. Error: {3}" -f $attempt,$MaxRetries,$delay,$msg)
      Start-Sleep -Seconds $delay
    }
  }
}

# True when an error (or one of its inner exceptions) is a Graph 404 / Request_ResourceNotFound.
function Test-LicensesGraphNotFound {
  param([Parameter(Mandatory)]$ErrorRecord)
  $exception = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $ErrorRecord.Exception } else { $ErrorRecord }
  while ($exception) {
    $status = $null
    try { if ($exception.PSObject.Properties['ResponseStatusCode']) { $status = [int]$exception.ResponseStatusCode } } catch { $status = $null }
    try { if ($null -eq $status -and $exception.Response) { $status = [int]$exception.Response.StatusCode } } catch { $status = $null }
    if ($status -eq 404 -or [string]$exception.Message -match 'Request_ResourceNotFound|Status:\s*404\b') { return $true }
    $exception = $exception.InnerException
  }
  return $false
}

# Classifies a failed user. Only a 404 on the license details, confirmed by a direct directory lookup,
# counts as a user deleted during collection; its assignment paths are removed. Anything else is a failure.
function Register-LicensesUserFailure {
  param(
    [Parameter(Mandatory)]$ErrorRecord,
    [Parameter(Mandatory)][string]$UserId,
    [Parameter(Mandatory)][bool]$LicenseDetailStage,
    [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$AssignmentPaths,
    [Parameter(Mandatory)][int]$AssignmentPathCountBefore,
    [Parameter(Mandatory)][hashtable]$State
  )
  if ($LicenseDetailStage -and -not [string]::IsNullOrWhiteSpace($UserId) -and (Test-LicensesGraphNotFound -ErrorRecord $ErrorRecord)) {
    $confirmed = $false
    try {
      $null = Invoke-GraphWithRetry { Get-MgUser -UserId $UserId -Property 'Id' -ErrorAction Stop }
      WriteLog -Message ("License details returned 404 but user {0} still exists; the user counts as a resolution failure." -f $UserId) 'ERROR'
    }
    catch {
      if (Test-LicensesGraphNotFound -ErrorRecord $_) { $confirmed = $true }
      else { WriteLog -Message ("Deletion of user {0} could not be confirmed: {1}" -f $UserId, $_.Exception.Message) 'ERROR' }
    }
    if ($confirmed) {
      while ($AssignmentPaths.Count -gt $AssignmentPathCountBefore) { $AssignmentPaths.RemoveAt($AssignmentPaths.Count - 1) }
      $State.UsersDeletedDuringCollection = [int]$State.UsersDeletedDuringCollection + 1
      WriteLog -Message ("User {0} was deleted from the directory during collection and is excluded from the licensing CSVs." -f $UserId) 'WARNING'
      return 'Deleted'
    }
  }
  $State.UserResolutionFailures = [int]$State.UserResolutionFailures + 1
  return 'Failure'
}

# Stops publication on any unresolved user, or when more users were deleted than the configured limit.
function Assert-LicensesUserResolution {
  param(
    [Parameter(Mandatory)][hashtable]$State,
    [Parameter(Mandatory)][int]$ListedUsers,
    [int]$MaxDeletedCount = 100,
    [double]$MaxDeletedPercent = 1
  )
  if ([int]$State.UserResolutionFailures -gt 0) {
    throw "License collection is incomplete: $([int]$State.UserResolutionFailures) user(s) could not be resolved. Current CSV publication is stopped."
  }
  $deleted = [int]$State.UsersDeletedDuringCollection
  if ($deleted -eq 0) { return $null }
  $limit = [int][math]::Min([math]::Max(0, $MaxDeletedCount), [math]::Ceiling([math]::Max(0, $ListedUsers) * [math]::Max(0, $MaxDeletedPercent) / 100.0))
  if ($deleted -gt $limit) {
    throw "License collection is incomplete: $deleted user(s) were deleted during collection, above the limit of $limit; a mass deletion needs review. Current CSV publication is stopped."
  }
  WriteLog -Message ("Users deleted during collection: {0} of {1} listed (limit {2}); they are excluded from the licensing CSVs." -f $deleted, $ListedUsers, $limit) 'WARNING'
  return ('UsersDeletedDuringCollection={0}' -f $deleted)
}

function Ensure-GraphModules {
  [CmdletBinding()]
  param()

  $requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups'
  )

  foreach ($moduleName in $requiredModules) {
    $module = Get-Module -ListAvailable -Name $moduleName |
      Sort-Object Version -Descending |
      Select-Object -First 1

    if (-not $module) {
      throw "Required Microsoft Graph module '$moduleName' is not installed. Install it with: Install-Module $moduleName -Scope CurrentUser"
    }

    Import-Module $moduleName -ErrorAction Stop | Out-Null
  }
}

function Get-InnerExceptionSummary {
  param([System.Exception]$Exception)

  $innerMessages = New-Object System.Collections.Generic.List[string]
  $inner = if ($Exception) { $Exception.InnerException } else { $null }
  while ($null -ne $inner) {
    if (-not [string]::IsNullOrWhiteSpace($inner.Message)) {
      $innerMessages.Add($inner.Message) | Out-Null
    }
    $inner = $inner.InnerException
  }

  return ($innerMessages -join " | ")
}

function Get-GeneratedCsvSummary {
  $paths = @()
  if ($global:csvGeneratedPaths) {
    $paths = @($global:csvGeneratedPaths | Sort-Object -Unique)
  }

  if ($paths.Count -eq 0) {
    return ""
  }

  return ($paths | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "; "
}

function Remove-LegacyDetailedServicePlansExport {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$CurrentOutputPath,
    [Parameter(Mandatory)][string]$LatestOutputPath
  )

  if (Test-SmartM365MaxItemsMode) {
    WriteLog -Message "Legacy detailed ServicePlans retirement skipped during MaxItems validation." "INFO"
    return
  }

  $legacyFileName = "M365_Licenses_ServicePlans_Detailed.csv"
  if (-not $PSCmdlet.ShouldProcess($legacyFileName, 'Retire legacy detailed ServicePlans export after compact publication')) {
    return
  }

  $legacyLatestPath = Join-Path -Path $LatestOutputPath -ChildPath $legacyFileName

  if ($global:EnableSharePointUpload) {
    $removedFromSharePoint = Remove-SmartM365SharePointFile -LocalFilePath $legacyLatestPath
    if (-not $removedFromSharePoint) {
      WriteLog -Message "Legacy detailed ServicePlans SharePoint file could not be confirmed as removed." "WARN"
    }
  }

  $legacyPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  [void]$legacyPaths.Add((Join-Path -Path $CurrentOutputPath -ChildPath $legacyFileName))
  [void]$legacyPaths.Add($legacyLatestPath)

  foreach ($legacyPath in $legacyPaths) {
    if (-not (Test-Path -LiteralPath $legacyPath)) {
      continue
    }

    try {
      Remove-Item -LiteralPath $legacyPath -Force -ErrorAction Stop
      WriteLog -Message ("Legacy detailed ServicePlans CSV removed: {0}" -f $legacyPath) "INFO"
    }
    catch {
      throw ("Failed to remove legacy detailed ServicePlans CSV '{0}': {1}" -f $legacyPath, $_.Exception.Message)
    }
  }

  WriteLog -Message "Legacy export M365_Licenses_ServicePlans_Detailed.csv is disabled." "INFO"
}
function Send-LicensesInventoryErrorNotification {
  param(
    [Parameter(Mandatory)]$ErrorRecord,
    [string]$Operation,
    [string]$OutputPath
  )

  try {
    $exception = $ErrorRecord.Exception
    $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
    $errorContext = @(
      "Script: $scriptName"
      "Tenant/Organization: $OrgDomain"
      "Operation: $Operation"
      "Error: $($exception.Message)"
      "Output path: $OutputPath"
    ) -join "`n"

    $helpUrl = "https://chat.openai.com/?q={0}" -f [System.Uri]::EscapeDataString("Help troubleshoot this SmartM365 M365 licenses inventory error:`n$errorContext")
    $facts = @{
      "Script name"         = $scriptName
      "Tenant/Organization" = $OrgDomain
      "Failed operation"    = $Operation
      "Exception message"   = $exception.Message
      "Inner exception"     = Get-InnerExceptionSummary -Exception $exception
      "Log path"            = $global:LogTextFile
      "Transcript path"     = $global:logTranscriptFile
      "Output path"         = $OutputPath
      "Generated CSV files" = Get-GeneratedCsvSummary
    }

    Send-SmartM365TeamsNotification `
      -Title "SmartM365 M365 licenses inventory failed" `
      -Message "A terminal error occurred in Microsoft 365 licenses inventory." `
      -Level "ERROR" `
      -Channel "Alerts" `
      -Facts $facts `
      -HelpUrl $helpUrl | Out-Null
  }
  catch {
    WriteLog -Message ("Failed to send Teams error notification: {0}" -f $_.Exception.Message) "ERROR"
  }
}

function Send-LicensesInventorySuccessNotification {
  param(
    [int]$UsersProcessed,
    [int]$UserLicenseRows,
    [int]$ServicePlanRows,
    [int]$TenantSkuRows,
    [int]$GroupRows,
    [int]$UsersDeletedDuringCollection = 0,
    [string]$OutputPath
  )

  try {
    $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
    $resultSummary = "Microsoft 365 licenses inventory completed without error. Users processed: {0}; user license rows: {1}; service plan rows: {2}; tenant SKUs: {3}; groups: {4}; users deleted during collection (excluded): {5}." -f $UsersProcessed, $UserLicenseRows, $ServicePlanRows, $TenantSkuRows, $GroupRows, $UsersDeletedDuringCollection
    $facts = @{
      "Script name"         = $scriptName
      "Tenant/Organization" = $OrgDomain
      "Users processed"     = $UsersProcessed
      "User license rows"   = $UserLicenseRows
      "Service plan rows"   = $ServicePlanRows
      "Tenant SKU rows"     = $TenantSkuRows
      "Group rows"          = $GroupRows
      "Users deleted during collection" = $UsersDeletedDuringCollection
      "Output path"         = $OutputPath
      "Generated CSV files" = Get-GeneratedCsvSummary
      "Log path"            = $global:LogTextFile
      "Transcript path"     = $global:logTranscriptFile
    }

    Send-SmartM365TeamsNotification `
      -Title "SmartM365 M365 licenses inventory success" `
      -Message $resultSummary `
      -Level "SUCCESS" `
      -Channel "Infos" `
      -ResultSummary $resultSummary `
      -Facts $facts | Out-Null
  }
  catch {
    WriteLog -Message ("Failed to send Teams completion notification: {0}" -f $_.Exception.Message) "WARN"
  }
}

function Get-LicensesSuiteProduct {
  param([AllowNull()][string]$SkuId, [AllowNull()][string]$SkuPartNumber)
  # Suites by part number, plus EEA and no-Teams editions by SkuId GUID (Microsoft product names and service plan identifiers CSV).
  # Unattended licenses are not user suites and stay out.
  $part = ([string]$SkuPartNumber).Trim().ToUpperInvariant()
  $id = ([string]$SkuId).Trim().ToLowerInvariant()
  if ($part -in @('M365_F1','M365_F1_COMM') -or $id -in @('0666269f-b167-4c5b-a76f-fc574f2b1118')) { return 'Microsoft 365 F1' }
  if ($part -eq 'SPE_F1' -or $id -in @('f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87')) { return 'Microsoft 365 F3' }
  if ($part -eq 'SPE_E3' -or $id -in @('dcf0408c-aaec-446c-afd4-43e3683943ea','c2fe850d-fbbb-4858-b67d-bd0c6e746da3','602e6573-55a3-46b1-a1a0-cc267991501a')) { return 'Microsoft 365 E3' }
  if ($part -eq 'SPE_E5' -or $id -in @('18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e','3271cf8e-2be5-4a09-a549-70fd05baaa17','1e988bf3-8b7c-4731-bec0-4e2a2946600c',
      '6ee4114a-9b2d-4577-9e7a-49fa43d222d3','90277bc7-a6fe-4181-99d8-712b08b8d32b','a640eead-25f6-4bec-97e3-23cfd382d7c2')) { return 'Microsoft 365 E5' }
  return ''
}

function Get-LicensesFocusedSummaryRows {
  param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows)

  $products = @(
    @{ Name = 'Microsoft 365 F1'; PartNumbers = @('M365_F1', 'M365_F1_COMM') }
    @{ Name = 'Microsoft 365 F3'; PartNumbers = @('SPE_F1') }
    @{ Name = 'Microsoft 365 E3'; PartNumbers = @('SPE_E3') }
    @{ Name = 'Microsoft 365 E5'; PartNumbers = @('SPE_E5') }
  )
  foreach ($product in $products) {
    $enabled = [long]0
    $consumed = [long]0
    $found = $false
    foreach ($tenantRow in $TenantRows) {
      if ((Get-LicensesSuiteProduct -SkuId (Get-LicensesOptionalProperty $tenantRow 'Id') -SkuPartNumber $tenantRow.TenantSkuPartNumber) -ne $product.Name) { continue }
      if ($null -eq $tenantRow.TenantPrepaidEnabled -or $null -eq $tenantRow.TenantConsumedUnits) {
        throw "License counts are unavailable for SKU '$($tenantRow.TenantSkuPartNumber)'."
      }
      $found = $true
      $enabled += [long]$tenantRow.TenantPrepaidEnabled
      $consumed += [long]$tenantRow.TenantConsumedUnits
    }
    [pscustomobject]@{
      Product = $product.Name
      Enabled = $enabled
      Consumed = $consumed
      Subscribed = $found
    }
  }
}

function Get-LicensesOtherProduct {
  param([AllowNull()][string]$SkuPartNumber)
  $sku = ([string]$SkuPartNumber).Trim()
  if ($sku -in @('Microsoft_365_Copilot','M365_COPILOT')) { return 'Microsoft 365 Copilot' }
  if ($sku -match '^(DYN365_|DYNAMICS_365_)' -and $sku -notmatch '(SANDBOX|TRIAL|PREVIEW|VIRAL|FREE|DEMO|TEST)') { return 'Dynamics 365' }
  if ($sku -in @('POWER_BI_PRO','PBI_PREMIUM_PER_USER')) { return 'Power BI' }
  return ''
}

function Get-LicensesAdditionalOverviewRows {
  param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows)

  $totals = @{}
  foreach ($name in @('Microsoft 365 Copilot','Dynamics 365','Power BI')) {
    $totals[$name] = @{ Enabled=[long]0; Consumed=[long]0; Subscribed=$false }
  }
  foreach ($tenantRow in $TenantRows) {
    $sku = ([string]$tenantRow.TenantSkuPartNumber).Trim()
    $name = Get-LicensesOtherProduct -SkuPartNumber $sku
    if (-not $name) { continue }
    if ($null -eq $tenantRow.TenantPrepaidEnabled -or $null -eq $tenantRow.TenantConsumedUnits) {
      throw "License counts are unavailable for SKU '$sku'."
    }
    $totals[$name].Enabled += [long]$tenantRow.TenantPrepaidEnabled
    $totals[$name].Consumed += [long]$tenantRow.TenantConsumedUnits
    $totals[$name].Subscribed = $true
  }
  foreach ($name in @('Microsoft 365 Copilot','Dynamics 365','Power BI')) {
    [pscustomobject]@{ Product=$name; Enabled=$totals[$name].Enabled; Consumed=$totals[$name].Consumed; Subscribed=$totals[$name].Subscribed }
  }
}

function New-LicensesOverviewCardHtml {
  param([Parameter(Mandatory)]$Row, [int]$Width, [Parameter(Mandatory)][string]$Accent)
  $label = [System.Net.WebUtility]::HtmlEncode(([string]$Row.Product).Replace('Microsoft 365 ',''))
  $percent = if ([long]$Row.Enabled -gt 0) {
    (([decimal]$Row.Consumed * 100 / [decimal]$Row.Enabled).ToString('0.#', [Globalization.CultureInfo]::InvariantCulture) + '%')
  } else { 'N/A' }
  $status = if ($Row.Subscribed) { 'License utilization' } else { 'Not subscribed' }
  return '<td width="{0}%" style="width:{0}%;padding:5px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:#ffffff;border:1px solid #dce6ed;border-top:4px solid {1};"><tr><td style="padding:12px 10px;"><div style="font-size:12px;line-height:17px;font-weight:700;color:#334155;">{2}</div><div style="margin-top:6px;font-size:26px;line-height:30px;font-weight:700;color:{1};">{3}</div><div style="font-size:10px;line-height:14px;color:#64748b;">{4}</div><div style="margin-top:8px;font-size:12px;line-height:17px;color:#334155;"><strong>{5}</strong> used of <strong>{6}</strong> enabled</div></td></tr></table></td>' -f `
    $Width,$Accent,$label,$percent,$status,$Row.Consumed,$Row.Enabled
}

function ConvertTo-LicensesActivityDate {
  param([AllowNull()]$Value)
  $valueText = ([string]$Value).Trim()
  if (-not $valueText) { return $null }
  $adDate = [datetime]::MinValue
  if ([datetime]::TryParseExact($valueText, [string[]]@('dd/MM/yyyy HH:mm:ss','dd/MM/yyyy H:mm:ss'),
      [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$adDate)) {
    return $adDate.Date
  }
  $parsed = [datetimeoffset]::MinValue
  if ([datetimeoffset]::TryParse($valueText, [Globalization.CultureInfo]::InvariantCulture,
      [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
    return $parsed.UtcDateTime.Date
  }
  return $null
}

function ConvertTo-LicensesMailboxSizeGb {
  param([AllowNull()]$Value)
  $sizeText = ([string]$Value).Trim()
  if ($sizeText -notmatch '^\d+(?:[.,]\d+)?$') { return $null }
  $size = [decimal]0
  if ([decimal]::TryParse($sizeText.Replace(',','.'), [Globalization.NumberStyles]::AllowDecimalPoint,
      [Globalization.CultureInfo]::InvariantCulture, [ref]$size)) { return $size }
  return $null
}

function Get-LicensesQualifiedPartialAdCoverage {
  param([Parameter(Mandatory)]$Manifest, [Parameter(Mandatory)]$FileRecord)
  $adScope = 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'
  if ($Manifest.Producer -cne 'SmartM365-ActiveDirectory-Inventory.ps1' -or
      $Manifest.Status -cne 'Completed' -or $Manifest.Scope -cne $adScope -or
      $Manifest.ScopeQualification -cne 'ConsumerScope' -or $Manifest.Errors -ne 0 -or
      $Manifest.ConsumerScopeQualified -isnot [bool] -or -not $Manifest.ConsumerScopeQualified -or
      $Manifest.IsPartialInventory -isnot [bool] -or -not $Manifest.IsPartialInventory -or
      $Manifest.FullInventoryQualified -isnot [bool] -or $Manifest.FullInventoryQualified -or
      $FileRecord.File -cne 'AD_Users_AllDomains.csv' -or $FileRecord.Status -cne 'Success' -or
      $FileRecord.Errors -ne 0 -or $FileRecord.Scope -cne $adScope -or
      $FileRecord.IsPartialInventory -isnot [bool] -or -not $FileRecord.IsPartialInventory -or
      @($Manifest.RequiredFiles) -notcontains 'AD_Users_AllDomains.csv') {
    throw 'Partial AD producer/file identity, status or consumer scope is not qualified.'
  }
  $sets = @{}
  foreach ($coverage in @($Manifest.DomainCoverage, $FileRecord.DomainCoverage)) {
    if ($coverage.Kind -cne 'ADDomainCoverage' -or $coverage.Status -cne 'PartialAccepted' -or
        $coverage.Reason -cne 'NonBlockingDomainErrors') { throw 'Accepted AD domain coverage is missing.' }
    $current = @{}
    foreach ($field in @('ExpectedDomains','CollectedDomains','UnavailableDomains','NonBlockingDomainErrors')) {
      $set = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
      foreach ($domain in @($coverage.$field)) {
        if ($domain -isnot [string] -or $domain -notmatch '^[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?$' -or
            -not $set.Add($domain)) { throw "Invalid or duplicate AD coverage domain: $field" }
      }
      $current[$field] = $set
    }
    if (-not $current.ExpectedDomains.Count -or -not $current.CollectedDomains.Count -or
        -not $current.UnavailableDomains.Count -or
        $current.CollectedDomains.Overlaps($current.UnavailableDomains) -or
        -not $current.UnavailableDomains.IsSubsetOf($current.NonBlockingDomainErrors)) {
      throw 'Partial AD domains are empty, overlapping or not explicitly tolerated.'
    }
    $union = [Collections.Generic.HashSet[string]]::new($current.CollectedDomains, [StringComparer]::OrdinalIgnoreCase)
    $union.UnionWith($current.UnavailableDomains)
    if (-not $union.SetEquals($current.ExpectedDomains)) { throw 'AD coverage does not account for every expected domain.' }
    if ($sets.Count) {
      foreach ($field in $current.Keys) {
        if (-not $sets[$field].SetEquals($current[$field])) { throw 'AD file and producer domain coverage differ.' }
      }
    }
    else { $sets = $current }
  }
  return $Manifest.DomainCoverage
}

# Write or read target of a DATA-LAST file of this run: its staged copy while Core staging is active.
function Get-LicensesPublicationPath {
  param([Parameter(Mandatory)][string]$Path,[switch]$Artifact)
  if (-not (Get-Command -Name Resolve-SmartM365SourcePublicationPath -ErrorAction SilentlyContinue)) { return $Path }
  return (Resolve-SmartM365SourcePublicationPath -Path $Path -Artifact:$Artifact)
}

function Get-LicensesCsvSource {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$FileName,
    [Parameter(Mandatory)][string[]]$Columns,
    [Parameter(Mandatory)][datetime]$AsOfUtc,
    [string]$CollectorManifestName = '',
    [switch]$RequireFileReceipt,
    [switch]$RequireConsumerScope,
    [switch]$AllowQualifiedPartialAdScope,
    [AllowNull()]$CurrentRunPublication
  )
  $path = Join-Path -Path $Folder -ChildPath $FileName
  $canonicalPath = [IO.Path]::GetFullPath($path)
  $currentRunSource = $CurrentRunPublication -and $CollectorManifestName -ceq 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  if ($currentRunSource) {
    # A staged run has not promoted its CSVs yet: read the copy that its completion will promote.
    $stagedPath = Get-LicensesPublicationPath -Path $path
    if ($stagedPath -cne $path -and (Test-Path -LiteralPath $stagedPath -PathType Leaf)) { $path = $stagedPath }
  }
  $source = [pscustomobject]@{ Name=$FileName; Path=$path; Ready=$false; Reason=''; Date=''; ModifiedUtc=$null; Provenance='CSV only'; Forced=$false; PartialScopeAccepted=$false; DomainCoverage=$null; SHA256=''; ReceiptPath=''; ReceiptHash=''; Proof=$null }
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $source.Reason='missing'; return $source }
  $item = Get-Item -LiteralPath $path
  $source.ModifiedUtc = $item.LastWriteTimeUtc
  $source.Date = $item.LastWriteTimeUtc.ToString('yyyy-MM-dd')
  $source.SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
  if ($item.LastWriteTimeUtc.Date -lt $AsOfUtc.Date.AddDays(-14)) { $source.Reason='older than 14 days'; return $source }
  $header = Get-Content -LiteralPath $path -TotalCount 1 -ErrorAction Stop
  $foundColumns = @($header.TrimStart([char]0xFEFF).Replace('"','').Split(','))
  foreach ($column in $Columns) {
    if ($column -notin $foundColumns) { $source.Reason="missing column: $column"; return $source }
  }
  # During the licensing run its own receipt is still Running: prove the file by this run's publication instead.
  if ($currentRunSource) {
    if (-not $CurrentRunPublication.Hashes.ContainsKey($canonicalPath)) { $source.Reason = 'file was not published by the current licensing run'; return $source }
    if ([string]$CurrentRunPublication.Hashes[$canonicalPath] -cne $source.SHA256) { $source.Reason = 'file hash differs from the current-run publication'; return $source }
    $source.Provenance = 'current run publication'
    $source.Proof = [pscustomobject]@{ RunId = [string]$CurrentRunPublication.RunId }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $source.SHA256) { $source.Reason='source changed during validation'; return $source }
    $source.Ready = $true
    return $source
  }
  if ($CollectorManifestName) {
    $manifestPath = Join-Path -Path $Folder -ChildPath $CollectorManifestName
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
      try {
        $manifestHash=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Assert-SmartM365SourcePublication -ReceiptPath $manifestPath -Proof $manifest
        $partialAdCoverage = $null
        if ($AllowQualifiedPartialAdScope -and $FileName -ceq 'AD_Users_AllDomains.csv' -and [bool]$manifest.IsPartialInventory) {
          try {
            $adFile = @($manifest.Files | Where-Object { $_.File -ceq $FileName })
            if ($adFile.Count -ne 1) { throw 'Exactly one AD file receipt is required.' }
            $partialAdCoverage = Get-LicensesQualifiedPartialAdCoverage -Manifest $manifest -FileRecord $adFile[0]
            if ($foundColumns -notcontains 'DomainName') { throw 'Partial AD CSV requires DomainName.' }
          }
          catch { $source.Reason = "collector receipt partial AD scope is not qualified: $($_.Exception.Message)"; return $source }
        }
        if ([string]$manifest.Status -ne 'Completed' -or ([bool]$manifest.IsPartialInventory -and -not $partialAdCoverage)) {
          $source.Reason = "collector receipt is $($manifest.Status) or partial"
          return $source
        }
        if ($RequireConsumerScope -and -not [bool]$manifest.ConsumerScopeQualified) {
          $source.Reason = 'collector receipt does not qualify the consumer scope'
          return $source
        }
        $fileReceipts = @($manifest.Files | Where-Object { [string]$_.File -eq $FileName })
        if (($RequireFileReceipt -or $manifest.PSObject.Properties['PublicationProtocol']) -and $fileReceipts.Count -eq 0) {
          $source.Reason = 'file missing from collector receipt'
          return $source
        }
        if ($fileReceipts.Count -gt 1 -or ($fileReceipts.Count -eq 1 -and
            ([string]$fileReceipts[0].Status -ne 'Success' -or ([bool]$fileReceipts[0].IsPartialInventory -and -not $partialAdCoverage)))) {
          $source.Reason = 'file receipt is not a single successful complete export'
          return $source
        }
        if ($RequireFileReceipt -or $partialAdCoverage -or $manifest.PSObject.Properties['PublicationProtocol']) {
          if (-not [string]$fileReceipts[0].SHA256 -or
              (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne [string]$fileReceipts[0].SHA256) {
            $source.Reason = 'file hash differs from collector receipt'
            return $source
          }
        }
        $source.Provenance = 'collector completed'
        if ($partialAdCoverage) {
          $source.PartialScopeAccepted = $true
          $source.DomainCoverage = $partialAdCoverage
          $source.Reason = 'Qualified partial AD scope: {0}/{1} domains collected; tolerated unavailable domains: {2}. Full inventory coverage is not qualified.' -f
            @($partialAdCoverage.CollectedDomains).Count, @($partialAdCoverage.ExpectedDomains).Count, (@($partialAdCoverage.UnavailableDomains) -join ', ')
          $source.Provenance = $source.Reason
        }
        $source.ReceiptPath=$manifestPath
        $source.ReceiptHash=$manifestHash
        if ((Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash -cne $manifestHash) { throw 'collector receipt changed during validation' }
        $source.Proof=$manifest
      }
      catch { $source.Reason='collector receipt is unreadable'; return $source }
    }
    elseif ($RequireFileReceipt) { $source.Reason='collector receipt is missing'; return $source }
  }
  if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $source.SHA256) { $source.Reason='source changed during validation'; return $source }
  $source.Ready = $true
  return $source
}

function Import-LicensesSourceCsv {
  param([Parameter(Mandatory)]$Source)
  if ($Source.ReceiptPath) {
    Assert-SmartM365SourcePublication -ReceiptPath $Source.ReceiptPath -Proof $Source.Proof
    if ((Get-FileHash -LiteralPath $Source.ReceiptPath).Hash -cne $Source.ReceiptHash) { throw 'collector receipt changed during read' }
  }
  if ((Get-FileHash -LiteralPath $Source.Path).Hash -cne $Source.SHA256) { throw 'source changed before read' }
  $data=@(Import-Csv -LiteralPath $Source.Path -ErrorAction Stop)
  if ((Get-FileHash -LiteralPath $Source.Path).Hash -cne $Source.SHA256) { throw 'source changed during read' }
  if ($Source.ReceiptPath) {
    Assert-SmartM365SourcePublication -ReceiptPath $Source.ReceiptPath -Proof $Source.Proof
    if ((Get-FileHash -LiteralPath $Source.ReceiptPath).Hash -cne $Source.ReceiptHash) { throw 'collector receipt changed during read' }
  }
  return $data
}

function Read-LicensesIndexedSource {
  param(
    [Parameter(Mandatory)]$Source,
    [Parameter(Mandatory)][string]$ExpectedTenantKey,
    [Parameter(Mandatory)][string]$KeyColumn,
    [Parameter(Mandatory)][System.Collections.IDictionary]$WantedKeys,
    [Parameter(Mandatory)][datetime]$AsOfUtc,
    [string]$RefreshColumn,
    [string]$PeriodColumn,
    [string]$ExpectedPeriod
  )
  $index = @{}
  $duplicates = @{}
  $refreshDate = $null
  if (-not $Source.Ready) { return [pscustomobject]@{ Ready=$false; Rows=$index; Duplicates=$duplicates; Source=$Source } }
  try {
    foreach ($row in (Import-LicensesSourceCsv -Source $Source)) {
      if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
      if ($PeriodColumn -and ([string]$row.$PeriodColumn).Trim() -ne $ExpectedPeriod) { throw 'unexpected report period' }
      if ($RefreshColumn) {
        $rowDate = ConvertTo-LicensesActivityDate $row.$RefreshColumn
        if ($null -eq $rowDate -or $rowDate -lt $AsOfUtc.Date.AddDays(-14)) { throw 'report refresh date is missing or older than 14 days' }
        if ($null -eq $refreshDate) { $refreshDate = $rowDate }
        elseif ($refreshDate -ne $rowDate) { throw 'mixed report refresh dates' }
      }
      $key = ([string]$row.$KeyColumn).Trim().ToLowerInvariant()
      if (-not $key -or -not $WantedKeys.Contains($key)) { continue }
      if ($index.ContainsKey($key)) { $duplicates[$key] = $true; continue }
      $index[$key] = $row
    }
    if ($RefreshColumn -and $null -eq $refreshDate) { throw 'empty report' }
    if ($refreshDate) { $Source.Date = $refreshDate.ToString('yyyy-MM-dd') }
  }
  catch {
    $Source.Ready = $false
    $Source.Reason = $_.Exception.Message
    $index = @{}
    $duplicates = @{}
  }
  return [pscustomobject]@{ Ready=$Source.Ready; Rows=$index; Duplicates=$duplicates; Source=$Source }
}

function Get-LicensesMailboxGapSummary {
  param(
    [Parameter(Mandatory)]$MailboxSource,
    [Parameter(Mandatory)]$OnPremSource,
    [Parameter(Mandatory)]$ActiveSource,
    [Parameter(Mandatory)][string]$ExpectedTenantKey,
    [Parameter(Mandatory)][System.Collections.IDictionary]$TargetSuitesByUser,
    [Parameter(Mandatory)][System.Collections.IDictionary]$AllSkusByUser
  )

  $userMailboxes = [pscustomobject]@{ Available=$false; Universe=0; Total=0; OtherSkus=0; NoSkus=0; Unknown=0; EntraEnabled=0; EntraDisabled=0; EntraStateUnknown=0; EntraStateAvailable=$false; EntraStateReason=''; Reason='' }
  $noUserMailbox = [pscustomobject]@{ Available=$false; Universe=0; Total=0; OtherSkus=0; NoSkus=0; MemberEnabled=0; MemberDisabled=0; Guests=0; Unknown=0; Members=@{}; Reason='' }
  if (-not $MailboxSource.Ready) {
    $userMailboxes.Reason = [string]$MailboxSource.Reason
    $noUserMailbox.Reason = [string]$MailboxSource.Reason
    return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox }
  }

  $mailboxTypes = @{}
  $duplicateMailboxIds = @{}
  $unqualifiedUserMailboxIds = @{}
  $userMailboxCandidateIds = @{}
  $unjoinedMailboxRows = 0
  try {
    foreach ($row in (Import-LicensesSourceCsv -Source $MailboxSource)) {
      if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
      $id = ([string]$row.ExternalDirectoryObjectId).Trim().ToLowerInvariant()
      $kind = ([string]$row.RecipientTypeDetails).Trim()
      $observed = ([string]$row.NativeIdentityStatus).Trim() -eq 'Observed'
      if (-not $id) {
        $unjoinedMailboxRows++
        if ($kind -eq 'UserMailbox') { $userMailboxes.Unknown++ }
        continue
      }
      if ($mailboxTypes.ContainsKey($id)) {
        $duplicateMailboxIds[$id] = $true
        if ($kind -eq 'UserMailbox' -or $mailboxTypes[$id] -eq 'UserMailbox') { $unqualifiedUserMailboxIds[$id] = $true }
        continue
      }
      $mailboxTypes[$id] = if ($observed -and $kind) { $kind } else { 'Unknown' }
      if ((-not $observed -or -not $kind) -and $kind -eq 'UserMailbox' -and
          -not $TargetSuitesByUser.ContainsKey($id)) { $unqualifiedUserMailboxIds[$id] = $true }
    }
    foreach ($id in $mailboxTypes.Keys) {
      if ($duplicateMailboxIds.ContainsKey($id) -or $mailboxTypes[$id] -eq 'Unknown') {
        if (-not $TargetSuitesByUser.ContainsKey($id) -and $unqualifiedUserMailboxIds.ContainsKey($id)) { $userMailboxes.Unknown++ }
        continue
      }
      if ($mailboxTypes[$id] -ne 'UserMailbox') { continue }
      $userMailboxes.Universe++
      if ($TargetSuitesByUser.ContainsKey($id)) { continue }
      $userMailboxes.Total++
      $userMailboxCandidateIds[$id] = $true
      if ($AllSkusByUser.ContainsKey($id)) { $userMailboxes.OtherSkus++ }
      else { $userMailboxes.NoSkus++ }
    }
    $userMailboxes.Available = $true
  }
  catch { $userMailboxes.Reason=$_.Exception.Message; $noUserMailbox.Reason=$_.Exception.Message; return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox } }

  $accounts = @{}
  $duplicateAccounts = @{}
  $activeReadError = ''
  if ($ActiveSource.Ready) {
    try {
      $accountRows = 0
      foreach ($row in (Import-LicensesSourceCsv -Source $ActiveSource)) {
        $accountRows++
        if (-not $row.PSObject.Properties['UserType']) { throw 'missing UserType column' }
        if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        $id = ([string]$row.'Object Id').Trim().ToLowerInvariant()
        if (-not $id) { $noUserMailbox.Unknown++; continue }
        if ($accounts.ContainsKey($id)) { $duplicateAccounts[$id] = $true; continue }
        $accounts[$id] = $row
      }
      if ($accountRows -eq 0) { throw 'empty Entra user export' }
      foreach ($id in $userMailboxCandidateIds.Keys) {
        if (-not $accounts.ContainsKey($id) -or $duplicateAccounts.ContainsKey($id)) {
          $userMailboxes.EntraStateUnknown++
          continue
        }
        $enabled = ([string]$accounts[$id].AccountEnabled).Trim().ToLowerInvariant()
        if ($enabled -eq 'true') { $userMailboxes.EntraEnabled++ }
        elseif ($enabled -eq 'false') { $userMailboxes.EntraDisabled++ }
        else { $userMailboxes.EntraStateUnknown++ }
      }
      if ($userMailboxes.EntraEnabled + $userMailboxes.EntraDisabled + $userMailboxes.EntraStateUnknown -ne $userMailboxes.Total) {
        throw 'UserMailbox Entra states do not reconcile with the unlicensed mailbox total.'
      }
      $userMailboxes.EntraStateAvailable = $true
    }
    catch { $activeReadError = $_.Exception.Message; $userMailboxes.EntraStateReason = $activeReadError }
  }
  else { $userMailboxes.EntraStateReason = [string]$ActiveSource.Reason }

  if (-not $ActiveSource.Ready -or $activeReadError) {
    $noUserMailbox.Reason = if ($activeReadError) { $activeReadError } else { [string]$ActiveSource.Reason }
  }
  elseif (-not $OnPremSource.Ready) {
    $noUserMailbox.Reason = "on-premises UserMailbox source: $($OnPremSource.Reason)"
  }
  elseif ($unjoinedMailboxRows -gt 0) {
    $noUserMailbox.Reason = "$unjoinedMailboxRows mailbox rows have no Entra ID; absence of a mailbox cannot be established"
  }
  else {
    try {
      $byImmutable = @{}
      $byUpn = @{}
      $duplicateImmutable = @{}
      $duplicateUpn = @{}
      foreach ($id in $accounts.Keys) {
        $account = $accounts[$id]
        $immutable = ([string]$account.OnPremisesImmutableId).Trim().ToLowerInvariant()
        $upn = ([string]$account.'User principal name').Trim().ToLowerInvariant()
        if ($immutable) {
          if ($byImmutable.ContainsKey($immutable)) { $duplicateImmutable[$immutable] = $true }
          else { $byImmutable[$immutable] = $id }
        }
        if ($upn) {
          if ($byUpn.ContainsKey($upn)) { $duplicateUpn[$upn] = $true }
          else { $byUpn[$upn] = $id }
        }
      }
      $onPremTypes = @{}
      $onPremAmbiguousAccounts = @{}
      $onPremConflicts = 0
      $onPremTypeConflicts = 0
      $onPremUnmatched = 0
      foreach ($row in (Import-LicensesSourceCsv -Source $OnPremSource)) {
        if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present in on-premises mailboxes' }
        if (([string]$row.NativeIdentityStatus).Trim() -ne 'Observed') { throw 'on-premises mailbox identity is not observed' }
        $kind = ([string]$row.RecipientType).Trim()
        if (-not $kind) { throw 'on-premises mailbox type is missing' }
        $guidValue = [guid]::Empty
        if (-not [guid]::TryParse(([string]$row.ObjectGUID).Trim(), [ref]$guidValue)) { throw 'on-premises mailbox ObjectGUID is missing or invalid' }
        $immutable = [Convert]::ToBase64String($guidValue.ToByteArray()).ToLowerInvariant()
        $upn = ([string]$row.UserPrincipalName).Trim().ToLowerInvariant()
        $immutableId = if ($byImmutable.ContainsKey($immutable)) { $byImmutable[$immutable] } else { '' }
        $upnId = if ($upn -and $byUpn.ContainsKey($upn)) { $byUpn[$upn] } else { '' }
        if ($duplicateImmutable.ContainsKey($immutable) -or ($upnId -and $duplicateUpn.ContainsKey($upn))) {
          throw 'ambiguous on-premises mailbox identity match'
        }
        if ($immutableId -and $upnId -and $immutableId -ne $upnId) {
          $onPremAmbiguousAccounts[$immutableId] = $true
          $onPremAmbiguousAccounts[$upnId] = $true
          $onPremConflicts++
          continue
        }
        $id = if ($immutableId) { $immutableId } else { $upnId }
        if (-not $id) { $onPremUnmatched++; continue }
        if ($onPremTypes.ContainsKey($id)) { throw 'multiple on-premises mailboxes match one Entra account' }
        if ($mailboxTypes.ContainsKey($id) -and $mailboxTypes[$id] -ne $kind) {
          $onPremAmbiguousAccounts[$id] = $true
          $onPremTypeConflicts++
          continue
        }
        $onPremTypes[$id] = $kind
      }
      if ($onPremConflicts -gt 0 -or $onPremTypeConflicts -gt 0 -or $onPremUnmatched -gt 0) {
        $OnPremSource.Provenance = "collector completed; $onPremConflicts identity conflicts, $onPremTypeConflicts EXO/on-premises type conflicts and $onPremUnmatched unmatched mailbox rows"
      }
      foreach ($id in $accounts.Keys) {
        if ($duplicateAccounts.ContainsKey($id) -or $onPremAmbiguousAccounts.ContainsKey($id) -or $duplicateMailboxIds.ContainsKey($id) -or
            ($mailboxTypes.ContainsKey($id) -and $mailboxTypes[$id] -eq 'Unknown')) {
          if (-not $TargetSuitesByUser.ContainsKey($id)) { $noUserMailbox.Unknown++ }
          continue
        }
        if (($mailboxTypes.ContainsKey($id) -and $mailboxTypes[$id] -ne 'UserMailbox') -or
            ($onPremTypes.ContainsKey($id) -and $onPremTypes[$id] -ne 'UserMailbox')) { continue }
        $noUserMailbox.Universe++
        if ($TargetSuitesByUser.ContainsKey($id)) { continue }
        if ($mailboxTypes.ContainsKey($id) -or $onPremTypes.ContainsKey($id)) { continue }
        $userType = ([string]$accounts[$id].UserType).Trim()
        $enabled = ([string]$accounts[$id].AccountEnabled).Trim().ToLowerInvariant()
        if ($userType -eq 'Member' -and $enabled -notin @('true','false')) { $noUserMailbox.Unknown++; continue }
        if ($userType -notin @('Member','Guest')) { $noUserMailbox.Unknown++; continue }
        $noUserMailbox.Total++
        if ($AllSkusByUser.ContainsKey($id)) { $noUserMailbox.OtherSkus++ }
        else { $noUserMailbox.NoSkus++ }
        if ($userType -eq 'Guest') { $noUserMailbox.Guests++ }
        elseif ($enabled -eq 'true') { $noUserMailbox.MemberEnabled++; $noUserMailbox.Members[$id] = $accounts[$id] }
        else { $noUserMailbox.MemberDisabled++; $noUserMailbox.Members[$id] = $accounts[$id] }
      }
      $noUserMailbox.Available = $true
    }
    catch { $noUserMailbox.Reason=$_.Exception.Message }
  }
  return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox }
}

function Get-LicensesAdAccountActivitySummary {
  param(
    [Parameter(Mandatory)]$MailboxGap,
    [Parameter(Mandatory)]$AdSource,
    [Parameter(Mandatory)][System.Collections.IDictionary]$AdByImmutable,
    [Parameter(Mandatory)][System.Collections.IDictionary]$AdByUpn,
    [Parameter(Mandatory)][System.Collections.IDictionary]$DuplicateAdUpns,
    [Parameter(Mandatory)][datetime]$Cutoff
  )
  $result = [pscustomobject]@{
    Available=$false; Provisional=([bool]$AdSource.Forced -or ($AdSource.PSObject.Properties['PartialScopeAccepted'] -and $AdSource.PartialScopeAccepted)); Members=0; AdObserved=0
    AdEnabledRecent=0; AdEnabledInactive=0; AdEnabledNoDate=0; AdEnabledUnknown=0
    AdDisabled=0; Ambiguous=0; NoObservedMatch=0; UnmatchedScopeUnknown=0; Reason=''
  }
  if (-not $MailboxGap.NoUserMailbox.Available) {
    $result.Reason = [string]$MailboxGap.NoUserMailbox.Reason
    return $result
  }
  if (-not $AdSource.Ready) {
    $result.Reason = [string]$AdSource.Reason
    return $result
  }
  foreach ($account in $MailboxGap.NoUserMailbox.Members.Values) {
    $result.Members++
    $immutable = ([string]$account.OnPremisesImmutableId).Trim().ToLowerInvariant()
    $upn = ([string]$account.'User principal name').Trim().ToLowerInvariant()
    $byImmutable = if ($immutable -and $AdByImmutable.ContainsKey($immutable)) { $AdByImmutable[$immutable] } else { $null }
    $upnDuplicated = ($upn -and $DuplicateAdUpns.ContainsKey($upn))
    $byUpn = if ($upn -and -not $upnDuplicated -and $AdByUpn.ContainsKey($upn)) { $AdByUpn[$upn] } else { $null }
    if (($byImmutable -and $byUpn -and -not [object]::ReferenceEquals($byImmutable,$byUpn)) -or ($upnDuplicated -and -not $byImmutable)) {
      $result.Ambiguous++
      continue
    }
    $ad = if ($byImmutable) { $byImmutable } else { $byUpn }
    if (-not $ad) {
      if ($AdSource.PSObject.Properties['PartialScopeAccepted'] -and $AdSource.PartialScopeAccepted) { $result.UnmatchedScopeUnknown++ }
      else { $result.NoObservedMatch++ }
      continue
    }
    $result.AdObserved++
    $enabled = ([string]$ad.Enabled).Trim().ToLowerInvariant()
    if ($enabled -eq 'false') { $result.AdDisabled++; continue }
    if ($enabled -ne 'true') { $result.AdEnabledUnknown++; continue }
    $lastLogon = ConvertTo-LicensesActivityDate $ad.LastLogonDate
    if (-not $lastLogon) { $result.AdEnabledNoDate++ }
    elseif ($lastLogon -ge $Cutoff) { $result.AdEnabledRecent++ }
    else { $result.AdEnabledInactive++ }
  }
  if ($result.AdObserved + $result.Ambiguous + $result.NoObservedMatch + $result.UnmatchedScopeUnknown -ne $result.Members -or
      $result.AdEnabledRecent + $result.AdEnabledInactive + $result.AdEnabledNoDate + $result.AdEnabledUnknown + $result.AdDisabled -ne $result.AdObserved) {
    throw 'AD account activity categories do not reconcile with section 07 members.'
  }
  $result.Available = $true
  if ($AdSource.PSObject.Properties['PartialScopeAccepted'] -and $AdSource.PartialScopeAccepted) {
    $result.NoObservedMatch = $null
    $result.Reason = $AdSource.Reason + " $($result.UnmatchedScopeUnknown) unmatched section 07 member(s) remain unassessed; their AD domain cannot be inferred from their UPN. Members retains the Entra section 07 population, not an exhaustive AD population."
  }
  return $result
}

function Get-LicensesFocusedUsageRows {
  param(
    [Parameter(Mandatory)][string]$CsvFolderPath,
    [Parameter(Mandatory)][string]$ExpectedTenantKey,
    [datetimeoffset]$LicenseSnapshotUtc = [datetimeoffset]::MinValue,
    [datetime]$AsOfUtc = [datetime]::UtcNow,
    [switch]$ForceAdCsvAnalysis,
    [AllowNull()]$CurrentRunPublication
  )
  $products = @(
    @{ Name='Microsoft 365 F1'; PartNumbers=@('M365_F1','M365_F1_COMM') }
    @{ Name='Microsoft 365 F3'; PartNumbers=@('SPE_F1') }
    @{ Name='Microsoft 365 E3'; PartNumbers=@('SPE_E3') }
    @{ Name='Microsoft 365 E5'; PartNumbers=@('SPE_E5') }
  )
  $sources = [System.Collections.Generic.List[object]]::new()
  $licenseSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'M365_Licenses_Users.csv' -Columns @('TenantKey','UserId','SkuPartNumber') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt' -RequireFileReceipt -CurrentRunPublication $CurrentRunPublication
  if ($licenseSource.Ready -and $LicenseSnapshotUtc -ne [datetimeoffset]::MinValue -and
      [math]::Abs(($licenseSource.ModifiedUtc - $LicenseSnapshotUtc.UtcDateTime).TotalHours) -gt 24) {
    $licenseSource.Ready = $false
    $licenseSource.Reason = 'license user assignments do not match the tenant snapshot date (over 24 hours apart)'
  }
  $sources.Add($licenseSource)
  $productUsers = @{}
  foreach ($product in $products) { $productUsers[$product.Name] = @{} }
  $targetSuitesByUser = @{}
  $allSkusByUser = @{}
  $licenseIdentityByUser = @{}
  if ($licenseSource.Ready) {
    try {
      foreach ($row in (Import-LicensesSourceCsv -Source $licenseSource)) {
        if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        $rowUserId = ([string]$row.UserId).Trim().ToLowerInvariant()
        $rowSku = ([string]$row.SkuPartNumber).Trim().ToUpperInvariant()
        if ($rowUserId -and $rowSku) {
          if (-not $allSkusByUser.ContainsKey($rowUserId)) { $allSkusByUser[$rowUserId] = @{} }
          $allSkusByUser[$rowUserId][$rowSku] = $true
          if (-not $licenseIdentityByUser.ContainsKey($rowUserId)) {
            $licenseIdentityByUser[$rowUserId] = [pscustomobject]@{
              UserPrincipalName = if ($row.PSObject.Properties['User principal name']) { [string]$row.'User principal name' } else { '' }
              DisplayName = if ($row.PSObject.Properties['Display name']) { [string]$row.'Display name' } else { '' }
            }
          }
        }
        $rowProduct = Get-LicensesSuiteProduct -SkuId (Get-LicensesOptionalProperty $row 'SkuId') -SkuPartNumber $row.SkuPartNumber
        foreach ($product in $products) {
          if ($rowProduct -ne $product.Name) { continue }
          $userId = $rowUserId
          if (-not $userId) { throw 'a target license row has no UserId' }
          $productUsers[$product.Name][$userId] = $true
          if (-not $targetSuitesByUser.ContainsKey($userId)) { $targetSuitesByUser[$userId] = @{} }
          $targetSuitesByUser[$userId][$product.Name] = $true
          break
        }
      }
    }
    catch { $licenseSource.Ready=$false; $licenseSource.Reason=$_.Exception.Message }
  }
  if (-not $licenseSource.Ready) {
    $unavailableRows = foreach ($product in $products) {
      [pscustomobject]@{ Product=$product.Name; Counts=@{}; Available=$false }
    }
    return [pscustomobject]@{ Rows=@($unavailableRows); RecoveryDetails=@(); DowngradeReview=$null; DowngradeDetails=@(); Sources=$sources.ToArray(); SharedSourceReady=$false; IntuneSourceReady=$false; AdSourceForced=$false; MailboxGap=$null; AdGapActivity=$null }
  }

  $sharedSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'Exchange_EXO_Mailboxes_AllDomains.csv' -Columns @(
    'TenantKey','ExternalDirectoryObjectId','RecipientTypeDetails','NativeIdentityStatus','TotalItemSizeGB',
    'ArchiveStatus','LitigationHoldEnabled','RetentionHoldEnabled'
  ) -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
  $sources.Add($sharedSource)
  $mailboxes = Read-LicensesIndexedSource -Source $sharedSource -ExpectedTenantKey $ExpectedTenantKey -KeyColumn 'ExternalDirectoryObjectId' -WantedKeys $targetSuitesByUser -AsOfUtc $AsOfUtc

  $activeSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'M365_Users_Active.csv' -Columns @('TenantKey','Object Id','User principal name','AccountEnabled','OnPremisesImmutableId','LastSuccessfulSignInDateTime') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-ActiveUsers-Inventory.current.json.txt'
  $sources.Add($activeSource)
  $active = Read-LicensesIndexedSource -Source $activeSource -ExpectedTenantKey $ExpectedTenantKey -KeyColumn 'Object Id' -WantedKeys $targetSuitesByUser -AsOfUtc $AsOfUtc
  $onPremSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'Exchange_OnPrem_Mailboxes_AllDomains.csv' -Columns @('TenantKey','ObjectGUID','UserPrincipalName','RecipientType','NativeIdentityStatus') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-Exchange-Local-Mailboxes-Inventory.current.json.txt' -RequireFileReceipt -RequireConsumerScope
  $sources.Add($onPremSource)
  $mailboxGap = Get-LicensesMailboxGapSummary -MailboxSource $sharedSource -OnPremSource $onPremSource -ActiveSource $activeSource -ExpectedTenantKey $ExpectedTenantKey -TargetSuitesByUser $targetSuitesByUser -AllSkusByUser $allSkusByUser
  $wantedUpns = @{}
  $wantedImmutableIds = @{}
  if ($active.Ready) {
    foreach ($user in $active.Rows.Values) {
      $upn = ([string]$user.'User principal name').Trim().ToLowerInvariant()
      if ($upn) { $wantedUpns[$upn] = $true }
      $immutableId = ([string]$user.OnPremisesImmutableId).Trim().ToLowerInvariant()
      if ($immutableId) { $wantedImmutableIds[$immutableId] = $true }
    }
  }
  if ($mailboxGap.NoUserMailbox.Available) {
    foreach ($user in $mailboxGap.NoUserMailbox.Members.Values) {
      $upn = ([string]$user.'User principal name').Trim().ToLowerInvariant()
      if ($upn) { $wantedUpns[$upn] = $true }
      $immutableId = ([string]$user.OnPremisesImmutableId).Trim().ToLowerInvariant()
      if ($immutableId) { $wantedImmutableIds[$immutableId] = $true }
    }
  }

  $intuneSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'Intune_Devices_Inventory.csv' -Columns @(
    'TenantKey','Device ID','OS','UserId','Primary user UPN'
  ) -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
  $sources.Add($intuneSource)
  $intunePrimaryPcUsers = @{}
  if ($intuneSource.Ready) {
    try {
      $deviceOwners = @{}
      foreach ($device in (Import-LicensesSourceCsv -Source $intuneSource)) {
        if ([string]$device.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        if (([string]$device.OS).Trim() -ine 'Windows') { continue }
        $deviceId = ([string]$device.'Device ID').Trim().ToLowerInvariant()
        if (-not $deviceId) { throw 'a Windows Intune device has no Device ID' }
        $primaryUserId = ([string]$device.UserId).Trim().ToLowerInvariant()
        if ($deviceOwners.ContainsKey($deviceId)) {
          if ($deviceOwners[$deviceId] -ne $primaryUserId) { throw 'conflicting primary users for an Intune device' }
          continue
        }
        $deviceOwners[$deviceId] = $primaryUserId
        if ($primaryUserId -and $targetSuitesByUser.ContainsKey($primaryUserId)) {
          $intunePrimaryPcUsers[$primaryUserId] = $true
        }
      }
    }
    catch {
      $intuneSource.Ready = $false
      $intuneSource.Reason = $_.Exception.Message
      $intunePrimaryPcUsers = @{}
    }
  }

  $adSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD','UserPrincipalName','Enabled','LastLogonDate') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt' -AllowQualifiedPartialAdScope
  $adReceiptReason = ''
  if ($ForceAdCsvAnalysis -and -not $adSource.Ready -and $adSource.Reason -match '^(collector receipt|file receipt)') {
    $adReceiptReason = [string]$adSource.Reason
    $adSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD','UserPrincipalName','Enabled','LastLogonDate') -AsOfUtc $AsOfUtc
  }
  $sources.Add($adSource)
  $adByImmutable = @{}
  $adByUpn = @{}
  $adDuplicateUpns = @{}
  if ($adSource.Ready) {
    try {
      Import-LicensesSourceCsv -Source $adSource | ForEach-Object {
        if ([string]$_.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        if ($adSource.PartialScopeAccepted -and ([string]$_.DomainName).Trim() -notin @($adSource.DomainCoverage.CollectedDomains)) {
          throw 'AD CSV row is outside the qualified collected domains.'
        }
        $immutableId = ([string]$_.ImmutableId_AD).Trim().ToLowerInvariant()
        $upn = ([string]$_.UserPrincipalName).Trim().ToLowerInvariant()
        $adValue = [pscustomobject]@{ LastLogonDate=$_.LastLogonDate; Enabled=if ($_.PSObject.Properties['Enabled']) { [string]$_.Enabled } else { '' } }
        if ($immutableId -and $wantedImmutableIds.ContainsKey($immutableId)) {
          if ($adByImmutable.ContainsKey($immutableId)) { throw 'duplicate AD immutable ID' }
          $adByImmutable[$immutableId] = $adValue
        }
        if ($upn -and $wantedUpns.ContainsKey($upn)) {
          if ($adByUpn.ContainsKey($upn)) { $adDuplicateUpns[$upn] = $true }
          else { $adByUpn[$upn] = $adValue }
        }
      }
    }
    catch { $adSource.Ready=$false; $adSource.PartialScopeAccepted=$false; $adSource.Reason=$_.Exception.Message; $adByImmutable=@{}; $adByUpn=@{} }
  }
  if ($adSource.Ready -and $adReceiptReason) {
    $adSource.Forced = $true
    $adSource.Provenance = "FORCED AD CSV; ignored collector receipt: $adReceiptReason"
    WriteLog -Message ("AD CSV analysis forced despite collector receipt: {0}. AD/Entra inactivity is provisional." -f $adReceiptReason) 'WARNING'
  }
  $adGapActivity = Get-LicensesAdAccountActivitySummary -MailboxGap $mailboxGap -AdSource $adSource -AdByImmutable $adByImmutable -AdByUpn $adByUpn -DuplicateAdUpns $adDuplicateUpns -Cutoff $AsOfUtc.Date.AddDays(-90)

  $reportDefinitions = @(
    @{ Name='M365_Users_Activity.csv'; Key='UserPrincipalName'; Columns=@('TenantKey','UserPrincipalName','ReportPeriod','ReportRefreshDate','LastActivityDate','IsDeleted'); Refresh='ReportRefreshDate'; Period='ReportPeriod'; Expected='D180' }
    @{ Name='M365_Mailbox_Usage.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Is Deleted'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
    @{ Name='M365_Email_Activity.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Is Deleted','Send Count','Read Count'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
    @{ Name='M365_Apps_Usage_180D.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Windows','Mac'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
    @{ Name='M365_OneDrive_Usage.csv'; Key='Owner Principal Name'; Columns=@('TenantKey','Owner Principal Name','Is Deleted','Storage Used (Byte)','Report Refresh Date','Report Period'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
  )
  # Usage reports are used only when the completed UserActivity receipt lists them with the same hash.
  $reports = @{}
  foreach ($definition in $reportDefinitions) {
    $source = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName $definition.Name -Columns $definition.Columns -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt' -RequireFileReceipt
    $sources.Add($source)
    $reports[$definition.Name] = Read-LicensesIndexedSource -Source $source -ExpectedTenantKey $ExpectedTenantKey -KeyColumn $definition.Key -WantedKeys $wantedUpns -AsOfUtc $AsOfUtc -RefreshColumn $definition.Refresh -PeriodColumn $definition.Period -ExpectedPeriod $definition.Expected
  }
  $oneDrive = $reports['M365_OneDrive_Usage.csv']

  $cutoff = $AsOfUtc.Date.AddDays(-90)
  $recoveryDetails = [System.Collections.Generic.List[object]]::new()
  $metricRows = foreach ($product in $products) {
    $count = @{ Assigned=0; Disabled=0; DisabledUnknown=0; AdEntraInactive=0; AdEntraUnknown=0; MailboxInactive=0; MailboxUnknown=0; M365Inactive=0; M365Unknown=0; LocalAppsInactive=0; LocalAppsUnknown=0; Multiple=0; MultipleUnknown=0; MultipleAll=0; MultipleAllUnknown=0; SharedLicensed=0; SharedUnder50=0; SharedEligible=0; SharedUnknown=0; RecoveryCandidates=0; RecoveryUnknown=0; RecoveryPrimaryPc=0; RecoveryPrimaryPcUnknown=0 }
    if (-not $licenseSource.Ready) { [pscustomobject]@{ Product=$product.Name; Counts=$count; Available=$false }; continue }
    $candidateUsers = @{}
    $candidateReasons = @{}
    $unknownUsers = @{}
    foreach ($userId in $productUsers[$product.Name].Keys) {
      $count.Assigned++
      $sharedStatus = 'NotShared'
      $mailboxKind = 'Other'
      if (-not $mailboxes.Ready -or $mailboxes.Duplicates.ContainsKey($userId)) {
        $sharedStatus = 'Unknown'
        $mailboxKind = 'Unknown'
        $count.SharedUnknown++
      }
      elseif ($mailboxes.Rows.ContainsKey($userId)) {
        $mailbox = $mailboxes.Rows[$userId]
        $recipientType = ([string]$mailbox.RecipientTypeDetails).Trim()
        if (-not $recipientType -or ([string]$mailbox.NativeIdentityStatus).Trim() -ne 'Observed') {
          $sharedStatus = 'Unknown'
          $mailboxKind = 'Unknown'
          $count.SharedUnknown++
        }
        elseif ($recipientType -eq 'SharedMailbox') {
          $mailboxKind = 'Shared'
          $count.SharedLicensed++
          $sizeGb = ConvertTo-LicensesMailboxSizeGb $mailbox.TotalItemSizeGB
          $archiveStatus = ([string]$mailbox.ArchiveStatus).Trim().ToLowerInvariant()
          $litigationHold = ([string]$mailbox.LitigationHoldEnabled).Trim().ToLowerInvariant()
          $retentionHold = ([string]$mailbox.RetentionHoldEnabled).Trim().ToLowerInvariant()
          if ($null -ne $sizeGb -and $sizeGb -lt 50) { $count.SharedUnder50++ }
          if ($null -eq $sizeGb -or $archiveStatus -notin @('none','disabled','false','active','enabled','true') -or
              $litigationHold -notin @('true','false') -or $retentionHold -notin @('true','false')) {
            $sharedStatus = 'Unknown'
            $count.SharedUnknown++
          }
          elseif ($sizeGb -ge 50 -or $archiveStatus -in @('active','enabled','true') -or
                  $litigationHold -eq 'true' -or $retentionHold -eq 'true') { $sharedStatus = 'Blocked' }
          else { $sharedStatus = 'Eligible'; $count.SharedEligible++ }
        }
      }
      if ($sharedStatus -eq 'Eligible') { $candidateUsers[$userId] = $true; $candidateReasons[$userId] = 'Eligible shared mailbox under 50 GB' }
      elseif ($sharedStatus -eq 'Unknown') { $unknownUsers[$userId] = $true }
      if ($mailboxKind -eq 'Shared') { continue }
      if ($mailboxKind -eq 'Unknown') {
        $count.DisabledUnknown++; $count.AdEntraUnknown++; $count.MailboxUnknown++; $count.M365Unknown++
        $count.MultipleUnknown++; $count.MultipleAllUnknown++
        if ($product.Name -in @('Microsoft 365 E3','Microsoft 365 E5')) { $count.LocalAppsUnknown++ }
        continue
      }
      if ($targetSuitesByUser[$userId].Count -gt 1) { $count.Multiple++ }
      if ($allSkusByUser[$userId].Count -gt 1) { $count.MultipleAll++ }
      if (-not $active.Ready -or -not $active.Rows.ContainsKey($userId) -or $active.Duplicates.ContainsKey($userId)) {
        $count.DisabledUnknown++; $count.AdEntraUnknown++; $count.MailboxUnknown++; $count.M365Unknown++
        if ($sharedStatus -eq 'NotShared') { $unknownUsers[$userId] = $true }
        if ($product.Name -in @('Microsoft 365 E3','Microsoft 365 E5')) { $count.LocalAppsUnknown++ }
        continue
      }
      $user = $active.Rows[$userId]
      $enabledText = ([string]$user.AccountEnabled).Trim().ToLowerInvariant()
      if ($enabledText -eq 'false') {
        $count.Disabled++
        $candidateUsers[$userId] = $true
        $candidateReasons[$userId] = 'Disabled account'
        continue
      }
      if ($enabledText -ne 'true') {
        $count.DisabledUnknown++; $count.AdEntraUnknown++; $count.MailboxUnknown++; $count.M365Unknown++
        if ($sharedStatus -eq 'NotShared') { $unknownUsers[$userId] = $true }
        if ($product.Name -in @('Microsoft 365 E3','Microsoft 365 E5')) { $count.LocalAppsUnknown++ }
        continue
      }
      $upn = ([string]$user.'User principal name').Trim().ToLowerInvariant()
      $entraDate = ConvertTo-LicensesActivityDate $user.LastSuccessfulSignInDateTime
      $immutableId = ([string]$user.OnPremisesImmutableId).Trim().ToLowerInvariant()
      $adRow = $null
      if ($adSource.Ready -and $immutableId) {
        if ($adByImmutable.ContainsKey($immutableId)) { $adRow = $adByImmutable[$immutableId] }
        elseif ($upn -and $adByUpn.ContainsKey($upn) -and -not $adDuplicateUpns.ContainsKey($upn)) { $adRow = $adByUpn[$upn] }
      }
      $adDate = if ($adRow) { ConvertTo-LicensesActivityDate $adRow.LastLogonDate } else { $null }
      if (($entraDate -and $entraDate -ge $cutoff) -or ($adDate -and $adDate -ge $cutoff)) { }
      elseif ($entraDate -and (-not $immutableId -or $adDate)) { $count.AdEntraInactive++ }
      else { $count.AdEntraUnknown++ }

      $mail = $reports['M365_Mailbox_Usage.csv']
      $email = $reports['M365_Email_Activity.csv']
      $mailRow = if ($mail.Ready -and $upn -and $mail.Rows.ContainsKey($upn) -and -not $mail.Duplicates.ContainsKey($upn)) { $mail.Rows[$upn] } else { $null }
      $emailRow = if ($email.Ready -and $upn -and $email.Rows.ContainsKey($upn) -and -not $email.Duplicates.ContainsKey($upn)) { $email.Rows[$upn] } else { $null }
      if ($mailRow -and ([string]$mailRow.'Is Deleted').Trim().ToLowerInvariant() -eq 'true') { $mailRow = $null }
      if ($emailRow -and ([string]$emailRow.'Is Deleted').Trim().ToLowerInvariant() -eq 'true') { $emailRow = $null }
      $mailDate = if ($mailRow) { ConvertTo-LicensesActivityDate $mailRow.'Last Activity Date' } else { $null }
      $emailDate = if ($emailRow) { ConvertTo-LicensesActivityDate $emailRow.'Last Activity Date' } else { $null }
      $emailUserAction = $false
      if ($emailRow -and $emailDate -and $emailDate -ge $cutoff) {
        foreach ($field in @('Send Count','Read Count','Meeting Created Count','Meeting Interacted Count')) {
          $number = 0L
          if ([long]::TryParse(([string]$emailRow.$field).Trim(), [ref]$number) -and $number -gt 0) { $emailUserAction=$true; break }
        }
      }
      if ($mailRow) {
        if ((-not $mailDate -or $mailDate -lt $cutoff) -and -not $emailUserAction) { $count.MailboxInactive++ }
      }
      else { $count.MailboxUnknown++ }

      $m365 = $reports['M365_Users_Activity.csv']
      $apps = $reports['M365_Apps_Usage_180D.csv']
      $m365Row = if ($m365.Ready -and $upn -and $m365.Rows.ContainsKey($upn) -and -not $m365.Duplicates.ContainsKey($upn)) { $m365.Rows[$upn] } else { $null }
      if ($m365Row -and ([string]$m365Row.IsDeleted).Trim().ToLowerInvariant() -eq 'true') { $m365Row = $null }
      $appsRow = if ($apps.Ready -and $upn -and $apps.Rows.ContainsKey($upn) -and -not $apps.Duplicates.ContainsKey($upn)) { $apps.Rows[$upn] } else { $null }
      $m365Date = if ($m365Row) { ConvertTo-LicensesActivityDate $m365Row.LastActivityDate } else { $null }
      $appsDate = if ($appsRow) { ConvertTo-LicensesActivityDate $appsRow.'Last Activity Date' } else { $null }
      $hasRecentM365 = ($m365Date -and $m365Date -ge $cutoff) -or ($appsDate -and $appsDate -ge $cutoff) -or ($mailDate -and $mailDate -ge $cutoff) -or $emailUserAction
      if (-not $hasRecentM365 -and $m365Row) {
        $count.M365Inactive++
        if ($sharedStatus -eq 'NotShared') { $candidateUsers[$userId] = $true; $candidateReasons[$userId] = 'No M365 activity in 90 days' }
      }
      elseif (-not $hasRecentM365) {
        $count.M365Unknown++
        if ($sharedStatus -eq 'NotShared') { $unknownUsers[$userId] = $true }
      }

      if ($product.Name -in @('Microsoft 365 E3','Microsoft 365 E5')) {
        if ($appsRow) {
          $windows = ([string]$appsRow.Windows).Trim().ToLowerInvariant()
          $mac = ([string]$appsRow.Mac).Trim().ToLowerInvariant()
          if ($windows -notin @('yes','no') -or $mac -notin @('yes','no')) { $count.LocalAppsUnknown++ }
          elseif ($windows -eq 'no' -and $mac -eq 'no') { $count.LocalAppsInactive++ }
        }
        else { $count.LocalAppsUnknown++ }
      }
    }
    $count.RecoveryCandidates = $candidateUsers.Count
    $count.RecoveryUnknown = $unknownUsers.Count
    $count.RecoveryPrimaryPcUnknown = $unknownUsers.Count
    if ($intuneSource.Ready) {
      foreach ($userId in $candidateUsers.Keys) {
        if ($intunePrimaryPcUsers.ContainsKey($userId)) { $count.RecoveryPrimaryPc++ }
      }
    }
    else { $count.RecoveryPrimaryPcUnknown += $candidateUsers.Count }
    foreach ($userId in $candidateUsers.Keys) {
      $identity = if ($licenseIdentityByUser.ContainsKey($userId)) { $licenseIdentityByUser[$userId] } else { $null }
      $activeUser = if ($active.Ready -and $active.Rows.ContainsKey($userId) -and -not $active.Duplicates.ContainsKey($userId)) { $active.Rows[$userId] } else { $null }
      $upn = if ($activeUser -and $activeUser.'User principal name') { [string]$activeUser.'User principal name' } elseif ($identity) { [string]$identity.UserPrincipalName } else { '' }
      $displayName = if ($activeUser -and $activeUser.PSObject.Properties['Display name'] -and $activeUser.'Display name') { [string]$activeUser.'Display name' } elseif ($identity) { [string]$identity.DisplayName } else { '' }
      $detailUpn = $upn.Trim().ToLowerInvariant()
      $adActivityDate = $null
      $m365ActivityDate = $null
      $detailMailbox = if ($mailboxes.Ready -and $mailboxes.Rows.ContainsKey($userId) -and -not $mailboxes.Duplicates.ContainsKey($userId)) { $mailboxes.Rows[$userId] } else { $null }
      $isSharedMailbox = $detailMailbox -and ([string]$detailMailbox.RecipientTypeDetails).Trim() -eq 'SharedMailbox'
      if (-not $isSharedMailbox -and $activeUser -and $adSource.Ready) {
        $detailImmutable = ([string]$activeUser.OnPremisesImmutableId).Trim().ToLowerInvariant()
        $byImmutable = if ($detailImmutable -and $adByImmutable.ContainsKey($detailImmutable)) { $adByImmutable[$detailImmutable] } else { $null }
        $byUpn = if ($detailUpn -and -not $adDuplicateUpns.ContainsKey($detailUpn) -and $adByUpn.ContainsKey($detailUpn)) { $adByUpn[$detailUpn] } else { $null }
        if (-not ($byImmutable -and $byUpn -and -not [object]::ReferenceEquals($byImmutable,$byUpn)) -and
            -not ($detailUpn -and $adDuplicateUpns.ContainsKey($detailUpn) -and -not $byImmutable)) {
          $adMatch = if ($byImmutable) { $byImmutable } else { $byUpn }
          if ($adMatch) { $adActivityDate = ConvertTo-LicensesActivityDate $adMatch.LastLogonDate }
        }
      }
      $m365Report = $reports['M365_Users_Activity.csv']
      if (-not $isSharedMailbox -and $m365Report.Ready -and $detailUpn -and
          $m365Report.Rows.ContainsKey($detailUpn) -and -not $m365Report.Duplicates.ContainsKey($detailUpn)) {
        $m365ActivityRow = $m365Report.Rows[$detailUpn]
        if (([string]$m365ActivityRow.IsDeleted).Trim().ToLowerInvariant() -ne 'true') {
          $m365ActivityDate = ConvertTo-LicensesActivityDate $m365ActivityRow.LastActivityDate
        }
      }
      $recoveryDetails.Add([pscustomobject]@{
        License = $product.Name
        UserId = $userId
        UserPrincipalName = $upn
        DisplayName = $displayName
        RecoveryReason = [string]$candidateReasons[$userId]
        LastAdActivityDate = if ($adActivityDate) { $adActivityDate.ToString('yyyy-MM-dd') } else { 'N/D' }
        LastM365ActivityDate = if ($m365ActivityDate) { $m365ActivityDate.ToString('yyyy-MM-dd') } else { 'N/D' }
        PrimaryOnIntuneWindowsPc = if (-not $intuneSource.Ready) { 'N/D' } elseif ($intunePrimaryPcUsers.ContainsKey($userId)) { 'Yes' } else { 'No' }
        TargetSuites = (@($targetSuitesByUser[$userId].Keys | Sort-Object) -join ', ')
      })
    }
    [pscustomobject]@{ Product=$product.Name; Counts=$count; Available=$true }
  }
  $e3Review = [pscustomobject]@{ Available=$false; Assigned=0; Candidates=0; Unknown=0; Excluded=0; RecoveryExcluded=0; Reason='' }
  $e3Review.Assigned = $productUsers['Microsoft 365 E3'].Count
  $e3Review.Available = $true
  $e3ReviewDetails = [System.Collections.Generic.List[object]]::new()
  $e3RecoveryIds = @{}
  foreach ($item in $recoveryDetails) { if ($item.License -eq 'Microsoft 365 E3') { $e3RecoveryIds[[string]$item.UserId] = $true } }
  $appsReport = $reports['M365_Apps_Usage_180D.csv']
  foreach ($userId in $productUsers['Microsoft 365 E3'].Keys) {
    if ($e3RecoveryIds.ContainsKey($userId)) { $e3Review.RecoveryExcluded++; continue }
    if ($targetSuitesByUser[$userId].Count -gt 1) { $e3Review.Excluded++; continue }
    if (-not $active.Ready -or $active.Duplicates.ContainsKey($userId) -or -not $active.Rows.ContainsKey($userId) -or
        -not $mailboxes.Ready -or $mailboxes.Duplicates.ContainsKey($userId) -or -not $mailboxes.Rows.ContainsKey($userId)) { $e3Review.Unknown++; continue }
    $user = $active.Rows[$userId]
    $mailbox = $mailboxes.Rows[$userId]
    $enabled = ([string]$user.AccountEnabled).Trim().ToLowerInvariant()
    $kind = ([string]$mailbox.RecipientTypeDetails).Trim()
    $identityStatus = ([string]$mailbox.NativeIdentityStatus).Trim()
    if ($enabled -notin @('true','false') -or -not $kind -or $identityStatus -ne 'Observed') { $e3Review.Unknown++; continue }
    if ($enabled -eq 'false' -or $kind -ne 'UserMailbox') { $e3Review.Excluded++; continue }
    $sizeGb = ConvertTo-LicensesMailboxSizeGb $mailbox.TotalItemSizeGB
    $archive = ([string]$mailbox.ArchiveStatus).Trim().ToLowerInvariant()
    $litigation = ([string]$mailbox.LitigationHoldEnabled).Trim().ToLowerInvariant()
    $retention = ([string]$mailbox.RetentionHoldEnabled).Trim().ToLowerInvariant()
    if ($null -eq $sizeGb -or $archive -notin @('none','disabled','false','active','enabled','true') -or
        $litigation -notin @('true','false') -or $retention -notin @('true','false')) { $e3Review.Unknown++; continue }
    if ($sizeGb -ge 2 -or $archive -in @('active','enabled','true') -or $litigation -eq 'true' -or $retention -eq 'true') { $e3Review.Excluded++; continue }
    $upn = ([string]$user.'User principal name').Trim().ToLowerInvariant()
    if (-not $upn -or -not $appsReport.Ready -or $appsReport.Duplicates.ContainsKey($upn) -or -not $appsReport.Rows.ContainsKey($upn) -or
        -not $oneDrive.Ready -or $oneDrive.Duplicates.ContainsKey($upn) -or -not $oneDrive.Rows.ContainsKey($upn)) { $e3Review.Unknown++; continue }
    $appsRow = $appsReport.Rows[$upn]
    $driveRow = $oneDrive.Rows[$upn]
    $windows = ([string]$appsRow.Windows).Trim().ToLowerInvariant()
    $mac = ([string]$appsRow.Mac).Trim().ToLowerInvariant()
    $deleted = ([string]$driveRow.'Is Deleted').Trim().ToLowerInvariant()
    $bytesText = ([string]$driveRow.'Storage Used (Byte)').Trim()
    $bytes = 0L
    if ($windows -notin @('yes','no') -or $mac -notin @('yes','no') -or $deleted -notin @('true','false') -or
        ($deleted -eq 'false' -and (-not [long]::TryParse($bytesText, [ref]$bytes) -or $bytes -lt 0))) { $e3Review.Unknown++; continue }
    if ($windows -eq 'yes' -or $mac -eq 'yes' -or $deleted -eq 'true' -or $bytes -ge 2GB) { $e3Review.Excluded++; continue }
    $e3Review.Candidates++
    $identity = if ($licenseIdentityByUser.ContainsKey($userId)) { $licenseIdentityByUser[$userId] } else { $null }
    $displayName = if ($user.PSObject.Properties['Display name'] -and $user.'Display name') { [string]$user.'Display name' } elseif ($identity) { [string]$identity.DisplayName } else { '' }
    $e3ReviewDetails.Add([pscustomobject]@{
      UserId = $userId
      UserPrincipalName = [string]$user.'User principal name'
      DisplayName = $displayName
      MailboxSizeGB = $sizeGb
      OneDriveUsedBytes = $bytes
      OneDriveUsedGB = [math]::Round(([decimal]$bytes / [decimal]1GB),4)
      ArchiveStatus = [string]$mailbox.ArchiveStatus
      LitigationHoldEnabled = [string]$mailbox.LitigationHoldEnabled
      RetentionHoldEnabled = [string]$mailbox.RetentionHoldEnabled
      WindowsAppsUse180d = [string]$appsRow.Windows
      MacAppsUse180d = [string]$appsRow.Mac
    })
  }
  if ($e3Review.Candidates + $e3Review.Unknown + $e3Review.Excluded + $e3Review.RecoveryExcluded -ne $e3Review.Assigned) {
    throw 'E3 to F3 review categories do not reconcile with assigned E3 users.'
  }
  if ($e3ReviewDetails.Count -ne $e3Review.Candidates) { throw 'E3 to F3 review details do not reconcile with candidate count.' }
  return [pscustomobject]@{ Rows=@($metricRows); RecoveryDetails=$recoveryDetails.ToArray(); DowngradeReview=$e3Review; DowngradeDetails=$e3ReviewDetails.ToArray(); Sources=$sources.ToArray(); SharedSourceReady=$mailboxes.Ready; IntuneSourceReady=$intuneSource.Ready; AdSourceForced=$adSource.Forced; MailboxGap=$mailboxGap; AdGapActivity=$adGapActivity }
}

function Format-LicensesMetric {
  param(
    [AllowNull()]$Counts,
    [Parameter(Mandatory)][string]$ValueName,
    [Parameter(Mandatory)][string]$UnknownName,
    [bool]$Available
  )
  if (-not $Available -or $null -eq $Counts) { return 'N/D' }
  $value = [long]$Counts[$ValueName]
  $unknown = [long]$Counts[$UnknownName]
  if ($unknown -eq 0) { return [string]$value }
  if ($value -eq 0 -and $unknown -eq [long]$Counts.Assigned) { return "N/D ($unknown)" }
  return "$value (N/D: $unknown)"
}

function Get-LicensesAssignmentIssues {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][object[]]$TenantRows,
    [Parameter(Mandatory)][datetimeoffset]$LicenseSnapshotUtc,
    [datetime]$AsOfUtc = [datetime]::UtcNow,
    [AllowNull()]$CurrentRunPublication
  )
  $productParts = [ordered]@{
    'Microsoft 365 F1' = @('M365_F1','M365_F1_COMM')
    'Microsoft 365 F3' = @('SPE_F1')
    'Microsoft 365 E3' = @('SPE_E3')
    'Microsoft 365 E5' = @('SPE_E5')
  }
  $counts = [ordered]@{}
  foreach ($product in $productParts.Keys) { $counts[$product] = [ordered]@{ CapacityWait=0; OtherErrors=0 } }
  $specs = @(
    @{File='M365_Licenses_AssignmentPaths.csv'; Columns=@('TenantKey','UserId','SkuId','AssignedByGroupId','AssignmentRoute','AssignmentState','AssignmentError','LastUpdatedDateTime'); Receipt='SmartInventory_SmartM365-Licences-Inventory.current.json.txt'},
    @{File='M365_Licenses_Users.csv'; Columns=@('TenantKey','UserId','SkuPartNumber'); Receipt='SmartInventory_SmartM365-Licences-Inventory.current.json.txt'},
    @{File='M365_Users_Active.csv'; Columns=@('TenantKey','Object Id','User principal name','Display name','AccountEnabled','UserType'); Receipt='SmartInventory_SmartM365-ActiveUsers-Inventory.current.json.txt'},
    @{File='M365_EntraGroups_All.csv'; Columns=@('TenantKey','GroupId','DisplayName'); Receipt='SmartInventory_SmartM365-Licences-Inventory.current.json.txt'}
  )
  $sources = [System.Collections.Generic.List[object]]::new()
  foreach ($spec in $specs) {
    $source = Get-LicensesCsvSource -Folder $Folder -FileName $spec.File -Columns $spec.Columns -AsOfUtc $AsOfUtc -CollectorManifestName $spec.Receipt -RequireFileReceipt -CurrentRunPublication $CurrentRunPublication
    if ($source.Ready -and $spec.File -ne 'M365_Users_Active.csv' -and
        [math]::Abs(($source.ModifiedUtc - $LicenseSnapshotUtc.UtcDateTime).TotalHours) -gt 24) {
      $source.Ready = $false
      $source.Reason = 'assignment source does not match the tenant snapshot date (over 24 hours apart)'
    }
    $sources.Add($source)
  }
  $result = [pscustomobject]@{ Available=$false; Reason=''; Counts=$counts; Rows=@(); Sources=$sources.ToArray() }
  $rejected = @($sources | Where-Object { -not $_.Ready })
  if ($rejected.Count -gt 0) {
    $result.Reason = (@($rejected | ForEach-Object { '{0}: {1}' -f $_.Name,$_.Reason }) -join '; ')
    return $result
  }
  $runIds = @($sources | Where-Object { $_.Name -ne 'M365_Users_Active.csv' } | ForEach-Object { [string]$_.Proof.RunId } | Sort-Object -Unique)
  if ($runIds.Count -ne 1 -or [string]::IsNullOrWhiteSpace($runIds[0])) {
    $result.Reason = 'license assignment files do not share one completed collection run'
    return $result
  }
  try {
    $skuProduct = @{}
    foreach ($tenantRow in $TenantRows) {
      $suite = Get-LicensesSuiteProduct -SkuId $tenantRow.Id -SkuPartNumber $tenantRow.TenantSkuPartNumber
      foreach ($product in $productParts.Keys) {
        if ($suite -ne $product) { continue }
        $skuId = ([string]$tenantRow.Id).Trim().ToLowerInvariant()
        if (-not $skuId) { throw "Tenant SKU ID is missing for $product." }
        if ($skuProduct.ContainsKey($skuId) -and $skuProduct[$skuId] -ne $product) { throw "Tenant SKU ID maps to multiple target suites: $skuId" }
        $skuProduct[$skuId] = $product
      }
    }
    $effective = @{}
    foreach ($row in (Import-LicensesSourceCsv -Source $sources[1])) {
      if ([string]$row.TenantKey -ine $TenantKey) { throw 'another tenant is present in effective license assignments' }
      $userId = ([string]$row.UserId).Trim().ToLowerInvariant()
      $suite = Get-LicensesSuiteProduct -SkuId (Get-LicensesOptionalProperty $row 'SkuId') -SkuPartNumber $row.SkuPartNumber
      if (-not $userId) { throw 'effective license assignment has no UserId' }
      foreach ($product in $productParts.Keys) {
        if ($suite -eq $product) { $effective["$userId|$product"] = $true; break }
      }
    }
    $paths = @{}
    $wantedUsers = @{}
    foreach ($row in (Import-LicensesSourceCsv -Source $sources[0])) {
      if ([string]$row.TenantKey -ine $TenantKey) { throw 'another tenant is present in license assignment paths' }
      $skuId = ([string]$row.SkuId).Trim().ToLowerInvariant()
      if (-not $skuProduct.ContainsKey($skuId)) { continue }
      $state = ([string]$row.AssignmentState).Trim()
      $errorCode = ([string]$row.AssignmentError).Trim()
      if ($state -notin @('Error','ActiveWithError') -and $errorCode -in @('','None')) { continue }
      $userId = ([string]$row.UserId).Trim().ToLowerInvariant()
      if (-not $userId) { throw 'failed license assignment has no UserId' }
      $groupId = ([string]$row.AssignedByGroupId).Trim()
      $route = ([string]$row.AssignmentRoute).Trim()
      if (($groupId -and $route -ne 'Group') -or (-not $groupId -and $route -ne 'Direct')) { throw 'failed license assignment route differs from its group ID' }
      $product = $skuProduct[$skuId]
      $hasEffective = $effective.ContainsKey("$userId|$product")
      $category = if ($errorCode -eq 'CountViolation' -and -not $hasEffective) { 'Capacity wait' } else { 'Licensing error' }
      $key = "$userId|$product|$category|$errorCode"
      if (-not $paths.ContainsKey($key)) { $paths[$key] = [System.Collections.Generic.List[object]]::new() }
      $paths[$key].Add([pscustomobject]@{ UserId=$userId; Product=$product; SkuId=$skuId; Category=$category; ErrorCode=$errorCode; State=$state; GroupId=$groupId; Route=$route; Updated=[string]$row.LastUpdatedDateTime; HasEffective=$hasEffective })
      $wantedUsers[$userId] = $true
    }
    $users = Read-LicensesIndexedSource -Source $sources[2] -ExpectedTenantKey $TenantKey -KeyColumn 'Object Id' -WantedKeys $wantedUsers -AsOfUtc $AsOfUtc
    if (-not $users.Ready) { throw "Active user source: $($users.Source.Reason)" }
    $groups = @{}
    foreach ($row in (Import-LicensesSourceCsv -Source $sources[3])) {
      if ([string]$row.TenantKey -ine $TenantKey) { throw 'another tenant is present in licensing groups' }
      $groupId = ([string]$row.GroupId).Trim().ToLowerInvariant()
      if (-not $groupId -or $groups.ContainsKey($groupId)) { throw 'licensing group ID is missing or duplicated' }
      $groups[$groupId] = [string]$row.DisplayName
    }
    $rows = foreach ($key in ($paths.Keys | Sort-Object)) {
      $items = @($paths[$key])
      $first = $items[0]
      $identity = if ($users.Rows.ContainsKey($first.UserId) -and -not $users.Duplicates.ContainsKey($first.UserId)) { $users.Rows[$first.UserId] } else { $null }
      $groupIds = @($items.GroupId | Where-Object { $_ } | Sort-Object -Unique)
      $groupNames = @($groupIds | ForEach-Object { if ($groups.ContainsKey($_.ToLowerInvariant())) { $groups[$_.ToLowerInvariant()] } else { 'N/D' } })
      [pscustomobject][ordered]@{
        Product=$first.Product; Category=$first.Category; UserId=$first.UserId
        UserPrincipalName=if ($identity) { [string]$identity.'User principal name' } else { '' }
        DisplayName=if ($identity) { [string]$identity.'Display name' } else { '' }
        IdentityStatus=if ($users.Duplicates.ContainsKey($first.UserId)) { 'Ambiguous' } elseif ($identity) { 'Matched' } else { 'Not found' }
        AccountEnabled=if ($identity) { [string]$identity.AccountEnabled } else { '' }
        UserType=if ($identity) { [string]$identity.UserType } else { '' }
        SkuIds=(@($items.SkuId | Sort-Object -Unique) -join ';')
        AssignmentStates=(@($items.State | Sort-Object -Unique) -join ';')
        AssignmentError=$first.ErrorCode; AssignmentRoutes=(@($items.Route | Sort-Object -Unique) -join ';')
        AssignedByGroupIds=($groupIds -join ';'); AssignedByGroupNames=($groupNames -join ';')
        HasEffectiveTargetSuite=[bool]$first.HasEffective
        LastUpdatedDateTime=(@($items.Updated | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1) -join '')
        AssignmentPaths=$items.Count
      }
    }
    foreach ($product in $productParts.Keys) {
      $counts[$product].CapacityWait = @($rows | Where-Object { $_.Product -eq $product -and $_.Category -eq 'Capacity wait' } | Select-Object -ExpandProperty UserId -Unique).Count
      $counts[$product].OtherErrors = @($rows | Where-Object { $_.Product -eq $product -and $_.Category -eq 'Licensing error' } | Select-Object -ExpandProperty UserId -Unique).Count
    }
    $result.Available = $true
    $result.Rows = @($rows)
  }
  catch { $result.Reason = $_.Exception.Message; $result.Rows = @() }
  return $result
}

function Get-LicensesAlertSummary {
  param([AllowNull()]$AssignmentIssues)
  $products = [ordered]@{ F1='Microsoft 365 F1'; F3='Microsoft 365 F3'; E3='Microsoft 365 E3'; E5='Microsoft 365 E5' }
  $result = [pscustomobject]@{
    AwaitingCapacity='N/D'; LicensingErrors='N/D'
    AwaitingBreakdown='F1 N/D · F3 N/D · E3 N/D · E5 N/D'
    ErrorsBreakdown='F1 N/D · F3 N/D · E3 N/D · E5 N/D'
  }
  if (-not $AssignmentIssues -or -not $AssignmentIssues.Available) { return $result }

  $waitingUsers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $errorUsers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($row in @($AssignmentIssues.Rows)) {
    if ($row.Category -eq 'Capacity wait') { [void]$waitingUsers.Add([string]$row.UserId) }
    elseif ($row.Category -eq 'Licensing error') { [void]$errorUsers.Add([string]$row.UserId) }
  }
  $result.AwaitingCapacity = [string]$waitingUsers.Count
  $result.LicensingErrors = [string]$errorUsers.Count
  $result.AwaitingBreakdown = (@(foreach ($label in $products.Keys) { '{0} {1}' -f $label,[long]$AssignmentIssues.Counts[$products[$label]].CapacityWait }) -join ' · ')
  $result.ErrorsBreakdown = (@(foreach ($label in $products.Keys) { '{0} {1}' -f $label,[long]$AssignmentIssues.Counts[$products[$label]].OtherErrors }) -join ' · ')
  return $result
}

function New-LicensesRecoveryWorkbook {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][object[]]$SummaryRows,
    [AllowNull()]$Usage,
    [AllowNull()]$AssignmentIssues,
    [Parameter(Mandatory)][string]$CollectedAtUtc
  )

  Initialize-SmartM365ImportExcel | Out-Null
  $usageByProduct = @{}
  if ($Usage) { foreach ($item in $Usage.Rows) { $usageByProduct[$item.Product] = $item } }
  $details = @(if ($Usage -and $Usage.RecoveryDetails) { $Usage.RecoveryDetails })
  $downgradeDetails = @(if ($Usage -and $Usage.DowngradeDetails) { $Usage.DowngradeDetails })
  $downgradeReview = if ($Usage) { $Usage.DowngradeReview } else { $null }
  if ($downgradeReview -and $downgradeReview.Available -and
      [long]$downgradeReview.Candidates -ne $downgradeDetails.Count) {
    throw ("E3 to F3 review detail count differs from the email KPI: {0} versus {1}." -f $downgradeDetails.Count,$downgradeReview.Candidates)
  }
  $summary = foreach ($product in $SummaryRows) {
    $metric = if ($usageByProduct.ContainsKey($product.Product)) { $usageByProduct[$product.Product] } else { $null }
    $available = ($metric -and $metric.Available)
    $expected = if ($available) { [long]$metric.Counts.RecoveryCandidates } else { $null }
    $actual = @($details | Where-Object License -eq $product.Product).Count
    if ($available -and $expected -ne $actual) {
      throw ("Recovery detail count differs from the email KPI for {0}: {1} versus {2}." -f $product.Product,$actual,$expected)
    }
    [pscustomobject]@{
      License = $product.Product
      RecoveryCandidates = if ($available) { $expected } else { 'N/D' }
      UnknownUsers = if ($available) { [long]$metric.Counts.RecoveryUnknown } else { 'N/D' }
      E3toF3ReviewCandidates = if ($product.Product -ne 'Microsoft 365 E3') { 'N/A' } elseif ($downgradeReview -and $downgradeReview.Available) { [long]$downgradeReview.Candidates } else { 'N/D' }
      AwaitingLicense = if ($AssignmentIssues -and $AssignmentIssues.Available) { [long]$AssignmentIssues.Counts[$product.Product].CapacityWait } else { 'N/D' }
      LicensingErrors = if ($AssignmentIssues -and $AssignmentIssues.Available) { [long]$AssignmentIssues.Counts[$product.Product].OtherErrors } else { 'N/D' }
      Qualification = if (-not $available) { 'Source not qualified' } else { 'Qualified' }
      LicenseSnapshotUtc = $CollectedAtUtc
    }
  }
  $summary | Export-Excel -Path $Path -WorksheetName 'Summary' -TableName 'LicenseRecoverySummary' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -ErrorAction Stop
  if ($details.Count -gt 0) {
    $safeDetails = foreach ($item in ($details | Sort-Object License,UserPrincipalName,UserId)) {
      $safeUpn = [string]$item.UserPrincipalName
      $safeDisplayName = [string]$item.DisplayName
      if ($safeUpn -match '^\s*[=+\-@]') { $safeUpn = "'$safeUpn" }
      if ($safeDisplayName -match '^\s*[=+\-@]') { $safeDisplayName = "'$safeDisplayName" }
      $adDateText = ([string]$item.LastAdActivityDate).Trim()
      $m365DateText = ([string]$item.LastM365ActivityDate).Trim()
      [pscustomobject]@{
        License = $item.License
        UserId = $item.UserId
        UserPrincipalName = $safeUpn
        DisplayName = $safeDisplayName
        RecoveryReason = $item.RecoveryReason
        LastAdActivityDate = if ($adDateText -and $adDateText -ne 'N/D') { [datetime]::ParseExact($adDateText,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture) } else { $null }
        LastM365ActivityDate = if ($m365DateText -and $m365DateText -ne 'N/D') { [datetime]::ParseExact($m365DateText,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture) } else { $null }
        PrimaryOnIntuneWindowsPc = $item.PrimaryOnIntuneWindowsPc
        TargetSuites = $item.TargetSuites
      }
    }
    $safeDetails | Export-Excel -Path $Path -WorksheetName 'Recovery candidates' -TableName 'LicenseRecoveryCandidates' -Append -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -ErrorAction Stop
    $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
    try {
      $sheet = $package.Workbook.Worksheets['Recovery candidates']
      $sheet.Cells[2,6,$sheet.Dimension.End.Row,7].Style.Numberformat.Format = 'yyyy-mm-dd'
    }
    finally { Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop }
  }
  else {
    $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
    try {
      $sheet = $package.Workbook.Worksheets.Add('Recovery candidates')
      $headers = @('License','UserId','UserPrincipalName','DisplayName','RecoveryReason','LastAdActivityDate','LastM365ActivityDate','PrimaryOnIntuneWindowsPc','TargetSuites')
      for ($column=0; $column -lt $headers.Count; $column++) { $sheet.Cells[1,($column+1)].Value = $headers[$column] }
      $sheet.Cells[1,1,1,$headers.Count].Style.Font.Bold = $true
      $sheet.View.FreezePanes(2,1)
    }
    finally { Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop }
  }
  if ($downgradeDetails.Count -gt 0) {
    $safeDowngradeDetails = foreach ($item in ($downgradeDetails | Sort-Object UserPrincipalName,UserId)) {
      $safeUpn = [string]$item.UserPrincipalName
      $safeDisplayName = [string]$item.DisplayName
      if ($safeUpn -match '^\s*[=+\-@]') { $safeUpn = "'$safeUpn" }
      if ($safeDisplayName -match '^\s*[=+\-@]') { $safeDisplayName = "'$safeDisplayName" }
      [pscustomobject]@{
        UserId = [string]$item.UserId
        UserPrincipalName = $safeUpn
        DisplayName = $safeDisplayName
        MailboxSizeGB = [decimal]$item.MailboxSizeGB
        OneDriveUsedGB = [decimal]$item.OneDriveUsedGB
        OneDriveUsedBytes = [long]$item.OneDriveUsedBytes
        ArchiveStatus = [string]$item.ArchiveStatus
        LitigationHoldEnabled = [string]$item.LitigationHoldEnabled
        RetentionHoldEnabled = [string]$item.RetentionHoldEnabled
        WindowsAppsUse180d = [string]$item.WindowsAppsUse180d
        MacAppsUse180d = [string]$item.MacAppsUse180d
      }
    }
    $safeDowngradeDetails | Export-Excel -Path $Path -WorksheetName 'E3 to F3 review' -TableName 'LicenseE3ToF3Review' -Append -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -ErrorAction Stop
  }
  else {
    $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
    try {
      $sheet = $package.Workbook.Worksheets.Add('E3 to F3 review')
      $headers = @('UserId','UserPrincipalName','DisplayName','MailboxSizeGB','OneDriveUsedGB','OneDriveUsedBytes','ArchiveStatus','LitigationHoldEnabled','RetentionHoldEnabled','WindowsAppsUse180d','MacAppsUse180d')
      for ($column=0; $column -lt $headers.Count; $column++) { $sheet.Cells[1,($column+1)].Value = $headers[$column] }
      $sheet.Cells[1,1,1,$headers.Count].Style.Font.Bold = $true
      $sheet.View.FreezePanes(2,1)
    }
    finally { Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop }
  }
  foreach ($category in @('Capacity wait','Licensing error')) {
    $sheetName = if ($category -eq 'Capacity wait') { 'Awaiting licenses' } else { 'Licensing errors' }
    $tableName = if ($category -eq 'Capacity wait') { 'LicenseCapacityWait' } else { 'LicenseAssignmentErrors' }
    $issueRows = @(if ($AssignmentIssues -and $AssignmentIssues.Available) { $AssignmentIssues.Rows | Where-Object Category -eq $category })
    if ($issueRows.Count -gt 0) {
      $safeRows = foreach ($item in ($issueRows | Sort-Object Product,UserPrincipalName,UserId)) {
        $upn = [string]$item.UserPrincipalName
        $display = [string]$item.DisplayName
        $groupNames = [string]$item.AssignedByGroupNames
        if ($upn -match '^\s*[=+\-@]') { $upn = "'$upn" }
        if ($display -match '^\s*[=+\-@]') { $display = "'$display" }
        if ($groupNames -match '^\s*[=+\-@]') { $groupNames = "'$groupNames" }
        $updated = ([string]$item.LastUpdatedDateTime).Trim()
        [pscustomobject][ordered]@{
          Product=$item.Product; UserId=$item.UserId; UserPrincipalName=$upn; DisplayName=$display
          AccountEnabled=$item.AccountEnabled; UserType=$item.UserType; IdentityStatus=$item.IdentityStatus
          AssignmentError=$item.AssignmentError; AssignmentStates=$item.AssignmentStates
          AssignmentRoutes=$item.AssignmentRoutes; AssignedByGroupNames=$groupNames; AssignedByGroupIds=$item.AssignedByGroupIds
          HasEffectiveTargetSuite=$item.HasEffectiveTargetSuite
          LastUpdatedDateTime=if ($updated) { [datetimeoffset]::Parse($updated,[Globalization.CultureInfo]::InvariantCulture).UtcDateTime } else { $null }
          AssignmentPaths=$item.AssignmentPaths
        }
      }
      $safeRows | Export-Excel -Path $Path -WorksheetName $sheetName -TableName $tableName -Append -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -ErrorAction Stop
      $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
      try {
        $sheet = $package.Workbook.Worksheets[$sheetName]
        $sheet.Cells[2,14,$sheet.Dimension.End.Row,14].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm:ss'
      }
      finally { Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop }
    }
    else {
      $package = Open-ExcelPackage -Path $Path -ErrorAction Stop
      try {
        $sheet = $package.Workbook.Worksheets.Add($sheetName)
        $headers = @('Product','UserId','UserPrincipalName','DisplayName','AccountEnabled','UserType','IdentityStatus','AssignmentError','AssignmentStates','AssignmentRoutes','AssignedByGroupNames','AssignedByGroupIds','HasEffectiveTargetSuite','LastUpdatedDateTime','AssignmentPaths')
        for ($column=0; $column -lt $headers.Count; $column++) { $sheet.Cells[1,($column+1)].Value = $headers[$column] }
        $sheet.Cells[1,1,1,$headers.Count].Style.Font.Bold = $true
        $sheet.View.FreezePanes(2,1)
      }
      finally { Close-ExcelPackage -ExcelPackage $package -ErrorAction Stop }
    }
  }
  $file = Get-Item -LiteralPath $Path -ErrorAction Stop
  if ($file.Length -gt 3MB) {
    throw ('Recovery workbook is {0:N1} MB; Graph mail attachments must be at most 3 MB.' -f ($file.Length / 1MB))
  }
  return $file.FullName
}

function Publish-LicensesReportCsv {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)]$Snapshot
  )
  $summaryColumns = @('SnapshotId','TenantKey','GeneratedAtUtc','LicenseCollectedAtUtc','Product','Category','EnabledUnits','ConsumedUnits','AvailableUnits','Subscribed','AssignedUsers','AssignmentCapacityWait','AssignmentOtherErrors','AssignmentEvidenceStatus','RecoveryCandidates','RecoveryUnknown','DisabledUsers','DisabledUnknown','NoM365Activity','NoM365ActivityUnknown','MultipleTargetSuites','MultipleTargetSuitesUnknown','MultipleAssignedSkus','MultipleAssignedSkusUnknown','NoAdOrEntraActivity','NoAdOrEntraActivityUnknown','NoMailboxActivity','NoMailboxActivityUnknown','NoLocalAppsUse','NoLocalAppsUseUnknown','CandidatesPrimaryIntunePc','CandidatesPrimaryIntunePcUnknown','SharedMailboxesLicensed','SharedUnder50Gb','SharedRemovalCandidates','SharedUnknown','E3ToF3ReviewAssigned','E3ToF3ReviewCandidates','E3ToF3ReviewUnknown','E3ToF3ReviewExcluded','E3ToF3ReviewRecoveryExcluded','EvidenceStatus','WorkbookQualification')
  $candidateColumns = @('SnapshotId','TenantKey','CandidateType','Product','UserId','TenantUserKey','IdentityJoinStatus','UserPrincipalName','DisplayName','Reason','LastAdActivityDate','LastM365ActivityDate','PrimaryOnIntuneWindowsPc','TargetSuites','MailboxSizeGB','OneDriveUsedGB','OneDriveUsedBytes','ArchiveStatus','LitigationHoldEnabled','RetentionHoldEnabled','WindowsAppsUse180d','MacAppsUse180d')
  $assignmentColumns = @('SnapshotId','TenantKey','Product','Category','UserId','UserPrincipalName','DisplayName','IdentityStatus','AccountEnabled','UserType','SkuIds','AssignmentStates','AssignmentError','AssignmentRoutes','AssignedByGroupIds','AssignedByGroupNames','HasEffectiveTargetSuite','LastUpdatedDateTime','AssignmentPaths')
  $gapColumns = @('SnapshotId','TenantKey','Section','Metric','Value','EvidenceStatus')
  $sourceColumns = @('SnapshotId','TenantKey','Name','Ready','Provisional','Date','Reason','SHA256')
  $provisional = @($Snapshot.Sources | Where-Object Provisional).Count -gt 0
  $licenseProvisional = @($Snapshot.Sources | Where-Object { $_.Name -eq 'M365_Licenses_Users.csv' -and $_.Provisional }).Count -gt 0
  $summary = @(
    foreach ($category in @('Suite','Other paid product')) {
      $products = if ($category -eq 'Suite') { $Snapshot.Products } else { $Snapshot.OtherProducts }
      foreach ($product in $products) {
        $counts = if ($category -eq 'Suite') { $product.Counts } else { $null }
        $count = { param($key) if ($null -ne $counts) { $counts[$key] } else { $null } }
        $review = if ($product.Product -eq 'Microsoft 365 E3' -and $Snapshot.E3ToF3Review -and $Snapshot.E3ToF3Review.Available) { $Snapshot.E3ToF3Review } else { $null }
        [pscustomobject][ordered]@{
          SnapshotId=$Snapshot.SnapshotId; TenantKey=$Snapshot.TenantKey; GeneratedAtUtc=$Snapshot.GeneratedAtUtc; LicenseCollectedAtUtc=$Snapshot.LicenseCollectedAtUtc
          Product=$product.Product; Category=$category; EnabledUnits=$product.Enabled; ConsumedUnits=$product.Consumed
          AvailableUnits=([long]$product.Enabled - [long]$product.Consumed); Subscribed=[bool]$product.Subscribed
          AssignedUsers=(& $count 'Assigned'); RecoveryCandidates=(& $count 'RecoveryCandidates'); RecoveryUnknown=(& $count 'RecoveryUnknown')
          AssignmentCapacityWait=if ($category -eq 'Suite' -and $Snapshot.AssignmentIssues.Available) { $Snapshot.AssignmentIssues.Counts[$product.Product].CapacityWait } else { $null }
          AssignmentOtherErrors=if ($category -eq 'Suite' -and $Snapshot.AssignmentIssues.Available) { $Snapshot.AssignmentIssues.Counts[$product.Product].OtherErrors } else { $null }
          AssignmentEvidenceStatus=if ($category -ne 'Suite') { 'N/A' } elseif ($Snapshot.AssignmentIssues.Available) { 'Qualified' } else { 'N/D' }
          DisabledUsers=(& $count 'Disabled'); DisabledUnknown=(& $count 'DisabledUnknown')
          NoM365Activity=(& $count 'M365Inactive'); NoM365ActivityUnknown=(& $count 'M365Unknown')
          MultipleTargetSuites=(& $count 'Multiple'); MultipleTargetSuitesUnknown=(& $count 'MultipleUnknown')
          MultipleAssignedSkus=(& $count 'MultipleAll'); MultipleAssignedSkusUnknown=(& $count 'MultipleAllUnknown')
          NoAdOrEntraActivity=(& $count 'AdEntraInactive'); NoAdOrEntraActivityUnknown=(& $count 'AdEntraUnknown')
          NoMailboxActivity=(& $count 'MailboxInactive'); NoMailboxActivityUnknown=(& $count 'MailboxUnknown')
          NoLocalAppsUse=(& $count 'LocalAppsInactive'); NoLocalAppsUseUnknown=(& $count 'LocalAppsUnknown')
          CandidatesPrimaryIntunePc=(& $count 'RecoveryPrimaryPc'); CandidatesPrimaryIntunePcUnknown=(& $count 'RecoveryPrimaryPcUnknown')
          SharedMailboxesLicensed=(& $count 'SharedLicensed'); SharedUnder50Gb=(& $count 'SharedUnder50')
          SharedRemovalCandidates=(& $count 'SharedEligible'); SharedUnknown=(& $count 'SharedUnknown')
          E3ToF3ReviewAssigned=if ($review) { $review.Assigned } else { $null }
          E3ToF3ReviewCandidates=if ($review) { $review.Candidates } else { $null }; E3ToF3ReviewUnknown=if ($review) { $review.Unknown } else { $null }
          E3ToF3ReviewExcluded=if ($review) { $review.Excluded } else { $null }
          E3ToF3ReviewRecoveryExcluded=if ($review) { $review.RecoveryExcluded } else { $null }
          EvidenceStatus=if ($category -ne 'Suite') { 'Capacity only' } elseif (-not $product.UsageAvailable) { 'N/D' } elseif ($provisional) { 'Provisional' } else { 'Qualified' }
          WorkbookQualification=if ($category -ne 'Suite') { '' } elseif (-not $product.UsageAvailable) { 'Source not qualified' } elseif ($licenseProvisional) { 'Provisional: license receipt bypassed' } else { 'Qualified' }
        }
      }
    }
  )
  $candidates = @(
    foreach ($candidateType in @('Recovery','E3 to F3 review')) {
      $details = if ($candidateType -eq 'Recovery') { $Snapshot.RecoveryCandidates } else { $Snapshot.DowngradeCandidates }
      foreach ($detail in $details) {
        $recovery = $candidateType -eq 'Recovery'
        [pscustomobject][ordered]@{
          SnapshotId=$Snapshot.SnapshotId; TenantKey=$Snapshot.TenantKey; CandidateType=$candidateType
          Product=if ($recovery) { $detail.License } else { 'Microsoft 365 E3' }; UserId=$detail.UserId
          TenantUserKey=''; IdentityJoinStatus='Not joined in SmartM365'; UserPrincipalName=$detail.UserPrincipalName; DisplayName=$detail.DisplayName
          Reason=if ($recovery) { $detail.RecoveryReason } else { 'Manual downgrade review' }
          LastAdActivityDate=if ($recovery -and $detail.LastAdActivityDate -ne 'N/D') { $detail.LastAdActivityDate } else { '' }
          LastM365ActivityDate=if ($recovery -and $detail.LastM365ActivityDate -ne 'N/D') { $detail.LastM365ActivityDate } else { '' }
          PrimaryOnIntuneWindowsPc=if ($recovery) { $detail.PrimaryOnIntuneWindowsPc } else { '' }
          TargetSuites=if ($recovery) { $detail.TargetSuites } else { 'Microsoft 365 E3' }
          MailboxSizeGB=if ($recovery) { '' } else { $detail.MailboxSizeGB }; OneDriveUsedGB=if ($recovery) { '' } else { $detail.OneDriveUsedGB }
          OneDriveUsedBytes=if ($recovery) { '' } else { $detail.OneDriveUsedBytes }; ArchiveStatus=if ($recovery) { '' } else { $detail.ArchiveStatus }
          LitigationHoldEnabled=if ($recovery) { '' } else { $detail.LitigationHoldEnabled }; RetentionHoldEnabled=if ($recovery) { '' } else { $detail.RetentionHoldEnabled }
          WindowsAppsUse180d=if ($recovery) { '' } else { $detail.WindowsAppsUse180d }; MacAppsUse180d=if ($recovery) { '' } else { $detail.MacAppsUse180d }
        }
      }
    }
  )
  $assignmentIssues = @($Snapshot.AssignmentIssues.Rows | ForEach-Object {
    [pscustomobject][ordered]@{
      SnapshotId=$Snapshot.SnapshotId; TenantKey=$Snapshot.TenantKey; Product=$_.Product; Category=$_.Category
      UserId=$_.UserId; UserPrincipalName=$_.UserPrincipalName; DisplayName=$_.DisplayName; IdentityStatus=$_.IdentityStatus
      AccountEnabled=$_.AccountEnabled; UserType=$_.UserType; SkuIds=$_.SkuIds; AssignmentStates=$_.AssignmentStates
      AssignmentError=$_.AssignmentError; AssignmentRoutes=$_.AssignmentRoutes; AssignedByGroupIds=$_.AssignedByGroupIds
      AssignedByGroupNames=$_.AssignedByGroupNames; HasEffectiveTargetSuite=$_.HasEffectiveTargetSuite
      LastUpdatedDateTime=$_.LastUpdatedDateTime; AssignmentPaths=$_.AssignmentPaths
    }
  })
  $gapMetrics = [ordered]@{
    UserMailboxes=@('Total','Universe','OtherSkus','NoSkus','Unknown','EntraEnabled','EntraDisabled','EntraStateUnknown')
    NoUserMailbox=@('Total','Universe','OtherSkus','NoSkus','Unknown','Guests','MemberEnabled','MemberDisabled')
    AdGapActivity=@('Members','AdObserved','NoObservedMatch','Ambiguous','AdEnabledRecent','AdEnabledInactive','AdEnabledNoDate','AdDisabled','AdEnabledUnknown')
  }
  $gaps = @(
    foreach ($section in $gapMetrics.Keys) {
      $record = if ($section -eq 'AdGapActivity') { $Snapshot.AdGapActivity } elseif ($Snapshot.MailboxGap) { $Snapshot.MailboxGap[$section] } else { $null }
      foreach ($metric in $gapMetrics[$section]) {
        $value = if ($record -and $record.Available) { $record[$metric] } else { $null }
        [pscustomobject][ordered]@{ SnapshotId=$Snapshot.SnapshotId; TenantKey=$Snapshot.TenantKey; Section=$section; Metric=$metric; Value=$value
          EvidenceStatus=if ($null -eq $value) { 'N/D' } elseif ($record.Contains('Provisional') -and $record.Provisional) { 'Provisional' } else { 'Qualified' } }
      }
    }
  )
  $sources = @($Snapshot.Sources | ForEach-Object {
    [pscustomobject][ordered]@{ SnapshotId=$Snapshot.SnapshotId; TenantKey=$Snapshot.TenantKey; Name=$_.Name; Ready=$_.Ready; Provisional=$_.Provisional; Date=$_.Date; Reason=$_.Reason; SHA256=$_.SHA256 }
  })
  if ($Snapshot.AssignmentIssues.Available) {
    foreach ($product in @($Snapshot.Products)) {
      foreach ($check in @(@{Category='Capacity wait';Field='CapacityWait'},@{Category='Licensing error';Field='OtherErrors'})) {
        $actual = @($assignmentIssues | Where-Object { $_.Product -eq $product.Product -and $_.Category -eq $check.Category } | Select-Object -ExpandProperty UserId -Unique).Count
        $expected = [long]$Snapshot.AssignmentIssues.Counts[$product.Product][$check.Field]
        if ($actual -ne $expected) { throw "License assignment issue CSV does not reconcile with $($product.Product) $($check.Category): $actual versus $expected." }
      }
    }
  }
  if (@($summary).Count -ne 7 -or @($candidates | Where-Object CandidateType -eq 'Recovery').Count -ne @($Snapshot.RecoveryCandidates).Count -or
      @($candidates | Where-Object CandidateType -eq 'E3 to F3 review').Count -ne @($Snapshot.DowngradeCandidates).Count -or
      @($assignmentIssues).Count -ne @($Snapshot.AssignmentIssues.Rows).Count -or @($gaps).Count -ne 25) {
    throw 'License report CSV rows do not reconcile with the mail snapshot.'
  }
  $publishedPaths = [System.Collections.Generic.List[string]]::new()
  foreach ($spec in @(
    @{ Name='M365_Licenses_ReportSummary.csv'; Columns=$summaryColumns; Rows=$summary },
    @{ Name='M365_Licenses_ReportCandidates.csv'; Columns=$candidateColumns; Rows=$candidates },
    @{ Name='M365_Licenses_ReportAssignmentIssues.csv'; Columns=$assignmentColumns; Rows=$assignmentIssues },
    @{ Name='M365_Licenses_ReportGaps.csv'; Columns=$gapColumns; Rows=$gaps },
    @{ Name='M365_Licenses_ReportSources.csv'; Columns=$sourceColumns; Rows=$sources }
  )) {
    $path = Get-LicensesPublicationPath -Path (Join-Path $Folder $spec.Name)
    $pending = "$path.$($Snapshot.SnapshotId).pending"
    try {
      $lines = if (@($spec.Rows).Count -gt 0) { @($spec.Rows | Select-Object -Property $spec.Columns | ConvertTo-Csv -NoTypeInformation) }
               else { @(($spec.Columns -join ',')) }
      $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($lines -join "`r`n") + "`r`n")
      $stream = [IO.File]::Open($pending, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
      try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
      [IO.File]::Move($pending, $path, $true)
      $readBack = @(Import-Csv -LiteralPath $path -ErrorAction Stop)
      if ($readBack.Count -ne @($spec.Rows).Count -or @($readBack | Where-Object { $_.SnapshotId -cne $Snapshot.SnapshotId }).Count -gt 0) {
        throw "Published license report CSV does not match its snapshot: $($spec.Name)"
      }
      $publishedPaths.Add($path)
    }
    finally { if (Test-Path -LiteralPath $pending -PathType Leaf) { Remove-Item -LiteralPath $pending -Force -ErrorAction SilentlyContinue } }
  }
  # Registration and upload belong to the caller so the receipt records the files before any copy.
  return $publishedPaths.ToArray()
}

function Publish-LicensesReportSnapshot {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][string]$CollectedAtUtc,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SummaryRows,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AdditionalRows,
    [AllowNull()]$Usage,
    [AllowNull()]$AssignmentIssues,
    [AllowNull()]$ReceiptRun
  )
  $snapshotId = [guid]::NewGuid().ToString('N')
  $metricByProduct = @{}
  if ($Usage) { foreach ($metric in @($Usage.Rows)) { $metricByProduct[[string]$metric.Product] = $metric } }
  $products = foreach ($row in $SummaryRows) {
    $metric = if ($metricByProduct.ContainsKey([string]$row.Product)) { $metricByProduct[[string]$row.Product] } else { $null }
    $counts = $null
    if ($metric -and $metric.Available) {
      $counts = [ordered]@{}
      foreach ($key in @('Assigned','Disabled','DisabledUnknown','AdEntraInactive','AdEntraUnknown',
                         'MailboxInactive','MailboxUnknown','M365Inactive','M365Unknown',
                         'LocalAppsInactive','LocalAppsUnknown','Multiple','MultipleUnknown',
                         'MultipleAll','MultipleAllUnknown','SharedLicensed','SharedUnder50',
                         'SharedEligible','SharedUnknown','RecoveryCandidates','RecoveryUnknown',
                         'RecoveryPrimaryPc','RecoveryPrimaryPcUnknown')) {
        $counts[$key] = if (($key -like 'LocalApps*' -and $row.Product -notin @('Microsoft 365 E3','Microsoft 365 E5')) -or
                            ($key -like 'Shared*' -and -not $Usage.SharedSourceReady) -or
                            ($key -like 'RecoveryPrimaryPc*' -and -not $Usage.IntuneSourceReady)) {
          $null
        } else { [long]$metric.Counts[$key] }
      }
    }
    [ordered]@{
      Product = [string]$row.Product
      Enabled = [long]$row.Enabled
      Consumed = [long]$row.Consumed
      Subscribed = [bool]$row.Subscribed
      UsageAvailable = [bool]($metric -and $metric.Available)
      Counts = $counts
    }
  }
  $otherProducts = foreach ($row in $AdditionalRows) {
    [ordered]@{ Product=[string]$row.Product; Enabled=[long]$row.Enabled; Consumed=[long]$row.Consumed; Subscribed=[bool]$row.Subscribed }
  }
  $sources = [System.Collections.Generic.List[object]]::new()
  # This run's tenant CSV: its staged copy while the run is staged, never the previous run's DATA-LAST file.
  $tenantPath = Get-LicensesPublicationPath -Path (Join-Path $Folder 'M365_Licenses_Tenant.csv')
  if (-not (Test-Path -LiteralPath $tenantPath -PathType Leaf)) { throw 'Tenant license CSV is missing before report snapshot publication.' }
  $tenantHashBefore = (Get-FileHash -LiteralPath $tenantPath -Algorithm SHA256).Hash
  $tenantSource = Read-LicensesTenantSnapshot -Path $tenantPath -ExpectedTenantKey $TenantKey
  foreach ($pair in @(
      @{ Sent=$SummaryRows; Published=@(Get-LicensesFocusedSummaryRows -TenantRows $tenantSource.Rows) },
      @{ Sent=$AdditionalRows; Published=@(Get-LicensesAdditionalOverviewRows -TenantRows $tenantSource.Rows) })) {
    for ($i=0; $i -lt $pair.Sent.Count; $i++) {
      if ([string]$pair.Sent[$i].Product -ne [string]$pair.Published[$i].Product -or
          [long]$pair.Sent[$i].Enabled -ne [long]$pair.Published[$i].Enabled -or
          [long]$pair.Sent[$i].Consumed -ne [long]$pair.Published[$i].Consumed -or
          [bool]$pair.Sent[$i].Subscribed -ne [bool]$pair.Published[$i].Subscribed) {
        throw 'Mail capacity differs from the published tenant license CSV; report snapshot was not published.'
      }
    }
  }
  if ((Get-FileHash -LiteralPath $tenantPath -Algorithm SHA256).Hash -cne $tenantHashBefore) {
    throw 'Tenant license CSV changed during report snapshot preparation.'
  }
  $sources.Add([ordered]@{ Name='M365_Licenses_Tenant.csv'; Ready=$true; Provisional=$false; Date=$tenantSource.CollectedAtUtc; Reason=''; SHA256=$tenantHashBefore })
  $seenSources = @{}
  $reportSources = [System.Collections.Generic.List[object]]::new()
  if ($Usage) { foreach ($source in @($Usage.Sources)) { $reportSources.Add($source) } }
  if ($AssignmentIssues) { foreach ($source in @($AssignmentIssues.Sources)) { $reportSources.Add($source) } }
  foreach ($source in $reportSources) {
      if ($seenSources.ContainsKey([string]$source.Name)) { continue }
      $seenSources[[string]$source.Name] = $true
      $hash = ''
      if ($source.Ready) {
        if (-not (Test-Path -LiteralPath $source.Path -PathType Leaf)) { throw "Qualified source disappeared before snapshot publication: $($source.Name)" }
        if ($source.ModifiedUtc -and (Get-Item -LiteralPath $source.Path -ErrorAction Stop).LastWriteTimeUtc -ne $source.ModifiedUtc) {
          throw "Qualified source changed during license report preparation: $($source.Name)"
        }
        $hash = (Get-FileHash -LiteralPath $source.Path -Algorithm SHA256).Hash
        if ($source.PSObject.Properties['SHA256'] -and $source.SHA256 -and $hash -cne [string]$source.SHA256) {
          throw "Qualified source bytes changed during license report preparation: $($source.Name)"
        }
        if ($source.PSObject.Properties['ReceiptPath'] -and $source.ReceiptPath) {
          Assert-SmartM365SourcePublication -ReceiptPath $source.ReceiptPath -Proof $source.Proof
          if ((Get-FileHash -LiteralPath $source.ReceiptPath -Algorithm SHA256).Hash -cne [string]$source.ReceiptHash) {
            throw "Qualified source receipt changed during license report preparation: $($source.Name)"
          }
        }
      }
      $sourceMetadata = [ordered]@{ Name=[string]$source.Name; Ready=[bool]$source.Ready; Provisional=([bool]$source.Forced -or ($source.PSObject.Properties['PartialScopeAccepted'] -and $source.PartialScopeAccepted)); Date=[string]$source.Date; Reason=[string]$source.Reason; SHA256=$hash }
      if ($source.PSObject.Properties['PartialScopeAccepted'] -and $source.PartialScopeAccepted) { $sourceMetadata.DomainCoverage = $source.DomainCoverage }
      $sources.Add($sourceMetadata)
  }
  $gap = $null
  if ($Usage -and $Usage.MailboxGap) {
    $gap = [ordered]@{}
    foreach ($name in @('UserMailboxes','NoUserMailbox')) {
      $item = $Usage.MailboxGap.$name
      if (-not $item -or -not $item.Available) {
        $gap[$name] = [ordered]@{ Available=$false; Reason=if ($item) { [string]$item.Reason } else { 'source unavailable' } }
        continue
      }
      $values = [ordered]@{ Available=$true; Reason='' }
      foreach ($key in @('Total','Universe','OtherSkus','NoSkus','Unknown','Guests','MemberEnabled','MemberDisabled','EntraEnabled','EntraDisabled','EntraStateUnknown')) {
        $property = $item.PSObject.Properties[$key]
        if ($property) { $values[$key] = [long]$property.Value }
      }
      if ($name -eq 'UserMailboxes') {
        $values.EntraStateAvailable = [bool]$item.EntraStateAvailable
        if (-not $item.EntraStateAvailable) {
          foreach ($key in @('EntraEnabled','EntraDisabled','EntraStateUnknown')) { $values[$key] = $null }
        }
      }
      $gap[$name] = $values
    }
  }
  $adGap = $null
  if ($Usage -and $Usage.AdGapActivity) {
    $ad = $Usage.AdGapActivity
    $adGap = [ordered]@{ Available=[bool]$ad.Available; Provisional=[bool]$ad.Provisional; Reason=[string]$ad.Reason }
    if ($ad.Available) {
      foreach ($key in @('Members','AdObserved','NoObservedMatch','Ambiguous','AdEnabledRecent','AdEnabledInactive','AdEnabledNoDate','AdDisabled','AdEnabledUnknown')) {
        $adGap[$key] = if ($null -eq $ad.$key) { $null } else { [long]$ad.$key }
      }
      $adGap.UnmatchedScopeUnknown = [long]$ad.UnmatchedScopeUnknown
    }
  }
  $recovery = @(if ($Usage) { $Usage.RecoveryDetails })
  $downgrade = @(if ($Usage) { $Usage.DowngradeDetails })
  $assignment = if ($AssignmentIssues) { $AssignmentIssues } else { [pscustomobject]@{ Available=$false; Reason='assignment issues not analyzed'; Counts=@{}; Rows=@() } }
  foreach ($product in $products) {
    if ($product.UsageAvailable -and @($recovery | Where-Object License -eq $product.Product).Count -ne $product.Counts.RecoveryCandidates) {
      throw "Recovery detail does not reconcile with $($product.Product)."
    }
  }
  if ($Usage -and $Usage.DowngradeReview -and $Usage.DowngradeReview.Available -and
      $downgrade.Count -ne [long]$Usage.DowngradeReview.Candidates) { throw 'Downgrade detail does not reconcile with its KPI.' }
  $snapshot = [ordered]@{
    SchemaVersion = 2
    LogicVersion = '1.44'
    SnapshotId = $snapshotId
    TenantKey = $TenantKey
    GeneratedAtUtc = [datetimeoffset]::UtcNow.ToString('o')
    LicenseCollectedAtUtc = $CollectedAtUtc
    TenantFileCollectedAtUtc = $tenantSource.CollectedAtUtc
    Products = @($products)
    OtherProducts = @($otherProducts)
    E3ToF3Review = if ($Usage -and $Usage.DowngradeReview) { $Usage.DowngradeReview } else { $null }
    MailboxGap = $gap
    AdGapActivity = $adGap
    RecoveryCandidates = $recovery
    DowngradeCandidates = $downgrade
    AssignmentIssues = [ordered]@{ Available=[bool]$assignment.Available; Reason=[string]$assignment.Reason
      Counts=if ($assignment.Available) { $assignment.Counts } else { $null }; Rows=@($assignment.Rows) }
    Sources = $sources.ToArray()
    # Additive in schema 2: lets the email render from the published snapshot without recomputation.
    Rendering = [ordered]@{
      UsageAvailable = [bool]$Usage
      SharedSourceReady = [bool]($Usage -and $Usage.SharedSourceReady)
      IntuneSourceReady = [bool]($Usage -and $Usage.IntuneSourceReady)
      AdSourceForced = [bool]($Usage -and $Usage.AdSourceForced)
      EntraStateReason = if ($Usage -and $Usage.MailboxGap -and $Usage.MailboxGap.UserMailboxes) { [string]$Usage.MailboxGap.UserMailboxes.EntraStateReason } else { '' }
      UsageSources = @(if ($Usage) {
        foreach ($source in @($Usage.Sources)) {
          [ordered]@{ Name=[string]$source.Name; Ready=[bool]$source.Ready; Forced=[bool]$source.Forced
            PartialScopeAccepted=[bool]($source.PSObject.Properties['PartialScopeAccepted'] -and $source.PartialScopeAccepted)
            Date=[string]$source.Date; Provenance=[string]$source.Provenance; Reason=[string]$source.Reason }
        }
      })
    }
  }
  # Additive in schema 2: the licensing receipt run that proves this report's licensing inputs.
  if ($ReceiptRun) { $snapshot.ReceiptRun = [ordered]@{ Producer=[string]$ReceiptRun.Producer; RunId=[string]$ReceiptRun.RunId; StartedAtUtc=[string]$ReceiptRun.StartedAtUtc } }
  # Promoted with the report CSVs and the receipt when the run is staged.
  $path = Get-LicensesPublicationPath -Path (Join-Path $Folder 'M365_Licenses_ReportSnapshot.json.txt') -Artifact
  $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($snapshot | ConvertTo-Json -Depth 16 -Compress))
  $document = [Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json -ErrorAction Stop
  if ([string]$document.SnapshotId -ne $snapshotId -or [string]$document.TenantKey -ne $TenantKey -or
      @($document.Products).Count -ne 4 -or @($document.OtherProducts).Count -ne 3) {
    throw 'License report snapshot failed its identity or product validation.'
  }
  $pending = "$path.$snapshotId.pending"
  try {
    $stream = [IO.File]::Open($pending, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
    $expectedHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    if ((Get-FileHash -LiteralPath $pending -Algorithm SHA256).Hash -ne $expectedHash) { throw 'Staged license report snapshot hash differs.' }
    [IO.File]::Move($pending, $path, $true)
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $expectedHash) { throw 'Published license report snapshot hash differs.' }
  }
  finally { if (Test-Path -LiteralPath $pending -PathType Leaf) { Remove-Item -LiteralPath $pending -Force -ErrorAction SilentlyContinue } }
  $csvPaths = @(Publish-LicensesReportCsv -Folder $Folder -Snapshot $snapshot)
  return [pscustomobject]@{ Id=$snapshotId; Path=$path; Data=$snapshot; CsvPaths=$csvPaths }
}

function Write-LicensesDailyMailState {
  param(
    [Parameter(Mandatory)][System.IO.FileStream]$Stream,
    [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes
  )
  $Stream.SetLength(0)
  $Stream.Position = 0
  $Stream.Write($Bytes, 0, $Bytes.Length)
  $Stream.Flush($true)
}

function Enter-LicensesDailyMailGate {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$TenantKey,
    [switch]$Force,
    [datetimeoffset]$NowUtc = [datetimeoffset]::UtcNow
  )
  if (-not (Test-Path -LiteralPath (Split-Path -Path $Path -Parent) -PathType Container)) {
    throw "License summary mail state folder does not exist: $Path"
  }
  $timeZone = try { [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time') }
              catch { [TimeZoneInfo]::FindSystemTimeZoneById('Europe/Paris') }
  $day = [TimeZoneInfo]::ConvertTime($NowUtc, $timeZone).ToString('yyyy-MM-dd')
  $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  try {
    if ($stream.Length -gt 16384) { throw "License summary mail state is unexpectedly large: $Path" }
    $previous = [byte[]]::new([int]$stream.Length)
    if ($previous.Length -gt 0) {
      $stream.Position = 0
      $read = 0
      while ($read -lt $previous.Length) {
        $count = $stream.Read($previous, $read, $previous.Length - $read)
        if ($count -eq 0) { throw "License summary mail state could not be read: $Path" }
        $read += $count
      }
      try {
        $state = [Text.Encoding]::UTF8.GetString($previous) | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if ($state -isnot [System.Collections.IDictionary] -or
            -not $state.Contains('TenantKey') -or -not $state.Contains('LocalDate') -or
            [string]$state.Status -notin @('Sent','Pending') -or
            [string]$state.LocalDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
          throw 'Invalid license summary mail state structure.'
        }
      }
      catch {
        if (-not $Force) { throw "License summary mail state is unreadable; use -ForceLicenseSummaryEmail after checking delivery: $Path" }
        $state = $null
      }
      if ($state -and [string]$state.TenantKey -ne $TenantKey) {
        throw "License summary mail state belongs to another tenant: $Path"
      }
      if ($state -and [string]$state.LocalDate -eq $day -and -not $Force) {
        $reason = if ([string]$state.Status -eq 'Sent') { "already sent on $day" }
                  elseif ([string]$state.Status -eq 'Pending') { "delivery is pending or uncertain on $day" }
                  else { throw "License summary mail state has an unknown status: $Path" }
        $stream.Dispose()
        return [pscustomobject]@{ Skip=$true; Reason=$reason; Stream=$null; Day=$day; PreviousBytes=$previous }
      }
    }
    $pending = [ordered]@{ TenantKey=$TenantKey; LocalDate=$day; Status='Pending'; LastAttemptUtc=$NowUtc.ToString('o') }
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($pending | ConvertTo-Json -Compress))
    Write-LicensesDailyMailState -Stream $stream -Bytes $bytes
    return [pscustomobject]@{ Skip=$false; Reason=''; Stream=$stream; Day=$day; PreviousBytes=$previous }
  }
  catch {
    $stream.Dispose()
    throw
  }
}

function Get-LicensesOptionalProperty {
  param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)
  if ($null -eq $InputObject) { return $null }
  $property = $InputObject.PSObject.Properties[$Name]
  if ($property) { return $property.Value }
  return $null
}

function Get-LicensesDecisionSettings {
  param([AllowNull()]$Config, [Parameter(Mandatory)][string]$ExclusionsPath)
  # Exclusions are versioned with the script so repository updates always reach the servers;
  # the local JSON can only add exclusions, never remove a versioned one.
  if (-not (Test-Path -LiteralPath $ExclusionsPath -PathType Leaf)) { throw "License decision exclusions are missing: $ExclusionsPath" }
  try { $document = [IO.File]::ReadAllText($ExclusionsPath) | ConvertFrom-Json -ErrorAction Stop }
  catch { throw "License decision exclusions are not valid JSON: $($_.Exception.Message)" }
  if (-not [string]$document.Version -or @($document.Exclusions).Count -eq 0) { throw 'License decision exclusions need a Version and at least one exclusion.' }
  $excluded = @{}
  foreach ($item in @($document.Exclusions)) {
    $sku = ([string]$item.SkuPartNumber).Trim()
    if (-not $sku -or -not ([string]$item.Reason).Trim()) { throw 'Each versioned license exclusion needs SkuPartNumber and Reason.' }
    $excluded[$sku.ToUpperInvariant()] = [string]$item.Reason
  }
  foreach ($item in @(Get-ScriptLocalConfigValue -Config $Config -Name 'LicenseDecisionAdditionalExcludedSkus' -DefaultValue @())) {
    if ($null -eq $item) { continue }
    $sku = ([string]$item.SkuPartNumber).Trim()
    if (-not $sku -or -not ([string]$item.Reason).Trim()) { throw 'LicenseDecisionAdditionalExcludedSkus entries need SkuPartNumber and Reason.' }
    if (-not $excluded.ContainsKey($sku.ToUpperInvariant())) { $excluded[$sku.ToUpperInvariant()] = 'Local: {0}' -f [string]$item.Reason }
  }
  $settings = [pscustomobject]@{
    CopilotInactivityDays = [int](Get-ScriptLocalConfigValue -Config $Config -Name 'LicenseDecisionCopilotInactivityDays' -DefaultValue 60)
    CopilotAssignmentGraceDays = [int](Get-ScriptLocalConfigValue -Config $Config -Name 'LicenseDecisionCopilotAssignmentGraceDays' -DefaultValue 30)
    MailboxReviewSizeGB = [double](Get-ScriptLocalConfigValue -Config $Config -Name 'LicenseDecisionMailboxReviewSizeGB' -DefaultValue 50)
    ExcludedSkus = $excluded
  }
  if ($settings.CopilotInactivityDays -lt 1 -or $settings.CopilotAssignmentGraceDays -lt 0 -or $settings.MailboxReviewSizeGB -le 0) {
    throw 'License decision settings must be positive (inactivity days, grace days, mailbox review size).'
  }
  return $settings
}

function Read-LicensesDependencyRules {
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "License dependency rules are missing: $Path" }
  $document = [IO.File]::ReadAllText($Path) | ConvertFrom-Json -ErrorAction Stop
  $rules = [System.Collections.Generic.List[object]]::new()
  foreach ($dependency in @($document.Dependencies)) {
    $dependents = @($dependency.Dependent | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
    $qualifying = @($dependency.QualifyingAnyOf | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
    if ($dependents.Count -eq 0 -or $qualifying.Count -eq 0 -or -not [string]$dependency.Source -or -not [string]$dependency.SourceDate) {
      throw 'Each license dependency needs Dependent, QualifyingAnyOf, Source and SourceDate.'
    }
    $rules.Add([pscustomobject]@{ Dependents=$dependents; Qualifying=$qualifying; Source=[string]$dependency.Source; SourceDate=[string]$dependency.SourceDate })
  }
  return $rules.ToArray()
}

function Get-LicensesAccountState {
  param([AllowNull()]$ActiveRow, [Parameter(Mandatory)]$AdSource, [System.Collections.IDictionary]$AdByImmutable, [System.Collections.IDictionary]$AdByUpn, [System.Collections.IDictionary]$DuplicateAdUpns)
  # Entra and AD together; a disagreement is a review, never a removal.
  if ($null -eq $ActiveRow) { return [pscustomobject]@{ State='Unknown'; Reason='Account not found in the Entra active users export'; Entra='N/D'; Ad='N/D' } }
  $entra = ([string]$ActiveRow.AccountEnabled).Trim().ToLowerInvariant()
  if ($entra -notin @('true','false')) { return [pscustomobject]@{ State='Unknown'; Reason='Entra account state is not measured'; Entra='N/D'; Ad='N/D' } }
  $immutable = ([string]$ActiveRow.OnPremisesImmutableId).Trim().ToLowerInvariant()
  if (-not $immutable) {
    $state = if ($entra -eq 'false') { 'Disabled' } else { 'Enabled' }
    return [pscustomobject]@{ State=$state; Reason='Cloud-only account'; Entra=$entra; Ad='Not synchronized' }
  }
  if (-not $AdSource.Ready) {
    if ($entra -eq 'true') { return [pscustomobject]@{ State='Enabled'; Reason=''; Entra=$entra; Ad='N/D' } }
    return [pscustomobject]@{ State='Review'; Reason='Entra disabled; AD state is not measured'; Entra=$entra; Ad='N/D' }
  }
  $upn = ([string]$ActiveRow.'User principal name').Trim().ToLowerInvariant()
  $adRow = $null
  if ($AdByImmutable.Contains($immutable)) { $adRow = $AdByImmutable[$immutable] }
  elseif ($upn -and $AdByUpn.Contains($upn) -and -not $DuplicateAdUpns.Contains($upn)) { $adRow = $AdByUpn[$upn] }
  if ($null -eq $adRow) {
    if ($entra -eq 'true') { return [pscustomobject]@{ State='Enabled'; Reason=''; Entra=$entra; Ad='Not found' } }
    $why = if ($AdSource.PartialScopeAccepted) { 'Entra disabled; AD account not found in the collected domains' } else { 'Entra disabled; synchronized AD account not found' }
    return [pscustomobject]@{ State='Review'; Reason=$why; Entra=$entra; Ad='Not found' }
  }
  $ad = ([string]$adRow.Enabled).Trim().ToLowerInvariant()
  if ($ad -notin @('true','false')) { return [pscustomobject]@{ State='Review'; Reason='AD account state is not measured'; Entra=$entra; Ad='N/D' } }
  if ($entra -eq 'false' -and $ad -eq 'false') { return [pscustomobject]@{ State='Disabled'; Reason='Disabled in Entra and AD'; Entra=$entra; Ad=$ad } }
  if ($entra -eq 'true' -and $ad -eq 'true') { return [pscustomobject]@{ State='Enabled'; Reason=''; Entra=$entra; Ad=$ad } }
  $conflict = if ($entra -eq 'false') { 'Entra disabled but AD enabled' } else { 'AD disabled but Entra enabled' }
  return [pscustomobject]@{ State='Review'; Reason="Directory state conflict: $conflict"; Entra=$entra; Ad=$ad }
}

function Get-LicensesMailboxGuard {
  param([AllowNull()]$Mailbox, [bool]$SourceReady, [double]$ReviewSizeGB)
  # Same retention checks as the suite recovery: size, archive, litigation and retention holds; unknown fails closed.
  if (-not $SourceReady) { return [pscustomobject]@{ Status='Unknown'; Reason='EXO mailbox source is not qualified' } }
  if ($null -eq $Mailbox) { return [pscustomobject]@{ Status='Unknown'; Reason='No observed EXO mailbox for the account' } }
  if (([string]$Mailbox.NativeIdentityStatus).Trim() -ne 'Observed') { return [pscustomobject]@{ Status='Unknown'; Reason='EXO mailbox identity is not observed' } }
  $size = ConvertTo-LicensesMailboxSizeGb $Mailbox.TotalItemSizeGB
  $archive = ([string]$Mailbox.ArchiveStatus).Trim().ToLowerInvariant()
  $litigation = ([string]$Mailbox.LitigationHoldEnabled).Trim().ToLowerInvariant()
  $retention = ([string]$Mailbox.RetentionHoldEnabled).Trim().ToLowerInvariant()
  if ($null -eq $size -or $archive -notin @('none','disabled','false','active','enabled','true') -or
      $litigation -notin @('true','false') -or $retention -notin @('true','false')) {
    return [pscustomobject]@{ Status='Unknown'; Reason='Mailbox size, archive or hold state is not measured' }
  }
  $blocked = @()
  if ($size -gt $ReviewSizeGB) { $blocked += ('mailbox above {0} GB' -f $ReviewSizeGB) }
  if ($archive -in @('active','enabled','true')) { $blocked += 'archive active' }
  if ($litigation -eq 'true') { $blocked += 'litigation hold' }
  if ($retention -eq 'true') { $blocked += 'retention hold' }
  if ($blocked.Count -gt 0) { return [pscustomobject]@{ Status='Blocked'; Reason=('Mailbox retention: ' + ($blocked -join ', ')) } }
  return [pscustomobject]@{ Status='Clear'; Reason='' }
}

function Get-LicensesCopilotUsageState {
  param([Parameter(Mandatory)]$Source, [datetime]$AsOfUtc, [int]$MinimumPeriodDays = 60)
  # Not measured, never zero: absent, stale, short-period or anonymized reports produce no candidate.
  $state = [pscustomobject]@{ Measured=$false; Reason=''; Rows=@{}; WindowEnd=$null }
  if (-not $Source.Ready) { $state.Reason = "Copilot usage report not qualified: $($Source.Reason)"; return $state }
  try {
    $rows = @(Import-LicensesSourceCsv -Source $Source)
    if ($rows.Count -eq 0) { $state.Reason = 'Copilot usage report is empty'; return $state }
    $refresh = $null; $index = @{}
    foreach ($row in $rows) {
      $period = ([string]$row.ReportPeriod).Trim()
      if ($period -notmatch '^D(\d+)$' -or [int]$Matches[1] -lt $MinimumPeriodDays) { $state.Reason = 'Copilot usage report period is shorter than the inactivity window'; return $state }
      $date = ConvertTo-LicensesActivityDate $row.ReportRefreshDate
      if ($null -eq $date) { $state.Reason = 'Copilot usage report refresh date is missing'; return $state }
      if ($null -eq $refresh) { $refresh = $date } elseif ($refresh -ne $date) { $state.Reason = 'Copilot usage report has mixed refresh dates'; return $state }
      $upn = ([string]$row.UserPrincipalName).Trim().ToLowerInvariant()
      if ($upn -notmatch '^[^@\s]+@[^@\s]+$') { $state.Reason = 'Copilot usage report user names are concealed (report anonymization)'; return $state }
      if ($index.ContainsKey($upn)) { $state.Reason = 'Copilot usage report has duplicate user rows'; return $state }
      $index[$upn] = $row
    }
    if ($refresh -lt $AsOfUtc.Date.AddDays(-14)) { $state.Reason = 'Copilot usage report refresh date is older than 14 days'; return $state }
    $state.Measured = $true; $state.Rows = $index; $state.WindowEnd = $refresh
  }
  catch { $state.Reason = "Copilot usage report unreadable: $($_.Exception.Message)" }
  return $state
}

function Get-LicensesCopilotActivity {
  param([Parameter(Mandatory)]$Row)
  # Only licensed Microsoft 365 Copilot app activity justifies the license; chat, Edge and agents do not.
  $licensed = $null; $chat = $null
  foreach ($column in @('Microsoft365CopilotLastActivityDate','TeamsCopilotLastActivityDate','WordCopilotLastActivityDate','ExcelCopilotLastActivityDate','PowerPointCopilotLastActivityDate','OutlookCopilotLastActivityDate','OneNoteCopilotLastActivityDate','LoopCopilotLastActivityDate')) {
    $date = ConvertTo-LicensesActivityDate $Row.$column
    if ($date -and (-not $licensed -or $date -gt $licensed)) { $licensed = $date }
  }
  foreach ($column in @('CopilotChatLastActivityDate','CopilotChatWorkLastActivityDate','CopilotChatWebLastActivityDate','EdgeLastActivityDate','CopilotAgentLastActivityDate')) {
    $date = ConvertTo-LicensesActivityDate $Row.$column
    if ($date -and (-not $chat -or $date -gt $chat)) { $chat = $date }
  }
  return [pscustomobject]@{ Licensed=$licensed; Chat=$chat }
}

function Get-LicensesSkuDecisions {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows,
    [AllowNull()]$CurrentRunPublication,
    [AllowEmptyCollection()][object[]]$RecoveryDetails = @(),
    [AllowEmptyCollection()][object[]]$DowngradeDetails = @(),
    [Parameter(Mandatory)]$Settings,
    [Parameter(Mandatory)][object[]]$DependencyRules,
    [datetime]$AsOfUtc = [datetime]::UtcNow
  )
  $licensesReceipt = 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  $sources = [System.Collections.Generic.List[object]]::new()
  $usersSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Licenses_Users.csv' -Columns @('TenantKey','UserId','SkuId','SkuPartNumber','Source','GroupsAssigningSku') -AsOfUtc $AsOfUtc -CollectorManifestName $licensesReceipt -RequireFileReceipt -CurrentRunPublication $CurrentRunPublication
  $pathsSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Licenses_AssignmentPaths.csv' -Columns @('TenantKey','UserId','SkuId','AssignmentRoute','LastUpdatedDateTime') -AsOfUtc $AsOfUtc -CollectorManifestName $licensesReceipt -RequireFileReceipt -CurrentRunPublication $CurrentRunPublication
  $catalogSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Licenses_ServicePlans_Catalog.csv' -Columns @('TenantKey','SkuId','PlanName') -AsOfUtc $AsOfUtc -CollectorManifestName $licensesReceipt -RequireFileReceipt -CurrentRunPublication $CurrentRunPublication
  foreach ($source in @($usersSource,$pathsSource,$catalogSource)) {
    $sources.Add($source)
    if (-not $source.Ready) { throw "License decision source not qualified: $($source.Name): $($source.Reason)" }
  }
  $activeSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Users_Active.csv' -Columns @('TenantKey','Object Id','User principal name','AccountEnabled','OnPremisesImmutableId') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-ActiveUsers-Inventory.current.json.txt'
  $adSource = Get-LicensesCsvSource -Folder $Folder -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD','UserPrincipalName','Enabled') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt' -AllowQualifiedPartialAdScope
  $exoSource = Get-LicensesCsvSource -Folder $Folder -FileName 'Exchange_EXO_Mailboxes_AllDomains.csv' -Columns @('TenantKey','ExternalDirectoryObjectId','RecipientTypeDetails','NativeIdentityStatus','TotalItemSizeGB','ArchiveStatus','LitigationHoldEnabled','RetentionHoldEnabled') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
  $copilotSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Copilot_UserUsage.csv' -Columns @('TenantKey','ReportPeriod','ReportRefreshDate','UserPrincipalName','Microsoft365CopilotLastActivityDate','CopilotChatLastActivityDate') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-CopilotUsage-Inventory.current.json.txt'
  $phoneSource = Get-LicensesCsvSource -Folder $Folder -FileName 'M365_Teams_PhoneAssignments.csv' -Columns @('TenantKey','AssignedPstnTargetId','TargetType') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-TeamsPhonePstnUsage-Inventory.current.json.txt'
  foreach ($source in @($activeSource,$adSource,$exoSource,$copilotSource,$phoneSource)) { $sources.Add($source) }
  if (-not $activeSource.Ready) { throw "License decision source not qualified: M365_Users_Active.csv: $($activeSource.Reason)" }

  $skuById = @{}
  foreach ($row in $TenantRows) {
    $id = ([string]$row.Id).Trim().ToLowerInvariant()
    if (-not $id) { throw 'Tenant SKU ID is missing.' }
    $skuById[$id] = $row
  }
  $plansBySku = @{}
  foreach ($row in (Import-LicensesSourceCsv -Source $catalogSource)) {
    if ([string]$row.TenantKey -cne $TenantKey) { throw 'another tenant is present in the service-plan catalog' }
    $id = ([string]$row.SkuId).Trim().ToLowerInvariant()
    if (-not $plansBySku.ContainsKey($id)) { $plansBySku[$id] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
    [void]$plansBySku[$id].Add(([string]$row.PlanName).Trim())
  }
  # Mailbox and archive plans grant a mailbox; EXCHANGE_S_FOUNDATION only supports other workloads.
  $exchangePlans = @('EXCHANGE_S_ENTERPRISE','EXCHANGE_S_STANDARD','EXCHANGE_S_DESKLESS','EXCHANGE_S_ESSENTIALS','EXCHANGE_S_ARCHIVE_ADDON','EXCHANGE_S_ARCHIVE')
  $hasPlan = { param($skuId, [string[]]$names) $set = $plansBySku[$skuId]; if (-not $set) { return $false }; foreach ($name in $names) { if ($set.Contains($name)) { return $true } }; return $false }

  $assignments = @{}
  $userIdentity = @{}
  foreach ($row in (Import-LicensesSourceCsv -Source $usersSource)) {
    if ([string]$row.TenantKey -cne $TenantKey) { throw 'another tenant is present in license assignments' }
    $userId = ([string]$row.UserId).Trim().ToLowerInvariant(); $skuId = ([string]$row.SkuId).Trim().ToLowerInvariant()
    if (-not $userId -or -not $skuId) { throw 'a license assignment has no UserId or SkuId' }
    $key = "$userId|$skuId"
    if (-not $assignments.ContainsKey($key)) {
      $assignments[$key] = [pscustomobject]@{ UserId=$userId; SkuId=$skuId; SkuPartNumber=([string]$row.SkuPartNumber).Trim(); Route=([string]$row.Source).Trim(); Groups=([string]$row.GroupsAssigningSku).Trim() }
    }
    if (-not $userIdentity.ContainsKey($userId)) {
      $userIdentity[$userId] = [pscustomobject]@{ UserPrincipalName=[string]$row.'User principal name'; DisplayName=[string]$row.'Display name' }
    }
  }
  $pathDates = @{}
  foreach ($row in (Import-LicensesSourceCsv -Source $pathsSource)) {
    if ([string]$row.TenantKey -cne $TenantKey) { throw 'another tenant is present in license assignment paths' }
    $key = '{0}|{1}' -f ([string]$row.UserId).Trim().ToLowerInvariant(), ([string]$row.SkuId).Trim().ToLowerInvariant()
    if (-not $pathDates.ContainsKey($key)) { $pathDates[$key] = [System.Collections.Generic.List[object]]::new() }
    $parsed = [datetimeoffset]::MinValue
    $text = ([string]$row.LastUpdatedDateTime).Trim()
    $pathDates[$key].Add($(if ($text -and [datetimeoffset]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { $parsed.UtcDateTime } else { $null }))
  }
  $skusByUser = @{}
  foreach ($assignment in $assignments.Values) {
    if (-not $skusByUser.ContainsKey($assignment.UserId)) { $skusByUser[$assignment.UserId] = [System.Collections.Generic.List[object]]::new() }
    $skusByUser[$assignment.UserId].Add($assignment)
  }
  $wanted = @{}; foreach ($userId in $skusByUser.Keys) { $wanted[$userId] = $true }
  $active = Read-LicensesIndexedSource -Source $activeSource -ExpectedTenantKey $TenantKey -KeyColumn 'Object Id' -WantedKeys $wanted -AsOfUtc $AsOfUtc
  if (-not $active.Ready) { throw "License decision source not qualified: M365_Users_Active.csv: $($active.Source.Reason)" }
  $wantedUpns = @{}; $wantedImmutable = @{}
  foreach ($row in $active.Rows.Values) {
    $upn = ([string]$row.'User principal name').Trim().ToLowerInvariant(); if ($upn) { $wantedUpns[$upn] = $true }
    $immutable = ([string]$row.OnPremisesImmutableId).Trim().ToLowerInvariant(); if ($immutable) { $wantedImmutable[$immutable] = $true }
  }
  $adByImmutable = @{}; $adByUpn = @{}; $duplicateAdUpns = @{}
  if ($adSource.Ready) {
    try {
      foreach ($row in (Import-LicensesSourceCsv -Source $adSource)) {
        if ([string]$row.TenantKey -cne $TenantKey) { throw 'another tenant is present in AD users' }
        if ($adSource.PartialScopeAccepted -and ([string]$row.DomainName).Trim() -notin @($adSource.DomainCoverage.CollectedDomains)) { throw 'AD CSV row is outside the qualified collected domains.' }
        $immutable = ([string]$row.ImmutableId_AD).Trim().ToLowerInvariant(); $upn = ([string]$row.UserPrincipalName).Trim().ToLowerInvariant()
        if ($immutable -and $wantedImmutable.ContainsKey($immutable)) {
          if ($adByImmutable.ContainsKey($immutable)) { throw 'duplicate AD immutable ID' }
          $adByImmutable[$immutable] = $row
        }
        if ($upn -and $wantedUpns.ContainsKey($upn)) { if ($adByUpn.ContainsKey($upn)) { $duplicateAdUpns[$upn] = $true } else { $adByUpn[$upn] = $row } }
      }
    }
    catch { $adSource.Ready = $false; $adSource.Reason = $_.Exception.Message; $adByImmutable = @{}; $adByUpn = @{}; $duplicateAdUpns = @{} }
  }
  $mailboxes = Read-LicensesIndexedSource -Source $exoSource -ExpectedTenantKey $TenantKey -KeyColumn 'ExternalDirectoryObjectId' -WantedKeys $wanted -AsOfUtc $AsOfUtc
  $resourceTargets = @{}
  if ($phoneSource.Ready) {
    try {
      foreach ($row in (Import-LicensesSourceCsv -Source $phoneSource)) {
        if ([string]$row.TenantKey -cne $TenantKey) { throw 'another tenant is present in phone assignments' }
        if (([string]$row.TargetType) -match '(?i)resource') { $resourceTargets[([string]$row.AssignedPstnTargetId).Trim().ToLowerInvariant()] = $true }
      }
    }
    catch { $phoneSource.Ready = $false; $phoneSource.Reason = $_.Exception.Message; $resourceTargets = @{} }
  }
  $copilotState = Get-LicensesCopilotUsageState -Source $copilotSource -AsOfUtc $AsOfUtc -MinimumPeriodDays $Settings.CopilotInactivityDays

  # Suite candidates already counted in ReportCandidates are referenced, never recounted.
  $referenced = @{}
  foreach ($detail in $RecoveryDetails) { $referenced[('{0}|{1}' -f ([string]$detail.UserId).ToLowerInvariant(), [string]$detail.License)] = 'ReportCandidates:Recovery' }
  foreach ($detail in $DowngradeDetails) { $key = '{0}|Microsoft 365 E3' -f ([string]$detail.UserId).ToLowerInvariant(); if (-not $referenced.ContainsKey($key)) { $referenced[$key] = 'ReportCandidates:E3 to F3 review' } }

  $copilotParts = @('MICROSOFT_365_COPILOT','M365_COPILOT')
  $rows = [System.Collections.Generic.List[object]]::new()
  $stateByUser = @{}
  $copilotUnmeasured = 0; $copilotGrace = 0
  foreach ($userId in $skusByUser.Keys) {
    $activeRow = if ($active.Rows.ContainsKey($userId) -and -not $active.Duplicates.ContainsKey($userId)) { $active.Rows[$userId] } else { $null }
    $state = Get-LicensesAccountState -ActiveRow $activeRow -AdSource $adSource -AdByImmutable $adByImmutable -AdByUpn $adByUpn -DuplicateAdUpns $duplicateAdUpns
    $stateByUser[$userId] = $state
    $identity = $userIdentity[$userId]
    $mailbox = if ($mailboxes.Ready -and $mailboxes.Rows.ContainsKey($userId) -and -not $mailboxes.Duplicates.ContainsKey($userId)) { $mailboxes.Rows[$userId] } else { $null }
    $recipient = if ($mailbox) { ([string]$mailbox.RecipientTypeDetails).Trim() } else { '' }
    $resourceReason = if ($recipient -in @('RoomMailbox','EquipmentMailbox','SchedulingMailbox')) { "Resource account ($recipient, sign-in blocked by design)" }
      elseif ($resourceTargets.ContainsKey($userId)) { 'Resource account (Teams Phone resource account, sign-in blocked by design)' }
      elseif (@(foreach ($owned in $skusByUser[$userId]) { if (& $hasPlan $owned.SkuId @('MCOEV_VIRTUALUSER')) { $owned } }).Count -gt 0) { 'Resource account (Teams Phone resource account license, sign-in blocked by design)' }
      else { '' }
    foreach ($assignment in $skusByUser[$userId]) {
      $part = $assignment.SkuPartNumber.ToUpperInvariant()
      $suiteKey = "$userId|$(Get-LicensesSuiteProduct -SkuId $assignment.SkuId -SkuPartNumber $assignment.SkuPartNumber)"
      if ($Settings.ExcludedSkus.ContainsKey($part)) { continue }
      $tenantSku = $skuById[$assignment.SkuId]
      $base = [ordered]@{
        CandidateType=''; Recommendation=''; Reason=''; SkuId=$assignment.SkuId; SkuPartNumber=$assignment.SkuPartNumber
        SkuDisplayName=$(if ($tenantSku) { [string](Get-LicensesOptionalProperty $tenantSku 'TenantSkuDisplayName') } else { '' })
        UserId=$userId; UserPrincipalName=$(if ($activeRow) { [string]$activeRow.'User principal name' } elseif ($identity) { $identity.UserPrincipalName } else { '' })
        DisplayName=$(if ($identity) { $identity.DisplayName } else { '' })
        EntraAccountEnabled=$state.Entra; AdAccountEnabled=$state.Ad; AccountState=$state.State
        AssignmentRoutes=$assignment.Route; AssignedByGroupNames=$assignment.Groups
        AssignmentLastUpdatedDateTime=''; LastLicensedActivityDate=''; LastChatActivityDate=''; ActivityWindowEnd=''
        MailboxGuard='Not applicable'; ResourceAccount=$(if ($resourceReason) { 'Yes' } else { 'No' }); DependencyStatus='Not applicable'; DependentSkus=''; CountedIn=''
      }
      $dates = @(if ($pathDates.ContainsKey("$userId|$($assignment.SkuId)")) { $pathDates["$userId|$($assignment.SkuId)"] })
      $knownDates = @($dates | Where-Object { $null -ne $_ })
      if ($knownDates.Count -gt 0) { $base.AssignmentLastUpdatedDateTime = (@($knownDates | Sort-Object -Descending)[0]).ToString('o') }
      if ($state.State -in @('Disabled','Review')) {
        $base.CandidateType = 'Disabled account'
        $reasons = [System.Collections.Generic.List[string]]::new()
        $recommendation = 'Recommended'
        if ($state.State -eq 'Review') { $recommendation = 'Review'; $reasons.Add($state.Reason) } else { $reasons.Add($state.Reason) }
        if ($resourceReason) { $recommendation = 'Review'; $reasons.Add($resourceReason) }
        if (& $hasPlan $assignment.SkuId $exchangePlans) {
          $guard = Get-LicensesMailboxGuard -Mailbox $mailbox -SourceReady $mailboxes.Ready -ReviewSizeGB $Settings.MailboxReviewSizeGB
          $base.MailboxGuard = if ($guard.Status -eq 'Clear') { 'Clear' } else { "$($guard.Status): $($guard.Reason)" }
          if ($guard.Status -eq 'Blocked') { $recommendation = 'Blocked by dependency'; $reasons.Add($guard.Reason) }
          elseif ($guard.Status -eq 'Unknown' -and $recommendation -ne 'Blocked by dependency') { $recommendation = 'Review'; $reasons.Add($guard.Reason) }
        }
        $countedIn = $referenced[$suiteKey]
        if ($countedIn) { $base.CountedIn = $countedIn; $recommendation = 'Referenced'; $reasons.Insert(0, "Counted in $countedIn") }
        $base.Recommendation = $recommendation; $base.Reason = ($reasons | Where-Object { $_ }) -join '; '
        $rows.Add([pscustomobject]$base)
        continue
      }
      $countedIn = $referenced[$suiteKey]
      if ($state.State -eq 'Enabled' -and $countedIn) {
        $base.CandidateType = 'Report candidate'; $base.Recommendation = 'Referenced'; $base.CountedIn = $countedIn; $base.Reason = "Counted in $countedIn"
        $rows.Add([pscustomobject]$base)
        continue
      }
      if ($state.State -ne 'Enabled' -or $part -notin $copilotParts) { continue }
      if (-not $copilotState.Measured) { $copilotUnmeasured++; continue }
      $upn = ([string]$base.UserPrincipalName).Trim().ToLowerInvariant()
      if (-not $upn -or -not $copilotState.Rows.ContainsKey($upn)) { $copilotUnmeasured++; continue }
      $activity = Get-LicensesCopilotActivity -Row $copilotState.Rows[$upn]
      $windowEnd = $copilotState.WindowEnd; $cutoff = $windowEnd.AddDays(-$Settings.CopilotInactivityDays)
      $base.ActivityWindowEnd = $windowEnd.ToString('yyyy-MM-dd')
      $base.LastLicensedActivityDate = if ($activity.Licensed) { $activity.Licensed.ToString('yyyy-MM-dd') } else { '' }
      $base.LastChatActivityDate = if ($activity.Chat) { $activity.Chat.ToString('yyyy-MM-dd') } else { '' }
      if ($activity.Licensed -and $activity.Licensed -ge $cutoff) { continue }
      $recommendation = 'Recommended'
      $reasons = [System.Collections.Generic.List[string]]::new()
      if ($knownDates.Count -eq 0 -or $knownDates.Count -ne $dates.Count) { $recommendation = 'Review'; $reasons.Add('Assignment date unknown') }
      else {
        # LastUpdatedDateTime is a lower bound of the assignment age; a recent update only delays the candidate.
        $latest = @($knownDates | Sort-Object -Descending)[0]
        if (($AsOfUtc - $latest).TotalDays -lt $Settings.CopilotAssignmentGraceDays) { $copilotGrace++; continue }
      }
      if ($activity.Chat -and $activity.Chat -ge $cutoff) { $recommendation = 'Review'; $reasons.Insert(0, 'Copilot Chat only') }
      else { $reasons.Insert(0, ('No Microsoft 365 Copilot activity in {0} days' -f $Settings.CopilotInactivityDays)) }
      $base.CandidateType = 'Copilot without activity'; $base.Recommendation = $recommendation; $base.Reason = $reasons -join '; '
      $rows.Add([pscustomobject]$base)
    }
  }

  # Dependency guard: a qualifying license is not removable while a dependent license on the same user would keep no qualifying license.
  $byUserSku = @{}
  $removableByUser = @{}
  foreach ($row in $rows) {
    $byUserSku["$($row.UserId)|$($row.SkuPartNumber.ToUpperInvariant())"] = $row
    if ($row.Recommendation -in @('Recommended','Referenced')) {
      if (-not $removableByUser.ContainsKey($row.UserId)) { $removableByUser[$row.UserId] = [System.Collections.Generic.List[string]]::new() }
      $removableByUser[$row.UserId].Add($row.SkuPartNumber.ToUpperInvariant())
    }
  }
  foreach ($row in $rows) {
    if ($row.Recommendation -notin @('Recommended','Referenced')) { continue }
    $part = $row.SkuPartNumber.ToUpperInvariant()
    $owned = @(foreach ($assignment in $skusByUser[$row.UserId]) { $assignment.SkuPartNumber.ToUpperInvariant() })
    $removable = @($removableByUser[$row.UserId])
    $blockedBy = [System.Collections.Generic.List[string]]::new(); $together = [System.Collections.Generic.List[string]]::new()
    foreach ($rule in $DependencyRules) {
      if ($part -notin $rule.Qualifying) { continue }
      # An E3 to F3 review keeps F3, which still qualifies when the rule lists it.
      if ($row.CountedIn -eq 'ReportCandidates:E3 to F3 review' -and 'SPE_F1' -in $rule.Qualifying) { continue }
      foreach ($dependent in @($owned | Where-Object { $_ -in $rule.Dependents })) {
        $remaining = @($owned | Where-Object { $_ -in $rule.Qualifying -and $_ -notin $removable })
        if ($remaining.Count -gt 0) { continue }
        $dependentRow = $byUserSku["$($row.UserId)|$dependent"]
        if ($dependentRow -and $dependentRow.Recommendation -in @('Recommended','Referenced')) { $together.Add($dependent) } else { $blockedBy.Add($dependent) }
      }
    }
    if ($blockedBy.Count -gt 0) {
      $row.DependencyStatus = 'Blocked by dependency'; $row.DependentSkus = (@($blockedBy | Sort-Object -Unique) -join ';')
      if ($row.Recommendation -eq 'Recommended') { $row.Recommendation = 'Blocked by dependency' }
      $row.Reason = (@($row.Reason, ('Dependent license keeps this license required: {0}' -f $row.DependentSkus)) | Where-Object { $_ }) -join '; '
    }
    elseif ($together.Count -gt 0) { $row.DependencyStatus = 'Dependents removed together'; $row.DependentSkus = (@($together | Sort-Object -Unique) -join ';') }
    else { $row.DependencyStatus = 'Clear' }
  }

  $dynamicsPattern = '^(DYN365_|DYNAMICS_365|D365_)'
  # One pass per collection: per-SKU counters instead of filtering every collection for every SKU.
  $tally = @{}
  foreach ($id in $skuById.Keys) { $tally[$id] = @{ Users=@{}; Disabled=0; Enabled=0; Recommended=0; Review=0; Blocked=0; Referenced=0 } }
  foreach ($assignment in $assignments.Values) {
    if (-not $tally.ContainsKey($assignment.SkuId)) { continue }
    $counter = $tally[$assignment.SkuId]; $counter.Users[$assignment.UserId] = $true
    $userState = $stateByUser[$assignment.UserId].State
    if ($userState -eq 'Disabled') { $counter.Disabled++ } elseif ($userState -eq 'Enabled') { $counter.Enabled++ }
  }
  foreach ($row in $rows) {
    if (-not $tally.ContainsKey($row.SkuId)) { continue }
    switch ($row.Recommendation) { 'Recommended' { $tally[$row.SkuId].Recommended++ } 'Review' { $tally[$row.SkuId].Review++ } 'Blocked by dependency' { $tally[$row.SkuId].Blocked++ } 'Referenced' { $tally[$row.SkuId].Referenced++ } }
  }
  $capacity = foreach ($id in ($skuById.Keys | Sort-Object)) {
    $tenantSku = $skuById[$id]; $part = ([string]$tenantSku.TenantSkuPartNumber).Trim(); $upper = $part.ToUpperInvariant()
    $counter = $tally[$id]
    $usage = 'Not applicable'; $usageReason = 'No usage criterion for this license'
    if ($upper -in $copilotParts) {
      if ($copilotState.Measured) { $usage = 'Measured'; $usageReason = ('Licensed Microsoft 365 Copilot activity, window ending {0}' -f $copilotState.WindowEnd.ToString('yyyy-MM-dd')) }
      else { $usage = 'Not measured'; $usageReason = $copilotState.Reason }
    }
    elseif ($upper -match $dynamicsPattern) { $usage = 'Not measured'; $usageReason = 'No Dynamics 365 product telemetry; active users are a signal only' }
    $enabled = [long]$tenantSku.TenantPrepaidEnabled; $consumed = [long]$tenantSku.TenantConsumedUnits
    [pscustomobject][ordered]@{
      SkuId=$id; SkuPartNumber=$part; SkuDisplayName=[string](Get-LicensesOptionalProperty $tenantSku 'TenantSkuDisplayName')
      ReportProduct=$(if ($suite = Get-LicensesSuiteProduct -SkuId $id -SkuPartNumber $part) { $suite } else { Get-LicensesOtherProduct -SkuPartNumber $part })
      Excluded=[bool]$Settings.ExcludedSkus.ContainsKey($upper); ExclusionReason=$(if ($Settings.ExcludedSkus.ContainsKey($upper)) { $Settings.ExcludedSkus[$upper] } else { '' })
      EnabledUnits=$enabled; WarningUnits=(Get-LicensesOptionalProperty $tenantSku 'TenantPrepaidWarning'); SuspendedUnits=(Get-LicensesOptionalProperty $tenantSku 'TenantPrepaidSuspended'); LockedOutUnits=(Get-LicensesOptionalProperty $tenantSku 'TenantPrepaidLockedOut')
      ConsumedUnits=$consumed; AvailableUnits=($enabled - $consumed)
      AssignedUsers=$counter.Users.Count
      DisabledAccountUsers=$counter.Disabled
      ActiveUsersUsageNotMeasured=$(if ($usage -eq 'Not measured') { $counter.Enabled } else { 0 })
      RecommendedCandidates=$counter.Recommended; ReviewCandidates=$counter.Review
      BlockedCandidates=$counter.Blocked; ReferencedCandidates=$counter.Referenced
      CopilotGraceUsers=$(if ($upper -in $copilotParts) { $copilotGrace } else { 0 })
      UsageEvidence=$usage; UsageEvidenceReason=$usageReason
    }
  }
  $excludedRows = @($rows | Where-Object { $Settings.ExcludedSkus.ContainsKey($_.SkuPartNumber.ToUpperInvariant()) }).Count
  if ($excludedRows -ne 0) { throw 'Excluded SKUs must not produce candidates.' }
  return [pscustomobject]@{ Candidates=$rows.ToArray(); Capacity=@($capacity); Sources=$sources.ToArray(); CopilotMeasured=$copilotState.Measured; CopilotReason=$copilotState.Reason }
}

function Publish-LicensesSkuReportCsv {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$SnapshotId, [Parameter(Mandatory)][string]$TenantKey, [Parameter(Mandatory)][string]$ReceiptRunId, [Parameter(Mandatory)]$Decisions)
  $candidateColumns = @('SnapshotId','TenantKey','ReceiptRunId','CandidateType','Recommendation','Reason','SkuId','SkuPartNumber','SkuDisplayName','UserId','UserPrincipalName','DisplayName','EntraAccountEnabled','AdAccountEnabled','AccountState','AssignmentRoutes','AssignedByGroupNames','AssignmentLastUpdatedDateTime','LastLicensedActivityDate','LastChatActivityDate','ActivityWindowEnd','MailboxGuard','ResourceAccount','DependencyStatus','DependentSkus','CountedIn')
  $capacityColumns = @('SnapshotId','TenantKey','ReceiptRunId','SkuId','SkuPartNumber','SkuDisplayName','ReportProduct','Excluded','ExclusionReason','EnabledUnits','WarningUnits','SuspendedUnits','LockedOutUnits','ConsumedUnits','AvailableUnits','AssignedUsers','DisabledAccountUsers','ActiveUsersUsageNotMeasured','RecommendedCandidates','ReviewCandidates','BlockedCandidates','ReferencedCandidates','CopilotGraceUsers','UsageEvidence','UsageEvidenceReason')
  $identity = [ordered]@{ SnapshotId=$SnapshotId; TenantKey=$TenantKey; ReceiptRunId=$ReceiptRunId }
  $paths = [System.Collections.Generic.List[string]]::new()
  foreach ($spec in @(
      @{ Name='M365_Licenses_ReportSkuCandidates.csv'; Columns=$candidateColumns; Rows=@($Decisions.Candidates) },
      @{ Name='M365_Licenses_ReportSkuCapacity.csv'; Columns=$capacityColumns; Rows=@($Decisions.Capacity) })) {
    $rows = @(foreach ($row in $spec.Rows) { $values = [ordered]@{} + $identity; foreach ($property in $row.PSObject.Properties) { $values[$property.Name] = $property.Value }; [pscustomobject]$values })
    $path = Get-LicensesPublicationPath -Path (Join-Path $Folder $spec.Name)
    $pending = "$path.$SnapshotId.pending"
    try {
      $lines = if ($rows.Count -gt 0) { @($rows | Select-Object -Property $spec.Columns | ConvertTo-Csv -NoTypeInformation) } else { @(($spec.Columns | ForEach-Object { '"{0}"' -f $_ }) -join ',') }
      $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($lines -join "`r`n") + "`r`n")
      $stream = [IO.File]::Open($pending, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
      try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
      [IO.File]::Move($pending, $path, $true)
      $readBack = @(Import-Csv -LiteralPath $path -ErrorAction Stop)
      if ($readBack.Count -ne $rows.Count -or @($readBack | Where-Object { $_.SnapshotId -cne $SnapshotId -or $_.ReceiptRunId -cne $ReceiptRunId }).Count -gt 0) {
        throw "Published license SKU decision CSV does not match its snapshot: $($spec.Name)"
      }
      $paths.Add($path)
    }
    finally { if (Test-Path -LiteralPath $pending -PathType Leaf) { Remove-Item -LiteralPath $pending -Force -ErrorAction SilentlyContinue } }
  }
  return $paths.ToArray()
}

function Publish-LicensesRunReport {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][string]$CollectedAtUtc,
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows,
    [switch]$ForceAdCsvAnalysis,
    [string]$DependencyRulesPath,
    [string]$ExclusionsPath
  )
  $result = [pscustomobject]@{ Published=$false; UploadSatisfied=$false; SnapshotId=''; ReceiptRunId=''; Reason='' }
  $snapshot = $null
  $skuDecisionPaths = @()
  try {
    # Licensing files are proven by the active receipt's own publication hashes; the receipt completes after the report.
    $publication = Get-SmartM365CurrentSourcePublication
    if (-not $publication -or $publication.Producer -cne 'SmartM365-Licences-Inventory.ps1' -or
        [IO.Path]::GetFullPath($Folder).TrimEnd([char[]]'\/') -ine $publication.SourceRoot) {
      throw 'No active licensing source receipt for this folder; the report cannot be qualified by the current run.'
    }
    $result.ReceiptRunId = [string]$publication.RunId
    $proof = [pscustomobject]@{ RunId=[string]$publication.RunId; Hashes=$publication.Hashes }
    $summaryRows = @(Get-LicensesFocusedSummaryRows -TenantRows $TenantRows)
    $additionalRows = @(Get-LicensesAdditionalOverviewRows -TenantRows $TenantRows)
    $snapshotTime = [datetimeoffset]::Parse($CollectedAtUtc, [Globalization.CultureInfo]::InvariantCulture)
    $usage = $null
    try {
      $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $Folder -ExpectedTenantKey $TenantKey -LicenseSnapshotUtc $snapshotTime -ForceAdCsvAnalysis:$ForceAdCsvAnalysis -CurrentRunPublication $proof
    }
    catch { WriteLog -Message ("Focused license usage metrics unavailable: {0}" -f $_.Exception.Message) 'WARNING' }
    $assignmentIssues = $null
    try {
      $assignmentIssues = Get-LicensesAssignmentIssues -Folder $Folder -TenantKey $TenantKey -TenantRows $TenantRows -LicenseSnapshotUtc $snapshotTime -CurrentRunPublication $proof
    }
    catch { WriteLog -Message ("License assignment issues unavailable: {0}" -f $_.Exception.Message) 'WARNING' }
    $snapshot = Publish-LicensesReportSnapshot -Folder $Folder -TenantKey $TenantKey -CollectedAtUtc $CollectedAtUtc -SummaryRows $summaryRows -AdditionalRows $additionalRows -Usage $usage -AssignmentIssues $assignmentIssues -ReceiptRun $publication
    foreach ($csvPath in @($snapshot.CsvPaths)) { Register-SmartM365SourceCsv -Path $csvPath }
    $result.SnapshotId = [string]$snapshot.Id
    $result.Published = $true
    WriteLog -Message ("License report snapshot {0} published; its CSV files are optional outputs of the licensing receipt." -f $snapshot.Id) 'INFO'
    # SKU decisions are additive: their failure never withdraws the published report.
    try {
      if (-not $ExclusionsPath) { $ExclusionsPath = Join-Path $PSScriptRoot 'SmartM365-LicenseExclusions.json.txt' }
      $settings = Get-LicensesDecisionSettings -Config $ScriptLocalConfig -ExclusionsPath $ExclusionsPath
      if (-not $DependencyRulesPath) { $DependencyRulesPath = Join-Path $PSScriptRoot 'SmartM365-LicenseDependencies.json.txt' }
      $rules = Read-LicensesDependencyRules -Path $DependencyRulesPath
      $decisions = Get-LicensesSkuDecisions -Folder $Folder -TenantKey $TenantKey -TenantRows $TenantRows -CurrentRunPublication $proof -RecoveryDetails @(if ($usage) { $usage.RecoveryDetails }) -DowngradeDetails @(if ($usage) { $usage.DowngradeDetails }) -Settings $settings -DependencyRules $rules
      $skuDecisionPaths = @(Publish-LicensesSkuReportCsv -Folder $Folder -SnapshotId $result.SnapshotId -TenantKey $TenantKey -ReceiptRunId $result.ReceiptRunId -Decisions $decisions)
      foreach ($csvPath in $skuDecisionPaths) { Register-SmartM365SourceCsv -Path $csvPath }
      if (-not $decisions.CopilotMeasured) { WriteLog -Message ("Copilot usage not measured: {0}. No Copilot candidate is produced." -f $decisions.CopilotReason) 'INFO' }
      WriteLog -Message ("License SKU decisions published with snapshot {0}: {1} candidate row(s), {2} SKU capacity row(s)." -f $result.SnapshotId, @($decisions.Candidates).Count, @($decisions.Capacity).Count) 'INFO'
    }
    catch {
      $skuDecisionPaths = @()
      $note = 'License SKU decisions not published in this run: M365_Licenses_ReportSkuCandidates.csv and M365_Licenses_ReportSkuCapacity.csv in DATA-LAST belong to an earlier run and are not recorded by this receipt. Reason: {0}' -f $_.Exception.Message
      WriteLog -Message ("{0} The license report is not affected." -f $note) 'WARNING'
      try { Add-SmartM365SourceRunNote -Message $note } catch { WriteLog -Message ("License SKU decision note could not be added to the run state: {0}" -f $_.Exception.Message) 'WARNING' }
    }
  }
  catch {
    $result.Reason = $_.Exception.Message
    $note = 'License report not republished in this run: M365_Licenses_ReportSnapshot.json.txt and M365_Licenses_Report*.csv in DATA-LAST belong to an earlier run and are not recorded by this receipt. Reason: {0}' -f $result.Reason
    WriteLog -Message ("{0} The licensing receipt is not affected; the summary email will not be sent." -f $note) 'WARNING'
    try { Add-SmartM365SourceRunNote -Message $note } catch { WriteLog -Message ("License report note could not be added to the run state: {0}" -f $_.Exception.Message) 'WARNING' }
    return $result
  }
  $reportPaths = @(@($snapshot.Path) + @($snapshot.CsvPaths) | ForEach-Object { Join-Path $Folder ([IO.Path]::GetFileName([string]$_)) })
  $skuPaths = @($skuDecisionPaths | ForEach-Object { Join-Path $Folder ([IO.Path]::GetFileName([string]$_)) })
  if ([string]$snapshot.Path -cne $reportPaths[0]) {
    # Staged: these copies reach DATA-LAST with the receipt; upload them after the completion.
    $result | Add-Member -NotePropertyName UploadDeferred -NotePropertyValue $true
    $result | Add-Member -NotePropertyName ReportPaths -NotePropertyValue $reportPaths
    $result | Add-Member -NotePropertyName SkuPaths -NotePropertyValue $skuPaths
    return $result
  }
  return (Invoke-LicensesReportUpload -Result $result -ReportPaths $reportPaths -SkuPaths $skuPaths)
}

function Invoke-LicensesReportUpload {
  param([Parameter(Mandatory)]$Result,[AllowEmptyCollection()][string[]]$ReportPaths=@(),[AllowEmptyCollection()][string[]]$SkuPaths=@())
  $uploadSetting = Get-Variable -Name EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
  if (-not ($uploadSetting -and [bool]$uploadSetting.Value)) { $Result.UploadSatisfied = $true; return $Result }
  $result = $Result
  try {
    foreach ($path in $ReportPaths) {
      if (-not (Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path)) { throw ("License report file was not uploaded to SharePoint: {0}" -f [IO.Path]::GetFileName($path)) }
    }
    $result.UploadSatisfied = $true
  }
  catch {
    $result.Reason = $_.Exception.Message
    $note = 'License report {0} is recorded by this receipt but was not uploaded to SharePoint. Reason: {1}' -f $result.SnapshotId, $result.Reason
    WriteLog -Message ("{0} The summary email will not be sent." -f $note) 'WARNING'
    try { Add-SmartM365SourceRunNote -Message $note } catch { WriteLog -Message ("License report note could not be added to the run state: {0}" -f $_.Exception.Message) 'WARNING' }
  }
  foreach ($path in $SkuPaths) {
    try { $uploaded = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path } catch { $uploaded = $false }
    if (-not $uploaded) {
      $note = 'License SKU decision file {0} is recorded by this receipt but was not uploaded to SharePoint.' -f [IO.Path]::GetFileName($path)
      WriteLog -Message $note 'WARNING'
      try { Add-SmartM365SourceRunNote -Message $note } catch { WriteLog -Message ("License SKU decision note could not be added to the run state: {0}" -f $_.Exception.Message) 'WARNING' }
    }
  }
  return $result
}

function ConvertFrom-LicensesJsonElement {
  param([Parameter(Mandatory)][System.Text.Json.JsonElement]$Element)
  # System.Text.Json keeps ISO date strings as strings, unlike ConvertFrom-Json.
  switch ($Element.ValueKind) {
    'Object' {
      $properties = [ordered]@{}
      foreach ($property in $Element.EnumerateObject()) { $properties[$property.Name] = ConvertFrom-LicensesJsonElement -Element $property.Value }
      return [pscustomobject]$properties
    }
    'Array' {
      $items = [System.Collections.Generic.List[object]]::new()
      foreach ($item in $Element.EnumerateArray()) { $items.Add((ConvertFrom-LicensesJsonElement -Element $item)) }
      return ,$items.ToArray()
    }
    'String' { return $Element.GetString() }
    'Number' {
      $integer = 0L
      if ($Element.TryGetInt64([ref]$integer)) { return $integer }
      return $Element.GetDecimal()
    }
    'True' { return $true }
    'False' { return $false }
    default { return $null }
  }
}

function ConvertTo-LicensesCountTable {
  param([AllowNull()]$Value)
  $table = @{}
  if ($null -eq $Value) { return $table }
  foreach ($property in $Value.PSObject.Properties) { $table[$property.Name] = $property.Value }
  return $table
}

function Read-LicensesPublishedReport {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey
  )
  $receiptPath = Join-Path -Path $Folder -ChildPath 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'Licensing source receipt is missing; the published license report cannot be qualified.' }
  $receiptHash = (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash
  $receipt = [IO.File]::ReadAllText($receiptPath) | ConvertFrom-Json -ErrorAction Stop
  if ([string]$receipt.Status -cne 'Completed' -or [string]$receipt.TenantKey -cne $TenantKey -or
      [string]$receipt.Producer -cne 'SmartM365-Licences-Inventory.ps1') {
    throw 'Licensing source receipt is not a completed receipt of this tenant.'
  }
  Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
  $snapshotPath = Join-Path -Path $Folder -ChildPath 'M365_Licenses_ReportSnapshot.json.txt'
  if (-not (Test-Path -LiteralPath $snapshotPath -PathType Leaf)) { throw 'Published license report snapshot is missing.' }
  $snapshotBytes = [IO.File]::ReadAllBytes($snapshotPath)
  $json = [System.Text.Json.JsonDocument]::Parse([System.ReadOnlyMemory[byte]]::new($snapshotBytes))
  try { $document = ConvertFrom-LicensesJsonElement -Element $json.RootElement }
  finally { $json.Dispose() }
  if ($document.SchemaVersion -ne 2 -or [string]$document.TenantKey -cne $TenantKey -or [string]$document.SnapshotId -notmatch '^[a-f0-9]{32}$' -or
      @($document.Products).Count -ne 4 -or @($document.OtherProducts).Count -ne 3 -or -not $document.PSObject.Properties['Rendering']) {
    throw 'Published license report snapshot identity or structure is not supported.'
  }
  if ($document.PSObject.Properties['ReceiptRun'] -and [string]$document.ReceiptRun.RunId -cne [string]$receipt.RunId) {
    throw 'Published license report snapshot belongs to another licensing receipt run.'
  }
  foreach ($name in @('M365_Licenses_ReportSummary.csv','M365_Licenses_ReportCandidates.csv','M365_Licenses_ReportAssignmentIssues.csv','M365_Licenses_ReportGaps.csv','M365_Licenses_ReportSources.csv')) {
    $records = @($receipt.Files | Where-Object { [string]$_.File -ceq $name })
    if ($records.Count -ne 1) { throw "License report CSV is not recorded by the current licensing receipt: $name" }
    $path = Join-Path -Path $Folder -ChildPath $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne [string]$records[0].SHA256) {
      throw "License report CSV differs from the current licensing receipt: $name"
    }
    foreach ($row in (Import-Csv -LiteralPath $path -ErrorAction Stop)) {
      if ([string]$row.TenantKey -cne $TenantKey -or [string]$row.SnapshotId -cne [string]$document.SnapshotId) { throw "License report CSV belongs to another snapshot: $name" }
    }
  }
  if ((Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash -cne $receiptHash -or
      [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($snapshotBytes)) -cne (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash) {
    throw 'Published license report changed while it was being read.'
  }

  $rendering = $document.Rendering
  $summaryRows = @(foreach ($product in @($document.Products)) {
    [pscustomobject]@{ Product=[string]$product.Product; Enabled=[long]$product.Enabled; Consumed=[long]$product.Consumed; Subscribed=[bool]$product.Subscribed }
  })
  $additionalRows = @(foreach ($product in @($document.OtherProducts)) {
    [pscustomobject]@{ Product=[string]$product.Product; Enabled=[long]$product.Enabled; Consumed=[long]$product.Consumed; Subscribed=[bool]$product.Subscribed }
  })
  $usage = $null
  if ($rendering.UsageAvailable) {
    $mailboxGap = $null
    if ($null -ne $document.MailboxGap) {
      $userMailboxes = $document.MailboxGap.UserMailboxes
      $userMailboxes | Add-Member -NotePropertyName EntraStateReason -NotePropertyValue ([string]$rendering.EntraStateReason) -Force
      $mailboxGap = [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$document.MailboxGap.NoUserMailbox }
    }
    $usage = [pscustomobject]@{
      Rows = @(foreach ($product in @($document.Products)) {
        [pscustomobject]@{ Product=[string]$product.Product; Counts=(ConvertTo-LicensesCountTable $product.Counts); Available=[bool]$product.UsageAvailable }
      })
      RecoveryDetails = @($document.RecoveryCandidates)
      DowngradeReview = $document.E3ToF3Review
      DowngradeDetails = @($document.DowngradeCandidates)
      Sources = @($rendering.UsageSources)
      SharedSourceReady = [bool]$rendering.SharedSourceReady
      IntuneSourceReady = [bool]$rendering.IntuneSourceReady
      AdSourceForced = [bool]$rendering.AdSourceForced
      MailboxGap = $mailboxGap
      AdGapActivity = $document.AdGapActivity
    }
  }
  $issues = $document.AssignmentIssues
  $assignmentIssues = $null
  # 'assignment issues not analyzed' marks an absent analysis in the snapshot, not an unavailable source.
  if ($issues.Available -or [string]$issues.Reason -cne 'assignment issues not analyzed') {
    $counts = @{}
    if ($issues.Available) { foreach ($property in $issues.Counts.PSObject.Properties) { $counts[$property.Name] = ConvertTo-LicensesCountTable $property.Value } }
    $assignmentIssues = [pscustomobject]@{ Available=[bool]$issues.Available; Reason=[string]$issues.Reason; Counts=$counts; Rows=@($issues.Rows) }
  }
  return [pscustomobject]@{
    SnapshotId = [string]$document.SnapshotId
    CollectedAtUtc = [string]$document.LicenseCollectedAtUtc
    SummaryRows = $summaryRows
    AdditionalRows = $additionalRows
    Usage = $usage
    AssignmentIssues = $assignmentIssues
  }
}

function Send-LicensesFocusedSummaryEmail {
  param(
    [Parameter(Mandatory)][string]$CsvFolderPath,
    [Parameter(Mandatory)][string]$ExpectedTenantKey,
    [Parameter(Mandatory)][string]$MailStatePath,
    [switch]$Manual,
    [switch]$ForceLicenseSummaryEmail
  )

  if (-not $Manual -and -not [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableLicenseSummaryEmail' -DefaultValue $true)) {
    WriteLog -Message 'Focused license summary email is disabled by configuration.' 'INFO'
    return $false
  }
  if (-not $Manual -and (Test-SmartM365MaxItemsMode)) {
    WriteLog -Message 'Focused license summary email skipped for a sampled inventory.' 'INFO'
    return $false
  }

  $attachmentPath = $null
  $mailGate = $null
  $sendStarted = $false
  try {
    $mailTo = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
    $mailFrom = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($mailTo) -or [string]::IsNullOrWhiteSpace($mailFrom)) {
      throw 'A report recipient (To) and sender (From) are required for the focused license summary email.'
    }
    if ([string]::IsNullOrWhiteSpace($ExpectedTenantKey)) { throw 'Tenant key is required for the daily license summary mail limit.' }
    $mailGate = Enter-LicensesDailyMailGate -Path $MailStatePath -TenantKey $ExpectedTenantKey -Force:$ForceLicenseSummaryEmail
    if ($mailGate.Skip) {
      WriteLog -Message ("License summary email skipped: {0} (Europe/Paris). Use -ForceLicenseSummaryEmail to send again." -f $mailGate.Reason) 'INFO'
      return $false
    }

    # The email never recomputes: it renders the report already published and qualified by the receipt.
    $publishedReport = Read-LicensesPublishedReport -Folder $CsvFolderPath -TenantKey $ExpectedTenantKey
    $summaryRows = @($publishedReport.SummaryRows)
    $additionalRows = @($publishedReport.AdditionalRows)
    $usage = $publishedReport.Usage
    $assignmentIssues = $publishedReport.AssignmentIssues
    $CollectedAtUtc = $publishedReport.CollectedAtUtc
    $reportSnapshot = [pscustomobject]@{ Id=$publishedReport.SnapshotId }
    $usageByProduct = @{}
    if ($usage) { foreach ($usageRow in $usage.Rows) { $usageByProduct[$usageRow.Product] = $usageRow } }
    $recoveryByProduct = @{}
    foreach ($name in @('Microsoft 365 E5','Microsoft 365 E3')) {
      $metric = if ($usageByProduct.ContainsKey($name)) { $usageByProduct[$name] } else { $null }
      $counts = if ($metric) { $metric.Counts } else { $null }
      $recoveryByProduct[$name] = Format-LicensesMetric -Counts $counts -ValueName 'RecoveryCandidates' -UnknownName 'RecoveryUnknown' -Available ($metric -and $metric.Available)
    }
    $topRecoveryE5 = $recoveryByProduct['Microsoft 365 E5']
    $topRecoveryE3 = $recoveryByProduct['Microsoft 365 E3']
    $f3F1Counts = @{ Assigned=0; RecoveryCandidates=0; RecoveryUnknown=0 }
    $f3F1Available = $true
    foreach ($name in @('Microsoft 365 F3','Microsoft 365 F1')) {
      if (-not $usageByProduct.ContainsKey($name) -or -not $usageByProduct[$name].Available) { $f3F1Available = $false; continue }
      foreach ($key in @('Assigned','RecoveryCandidates','RecoveryUnknown')) { $f3F1Counts[$key] += [long]$usageByProduct[$name].Counts[$key] }
    }
    $topRecoveryF3F1 = Format-LicensesMetric -Counts $f3F1Counts -ValueName 'RecoveryCandidates' -UnknownName 'RecoveryUnknown' -Available $f3F1Available
    $licenseAlerts = Get-LicensesAlertSummary -AssignmentIssues $assignmentIssues
    $suiteColors = @('#0f766e','#2563eb','#6d28d9','#475569')
    $suiteCards = for ($index=0; $index -lt $summaryRows.Count; $index++) {
      New-LicensesOverviewCardHtml -Row $summaryRows[$index] -Width 25 -Accent $suiteColors[$index]
    }
    $otherColors = @('#d97706','#7c3aed','#0284c7')
    $otherCards = for ($index=0; $index -lt $additionalRows.Count; $index++) {
      $width = if ($index -eq 2) { 34 } else { 33 }
      New-LicensesOverviewCardHtml -Row $additionalRows[$index] -Width $width -Accent $otherColors[$index]
    }
    $percentCulture = [Globalization.CultureInfo]::InvariantCulture
    $capacityRows = @()
    $assignmentRows = @()
    $recoveryRows = @()
    $activityRows = @()
    foreach ($row in $summaryRows) {
      $status = if ($row.Subscribed) { 'Subscribed' } else { 'Not subscribed' }
      $metrics = if ($usageByProduct.ContainsKey($row.Product)) { $usageByProduct[$row.Product] } else { $null }
      $counts = if ($metrics) { $metrics.Counts } else { $null }
      $assigned = if ($metrics -and $metrics.Available) { [string]$counts.Assigned } else { 'N/D' }
      $multiple = Format-LicensesMetric -Counts $counts -ValueName 'Multiple' -UnknownName 'MultipleUnknown' -Available ($metrics -and $metrics.Available)
      $multipleAll = Format-LicensesMetric -Counts $counts -ValueName 'MultipleAll' -UnknownName 'MultipleAllUnknown' -Available ($metrics -and $metrics.Available)
      $disabled = Format-LicensesMetric -Counts $counts -ValueName 'Disabled' -UnknownName 'DisabledUnknown' -Available ($metrics -and $metrics.Available)
      $adEntra = Format-LicensesMetric -Counts $counts -ValueName 'AdEntraInactive' -UnknownName 'AdEntraUnknown' -Available ($metrics -and $metrics.Available)
      $mailbox = Format-LicensesMetric -Counts $counts -ValueName 'MailboxInactive' -UnknownName 'MailboxUnknown' -Available ($metrics -and $metrics.Available)
      $m365 = Format-LicensesMetric -Counts $counts -ValueName 'M365Inactive' -UnknownName 'M365Unknown' -Available ($metrics -and $metrics.Available)
      $apps = if ($row.Product -in @('Microsoft 365 E3','Microsoft 365 E5')) {
        Format-LicensesMetric -Counts $counts -ValueName 'LocalAppsInactive' -UnknownName 'LocalAppsUnknown' -Available ($metrics -and $metrics.Available)
      } else { 'N/A' }
      $recovery = Format-LicensesMetric -Counts $counts -ValueName 'RecoveryCandidates' -UnknownName 'RecoveryUnknown' -Available ($metrics -and $metrics.Available)
      $recoveryPrimaryPc = Format-LicensesMetric -Counts $counts -ValueName 'RecoveryPrimaryPc' -UnknownName 'RecoveryPrimaryPcUnknown' -Available ($usage -and $usage.IntuneSourceReady -and $metrics -and $metrics.Available)
      $availableUnits = [long]$row.Enabled - [long]$row.Consumed
      $usedPercent = if ([long]$row.Enabled -gt 0) { (([decimal]$row.Consumed * 100 / [decimal]$row.Enabled).ToString('0.#', $percentCulture) + '%') } else { 'N/A' }
      $freePercent = if ([long]$row.Enabled -gt 0) { (([decimal]$availableUnits * 100 / [decimal]$row.Enabled).ToString('0.#', $percentCulture) + '%') } else { 'N/A' }
      $consumedCapacity = '{0} ({1})' -f $row.Consumed, $usedPercent
      $availableCapacity = '{0} ({1})' -f $availableUnits, $freePercent
      $license = [System.Net.WebUtility]::HtmlEncode($row.Product.Replace('Microsoft 365 ',''))
      $cell = 'padding:10px 8px;border-bottom:1px solid #e2e8f0;font-size:12px;line-height:17px;color:#334155;vertical-align:top;'
      $lead = 'padding:10px 8px;border-bottom:1px solid #e2e8f0;font-size:12px;line-height:17px;color:#0f172a;font-weight:700;vertical-align:top;'
      $capacityRows += '<tr><td style="{0}">{1}</td><td style="{2}">{3}</td><td style="{2}">{4}</td><td style="{2}">{5}</td><td style="{2}">{6}</td><td style="{2}">{7}</td></tr>' -f $lead,$license,$cell,$row.Enabled,$consumedCapacity,$availableCapacity,$assigned,$status
      $capacityWait = if ($assignmentIssues -and $assignmentIssues.Available) { [string]$assignmentIssues.Counts[$row.Product].CapacityWait } else { 'N/D' }
      $otherErrors = if ($assignmentIssues -and $assignmentIssues.Available) { [string]$assignmentIssues.Counts[$row.Product].OtherErrors } else { 'N/D' }
      $assignmentRows += '<tr><td style="{0}">{1}</td><td style="{2}">{3}</td><td style="{2}">{4}</td></tr>' -f $lead,$license,$cell,$capacityWait,$otherErrors
      $recoveryRows += '<tr><td style="{0}">{1}</td><td style="{2}font-weight:700;color:#0f766e;">{3}</td><td style="{2}">{4}</td><td style="{2}">{5}</td><td style="{2}">{6}</td><td style="{2}">{7}</td></tr>' -f $lead,$license,$cell,$recovery,$disabled,$m365,$recoveryPrimaryPc,$multiple
      $activityRows += '<tr><td style="{0}">{1}</td><td style="{2}">{3}</td><td style="{2}">{4}</td><td style="{2}">{5}</td><td style="{2}">{6}</td></tr>' -f $lead,$license,$cell,$adEntra,$mailbox,$apps,$multipleAll
    }
    $sharedRows = foreach ($row in $summaryRows) {
      $metrics = if ($usageByProduct.ContainsKey($row.Product)) { $usageByProduct[$row.Product] } else { $null }
      $sharedReady = ($usage -and $usage.SharedSourceReady -and $metrics -and $metrics.Available)
      $licensed = if ($sharedReady) { [string]$metrics.Counts.SharedLicensed } else { 'N/D' }
      $under50 = if ($sharedReady) { [string]$metrics.Counts.SharedUnder50 } else { 'N/D' }
      $eligible = if ($sharedReady) { [string]$metrics.Counts.SharedEligible } else { 'N/D' }
      $unknown = if ($sharedReady) { [string]$metrics.Counts.SharedUnknown } else { 'N/D' }
      '<tr><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;font-weight:700;">{0}</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">{1}</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">{2}</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;font-weight:700;color:#0f766e;">{3}</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">{4}</td></tr>' -f `
         [System.Net.WebUtility]::HtmlEncode($row.Product.Replace('Microsoft 365 ','')), $licensed, $under50, $eligible, $unknown
    }
    $gap05 = if ($usage -and $usage.MailboxGap) { $usage.MailboxGap.UserMailboxes } else { $null }
    $gap06 = if ($usage -and $usage.MailboxGap) { $usage.MailboxGap.NoUserMailbox } else { $null }
    $gap05Total = if ($gap05 -and $gap05.Available) { [string]$gap05.Total } else { 'N/D' }
    $gap05Universe = if ($gap05 -and $gap05.Available) { [string]$gap05.Universe } else { 'N/D' }
    $gap05Percent = if ($gap05 -and $gap05.Available -and $gap05.Universe -gt 0) { (([decimal]$gap05.Total * 100 / [decimal]$gap05.Universe).ToString('0.#', $percentCulture) + '%') } elseif ($gap05 -and $gap05.Available) { 'N/A' } else { 'N/D' }
    $gap05Other = if ($gap05 -and $gap05.Available) { [string]$gap05.OtherSkus } else { 'N/D' }
    $gap05None = if ($gap05 -and $gap05.Available) { [string]$gap05.NoSkus } else { 'N/D' }
    $gap05Unknown = if ($gap05 -and $gap05.Available) { [string]$gap05.Unknown } else { 'N/D' }
    $gap05EntraReady = ($gap05 -and $gap05.Available -and $gap05.EntraStateAvailable)
    $gap05EntraEnabled = if ($gap05EntraReady) { [string]$gap05.EntraEnabled } else { 'N/D' }
    $gap05EntraDisabled = if ($gap05EntraReady) { [string]$gap05.EntraDisabled } else { 'N/D' }
    $gap05EntraUnknown = if ($gap05EntraReady) { [string]$gap05.EntraStateUnknown } else { 'N/D' }
    $gap06Total = if ($gap06 -and $gap06.Available) { [string]$gap06.Total } else { 'N/D' }
    $gap06Universe = if ($gap06 -and $gap06.Available) { [string]$gap06.Universe } else { 'N/D' }
    $gap06Percent = if ($gap06 -and $gap06.Available -and $gap06.Universe -gt 0) { (([decimal]$gap06.Total * 100 / [decimal]$gap06.Universe).ToString('0.#', $percentCulture) + '%') } elseif ($gap06 -and $gap06.Available) { 'N/A' } else { 'N/D' }
    $gap06Other = if ($gap06 -and $gap06.Available) { [string]$gap06.OtherSkus } else { 'N/D' }
    $gap06None = if ($gap06 -and $gap06.Available) { [string]$gap06.NoSkus } else { 'N/D' }
    $gap06Guests = if ($gap06 -and $gap06.Available) { [string]$gap06.Guests } else { 'N/D' }
    $gap06MemberEnabled = if ($gap06 -and $gap06.Available) { [string]$gap06.MemberEnabled } else { 'N/D' }
    $gap06MemberDisabled = if ($gap06 -and $gap06.Available) { [string]$gap06.MemberDisabled } else { 'N/D' }
    $gap06Unknown = if ($gap06 -and $gap06.Available) { [string]$gap06.Unknown } else { 'N/D' }
    $adGap = if ($usage) { $usage.AdGapActivity } else { $null }
    $adGapReady = ($adGap -and $adGap.Available)
    $adGapMembers = if ($adGapReady) { [string]$adGap.Members } else { 'N/D' }
    $adGapObserved = if ($adGapReady) { [string]$adGap.AdObserved } else { 'N/D' }
    $adGapNoMatch = if ($adGapReady -and $null -ne $adGap.NoObservedMatch) { [string]$adGap.NoObservedMatch } else { 'N/D' }
    $adGapAmbiguous = if ($adGapReady) { [string]$adGap.Ambiguous } else { 'N/D' }
    $adGapRecent = if ($adGapReady) { [string]$adGap.AdEnabledRecent } else { 'N/D' }
    $adGapInactive = if ($adGapReady) { [string]$adGap.AdEnabledInactive } else { 'N/D' }
    $adGapNoDate = if ($adGapReady) { [string]$adGap.AdEnabledNoDate } else { 'N/D' }
    $adGapDisabled = if ($adGapReady) { [string]$adGap.AdDisabled } else { 'N/D' }
    $adGapStateUnknown = if ($adGapReady) { [string]$adGap.AdEnabledUnknown } else { 'N/D' }
    $adGapStatus = if ($adGapReady -and $adGap.Provisional) { 'Provisional AD source' } elseif ($adGapReady) { 'Qualified AD source' } else { 'AD source N/D' }
    $e3Review = if ($usage) { $usage.DowngradeReview } else { $null }
    $e3ReviewAssigned = if ($e3Review -and $e3Review.Available) { [string]$e3Review.Assigned } else { 'N/D' }
    $e3ReviewCandidates = if ($e3Review -and $e3Review.Available) { [string]$e3Review.Candidates } else { 'N/D' }
    $e3ReviewUnknown = if ($e3Review -and $e3Review.Available) { [string]$e3Review.Unknown } else { 'N/D' }
    $e3ReviewExcluded = if ($e3Review -and $e3Review.Available) { [string]$e3Review.Excluded } else { 'N/D' }
    $e3ReviewRecovery = if ($e3Review -and $e3Review.Available) { [string]$e3Review.RecoveryExcluded } else { 'N/D' }
    $gapNotes = @()
    foreach ($item in @(@{ Section='06'; Value=$gap05 }, @{ Section='07'; Value=$gap06 })) {
      if ($item.Value -and -not $item.Value.Available -and $item.Value.Reason) {
        $gapNotes += '<p style="margin:7px 0 0;font-size:11px;color:#9a3412;">Section {0} N/D: {1}</p>' -f `
          $item.Section, [System.Net.WebUtility]::HtmlEncode([string]$item.Value.Reason)
      }
    }
    if ($gap05 -and $gap05.Available -and -not $gap05.EntraStateAvailable -and $gap05.EntraStateReason) {
      $gapNotes += '<p style="margin:7px 0 0;font-size:11px;color:#9a3412;">Section 06 Entra state N/D: {0}</p>' -f `
        [System.Net.WebUtility]::HtmlEncode([string]$gap05.EntraStateReason)
    }
    if ($adGapReady -and $adGap.Provisional -and $adGap.Reason) {
      $gapNotes += '<p style="margin:7px 0 0;font-size:11px;color:#9a3412;">AD activity scope: {0}</p>' -f [System.Net.WebUtility]::HtmlEncode([string]$adGap.Reason)
    }
    $sourceRows = if ($usage) {
      foreach ($source in $usage.Sources) {
        $state = if ($source.Ready -and ($source.Forced -or ($source.PSObject.Properties['PartialScopeAccepted'] -and $source.PartialScopeAccepted))) { "PROVISIONAL ($($source.Date); $($source.Provenance))" }
                 elseif ($source.Ready) { "Ready ($($source.Date); $($source.Provenance))" }
                 else { "N/D ($($source.Reason))" }
        '<li>{0}: {1}</li>' -f [System.Net.WebUtility]::HtmlEncode($source.Name), [System.Net.WebUtility]::HtmlEncode($state)
      }
    } else { @('<li>Usage sources: N/D</li>') }
    $sourceNote = if ($Manual) { '<span style="color:#0f766e;font-weight:700;">Source: existing published CSV. No new inventory was run.</span>' } else { '<span style="color:#0f766e;font-weight:700;">Source: published inventory.</span>' }
    $adOverrideNote = if ($usage -and $usage.AdSourceForced) { '<div style="margin:0 0 16px;padding:11px 14px;background:#fff7ed;border-left:4px solid #d97706;font-size:12px;line-height:18px;color:#7c2d12;"><strong>Provisional AD/Entra indicator.</strong> The AD CSV was analyzed despite a rejected collector receipt. Confirm AD inventory completeness before using its inactivity counts for a license decision.</div>' } else { '' }
    $subject = 'Microsoft 365 license overview and recovery'
    $assignmentStatusNote = if ($assignmentIssues -and -not $assignmentIssues.Available -and $assignmentIssues.Reason) {
      '<p style="margin:5px 0 14px;font-size:11px;color:#b45309;">Assignment issues N/D: {0}</p>' -f [System.Net.WebUtility]::HtmlEncode([string]$assignmentIssues.Reason)
    } else { '' }
    $tableStyle = 'width:100%;border-collapse:collapse;table-layout:fixed;font-family:Segoe UI,Arial,sans-serif;font-size:12px;line-height:17px;color:#334155;'
    $headStyle = 'padding:9px 8px;background:#eaf1f8;border-bottom:2px solid #cbd5e1;text-align:left;font-size:11px;line-height:15px;color:#334155;vertical-align:bottom;'
    $bodyHtml = @"
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:700px;margin:0 auto;">
  <p style="margin:0 0 5px;font-size:11px;letter-spacing:1px;font-weight:700;color:#0f766e;">MICROSOFT 365 &middot; LICENSE SUMMARY</p>
  <h1 style="margin:0 0 8px;font-size:24px;line-height:30px;color:#0f172a;">License overview and recovery</h1>
  <p style="margin:0 0 18px;font-size:12px;line-height:18px;color:#64748b;">$([System.Net.WebUtility]::HtmlEncode([string]$OrgDomain)) &nbsp;&middot;&nbsp; Snapshot $([System.Net.WebUtility]::HtmlEncode($CollectedAtUtc)) UTC</p>
  $adOverrideNote
  <h2 style="margin:0 0 10px;font-size:18px;line-height:24px;color:#0f172a;">License overview</h2>
  <h3 style="margin:0 0 7px;font-size:14px;line-height:20px;color:#0f172a;">License alerts</h3>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#fff7ed;border:1px solid #fed7aa;">
    <tr>
      <td width="50%" style="width:50%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #d97706;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#b45309;">$($licenseAlerts.AwaitingCapacity)</div><div style="font-size:12px;line-height:17px;color:#334155;">Awaiting capacity</div><div style="margin-top:5px;font-size:11px;line-height:16px;color:#64748b;">$($licenseAlerts.AwaitingBreakdown)</div></td></tr></table></td>
      <td width="50%" style="width:50%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #dc2626;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#b91c1c;">$($licenseAlerts.LicensingErrors)</div><div style="font-size:12px;line-height:17px;color:#334155;">Licensing errors</div><div style="margin-top:5px;font-size:11px;line-height:16px;color:#64748b;">$($licenseAlerts.ErrorsBreakdown)</div></td></tr></table></td>
    </tr>
  </table>
  <p style="margin:7px 0 18px;font-size:11px;line-height:16px;color:#64748b;">Each alert total counts distinct users across F1/F3/E3/E5; the suite breakdown may count a user more than once. Alerts come from the collected assignment snapshot and are separate from recovery candidates. N/D means the assignment sources could not be qualified.</p>
  <h2 style="margin:14px 0 10px;font-size:18px;line-height:24px;color:#0f172a;">License recovery overview</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#eff6f8;border:1px solid #cbdfe2;">
    <tr>
      <td width="25%" style="width:25%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #7c3aed;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#6d28d9;">$topRecoveryE5</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates E5</div></td></tr></table></td>
      <td width="25%" style="width:25%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #2563eb;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#1d4ed8;">$topRecoveryE3</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates E3</div></td></tr></table></td>
      <td width="25%" style="width:25%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #0f766e;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#0f766e;">$topRecoveryF3F1</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates F3/F1</div></td></tr></table></td>
      <td width="25%" style="width:25%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #d97706;"><tr><td style="padding:12px 10px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#b45309;">$e3ReviewCandidates</div><div style="font-size:12px;line-height:17px;color:#334155;">E3 to F3 downgrade review</div></td></tr></table></td>
    </tr>
  </table>
  <p style="margin:8px 0 20px;font-size:11px;line-height:16px;color:#64748b;">Recovery cards count license assignments for each suite. F3/F1 adds both suite counts; a user with both may count twice. The E3 to F3 card counts separate manual review candidates, excludes users already in recovery and is not a recoverable license count. N/D indicates unqualified users.</p>
  <p style="margin:0 0 5px;font-size:11px;line-height:16px;font-weight:700;color:#475569;">MICROSOFT 365 SUITES</p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#eff6f8;border:1px solid #cbdfe2;"><tr>$($suiteCards -join "`n")</tr></table>
  <p style="margin:14px 0 5px;font-size:11px;line-height:16px;font-weight:700;color:#475569;">COPILOT, DYNAMICS 365 AND POWER BI</p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#f8f5f1;border:1px solid #e5ddd2;"><tr>$($otherCards -join "`n")</tr></table>
  <p style="margin:7px 0 17px;font-size:11px;line-height:16px;color:#64748b;">Enabled and used counts come from the tenant subscription snapshot. Used means consumed license units, not measured app activity. Percent used = used / enabled; N/A means no enabled units. Dynamics 365 and Power BI add SKU units, not distinct users. Copilot covers Microsoft 365 Copilot; Dynamics 365 excludes sandbox, trial and preview SKUs; Power BI covers Pro and Premium Per User, excluding free Standard.</p>
  <h2 style="margin:0 0 8px;font-size:16px;line-height:22px;color:#0f172a;">01 &nbsp; License capacity</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Enabled units</th><th style="$headStyle">Consumed (used %)</th><th style="$headStyle">Available (free %)</th><th style="$headStyle">Assigned users</th><th style="$headStyle">Status</th></tr></thead><tbody>$($capacityRows -join "`n")</tbody></table>
  <h3 style="margin:18px 0 7px;font-size:14px;line-height:20px;color:#0f172a;">License assignment issues</h3>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Awaiting capacity</th><th style="$headStyle">Other licensing errors</th></tr></thead><tbody>$($assignmentRows -join "`n")</tbody></table>
  <p style="margin:6px 0 16px;font-size:11px;line-height:16px;color:#64748b;">Awaiting capacity counts distinct users with a CountViolation assignment path and no effective F1/F3/E3/E5 suite of the same product. Other errors include different Graph assignment failures and failed paths where another route already grants the suite. The Excel tabs Awaiting licenses and Licensing errors list the affected users, error codes and assigning groups. These counts reflect the collected Graph snapshot and can differ from the live admin center.</p>
  $assignmentStatusNote
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">02 &nbsp; Recovery by license</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Recovery candidates</th><th style="$headStyle">Disabled users</th><th style="$headStyle">No M365 activity (90d)</th><th style="$headStyle">Candidates primary on Intune PC</th><th style="$headStyle">Multiple target suites</th></tr></thead><tbody>$($recoveryRows -join "`n")</tbody></table>
  <p style="margin:7px 0 0;font-size:11px;line-height:16px;color:#64748b;">The attached Excel workbook lists each qualified recovery candidate once per license, with the recovery reason, Intune primary-PC indicator, last AD logon date and last M365 activity date. Its Summary sheet reconciles with this table. Activity dates are sortable Excel dates; an empty date cell means no qualified date was observed. AD LastLogonDate is replicated and approximate. Dates are not populated for identified shared mailboxes.</p>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">03 &nbsp; E3 to F3 downgrade review</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">E3 assignments</th><th style="$headStyle">Review candidates</th><th style="$headStyle">N/D</th><th style="$headStyle">Excluded by checks</th><th style="$headStyle">Already in recovery</th></tr></thead><tbody><tr><td style="$cell">$e3ReviewAssigned</td><td style="$cell"><strong style="color:#0f766e;">$e3ReviewCandidates</strong></td><td style="$cell">$e3ReviewUnknown</td><td style="$cell">$e3ReviewExcluded</td><td style="$cell">$e3ReviewRecovery</td></tr></tbody></table>
  <p style="margin:7px 0 0;font-size:11px;line-height:16px;color:#64748b;">The attached Excel workbook has an E3 to F3 review tab listing every qualified review candidate and the values used by these checks. Review candidates are enabled E3 UserMailbox users with mailbox size below 2 GB, no active archive or hold, no Windows/Mac Apps use in the 180-day report, OneDrive storage below 2 GB, and no second F1/F3/E5 suite. Missing or unqualified data is N/D. This is a manual downgrade review, not a recoverable E3 license count: confirm frontline eligibility, required E3 features, device rights and the commercial contract before changing a license. Usage reports, including OneDrive, are used only when the completed UserActivity receipt lists them with the same hash.</p>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">04 &nbsp; Activity and overlap</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">No AD or Entra activity (90d)</th><th style="$headStyle">No mailbox activity (90d)</th><th style="$headStyle">No local Apps use (180d)</th><th style="$headStyle">Multiple assigned SKUs</th></tr></thead><tbody>$($activityRows -join "`n")</tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">05 &nbsp; Licensed shared mailboxes</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Shared mailboxes with target SKU</th><th style="$headStyle">Under 50 GB</th><th style="$headStyle">Removal candidates after archive and hold checks</th><th style="$headStyle">Not qualified</th></tr></thead><tbody>$($sharedRows -join "`n")</tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">06 &nbsp; User mailboxes without licence F1/F3/E3/E5</h2>
  <p style="margin:0 0 8px;padding:11px 13px;background:#eff6f8;border-left:4px solid #0f766e;font-size:12px;color:#334155;"><strong style="font-size:19px;color:#0f766e;">$gap05Total</strong> of $gap05Universe qualified EXO UserMailbox &nbsp;&middot;&nbsp; <strong>$gap05Percent</strong></p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">Entra enabled</th><th style="$headStyle">Entra disabled</th><th style="$headStyle">Entra state N/D</th><th style="$headStyle">With other SKUs</th><th style="$headStyle">With no assigned SKU</th><th style="$headStyle">Not qualified</th></tr></thead><tbody><tr><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05EntraEnabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05EntraDisabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05EntraUnknown</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05Other</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05None</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05Unknown</td></tr></tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">07 &nbsp; Without User mailboxes + without licence F1/F3/E3/E5</h2>
  <p style="margin:0 0 8px;padding:11px 13px;background:#eff6f8;border-left:4px solid #2563eb;font-size:12px;color:#334155;"><strong style="font-size:19px;color:#1d4ed8;">$gap06Total</strong> of $gap06Universe qualified Entra accounts &nbsp;&middot;&nbsp; <strong>$gap06Percent</strong></p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">Member enabled</th><th style="$headStyle">Member disabled</th><th style="$headStyle">Guests</th><th style="$headStyle">With other SKUs</th><th style="$headStyle">With no assigned SKU</th><th style="$headStyle">Not qualified</th></tr></thead><tbody><tr><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06MemberEnabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06MemberDisabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Guests</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Other</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06None</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Unknown</td></tr></tbody></table>
  <p style="margin:15px 0 6px;font-size:12px;font-weight:700;color:#334155;">AD accounts and activity among section 07 Members &nbsp;&middot;&nbsp; $adGapStatus</p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">Members</th><th style="$headStyle">AD accounts observed</th><th style="$headStyle">No AD match observed</th><th style="$headStyle">Ambiguous AD match</th></tr></thead><tbody><tr><td style="$cell">$adGapMembers</td><td style="$cell">$adGapObserved</td><td style="$cell">$adGapNoMatch</td><td style="$cell">$adGapAmbiguous</td></tr></tbody></table>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">AD enabled, active in 90d</th><th style="$headStyle">AD enabled, inactive 90d</th><th style="$headStyle">AD enabled, no logon date</th><th style="$headStyle">AD disabled</th><th style="$headStyle">AD state N/D</th></tr></thead><tbody><tr><td style="$cell">$adGapRecent</td><td style="$cell">$adGapInactive</td><td style="$cell">$adGapNoDate</td><td style="$cell">$adGapDisabled</td><td style="$cell">$adGapStateUnknown</td></tr></tbody></table>
  <p style="margin:5px 0 0;font-size:11px;line-height:16px;color:#64748b;">AD activity uses the replicated LastLogonDate and is approximate. No AD match means no unique match in the available AD export, not proof that no AD account exists. AD and Entra enabled states are separate.</p>
  $($gapNotes -join "`n")
  <div style="margin:22px 0 0;padding:14px 16px;background:#f8fafc;border-left:3px solid #94a3b8;font-size:11px;line-height:17px;color:#475569;">
    <strong style="color:#0f172a;">How to read this report</strong><br />
    Recovery candidates are distinct licensed users per product with a disabled non-shared account, no observed M365 activity in 90 days, or a qualifying shared mailbox under 50 GB. Disabled users and activity/overlap indicators exclude identified shared mailboxes; an unqualified mailbox type is N/D. M365 activity includes Exchange, OneDrive, SharePoint, Teams, Skype for Business and Yammer, plus qualified mailbox, email-action and Apps usage reports. Shared mailbox candidates exclude active archives and litigation or retention holds. The Intune PC column counts recovery candidates assigned as Primary User of a Windows device; this assignment does not prove recent PC use and does not change the recovery count. Having no primary Intune PC also does not prove that a license is unused. Review advanced compliance features, assignment path and the commercial contract before removing a license. Indicators overlap and must not be added together.<br /><br />
    F1 includes M365_F1 and M365_F1_COMM. Multiple assigned SKUs include add-ons, trials and free products. Multiple target suites count non-shared users assigned to at least two distinct F1/F3/E3/E5 suites; the two F1 SKU variants count as one suite. Neither count alone proves redundant seats. Local Apps usage applies only to E3/E5 and uses the available 180-day Windows/Mac report. Sections 06 and 07 count distinct EXO UserMailbox and Entra account IDs with none of the four target suites; other SKUs are allowed and shown separately. Section 06 percent uses qualified EXO UserMailbox as its denominator. Its Entra enabled, disabled and state N/D counts add up to the mailbox total when the Entra source is qualified; these are Entra account states, not AD states. Section 07 percent uses qualified Entra accounts after excluding identified shared, room and other technical mailbox accounts in EXO and Exchange on-premises; its Member enabled, Member disabled and Guest groups add up to the qualified total. On-premises mailboxes match by ObjectGUID/ImmutableId first, then by unique UPN; an unqualified on-premises source makes section 07 N/D. These are coverage indicators, not license recovery candidates. N/D means a source or identity cannot be qualified; sources older than 14 days are excluded.
  </div>
  <h2 style="margin:22px 0 8px;font-size:14px;line-height:20px;color:#334155;">Source freshness</h2>
  <ul style="margin:0;padding-left:18px;font-size:11px;line-height:18px;color:#64748b;">$($sourceRows -join "`n")</ul>
  <p style="margin:14px 0 0;font-size:11px;line-height:16px;color:#64748b;">$sourceNote</p>
</div>
"@
    $bodyRow = '<tr><td style="padding:18px 24px 22px 24px;font-size:13px;line-height:19px;color:#334155;">{0}</td></tr>' -f $bodyHtml
    $executionFooter = 'Host: {0} | Generated: {1} | Snapshot: {2}' -f [string]$env:COMPUTERNAME, (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'), $reportSnapshot.Id
    $mailBody = New-SmartM365EmailBody -Title $subject -Category 'SmartM365' -HostName '' -GeneratedAt '' -BodyHtml $bodyRow -Footer $executionFooter
    $attachmentPath = Join-Path ([System.IO.Path]::GetTempPath()) ("M365_Licenses_RecoveryCandidates_{0}_{1}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'),[guid]::NewGuid().ToString('N').Substring(0,8))
    [void](New-LicensesRecoveryWorkbook -Path $attachmentPath -SummaryRows $summaryRows -Usage $usage -AssignmentIssues $assignmentIssues -CollectedAtUtc $CollectedAtUtc)
    $sendStarted = $true
    [void](Send-SmartM365Mail -From $mailFrom -To $mailTo -Subject $subject -BodyHtml $mailBody -MailPurpose Report -Attachments @($attachmentPath) -AllowAttachments -SuppressAttachmentLinks)
    $sentState = [ordered]@{ TenantKey=$ExpectedTenantKey; LocalDate=$mailGate.Day; Status='Sent'; SentAtUtc=[datetimeoffset]::UtcNow.ToString('o') }
    Write-LicensesDailyMailState -Stream $mailGate.Stream -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($sentState | ConvertTo-Json -Compress)))
    WriteLog -Message 'Microsoft 365 license overview and recovery email sent with recovery candidate workbook.' 'INFO'
    return $true
  }
  catch {
    if ($mailGate -and -not $mailGate.Skip -and -not $sendStarted) {
      try { Write-LicensesDailyMailState -Stream $mailGate.Stream -Bytes $mailGate.PreviousBytes }
      catch { WriteLog -Message ("License summary mail state could not be restored: {0}" -f $_.Exception.Message) 'WARNING' }
    }
    if ($Manual) { throw }
    WriteLog -Message ("Focused license summary email failed: {0}" -f $_.Exception.Message) 'WARNING'
  }
  finally {
    if ($mailGate -and $mailGate.Stream) { $mailGate.Stream.Dispose() }
    if ($attachmentPath -and (Test-Path -LiteralPath $attachmentPath -PathType Leaf)) {
      try { Remove-Item -LiteralPath $attachmentPath -Force -ErrorAction Stop } catch { Write-Warning "Temporary mail attachment cleanup failed: $($_.Exception.Message)" }
    }
  }
}

function Read-LicensesTenantSnapshot {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$ExpectedTenantKey
  )

  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Published tenant license CSV not found: $Path" }
  $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
  if ($rows.Count -eq 0) { throw "Published tenant license CSV has no data rows: $Path" }
  $requiredColumns = @('TenantKey','TenantSkuPartNumber','TenantPrepaidEnabled','TenantConsumedUnits','CollectedAtUtc')
  $columns = @($rows[0].PSObject.Properties.Name)
  foreach ($column in $requiredColumns) {
    if ($column -notin $columns) { throw "Published tenant license CSV is missing '$column': $Path" }
  }

  $collectedAtUtc = [string]$rows[0].CollectedAtUtc
  $parsedTimestamp = [datetimeoffset]::MinValue
  if ($collectedAtUtc -notmatch '(Z|[+-][0-9]{2}:[0-9]{2})$' -or
      -not [datetimeoffset]::TryParse($collectedAtUtc, [ref]$parsedTimestamp)) {
    throw "Published tenant license CSV has an invalid CollectedAtUtc: $Path"
  }
  foreach ($row in $rows) {
    if ([string]$row.TenantKey -ine $ExpectedTenantKey) {
      throw "Published tenant license CSV contains a row for another tenant: $Path"
    }
    if ([string]::IsNullOrWhiteSpace([string]$row.TenantSkuPartNumber)) {
      throw "Published tenant license CSV contains a row without a SKU part number: $Path"
    }
    if ([string]$row.CollectedAtUtc -cne $collectedAtUtc) {
      throw "Published tenant license CSV contains mixed collection timestamps: $Path"
    }
  }
  return [pscustomobject]@{ Rows = $rows; CollectedAtUtc = $parsedTimestamp.ToUniversalTime().ToString('o') }
}

function To-GuidOrNull {
  param($Value)
  try {
    if ($null -eq $Value) { return $null }
    if ($Value -is [Guid]) { return $Value }
    if ("$Value".Trim()) { return [Guid]$Value }
  } catch { return $null }
  return $null
}

# SKUs of the license details that the user's assignment states do not cover.
function Get-LicensesUncoveredSkuId {
  param([Parameter(Mandatory)]$User, [AllowNull()][object[]]$LicenseDetails)
  $assigned = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($state in @($User.LicenseAssignmentStates)) { if ($null -ne $state -and $state.SkuId) { [void]$assigned.Add([string]$state.SkuId) } }
  @(@($LicenseDetails) | Where-Object { $null -ne $_ -and $_.SkuId -and -not $assigned.Contains([string]$_.SkuId) } | ForEach-Object { [string]$_.SkuId } | Sort-Object -Unique)
}

# Assignment paths and service-plan states of one user come from one snapshot. When the license details
# (read user by user) list a SKU that the user list (read at the start of the run) did not assign, the
# license changed in between: the user and the license details are read again and used for both files.
# An assignment is never inferred from license details; a gap after the fresh read fails the user.
function Resolve-LicensesUserSnapshot {
  param(
    [Parameter(Mandatory)]$User,
    [AllowNull()][object[]]$LicenseDetails,
    [Parameter(Mandatory)][scriptblock]$ReadUser,
    [Parameter(Mandatory)][scriptblock]$ReadLicenseDetails
  )
  $uncovered = @(Get-LicensesUncoveredSkuId -User $User -LicenseDetails $LicenseDetails)
  if ($uncovered.Count -eq 0) { return [pscustomobject]@{ User = $User; LicenseDetails = @($LicenseDetails); Refreshed = $false; UncoveredSkuIds = @() } }
  $freshUser = & $ReadUser
  if ($null -eq $freshUser -or [string]$freshUser.Id -ne [string]$User.Id) { throw "License assignment re-read returned no record for user '$($User.Id)'." }
  $freshDetails = @(& $ReadLicenseDetails)
  $stillUncovered = @(Get-LicensesUncoveredSkuId -User $freshUser -LicenseDetails $freshDetails)
  if ($stillUncovered.Count -gt 0) {
    throw ("License details of user '{0}' list SKU(s) {1} without a license assignment state after a fresh read; no assignment is inferred." -f $User.Id, ($stillUncovered -join ','))
  }
  [pscustomobject]@{ User = $freshUser; LicenseDetails = $freshDetails; Refreshed = $true; UncoveredSkuIds = $uncovered }
}

# Same rule as the workplace preparation: every user/SKU pair of the service-plan states has an assignment path.
function Assert-LicensesPlanStateCoverage {
  param(
    [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$AssignmentPaths,
    [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$PlanStatePairs
  )
  $pathPairs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($row in $AssignmentPaths) { [void]$pathPairs.Add(('{0}|{1}' -f $row.UserId, $row.SkuId)) }
  $missing = @($PlanStatePairs | Where-Object { -not $pathPairs.Contains($_) } | Sort-Object)
  if ($missing.Count -gt 0) {
    throw ("License collection is inconsistent: {0} user/SKU pair(s) have service-plan states without a license assignment path (first: {1}). Current CSV publication is stopped." -f $missing.Count, (@($missing | Select-Object -First 3) -join ', '))
  }
}

function ConvertTo-LicensesAssignmentPathRows {
  param([Parameter(Mandatory)]$User, [datetimeoffset]$CollectedAtUtc = [datetimeoffset]::UtcNow)
  $userId = [string]$User.Id
  if ([string]::IsNullOrWhiteSpace($userId)) { throw 'License assignment user lacks an immutable Id.' }
  $seen = @{}
  foreach ($state in @($User.LicenseAssignmentStates)) {
    if ($null -eq $state) { continue }
    $skuId = [string]$state.SkuId
    if ([string]::IsNullOrWhiteSpace($skuId)) { throw "License assignment state lacks SkuId for user '$userId'." }
    $groupId = [string]$state.AssignedByGroup
    $row = [pscustomobject][ordered]@{
      UserId = $userId
      SkuId = $skuId
      AssignedByGroupId = $groupId
      AssignmentRoute = if ([string]::IsNullOrWhiteSpace($groupId)) { 'Direct' } else { 'Group' }
      AssignmentState = [string]$state.State
      AssignmentError = [string]$state.Error
      DisabledPlanIds = (@($state.DisabledPlans | Where-Object { $null -ne $_ } | Sort-Object -Unique) -join ';')
      LastUpdatedDateTime = if ($null -eq $state.LastUpdatedDateTime) { $null } else { ([datetimeoffset]$state.LastUpdatedDateTime).ToUniversalTime().ToString('o') }
      CollectedAtUtc = $CollectedAtUtc.ToUniversalTime().ToString('o')
    }
    $key = "$userId|$skuId|$groupId"
    $signature = $row | ConvertTo-Json -Compress
    if ($seen.ContainsKey($key)) {
      if ($seen[$key] -cne $signature) { throw "Conflicting license assignment states for '$key'." }
      continue
    }
    $seen[$key] = $signature
    $row
  }
}

function ConvertTo-LicensesTenantSkuEvidence {
  param([Parameter(Mandatory)]$Sku)
  function Read-EvidenceProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
  }
  [pscustomobject]@{
    PartNumber = $Sku.SkuPartNumber
    PrepaidEnabled = Read-EvidenceProperty $Sku.PrepaidUnits 'Enabled'
    PrepaidWarning = Read-EvidenceProperty $Sku.PrepaidUnits 'Warning'
    PrepaidSuspended = Read-EvidenceProperty $Sku.PrepaidUnits 'Suspended'
    PrepaidLockedOut = Read-EvidenceProperty $Sku.PrepaidUnits 'LockedOut'
    ConsumedUnits = $Sku.ConsumedUnits
    CapabilityStatus = $Sku.CapabilityStatus
    AppliesTo = $Sku.AppliesTo
    SubscriptionIds = (@($Sku.SubscriptionIds) -join ';')
  }
}

# Group cache + aggregator (for Groups CSV)
$groupNameCache = @{}
function Get-GroupNameCached {
  param([Parameter(Mandatory)][string]$GroupId)
  if (-not $GroupId) { return $null }
  if ($groupNameCache.ContainsKey($GroupId)) { return $groupNameCache[$GroupId] }
  try { $g = Invoke-GraphWithRetry { Get-MgGroup -GroupId $GroupId -Property "displayName" }; $name = $g.DisplayName }
  catch { $name = $GroupId }
  $groupNameCache[$GroupId] = $name
  return $name
}

$groupAgg = @{}
function Add-GroupAgg {
  param([string]$GroupId,[Guid]$SkuId,[string]$UserId)
  if (-not $GroupId -or -not $SkuId -or -not $UserId) { return }
  if (-not $groupAgg.ContainsKey($GroupId)) {
    $groupAgg[$GroupId] = @{
      DisplayName = $null
      SkuIds      = New-Object System.Collections.Generic.HashSet[System.Guid]
      UserIds     = New-Object System.Collections.Generic.HashSet[string]
    }
  }
  [void]$groupAgg[$GroupId].SkuIds.Add($SkuId)
  [void]$groupAgg[$GroupId].UserIds.Add($UserId)
  if (-not $groupAgg[$GroupId].DisplayName) { $groupAgg[$GroupId].DisplayName = Get-GroupNameCached -GroupId $GroupId }
}

function Get-ServicePlanStateCode {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][bool]$IsEnabled,
    [AllowEmptyString()][string]$PlanStatus
  )

  $normalizedStatus = if ([string]::IsNullOrWhiteSpace($PlanStatus)) { "" } else { $PlanStatus.Trim() }
  if ($IsEnabled) {
    switch ($normalizedStatus) {
      "Success"             { return "A" }
      "PendingActivation"   { return "PA" }
      "PendingInput"        { return "PI" }
      "PendingProvisioning" { return "PP" }
      "Error"               { return "E" }
      default                { return ("EN:{0}" -f $normalizedStatus) }
    }
  }

  if ($normalizedStatus -eq "Disabled") {
    return "D"
  }
  return ("DIS:{0}" -f $normalizedStatus)
}

function ConvertFrom-ServicePlanStateCode {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$StateCode)

  switch ($StateCode) {
    "A"  { return [pscustomobject]@{ IsEnabled = $true;  PlanStatus = "Success" } }
    "D"  { return [pscustomobject]@{ IsEnabled = $false; PlanStatus = "Disabled" } }
    "PA" { return [pscustomobject]@{ IsEnabled = $true;  PlanStatus = "PendingActivation" } }
    "PI" { return [pscustomobject]@{ IsEnabled = $true;  PlanStatus = "PendingInput" } }
    "PP" { return [pscustomobject]@{ IsEnabled = $true;  PlanStatus = "PendingProvisioning" } }
    "E"  { return [pscustomobject]@{ IsEnabled = $true;  PlanStatus = "Error" } }
  }

  if ($StateCode.StartsWith("EN:", [System.StringComparison]::Ordinal)) {
    return [pscustomobject]@{ IsEnabled = $true; PlanStatus = $StateCode.Substring(3) }
  }
  if ($StateCode.StartsWith("DIS:", [System.StringComparison]::Ordinal)) {
    return [pscustomobject]@{ IsEnabled = $false; PlanStatus = $StateCode.Substring(4) }
  }

  throw "Unsupported service-plan StateCode '$StateCode'."
}

function ConvertTo-LicensesCsvField {
  [CmdletBinding()]
  param([AllowNull()]$Value)

  $text = if ($null -eq $Value) { '' } else { [string]$Value }
  if ($text -match "`r|`n") {
    throw "Service-plan CSV field contains a line break, which is not supported by the disk-backed writer."
  }
  return ('"{0}"' -f $text.Replace('"', '""'))
}

function New-LicensesServicePlanStateWriter {
  [CmdletBinding()]
  [OutputType([System.IO.StreamWriter])]
  param([Parameter(Mandatory)][string]$Path)

  $folder = Split-Path -Path $Path -Parent
  if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder)) {
    New-Item -Path $folder -ItemType Directory -Force -ErrorAction Stop | Out-Null
  }
  if (Test-Path -LiteralPath $Path) {
    Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
  }

  $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.UTF8Encoding]::new($false))
  $writer.WriteLine('"TenantKey","UserId","SkuId","PlanId","StateCode"')
  return $writer
}

function Write-LicensesServicePlanStateRow {
  param(
    [Parameter(Mandatory)][System.IO.StreamWriter]$Writer,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][string]$UserId,
    [Parameter(Mandatory)][string]$SkuId,
    [Parameter(Mandatory)][string]$PlanId,
    [Parameter(Mandatory)][string]$StateCode,
    [Parameter(Mandatory)][int64]$RowNumber
  )

  try {
    if (
      [string]::IsNullOrWhiteSpace($TenantKey) -or
      [string]::IsNullOrWhiteSpace($UserId) -or
      [string]::IsNullOrWhiteSpace($SkuId) -or
      [string]::IsNullOrWhiteSpace($PlanId) -or
      [string]::IsNullOrWhiteSpace($StateCode)
    ) {
      throw 'A critical compact service-plan field is empty.'
    }
    if (
      $TenantKey -match "`r|`n" -or
      $UserId -match "`r|`n" -or
      $SkuId -match "`r|`n" -or
      $PlanId -match "`r|`n" -or
      $StateCode -match "`r|`n"
    ) {
      throw 'A compact service-plan CSV field contains a line break.'
    }

    $Writer.WriteLine(
      ('"{0}","{1}","{2}","{3}","{4}"' -f
        $TenantKey.Replace('"', '""'),
        $UserId.Replace('"', '""'),
        $SkuId.Replace('"', '""'),
        $PlanId.Replace('"', '""'),
        $StateCode.Replace('"', '""'))
    )
    if (($RowNumber % 5000) -eq 0) { $Writer.Flush() }
  }
  catch {
    $exception = [System.IO.IOException]::new("Failed to write compact service-plan row ${RowNumber}: $($_.Exception.Message)", $_.Exception)
    $exception.Data['SmartM365FatalServicePlanStateWrite'] = $true
    throw $exception
  }
}

function Close-LicensesServicePlanStateWriter {
  [CmdletBinding()]
  param([AllowNull()][System.IO.StreamWriter]$Writer)

  if ($null -eq $Writer) { return }
  $Writer.Flush()
  $Writer.Dispose()
}

function Get-LicensesCsvDataRowCount {
  [CmdletBinding()]
  [OutputType([int64])]
  param([Parameter(Mandatory)][string]$Path)

  [int64]$lineCount = 0
  $reader = [System.IO.StreamReader]::new($Path)
  try {
    while ($null -ne $reader.ReadLine()) { $lineCount++ }
  }
  finally {
    $reader.Dispose()
  }
  if ($lineCount -eq 0) { return 0L }
  return ($lineCount - 1L)
}

function Assert-LicensesServicePlanStateCsvFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][int64]$ExpectedRows
  )

  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "Compact service-plan CSV was not created: $Path"
  }
  $expectedHeader = '"TenantKey","UserId","SkuId","PlanId","StateCode"'
  $reader = [System.IO.StreamReader]::new($Path)
  try {
    $actualHeader = $reader.ReadLine()
  }
  finally {
    $reader.Dispose()
  }
  if ($actualHeader -cne $expectedHeader) {
    throw "Compact service-plan CSV header mismatch. Expected '$expectedHeader', got '$actualHeader'."
  }

  $validationStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  $actualRows = Get-LicensesCsvDataRowCount -Path $Path
  $validationStopwatch.Stop()
  if ($actualRows -ne $ExpectedRows) {
    throw "Compact service-plan CSV row count mismatch. Expected $ExpectedRows, got $actualRows."
  }
  WriteLog -Message ("Streaming file validation passed for 'M365_Licenses_UserServicePlanStates'. Rows: {0:N0}; elapsed={1:N1} sec; working set={2:N1} MB." -f $actualRows, $validationStopwatch.Elapsed.TotalSeconds, ((Get-Process -Id $PID).WorkingSet64 / 1MB)) 'INFO'
}

function Copy-LicensesCsvAtomically {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$DestinationPath
  )

  $destinationFolder = Split-Path -Path $DestinationPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($destinationFolder) -and -not (Test-Path -LiteralPath $destinationFolder)) {
    New-Item -Path $destinationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
  }
  $stagingPath = "$DestinationPath.building-$PID"
  try {
    Copy-Item -LiteralPath $SourcePath -Destination $stagingPath -Force -ErrorAction Stop
    Move-Item -LiteralPath $stagingPath -Destination $DestinationPath -Force -ErrorAction Stop
  }
  finally {
    if (Test-Path -LiteralPath $stagingPath) {
      Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue
    }
  }
}

function Publish-LicensesServicePlanStateCsvFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$BuildingPath,
    [Parameter(Mandatory)][int64]$ExpectedRows,
    [Parameter(Mandatory)][string]$CurrentOutputPath,
    [Parameter(Mandatory)][string]$LatestOutputPath
  )

  Assert-LicensesServicePlanStateCsvFile -Path $BuildingPath -ExpectedRows $ExpectedRows

  $baseFileName = 'M365_Licenses_UserServicePlanStates'
  $runBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName $baseFileName
  $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $timestampedPath = Join-Path -Path $CurrentOutputPath -ChildPath "$runBaseFileName`_$timestamp.csv"
  $currentPath = Join-Path -Path $CurrentOutputPath -ChildPath "$runBaseFileName.csv"
  $latestPath = if ([string]::IsNullOrWhiteSpace($LatestOutputPath)) { $currentPath } else { Join-Path -Path $LatestOutputPath -ChildPath "$runBaseFileName.csv" }
  if (Test-SmartM365MaxItemsMode) {
    WriteLog -Message ("MaxItems test CSV publication active. CSV paths are suffixed with {0}; standard Power BI filenames are not updated." -f (Get-SmartM365MaxItemsSuffix)) 'WARNING'
  }

  Move-Item -LiteralPath $BuildingPath -Destination $timestampedPath -Force -ErrorAction Stop
  WriteLog -Message "CSV exported to: $timestampedPath" 'INFO'
  if (-not $global:csvGeneratedPaths) {
    $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  }
  [void]$global:csvGeneratedPaths.Add($timestampedPath)

  Copy-LicensesCsvAtomically -SourcePath $timestampedPath -DestinationPath $currentPath
  [void]$global:csvGeneratedPaths.Add($currentPath)
  WriteLog -Message "CSV copied to: $currentPath" 'INFO'

  $latestWritePath = $latestPath
  if ($latestPath -ine $currentPath) {
    # Shared atomic copy: Publishing fence, receipt registration and source staging when enabled.
    $latestWritePath = Get-LicensesPublicationPath -Path $latestPath
    Copy-SmartM365FileAtomically -SourcePath $timestampedPath -DestinationPath $latestWritePath -VerifyHash
    [void]$global:csvGeneratedPaths.Add($latestPath)
    WriteLog -Message $(if ($latestWritePath -cne $latestPath) { "CSV staged for global path: $latestPath" } else { "CSV copied to global path: $latestPath" }) 'INFO'
  }

  $global:csvFilePath1 = $timestampedPath
  $global:csvFilePath2 = $currentPath
  $global:csvFilePath3 = $latestPath

  if ($global:RetentionMaxCSV -gt 0) {
    Remove-SmartM365TimestampedFilesOlderThan -FolderPath $CurrentOutputPath -FilePattern "$runBaseFileName`_*.csv" -RetentionDays 7 -LogFile $global:logTextFile
  }

  if ($latestWritePath -cne $latestPath) { Add-SmartM365StagedPublicationAction -Path $latestPath -SharePointUpload }
  else { Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestPath | Out-Null }
  WriteLog -Message 'Automatic WeeklyHistory publication skipped for this CSV; the caller will publish the validated dataset group.' 'INFO'
  return [pscustomobject]@{
    TimestampedPath = $timestampedPath
    CurrentPath     = $currentPath
    PublishedPath   = $latestWritePath
    RowCount        = $ExpectedRows
  }
}

function New-LicensesDetailedServicePlanHistorySource {
  [CmdletBinding()]
  [OutputType([int64])]
  param(
    [Parameter(Mandatory)][string]$CompactCsvPath,
    [Parameter(Mandatory)][string]$DetailedCsvPath
  )

  $buildingPath = "$DetailedCsvPath.building"
  try {
    [int64]$detailCount = 0
    Import-Csv -LiteralPath $CompactCsvPath |
      ForEach-Object {
        $decodedState = ConvertFrom-ServicePlanStateCode -StateCode ([string]$_.StateCode)
        $detailCount++
        [pscustomobject][ordered]@{
          'TenantKey'       = $_.TenantKey
          'OrganizationKey' = $global:SmartM365OrganizationKey
          'EnvironmentKey'  = $global:SmartM365EnvironmentKey
          'TenantId'        = $global:SmartM365TenantId
          'UserId'          = $_.UserId
          'SkuId'           = $_.SkuId
          'PlanId'          = $_.PlanId
          'IsEnabled'       = $decodedState.IsEnabled
          'PlanStatus'      = $decodedState.PlanStatus
        }
      } |
      Export-Csv -LiteralPath $buildingPath -NoTypeInformation -Encoding UTF8

    Move-Item -LiteralPath $buildingPath -Destination $DetailedCsvPath -Force -ErrorAction Stop
    return $detailCount
  }
  finally {
    if (Test-Path -LiteralPath $buildingPath) {
      Remove-Item -LiteralPath $buildingPath -Force -ErrorAction SilentlyContinue
    }
  }
}

function Get-LicensesIsoWeekName {
  [CmdletBinding()]
  param([datetime]$Date = (Get-Date))

  $isoYear = [System.Globalization.ISOWeek]::GetYear($Date)
  $isoWeek = [System.Globalization.ISOWeek]::GetWeekOfYear($Date)
  return '{0}-W{1:00}' -f $isoYear, $isoWeek
}

function Remove-LegacyWeeklyServicePlanStateDuplicates {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$HistoryRootPath,
    [int64]$ExpectedCurrentWeekRows = 0
  )

  if (-not (Test-Path -LiteralPath $HistoryRootPath -PathType Container)) { return 0 }
  $expectedDetailedHeader = '"TenantKey","OrganizationKey","EnvironmentKey","TenantId","UserId","SkuId","PlanId","IsEnabled","PlanStatus"'
  $currentWeekName = Get-LicensesIsoWeekName
  $removedCount = 0
  $weekFolders = @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-W\d{2}$' })
  foreach ($weekFolder in $weekFolders) {
    $legacyPath = Join-Path -Path $weekFolder.FullName -ChildPath 'M365_Licenses_UserServicePlanStates.csv'
    $detailedPath = Join-Path -Path $weekFolder.FullName -ChildPath 'M365_Licenses_UserServicePlanStates_Detailed.csv'
    if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf) -or -not (Test-Path -LiteralPath $detailedPath -PathType Leaf)) { continue }

    $reader = [System.IO.StreamReader]::new($detailedPath)
    try { $detailedHeader = $reader.ReadLine() }
    finally { $reader.Dispose() }
    if ($detailedHeader -cne $expectedDetailedHeader -or (Get-Item -LiteralPath $detailedPath).Length -le $expectedDetailedHeader.Length) {
      WriteLog -Message ("Legacy WeeklyHistory duplicate preserved because the replacement detailed CSV is invalid: {0}" -f $detailedPath) 'WARNING'
      continue
    }

    $detailedRows = Get-LicensesCsvDataRowCount -Path $detailedPath
    $legacyRows = Get-LicensesCsvDataRowCount -Path $legacyPath
    $replacementCountValid = if ($weekFolder.Name -eq $currentWeekName -and $ExpectedCurrentWeekRows -gt 0) {
      $detailedRows -eq $ExpectedCurrentWeekRows
    }
    else {
      $detailedRows -eq $legacyRows
    }
    if (-not $replacementCountValid) {
      WriteLog -Message ("Legacy WeeklyHistory duplicate preserved because row counts do not prove a complete replacement. Week={0}; LegacyRows={1}; DetailedRows={2}; ExpectedCurrentRows={3}." -f $weekFolder.Name, $legacyRows, $detailedRows, $ExpectedCurrentWeekRows) 'WARNING'
      continue
    }

    # Resolve and validate before deleting any duplicate CSV. A corrupt preferred
    # manifest must not be replaced with freshly calculated historical metadata.
    $manifestPath = Resolve-SmartM365OwnedJsonPath -Path (Join-Path $weekFolder.FullName 'manifest.json') -Owner 'M365 licenses inventory' -Validate {
      param($document)
      if ($document.HistoryLabel -ne 'M365 licenses inventory' -or $document.Week -ne $weekFolder.Name) { throw 'Licensing weekly manifest owner/week mismatch.' }
    }
    $previousManifestDocument = if (Test-Path -LiteralPath $manifestPath) { Read-SmartM365JsonDocument $manifestPath } else { $null }

    if ($global:EnableSharePointUpload) {
      $removedFromSharePoint = Remove-SmartM365SharePointFile -LocalFilePath $legacyPath
      if (-not $removedFromSharePoint) {
        WriteLog -Message ("Legacy WeeklyHistory duplicate could not be confirmed as removed from SharePoint: {0}" -f $legacyPath) 'WARNING'
      }
    }

    Remove-Item -LiteralPath $legacyPath -Force -ErrorAction Stop
    $removedCount++
    WriteLog -Message ("Legacy WeeklyHistory service-plan duplicate removed: {0}" -f $legacyPath) 'INFO'

    $manifestObject = if ($previousManifestDocument) { $previousManifestDocument.Document } else { [pscustomobject][ordered]@{
      UpdatedAt       = (Get-Date).ToString('o')
      Week            = $weekFolder.Name
      HistoryLabel    = 'M365 licenses inventory'
      HistoryRootPath = $HistoryRootPath
      Files           = @(Get-ChildItem -LiteralPath $weekFolder.FullName -Filter '*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { $_.Name })
    } }
    $manifestObject | Add-Member -NotePropertyName Files -NotePropertyValue @(Get-ChildItem -LiteralPath $weekFolder.FullName -Filter '*.csv' -File -ErrorAction Stop | Sort-Object Name | ForEach-Object { $_.Name }) -Force
    $manifestObject | Add-Member -NotePropertyName UpdatedAt -NotePropertyValue (Get-Date).ToString('o') -Force
    $manifestContent = $manifestObject | ConvertTo-Json -Depth 7
    $expectedHash = if ($previousManifestDocument) { $previousManifestDocument.SHA256 } else { 'ABSENT' }
    $null = Write-SmartM365JsonBytesAtomically -Path $manifestPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($manifestContent)) -ExpectedSHA256 $expectedHash -Validate { param($document) if (-not $document.PSObject.Properties['Files']) { throw 'Licensing manifest Files missing.' } }
    if ($global:EnableSharePointUpload) {
      Invoke-SmartM365SharePointCsvUpload -LocalFilePath $manifestPath | Out-Null
    }
  }
  return $removedCount
}

function Publish-LicensesWeeklyHistory {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ServicePlanStateCsvPath,
    [Parameter(Mandatory)][int64]$ExpectedServicePlanRows,
    [Parameter(Mandatory)][string]$CurrentOutputPath,
    [Parameter(Mandatory)][string]$LatestOutputPath
  )

  if (Test-SmartM365MaxItemsMode) {
    WriteLog -Message 'Consolidated Licensing WeeklyHistory publication skipped during MaxItems validation.' 'INFO'
    return
  }

  $weeklyHistoryEnabled = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableWeeklyHistory' -DefaultValue $true)
  if (-not $weeklyHistoryEnabled) {
    WriteLog -Message 'Licensing WeeklyHistory publication is disabled by configuration.' 'INFO'
    return
  }

  $historyRootPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue '')
  if ([string]::IsNullOrWhiteSpace($historyRootPath)) {
    $historyRootPath = Join-Path -Path $CurrentOutputPath -ChildPath 'WeeklyHistory'
  }
  $retentionWeeks = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)

  $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $detailedSourcePath = Join-Path -Path $CurrentOutputPath -ChildPath "M365_Licenses_UserServicePlanStates_Detailed_$timestamp.csv"

  try {
    $detailCount = New-LicensesDetailedServicePlanHistorySource -CompactCsvPath $ServicePlanStateCsvPath -DetailedCsvPath $detailedSourcePath
    if ($detailCount -ne $ExpectedServicePlanRows) {
      throw "Detailed service-plan WeeklyHistory row count mismatch. Expected $ExpectedServicePlanRows, got $detailCount."
    }
    WriteLog -Message ("Detailed service-plan WeeklyHistory source created from the compact CSV: {0:N0} row(s); working set={1:N1} MB." -f $detailCount, ((Get-Process -Id $PID).WorkingSet64 / 1MB)) 'INFO'

    $latestNames = @(
      'M365_Licenses_Users.csv'
      'M365_Licenses_ServicePlans_Catalog.csv'
      'M365_Licenses_ServicePlans.csv'
      'M365_Licenses_Tenant.csv'
      'M365_Licenses_Groups.csv'
    )
    $weeklySourceFiles = New-Object System.Collections.Generic.List[string]
    foreach ($latestName in $latestNames) {
      $latestPath = Join-Path -Path $LatestOutputPath -ChildPath $latestName
      $runCopyPath = Get-LicensesPublicationPath -Path $latestPath
      if ($global:csvGeneratedPaths -and $global:csvGeneratedPaths.Contains($latestPath) -and (Test-Path -LiteralPath $runCopyPath -PathType Leaf)) {
        $weeklySourceFiles.Add($runCopyPath) | Out-Null
      }
    }
    $weeklySourceFiles.Add($detailedSourcePath) | Out-Null

    Save-SmartM365WeeklyInventoryHistory `
      -SourceFiles $weeklySourceFiles.ToArray() `
      -HistoryRootPath $historyRootPath `
      -RetentionWeeks $retentionWeeks `
      -HistoryLabel 'M365 licenses inventory' `
      -UploadChangedFilesOnly

    try {
      Remove-LegacyWeeklyServicePlanStateDuplicates -HistoryRootPath $historyRootPath -ExpectedCurrentWeekRows $detailCount | Out-Null
    }
    catch {
      WriteLog -Message ("Legacy WeeklyHistory service-plan duplicate cleanup failed after the current dataset was published: {0}" -f $_) 'WARNING'
    }
  }
  finally {
    if (Test-Path -LiteralPath $detailedSourcePath) {
      Remove-Item -LiteralPath $detailedSourcePath -Force -ErrorAction SilentlyContinue
    }
  }
}
# ==========================================================
# Main
# ==========================================================
$ScriptVersion = "1.58"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicensesCsvLogFolderPath' -DefaultValue $OutputPath
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$defaultLicenseSummaryMailStatePath = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'M365_Licenses_SummaryEmail_SendState.json.txt' } else { '' }
$licenseSummaryMailStatePath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicenseSummaryMailStatePath' -DefaultValue $defaultLicenseSummaryMailStatePath)
# Includes email-only mode; the workbook remains optional for raw CSV collection.
try { Initialize-SmartM365ImportExcel | Out-Null }
catch { Write-Warning ("Optional license workbook unavailable at startup: {0}" -f $_.Exception.Message) }
if ($SendLicenseSummaryEmailOnly) {
  if ($MaxItems -gt 0 -or $TopUsers -gt 0 -or $FastSample -or $InteractiveAuth) {
    throw '-SendLicenseSummaryEmailOnly cannot be combined with collection or sampling switches.'
  }
  if ($ForceAdCsvAnalysis) {
    throw '-ForceAdCsvAnalysis applies to a full collection run; email-only mode sends the license report already published by that run.'
  }
  if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { throw 'LatestCsvFolderPath is required for email-only mode.' }
  $expectedTenantKey = if ($global:SmartM365TenantKey) { [string]$global:SmartM365TenantKey } else { $Tenant }
  $global:AppId = [string]$AppId
  $global:TenantId = [string]$TenantId
  $global:Thumb = [string]$Thumb
  $global:Thumbprint = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumbprint' -DefaultValue $Thumb)
  $global:OrgDomain = [string]$OrgDomain
  $mailSent = Send-LicensesFocusedSummaryEmail -CsvFolderPath $LatestCsvFolderPath -ExpectedTenantKey $expectedTenantKey -MailStatePath $licenseSummaryMailStatePath -Manual -ForceLicenseSummaryEmail:$ForceLicenseSummaryEmail
  if ($mailSent) { Write-Host ("License summary email sent from the published license report: {0}" -f $LatestCsvFolderPath) }
  else { Write-Host 'License summary email skipped: already sent or pending today (Europe/Paris).' }
  return
}
$connectedGraphInThisRun = $false
$currentOperation = "Initialize script environment"
$usersProcessedCount = 0
$userLicenseRowCount = 0
$servicePlanRowCount = 0
$tenantSkuRowCount = 0
$groupRowCount = 0
$servicePlanStateWriter = $null
$servicePlanStateBuildingPath = ''
$servicePlanStateLatestPath = ''
$runFinalized = $false
$inventoryEnvironmentInitialized = $false
$transcriptStarted = $false

try {
  # ------------------------
  # Initialize script environment
  # ------------------------
  $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
  $inventoryEnvironmentInitialized = $true
  Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
  # DATA-LAST changes only when this run completes; a failed or running run leaves the previous set intact.
  Enable-SmartM365StagedSourcePublication
  Start-Transcript -Path $global:logTranscriptFile -Append
  $transcriptStarted = $true
  WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
  $OutputPath = $InitializeOutputPath
  Enable-SmartM365DailyCaptureRetention -OutputRoot $OutputPath
  WriteLog -Message "Starting $TaskName..."
  $currentOperation = "Load Microsoft Graph modules"
  Ensure-GraphModules

  # ------------------------
  # Connect to Microsoft Graph via SmartM365.Core / Connect-SmartM365CloudSession
  # ------------------------
  if ($Connect) {
    Write-Host "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
  } else {
    Write-Host "Disconnecting any existing Microsoft Graph session before connecting..." -ForegroundColor Cyan
  }

  $currentOperation = "Disconnect existing Microsoft Graph session"
  try {
    Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
  }
  catch {
    WriteLog -Message ("Existing Microsoft Graph session cleanup did not complete: {0}" -f $_.Exception.Message) "WARN"
  }

  $currentOperation = "Connect to Microsoft Graph"
  $connectParams = @{
    ExchangeOnline = $false
    Graph          = $true
    GraphScopes    = @("Directory.Read.All", "User.Read.All")
  }

  if (-not $InteractiveAuth) {
    # Default: app-only certificate authentication
    $connectParams.AppId        = $AppId
    $connectParams.Thumbprint   = $Thumb
    $connectParams.TenantId     = $TenantId
    $connectParams.Organization = $OrgDomain
    WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"
  } else {
    WriteLog -Message "Connecting to Microsoft Graph with interactive authentication." "INFO"
  }

  $connectResult = Connect-SmartM365CloudSession @connectParams

  if (-not $connectResult.GraphConnected) {
    throw "Failed to connect to Microsoft Graph."
  }

  $connectedGraphInThisRun = $connectResult.GraphConnected

  $currentOperation = "Run preflight checks"
  Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('Directory.Read.All','User.Read.All','Group.Read.All') -GraphProbeUris @(
    'https://graph.microsoft.com/v1.0/subscribedSkus',
    'https://graph.microsoft.com/v1.0/users?$top=1',
    'https://graph.microsoft.com/v1.0/groups?$top=1'
  ) -RequiredModules @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups'
  ) -RequiredCommands @(
    'Get-MgSubscribedSku',
    'Get-MgUser',
    'Get-MgUserLicenseDetail',
    'Get-MgGroup'
  ) | Out-Null
  # ---------------- Tenant SKUs ----------------
  $currentOperation = "Read tenant subscribed SKUs"
  WriteLog -Message "Reading tenant SubscribedSkus..."
  $subscribedSkus = Invoke-GraphWithRetry { Get-MgSubscribedSku -All }
  $skusCollectedAtUtc = [datetimeoffset]::UtcNow.ToString('o')
  $tenantSkuMap = @{}
  foreach ($s in $subscribedSkus) {
    $tenantSkuMap[$s.SkuId] = ConvertTo-LicensesTenantSkuEvidence -Sku $s
  }

  # ---------------- CSV mapping + enforcement ----------------
  $currentOperation = "Load SKU and service plan mapping CSV"
  $SkuMapByPart,$SkuMapById,$SvcMapByName,$SvcMapById = Load-SkuNameMap -Path $SkuNameCsvPath
  if ($RequireSkuNameCsv) {
    $csvExists = Test-Path -LiteralPath $SkuNameCsvPath
    $hasSkuMap = ($SkuMapByPart.Count -gt 0) -or ($SkuMapById.Count -gt 0)
    $hasSvcMap = ($SvcMapById.Count   -gt 0) # plan-friendly required for ServicePlans
    if (-not $csvExists) {
      Write-Error "RequireSkuNameCsv is set, but CSV not found: $SkuNameCsvPath"
      throw "Required SKU mapping CSV not found."
    }
    if (-not $hasSkuMap) {
      Write-Error "RequireSkuNameCsv is set, but CSV contains no usable SKU mappings."
      throw "Required SKU mapping CSV contains no SKU mappings."
    }
    if ($ServicePlans -and -not $hasSvcMap) {
      Write-Error "RequireSkuNameCsv is set, but CSV contains no usable Service Plan friendly mappings."
      throw "Required SKU mapping CSV contains no Service Plan mappings."
    }
  }

  # ---------------- Users (with LicenseAssignmentStates) ----------------
  $currentOperation = "Retrieve users from Microsoft Graph"
  WriteLog -Message "Retrieving users..."
  $usersAll = Invoke-GraphWithRetry {
    Get-MgUser -All -Property "DisplayName","UserPrincipalName","Id","AssignedLicenses","Mail","ProxyAddresses","LicenseAssignmentStates"
  }
  $licenseUsersCollectedAtUtc = [datetimeoffset]::UtcNow

  $currentOperation = 'Retrieve the complete Entra group inventory'
  $directoryGroups = @(Invoke-GraphWithRetry {
    Get-MgGroup -All -Property 'id','displayName','description','mail','mailEnabled','securityEnabled','groupTypes','membershipRule','membershipRuleProcessingState','onPremisesSyncEnabled','onPremisesSecurityIdentifier','createdDateTime' -ErrorAction Stop
  })
  $groupsCollectedAtUtc = [datetimeoffset]::UtcNow.ToString('o')
  $directoryGroupRows = @($directoryGroups | ForEach-Object {
    $groupNameCache[[string]$_.Id] = [string]$_.DisplayName
    [pscustomobject][ordered]@{
      GroupId = $_.Id; DisplayName = $_.DisplayName; Description = $_.Description
      Mail = $_.Mail; MailEnabled = $_.MailEnabled; SecurityEnabled = $_.SecurityEnabled
      GroupTypes = (@($_.GroupTypes) -join ';'); MembershipRule = $_.MembershipRule
      MembershipRuleProcessingState = $_.MembershipRuleProcessingState
      OnPremisesSyncEnabled = $_.OnPremisesSyncEnabled
      OnPremisesSecurityIdentifier = $_.OnPremisesSecurityIdentifier
      CreatedDateTime = $_.CreatedDateTime
      CollectedAtUtc = $groupsCollectedAtUtc
    }
  })

  # Sampling: prefer users with group-based licensing
  if ($TopUsers -gt 0) {
    # Legacy sampling must have the same canonical-file protection as MaxItems.
    $global:SmartM365MaxItems = $TopUsers
    $global:SmartM365TestMaxItems = $TopUsers
    $global:SmartM365IsMaxItemsRun = $true
    $groupBased=@(); $directOnly=@()
    foreach ($u in $usersAll) {
      $hasGroup = $false
      foreach ($st in (@($u.LicenseAssignmentStates) | ForEach-Object { $_ })) {
        if ($st.AssignedByGroup) { $hasGroup = $true; break }
      }
      if ($hasGroup) {
        $groupBased += $u
      } elseif ($u.AssignedLicenses -and $u.AssignedLicenses.Count -gt 0) {
        $directOnly += $u
      }
    }
    $users = New-Object System.Collections.Generic.List[object]
    foreach ($u in $groupBased) {
      if ($users.Count -lt $TopUsers) { $users.Add($u) | Out-Null }
    }
    if ($users.Count -lt $TopUsers) {
      foreach ($u in $directOnly) {
        if ($users.Count -lt $TopUsers) { $users.Add($u) | Out-Null } else { break }
      }
    }
    if ($users.Count -eq 0) {
      $users = $usersAll | Select-Object -First $TopUsers
    }
    WriteLog -Message "Users to process (sample): $($users.Count) (of $(@($usersAll).Count))"
  } else {
    $users = $usersAll
    WriteLog -Message "Users to process: $(@($users).Count)"
  }

  # ---------------- Build rows ----------------
  $currentOperation = "Resolve user licenses and service plans"
  WriteLog -Message "Resolving licenses and building rows..."
  $resultsUsers       = New-Object System.Collections.Generic.List[object]
  $assignmentPaths    = [System.Collections.Generic.List[object]]::new()
  $userResolution = @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = 0 }
  $usersDeletedQualification = $null
  $rowsPlansExchange  = New-Object System.Collections.Generic.List[object]
  $servicePlanCatalog = @{}
  # Tenant catalog includes plans in SKUs with no assigned users.
  foreach ($subscribedSku in $subscribedSkus) {
    foreach ($plan in @($subscribedSku.ServicePlans)) {
      $catalogKey = "{0}|{1}" -f $subscribedSku.SkuId, $plan.ServicePlanId
      $servicePlanCatalog[$catalogKey] = [pscustomobject][ordered]@{
        SkuId = $subscribedSku.SkuId; SkuPartNumber = $subscribedSku.SkuPartNumber
        'SKU name' = Get-SkuDisplayName -SkuId $subscribedSku.SkuId -SkuPartNumber $subscribedSku.SkuPartNumber -MapByPart $SkuMapByPart -MapById $SkuMapById
        PlanId = $plan.ServicePlanId; PlanName = $plan.ServicePlanName
        PlanDisplayName = Get-ServiceFriendly -PlanName $plan.ServicePlanName -PlanId $plan.ServicePlanId -ByName $null -ById $SvcMapById
        AppliesTo = $plan.AppliesTo; TenantProvisioningStatus = $plan.ProvisioningStatus
        CollectedAtUtc = $skusCollectedAtUtc
      }
    }
  }
  if ($ServicePlans) {
    $stateRunBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'M365_Licenses_UserServicePlanStates'
    $servicePlanStateBuildingPath = Join-Path -Path $OutputPath -ChildPath ("{0}_{1}.building.csv" -f $stateRunBaseFileName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $servicePlanStateWriter = New-LicensesServicePlanStateWriter -Path $servicePlanStateBuildingPath
  }

  $uTotal = ($users | Measure-Object).Count
  $uIndex = 0
  # User/SKU pairs written to the service-plan states, checked against the assignment paths before publication.
  $servicePlanStatePairs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $userSnapshotRefreshes = 0

  foreach ($u in $users) {
    $uIndex++
    $pct = if ($uTotal -gt 0) { [math]::Floor(($uIndex * 100.0) / $uTotal) } else { 100 }
    Write-Progress -Activity "Processing users" -Status "[$uIndex/$uTotal] $($u.UserPrincipalName)" -PercentComplete $pct

    $licenseDetailStage = $false
    $assignmentPathCountBefore = $assignmentPaths.Count
    try {
      $upn         = $u.UserPrincipalName
      $displayName = $u.DisplayName
      $uid         = $u.Id
      $primarySmtp = Get-PrimarySmtpAddress -User $u
      # Keep error/disabled paths even when Graph returns no effective license detail.
      foreach ($assignmentPath in @(ConvertTo-LicensesAssignmentPathRows -User $u -CollectedAtUtc $licenseUsersCollectedAtUtc)) {
        $assignmentPaths.Add($assignmentPath)
      }

      # Direct disabled plans per SKU (raw)
      $disabledPlansBySku=@{}
      foreach ($al in (@($u.AssignedLicenses) | ForEach-Object { $_ })) {
        if ($al.SkuId) { $disabledPlansBySku[$al.SkuId]=@($al.DisabledPlans) }
      }

      # License details (Direct + Group)
      $licenseDetailStage = $true
      $licenseDetails = Invoke-GraphWithRetry { Get-MgUserLicenseDetail -UserId $uid -ErrorAction Stop }
      # A license changed since the user list: paths and plans of this user come from one fresh read.
      $snapshot = Resolve-LicensesUserSnapshot -User $u -LicenseDetails @($licenseDetails) `
        -ReadUser { Invoke-GraphWithRetry { Get-MgUser -UserId $uid -Property "DisplayName","UserPrincipalName","Id","AssignedLicenses","Mail","ProxyAddresses","LicenseAssignmentStates" -ErrorAction Stop } } `
        -ReadLicenseDetails { Invoke-GraphWithRetry { Get-MgUserLicenseDetail -UserId $uid -ErrorAction Stop } }
      $licenseDetailStage = $false
      if ($snapshot.Refreshed) {
        $userSnapshotRefreshes++
        $u = $snapshot.User
        $licenseDetails = $snapshot.LicenseDetails
        while ($assignmentPaths.Count -gt $assignmentPathCountBefore) { $assignmentPaths.RemoveAt($assignmentPaths.Count - 1) }
        foreach ($assignmentPath in @(ConvertTo-LicensesAssignmentPathRows -User $u -CollectedAtUtc ([datetimeoffset]::UtcNow))) { $assignmentPaths.Add($assignmentPath) }
        $disabledPlansBySku = @{}
        foreach ($al in (@($u.AssignedLicenses) | ForEach-Object { $_ })) {
          if ($al.SkuId) { $disabledPlansBySku[$al.SkuId] = @($al.DisabledPlans) }
        }
        WriteLog -Message ("License assignment of user {0} changed during collection (SKU {1}); assignment paths and service plans use one fresh read." -f $uid, ($snapshot.UncoveredSkuIds -join ',')) 'WARNING'
      }
      if (-not $licenseDetails -or $licenseDetails.Count -eq 0) { continue }

      # SKU names
      $skuDisplayById=@{}; $skuPartById=@{}
      foreach ($ld in $licenseDetails) {
        $sku = $ld.SkuId
        $pn  = if ($tenantSkuMap.ContainsKey($sku)) { $tenantSkuMap[$sku].PartNumber } else { $ld.SkuPartNumber }
        $skuPartById[$sku]=$pn
        if (-not $skuDisplayById.ContainsKey($sku)) {
          $skuDisplayById[$sku] = Get-SkuDisplayName -SkuId $sku -SkuPartNumber $pn -MapByPart $SkuMapByPart -MapById $SkuMapById
        }
      }

      # For each SKU
      $statesAll = @($u.LicenseAssignmentStates)
      foreach ($ld in $licenseDetails) {
        $skuId = $ld.SkuId

        # Source + group names (for Users CSV)
        $states   = @($statesAll | Where-Object { $_.SkuId -eq $skuId })
        $groupIds = @($states | Where-Object { $_.AssignedByGroup } | Select-Object -ExpandProperty AssignedByGroup -Unique)
        $groupNames = @()
        foreach ($gid in $groupIds) {
          $name = Get-GroupNameCached -GroupId $gid
          if ($name) { $groupNames += $name }
          $skuGuid = To-GuidOrNull $skuId
          if ($skuGuid -and $gid) { Add-GroupAgg -GroupId $gid -SkuId $skuGuid -UserId $uid }
        }
        $hasDirectState = (@($states | Where-Object { -not $_.AssignedByGroup }).Count -gt 0)

        function Add-UsersRow {
          param([string]$SourceVal,[string]$GroupName,[bool]$HasBoth)
          $groupField = if ($GroupName) { $GroupName } else { "" }
          $groupCount = if ($GroupName) { 1 } else { 0 }
          $rowUsers = [ordered]@{
            "Id"                  = "$uid-$skuId"
            "User principal name" = $upn
            "primarysmtp"         = $primarySmtp
            "Display name"        = $displayName
            "UserId"              = $uid
            "SkuId"               = $skuId
            "SkuPartNumber"       = $skuPartById[$skuId]
            "SKU name"            = $skuDisplayById[$skuId]
            "Source"              = $SourceVal
            "GroupsAssigningSku"  = $groupField
            "GroupCountForSku"    = $groupCount
            "HasDirectAndGroup"   = $HasBoth
          }
          foreach ($k in @($rowUsers.Keys)) {
            if ($rowUsers[$k] -is [string]) {
              $rowUsers[$k] = $rowUsers[$k] -replace "`r`n|`n|`r"," "
              $rowUsers[$k] = $rowUsers[$k] -replace '"',"'" 
            }
          }
          $resultsUsers.Add([PSCustomObject]$rowUsers) | Out-Null
        }

        # Row-splitting for Users CSV
        if ($hasDirectState -and $groupNames.Count -gt 0) {
          Add-UsersRow -SourceVal "Direct" -GroupName $null -HasBoth $true
          foreach ($gname in $groupNames) { Add-UsersRow -SourceVal "Group" -GroupName $gname -HasBoth $true }
        } elseif ($hasDirectState) {
          Add-UsersRow -SourceVal "Direct" -GroupName $null -HasBoth $false
        } elseif ($groupNames.Count -gt 0) {
          foreach ($gname in $groupNames) { Add-UsersRow -SourceVal "Group" -GroupName $gname -HasBoth $false }
        } else {
          Add-UsersRow -SourceVal "Unknown" -GroupName $null -HasBoth $false
        }

        # Compact user service-plan state fact and normalized service-plan catalog
        if ($ServicePlans) {
          $disabledIds = if ($disabledPlansBySku.ContainsKey($skuId)) { $disabledPlansBySku[$skuId] } else { @() }
          [void]$servicePlanStatePairs.Add(('{0}|{1}' -f $uid, $skuId))
          foreach ($sp in $ld.ServicePlans) {
            $planDisplayName = Get-ServiceFriendly -PlanName $sp.ServicePlanName -PlanId $sp.ServicePlanId -ByName $null -ById $SvcMapById
            $isEnabled = (-not ($disabledIds -contains $sp.ServicePlanId))

            $servicePlanRowCount++
            Write-LicensesServicePlanStateRow `
              -Writer $servicePlanStateWriter `
              -TenantKey $global:SmartM365TenantKey `
              -UserId ([string]$uid) `
              -SkuId ([string]$skuId) `
              -PlanId ([string]$sp.ServicePlanId) `
              -StateCode (Get-ServicePlanStateCode -IsEnabled $isEnabled -PlanStatus $sp.ProvisioningStatus) `
              -RowNumber $servicePlanRowCount

            $catalogKey = "{0}|{1}" -f $skuId, $sp.ServicePlanId
            if (-not $servicePlanCatalog.ContainsKey($catalogKey)) {
              $servicePlanCatalog[$catalogKey] = [PSCustomObject][ordered]@{
                "SkuId"           = $skuId
                "SkuPartNumber"   = $skuPartById[$skuId]
                "SKU name"        = $skuDisplayById[$skuId]
                "PlanId"          = $sp.ServicePlanId
                "PlanName"        = $sp.ServicePlanName
                "PlanDisplayName" = $planDisplayName
                "AppliesTo"       = $null
                "TenantProvisioningStatus" = $null
                "CollectedAtUtc"  = $licenseUsersCollectedAtUtc.ToString('o')
              }
            }

            if ($sp.ServicePlanName -like '*EXCHANGE*' -and $skuPartById[$skuId] -like '*SPE*') {
              $rowPlanExchange = [ordered]@{
                "Id"                  = "$uid-$skuId-$($sp.ServicePlanId)"
                "User principal name" = $upn
                "primarysmtp"         = $primarySmtp
                "Display name"        = $displayName
                "UserId"              = $uid
                "SkuId"               = $skuId
                "SkuPartNumber"       = $skuPartById[$skuId]
                "SKU name"            = $skuDisplayById[$skuId]
                "PlanId"              = $sp.ServicePlanId
                "PlanName"            = $sp.ServicePlanName
                "PlanDisplayName"     = $planDisplayName
                "IsEnabled"           = $isEnabled
                "PlanStatus"          = $sp.ProvisioningStatus
              }
              foreach ($k in @($rowPlanExchange.Keys)) {
                if ($rowPlanExchange[$k] -is [string]) {
                  $rowPlanExchange[$k] = $rowPlanExchange[$k] -replace "`r`n|`n|`r"," "
                  $rowPlanExchange[$k] = $rowPlanExchange[$k] -replace '"',"'"
                }
              }
              $rowsPlansExchange.Add([PSCustomObject]$rowPlanExchange) | Out-Null
            }
          }
        }
      }
    } catch {
      if ($_.Exception.Data['SmartM365FatalServicePlanStateWrite']) { throw }
      $disposition = Register-LicensesUserFailure -ErrorRecord $_ -UserId ([string]$u.Id) -LicenseDetailStage $licenseDetailStage `
        -AssignmentPaths $assignmentPaths -AssignmentPathCountBefore $assignmentPathCountBefore -State $userResolution
      if ($disposition -ne 'Deleted') { WriteLog -Message "Error processing user $($u.UserPrincipalName): $_" "ERROR" }
    }
  }
  if ($servicePlanStateWriter) {
    Close-LicensesServicePlanStateWriter -Writer $servicePlanStateWriter
    $servicePlanStateWriter = $null
    $stateBuildingFile = Get-Item -LiteralPath $servicePlanStateBuildingPath -ErrorAction Stop
    WriteLog -Message ("Disk-backed compact service-plan source completed: Rows={0:N0}; Size={1:N1} MB; WorkingSet={2:N1} MB." -f $servicePlanRowCount, ($stateBuildingFile.Length / 1MB), ((Get-Process -Id $PID).WorkingSet64 / 1MB)) 'INFO'
  }
  Write-Progress -Activity "Processing users" -Completed -Status "Done"
  $usersDeletedQualification = Assert-LicensesUserResolution -State $userResolution -ListedUsers $uTotal `
    -MaxDeletedCount ([int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicenseDeletedUsersMaxCount' -DefaultValue 100)) `
    -MaxDeletedPercent ([double](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicenseDeletedUsersMaxPercent' -DefaultValue 1))
  if ($usersDeletedQualification) {
    try { Add-SmartM365SourceRunNote -Message ('{0}: users deleted from the directory during collection are excluded from the licensing CSVs.' -f $usersDeletedQualification) }
    catch { WriteLog -Message ("Deleted-user note could not be added to the run state: {0}" -f $_.Exception.Message) 'WARNING' }
  }

  if ($userSnapshotRefreshes -gt 0) { WriteLog -Message ("Users whose license assignment changed during collection and were read again: {0}." -f $userSnapshotRefreshes) 'WARNING' }
  $currentOperation = 'Check license assignment paths against user service-plan states'
  Assert-LicensesPlanStateCoverage -AssignmentPaths $assignmentPaths -PlanStatePairs $servicePlanStatePairs

  $currentOperation = 'Export immutable license assignment paths'
  Assert-SmartM365CsvDataCompleteness -Data $directoryGroupRows -BaseFileName 'M365_EntraGroups_All' `
    -Columns @('GroupId','DisplayName','GroupTypes','SecurityEnabled')
  $pathBaseName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'M365_Licenses_AssignmentPaths'
  Export-SmartM365Csv -Data $assignmentPaths.ToArray() `
    -Columns @('UserId','SkuId','AssignedByGroupId','AssignmentRoute','AssignmentState','AssignmentError','DisabledPlanIds','LastUpdatedDateTime','CollectedAtUtc') `
    -TimestampedPath (Join-Path $OutputPath ("{0}_{1}.csv" -f $pathBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))) `
    -LatestPath (Join-Path $LatestCsvFolderPath "$pathBaseName.csv") -NoWeeklyHistory | Out-Null

  $groupBaseName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'M365_EntraGroups_All'
  Export-SmartM365Csv -Data $directoryGroupRows `
    -Columns @('GroupId','DisplayName','Description','Mail','MailEnabled','SecurityEnabled','GroupTypes','MembershipRule','MembershipRuleProcessingState','OnPremisesSyncEnabled','OnPremisesSecurityIdentifier','CreatedDateTime','CollectedAtUtc') `
    -TimestampedPath (Join-Path $OutputPath ("{0}_{1}.csv" -f $groupBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))) `
    -LatestPath (Join-Path $LatestCsvFolderPath "$groupBaseName.csv") -NoWeeklyHistory | Out-Null

  # ---------------- Export Users ----------------
  $currentOperation = "Export user license CSV"
  if (-not $resultsUsers -or $resultsUsers.Count -eq 0) {
    Write-Warning "No user license rows produced."
    $emptyUsersBaseName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'M365_Licenses_Users'
    Export-SmartM365Csv -Data @() `
      -Columns @('Id','User principal name','primarysmtp','Display name','UserId','SkuId','SkuPartNumber','SKU name','Source','GroupsAssigningSku','GroupCountForSku','HasDirectAndGroup') `
      -TimestampedPath (Join-Path $OutputPath ("{0}_{1}.csv" -f $emptyUsersBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))) `
      -LatestPath (Join-Path $LatestCsvFolderPath "$emptyUsersBaseName.csv") -NoWeeklyHistory | Out-Null
    WriteLog -Message 'Successful empty effective-license snapshot published; older assignments are not reused.'
  } else {
    Write-Host ""
    Write-Host "--- Export CSV (Licenses - Users) ---"
$BaseFileName = "M365_Licenses_Users"
    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
      -OutputPath $OutputPath `
      -GlobalPath $LatestCsvFolderPath `
      -Data $resultsUsers -Encoding "UTF8" -NoTypeInformation -Delimiter "," -SkipWeeklyHistory
  }

  # ---------------- Export compact ServicePlans state fact + catalog ----------------
  $currentOperation = "Export service plan CSV"
  if ($ServicePlans) {
    $catalogRows = @($servicePlanCatalog.Values | Sort-Object SkuPartNumber, PlanName, PlanId)
    $catalogBaseName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'M365_Licenses_ServicePlans_Catalog'
    Export-SmartM365Csv -Data $catalogRows `
      -Columns @('SkuId','SkuPartNumber','SKU name','PlanId','PlanName','PlanDisplayName','AppliesTo','TenantProvisioningStatus','CollectedAtUtc') `
      -TimestampedPath (Join-Path $OutputPath ("{0}_{1}.csv" -f $catalogBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'))) `
      -LatestPath (Join-Path $LatestCsvFolderPath "$catalogBaseName.csv") -NoWeeklyHistory | Out-Null
    if ($servicePlanRowCount -gt 0) {
      Write-Host ""
      Write-Host "--- Export CSV (User ServicePlan States - Compact) ---"
      $statePublication = Publish-LicensesServicePlanStateCsvFile `
        -BuildingPath $servicePlanStateBuildingPath `
        -ExpectedRows $servicePlanRowCount `
        -CurrentOutputPath $OutputPath `
        -LatestOutputPath $LatestCsvFolderPath
      $servicePlanStateLatestPath = [string]$statePublication.PublishedPath

      if ($rowsPlansExchange.Count -gt 0) {
        Write-Host ""
        Write-Host "--- Export CSV (ServicePlans - Exchange Online) ---"
        $BaseFileName = "M365_Licenses_ServicePlans"
        ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
          -OutputPath $OutputPath `
          -GlobalPath $LatestCsvFolderPath `
          -Data $rowsPlansExchange -Encoding "UTF8" -NoTypeInformation -Delimiter "," -SkipWeeklyHistory
        WriteLog -Message ("ServicePlans filtered CSV exported: {0} rows (PlanName contains EXCHANGE + SkuPartNumber contains SPE)." -f $rowsPlansExchange.Count)
      } else {
        Write-Host "No ServicePlans rows matching PlanName containing 'EXCHANGE' and SkuPartNumber containing 'SPE'."
        WriteLog -Message "No ServicePlans rows matching PlanName containing 'EXCHANGE' and SkuPartNumber containing 'SPE'."
      }

      if (-not (Test-SmartM365MaxItemsMode)) {
        $stateLatestPath = Join-Path -Path $LatestCsvFolderPath -ChildPath "M365_Licenses_UserServicePlanStates.csv"
        $catalogLatestPath = Join-Path -Path $LatestCsvFolderPath -ChildPath "M365_Licenses_ServicePlans_Catalog.csv"
        $publishedThisRun = $global:csvGeneratedPaths -and
          $global:csvGeneratedPaths.Contains($stateLatestPath) -and
          $global:csvGeneratedPaths.Contains($catalogLatestPath)
        if (-not $publishedThisRun -or -not (Test-Path -LiteralPath $stateLatestPath) -or -not (Test-Path -LiteralPath $catalogLatestPath)) {
          throw "Compact ServicePlans state and catalog CSV publication must succeed in the current run before retiring the legacy detailed CSV."
        }
      }

      Remove-LegacyDetailedServicePlansExport -CurrentOutputPath $OutputPath -LatestOutputPath $LatestCsvFolderPath
    } else {
      # A successful empty current state must replace, not reuse, an older state CSV.
      $statePublication = Publish-LicensesServicePlanStateCsvFile -BuildingPath $servicePlanStateBuildingPath `
        -ExpectedRows 0 -CurrentOutputPath $OutputPath -LatestOutputPath $LatestCsvFolderPath
      $servicePlanStateLatestPath = [string]$statePublication.PublishedPath
      Write-Warning "No ServicePlans rows produced."
      WriteLog -Message "No ServicePlans rows to export."
    }
  }
  # ---------------- Export Tenant ----------------
  $currentOperation = "Export tenant license CSV"
  $tenantRows = New-Object System.Collections.Generic.List[object]
  foreach ($kvp in $tenantSkuMap.GetEnumerator()) {
    $tenantRows.Add([PSCustomObject]@{
      "Id"                   = "$($kvp.Key)"
      "TenantSkuDisplayName" = Get-SkuDisplayName -SkuId $kvp.Key -SkuPartNumber $kvp.Value.PartNumber -MapByPart $SkuMapByPart -MapById $SkuMapById
      "TenantSkuPartNumber"  = $kvp.Value.PartNumber
      "TenantPrepaidEnabled" = $kvp.Value.PrepaidEnabled
      "TenantConsumedUnits"  = $kvp.Value.ConsumedUnits
      "TenantPrepaidWarning" = $kvp.Value.PrepaidWarning
      "TenantPrepaidSuspended" = $kvp.Value.PrepaidSuspended
      "TenantPrepaidLockedOut" = $kvp.Value.PrepaidLockedOut
      "CapabilityStatus" = $kvp.Value.CapabilityStatus
      "AppliesTo" = $kvp.Value.AppliesTo
      "SubscriptionIds" = $kvp.Value.SubscriptionIds
      "CollectedAtUtc" = $skusCollectedAtUtc
    }) | Out-Null
  }
  if ($tenantRows.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Export CSV (Licenses - Tenant) ---"
$BaseFileName = "M365_Licenses_Tenant"
    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
      -OutputPath $OutputPath `
      -GlobalPath $LatestCsvFolderPath `
      -Data $tenantRows -Encoding "UTF8" -NoTypeInformation -Delimiter "," -SkipWeeklyHistory
  }

  # ---------------- Export Groups ----------------
  $currentOperation = "Export group license CSV"
  $groupRows = New-Object System.Collections.Generic.List[object]
  foreach ($kvp in $groupAgg.GetEnumerator()) {
    $gid = $kvp.Key; $g = $kvp.Value
    $skuIds   = @($g.SkuIds)
    $skuParts = @()
    $skuNames = @()
    foreach ($sid in $skuIds) {
      $pn = if ($tenantSkuMap.ContainsKey($sid)) { $tenantSkuMap[$sid].PartNumber } else { "" }
      $skuParts += $pn
      $skuNames += (Get-SkuDisplayName -SkuId $sid -SkuPartNumber $pn -MapByPart $SkuMapByPart -MapById $SkuMapById)
    }
    $groupRows.Add([PSCustomObject]@{
      "Id"               = "$gid"
      "GroupId"          = $gid
      "GroupDisplayName" = $g.DisplayName
      "UsersCount"       = $g.UserIds.Count
      "DistinctSkuCount" = $skuIds.Count
      "SkuPartNumbers"   = ($skuParts | Where-Object { $_ } | Sort-Object -Unique) -join "; "
      "SKU names"        = ($skuNames | Where-Object { $_ } | Sort-Object -Unique) -join "; "
    }) | Out-Null
  }
  if ($groupRows.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Export CSV (Licenses - Groups) ---"
$BaseFileName = "M365_Licenses_Groups"
    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
      -OutputPath $OutputPath `
      -GlobalPath $LatestCsvFolderPath `
      -Data $groupRows -Encoding "UTF8" -NoTypeInformation -Delimiter "," -SkipWeeklyHistory
  } else {
    Write-Host "No groups with license assignments discovered during this run."
  }

  if ($ServicePlans -and $servicePlanRowCount -gt 0) {
    $currentOperation = "Publish consolidated Licensing WeeklyHistory"
    Publish-LicensesWeeklyHistory `
      -ServicePlanStateCsvPath $servicePlanStateLatestPath `
      -ExpectedServicePlanRows $servicePlanRowCount `
      -CurrentOutputPath $OutputPath `
      -LatestOutputPath $LatestCsvFolderPath
  }

  # ---------------- Disconnect / Cleanup ----------------
  $usersProcessedCount = if ($null -eq $users) { 0 } else { [int]$users.Count }
  $userLicenseRowCount = $resultsUsers.Count
  $tenantSkuRowCount = $tenantRows.Count
  $groupRowCount = $groupRows.Count

  # Published before the receipt completes so its CSVs are recorded as optional outputs of this run.
  $licenseRunReport = $null
  if ($MaxItems -eq 0 -and $TopUsers -eq 0 -and -not (Test-SmartM365MaxItemsMode)) {
    $currentOperation = 'Publish license report'
    $licenseRunReport = Publish-LicensesRunReport -Folder $LatestCsvFolderPath -TenantKey $global:SmartM365TenantKey -CollectedAtUtc $skusCollectedAtUtc `
      -TenantRows $tenantRows.ToArray() -ForceAdCsvAnalysis:$ForceAdCsvAnalysis
  }

  $currentOperation = 'Complete licensing source receipt'
  Set-SmartM365CmdbSourceScope -CompleteScope ($MaxItems -eq 0 -and $TopUsers -eq 0) -Scope 'CMDB:skus,license_paths,plans,user_plans,groups' `
    -Qualifications @($usersDeletedQualification | Where-Object { $_ })
  try {
    $completedReceiptPath = Complete-SmartM365CmdbSourceReceipt -Status 'Completed'
    if (-not $completedReceiptPath -or -not (Test-Path -LiteralPath $completedReceiptPath -PathType Leaf)) {
      throw 'Licensing source receipt was not written.'
    }
    $completedReceipt = Get-Content -LiteralPath $completedReceiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if ([string]$completedReceipt.Status -ne 'Completed' -or [bool]$completedReceipt.IsPartialInventory -or
        -not [bool]$completedReceipt.ConsumerScopeQualified) {
      throw ("Licensing source receipt is not qualified: {0}. {1}" -f $completedReceipt.Status, $completedReceipt.Error)
    }
  }
  catch { throw ("Licensing source receipt could not be completed; summary email was not sent: {0}" -f $_.Exception.Message) }
  if ($licenseRunReport -and $licenseRunReport.Published -and $licenseRunReport.PSObject.Properties['UploadDeferred']) {
    $currentOperation = 'Upload promoted license report'
    $licenseRunReport = Invoke-LicensesReportUpload -Result $licenseRunReport -ReportPaths $licenseRunReport.ReportPaths -SkuPaths $licenseRunReport.SkuPaths
  }

  $currentOperation = 'Send focused license summary email'
  if ($licenseRunReport -and $licenseRunReport.Published -and $licenseRunReport.UploadSatisfied) {
    [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $LatestCsvFolderPath -ExpectedTenantKey $global:SmartM365TenantKey -MailStatePath $licenseSummaryMailStatePath -ForceLicenseSummaryEmail:$ForceLicenseSummaryEmail)
  }
  else { WriteLog -Message 'Focused license summary email not sent: the license report was not published and uploaded in this run.' 'WARNING' }

  if ($connectedGraphInThisRun) {
    $currentOperation = "Disconnect Microsoft Graph"
    Write-Host ""
    Write-Host "--- Disconnect Cloud Services ---"
    try {
      Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true
    }
    catch {
      WriteLog -Message ("Microsoft Graph disconnect cleanup did not complete: {0}" -f $_.Exception.Message) "WARN"
    }
  }

  $currentOperation = "Apply retention cleanup"
  try {
    Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -LogFile $global:logTextFile
  }
  catch {
    WriteLog -Message ("CSV retention cleanup failed: {0}" -f $_.Exception.Message) "WARN"
  }
  try {
    RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
  }
  catch {
    WriteLog -Message ("Log retention cleanup failed: {0}" -f $_.Exception.Message) "WARN"
  }

  $currentOperation = "Send Teams completion notification"
  Send-LicensesInventorySuccessNotification `
    -UsersProcessed $usersProcessedCount `
    -UserLicenseRows $userLicenseRowCount `
    -ServicePlanRows $servicePlanRowCount `
    -TenantSkuRows $tenantSkuRowCount `
    -GroupRows $groupRowCount `
    -UsersDeletedDuringCollection ([int]$userResolution.UsersDeletedDuringCollection) `
    -OutputPath $OutputPath

  WriteLog -Message "$TaskName completed."
  try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
  Complete-SmartM365ExecutionContext -Status Auto
  $runFinalized = $true
}
catch [System.Management.Automation.PipelineStoppedException] {
  throw
}
catch {
  if ($servicePlanStateWriter) {
    try { Close-LicensesServicePlanStateWriter -Writer $servicePlanStateWriter } catch {}
    $servicePlanStateWriter = $null
  }
  $globalError = $_
  WriteLog -Message ("Global error in M365 licenses inventory: {0}" -f $globalError) "ERROR"
  Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red
  Send-LicensesInventoryErrorNotification -ErrorRecord $globalError -Operation $currentOperation -OutputPath $OutputPath

  # -------- Global error email notification --------
  try {
    $title = "M365 licenses inventory - ERROR"
    $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:logTextFile)
"@

    $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

    $attachments = @()
    if ($global:logTextFile -and (Test-Path $global:logTextFile)) {
      $attachments = @($global:logTextFile)
    }

    Send-SmartM365Mail -Subject $title -BodyHtml $bodyHtml -Attachments $attachments -SendMailMode Graph -HighPriority
  } catch {
    WriteLog -Message ("Failed to send global error notification email: {0}" -f $_) "ERROR"
  }

  try {
    if ($connectedGraphInThisRun) {
      Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true
    }
  } catch {
    WriteLog -Message ("Failed to disconnect Microsoft Graph after error: {0}" -f $_.Exception.Message) "WARNING"
  }

  try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
  Complete-SmartM365ExecutionContext -Status Auto
  $runFinalized = $true
  exit 1
}
finally {
  if (-not $runFinalized) {
    # Ctrl+C bypasses the normal completion path. Keep published CSVs intact,
    # but close local resources and invalidate the in-progress source receipt.
    if ($servicePlanStateWriter) {
      try { Close-LicensesServicePlanStateWriter -Writer $servicePlanStateWriter } catch {}
      $servicePlanStateWriter = $null
    }
    if ($inventoryEnvironmentInitialized) {
      try { WriteLog -Message ("Licenses inventory interrupted during: {0}" -f $currentOperation) -Level ERROR } catch {}
    }
    try { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch {}
    if ($connectedGraphInThisRun) {
      try { Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true } catch {}
    }
    if ($transcriptStarted) {
      try { Stop-Transcript | Out-Null } catch {}
      if ($global:logTranscriptFile) {
        try { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch {}
      }
    }
    if ($inventoryEnvironmentInitialized) {
      try { Write-SmartM365CompletionBanner -Status Failed -ClosedTranscriptPath ([string]$global:logTranscriptFile) -Force } catch {}
    }
    [Environment]::ExitCode = 130
  }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBN+ssiupTBRM/X
# eyyMjKE5IxP3wylPlyPPdB5i2UxmTaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDdglY1eo7/kl75BQcUxirjZwSNywtYTC/QouXklsDD2MA0GCSqG
# SIb3DQEBAQUABIIBgCbvhFgsT2FdkaM40Fqo2hijU5sfnpXv26gTOavGttpixYCo
# B1h1SzciSQPwiYlfZT4A+GgDjZZg0yGWIAAlPTotV5q4USQH9GR41NXWIcQ6aM2w
# 2Yj7zXJU3fvQee44OnDhYjSVSbjLX9hy480XGH3lubmG3UiL3bVAHevAud4KGROg
# AWhp0NFev+C7mrwC8gQwvOcjA7exQLpShz4XuuiK7IqFzQC/3Ps8uLvu1tVRk+Nn
# Dxzws8+8qGd1IAC070N3G4C8P/DQzcn+6wByqzlo0ferwoUBFfKZcUb5QVAaowOm
# CwllMnm5WjNY/c6OrefJOjl/++GcR8GSS8MReNUVEXGILl/LQo32RIS9R33Nn4iT
# v9DXpKAUKntwBesfuXgJxltCMUWmTltoPUYCx4qZSbBRVMrvhjgrof9U+/Op4jO+
# 0SmxoFyHrHIgnNu4Ca5H9j0qKoxahyVQb6N4+X/vgNc1HjBmwRJefIWZADOl7ylt
# UM36nLxq96pSmwTuOqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNDQ5
# MjlaMC8GCSqGSIb3DQEJBDEiBCDelpSjDhIcstq8vCgT+CckUB6x0XSFxlhS3JkM
# uCpFezANBgkqhkiG9w0BAQEFAASCAgCFB1drLNPKnXlfRsbE4zBrBEDe7FBidd3P
# z0XIuM8xsmYSKo2Ei43V01J0D8fouo4fbIX3/ffGIiEyanX4UAM40185KregSf5x
# nsuw/tYjsvrZ7Bgopk616wYphFbvHUfl0kfTnJW7Frf+5EOKuy+vrIlsCuglwJZ7
# IrrY2t4bTUXdDORRPtabkD0AdL3XOTznHTrp90SjhMRWhP0E7qKY3QmEPOJ+QY3i
# 5zX4f8iVaTpEns9Y+LLfMgy8xcZqaD8O6cUGKRtN+AIQ/8hJPKRVZ131ZrSQxEiH
# 6dRkMFQEh61Ici2CeXIl2tCDNxhf5+8gueWvkL4SexYR9TBaVyEHPXx58RY923TS
# RMBAndCqboJsR9J12AP4XxP1QeytNkD0SZ5puO551nZsYk1u74EwkqAn7R79+qdF
# TmA2yRREuua03f3YO8GIjJsGPC/6zs0BrUIvesV0sMR0/LBxkJWaibzAZUfFd54b
# ZXyb6DjVZ+uyX0/AI11BLq2Wgzg1D73EhfoWDpcfo+rg07Ef0eH8svvS44FTKRVj
# YyebstixUIjekUszq6eDJCEcGogjOq5e4F8KnKidCHbeE9y945kh1MsaKT7btX1r
# ZXvzp7bp/uf4noZz4VHLItv6Qp0wBJ1mxwa0ovnVS8ynK3Eb+15dZfy3OmnLnHFx
# z3cpXGH3VA==
# SIG # End signature block
