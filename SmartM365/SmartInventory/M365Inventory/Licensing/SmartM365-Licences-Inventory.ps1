<#
.SYNOPSIS
  Export M365 license assignments with Users/Tenant/Groups + one normalized ServicePlans CSV.
  Detects Direct vs Group via user.LicenseAssignmentStates.assignedByGroup.
  Maps SKU & Service Plan friendly names from the Microsoft CSV (default: script folder).
.VERSION
1.6


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Identity.DirectoryManagement; Microsoft.Graph.Users; Microsoft.Graph.Groups.
    Minimum Graph application permissions: Directory.Read.All; User.Read.All; Group.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
  Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.2
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

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
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

            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Edit the local JSON now if needed. Press Enter to continue with the current file values.' -ForegroundColor Yellow
            Read-Host 'Press Enter to continue'
        }
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
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
    Import-Module $modulePath -ErrorAction Stop
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
      $status429 = ($msg -match '429') -or ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429)
      $status503 = ($msg -match '503') -or ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 503)
      if (-not ($status429 -or $status503)) { throw }
      $retryAfter=$null; try {
        if ($_.Exception.Response -and $_.Exception.Response.Headers) { $retryAfter = $_.Exception.Response.Headers['Retry-After'] }
        elseif ($_.Exception.Data['Retry-After']) { $retryAfter = $_.Exception.Data['Retry-After'] }
      } catch {}
      if ($retryAfter) { $delay=[int]$retryAfter } else { $delay=[math]::Min(60, ($BaseDelaySeconds * [math]::Pow(2, $attempt))) }
      Write-Warning ("Graph throttled/unavailable (attempt {0}/{1}). Sleeping {2}s. Error: {3}" -f $attempt,$MaxRetries,$delay,$msg)
      Start-Sleep -Seconds $delay
      if ($attempt -ge $MaxRetries) { throw "Max retry attempts reached: $msg" }
    }
  }
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
    [string]$OutputPath
  )

  try {
    $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
    $resultSummary = "Microsoft 365 licenses inventory completed without error. Users processed: {0}; user license rows: {1}; service plan rows: {2}; tenant SKUs: {3}; groups: {4}." -f $UsersProcessed, $UserLicenseRows, $ServicePlanRows, $TenantSkuRows, $GroupRows
    $facts = @{
      "Script name"         = $scriptName
      "Tenant/Organization" = $OrgDomain
      "Users processed"     = $UsersProcessed
      "User license rows"   = $UserLicenseRows
      "Service plan rows"   = $ServicePlanRows
      "Tenant SKU rows"     = $TenantSkuRows
      "Group rows"          = $GroupRows
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

function To-GuidOrNull {
  param($Value)
  try {
    if ($null -eq $Value) { return $null }
    if ($Value -is [Guid]) { return $Value }
    if ("$Value".Trim()) { return [Guid]$Value }
  } catch { return $null }
  return $null
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

# ==========================================================
# Main
# ==========================================================
$ScriptVersion = "1.6"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicensesCsvLogFolderPath' -DefaultValue $OutputPath
$connectedGraphInThisRun = $false
$currentOperation = "Initialize script environment"
$usersProcessedCount = 0
$userLicenseRowCount = 0
$servicePlanRowCount = 0
$tenantSkuRowCount = 0
$groupRowCount = 0

try {
  # ------------------------
  # Initialize script environment
  # ------------------------
  $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
  Start-Transcript -Path $global:logTranscriptFile -Append
  WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
  $OutputPath = $InitializeOutputPath
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
  $tenantSkuMap = @{}
  foreach ($s in $subscribedSkus) {
    $tenantSkuMap[$s.SkuId] = @{
      PartNumber     = $s.SkuPartNumber
      PrepaidEnabled = if ($s.PrepaidUnits -and $s.PrepaidUnits.Enabled) { [int]$s.PrepaidUnits.Enabled } else { 0 }
      ConsumedUnits  = if ($s.ConsumedUnits) { [int]$s.ConsumedUnits } else { 0 }
    }
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

  # Sampling: prefer users with group-based licensing
  if ($TopUsers -gt 0) {
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
  $resultsUsers      = New-Object System.Collections.Generic.List[object]
  $rowsPlansDetailed = New-Object System.Collections.Generic.List[object]

  $uTotal = ($users | Measure-Object).Count
  $uIndex = 0

  foreach ($u in $users) {
    $uIndex++
    $pct = if ($uTotal -gt 0) { [math]::Floor(($uIndex * 100.0) / $uTotal) } else { 100 }
    Write-Progress -Activity "Processing users" -Status "[$uIndex/$uTotal] $($u.UserPrincipalName)" -PercentComplete $pct

    try {
      $upn         = $u.UserPrincipalName
      $displayName = $u.DisplayName
      $uid         = $u.Id
      $primarySmtp = Get-PrimarySmtpAddress -User $u

      # Direct disabled plans per SKU (raw)
      $disabledPlansBySku=@{}
      foreach ($al in (@($u.AssignedLicenses) | ForEach-Object { $_ })) {
        if ($al.SkuId) { $disabledPlansBySku[$al.SkuId]=@($al.DisabledPlans) }
      }

      # License details (Direct + Group)
      $licenseDetails = Invoke-GraphWithRetry { Get-MgUserLicenseDetail -UserId $uid -ErrorAction Stop }
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

        # Normalized ServicePlans CSV
        if ($ServicePlans) {
          $disabledIds = if ($disabledPlansBySku.ContainsKey($skuId)) { $disabledPlansBySku[$skuId] } else { @() }
          foreach ($sp in $ld.ServicePlans) {
            $rowPlan = [ordered]@{
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
              "PlanDisplayName"     = (Get-ServiceFriendly -PlanName $sp.ServicePlanName -PlanId $sp.ServicePlanId -ByName $null -ById $SvcMapById)
              "IsEnabled"           = (-not ($disabledIds -contains $sp.ServicePlanId))
              "PlanStatus"          = $sp.ProvisioningStatus
            }
            foreach ($k in @($rowPlan.Keys)) {
              if ($rowPlan[$k] -is [string]) {
                $rowPlan[$k] = $rowPlan[$k] -replace "`r`n|`n|`r"," "
                $rowPlan[$k] = $rowPlan[$k] -replace '"',"'" 
              }
            }
            $rowsPlansDetailed.Add([PSCustomObject]$rowPlan) | Out-Null
          }
        }
      }
    } catch {
      WriteLog -Message "Error processing user $($u.UserPrincipalName): $_" "ERROR"
    }
  }
  Write-Progress -Activity "Processing users" -Completed -Status "Done"

  # ---------------- Export Users ----------------
  $currentOperation = "Export user license CSV"
  if (-not $resultsUsers -or $resultsUsers.Count -eq 0) {
    Write-Warning "No user license rows produced."
    WriteLog -Message "No Users rows to export."
  } else {
    Write-Host ""
    Write-Host "--- Export CSV (Licenses - Users) ---"
$BaseFileName = "M365_Licenses_Users"
    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
      -OutputPath $OutputPath `
      -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
      -Data $resultsUsers -Encoding "UTF8" -NoTypeInformation -Delimiter ","
  }

  # ---------------- Export ServicePlans (single detailed CSV) ----------------
  $currentOperation = "Export service plan CSV"
  if ($ServicePlans) {
    if ($rowsPlansDetailed.Count -gt 0) {
      Write-Host ""
      Write-Host "--- Export CSV (ServicePlans - Detailed) ---"
$BaseFileName = "M365_Licenses_ServicePlans_Detailed"
      ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
        -OutputPath $OutputPath `
        -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
        -Data $rowsPlansDetailed -Encoding "UTF8" -NoTypeInformation -Delimiter ","

      # ---------------- Export ServicePlans (filtered: PlanName contains EXCHANGE + SkuPartNumber contains SPE) ----------------
      $rowsPlansFiltered = $rowsPlansDetailed | Where-Object { $_.PlanName -like '*EXCHANGE*' -and $_.SkuPartNumber -like '*SPE*' }
      if ($rowsPlansFiltered -and @($rowsPlansFiltered).Count -gt 0) {
        Write-Host ""
        Write-Host "--- Export CSV (ServicePlans - Exchange Online) ---"
$BaseFileName = "M365_Licenses_ServicePlans"
        ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
          -OutputPath $OutputPath `
          -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
          -Data @($rowsPlansFiltered) -Encoding "UTF8" -NoTypeInformation -Delimiter ","
        WriteLog -Message ("ServicePlans filtered CSV exported: {0} rows (PlanName contains EXCHANGE + SkuPartNumber contains SPE)." -f @($rowsPlansFiltered).Count)
      } else {
        Write-Host "No ServicePlans rows matching PlanName containing 'EXCHANGE' and SkuPartNumber containing 'SPE'."
        WriteLog -Message "No ServicePlans rows matching PlanName containing 'EXCHANGE' and SkuPartNumber containing 'SPE'."
      }
    } else {
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
    }) | Out-Null
  }
  if ($tenantRows.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Export CSV (Licenses - Tenant) ---"
$BaseFileName = "M365_Licenses_Tenant"
    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
      -OutputPath $OutputPath `
      -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
      -Data $tenantRows -Encoding "UTF8" -NoTypeInformation -Delimiter ","
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
      -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
      -Data $groupRows -Encoding "UTF8" -NoTypeInformation -Delimiter ","
  } else {
    Write-Host "No groups with license assignments discovered during this run."
  }

  # ---------------- Disconnect / Cleanup ----------------
  $usersProcessedCount = if ($null -eq $users) { 0 } else { [int]$users.Count }
  $userLicenseRowCount = $resultsUsers.Count
  $servicePlanRowCount = $rowsPlansDetailed.Count
  $tenantSkuRowCount = $tenantRows.Count
  $groupRowCount = $groupRows.Count

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
    RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
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
    -OutputPath $OutputPath

  WriteLog -Message "$TaskName completed."
  try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
  Complete-SmartM365ExecutionContext -Status Auto
}
catch {
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
  exit 1
}
