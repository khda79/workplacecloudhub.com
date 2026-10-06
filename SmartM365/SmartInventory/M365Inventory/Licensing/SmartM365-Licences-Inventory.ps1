<#
.SYNOPSIS
  Export M365 license assignments with Users/Tenant/Groups plus compact user service-plan state codes and a service-plan catalog.
  WeeklyHistory retains the detailed IsEnabled and PlanStatus representation.
  Detects Direct vs Group via user.LicenseAssignmentStates.assignedByGroup.
  Maps SKU & Service Plan friendly names from the Microsoft CSV (default: script folder).
  SendLicenseSummaryEmailOnly sends the license overview and recovery summary from existing published CSVs without collecting again.
  ForceAdCsvAnalysis uses a fresh, structurally valid AD CSV for this email when only its collector receipt is rejected.
  BypassLicenseUsersReceipt temporarily accepts the fresh license-users CSV when the previous completed receipt does not list it.
.VERSION
 1.31
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Identity.DirectoryManagement; Microsoft.Graph.Users; Microsoft.Graph.Groups.
    Minimum Graph application permissions: Directory.Read.All; User.Read.All; Group.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
  Author: https://github.com/khda79/workplacecloudhub.com
     Version : 1.31
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
  [switch]$ForceAdCsvAnalysis,
  [switch]$BypassLicenseUsersReceipt,
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
    Import-Module -Name $modulePath -MinimumVersion '1.0.65' -ErrorAction Stop
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
      $statusCode=$null
      try{if($_.Exception.Response){$statusCode=[int]$_.Exception.Response.StatusCode}}catch{}
      if($null-eq$statusCode){try{$statusCode=[int]$_.Exception.Data['StatusCode']}catch{}}
      $isTransient=$statusCode -in @(408,409,429,500,502,503,504) -or $msg -match '(?i)throttl|TooManyRequests|temporarily unavailable|timeout|timed out'
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
      if ([string]$tenantRow.TenantSkuPartNumber -notin $product.PartNumbers) { continue }
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

function Get-LicensesAdditionalOverviewRows {
  param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows)

  $totals = @{}
  foreach ($name in @('Microsoft 365 Copilot','Dynamics 365','Power BI')) {
    $totals[$name] = @{ Enabled=[long]0; Consumed=[long]0; Subscribed=$false }
  }
  foreach ($tenantRow in $TenantRows) {
    $sku = ([string]$tenantRow.TenantSkuPartNumber).Trim()
    $name = if ($sku -in @('Microsoft_365_Copilot','M365_COPILOT')) { 'Microsoft 365 Copilot' }
            elseif ($sku -match '^(DYN365_|DYNAMICS_365_)' -and $sku -notmatch '(SANDBOX|TRIAL|PREVIEW|VIRAL|FREE|DEMO|TEST)') { 'Dynamics 365' }
            elseif ($sku -in @('POWER_BI_PRO','PBI_PREMIUM_PER_USER')) { 'Power BI' }
            else { '' }
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
  $status = if ($Row.Subscribed) { 'Enabled licenses' } else { 'Not subscribed' }
  return '<td width="{0}%" style="width:{0}%;padding:5px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;background:#ffffff;border:1px solid #dce6ed;border-top:4px solid {1};"><tr><td style="padding:12px 10px;"><div style="font-size:12px;line-height:17px;font-weight:700;color:#334155;">{2}</div><div style="margin-top:6px;font-size:22px;line-height:26px;font-weight:700;color:#0f172a;">{3}</div><div style="font-size:10px;line-height:14px;color:#64748b;">{4}</div><div style="margin-top:8px;font-size:12px;line-height:17px;color:#334155;"><strong>{5}</strong> used &nbsp;&middot;&nbsp; <strong>{6}</strong> used</div></td></tr></table></td>' -f `
    $Width,$Accent,$label,$Row.Enabled,$status,$Row.Consumed,$percent
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

function Get-LicensesCsvSource {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$FileName,
    [Parameter(Mandatory)][string[]]$Columns,
    [Parameter(Mandatory)][datetime]$AsOfUtc,
    [string]$CollectorManifestName = '',
    [switch]$RequireFileReceipt,
    [switch]$RequireConsumerScope
  )
  $path = Join-Path -Path $Folder -ChildPath $FileName
  $source = [pscustomobject]@{ Name=$FileName; Path=$path; Ready=$false; Reason=''; Date=''; ModifiedUtc=$null; Provenance='CSV only'; Forced=$false }
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $source.Reason='missing'; return $source }
  $item = Get-Item -LiteralPath $path
  $source.ModifiedUtc = $item.LastWriteTimeUtc
  $source.Date = $item.LastWriteTimeUtc.ToString('yyyy-MM-dd')
  if ($item.LastWriteTimeUtc.Date -lt $AsOfUtc.Date.AddDays(-14)) { $source.Reason='older than 14 days'; return $source }
  $header = Get-Content -LiteralPath $path -TotalCount 1 -ErrorAction Stop
  $foundColumns = @($header.TrimStart([char]0xFEFF).Replace('"','').Split(','))
  foreach ($column in $Columns) {
    if ($column -notin $foundColumns) { $source.Reason="missing column: $column"; return $source }
  }
  if ($CollectorManifestName) {
    $manifestPath = Join-Path -Path $Folder -ChildPath $CollectorManifestName
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
      try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ([string]$manifest.Status -ne 'Completed' -or [bool]$manifest.IsPartialInventory) {
          $source.Reason = "collector receipt is $($manifest.Status) or partial"
          return $source
        }
        if ($RequireConsumerScope -and -not [bool]$manifest.ConsumerScopeQualified) {
          $source.Reason = 'collector receipt does not qualify the consumer scope'
          return $source
        }
        $fileReceipts = @($manifest.Files | Where-Object { [string]$_.File -eq $FileName })
        if ($RequireFileReceipt -and $fileReceipts.Count -eq 0) {
          $source.Reason = 'file missing from collector receipt'
          return $source
        }
        if ($fileReceipts.Count -gt 1 -or ($fileReceipts.Count -eq 1 -and
            ([string]$fileReceipts[0].Status -ne 'Success' -or [bool]$fileReceipts[0].IsPartialInventory))) {
          $source.Reason = 'file receipt is not a single successful complete export'
          return $source
        }
        if ($RequireFileReceipt) {
          if (-not [string]$fileReceipts[0].SHA256 -or
              (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne [string]$fileReceipts[0].SHA256) {
            $source.Reason = 'file hash differs from collector receipt'
            return $source
          }
        }
        $source.Provenance = 'collector completed'
      }
      catch { $source.Reason='collector receipt is unreadable'; return $source }
    }
    elseif ($RequireFileReceipt) { $source.Reason='collector receipt is missing'; return $source }
  }
  $source.Ready = $true
  return $source
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
    foreach ($row in (Import-Csv -LiteralPath $Source.Path -ErrorAction Stop)) {
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

  $userMailboxes = [pscustomobject]@{ Available=$false; Universe=0; Total=0; OtherSkus=0; NoSkus=0; Unknown=0; Reason='' }
  $noUserMailbox = [pscustomobject]@{ Available=$false; Universe=0; Total=0; OtherSkus=0; NoSkus=0; MemberEnabled=0; MemberDisabled=0; Guests=0; Unknown=0; Reason='' }
  if (-not $MailboxSource.Ready) {
    $userMailboxes.Reason = [string]$MailboxSource.Reason
    $noUserMailbox.Reason = [string]$MailboxSource.Reason
    return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox }
  }

  $mailboxTypes = @{}
  $duplicateMailboxIds = @{}
  $unqualifiedUserMailboxIds = @{}
  $unjoinedMailboxRows = 0
  try {
    foreach ($row in (Import-Csv -LiteralPath $MailboxSource.Path -ErrorAction Stop)) {
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
      if ($AllSkusByUser.ContainsKey($id)) { $userMailboxes.OtherSkus++ }
      else { $userMailboxes.NoSkus++ }
    }
    $userMailboxes.Available = $true
  }
  catch { $userMailboxes.Reason=$_.Exception.Message; $noUserMailbox.Reason=$_.Exception.Message; return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox } }

  if (-not $ActiveSource.Ready) {
    $noUserMailbox.Reason = [string]$ActiveSource.Reason
  }
  elseif (-not $OnPremSource.Ready) {
    $noUserMailbox.Reason = "on-premises UserMailbox source: $($OnPremSource.Reason)"
  }
  elseif ($unjoinedMailboxRows -gt 0) {
    $noUserMailbox.Reason = "$unjoinedMailboxRows mailbox rows have no Entra ID; absence of a mailbox cannot be established"
  }
  else {
    try {
      $accounts = @{}
      $duplicateAccounts = @{}
      $accountRows = 0
      foreach ($row in (Import-Csv -LiteralPath $ActiveSource.Path -ErrorAction Stop)) {
        $accountRows++
        if (-not $row.PSObject.Properties['UserType']) { throw 'missing UserType column' }
        if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        $id = ([string]$row.'Object Id').Trim().ToLowerInvariant()
        if (-not $id) { $noUserMailbox.Unknown++; continue }
        if ($accounts.ContainsKey($id)) { $duplicateAccounts[$id] = $true; continue }
        $accounts[$id] = $row
      }
      if ($accountRows -eq 0) { throw 'empty Entra user export' }
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
      foreach ($row in (Import-Csv -LiteralPath $OnPremSource.Path -ErrorAction Stop)) {
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
        elseif ($enabled -eq 'true') { $noUserMailbox.MemberEnabled++ }
        else { $noUserMailbox.MemberDisabled++ }
      }
      $noUserMailbox.Available = $true
    }
    catch { $noUserMailbox.Reason=$_.Exception.Message }
  }
  return [pscustomobject]@{ UserMailboxes=$userMailboxes; NoUserMailbox=$noUserMailbox }
}

function Get-LicensesFocusedUsageRows {
  param(
    [Parameter(Mandatory)][string]$CsvFolderPath,
    [Parameter(Mandatory)][string]$ExpectedTenantKey,
    [datetimeoffset]$LicenseSnapshotUtc = [datetimeoffset]::MinValue,
    [datetime]$AsOfUtc = [datetime]::UtcNow,
    [switch]$ForceAdCsvAnalysis,
    [switch]$BypassLicenseUsersReceipt
  )
  $products = @(
    @{ Name='Microsoft 365 F1'; PartNumbers=@('M365_F1','M365_F1_COMM') }
    @{ Name='Microsoft 365 F3'; PartNumbers=@('SPE_F1') }
    @{ Name='Microsoft 365 E3'; PartNumbers=@('SPE_E3') }
    @{ Name='Microsoft 365 E5'; PartNumbers=@('SPE_E5') }
  )
  $sources = [System.Collections.Generic.List[object]]::new()
  $licenseSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'M365_Licenses_Users.csv' -Columns @('TenantKey','UserId','SkuPartNumber') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt' -RequireFileReceipt
  $licenseReceiptBypassed = $false
  if ($BypassLicenseUsersReceipt -and -not $licenseSource.Ready -and $licenseSource.Reason -eq 'file missing from collector receipt') {
    $licenseSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'M365_Licenses_Users.csv' -Columns @('TenantKey','UserId','SkuPartNumber') -AsOfUtc $AsOfUtc
    $licenseReceiptBypassed = $licenseSource.Ready
  }
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
  if ($licenseSource.Ready) {
    try {
      foreach ($row in (Import-Csv -LiteralPath $licenseSource.Path -ErrorAction Stop)) {
        if ([string]$row.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        $rowUserId = ([string]$row.UserId).Trim().ToLowerInvariant()
        $rowSku = ([string]$row.SkuPartNumber).Trim().ToUpperInvariant()
        if ($rowUserId -and $rowSku) {
          if (-not $allSkusByUser.ContainsKey($rowUserId)) { $allSkusByUser[$rowUserId] = @{} }
          $allSkusByUser[$rowUserId][$rowSku] = $true
        }
        foreach ($product in $products) {
          if ([string]$row.SkuPartNumber -notin $product.PartNumbers) { continue }
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
    return [pscustomobject]@{ Rows=@($unavailableRows); Sources=$sources.ToArray(); SharedSourceReady=$false; IntuneSourceReady=$false; AdSourceForced=$false; LicenseSourceForced=$false; MailboxGap=$null }
  }
  if ($licenseReceiptBypassed) {
    $licenseSource.Forced = $true
    $licenseSource.Provenance = 'license-users CSV absent from the previous completed receipt'
    WriteLog -Message 'License users CSV analyzed without a file receipt; license assignment and recovery indicators are provisional.' 'WARNING'
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

  $intuneSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'Intune_Devices_Inventory.csv' -Columns @(
    'TenantKey','Device ID','OS','UserId','Primary user UPN'
  ) -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
  $sources.Add($intuneSource)
  $intunePrimaryPcUsers = @{}
  if ($intuneSource.Ready) {
    try {
      $deviceOwners = @{}
      foreach ($device in (Import-Csv -LiteralPath $intuneSource.Path -ErrorAction Stop)) {
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

  $adSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD','UserPrincipalName','LastLogonDate') -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
  $adReceiptReason = ''
  if ($ForceAdCsvAnalysis -and -not $adSource.Ready -and $adSource.Reason -match '^(collector receipt|file receipt)') {
    $adReceiptReason = [string]$adSource.Reason
    $adSource = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','ImmutableId_AD','UserPrincipalName','LastLogonDate') -AsOfUtc $AsOfUtc
  }
  $sources.Add($adSource)
  $adByImmutable = @{}
  $adByUpn = @{}
  $adDuplicateUpns = @{}
  if ($adSource.Ready) {
    try {
      Import-Csv -LiteralPath $adSource.Path -ErrorAction Stop | ForEach-Object {
        if ([string]$_.TenantKey -ine $ExpectedTenantKey) { throw 'another tenant is present' }
        $immutableId = ([string]$_.ImmutableId_AD).Trim().ToLowerInvariant()
        $upn = ([string]$_.UserPrincipalName).Trim().ToLowerInvariant()
        $adValue = [pscustomobject]@{ LastLogonDate=$_.LastLogonDate }
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
    catch { $adSource.Ready=$false; $adSource.Reason=$_.Exception.Message; $adByImmutable=@{}; $adByUpn=@{} }
  }
  if ($adSource.Ready -and $adReceiptReason) {
    $adSource.Forced = $true
    $adSource.Provenance = "FORCED AD CSV; ignored collector receipt: $adReceiptReason"
    WriteLog -Message ("AD CSV analysis forced despite collector receipt: {0}. AD/Entra inactivity is provisional." -f $adReceiptReason) 'WARNING'
  }

  $reportDefinitions = @(
    @{ Name='M365_Users_Activity.csv'; Key='UserPrincipalName'; Columns=@('TenantKey','UserPrincipalName','ReportPeriod','ReportRefreshDate','LastActivityDate','IsDeleted'); Refresh='ReportRefreshDate'; Period='ReportPeriod'; Expected='D180' }
    @{ Name='M365_Mailbox_Usage.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Is Deleted'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
    @{ Name='M365_Email_Activity.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Is Deleted','Send Count','Read Count'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
    @{ Name='M365_Apps_Usage_180D.csv'; Key='User Principal Name'; Columns=@('TenantKey','User Principal Name','Report Period','Report Refresh Date','Last Activity Date','Windows','Mac'); Refresh='Report Refresh Date'; Period='Report Period'; Expected='180' }
  )
  $reports = @{}
  foreach ($definition in $reportDefinitions) {
    $source = Get-LicensesCsvSource -Folder $CsvFolderPath -FileName $definition.Name -Columns $definition.Columns -AsOfUtc $AsOfUtc -CollectorManifestName 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt'
    $sources.Add($source)
    $reports[$definition.Name] = Read-LicensesIndexedSource -Source $source -ExpectedTenantKey $ExpectedTenantKey -KeyColumn $definition.Key -WantedKeys $wantedUpns -AsOfUtc $AsOfUtc -RefreshColumn $definition.Refresh -PeriodColumn $definition.Period -ExpectedPeriod $definition.Expected
  }

  $cutoff = $AsOfUtc.Date.AddDays(-90)
  $metricRows = foreach ($product in $products) {
    $count = @{ Assigned=0; Disabled=0; DisabledUnknown=0; AdEntraInactive=0; AdEntraUnknown=0; MailboxInactive=0; MailboxUnknown=0; M365Inactive=0; M365Unknown=0; LocalAppsInactive=0; LocalAppsUnknown=0; Multiple=0; MultipleUnknown=0; MultipleAll=0; MultipleAllUnknown=0; SharedLicensed=0; SharedUnder50=0; SharedEligible=0; SharedUnknown=0; RecoveryCandidates=0; RecoveryUnknown=0; RecoveryPrimaryPc=0; RecoveryPrimaryPcUnknown=0 }
    if (-not $licenseSource.Ready) { [pscustomobject]@{ Product=$product.Name; Counts=$count; Available=$false }; continue }
    $candidateUsers = @{}
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
      if ($sharedStatus -eq 'Eligible') { $candidateUsers[$userId] = $true }
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
        if ($sharedStatus -eq 'NotShared') { $candidateUsers[$userId] = $true }
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
    [pscustomobject]@{ Product=$product.Name; Counts=$count; Available=$true }
  }
  return [pscustomobject]@{ Rows=@($metricRows); Sources=$sources.ToArray(); SharedSourceReady=$mailboxes.Ready; IntuneSourceReady=$intuneSource.Ready; AdSourceForced=$adSource.Forced; LicenseSourceForced=$licenseSource.Forced; MailboxGap=$mailboxGap }
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

function Send-LicensesFocusedSummaryEmail {
  param(
    [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TenantRows,
    [Parameter(Mandatory)][string]$CollectedAtUtc,
    [string]$CsvFolderPath = '',
    [string]$ExpectedTenantKey = '',
    [switch]$Manual,
    [switch]$ForceAdCsvAnalysis,
    [switch]$BypassLicenseUsersReceipt
  )

  if (-not $Manual -and -not [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableLicenseSummaryEmail' -DefaultValue $true)) {
    WriteLog -Message 'Focused license summary email is disabled by configuration.' 'INFO'
    return
  }
  if (-not $Manual -and (Test-SmartM365MaxItemsMode)) {
    WriteLog -Message 'Focused license summary email skipped for a sampled inventory.' 'INFO'
    return
  }

  try {
    $mailTo = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
    $mailFrom = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($mailTo) -or [string]::IsNullOrWhiteSpace($mailFrom)) {
      throw 'A report recipient (To) and sender (From) are required for the focused license summary email.'
    }

    $summaryRows = @(Get-LicensesFocusedSummaryRows -TenantRows $TenantRows)
    $additionalRows = @(Get-LicensesAdditionalOverviewRows -TenantRows $TenantRows)
    $usage = $null
    if ($CsvFolderPath -and $ExpectedTenantKey) {
      try {
        $snapshotTime = [datetimeoffset]::Parse($CollectedAtUtc, [Globalization.CultureInfo]::InvariantCulture)
        $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $CsvFolderPath -ExpectedTenantKey $ExpectedTenantKey -LicenseSnapshotUtc $snapshotTime -ForceAdCsvAnalysis:$ForceAdCsvAnalysis -BypassLicenseUsersReceipt:$BypassLicenseUsersReceipt
      }
      catch { WriteLog -Message ("Focused license usage metrics unavailable: {0}" -f $_.Exception.Message) 'WARNING' }
    }
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
    $gap06Total = if ($gap06 -and $gap06.Available) { [string]$gap06.Total } else { 'N/D' }
    $gap06Universe = if ($gap06 -and $gap06.Available) { [string]$gap06.Universe } else { 'N/D' }
    $gap06Percent = if ($gap06 -and $gap06.Available -and $gap06.Universe -gt 0) { (([decimal]$gap06.Total * 100 / [decimal]$gap06.Universe).ToString('0.#', $percentCulture) + '%') } elseif ($gap06 -and $gap06.Available) { 'N/A' } else { 'N/D' }
    $gap06Other = if ($gap06 -and $gap06.Available) { [string]$gap06.OtherSkus } else { 'N/D' }
    $gap06None = if ($gap06 -and $gap06.Available) { [string]$gap06.NoSkus } else { 'N/D' }
    $gap06Guests = if ($gap06 -and $gap06.Available) { [string]$gap06.Guests } else { 'N/D' }
    $gap06MemberEnabled = if ($gap06 -and $gap06.Available) { [string]$gap06.MemberEnabled } else { 'N/D' }
    $gap06MemberDisabled = if ($gap06 -and $gap06.Available) { [string]$gap06.MemberDisabled } else { 'N/D' }
    $gap06Unknown = if ($gap06 -and $gap06.Available) { [string]$gap06.Unknown } else { 'N/D' }
    $gapNotes = @()
    foreach ($item in @(@{ Section='05'; Value=$gap05 }, @{ Section='06'; Value=$gap06 })) {
      if ($item.Value -and -not $item.Value.Available -and $item.Value.Reason) {
        $gapNotes += '<p style="margin:7px 0 0;font-size:11px;color:#9a3412;">Section {0} N/D: {1}</p>' -f `
          $item.Section, [System.Net.WebUtility]::HtmlEncode([string]$item.Value.Reason)
      }
    }
    $sourceRows = if ($usage) {
      foreach ($source in $usage.Sources) {
        $state = if ($source.Ready -and $source.Forced) { "PROVISIONAL ($($source.Date); $($source.Provenance))" }
                 elseif ($source.Ready) { "Ready ($($source.Date); $($source.Provenance))" }
                 else { "N/D ($($source.Reason))" }
        '<li>{0}: {1}</li>' -f [System.Net.WebUtility]::HtmlEncode($source.Name), [System.Net.WebUtility]::HtmlEncode($state)
      }
    } else { @('<li>Usage sources: N/D</li>') }
    $sourceNote = if ($Manual) { '<span style="color:#0f766e;font-weight:700;">Source: existing published CSV. No new inventory was run.</span>' } else { '<span style="color:#0f766e;font-weight:700;">Source: published inventory.</span>' }
    $adOverrideNote = if ($usage -and $usage.AdSourceForced) { '<div style="margin:0 0 16px;padding:11px 14px;background:#fff7ed;border-left:4px solid #d97706;font-size:12px;line-height:18px;color:#7c2d12;"><strong>Provisional AD/Entra indicator.</strong> The AD CSV was analyzed despite a rejected collector receipt. Confirm AD inventory completeness before using its inactivity counts for a license decision.</div>' } else { '' }
    $licenseOverrideNote = if ($usage -and $usage.LicenseSourceForced) { '<div style="margin:0 0 16px;padding:11px 14px;background:#fff7ed;border-left:4px solid #d97706;font-size:12px;line-height:18px;color:#7c2d12;"><strong>Provisional license assignment indicators.</strong> The license-users CSV is absent from the previous collector receipt. Confirm a new complete licensing collection before using recovery counts for a license decision.</div>' } else { '' }
    $subject = 'Microsoft 365 license overview and recovery'
    $tableStyle = 'width:100%;border-collapse:collapse;table-layout:fixed;font-family:Segoe UI,Arial,sans-serif;font-size:12px;line-height:17px;color:#334155;'
    $headStyle = 'padding:9px 8px;background:#eaf1f8;border-bottom:2px solid #cbd5e1;text-align:left;font-size:11px;line-height:15px;color:#334155;vertical-align:bottom;'
    $bodyHtml = @"
<div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:700px;margin:0 auto;">
  <p style="margin:0 0 5px;font-size:11px;letter-spacing:1px;font-weight:700;color:#0f766e;">MICROSOFT 365 &middot; LICENSE SUMMARY</p>
  <h1 style="margin:0 0 8px;font-size:24px;line-height:30px;color:#0f172a;">License overview and recovery</h1>
  <p style="margin:0 0 18px;font-size:12px;line-height:18px;color:#64748b;">$([System.Net.WebUtility]::HtmlEncode([string]$OrgDomain)) &nbsp;&middot;&nbsp; Snapshot $([System.Net.WebUtility]::HtmlEncode($CollectedAtUtc)) UTC</p>
  $adOverrideNote
  $licenseOverrideNote
  <h2 style="margin:0 0 10px;font-size:18px;line-height:24px;color:#0f172a;">License overview</h2>
  <p style="margin:0 0 5px;font-size:11px;line-height:16px;font-weight:700;color:#475569;">MICROSOFT 365 SUITES</p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#eff6f8;border:1px solid #cbdfe2;"><tr>$($suiteCards -join "`n")</tr></table>
  <p style="margin:14px 0 5px;font-size:11px;line-height:16px;font-weight:700;color:#475569;">COPILOT, DYNAMICS 365 AND POWER BI</p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#f8f5f1;border:1px solid #e5ddd2;"><tr>$($otherCards -join "`n")</tr></table>
  <p style="margin:7px 0 17px;font-size:11px;line-height:16px;color:#64748b;">Enabled and used counts come from the tenant subscription snapshot. Used means consumed license units, not measured app activity. Percent used = used / enabled; N/A means no enabled units. Dynamics 365 and Power BI add SKU units, not distinct users. Copilot covers Microsoft 365 Copilot; Dynamics 365 excludes sandbox, trial and preview SKUs; Power BI covers Pro and Premium Per User, excluding free Standard.</p>
  <h2 style="margin:0 0 8px;font-size:16px;line-height:22px;color:#0f172a;">01 &nbsp; License capacity</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Enabled units</th><th style="$headStyle">Consumed (used %)</th><th style="$headStyle">Available (free %)</th><th style="$headStyle">Assigned users</th><th style="$headStyle">Status</th></tr></thead><tbody>$($capacityRows -join "`n")</tbody></table>
  <h2 style="margin:24px 0 10px;font-size:18px;line-height:24px;color:#0f172a;">License recovery overview</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#eff6f8;border:1px solid #cbdfe2;">
    <tr>
      <td width="33%" style="width:33%;padding:8px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #7c3aed;"><tr><td style="padding:13px 12px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#6d28d9;">$topRecoveryE5</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates E5</div></td></tr></table></td>
      <td width="33%" style="width:33%;padding:8px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #2563eb;"><tr><td style="padding:13px 12px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#1d4ed8;">$topRecoveryE3</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates E3</div></td></tr></table></td>
      <td width="34%" style="width:34%;padding:8px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border-left:4px solid #0f766e;"><tr><td style="padding:13px 12px;"><div style="font-size:22px;line-height:27px;font-weight:700;color:#0f766e;">$topRecoveryF3F1</div><div style="font-size:12px;line-height:17px;color:#334155;">Recovery candidates F3/F1</div></td></tr></table></td>
    </tr>
  </table>
  <p style="margin:8px 0 20px;font-size:11px;line-height:16px;color:#64748b;">Each card counts license assignments for its suite. F3/F1 adds both suite counts; a user with both may count twice. N/D indicates unqualified users.</p>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">02 &nbsp; Recovery by license</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Recovery candidates</th><th style="$headStyle">Disabled users</th><th style="$headStyle">No M365 activity (90d)</th><th style="$headStyle">Candidates primary on Intune PC</th><th style="$headStyle">Multiple target suites</th></tr></thead><tbody>$($recoveryRows -join "`n")</tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">03 &nbsp; Activity and overlap</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">No AD or Entra activity (90d)</th><th style="$headStyle">No mailbox activity (90d)</th><th style="$headStyle">No local Apps use (180d)</th><th style="$headStyle">Multiple assigned SKUs</th></tr></thead><tbody>$($activityRows -join "`n")</tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">04 &nbsp; Licensed shared mailboxes</h2>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">License</th><th style="$headStyle">Shared mailboxes with target SKU</th><th style="$headStyle">Under 50 GB</th><th style="$headStyle">Removal candidates after archive and hold checks</th><th style="$headStyle">Not qualified</th></tr></thead><tbody>$($sharedRows -join "`n")</tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">05 &nbsp; User mailboxes without licence F1/F3/E3/E5</h2>
  <p style="margin:0 0 8px;padding:11px 13px;background:#eff6f8;border-left:4px solid #0f766e;font-size:12px;color:#334155;"><strong style="font-size:19px;color:#0f766e;">$gap05Total</strong> of $gap05Universe qualified EXO UserMailbox &nbsp;&middot;&nbsp; <strong>$gap05Percent</strong></p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">With other SKUs</th><th style="$headStyle">With no assigned SKU</th><th style="$headStyle">Not qualified</th></tr></thead><tbody><tr><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05Other</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05None</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap05Unknown</td></tr></tbody></table>
  <h2 style="margin:22px 0 8px;font-size:16px;line-height:22px;color:#0f172a;">06 &nbsp; Without User mailboxes + without licence F1/F3/E3/E5</h2>
  <p style="margin:0 0 8px;padding:11px 13px;background:#eff6f8;border-left:4px solid #2563eb;font-size:12px;color:#334155;"><strong style="font-size:19px;color:#1d4ed8;">$gap06Total</strong> of $gap06Universe qualified Entra accounts &nbsp;&middot;&nbsp; <strong>$gap06Percent</strong></p>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$tableStyle"><thead><tr><th style="$headStyle">Member enabled</th><th style="$headStyle">Member disabled</th><th style="$headStyle">Guests</th><th style="$headStyle">With other SKUs</th><th style="$headStyle">With no assigned SKU</th><th style="$headStyle">Not qualified</th></tr></thead><tbody><tr><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06MemberEnabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06MemberDisabled</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Guests</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Other</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06None</td><td style="padding:10px 8px;border-bottom:1px solid #e2e8f0;">$gap06Unknown</td></tr></tbody></table>
  $($gapNotes -join "`n")
  <div style="margin:22px 0 0;padding:14px 16px;background:#f8fafc;border-left:3px solid #94a3b8;font-size:11px;line-height:17px;color:#475569;">
    <strong style="color:#0f172a;">How to read this report</strong><br />
    Recovery candidates are distinct licensed users per product with a disabled non-shared account, no observed M365 activity in 90 days, or a qualifying shared mailbox under 50 GB. Disabled users and activity/overlap indicators exclude identified shared mailboxes; an unqualified mailbox type is N/D. M365 activity includes Exchange, OneDrive, SharePoint, Teams, Skype for Business and Yammer, plus qualified mailbox, email-action and Apps usage reports. Shared mailbox candidates exclude active archives and litigation or retention holds. The Intune PC column counts recovery candidates assigned as Primary User of a Windows device; this assignment does not prove recent PC use and does not change the recovery count. Having no primary Intune PC also does not prove that a license is unused. Review advanced compliance features, assignment path and the commercial contract before removing a license. Indicators overlap and must not be added together.<br /><br />
    F1 includes M365_F1 and M365_F1_COMM. Multiple assigned SKUs include add-ons, trials and free products. Multiple target suites count non-shared users assigned to at least two distinct F1/F3/E3/E5 suites; the two F1 SKU variants count as one suite. Neither count alone proves redundant seats. Local Apps usage applies only to E3/E5 and uses the available 180-day Windows/Mac report. Sections 05 and 06 count distinct EXO UserMailbox and Entra account IDs with none of the four target suites; other SKUs are allowed and shown separately. Section 05 percent uses qualified EXO UserMailbox as its denominator. Section 06 percent uses qualified Entra accounts after excluding identified shared, room and other technical mailbox accounts in EXO and Exchange on-premises; its Member enabled, Member disabled and Guest groups add up to the qualified total. On-premises mailboxes match by ObjectGUID/ImmutableId first, then by unique UPN; an unqualified on-premises source makes section 06 N/D. These are coverage indicators, not license recovery candidates. N/D means a source or identity cannot be qualified; sources older than 14 days are excluded.
  </div>
  <h2 style="margin:22px 0 8px;font-size:14px;line-height:20px;color:#334155;">Source freshness</h2>
  <ul style="margin:0;padding-left:18px;font-size:11px;line-height:18px;color:#64748b;">$($sourceRows -join "`n")</ul>
  <p style="margin:14px 0 0;font-size:11px;line-height:16px;color:#64748b;">$sourceNote</p>
</div>
"@
    $bodyRow = '<tr><td style="padding:18px 24px 22px 24px;font-size:13px;line-height:19px;color:#334155;">{0}</td></tr>' -f $bodyHtml
    $executionFooter = 'Host: {0} | Generated: {1}' -f [string]$env:COMPUTERNAME, (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    $mailBody = New-SmartM365EmailBody -Title $subject -Category 'SmartM365' -HostName '' -GeneratedAt '' -BodyHtml $bodyRow -Footer $executionFooter
    Send-SmartM365Mail -From $mailFrom -To $mailTo -Subject $subject -BodyHtml $mailBody -MailPurpose Report
    WriteLog -Message 'Microsoft 365 license overview and recovery email sent.' 'INFO'
  }
  catch {
    if ($Manual) { throw }
    WriteLog -Message ("Focused license summary email failed: {0}" -f $_.Exception.Message) 'WARNING'
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

  if ($latestPath -ine $currentPath) {
    Copy-LicensesCsvAtomically -SourcePath $timestampedPath -DestinationPath $latestPath
    [void]$global:csvGeneratedPaths.Add($latestPath)
    WriteLog -Message "CSV copied to global path: $latestPath" 'INFO'
  }

  $global:csvFilePath1 = $timestampedPath
  $global:csvFilePath2 = $currentPath
  $global:csvFilePath3 = $latestPath

  if ($global:RetentionMaxCSV -gt 0) {
    Remove-SmartM365TimestampedFilesOlderThan -FolderPath $CurrentOutputPath -FilePattern "$runBaseFileName`_*.csv" -RetentionDays 7 -LogFile $global:logTextFile
  }

  Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestPath | Out-Null
  WriteLog -Message 'Automatic WeeklyHistory publication skipped for this CSV; the caller will publish the validated dataset group.' 'INFO'
  return [pscustomobject]@{
    TimestampedPath = $timestampedPath
    CurrentPath     = $currentPath
    PublishedPath   = $latestPath
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
      if ($global:csvGeneratedPaths -and $global:csvGeneratedPaths.Contains($latestPath) -and (Test-Path -LiteralPath $latestPath -PathType Leaf)) {
        $weeklySourceFiles.Add($latestPath) | Out-Null
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
$ScriptVersion = "1.31"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LicensesCsvLogFolderPath' -DefaultValue $OutputPath
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
if ($SendLicenseSummaryEmailOnly) {
  if ($MaxItems -gt 0 -or $TopUsers -gt 0 -or $FastSample -or $InteractiveAuth) {
    throw '-SendLicenseSummaryEmailOnly cannot be combined with collection or sampling switches.'
  }
  if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { throw 'LatestCsvFolderPath is required for email-only mode.' }
  $tenantCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath 'M365_Licenses_Tenant.csv'
  $expectedTenantKey = if ($global:SmartM365TenantKey) { [string]$global:SmartM365TenantKey } else { $Tenant }
  $snapshot = Read-LicensesTenantSnapshot -Path $tenantCsvPath -ExpectedTenantKey $expectedTenantKey
  $global:AppId = [string]$AppId
  $global:TenantId = [string]$TenantId
  $global:Thumb = [string]$Thumb
  $global:Thumbprint = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumbprint' -DefaultValue $Thumb)
  $global:OrgDomain = [string]$OrgDomain
  Send-LicensesFocusedSummaryEmail -TenantRows $snapshot.Rows -CollectedAtUtc $snapshot.CollectedAtUtc -CsvFolderPath $LatestCsvFolderPath -ExpectedTenantKey $expectedTenantKey -Manual -ForceAdCsvAnalysis:$ForceAdCsvAnalysis -BypassLicenseUsersReceipt:$BypassLicenseUsersReceipt
  Write-Host ("License summary email sent from existing CSV: {0}" -f $tenantCsvPath)
  return
}
if ($BypassLicenseUsersReceipt) { throw '-BypassLicenseUsersReceipt requires -SendLicenseSummaryEmailOnly.' }
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

try {
  # ------------------------
  # Initialize script environment
  # ------------------------
  $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
  Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
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
  $userResolutionFailures = 0
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

  foreach ($u in $users) {
    $uIndex++
    $pct = if ($uTotal -gt 0) { [math]::Floor(($uIndex * 100.0) / $uTotal) } else { 100 }
    Write-Progress -Activity "Processing users" -Status "[$uIndex/$uTotal] $($u.UserPrincipalName)" -PercentComplete $pct

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

        # Compact user service-plan state fact and normalized service-plan catalog
        if ($ServicePlans) {
          $disabledIds = if ($disabledPlansBySku.ContainsKey($skuId)) { $disabledPlansBySku[$skuId] } else { @() }
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
      $userResolutionFailures++
      WriteLog -Message "Error processing user $($u.UserPrincipalName): $_" "ERROR"
    }
  }
  if ($servicePlanStateWriter) {
    Close-LicensesServicePlanStateWriter -Writer $servicePlanStateWriter
    $servicePlanStateWriter = $null
    $stateBuildingFile = Get-Item -LiteralPath $servicePlanStateBuildingPath -ErrorAction Stop
    WriteLog -Message ("Disk-backed compact service-plan source completed: Rows={0:N0}; Size={1:N1} MB; WorkingSet={2:N1} MB." -f $servicePlanRowCount, ($stateBuildingFile.Length / 1MB), ((Get-Process -Id $PID).WorkingSet64 / 1MB)) 'INFO'
  }
  Write-Progress -Activity "Processing users" -Completed -Status "Done"
  if ($userResolutionFailures -gt 0) {
    throw "License collection is incomplete: $userResolutionFailures user(s) could not be resolved. Current CSV publication is stopped."
  }

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

  $currentOperation = 'Complete licensing source receipt'
  Set-SmartM365CmdbSourceScope -CompleteScope ($MaxItems -eq 0 -and $TopUsers -eq 0) -Scope 'CMDB:skus,license_paths,plans,user_plans,groups'
  try { Complete-SmartM365CmdbSourceReceipt -Status 'Completed' | Out-Null }
  catch { WriteLog -Message ("Licensing source receipt could not be completed: {0}" -f $_.Exception.Message) 'WARNING' }

  $currentOperation = 'Send focused license summary email'
  Send-LicensesFocusedSummaryEmail -TenantRows $tenantRows.ToArray() -CollectedAtUtc $skusCollectedAtUtc -CsvFolderPath $LatestCsvFolderPath -ExpectedTenantKey $global:SmartM365TenantKey -ForceAdCsvAnalysis:$ForceAdCsvAnalysis

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
    -OutputPath $OutputPath

  WriteLog -Message "$TaskName completed."
  try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
  Complete-SmartM365ExecutionContext -Status Auto
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
  exit 1
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBq0DpZtWy6S3OF
# oPKKf6uY5/hnKzXnF4BFbgk8myWkcqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB7NX0etSzMnQoOGpf30/r6
# 1fpVOfrGI/rdVi2hIc2eqTANBgkqhkiG9w0BAQEFAASCAYA3ObbxIBhjv7CPl7Bq
# mQje+Wn7SQlo8TXAKDugXm/Vdq9HIlARbnV9zjIHFCKnVYFFEeAQK/cEIIYzO3MN
# 9Jlvjr2qDhv3AMosoud7SCzalnJKLZgpp0rEbJYRdEwl22cX1eO/JeynC3AQtCSI
# 2XDNtjh20r+E+H1GOUnnzmLIRg6NZm4MxEEX+KAllfMwCg33x8atVCjREqOnoPCM
# p7l6m68o9Hswy2nVyk2vdpoQqHG8Pqk3TMEW0mR4HtvyrFOKhZmQa6BGk1/zsVUA
# ZZLTW4lxBN/U/wgwozN7neMThyfzho0odQwYMIR+QJYvOWqWhBc9GAQmU0LBSCcZ
# 3mvHuW52cJXVahQLu4MXZPhR+dFZCO0cmVTMkRH62iOXwd8iovhGaAdfphTm3cGE
# XGonHpK0WNdz2KKmcwHYM78tmt9VBtVSiYMftu8JP/jEqyyBv3KhGga78FPzqRPz
# 00pVndI+ddLcD9/n1gM3pA/EmA4N196iLzxo3rRxQUgGSbY=
# SIG # End signature block
