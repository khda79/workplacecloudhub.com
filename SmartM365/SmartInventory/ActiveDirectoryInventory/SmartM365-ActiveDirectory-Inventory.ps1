<#
.SYNOPSIS
    Active Directory full inventory (OUs, Computers, Users, Groups, Contacts) across one or multiple domains.

.DESCRIPTION
    This script performs a comprehensive inventory of Active Directory across:
    - Organizational Units (OUs)
    - Computers
    - Users
    - Groups
    - Contacts

    It:
    - Discovers all domains in the forest or uses a subset passed via -TargetDomains
    - Exports detailed CSVs per domain in a "Not-CSV-Combined" folder
    - Combines all per-domain CSVs into global "AllDomains" CSVs
    - Analyzes duplicate UserPrincipalNames and SMTP proxy addresses across all domains
    - Uses the shared framework (SmartM365.Core / InitializeScriptEnvironment)
    - Logs to text + transcript
    - Copies combined CSVs to a local configuration path (LatestCsvFolderPath)
    - Cleans old CSV/log files
    - Sends an email notification in case of a global error (SendEmailHtmlReport)

.VERSION
1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory = $false)]
    [string[]]$TargetDomains,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)
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
$DomainFriendlyNames = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DomainFriendlyNames' -DefaultValue ([pscustomobject]@{})
$IntuneEnrollmentGroupPattern = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IntuneEnrollmentGroupPattern' -DefaultValue '*GG_INTUNE_ENROLLMENT*'
$Windows11UpgradeGroupPattern = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Windows11UpgradeGroupPattern' -DefaultValue '*GG_INTUNE_UPGRADEW11*'

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host ("Current PowerShell version: {0}" -f $PSVersionTable.PSVersion) -ForegroundColor Yellow
    exit 1
}

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host ("Failed to import SmartM365.Core module from '{0}' : {1}" -f $modulePath, $_) -ForegroundColor Red
    exit 1
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.0"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveDirectoryInventoryCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:LogTranscriptFile -Append

    WriteLog -Message ("Script environment initialized at {0}" -f $InitializeOutputPath)
    $OutputPath = $InitializeOutputPath
    WriteLog -Message ("Starting {0}" -f $TaskName)
    WriteLog -Message ("PowerShell Version: {0}" -f $PSVersionTable.PSVersion)
}
catch {
    Write-Host ("Initialization failed: {0}" -f $_) -ForegroundColor Red
    exit 1
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    WriteLog -Message "ActiveDirectory module imported successfully."
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null

    # ----------------------------------------------------------
    # RETRY CONFIGURATION
    # ----------------------------------------------------------
    $MaxRetries        = 3
    $RetryDelaySeconds = 30

    # ----------------------------------------------------------
    # TRANSIENT AD ERROR DETECTION
    # ----------------------------------------------------------
    function Test-IsTransientADError {
        param(
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        $ex = $ErrorRecord.Exception
        while ($null -ne $ex) {
            if ($ex -is [Microsoft.ActiveDirectory.Management.ADServerDownException]) { return $true }
            if ($ex.Message -match 'Unable to contact the server')                    { return $true }
            if ($ex.Message -match 'The server is not operational')                   { return $true }
            if ($ex.Message -match 'invalid enumeration context')                     { return $true }
            $ex = $ex.InnerException
        }
        return $false
    }

    $utcDate     = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $baseFolder  = Join-Path $OutputPath "Not-CSV-Combined"
    $tempFolder  = Join-Path $baseFolder $utcDate
    $null = New-Item -ItemType Directory -Path $tempFolder -Force

    WriteLog -Message ("Temporary per-domain export folder created: {0}" -f $tempFolder)

    function Get-DomainNameShort {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainName
        )

        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            return $null
        }

        $domainNameLower = $DomainName.Trim().ToLowerInvariant()
        $configuredName = $DomainFriendlyNames.PSObject.Properties[$domainNameLower]
        if ($null -ne $configuredName -and -not [string]::IsNullOrWhiteSpace([string]$configuredName.Value)) {
            return [string]$configuredName.Value
        }

        $dotPos = $domainNameLower.IndexOf('.')
        if ($dotPos -gt 0) {
            return $domainNameLower.Substring(0, $dotPos).ToUpperInvariant()
        }

        return $domainNameLower.ToUpperInvariant()
    }

    function Get-NormalizedDomainAndSam {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainNameShort,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$SamAccountName
        )

        $domainPart = if ([string]::IsNullOrWhiteSpace($DomainNameShort)) { '' } else { $DomainNameShort.Trim() }
        $samPart    = if ([string]::IsNullOrWhiteSpace($SamAccountName)) { '' } else { $SamAccountName.Trim() }

        if ($samPart.EndsWith('$')) {
            $samPart = $samPart.Substring(0, $samPart.Length - 1)
        }

        $value = "{0}\{1}" -f $domainPart, $samPart
        $value = $value.Trim().ToLowerInvariant()
        $value = $value -replace '\u00A0', ''
        $value = $value -replace ' ', ''
        $value = $value -replace '\t', ''

        return $value
    }

    function Convert-GuidToImmutableId {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$ObjectGuid
        )
        try {
            if ([string]::IsNullOrWhiteSpace($ObjectGuid)) {
                return $null
            }
            $guid  = [System.Guid]::Parse($ObjectGuid)
            $bytes = $guid.ToByteArray()
            return [System.Convert]::ToBase64String($bytes)
        }
        catch {
            return $null
        }
    }

    function Remove-OldFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [int]$OlderThanDays
        )

        if (-not (Test-Path -Path $Path)) {
            return
        }

        $limit = (Get-Date).AddDays(-$OlderThanDays)
        Get-ChildItem -Path $Path -Filter $Filter -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                    WriteLog -Message ("Deleted old file: {0}" -f $_.FullName)
                }
                catch {
                    WriteLog -Message ("Failed to delete old file '{0}': {1}" -f $_.FullName, $_)
                }
            }
    }

    function Combine-CsvFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourceFolder,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [string]$DestinationFile
        )

        $files = Get-ChildItem -Path $SourceFolder -Filter $Filter -File | Sort-Object Name
        if (-not $files) {
            WriteLog -Message ("No CSV files found for filter '{0}' in '{1}'" -f $Filter, $SourceFolder)
            return
        }

        $combined = foreach ($file in $files) {
            Import-Csv -Path $file.FullName
        }

        $combined | Export-Csv -Path $DestinationFile -NoTypeInformation -Encoding UTF8
        WriteLog -Message ("Combined {0} file(s) into '{1}'" -f $files.Count, $DestinationFile)
    }

    function Get-ADStringValue {
        param([object]$Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value.Count -eq 0) { return $null }
        return [string]($Value[0])
    }


    function Get-CountryNameFromCode {
        param(
            [Parameter(Mandatory = $true)]
            [string]$CountryCode
        )

        $CountryLookup = @{
            "AE" = "United Arab Emirates"
            "AT" = "Austria"
            "BE" = "Belgium"
            "BR" = "Brazil"
            "CH" = "Switzerland"
            "CL" = "Chile"
            "CN" = "China"
            "CZ" = "Czech Republic"
            "DE" = "Germany"
            "ES" = "Spain"
            "FR" = "France"
            "GB" = "United Kingdom"
            "HR" = "Croatia"
            "IE" = "Ireland"
            "IT" = "Italy"
            "LU" = "Luxembourg"
            "LV" = "Latvia"
            "MX" = "Mexico"
            "NL" = "Netherlands"
            "PL" = "Poland"
            "PT" = "Portugal"
            "SI" = "Slovenia"
            "UY" = "Uruguay"
        }

        if ($CountryLookup.ContainsKey($CountryCode.ToUpper())) {
            return $CountryLookup[$CountryCode.ToUpper()]
        }
        else {
            return "Country code not found"
        }
    }

    function Test-UserAccountControlFlag {
        param(
            [Parameter(Mandatory = $true)]
            [int]$UserAccountControlValue,

            [Parameter(Mandatory = $true)]
            [int]$FlagToCheck
        )

        return (($UserAccountControlValue -band $FlagToCheck) -eq $FlagToCheck)
    }


    function Test-GroupMembershipByName {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string[]]$GroupNames,

            [Parameter(Mandatory = $true)]
            [string]$GroupNameToFind
        )

        if ($null -eq $GroupNames -or $GroupNames.Count -eq 0) {
            return $false
        }

        return ($GroupNames -contains $GroupNameToFind)
    }


    function Get-ComputerGroupNames {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Computer,

            [Parameter(Mandatory = $true)]
            [string]$Server,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainSid,

            [Parameter(Mandatory = $true)]
            [bool]$ResolveNestedGroups,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupParentsByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameBySIDCache
        )

        if ($ResolveNestedGroups) {
            try {
                return @(Get-ADPrincipalGroupMembership -Identity $Computer -Server $Server | Select-Object -ExpandProperty Name)
            }
            catch {
                return @()
            }
        }

        $startDns = New-Object System.Collections.Generic.List[string]
        if ($Computer.MemberOf) {
            foreach ($memberDn in $Computer.MemberOf) {
                if (-not [string]::IsNullOrWhiteSpace([string]$memberDn)) {
                    [void]$startDns.Add([string]$memberDn)
                }
            }
        }

        if ($Computer.primaryGroupID -and $DomainSid) {
            try {
                $pgSid = ('{0}-{1}' -f $DomainSid, $Computer.primaryGroupID)
                $pgDn  = $null

                if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                    $pgDn = $GroupNameBySIDCache[$pgSid]
                }
                else {
                    $pgObj = Get-ADGroup -Identity $pgSid -Server $Server -Properties DistinguishedName
                    if ($pgObj -and $pgObj.DistinguishedName) {
                        $pgDn = [string]$pgObj.DistinguishedName
                        $GroupNameBySIDCache[$pgSid] = $pgDn
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                    [void]$startDns.Add($pgDn)
                }
            }
            catch { }
        }

        if ($startDns.Count -eq 0) {
            return @()
        }

        $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $queue   = New-Object System.Collections.Queue
        $names   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($dn in $startDns) {
            if ($visited.Add($dn)) {
                $queue.Enqueue($dn)
            }
        }

        while ($queue.Count -gt 0) {
            $gdn      = [string]$queue.Dequeue()
            $gName    = $null
            $gParents = $null

            if ($GroupNameByDNCache.ContainsKey($gdn)) {
                $gName = $GroupNameByDNCache[$gdn]
                if ($GroupParentsByDNCache.ContainsKey($gdn)) {
                    $gParents = $GroupParentsByDNCache[$gdn]
                }
            }
            else {
                try {
                    $gObj = Get-ADGroup -Identity $gdn -Server $Server -Properties Name, MemberOf
                    if ($gObj) {
                        $gName    = $gObj.Name
                        $gParents = @($gObj.MemberOf)
                        $GroupNameByDNCache[$gdn]    = $gName
                        $GroupParentsByDNCache[$gdn] = $gParents
                    }
                }
                catch {
                    $gParents = @()
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($gName)) {
                [void]$names.Add($gName)
            }

            if ($gParents) {
                foreach ($parentDn in $gParents) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$parentDn)) {
                        $parentDnString = [string]$parentDn
                        if ($visited.Add($parentDnString)) {
                            $queue.Enqueue($parentDnString)
                        }
                    }
                }
            }
        }

        return @($names)
    }


    # ----------------------------------------------------------
    # PATH VERIFICATION
    # ----------------------------------------------------------
    $destinationRootPath = $null
    try {
        $destinationRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''

        if ([string]::IsNullOrWhiteSpace($destinationRootPath)) {
            WriteLog -Message "WARNING: LatestCsvFolderPath not found in local configuration or returned empty. Combined CSV copy will be skipped."
            $destinationRootPath = $null
        }
        else {
            $destinationRootPath = $destinationRootPath.Trim()
            WriteLog -Message ("LatestCsvFolderPath resolved to: {0}" -f $destinationRootPath)
        }
    }
    catch {
        WriteLog -Message ("WARNING: Failed to resolve LatestCsvFolderPath: {0}. Combined CSV copy will be skipped." -f $_)
        $destinationRootPath = $null
    }

    if ($TargetDomains -and $TargetDomains.Count -gt 0) {
        $DomainsToProcess = $TargetDomains
        WriteLog -Message ("Using explicitly provided target domains: {0}" -f ($DomainsToProcess -join ', '))
    }
    else {
        try {
            $forest = Get-ADForest -ErrorAction Stop
            $DomainsToProcess = $forest.Domains
            WriteLog -Message ("Discovered forest domains: {0}" -f ($DomainsToProcess -join ', '))
        }
        catch {
            throw "Unable to retrieve forest domains. $_"
        }
    }

    foreach ($currentDomainName in $DomainsToProcess) {
        WriteLog -Message ("Starting inventory for domain '{0}'" -f $currentDomainName)

        $domainAttempt = 0
        $domainSuccess = $false

        while (-not $domainSuccess -and $domainAttempt -lt $MaxRetries) {
            $domainAttempt++
            if ($domainAttempt -gt 1) {
                WriteLog -Message ("Retrying inventory for domain '{0}' (attempt {1}/{2}) after {3}s delay..." -f $currentDomainName, $domainAttempt, $MaxRetries, $RetryDelaySeconds)
                Start-Sleep -Seconds $RetryDelaySeconds
            }

        try {

        $safeDomainFileName = $currentDomainName -replace '[^a-zA-Z0-9\.-]', '_'

        # ------------------------------------------------------
        # OU INVENTORY
        # ------------------------------------------------------
        try {
            $CurrentObjectType = "OUs"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_OUs_{0}.csv" -f $safeDomainFileName)

            Get-ADOrganizationalUnit -Filter * -Server $currentDomainName -Properties Name, DistinguishedName, description, managedBy |
                Select-Object `
                    @{Name = 'DomainName';   Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';   Expression = { $CurrentObjectType }},
                    Name,
                    DistinguishedName,
                    @{Name = 'Description'; Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    managedBy |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported OUs for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, (Import-Csv $outputCsvFilePath).Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("OU inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }

        # ------------------------------------------------------
        # COMPUTER INVENTORY
        # ------------------------------------------------------
        try {
            $CurrentObjectType = "Computers"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Computers_{0}.csv" -f $safeDomainFileName)

            $EnablePCFilter = $true
            if ($EnablePCFilter) {
                $computerFilter = {
                    (OperatingSystem -like "*Windows*" -and OperatingSystem -notlike "*Server*") -or (OperatingSystem -notlike "*")
                }
            }
            else {
                $computerFilter = { $true }
            }

            $GroupNameByDNCache     = @{}
            $GroupParentsByDNCache = @{}
            $GroupNameBySIDCache   = @{}

            try {
                $domainObj = Get-ADDomain -Server $currentDomainName
                $domainSid = $domainObj.DomainSID.Value
            }
            catch {
                $domainSid = $null
            }

            $ResolveNestedComputerGroups = $false

            Get-ADComputer -Filter $computerFilter -Server $currentDomainName -Properties SamAccountName, Name, DistinguishedName, Enabled, DNSHostName, OperatingSystem, operatingSystemHotfix, operatingSystemServicePack, operatingSystemVersion, LastLogonDate, LastLogonTimestamp, Description, IPv4Address, WhenCreated, WhenChanged, pwdLastSet, CanonicalName, MemberOf, primaryGroupID, ObjectGUID, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.DNSHostName) } |
                ForEach-Object {
                    $computer = $_
                    $computerGroupNames = Get-ComputerGroupNames -Computer $computer -Server $currentDomainName -DomainSid $domainSid -ResolveNestedGroups:$ResolveNestedComputerGroups -GroupNameByDNCache $GroupNameByDNCache -GroupParentsByDNCache $GroupParentsByDNCache -GroupNameBySIDCache $GroupNameBySIDCache
                    $hasGroupAddIntune  = [bool]($computerGroupNames | Where-Object { $_ -like $IntuneEnrollmentGroupPattern })
                    $hasGroupUpgradeW11 = [bool]($computerGroupNames | Where-Object { $_ -like $Windows11UpgradeGroupPattern })

                    [PSCustomObject][ordered]@{
                        DomainName              = $currentDomainName
                        ObjectType              = $CurrentObjectType
                        SamAccountName          = $computer.SamAccountName
                        Name                    = $computer.Name
                        DistinguishedName       = $computer.DistinguishedName
                        Enabled                 = $computer.Enabled
                        DNSHostName             = $computer.DNSHostName
                        OperatingSystem         = $computer.OperatingSystem
                        operatingSystemHotfix   = $computer.operatingSystemHotfix
                        operatingSystemServicePack = $computer.operatingSystemServicePack
                        operatingSystemVersion  = $computer.operatingSystemVersion
                        LastLogonDate           = if ($computer.LastLogonTimestamp -ne $null -and $computer.LastLogonTimestamp -ne 0) { [datetime]::FromFileTime($computer.LastLogonTimestamp) } else { '' }
                        Description             = $computer.Description -replace "`r", " -R " -replace "`n", " -N "
                        IPv4Address             = $computer.IPv4Address
                        WhenCreated             = $computer.WhenCreated
                        WhenChanged             = $computer.WhenChanged
                        pwdLastSetDate          = if ($computer.pwdLastSet -ne $null -and $computer.pwdLastSet -ne 0) { [datetime]::FromFileTime($computer.pwdLastSet) } else { '' }
                        CanonicalName           = $computer.CanonicalName -replace "`r", " -R " -replace "`n", " -N "
                        MemberOfDNs             = if ($computer.MemberOf) { ($computer.MemberOf) -join ';' } else { '' }
                        PrimaryGroupName        = if ($computer.primaryGroupID -and $domainSid) {
                            $pgSid = ('{0}-{1}' -f $domainSid, $computer.primaryGroupID)
                            if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                                $pgDn = $GroupNameBySIDCache[$pgSid]
                                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                                    if ($GroupNameByDNCache.ContainsKey($pgDn)) {
                                        $GroupNameByDNCache[$pgDn]
                                    }
                                    else {
                                        try {
                                            $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName
                                            $GroupNameByDNCache[$pgDn] = $g.Name
                                            $g.Name
                                        }
                                        catch { '' }
                                    }
                                }
                                else {
                                    ''
                                }
                            }
                            else {
                                try {
                                    $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName -Properties DistinguishedName
                                    if ($g -and $g.DistinguishedName) {
                                        $GroupNameBySIDCache[$pgSid] = [string]$g.DistinguishedName
                                        $GroupNameByDNCache[[string]$g.DistinguishedName] = $g.Name
                                    }
                                    $g.Name
                                }
                                catch { '' }
                            }
                        } else { '' }
                        ObjectGUID              = $computer.ObjectGUID
                        SID                     = $computer.ObjectSID.Value
                        SIDHistory              = if ($computer.SIDHistory) { ($computer.SIDHistory | ForEach-Object { $_.Value }) -join ';' } else { '' }
                        extensionAttribute1     = Get-ADStringValue $computer.extensionAttribute1
                        extensionAttribute2     = Get-ADStringValue $computer.extensionAttribute2
                        extensionAttribute3     = Get-ADStringValue $computer.extensionAttribute3
                        extensionAttribute4     = Get-ADStringValue $computer.extensionAttribute4
                        extensionAttribute5     = Get-ADStringValue $computer.extensionAttribute5
                        extensionAttribute6     = Get-ADStringValue $computer.extensionAttribute6
                        extensionAttribute7     = Get-ADStringValue $computer.extensionAttribute7
                        extensionAttribute8     = Get-ADStringValue $computer.extensionAttribute8
                        extensionAttribute9     = Get-ADStringValue $computer.extensionAttribute9
                        extensionAttribute10    = Get-ADStringValue $computer.extensionAttribute10
                        extensionAttribute11    = Get-ADStringValue $computer.extensionAttribute11
                        extensionAttribute12    = Get-ADStringValue $computer.extensionAttribute12
                        extensionAttribute13    = Get-ADStringValue $computer.extensionAttribute13
                        extensionAttribute14    = Get-ADStringValue $computer.extensionAttribute14
                        extensionAttribute15    = Get-ADStringValue $computer.extensionAttribute15
                        DomainNameShort         = Get-DomainNameShort -DomainName $currentDomainName
                        DomainAndSam            = Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $computer.SamAccountName
                        ImmutableId_AD          = Convert-GuidToImmutableId -ObjectGuid ([string]$computer.ObjectGUID)
                        Has_Group_AddIntune     = $hasGroupAddIntune
                        Has_Group_UpgradeW11    = $hasGroupUpgradeW11
                    }
                } |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Computers for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, (Import-Csv $outputCsvFilePath).Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Computer inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }

        # ------------------------------------------------------
        # USER INVENTORY
        # ------------------------------------------------------
        try {
            $CurrentObjectType = "Users"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Users_{0}.csv" -f $safeDomainFileName)

            [int]$DomainExcludedUsersNoUpn = 0
            try {
                $DomainExcludedUsersNoUpn = (Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(!(userPrincipalName=*)))" -Server $currentDomainName -ResultSetSize $null).Count
            }
            catch {
                WriteLog -Message ("WARNING: Failed to count users without UPN for domain '{0}': {1}" -f $currentDomainName, $_)
                $DomainExcludedUsersNoUpn = 0
            }
            WriteLog -Message ("Users excluded because of missing UPN for domain '{0}': {1}" -f $currentDomainName, $DomainExcludedUsersNoUpn)

            Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userPrincipalName=*))" -Server $currentDomainName -ResultSetSize $null -Properties SamAccountName, sAMAccountType, Name, DistinguishedName, UserPrincipalName, Enabled, manager, LastLogonTimestamp, DisplayName, GivenName, Surname, Description, Department, Title, Company, Office, TelephoneNumber, MobilePhone, EmailAddress, StreetAddress, City, PostalCode, Country, WhenCreated, WhenChanged, AccountExpirationDate, pwdLastSet, badPwdCount, badPasswordTime, LogonCount, userAccountControl, msDS-ManagedPassword, ProxyAddresses, MemberOf, CanonicalName, ObjectGUID, targetAddress, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Select-Object `
                    @{Name = 'DomainName';           Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';           Expression = { $CurrentObjectType }},
                    SamAccountName,
                    sAMAccountType,
                    @{Name = 'Name';                Expression = { $_.Name        -replace "`r", " -R " -replace "`n", " -N " }},
                    DistinguishedName,
                    UserPrincipalName,
                    Enabled,
                    manager,
                    LastLogonTimestamp,
                    @{Name = 'LastLogonDate';       Expression = {
                        if ($_.LastLogonTimestamp -eq 0 -or $_.LastLogonTimestamp -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.LastLogonTimestamp)
                        }
                    }},
                    @{Name = 'DisplayName';          Expression = { $_.DisplayName  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'GivenName';            Expression = { $_.GivenName    -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Surname';             Expression = { $_.Surname     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Description';         Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Department';          Expression = { $_.Department  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Title';               Expression = { $_.Title       -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Company';             Expression = { $_.Company     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Office';              Expression = { $_.Office      -replace "`r", " -R " -replace "`n", " -N " }},
                    TelephoneNumber,
                    MobilePhone,
                    EmailAddress,
                    @{Name = 'StreetAddress';       Expression = { $_.StreetAddress -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'City';                Expression = { $_.City          -replace "`r", " -R " -replace "`n", " -N " }},
                    PostalCode,
                    Country,
                    @{Name = 'CountryName';         Expression = {
                        if (-not [string]::IsNullOrWhiteSpace($_.Country)) {
                            Get-CountryNameFromCode -CountryCode $_.Country
                        }
                        else {
                            "Unknown"
                        }
                    }},
                    WhenCreated,
                    WhenChanged,
                    AccountExpirationDate,
                    @{Name = 'PasswordLastSetDate'; Expression = {
                        if ($_.pwdLastSet -eq 0 -or $_.pwdLastSet -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.pwdLastSet)
                        }
                    }},
                    badPwdCount,
                    @{Name = 'BadPasswordDate';     Expression = {
                        if ($_.badPasswordTime -eq 0 -or $_.badPasswordTime -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.badPasswordTime)
                        }
                    }},
                    LogonCount,
                    userAccountControl,
                    @{Name = 'IsNormalAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 64 }},
                    @{Name = 'IsScriptAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 1 }},
                    @{Name = 'IsPasswordNeverExpires';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536 }},
                    @{Name = 'IsTrustedForDelegation';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 524288 }},
                    @{Name = 'MemberOfGroups';          Expression = { $_.MemberOf      -join ";" }},
                    @{Name = 'ProxyAddresses';          Expression = { $_.ProxyAddresses -join ";" }},
                    @{Name = 'CanonicalName';           Expression = { $_.CanonicalName -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'ObjectGUID';              Expression = { $_.ObjectGUID }},
                    @{Name = 'TargetAddress';           Expression = { $_.targetAddress }},
                    @{Name = 'ObjectSID';               Expression = { $_.ObjectSID.Value }},
                    @{Name = 'ObjectSIDHistory';        Expression = {
                        if ($_.SIDHistory) {
                            ($_.SIDHistory | ForEach-Object { $_.Value }) -join ";"
                        }
                        else {
                            ""
                        }
                    }},
                    @{Name = 'extensionAttribute1';          Expression = { Get-ADStringValue $_.extensionAttribute1 }},
                    @{Name = 'extensionAttribute2';          Expression = { Get-ADStringValue $_.extensionAttribute2 }},
                    @{Name = 'extensionAttribute3';          Expression = { Get-ADStringValue $_.extensionAttribute3 }},
                    @{Name = 'extensionAttribute4';          Expression = { Get-ADStringValue $_.extensionAttribute4 }},
                    @{Name = 'extensionAttribute5';          Expression = { Get-ADStringValue $_.extensionAttribute5 }},
                    @{Name = 'extensionAttribute6';          Expression = { Get-ADStringValue $_.extensionAttribute6 }},
                    @{Name = 'extensionAttribute7';          Expression = { Get-ADStringValue $_.extensionAttribute7 }},
                    @{Name = 'extensionAttribute8';          Expression = { Get-ADStringValue $_.extensionAttribute8 }},
                    @{Name = 'extensionAttribute9';          Expression = { Get-ADStringValue $_.extensionAttribute9 }},
                    @{Name = 'extensionAttribute10';         Expression = { Get-ADStringValue $_.extensionAttribute10 }},
                    @{Name = 'extensionAttribute11';         Expression = { Get-ADStringValue $_.extensionAttribute11 }},
                    @{Name = 'extensionAttribute12';         Expression = { Get-ADStringValue $_.extensionAttribute12 }},
                    @{Name = 'extensionAttribute13';         Expression = { Get-ADStringValue $_.extensionAttribute13 }},
                    @{Name = 'extensionAttribute14';         Expression = { Get-ADStringValue $_.extensionAttribute14 }},
                    @{Name = 'extensionAttribute15';         Expression = { Get-ADStringValue $_.extensionAttribute15 }},
                    @{Name = 'MustChangePasswordAtNextLogon'; Expression = { ($_.pwdLastSet -eq 0) -and (-not (Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536)) }},
                    @{Name = 'DomainNameShort';         Expression = { Get-DomainNameShort -DomainName $currentDomainName }},
                    @{Name = 'DomainAndSam';            Expression = { Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $_.SamAccountName }},
                    @{Name = 'ImmutableId_AD';          Expression = { Convert-GuidToImmutableId -ObjectGuid ([string]$_.ObjectGUID) }} |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Users for domain '{0}' to '{1}'. Count: {2}. Excluded without UPN: {3}" -f $currentDomainName, $outputCsvFilePath, (Import-Csv $outputCsvFilePath).Count, $DomainExcludedUsersNoUpn)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("User inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }

        # ------------------------------------------------------
        # GROUP INVENTORY
        # ------------------------------------------------------
        try {
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Groups_{0}.csv" -f $safeDomainFileName)
            $GroupData = @(Get-ADGroup -Filter * -Server $currentDomainName -Properties CanonicalName, CN, Created, createTimeStamp, Deleted, Description, DisplayName, DistinguishedName, GroupCategory, GroupScope, GroupType, HomePage, LastKnownParent, mail, ManagedBy, MemberOf, Members, Modified, modifyTimeStamp, Name, ObjectCategory, ObjectClass, ObjectGUID, objectSid, ProtectedFromAccidentalDeletion, SamAccountName, SIDHistory, whenChanged, whenCreated |
                Select-Object `
                    @{Name = 'DomainName'; Expression = { [string]$currentDomainName }},
                    @{Name = 'CanonicalName'; Expression = { Get-ADStringValue $_.CanonicalName }},
                    @{Name = 'CN'; Expression = { Get-ADStringValue $_.CN }},
                    @{Name = 'Created'; Expression = { $_.Created }},
                    @{Name = 'createTimeStamp'; Expression = { $_.createTimeStamp }},
                    @{Name = 'Deleted'; Expression = { $_.Deleted }},
                    @{Name = 'Description'; Expression = { Get-ADStringValue $_.Description }},
                    @{Name = 'DisplayName'; Expression = { Get-ADStringValue $_.DisplayName }},
                    @{Name = 'DistinguishedName'; Expression = { Get-ADStringValue $_.DistinguishedName }},
                    @{Name = 'GroupCategory'; Expression = { Get-ADStringValue $_.GroupCategory }},
                    @{Name = 'GroupScope'; Expression = { Get-ADStringValue $_.GroupScope }},
                    @{Name = 'GroupType'; Expression = { $_.GroupType }},
                    @{Name = 'HomePage'; Expression = { Get-ADStringValue $_.HomePage }},
                    @{Name = 'LastKnownParent'; Expression = { Get-ADStringValue $_.LastKnownParent }},
                    @{Name = 'mail'; Expression = { Get-ADStringValue $_.mail }},
                    @{Name = 'ManagedBy'; Expression = { Get-ADStringValue $_.ManagedBy }},
                    @{Name = 'MemberOf'; Expression = { if ($_.MemberOf) { ($_.MemberOf -join ';') } else { '' } }},
                    @{Name = 'Members'; Expression = { if ($_.Members) { ($_.Members -join ';') } else { '' } }},
                    @{Name = 'Modified'; Expression = { $_.Modified }},
                    @{Name = 'modifyTimeStamp'; Expression = { $_.modifyTimeStamp }},
                    @{Name = 'Name'; Expression = { Get-ADStringValue $_.Name }},
                    @{Name = 'ObjectCategory'; Expression = { Get-ADStringValue $_.ObjectCategory }},
                    @{Name = 'ObjectClass'; Expression = { Get-ADStringValue $_.ObjectClass }},
                    @{Name = 'ObjectGUID'; Expression = { if ($_.ObjectGUID) { $_.ObjectGUID.Guid } else { $null } }},
                    @{Name = 'objectSid'; Expression = { $_.objectSid.Value }},
                    @{Name = 'ProtectedFromAccidentalDeletion'; Expression = { $_.ProtectedFromAccidentalDeletion }},
                    @{Name = 'SamAccountName'; Expression = { Get-ADStringValue $_.SamAccountName }},
                    @{Name = 'SIDHistory'; Expression = { if ($_.SIDHistory) { ($_.SIDHistory.Value) -join ';' } else { '' } }},
                    @{Name = 'whenChanged'; Expression = { $_.whenChanged }},
                    @{Name = 'whenCreated'; Expression = { $_.whenCreated }}
            )
            $GroupData | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8
            WriteLog -Message ("Exported Groups for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $GroupData.Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Group inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }

        # ------------------------------------------------------
        # CONTACT INVENTORY
        # ------------------------------------------------------
        try {
            $CurrentObjectType = "Contacts"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Contacts_{0}.csv" -f $safeDomainFileName)

            Get-ADObject -Filter { ObjectClass -eq "contact" } -Server $currentDomainName -Properties DisplayName, ProxyAddresses, Mail |
                Select-Object `
                    @{Name = 'DomainName';      Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';      Expression = { $CurrentObjectType }},
                    DisplayName,
                    @{Name = 'ProxyAddresses'; Expression = { $_.ProxyAddresses -join ";" }},
                    Mail |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Contacts for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, (Import-Csv $outputCsvFilePath).Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Contact inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }

        $domainSuccess = $true

        } # end outer domain try
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) {
                if ($domainAttempt -lt $MaxRetries) {
                    WriteLog -Message ("WARNING: Transient AD connectivity error for domain '{0}' (attempt {1}/{2}): {3}" -f $currentDomainName, $domainAttempt, $MaxRetries, $_.Exception.Message)
                    # Loop continues: next iteration will wait and retry
                }
                else {
                    WriteLog -Message ("ERROR: Transient AD connectivity error for domain '{0}' persisted after {1} attempt(s). Skipping domain. Error: {2}" -f $currentDomainName, $domainAttempt, $_.Exception.Message)
                }
            }
            else {
                WriteLog -Message ("FATAL: Unhandled error for domain '{0}' (attempt {1}/{2}): {3}. Continuing with next domain." -f $currentDomainName, $domainAttempt, $MaxRetries, $_)
                break
            }
        }

        } # end while retry loop

        if (-not $domainSuccess) {
            WriteLog -Message ("WARNING: Domain '{0}' could not be fully inventoried after {1} attempt(s). Skipping." -f $currentDomainName, $domainAttempt)
        }
    }

    # ------------------------------------------------------
    # COMBINE PER-DOMAIN CSV FILES
    # ------------------------------------------------------
    $combinedUsersCsv     = Join-Path $OutputPath "AD_Users_AllDomains.csv"
    $combinedComputersCsv = Join-Path $OutputPath "AD_Computers_AllDomains.csv"
    $combinedGroupsCsv    = Join-Path $OutputPath "AD_Groups_AllDomains.csv"
    $combinedOusCsv       = Join-Path $OutputPath "AD_OUs_AllDomains.csv"
    $combinedContactsCsv  = Join-Path $OutputPath "AD_Contacts_AllDomains.csv"

    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Users_*.csv"     -DestinationFile $combinedUsersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Computers_*.csv" -DestinationFile $combinedComputersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Groups_*.csv"    -DestinationFile $combinedGroupsCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_OUs_*.csv"       -DestinationFile $combinedOusCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Contacts_*.csv"  -DestinationFile $combinedContactsCsv

    # ------------------------------------------------------
    # Copy combined CSVs to the latest CSV folder
    # ------------------------------------------------------
    try {
        if ($destinationRootPath) {
            if (-not (Test-Path -LiteralPath $destinationRootPath)) {
                New-Item -Path $destinationRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $destinationRootPath)
            }
            foreach ($combinedCsv in @($combinedUsersCsv, $combinedComputersCsv, $combinedGroupsCsv, $combinedOusCsv, $combinedContactsCsv)) {
                if (Test-Path -Path $combinedCsv) {
                    $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($combinedCsv))
                    Copy-Item -LiteralPath $combinedCsv -Destination $destinationFile -Force -ErrorAction Stop
                    WriteLog -Message ("Copied combined CSV '{0}' to '{1}'" -f $combinedCsv, $destinationFile)
                    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
                }
                else {
                    WriteLog -Message ("Combined CSV not found, skipping copy: {0}" -f $combinedCsv)
                }
            }
        }
        else {
            WriteLog -Message "LatestCsvFolderPath unavailable. Combined CSV copy skipped."
        }
    }
    catch {
        WriteLog -Message ("Copy to LatestCsvFolderPath failed: {0}" -f $_)
    }

    # ------------------------------------------------------
    # DUPLICATE UPN ANALYSIS
    # ------------------------------------------------------
    try {
        WriteLog -Message "Starting duplicate UserPrincipalName analysis..."

        $duplicateUpnCsv = Join-Path $OutputPath "AD_Users_DuplicateUPN.csv"

        if (-not (Test-Path -Path $combinedUsersCsv)) {
            WriteLog -Message ("WARNING: Combined users CSV not found, skipping duplicate UPN analysis: {0}" -f $combinedUsersCsv)
        }
        else {
            $allUsers = Import-Csv -Path $combinedUsersCsv -Encoding UTF8
            WriteLog -Message ("Loaded {0} users from combined CSV for duplicate UPN analysis." -f $allUsers.Count)

            $upnGroups = $allUsers |
                Group-Object { $_.UserPrincipalName.ToLowerInvariant() } |
                Where-Object { $_.Count -gt 1 }

            $duplicateUpnRows = foreach ($grp in $upnGroups) {
                foreach ($u in $grp.Group) {
                    [PSCustomObject][ordered]@{
                        UserPrincipalName   = $u.UserPrincipalName
                        UPN_OccurrenceCount = $grp.Count
                        DomainName          = $u.DomainName
                        DomainNameShort     = $u.DomainNameShort
                        SamAccountName      = $u.SamAccountName
                        DisplayName         = $u.DisplayName
                        Enabled             = $u.Enabled
                        DistinguishedName   = $u.DistinguishedName
                    }
                }
            }

            $duplicateUpnRows |
                Sort-Object UserPrincipalName, DomainName |
                Export-Csv -Path $duplicateUpnCsv -NoTypeInformation -Encoding UTF8

            $upnDuplicateCount   = ($upnGroups   | Measure-Object).Count
            $upnAffectedAccounts = ($duplicateUpnRows | Measure-Object).Count
            WriteLog -Message ("Duplicate UPN analysis complete. Distinct duplicate UPNs: {0}. Affected accounts: {1}. Output: {2}" -f $upnDuplicateCount, $upnAffectedAccounts, $duplicateUpnCsv)

            if ($destinationRootPath -and (Test-Path -Path $duplicateUpnCsv)) {
                $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($duplicateUpnCsv))
                Copy-Item -LiteralPath $duplicateUpnCsv -Destination $destinationFile -Force -ErrorAction Stop
                WriteLog -Message ("Copied '{0}' to '{1}'" -f $duplicateUpnCsv, $destinationFile)
                Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
            }
        }
    }
    catch {
        WriteLog -Message ("Duplicate UPN analysis failed: {0}" -f $_)
    }

    # ------------------------------------------------------
    # DUPLICATE SMTP PROXY ADDRESS ANALYSIS
    # ------------------------------------------------------
    try {
        WriteLog -Message "Starting duplicate SMTP proxy address analysis..."

        $duplicateSmtpCsv = Join-Path $OutputPath "AD_Users_DuplicateSMTP.csv"

        if (-not (Test-Path -Path $combinedUsersCsv)) {
            WriteLog -Message ("WARNING: Combined users CSV not found, skipping duplicate SMTP analysis: {0}" -f $combinedUsersCsv)
        }
        else {
            # $allUsers may already be loaded from UPN analysis above; reload defensively
            if ($null -eq $allUsers) {
                $allUsers = Import-Csv -Path $combinedUsersCsv -Encoding UTF8
                WriteLog -Message ("Loaded {0} users from combined CSV for duplicate SMTP analysis." -f $allUsers.Count)
            }

            # Expand ProxyAddresses: one row per smtp: entry per user
            $smtpExpanded = foreach ($u in $allUsers) {
                if ([string]::IsNullOrWhiteSpace($u.ProxyAddresses)) { continue }
                foreach ($entry in ($u.ProxyAddresses -split ';')) {
                    $entry = $entry.Trim()
                    if ($entry -notmatch '^smtp:(.+)$') { continue }
                    [PSCustomObject]@{
                        SmtpAddress_Lower = $Matches[1].ToLowerInvariant()
                        IsUppercaseSMTP   = $entry.StartsWith('SMTP:')
                        SmtpAddress       = $Matches[1]
                        UserPrincipalName = $u.UserPrincipalName
                        DomainName        = $u.DomainName
                        DomainNameShort   = $u.DomainNameShort
                        SamAccountName    = $u.SamAccountName
                        DisplayName       = $u.DisplayName
                        Enabled           = $u.Enabled
                        DistinguishedName = $u.DistinguishedName
                    }
                }
            }

            $smtpGroups = $smtpExpanded |
                Group-Object SmtpAddress_Lower |
                Where-Object { $_.Count -gt 1 }

            $duplicateSmtpRows = foreach ($grp in $smtpGroups) {
                foreach ($e in $grp.Group) {
                    [PSCustomObject][ordered]@{
                        SmtpAddress          = $e.SmtpAddress
                        SMTP_OccurrenceCount = $grp.Count
                        IsUppercaseSMTP      = $e.IsUppercaseSMTP
                        UserPrincipalName    = $e.UserPrincipalName
                        DomainName           = $e.DomainName
                        DomainNameShort      = $e.DomainNameShort
                        SamAccountName       = $e.SamAccountName
                        DisplayName          = $e.DisplayName
                        Enabled              = $e.Enabled
                        DistinguishedName    = $e.DistinguishedName
                    }
                }
            }

            $duplicateSmtpRows |
                Sort-Object SmtpAddress, DomainName |
                Export-Csv -Path $duplicateSmtpCsv -NoTypeInformation -Encoding UTF8

            $smtpDuplicateCount   = ($smtpGroups        | Measure-Object).Count
            $smtpAffectedEntries  = ($duplicateSmtpRows  | Measure-Object).Count
            WriteLog -Message ("Duplicate SMTP analysis complete. Distinct duplicate addresses: {0}. Affected entries: {1}. Output: {2}" -f $smtpDuplicateCount, $smtpAffectedEntries, $duplicateSmtpCsv)

            if ($destinationRootPath -and (Test-Path -Path $duplicateSmtpCsv)) {
                $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($duplicateSmtpCsv))
                Copy-Item -LiteralPath $duplicateSmtpCsv -Destination $destinationFile -Force -ErrorAction Stop
                WriteLog -Message ("Copied '{0}' to '{1}'" -f $duplicateSmtpCsv, $destinationFile)
                Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
            }
        }
    }
    catch {
        WriteLog -Message ("Duplicate SMTP analysis failed: {0}" -f $_)
    }

    # ------------------------------------------------------
    # CLEANUP OLD FILES
    # ------------------------------------------------------
    Remove-OldFiles -Path $OutputPath -Filter "*.csv" -OlderThanDays 30
    Remove-OldFiles -Path $OutputPath -Filter "*.log" -OlderThanDays 30

    WriteLog -Message ("{0} completed successfully." -f $TaskName)
}
catch {
    $globalError = $_
    WriteLog -Message ("Fatal error in script: {0}" -f $globalError)

    try {
        $emailSubject = "[ERROR] $TaskName"
        $emailBody = @"
<html>
<body>
    <h3>Active Directory Inventory - Global Error</h3>
    <p><b>Script:</b> $($MyInvocation.MyCommand.Name)</p>
    <p><b>Version:</b> $ScriptVersion</p>
    <p><b>Host:</b> $env:COMPUTERNAME</p>
    <p><b>Date:</b> $(Get-Date)</p>
    <p><b>Error:</b></p>
    <pre>$globalError</pre>
</body>
</html>
"@
        SendEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody
        WriteLog -Message "Global error email notification sent."
    }
    catch {
        WriteLog -Message ("Failed to send global error email notification: {0}" -f $_)
    }
}
finally {
    try {
        WriteLog -Message ("Stopping transcript for script '{0}'" -f $MyInvocation.MyCommand.Name)
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    }
    catch {
        Write-Host ("Failed to stop transcript: {0}" -f $_) -ForegroundColor Yellow
    }
}
