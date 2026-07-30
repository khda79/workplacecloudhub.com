#Requires -Version 7.0
<#
.SYNOPSIS
    Exports Exchange Online quarantined email metadata and sends an Excel report.

.DESCRIPTION
    Connects to Exchange Online with certificate-based app-only authentication,
    retrieves email entities that are still blocked in quarantine, publishes a
    timestamped and latest CSV, creates an Excel workbook, and optionally sends
    the workbook through the shared SmartM365 mail transport.

    This script is read-only. It never previews, exports, releases, reports, or
    deletes quarantined messages.

.PARAMETER Tenant
    SmartM365 tenant profile key. Defaults to test.

.PARAMETER Connect
    Retained for launcher consistency. The script always disconnects any
    existing Exchange Online session before establishing its app-only session.

.PARAMETER LookbackDays
    Overrides the configured quarantine lookback period. Valid values are 1-30.

.PARAMETER MaxItems
    Limits collected rows for a bounded test. MaxItems runs use suffixed,
    timestamped artifacts and skip latest publication, SharePoint, weekly
    history, and report email.

.PARAMETER NoMail
    Generates and publishes the report without sending the report email.

.VERSION
    1.0.2

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ExchangeOnlineManagement; ImportExcel.
    Runtime: Exchange.ManageAsApp plus a supported read-only role that can use
    Get-QuarantineMessage (Global Reader is the SmartM365 baseline).
    Conditional: Mail.Send is required only when Graph mail is used.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'SmartM365.Core uses global execution context variables for logs, configuration, and generated CSV tracking.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'The final console summary is intentional for interactive runs.')]
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [ValidateRange(0, 30)]
    [int]$LookbackDays = 0,
    [ValidateRange(0, 1000000)]
    [int]$MaxItems = 0,
    [switch]$NoMail,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '1.0.2'
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$script:TaskName = "$script:ScriptBaseName v$script:ScriptVersion"
$script:CurrentOperation = 'Initialize'
$script:RunId = [guid]::NewGuid().ToString()
$script:RunStarted = Get-Date
$script:ExchangeConnected = $false
$script:TranscriptStarted = $false
$script:TimestampedCsvPath = ''
$script:LatestCsvPath = ''
$script:WorkbookPath = ''
$script:ReportRowCount = 0
$script:FailureRecord = $null

if ($MaxItems -gt 0) {
    $global:SmartM365MaxItems = $MaxItems
    $global:SmartM365TestMaxItems = $MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

$tenantContextPath = & {
    $directory = $PSScriptRoot
    while ($directory) {
        foreach ($candidate in @(
            (Join-Path -Path $directory -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $directory -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $directory -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $directory) { break }
        $directory = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

. $tenantContextPath
$script:TenantContext = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$contextDirectory = Split-Path -Path $tenantContextPath -Parent
$script:SmartM365Root = if ((Split-Path -Path $contextDirectory -Leaf) -ieq 'Config') {
    Split-Path -Path $contextDirectory -Parent
}
else {
    $contextDirectory
}

$coreModulePath = Join-Path -Path $script:SmartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $coreModulePath -MinimumVersion '1.0.37' -Force -ErrorAction Stop

$script:LocalConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "$script:ScriptBaseName.local.json"
$script:LocalTemplatePath = "$script:LocalConfigPath.template"
if (-not (Test-Path -LiteralPath $script:LocalConfigPath)) {
    Initialize-SmartM365LocalJsonFromTemplate `
        -Path $script:LocalConfigPath `
        -TemplatePath $script:LocalTemplatePath `
        -ConfigDescription 'script local configuration' | Out-Null
}
$script:ScriptConfig = Get-Content -LiteralPath $script:LocalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

function Resolve-QuarantineConfigToken {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($Value -isnot [string]) { return $Value }
    $resolved = $Value
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        foreach ($match in $tokenMatches) {
            $property = $script:TenantContext.PSObject.Properties[$match.Groups['Name'].Value]
            if ($property -and $null -ne $property.Value) {
                $resolved = $resolved.Replace($match.Value, [string]$property.Value)
            }
        }
    }
    return $resolved
}

function Get-QuarantineConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$DefaultValue
    )

    $localProperty = $script:ScriptConfig.PSObject.Properties[$Name]
    if ($localProperty -and $null -ne $localProperty.Value) {
        $localValue = $localProperty.Value
        if ($localValue -isnot [string] -or (
            -not [string]::IsNullOrWhiteSpace($localValue) -and
            $localValue.Trim() -notin @('__USE_GLOBAL__', 'USE_GLOBAL')
        )) {
            return Resolve-QuarantineConfigToken -Value $localValue
        }
    }

    $globalProperty = $script:TenantContext.PSObject.Properties[$Name]
    if ($globalProperty -and $null -ne $globalProperty.Value) {
        return Resolve-QuarantineConfigToken -Value $globalProperty.Value
    }

    return Resolve-QuarantineConfigToken -Value $DefaultValue
}

function ConvertTo-QuarantineBoolean {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Value,
        [bool]$DefaultValue = $false
    )

    if ($null -eq $Value) { return $DefaultValue }
    if ($Value -is [bool]) { return $Value }
    $parsed = $false
    if ([bool]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }
    return $DefaultValue
}

function Get-QuarantinePropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $InputObject) { return $null }
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($property) { return $property.Value }
    }
    return $null
}

function Join-QuarantineValue {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return (@($Value) |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique) -join '; '
}

function ConvertTo-QuarantineUtcText {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        return ([datetime]$Value).ToUniversalTime().ToString(
            'yyyy-MM-ddTHH:mm:ssZ',
            [Globalization.CultureInfo]::InvariantCulture
        )
    }
    catch {
        return ''
    }
}

function Protect-QuarantineSpreadsheetText {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ($text -match '^[=\+\-@\t\r]') { return "'$text" }
    return $text
}

function ConvertTo-QuarantineReportRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Message,
        [Parameter(Mandatory)][datetime]$CollectedAtUtc
    )

    $receivedValue = Get-QuarantinePropertyValue -InputObject $Message -Names @('ReceivedTime', 'ReceivedTimeUtc', 'Received')
    $receivedUtc = $null
    if ($null -ne $receivedValue -and -not [string]::IsNullOrWhiteSpace([string]$receivedValue)) {
        try { $receivedUtc = ([datetime]$receivedValue).ToUniversalTime() } catch { $receivedUtc = $null }
    }

    $ageDays = ''
    if ($null -ne $receivedUtc) {
        $ageDays = [math]::Max(0, [math]::Floor(($CollectedAtUtc - $receivedUtc).TotalDays))
    }

    $releasedValue = Get-QuarantinePropertyValue -InputObject $Message -Names @('Released')
    $systemReleasedValue = Get-QuarantinePropertyValue -InputObject $Message -Names @('SystemReleased')
    $releaseStatus = [string](Get-QuarantinePropertyValue -InputObject $Message -Names @('ReleaseStatus'))
    if ([string]::IsNullOrWhiteSpace($releaseStatus)) {
        if ((ConvertTo-QuarantineBoolean -Value $releasedValue) -or (ConvertTo-QuarantineBoolean -Value $systemReleasedValue)) {
            $releaseStatus = 'Released'
        }
        else {
            $releaseStatus = 'NotReleased'
        }
    }

    return [pscustomobject][ordered]@{
        RunId             = $script:RunId
        CollectedAtUtc    = $CollectedAtUtc.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
        ReceivedTimeUtc   = ConvertTo-QuarantineUtcText -Value $receivedValue
        AgeDays           = $ageDays
        SenderAddress     = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('SenderAddress', 'Sender'))
        RecipientAddress  = Protect-QuarantineSpreadsheetText -Value (Join-QuarantineValue -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('RecipientAddress', 'Recipients', 'QuarantineUser')))
        Subject            = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('Subject'))
        QuarantineType     = Protect-QuarantineSpreadsheetText -Value (Join-QuarantineValue -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('QuarantineTypes', 'Type')))
        PolicyName         = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('PolicyName'))
        PolicyType         = Protect-QuarantineSpreadsheetText -Value (Join-QuarantineValue -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('PolicyTypes', 'PolicyType')))
        Direction          = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('Direction'))
        ReleaseStatus      = Protect-QuarantineSpreadsheetText -Value $releaseStatus
        ExpiresUtc         = ConvertTo-QuarantineUtcText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('Expires', 'ExpiresTime', 'ExpirationTime'))
        Reported           = [string](ConvertTo-QuarantineBoolean -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('Reported')))
        SystemReleased     = [string](ConvertTo-QuarantineBoolean -Value $systemReleasedValue)
        MessageId          = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('MessageId', 'InternetMessageId'))
        Identity           = Protect-QuarantineSpreadsheetText -Value (Get-QuarantinePropertyValue -InputObject $Message -Names @('Identity'))
    }
}

function Get-SmartM365BlockedQuarantineMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$StartReceivedDate,
        [Parameter(Mandatory)][datetime]$EndReceivedDate,
        [Parameter(Mandatory)][string[]]$ReleaseStatuses,
        [bool]$IncludeBlockedSenders = $true,
        [ValidateRange(0, 1000000)][int]$Limit = 0,
        [scriptblock]$CommandInvoker = {
            param([hashtable]$Parameters)
            Get-QuarantineMessage @Parameters
        },
        [string[]]$SupportedParameterNames,
        [switch]$NoRetryDelay
    )

    $usesDefaultCommandInvoker = -not $PSBoundParameters.ContainsKey('CommandInvoker')
    if ($usesDefaultCommandInvoker) {
        $quarantineCommand = Get-Command -Name 'Get-QuarantineMessage' -ErrorAction Stop
        $SupportedParameterNames = @($quarantineCommand.Parameters.Keys)
    }
    elseif (-not $PSBoundParameters.ContainsKey('SupportedParameterNames')) {
        # Mock/custom invokers preserve the complete modern parameter contract by default.
        $SupportedParameterNames = @('EntityType', 'IncludeMessagesFromBlockedSenderAddress')
    }

    $supportsEntityType = $SupportedParameterNames -contains 'EntityType'
    $supportsBlockedSenderFilter = $SupportedParameterNames -contains 'IncludeMessagesFromBlockedSenderAddress'
    if (-not $supportsEntityType) {
        WriteLog -Message 'Get-QuarantineMessage does not expose -EntityType; using its email-only/default entity behavior.' -Level 'INFO'
    }
    if ($IncludeBlockedSenders -and -not $supportsBlockedSenderFilter) {
        WriteLog -Message 'Get-QuarantineMessage does not expose -IncludeMessagesFromBlockedSenderAddress; using its default blocked-sender behavior.' -Level 'INFO'
    }

    $pageSize = 1000
    $messages = [System.Collections.Generic.List[object]]::new()

    for ($page = 1; $page -le 1000; $page++) {
        $parameters = @{
            StartReceivedDate = $StartReceivedDate
            EndReceivedDate   = $EndReceivedDate
            ReleaseStatus     = $ReleaseStatuses
            PageSize          = $pageSize
            Page              = $page
            ErrorAction       = 'Stop'
        }
        if ($supportsEntityType) {
            $parameters['EntityType'] = 'Email'
        }
        if ($IncludeBlockedSenders -and $supportsBlockedSenderFilter) {
            $parameters['IncludeMessagesFromBlockedSenderAddress'] = $true
        }

        $pageRows = @()
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            try {
                $pageRows = @(& $CommandInvoker $parameters)
                break
            }
            catch {
                $isTransient = $_.Exception.Message -match 'throttl|TooManyRequests|temporar|timeout|429|500|502|503|504'
                if (-not $isTransient -or $attempt -eq 5) { throw }
                $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt) * 5)
                WriteLog -Message (
                    'Quarantine page {0} transient failure on attempt {1}/5; retry in {2}s: {3}' -f
                    $page, $attempt, $delaySeconds, $_.Exception.Message
                ) -Level 'WARNING'
                if (-not $NoRetryDelay) { Start-Sleep -Seconds $delaySeconds }
            }
        }

        WriteLog -Message ("Quarantine page {0}: {1} row(s)." -f $page, $pageRows.Count) -Level 'INFO'
        foreach ($message in $pageRows) {
            if ($null -eq $message) { continue }
            [void]$messages.Add($message)
            if ($Limit -gt 0 -and $messages.Count -ge $Limit) {
                return @($messages.ToArray() | Select-Object -First $Limit)
            }
        }

        if ($pageRows.Count -lt $pageSize) {
            return $messages.ToArray()
        }

        if ($page -eq 1000) {
            throw 'Quarantine pagination reached the supported 1000-page boundary before a final partial page was returned.'
        }
    }

    return $messages.ToArray()
}

function Export-SmartM365QuarantineWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$DetailColumns,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][datetime]$StartReceivedDateUtc,
        [Parameter(Mandatory)][datetime]$EndReceivedDateUtc,
        [Parameter(Mandatory)][int]$EffectiveLookbackDays
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    }

    $reasonGroups = @($Rows |
        Group-Object -Property QuarantineType |
        Sort-Object -Property @(
            @{ Expression = 'Count'; Descending = $true },
            @{ Expression = 'Name'; Descending = $false }
        ))
    $summaryRows = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @(
        [pscustomobject]@{ Metric = 'Tenant'; Value = [string]$global:SmartM365TenantKey },
        [pscustomobject]@{ Metric = 'Generated UTC'; Value = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') },
        [pscustomobject]@{ Metric = 'Lookback days'; Value = $EffectiveLookbackDays },
        [pscustomobject]@{ Metric = 'Window start UTC'; Value = $StartReceivedDateUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') },
        [pscustomobject]@{ Metric = 'Window end UTC'; Value = $EndReceivedDateUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') },
        [pscustomobject]@{ Metric = 'Blocked messages'; Value = $Rows.Count },
        [pscustomobject]@{ Metric = 'Oldest message age days'; Value = $(if ($Rows.Count -gt 0) { ($Rows | Measure-Object -Property AgeDays -Maximum).Maximum } else { 0 }) }
    )) {
        [void]$summaryRows.Add($item)
    }

    $summaryRows.ToArray() | Export-Excel `
        -Path $Path `
        -WorksheetName 'Summary' `
        -TableName 'QuarantineSummary' `
        -TableStyle Medium2 `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow `
        -AutoFilter

    if ($reasonGroups.Count -gt 0) {
        $reasonRows = @($reasonGroups | ForEach-Object {
            [pscustomobject]@{
                QuarantineType = $(if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { $_.Name })
                Count          = $_.Count
            }
        })
        $reasonRows | Export-Excel `
            -Path $Path `
            -WorksheetName 'By_Reason' `
            -TableName 'QuarantineByReason' `
            -TableStyle Medium4 `
            -AutoSize `
            -FreezeTopRow `
            -BoldTopRow `
            -AutoFilter
    }

    if ($Rows.Count -gt 0) {
        $Rows |
            Select-Object -Property $DetailColumns |
            Export-Excel `
                -Path $Path `
                -WorksheetName 'Quarantine_Messages' `
                -TableName 'QuarantineMessages' `
                -TableStyle Medium2 `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow `
                -AutoFilter
    }
    else {
        $package = Open-ExcelPackage -Path $Path
        try {
            $worksheet = $package.Workbook.Worksheets.Add('Quarantine_Messages')
            for ($columnIndex = 0; $columnIndex -lt $DetailColumns.Count; $columnIndex++) {
                $cell = $worksheet.Cells.Item(1, $columnIndex + 1)
                $cell.Value = $DetailColumns[$columnIndex]
                $cell.Style.Font.Bold = $true
            }
            $worksheet.View.FreezePanes(2, 1)
            $worksheet.Cells.Item(1, 1, 1, $DetailColumns.Count).AutoFilter = $true
            $worksheet.Cells.AutoFitColumns(8, 60)
            Close-ExcelPackage -ExcelPackage $package
            $package = $null
        }
        finally {
            if ($null -ne $package) { $package.Dispose() }
        }
    }

    return $Path
}

function Format-SmartM365QuarantineMailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][datetime]$StartReceivedDateUtc,
        [Parameter(Mandatory)][datetime]$EndReceivedDateUtc,
        [Parameter(Mandatory)][string]$WorkbookPath
    )

    $topReasons = @($Rows |
        Group-Object -Property QuarantineType |
        Sort-Object -Property @(
            @{ Expression = 'Count'; Descending = $true },
            @{ Expression = 'Name'; Descending = $false }
        ) |
        Select-Object -First 8)
    $reasonRows = if ($topReasons.Count -eq 0) {
        '<tr><td colspan="2" style="padding:8px;border-bottom:1px solid #e2e8f0;">No blocked messages</td></tr>'
    }
    else {
        @($topReasons | ForEach-Object {
            $reason = [System.Net.WebUtility]::HtmlEncode($(if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { $_.Name }))
            '<tr><td style="padding:8px;border-bottom:1px solid #e2e8f0;">{0}</td><td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;">{1}</td></tr>' -f $reason, $_.Count
        }) -join [Environment]::NewLine
    }

    $workbookName = [System.Net.WebUtility]::HtmlEncode((Split-Path -Path $WorkbookPath -Leaf))
    return @"
<!DOCTYPE html>
<html>
<body style="font-family:Segoe UI,Arial,sans-serif;color:#1f2937;">
  <h2 style="color:#0f4c81;">Exchange Online quarantine report</h2>
  <p>The attached Excel workbook contains read-only metadata for email messages that remain blocked in quarantine.</p>
  <table role="presentation" style="border-collapse:collapse;width:100%;max-width:720px;">
    <tr><td style="padding:8px;font-weight:600;">Tenant</td><td style="padding:8px;">$([System.Net.WebUtility]::HtmlEncode([string]$global:SmartM365TenantKey))</td></tr>
    <tr><td style="padding:8px;font-weight:600;">Blocked messages</td><td style="padding:8px;">$($Rows.Count)</td></tr>
    <tr><td style="padding:8px;font-weight:600;">Window UTC</td><td style="padding:8px;">$($StartReceivedDateUtc.ToString('yyyy-MM-dd HH:mm')) to $($EndReceivedDateUtc.ToString('yyyy-MM-dd HH:mm'))</td></tr>
    <tr><td style="padding:8px;font-weight:600;">Workbook</td><td style="padding:8px;">$workbookName</td></tr>
  </table>
  <h3 style="color:#0f4c81;">Top quarantine reasons</h3>
  <table style="border-collapse:collapse;width:100%;max-width:720px;border:1px solid #e2e8f0;">
    <tr><th style="padding:8px;text-align:left;background:#f8fafc;">Reason</th><th style="padding:8px;text-align:right;background:#f8fafc;">Count</th></tr>
    $reasonRows
  </table>
  <p style="font-size:12px;color:#64748b;">No message body or attachment content is collected by this report.</p>
</body>
</html>
"@
}

function Send-QuarantineTeamsSuccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Count,
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string]$XlsxPath
    )

    $summary = "Exchange Online quarantine report completed. Blocked messages: $Count."
    Send-SmartM365TeamsNotification `
        -Title 'SmartM365 EXO quarantine report completed' `
        -Message $summary `
        -Level 'SUCCESS' `
        -Channel 'Infos' `
        -ResultSummary $summary `
        -Facts @{
            'Script name'         = [System.IO.Path]::GetFileName($PSCommandPath)
            'Tenant/Organization' = [string]$script:OrgDomain
            'Computer'            = $env:COMPUTERNAME
            'Timestamp'           = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
            'Blocked messages'    = $Count
            'CSV path'            = $CsvPath
            'Excel path'          = $XlsxPath
            'Log path'            = $global:LogTextFile
        } | Out-Null
}

function Send-QuarantineTeamsFailure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [Parameter(Mandatory)][string]$Operation
    )

    $exception = $ErrorRecord.Exception
    $innerMessages = [System.Collections.Generic.List[string]]::new()
    $innerException = $exception.InnerException
    while ($null -ne $innerException) {
        if (-not [string]::IsNullOrWhiteSpace($innerException.Message)) {
            [void]$innerMessages.Add($innerException.Message)
        }
        $innerException = $innerException.InnerException
    }

    $helpText = @(
        "Script: $([System.IO.Path]::GetFileName($PSCommandPath))"
        "Tenant: $script:OrgDomain"
        "Operation: $Operation"
        "Error: $($exception.Message)"
    ) -join "`n"

    Send-SmartM365TeamsNotification `
        -Title 'SmartM365 EXO quarantine report failed' `
        -Message 'A terminal error occurred in the Exchange Online quarantine report.' `
        -Level 'ERROR' `
        -Channel 'Alerts' `
        -Facts @{
            'Script name'         = [System.IO.Path]::GetFileName($PSCommandPath)
            'Tenant/Organization' = [string]$script:OrgDomain
            'Computer'            = $env:COMPUTERNAME
            'Timestamp'           = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
            'Failed operation'    = $Operation
            'Exception message'   = $exception.Message
            'Inner exception'     = ($innerMessages -join ' | ')
            'Log path'            = $global:LogTextFile
            'Transcript path'     = $global:logTranscriptFile
            'CSV path'            = $script:TimestampedCsvPath
            'Excel path'          = $script:WorkbookPath
        } `
        -HelpUrl ("https://chat.openai.com/?q={0}" -f [uri]::EscapeDataString("Help troubleshoot this SmartM365 error:`n$helpText")) | Out-Null
}

$script:AppId = [string](Get-QuarantineConfigValue -Name 'AppId' -DefaultValue '')
$script:TenantId = [string](Get-QuarantineConfigValue -Name 'TenantId' -DefaultValue '')
$script:OrgDomain = [string](Get-QuarantineConfigValue -Name 'OrgDomain' -DefaultValue '')
$script:Thumbprint = [string](Get-QuarantineConfigValue -Name 'Thumbprint' -DefaultValue (Get-QuarantineConfigValue -Name 'Thumb' -DefaultValue ''))
$script:LatestCsvFolderPath = [string](Get-QuarantineConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
$script:OutputPathFromConfig = [string](Get-QuarantineConfigValue -Name 'QuarantineCsvLogFolderPath' -DefaultValue '')
$script:ConfiguredLookbackDays = [int](Get-QuarantineConfigValue -Name 'LookbackDays' -DefaultValue 30)
$script:ReleaseStatuses = @(
    Get-QuarantineConfigValue `
        -Name 'ReleaseStatuses' `
        -DefaultValue @('NotReleased', 'Requested', 'Denied', 'Error', 'PreparingToRelease')
) | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
$script:IncludeBlockedSenders = ConvertTo-QuarantineBoolean `
    -Value (Get-QuarantineConfigValue -Name 'IncludeMessagesFromBlockedSenderAddress' -DefaultValue $true) `
    -DefaultValue $true
$script:SendReportMail = ConvertTo-QuarantineBoolean `
    -Value (Get-QuarantineConfigValue -Name 'SendReportMail' -DefaultValue $true) `
    -DefaultValue $true

$global:AppId = $script:AppId
$global:TenantId = $script:TenantId
$global:OrgDomain = $script:OrgDomain
$global:Thumb = $script:Thumbprint
$global:Thumbprint = $script:Thumbprint
$global:RetentionMaxCSV = [int](Get-QuarantineConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-QuarantineConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = ConvertTo-QuarantineBoolean -Value (Get-QuarantineConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
if ($MaxItems -gt 0) {
    # A bounded test must not publish report artifacts or execution logs.
    $global:EnableSharePointUpload = $false
}
$global:SharePointSiteHostname = [string](Get-QuarantineConfigValue -Name 'SharePointSiteHostname' -DefaultValue '')
$global:SharePointSitePath = [string](Get-QuarantineConfigValue -Name 'SharePointSitePath' -DefaultValue '')
$global:SharePointLibraryDisplayName = [string](Get-QuarantineConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents')
$global:SharePointTargetFolderPath = [string](Get-QuarantineConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue '')
$global:EnableWeeklyHistory = ConvertTo-QuarantineBoolean -Value (Get-QuarantineConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true) -DefaultValue $true
$global:WeeklyHistoryFolderPath = [string](Get-QuarantineConfigValue -Name 'WeeklyHistoryFolderPath' -DefaultValue (Join-Path -Path $script:OutputPathFromConfig -ChildPath 'WeeklyHistory'))
$global:WeeklyHistoryRetentionWeeks = [int](Get-QuarantineConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
$global:SmtpServer = [string](Get-QuarantineConfigValue -Name 'SmtpServer' -DefaultValue '')
$global:SendMailMode = [string](Get-QuarantineConfigValue -Name 'SendMailMode' -DefaultValue '')
$global:From = [string](Get-QuarantineConfigValue -Name 'From' -DefaultValue '')
$global:To = [string](Get-QuarantineConfigValue -Name 'To' -DefaultValue '')
$global:Cc = [string](Get-QuarantineConfigValue -Name 'Cc' -DefaultValue '')
$global:ErrorMailTo = [string](Get-QuarantineConfigValue -Name 'ErrorMailTo' -DefaultValue '')

if ($LookbackDays -eq 0) { $LookbackDays = $script:ConfiguredLookbackDays }
if ($LookbackDays -lt 1 -or $LookbackDays -gt 30) {
    throw "Effective LookbackDays must be between 1 and 30. Current value: $LookbackDays"
}
if ($script:ReleaseStatuses.Count -eq 0) {
    throw 'ReleaseStatuses must contain at least one non-empty quarantine status.'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = $script:OutputPathFromConfig }

$detailColumns = @(
    'TenantKey',
    'RunId',
    'CollectedAtUtc',
    'ReceivedTimeUtc',
    'AgeDays',
    'SenderAddress',
    'RecipientAddress',
    'Subject',
    'QuarantineType',
    'PolicyName',
    'PolicyType',
    'Direction',
    'ReleaseStatus',
    'ExpiresUtc',
    'Reported',
    'SystemReleased',
    'MessageId',
    'Identity'
)
$csvColumns = @($detailColumns | Where-Object { $_ -ne 'TenantKey' })

try {
    $script:CurrentOperation = 'Initialize script environment'
    $OutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $script:ScriptBaseName -CallerScriptPath $PSCommandPath
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    $script:TranscriptStarted = $true
    Write-SmartM365ExecutionContext -ScriptPath $PSCommandPath -OutputPath $OutputPath
    WriteLog -Message ("Starting {0}. Tenant={1}; LookbackDays={2}; MaxItems={3}; NoMail={4}; Connect={5}" -f $script:TaskName, $Tenant, $LookbackDays, $MaxItems, $NoMail, $Connect) -Level 'INFO'

    $script:CurrentOperation = 'Validate configuration'
    $configurationErrors = [System.Collections.Generic.List[string]]::new()
    foreach ($requiredValue in @(
        @{ Name = 'AppId'; Value = $script:AppId },
        @{ Name = 'TenantId'; Value = $script:TenantId },
        @{ Name = 'OrgDomain'; Value = $script:OrgDomain },
        @{ Name = 'Thumbprint'; Value = $script:Thumbprint }
    )) {
        if ([string]::IsNullOrWhiteSpace([string]$requiredValue.Value)) {
            [void]$configurationErrors.Add("Missing required configuration value: $($requiredValue.Name)")
        }
    }
    if ($MaxItems -eq 0 -and [string]::IsNullOrWhiteSpace($script:LatestCsvFolderPath)) {
        [void]$configurationErrors.Add('Missing required configuration value: LatestCsvFolderPath')
    }
    $reportMailRequired = -not $NoMail -and $script:SendReportMail -and $MaxItems -eq 0
    if ($reportMailRequired -and [string]::IsNullOrWhiteSpace($global:From)) {
        [void]$configurationErrors.Add('Missing required report mail configuration value: From')
    }
    if ($reportMailRequired -and [string]::IsNullOrWhiteSpace($global:To)) {
        [void]$configurationErrors.Add('Missing required report mail configuration value: To')
    }
    if ($configurationErrors.Count -gt 0) {
        throw ('SmartM365 quarantine configuration error(s): {0}' -f ($configurationErrors -join '; '))
    }

    $script:CurrentOperation = 'Load required PowerShell modules'
    $preflightOutputPaths = @($OutputPath)
    if ($MaxItems -eq 0) { $preflightOutputPaths += $script:LatestCsvFolderPath }
    Invoke-SmartM365Preflight `
        -ScriptName $script:TaskName `
        -RequiredModules @('ExchangeOnlineManagement', 'ImportExcel') `
        -RequiredCommands @('Connect-ExchangeOnline', 'Export-Excel', 'Open-ExcelPackage', 'Close-ExcelPackage') `
        -OutputPaths $preflightOutputPaths | Out-Null

    $script:CurrentOperation = 'Disconnect existing Exchange Online session'
    Disconnect-SmartM365CloudSession -ExchangeOnline:$true -Graph:$false -VerboseDisconnect:$true

    $script:CurrentOperation = 'Connect to Exchange Online'
    $connection = Connect-SmartM365CloudSession `
        -AppId $script:AppId `
        -Thumbprint $script:Thumbprint `
        -TenantId $script:TenantId `
        -Organization $script:OrgDomain `
        -ExchangeOnline:$true `
        -Graph:$false
    if (-not $connection.ExchangeOnlineConnected) {
        throw 'Exchange Online app-only connection failed.'
    }
    $script:ExchangeConnected = $true

    $script:CurrentOperation = 'Validate Get-QuarantineMessage authorization'
    if (-not (Get-Command -Name Get-QuarantineMessage -ErrorAction SilentlyContinue)) {
        throw 'Get-QuarantineMessage is unavailable in the connected Exchange Online session.'
    }
    Invoke-SmartM365Preflight `
        -ScriptName $script:TaskName `
        -OutputPaths @($OutputPath) `
        -ExchangeOnlineProbeCommands @('Get-QuarantineMessage') | Out-Null

    $collectedAtUtc = (Get-Date).ToUniversalTime()
    $endReceivedDateUtc = $collectedAtUtc
    $startReceivedDateUtc = $endReceivedDateUtc.AddDays(-$LookbackDays)

    $script:CurrentOperation = 'Retrieve blocked quarantine messages'
    WriteLog -Message (
        'Retrieving quarantined email. StartUtc={0}; EndUtc={1}; Statuses={2}; IncludeBlockedSenders={3}.' -f
        $startReceivedDateUtc.ToString('o'),
        $endReceivedDateUtc.ToString('o'),
        ($script:ReleaseStatuses -join ','),
        $script:IncludeBlockedSenders
    ) -Level 'INFO'
    $rawMessages = @(
        Get-SmartM365BlockedQuarantineMessage `
            -StartReceivedDate $startReceivedDateUtc `
            -EndReceivedDate $endReceivedDateUtc `
            -ReleaseStatuses $script:ReleaseStatuses `
            -IncludeBlockedSenders $script:IncludeBlockedSenders `
            -Limit $MaxItems
    )

    $script:CurrentOperation = 'Normalize quarantine message metadata'
    $reportRows = @($rawMessages |
        ForEach-Object { ConvertTo-QuarantineReportRow -Message $_ -CollectedAtUtc $collectedAtUtc } |
        Where-Object { $_.ReleaseStatus -ne 'Released' -and $_.SystemReleased -ne 'True' } |
        Sort-Object -Property ReceivedTimeUtc, SenderAddress, RecipientAddress, Identity -Descending)
    $script:ReportRowCount = $reportRows.Count
    WriteLog -Message ("Blocked quarantine messages normalized: {0}." -f $script:ReportRowCount) -Level 'INFO'

    $stamp = $collectedAtUtc.ToString('yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)
    $maxItemsSuffix = if ($MaxItems -gt 0) { "_MAXITEMS-$MaxItems" } else { '' }
    $script:TimestampedCsvPath = Join-Path -Path $OutputPath -ChildPath "Exchange_EXO_QuarantineMessages_${stamp}${maxItemsSuffix}.csv"
    if ($MaxItems -eq 0) {
        $script:LatestCsvPath = Join-Path -Path $script:LatestCsvFolderPath -ChildPath 'Exchange_EXO_QuarantineMessages.csv'
    }

    $script:CurrentOperation = 'Publish quarantine CSV'
    $exportParameters = @{
        Data            = $(if ($reportRows.Count -gt 0) { $reportRows } else { $null })
        Columns         = $csvColumns
        TimestampedPath = $script:TimestampedCsvPath
    }
    if ($MaxItems -eq 0) {
        $exportParameters['LatestPath'] = $script:LatestCsvPath
    }
    else {
        $exportParameters['NoSharePointUpload'] = $true
        $exportParameters['NoWeeklyHistory'] = $true
    }
    Export-SmartM365Csv @exportParameters | Out-Null

    $script:CurrentOperation = 'Create quarantine Excel workbook'
    $script:WorkbookPath = Join-Path -Path $OutputPath -ChildPath "Exchange_EXO_QuarantineMessages_${stamp}${maxItemsSuffix}.xlsx"
    $workbookRows = if ($reportRows.Count -gt 0) {
        @($reportRows | Add-SmartM365TenantKey)
    }
    else {
        @()
    }
    Export-SmartM365QuarantineWorkbook `
        -Rows $workbookRows `
        -DetailColumns $detailColumns `
        -Path $script:WorkbookPath `
        -StartReceivedDateUtc $startReceivedDateUtc `
        -EndReceivedDateUtc $endReceivedDateUtc `
        -EffectiveLookbackDays $LookbackDays | Out-Null
    WriteLog -Message ("Excel workbook generated: {0}" -f $script:WorkbookPath) -Level 'SUCCESS'

    if ($MaxItems -eq 0 -and $global:EnableSharePointUpload) {
        $script:CurrentOperation = 'Upload quarantine workbook to SharePoint'
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $script:WorkbookPath | Out-Null
    }
    elseif ($MaxItems -gt 0) {
        WriteLog -Message 'MaxItems mode: workbook SharePoint upload skipped.' -Level 'WARNING'
    }

    $mailSuppressed = $NoMail -or -not $script:SendReportMail -or $MaxItems -gt 0
    if ($mailSuppressed) {
        $reason = if ($NoMail) {
            'NoMail switch'
        }
        elseif (-not $script:SendReportMail) {
            'SendReportMail=false'
        }
        else {
            'MaxItems test mode'
        }
        WriteLog -Message ("Quarantine report email skipped: {0}." -f $reason) -Level 'INFO'
    }
    else {
        $script:CurrentOperation = 'Send quarantine Excel report email'
        $mailSubject = '[SmartM365] Quarantaine EXO | {0} message(s) bloque(s) | {1}' -f $script:ReportRowCount, (Get-Date -Format 'yyyy-MM-dd')
        $mailBody = Format-SmartM365QuarantineMailBody `
            -Rows $reportRows `
            -StartReceivedDateUtc $startReceivedDateUtc `
            -EndReceivedDateUtc $endReceivedDateUtc `
            -WorkbookPath $script:WorkbookPath
        Send-SmartM365Mail `
            -Subject $mailSubject `
            -BodyHtml $mailBody `
            -Attachments @($script:WorkbookPath) `
            -AllowAttachments
        WriteLog -Message ("Quarantine report email sent: {0}" -f $mailSubject) -Level 'SUCCESS'
    }

    $script:CurrentOperation = 'Apply retention'
    if ($global:RetentionMaxCSV -gt 0 -and $MaxItems -eq 0) {
        $maxItemsCsvPaths = @(Get-ChildItem -LiteralPath $OutputPath -Filter 'Exchange_EXO_QuarantineMessages_*_MAXITEMS-*.csv' -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
        $maxItemsWorkbookPaths = @(Get-ChildItem -LiteralPath $OutputPath -Filter 'Exchange_EXO_QuarantineMessages_*_MAXITEMS-*.xlsx' -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
        RemoveOldFiles -Path $OutputPath -Filter 'Exchange_EXO_QuarantineMessages_*.csv' -KeepCount $global:RetentionMaxCSV -ExcludeFiles $maxItemsCsvPaths
        RemoveOldFiles -Path $OutputPath -Filter 'Exchange_EXO_QuarantineMessages_*.xlsx' -KeepCount $global:RetentionMaxCSV -ExcludeFiles $maxItemsWorkbookPaths
    }
    elseif ($MaxItems -gt 0) {
        WriteLog -Message 'MaxItems mode: production retention cleanup skipped.' -Level 'INFO'
    }

    $script:CurrentOperation = 'Send Teams completion notification'
    Send-QuarantineTeamsSuccess -Count $script:ReportRowCount -CsvPath $script:TimestampedCsvPath -XlsxPath $script:WorkbookPath

    $resultSummary = 'BlockedMessages={0}; CSV={1}; Excel={2}; MailSent={3}' -f
        $script:ReportRowCount,
        $script:TimestampedCsvPath,
        $script:WorkbookPath,
        (-not $mailSuppressed)
    WriteLog -Message ("Result summary: {0}" -f $resultSummary) -Level 'INFO'
    Write-Host ("EXO quarantine report completed. {0}" -f $resultSummary)
}
catch {
    $failure = $_
    $script:FailureRecord = $failure
    try {
        WriteLog -Message ("EXO quarantine report failed during {0}: {1}" -f $script:CurrentOperation, $failure.Exception.Message) -Level 'ERROR'
    }
    catch {
        Microsoft.PowerShell.Utility\Write-Warning ('Primary failure could not be written through WriteLog: {0}' -f $_.Exception.Message)
    }

    try {
        Send-QuarantineTeamsFailure -ErrorRecord $failure -Operation $script:CurrentOperation
    }
    catch {
        try { WriteLog -Message ("Teams failure notification failed: {0}" -f $_.Exception.Message) -Level 'ERROR' }
        catch { Microsoft.PowerShell.Utility\Write-Warning ('Teams failure notification logging failed: {0}' -f $_.Exception.Message) }
    }

    if ($NoMail -or $MaxItems -gt 0) {
        try {
            $errorMailReason = if ($NoMail) { 'NoMail switch' } else { 'MaxItems test mode' }
            WriteLog -Message ("Error email skipped: {0}." -f $errorMailReason) -Level 'INFO'
        }
        catch {
            Microsoft.PowerShell.Utility\Write-Warning ('Error email suppression logging failed: {0}' -f $_.Exception.Message)
        }
    }
    else {
        try {
            $errorSubject = '[SmartM365] EXO quarantine report - ERROR - {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            $errorBody = NewSimpleEmailBody `
                -Title 'Exchange Online quarantine report failed' `
                -Message ("Operation: {0}`nError: {1}" -f $script:CurrentOperation, $failure.Exception.Message)
            $errorAttachments = @()
            if ($global:LogTextFile -and (Test-Path -LiteralPath $global:LogTextFile)) {
                $errorAttachments = @($global:LogTextFile)
            }
            Send-SmartM365Mail `
                -To $global:ErrorMailTo `
                -Subject $errorSubject `
                -BodyHtml $errorBody `
                -Attachments $errorAttachments `
                -AllowAttachments
        }
        catch {
            try { WriteLog -Message ("Error email failed: {0}" -f $_.Exception.Message) -Level 'ERROR' }
            catch { Microsoft.PowerShell.Utility\Write-Warning ('Error email failure logging failed: {0}' -f $_.Exception.Message) }
        }
    }

    throw
}
finally {
    if ($script:ExchangeConnected) {
        try {
            Disconnect-SmartM365CloudSession -ExchangeOnline:$true -Graph:$false -VerboseDisconnect:$true
        }
        catch { Microsoft.PowerShell.Utility\Write-Warning ('Exchange Online disconnect failed: {0}' -f $_.Exception.Message) }
    }
    if ($script:TranscriptStarted) {
        try {
            Stop-Transcript | Out-Null
            Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile
        }
        catch { Microsoft.PowerShell.Utility\Write-Warning ('Transcript finalization failed: {0}' -f $_.Exception.Message) }
    }

    if ($null -ne $script:FailureRecord) {
        try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $script:FailureRecord -FailureStage $script:CurrentOperation }
        catch { Microsoft.PowerShell.Utility\Write-Warning ('Failure completion banner failed: {0}' -f $_.Exception.Message) }
    }
    else {
        try { Complete-SmartM365ExecutionContext -Status Auto }
        catch { Microsoft.PowerShell.Utility\Write-Warning ('Completion banner failed: {0}' -f $_.Exception.Message) }
    }
}
