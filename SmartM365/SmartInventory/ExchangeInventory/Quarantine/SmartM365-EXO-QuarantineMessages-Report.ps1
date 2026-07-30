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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCByPA5+LBzBIgk
# QCsX7c4T5tCRdZiHrzkq9x9uKO4dNaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIApNCU2Y1ugiMuALO9OpdJH/tLdc87A8+5nwqS5Eq0FgMA0GCSqG
# SIb3DQEBAQUABIIBgCpH67yX6ebDPwtNlc+vt7KiO+MyIImcTw6j+oDTwZ6ewxB+
# YSGRbV5IdITaRyXgOQH1FQCaC97Zkjg1bPT6wkKjaRnBHht+UI+TOiRqqU+U/a31
# dW7B1Q4frn9yRhR3fxeu8yLbYje4PqtpowR8Xfliby3k873sV5Cs3rNTiCM96H8M
# onfa88bfAnoNZY6W0uz0Ey8EthvQsU1rpgfEM67lIsRKPpA+ITFl+Zka4ur0ZMzq
# EN4yQ2T2ydIQvWHjdJWJU0QLlm9Us5jgWoUFn8bJmDUN0Wm/6XQ4acoRfKVUekl6
# 4xjwWUhk66nBncNmMrczglEDI/YKyXfu/F/2vINSAyeDOcmkXwdoLDdA9AVknFP7
# lleyflVbXs8yK6ivdOGhjzrE3JsNfDFeqk0kQ3YSURgY/UXCULw8sgQ7H8hvJkPw
# X946cJAGYgu7vpNCc8+KlPMhI3JqjRkBITLfYQiJBt2eiczOC8ORiOmTJ2Od46Dn
# pxiasAkLlcJgxdsFrKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MzAxMjQ3
# MDRaMC8GCSqGSIb3DQEJBDEiBCD86q7qA0OYqt+F8+NOZ2quu7R6sZcPBNTetUmz
# ahO7gjANBgkqhkiG9w0BAQEFAASCAgA40sBojgEVgHbC2WNwyDBRfT3E3fkYkSj+
# Y/K1vBx89AkwHB7H7ewQQOoQPaSAqspqEV5ZhBu4qrqKEtau5aPLGZ14G7xmIJOc
# XLpR2egEav3ep1fbFvm/YymPENSe3cJm68aiSmoBPQd6hl6p5OIbQVLxAUAPdO0Q
# Xlqiw0Fe+HSexH8TlKZFKeNJiVjkh3AleFdcgPBxb+oUY91gtGfeOgUb/MRiyp8Z
# keMtzdBaJiXSEcD67Q8uTE8+nCX3r65jnaIrmGRwqyCzhftMfztGa1+OV1bpgRi0
# 3SQ4jOB3sGpq3F+KUUEbN6Tl7lwbk4zHxXq7DajtoAw5LVdn2Oefb/0a8OXGmTfl
# 4ou+ksRAc43czsS1E0MT/J6vtHvUmHkpl+al+FZBAwoTLi5nIn70yFH1+Kq33wx7
# upOihkJBxM68DP3usk4JmxWki0tpDcUaG1qoerloDAlmsQnoc73NFwqltVGGXA6P
# EGriG298kzxhs5KGkwAOgM2TuzpkEhDJROUA9pEduSUwyG8tnA+93+0Bs94qjfk6
# XlIB3LQiTKlFo/9qF2uaLxzj9Fpg4NBzEmwR/dy8wnDdI3ntht0zOKevNFVOCIe6
# 8xJUrAoyrKGxenUcJurDJBN3hzK18FPRHm20yPl/gKEzyCD8NmNjOpJ2odaMQPgm
# sq2/bFhNDA==
# SIG # End signature block
