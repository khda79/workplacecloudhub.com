# Data adapters for rebuilding report mail from existing exports.
# Version: 1.0.0
Set-StrictMode -Version 2.0

$script:NativeReportModes = @{
    'SmartM365-Teams-Inventory' = 'MailOnly'
    'SmartM365-SPO-Inventory' = 'MailOnly'
    'SmartM365-Licences-Inventory' = 'SendLicenseSummaryEmailOnly'
    'SmartM365-Inventory-Orchestrator' = 'SendExecutionSummary'
}

$script:PublishedReportFiles = @{
    'SmartM365-ActiveDirectory-Inventory' = @('AD_Users_AllDomains.csv','AD_Computers_AllDomains.csv')
    'SmartM365-ActiveDirectory-HealthCheck' = @('AD_HealthCheck.csv')
    'SmartM365-EXO-QuarantineMessages-Report' = @('Exchange_EXO_QuarantineMessages.csv')
    'SmartM365-Exchange-Hybrid-FreeBusy-Inventory' = @('Exchange_Hybrid_FreeBusy_Summary.csv','Exchange_Hybrid_FreeBusy_Findings.csv')
    'SmartM365-Exchange-Local-Mailboxes-Inventory' = @('Exchange_OnPrem_Mailboxes_AllDomains.csv')
    'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory' = @('Exchange_OnPrem_Servers_Inventory_Summary.csv','Exchange_OnPrem_Infrastructure_PerServerSummary.csv','Exchange_OnPrem_MigrationReadiness_Config.csv')
    'SmartM365-WinUpdate_Status_From_Intune' = @('Intune_WindowsUpdate_Status.csv')
    'SmartM365-SharePoint-OnPrem-Content-Inventory' = @('SharePoint_OnPrem_SiteCollections.csv','SharePoint_OnPrem_Webs.csv','SharePoint_OnPrem_CollectionCoverage.csv','SharePoint_OnPrem_SiteAdministrators.csv')
    'SmartM365-WorkplaceEvidence-Prepare' = @()
}
. (Join-Path $PSScriptRoot '../Common/SmartM365.ReportRenderers.ps1')

function Get-SmartM365RebuiltReportMode {
    param([Parameter(Mandatory)][string]$Producer)
    if ($script:NativeReportModes.ContainsKey($Producer)) { return $script:NativeReportModes[$Producer] }
    if ($script:PublishedReportFiles.ContainsKey($Producer)) { return 'Exports' }
    throw 'This script has no approved mail-only data reader.'
}

function Read-RebuiltReportCsv {
    param([string]$Folder,[string]$Name,[string]$TenantKey)
    $path = Join-Path $Folder $Name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing current export: $Name" }
    $item = Get-Item -LiteralPath $path
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked current export is not accepted: $Name" }
    $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $header = Get-Content -LiteralPath $path -TotalCount 1
    $delimiter = if ($header -match '^"?TenantKey"?;') { ';' } else { ',' }
    if ($header -notmatch '^"?TenantKey"?[;,]') { throw "Current export has no tenant column: $Name" }
    $rows = @(Import-Csv -LiteralPath $path -Delimiter $delimiter)
    foreach ($row in $rows) {
        if ([string]$row.TenantKey -cne $TenantKey) { throw "Current export contains a different or empty tenant key: $Name" }
    }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $before) { throw "Current export changed while it was being read: $Name" }
    [pscustomobject]@{Name=$Name;Path=$path;Hash=$before;Rows=$rows;ModifiedUtc=$item.LastWriteTimeUtc}
}

function Read-RebuiltPreparedBatch {
    param([string]$LatestFolder,[string]$TenantKey)
    $folder = Join-Path (Split-Path -Parent $LatestFolder) 'DATA-POWERBI'
    $pointer = Read-SmartM365JsonDocument -Path (Join-Path $folder 'current.json.txt')
    $identity = $pointer.Document
    if ($identity.SchemaVersion -ne 1 -or $identity.TenantKey -cne $TenantKey -or $identity.BatchId -notmatch '^[a-zA-Z0-9][a-zA-Z0-9._-]{0,99}$') {
        throw 'The current prepared batch does not match the tenant.'
    }
    $batch = Join-Path $folder ('batches/' + $identity.BatchId)
    $manifest = Read-SmartM365JsonDocument -Path (Join-Path $batch 'batch.json.txt')
    $m = $manifest.Document
    if ($m.BatchId -cne $identity.BatchId -or $m.TenantKey -cne $TenantKey -or $manifest.SHA256 -ne $identity.ManifestSHA256) {
        throw 'The current prepared batch manifest does not match its pointer.'
    }
    $latest = Read-SmartM365JsonDocument -Path (Join-Path $folder 'current.json.txt')
    if ($latest.SHA256 -ne $pointer.SHA256) { throw 'The prepared batch changed while its metadata was being read.' }
    [pscustomobject]@{BatchId=$m.BatchId; CreatedUtc=$m.CreatedUtc; Files=$m.Files; ManifestHash=$manifest.SHA256}
}


function Get-RebuiltReportFileHash {
    param([string]$Path)
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Invalid report source: $($item.Name)" }
    (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
}

function Read-RebuiltReportInputs {
    param([string]$Producer,[string]$LatestFolder,[string]$TenantKey)
    $directory = Join-Path $LatestFolder '.report-input'
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { return }
    if ((Get-Item -LiteralPath $directory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked report input directory is not accepted.' }
    $path = Join-Path $directory ($Producer + '.clixml.txt')
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
    $hash = Get-RebuiltReportFileHash $path
    $manifest = Import-Clixml -LiteralPath $path
    if ($manifest -isnot [System.Collections.IDictionary] -or $manifest.SchemaVersion -ne 1 -or $manifest.Producer -cne $Producer -or
        $manifest.TenantKey -cne $TenantKey -or $manifest.Inputs -isnot [hashtable] -or $manifest.Files -isnot [hashtable]) {
        throw 'The saved report inputs do not match the selected tenant and collector.'
    }
    foreach ($name in $script:PublishedReportFiles[$Producer]) {
        if (-not $manifest.Files.ContainsKey($name)) { throw "Saved report inputs are missing current export: $name" }
    }
    foreach ($entry in $manifest.Files.GetEnumerator()) {
        if ([string]$entry.Key -notmatch '^[a-zA-Z0-9_-]+\.csv$') { throw 'Invalid export name in report inputs.' }
        $source = Join-Path $LatestFolder $entry.Key
        if (-not (Test-Path -LiteralPath $source -PathType Leaf) -or (Get-RebuiltReportFileHash $source) -cne $entry.Value) {
            throw "Saved report inputs do not match the latest export: $($entry.Key)"
        }
    }
    [pscustomobject]@{Path=$path; Hash=$hash; Manifest=$manifest}
}

function Get-RebuiltLegacyQuarantineInputs {
    param($Source,[string]$LatestFolder,[string]$TenantKey)
    $receiptPath = Join-Path $LatestFolder 'SmartInventory_SmartM365-EXO-QuarantineMessages-Report.current.json.txt'
    $receiptHash = Get-RebuiltReportFileHash $receiptPath
    $receipt = (Read-SmartM365JsonDocument -Path $receiptPath).Document
    $file = @($receipt.Files | Where-Object File -CEQ $Source.Name)
    $days = [int]$receipt.ScopeParameters.LookbackDays
    $time = @($Source.Rows | Select-Object -ExpandProperty CollectedAtUtc -Unique)
    if ($receipt.TenantKey -cne $TenantKey -or $receipt.Status -ne 'Completed' -or $file.Count -ne 1 -or
        $file[0].SHA256 -cne $Source.Hash -or $days -notin (1..30) -or $time.Count -ne 1) {
        throw 'This quarantine export needs the exact collection window from a recorded report input set.'
    }
    if ((Get-RebuiltReportFileHash $receiptPath) -cne $receiptHash) { throw 'The quarantine receipt changed while the report was being built.' }
    $end = ([datetimeoffset]$time[0]).UtcDateTime
    @{
        EndUtc=$end; LookbackDays=$days; WorkbookGeneratedUtc=([datetimeoffset]$receipt.CompletedAtUtc).UtcDateTime
        LocalDate=$end.ToLocalTime().ToString('yyyy-MM-dd')
    }
}

function Convert-RebuiltReportLinks {
    param([string]$BodyHtml,[object[]]$UploadRecords)
    if (-not $UploadRecords.Count) { return $BodyHtml }
    $safeRecords = @(foreach ($record in $UploadRecords) {
        if (-not $record.LocalFilePath) { continue }
        $uri = $null
        if ($record.WebUrl -and (-not [uri]::TryCreate([string]$record.WebUrl,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https')) { continue }
        $record
    })
    $previous = Get-Variable SmartM365SharePointUploadedFiles -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    try {
        $global:SmartM365SharePointUploadedFiles = $safeRecords
        Convert-SmartM365MailBodyLocalPathsToSharePointLinks -BodyHtml $BodyHtml
    } finally {
        $global:SmartM365SharePointUploadedFiles = $previous
    }
}

function New-SmartM365RebuiltReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Producer,[Parameter(Mandatory)][string]$LatestFolder,
          [Parameter(Mandatory)][string]$TenantKey,[Parameter(Mandatory)][string]$ScratchFolder)
    if (-not $script:PublishedReportFiles.ContainsKey($Producer)) { throw 'This script does not have an export report adapter.' }
    if (-not [IO.Path]::IsPathRooted($LatestFolder) -or -not (Test-Path -LiteralPath $LatestFolder -PathType Container) -or
        ((Get-Item -LiteralPath $LatestFolder).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Invalid current export directory.' }
    $sources = @{}
    # The remaining collectors save numeric summaries and run metadata tied to these same current files.
    $parseCsv = $Producer -in @('SmartM365-ActiveDirectory-HealthCheck','SmartM365-EXO-QuarantineMessages-Report',
        'SmartM365-Exchange-Hybrid-FreeBusy-Inventory','SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory')
    foreach ($name in $script:PublishedReportFiles[$Producer]) {
        $path = Join-Path $LatestFolder $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing current export: $name" }
        $sources[$name] = if ($parseCsv) { Read-RebuiltReportCsv -Folder $LatestFolder -Name $name -TenantKey $TenantKey }
            else { [pscustomobject]@{Path=$path; Hash=(Get-RebuiltReportFileHash $path)} }
    }
    $saved = Read-RebuiltReportInputs -Producer $Producer -LatestFolder $LatestFolder -TenantKey $TenantKey
    $inputs = if ($saved) { $saved.Manifest.Inputs }
    $uploadRecords = if ($inputs) {
        $inputs['PathSharePointUploads']
        $inputs.Remove('PathSharePointUploads')
    }
    if (-not $inputs -and $Producer -notin @('SmartM365-EXO-QuarantineMessages-Report','SmartM365-Exchange-Hybrid-FreeBusy-Inventory')) {
        throw "The current export is missing the original report inputs for $Producer. Wait for its next normal collection with this version."
    }

    $report = switch ($Producer) {
        'SmartM365-ActiveDirectory-Inventory' {
            $identity = $null
            $recordedIdentity = $inputs['Identity']
            $inputs.Remove('Identity')
            if ($recordedIdentity) {
                $diagnosticSources = @(Get-SmartM365AdDiagnosticWorkbookSource -Folder $LatestFolder)
                $diagnosticRows = @{}
                foreach ($spec in $diagnosticSources) {
                    if (-not $saved.Manifest.Files.ContainsKey($spec.Name)) { throw "The AD diagnostics are missing the verified export: $($spec.Name)" }
                    $sources[$spec.Name] = Read-RebuiltReportCsv -Folder $LatestFolder -Name $spec.Name -TenantKey $TenantKey
                    $diagnosticRows[$spec.TableName] = $sources[$spec.Name].Rows
                }
                $workbook = Join-Path $ScratchFolder 'AD_Users_IdentityAndMailRoutingIssues.xlsx'
                Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
                Export-SmartM365AdDiagnosticsWorkbook -Sources $diagnosticSources -Path $workbook -ImportExcelAlreadyLoaded | Out-Null
                $previewInputs = @{
                    DuplicateUpnRows=$diagnosticRows.DuplicateUPN; DuplicateSmtpRows=$diagnosticRows.DuplicateSMTP
                    DuplicateRemoteRows=$diagnosticRows.DuplicateRemoteRouting; RoutingIssueRows=$diagnosticRows.RemoteRoutingIssues
                }
                $userDomains = @($recordedIdentity.UserDomains | ForEach-Object { [pscustomobject]@{DomainName=$_} })
                $identity = [pscustomobject]@{
                    SummaryHtml=(New-SmartM365AdIdentitySummaryHtml -AllUsers $userDomains @previewInputs)
                    Sections=@(New-SmartM365AdIdentityMailSections @previewInputs -Limit $recordedIdentity.PreviewLimit | Where-Object { $_ })
                    SharePointUploads=$recordedIdentity.SharePointUploads; PathRows=$recordedIdentity.PathRows; WorkbookPath=$workbook
                }
            }
            $adReport = New-SmartM365AdInventoryMailReport @inputs -IdentityContext $identity
            if ($identity) {
                $reference = $recordedIdentity.WorkbookFile
                if (-not $inputs.ShowMailLinks) {
                    $reference = [pscustomobject]@{Name=$reference.Name;Path=$reference.Path;SizeKB=$reference.SizeKB;WebUrl='';SharePointPath=''}
                }
                $adReport.BodyHtml = Add-SmartM365MailFilesSection -BodyHtml $adReport.BodyHtml -ReferenceFiles @($reference)
                $adReport.SuppressAttachmentLinks = $true
            }
            $adReport.BodyHtml = Convert-RebuiltReportLinks -BodyHtml $adReport.BodyHtml -UploadRecords @($uploadRecords | Where-Object { $_ })
            $adReport
        }
        'SmartM365-ActiveDirectory-HealthCheck' {
            $source = $sources['AD_HealthCheck.csv']
            $runDate = $inputs.RunDateUtc
            $inputs.Remove('RunDateUtc')
            $body = ConvertTo-ReportHtml @inputs -r $source.Rows
            $body = Convert-RebuiltReportLinks -BodyHtml $body -UploadRecords @($uploadRecords | Where-Object { $_ })
            [pscustomobject]@{
                Subject="SMART365 - [$($inputs.Status.ToUpperInvariant())] Active Directory Health Check - $($inputs.ForestInfo.Name) - $runDate"
                BodyHtml=$body; VerboseLog=$true
            }
        }
        'SmartM365-EXO-QuarantineMessages-Report' {
            $source = $sources['Exchange_EXO_QuarantineMessages.csv']
            if (-not $inputs) { $inputs = Get-RebuiltLegacyQuarantineInputs -Source $source -LatestFolder $LatestFolder -TenantKey $TenantKey }
            $endUtc = ([datetimeoffset]$inputs.EndUtc).UtcDateTime
            $lookback = [int]$inputs.LookbackDays
            if ($lookback -notin (1..30)) { throw 'The recorded quarantine lookback is invalid.' }
            $startUtc = $endUtc.AddDays(-$lookback)
            $stamp = $endUtc.ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
            $workbook = Join-Path $ScratchFolder "Exchange_EXO_QuarantineMessages_$stamp.xlsx"
            $rows = @($source.Rows)
            foreach ($row in $rows) {
                if ($row.AgeDays -ne '') { $row.AgeDays = [double]::Parse($row.AgeDays,[Globalization.CultureInfo]::InvariantCulture) }
                # Old exports may predate spreadsheet text validation. Keep incoming subjects and values literal.
                foreach ($column in @('SenderAddress','RecipientAddress','Subject','QuarantineType','PolicyName','PolicyType','Direction','ReleaseStatus','MessageId','Identity')) {
                    if ($row.$column -match '^[\s]*[=+@-]|^[\t\r\n]') { $row.$column = "'" + $row.$column }
                }
            }
            Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
            Export-SmartM365QuarantineWorkbook -Rows $rows -DetailColumns @(Get-SmartM365QuarantineReportColumns) -Path $workbook -StartReceivedDateUtc $startUtc -EndReceivedDateUtc $endUtc -EffectiveLookbackDays $lookback -GeneratedUtc ([datetimeoffset]$inputs.WorkbookGeneratedUtc).UtcDateTime | Out-Null
            $body = Format-SmartM365QuarantineMailBody -Rows $rows -StartReceivedDateUtc $startUtc -EndReceivedDateUtc $endUtc -WorkbookPath $workbook
            if ($inputs['WorkbookFile']) { $body = Add-SmartM365MailFilesSection -BodyHtml $body -ReferenceFiles @($inputs.WorkbookFile) }
            [pscustomobject]@{Subject=(Get-SmartM365QuarantineReportSubject -Count $rows.Count -LocalDate $inputs.LocalDate); BodyHtml=$body
                Attachments=@($workbook); AllowAttachments=$true; SuppressAttachmentLinks=[bool]$inputs['WorkbookFile']}
        }
        'SmartM365-Exchange-Hybrid-FreeBusy-Inventory' {
            $summary = $sources['Exchange_Hybrid_FreeBusy_Summary.csv']
            $findings = $sources['Exchange_Hybrid_FreeBusy_Findings.csv']
            if (-not $summary.Rows.Count) { throw 'Free/Busy summary is empty.' }
            $status = @($summary.Rows | Select-Object -ExpandProperty RunStatus -Unique)
            $stamp = @($summary.Rows | Select-Object -ExpandProperty GenerationTime -Unique)
            if ($status.Count -ne 1 -or $stamp.Count -ne 1 -or -not $status[0] -or -not $stamp[0]) { throw 'Free/Busy exports do not identify a single run.' }
            foreach ($row in $findings.Rows) {
                if ($row.GenerationTime -cne $stamp[0]) { throw 'Free/Busy findings and summary are from different runs.' }
                $row.RowNumber = [int]$row.RowNumber
            }
            if (-not $inputs) {
                $summaryRow = $summary.Rows | Where-Object SideStatus -NE 'NotApplicable' | Select-Object -First 1
                if (-not $summaryRow) { throw 'Free/Busy summary has no applicable run.' }
                $inputs = @{ RunId=([datetimeoffset]$stamp[0]).ToString('yyyyMMdd-HHmmss'); RunStatus=$status[0]; RunMessage=$summaryRow.RunMessage; CheckerVersion=$summaryRow.FreeBusyCheckerVersion }
            }
            $body = ConvertTo-FreeBusyReportHtml -Summary $summary.Rows -Findings $findings.Rows -RunStatus $inputs.RunStatus -RunMessage $inputs.RunMessage -CheckerVersion $inputs.CheckerVersion -ReportRunId $inputs.RunId
            $body = Add-SmartM365FreeBusyTemplateMarker $body
            if ($inputs['SummaryFile']) {
                $body = Add-SmartM365MailFilesSection -BodyHtml $body -ReferenceFiles @($inputs.SummaryFile)
                [pscustomobject]@{Subject="SmartM365 Exchange Hybrid Free/Busy check - $($inputs.RunId)"; BodyHtml=$body}
            } else {
                [pscustomobject]@{Subject="SmartM365 Exchange Hybrid Free/Busy check - $($inputs.RunId)"; BodyHtml=$body; Attachments=@($summary.Path)}
            }
        }
        'SmartM365-Exchange-Local-Mailboxes-Inventory' {
            Import-Module (Join-Path $PSScriptRoot '../ExchangeInventory/OnPremises/Mailboxes/SmartM365.ExchangeProxyMail.psm1') -ErrorAction Stop
            $diff = if ($inputs.CurrentSnapshot) {
                New-SmartM365ExchangeMailboxDailyDiffSection -Current $inputs.CurrentSnapshot -History @($inputs.History)
            }
            $mailInputs = $inputs.Mail
            $body = New-SmartM365ExchangeLocalMailboxReportEmailBody @mailInputs -DailyDiffSection $diff
            [pscustomobject]@{Subject='SmartM365'; BodyHtml=$body}
        }
        'SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory' {
            Set-Alias -Name New-CoreSmartM365EmailBody -Value New-SmartM365EmailBody -Scope Local
            Set-Alias -Name New-CoreSmartM365EmailKpiGridHtml -Value New-SmartM365EmailKpiGridHtml -Scope Local
            Set-Alias -Name New-CoreSmartM365EmailTableHtml -Value New-SmartM365EmailTableHtml -Scope Local
            foreach ($name in @('Exchange_OnPrem_HealthChecker_Summary.csv','Exchange_OnPrem_HealthChecker_Findings.csv')) {
                if ($saved.Manifest.Files.ContainsKey($name)) { $sources[$name] = Read-RebuiltReportCsv -Folder $LatestFolder -Name $name -TenantKey $TenantKey }
            }
            $subject = $inputs.Subject
            $inputs.Remove('Subject')
            $summary = $sources['Exchange_OnPrem_Servers_Inventory_Summary.csv'].Rows
            if ($summary.Count -ne 1) { throw 'The current Exchange infrastructure summary must contain one row.' }
            $body = New-ServersAndStorageMailBodyHtml @inputs -Summary $summary[0] -PerServerSummary $sources['Exchange_OnPrem_Infrastructure_PerServerSummary.csv'].Rows `
                -ReadinessInventory $sources['Exchange_OnPrem_MigrationReadiness_Config.csv'].Rows `
                -HealthCheckerSummary $(if ($sources.ContainsKey('Exchange_OnPrem_HealthChecker_Summary.csv')) { $sources['Exchange_OnPrem_HealthChecker_Summary.csv'].Rows }) `
                -HealthCheckerFindings $(if ($sources.ContainsKey('Exchange_OnPrem_HealthChecker_Findings.csv')) { $sources['Exchange_OnPrem_HealthChecker_Findings.csv'].Rows })
            [pscustomobject]@{Subject=$subject; BodyHtml=$body}
        }
        'SmartM365-WinUpdate_Status_From_Intune' {
            New-SmartM365WindowsUpdateMailReport @inputs
        }
        'SmartM365-SharePoint-OnPrem-Content-Inventory' {
            $script:BuildEditionMapPath = Join-Path $PSScriptRoot '../SharePointInventory/OnPremises/SharePoint-OnPrem-BuildEditions.psd1'
            [pscustomobject]@{Subject='SharePoint on-prem inventory'; BodyHtml=(New-ContentMailHtml @inputs); SendMailMode='Graph'}
        }
        'SmartM365-WorkplaceEvidence-Prepare' {
            $batch = Read-RebuiltPreparedBatch -LatestFolder $LatestFolder -TenantKey $TenantKey
            if ($batch.BatchId -cne $saved.Manifest.BatchId -or $batch.ManifestHash -cne $saved.Manifest.BatchManifestHash) {
                throw 'The latest prepared batch does not match the original completion report inputs.'
            }
            [pscustomobject]@{Subject="$Producer completed - batch $($batch.BatchId)"; BodyHtml=(New-SmartM365PreparedCompletionMailBody @inputs)}
        }
    }
    foreach ($source in $sources.Values) {
        if ((Get-RebuiltReportFileHash $source.Path) -cne $source.Hash) { throw "Current export changed before the mail was ready: $($source.Path)" }
    }
    if ($saved) {
        foreach ($entry in $saved.Manifest.Files.GetEnumerator()) {
            if (-not $sources.ContainsKey($entry.Key) -and (Get-RebuiltReportFileHash (Join-Path $LatestFolder $entry.Key)) -cne $entry.Value) {
                throw "Current export changed before the mail was ready: $($entry.Key)"
            }
        }
        if ((Get-RebuiltReportFileHash $saved.Path) -cne $saved.Hash) { throw 'Saved report inputs changed before the mail was ready.' }
    }
    return $report
}

Export-ModuleMember -Function Get-SmartM365RebuiltReportMode, New-SmartM365RebuiltReport

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOwlxH4D3uKhum
# fCzgEiV79u6p3Lpmb6CqXlWaJLpwoqCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgrA6NP3PKUO7P
# KOfyi4Unf2pW4BbNVcST/9smlqz99E8wDQYJKoZIhvcNAQEBBQAEggGActEgRZhD
# PqU4jAQmPv826y9Fs1VZEIbDQNJRsZNM9G7Y6J/L0cozD1pxgd33QVW08sd3ecVN
# a7gmSzIYTcSrPW2VWwGS1mr58zybmt0FOrRFAPz2UpTDdQg4hEJ+OY4L5IjA9yIi
# ZRGSeC872bdBZWWIHlZhMs/1i0RpQ3S+bDpPLGVdXegpH9OwULH/Ys4FQAWDyIr3
# tPkpvwKFlBxGy6o00lAmWPo48D5MSdw9HqJ7nikayLm1wmV/FQmyGYWWoAfIQgTQ
# zlCBNp31Og7Ao3ZvUGLGPzWVJzbK4U5/ufOPfw0ZjIhc52nkupsjJjSsoDL9cdpp
# g5AuF/clay2RWiQEpng5U/SFC6tM8Xhcce7IsxwXaWLXqoeSwLQlluGI4v6JnAcF
# HJRc2avSHSVQdMqCsMq6jbMxzfULqjAqO4t2Hez2Sj18GPaRmEPjsKW2e2MxLW9P
# rP0AXmuWBSpZjdZ+OqzRofx92s2zIXB4qAVxFPlG+lR+jf4jEkLi3xC+oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxODI5WjA/BgkqhkiG9w0BCQQxMgQwzqNMAF3OYYM40TFDEeX2q9s0
# dN4jAcuiEt2AEneUokT9+xO8omyiPdJbxOpyJ0EGMA0GCSqGSIb3DQEBAQUABIIC
# AKK7RAz3A/zpHxBCnyrAsk8ArwHdoUYML50D/A7Wsnafd+Gy2Po0/WYvXKafhdjn
# FXTWUcW4FuSgr4KVY/VDnlQ/lZB+8mmgw1XfF+r9jf1vhDWvJ/sXD4w7HT4ebVAj
# /ZFG5lrowmPkXsmNh+DUMLWZihkli+PNJJ8kkmFp6aa27FaAmW4GNwiZtx7TctuX
# 7QlX+n+S1l8dx4MlLLLq94JGRAwivLw9948MYGIDDxiQl60V0l5GtuHOCRVAdewZ
# OvYEqtWGtvBAKBuRZS+Qh55dMUCE0/ewGXl1kQGg/qq6gExVzST+CUVUWYfze0cU
# X/o00xcur3Y5ZJLW5Y0zjzH3VYsrFyeOylZiiH6n1vv4o4BZCBd3dOll4xOlK3My
# ZfNDcimxfIcjJCPi9x9JUGEUuZytCZXDUliXK8XH9Aok0/5MCks59gJUrkwTCeCh
# IO8cfwazFG8s+5aHezcSjhJzDrvdvSC1dE/SkIveulBzolVQfGkculxX2xcLYJLc
# L+a+2LjddIwKMcd78urS16EdrBAol6SmTR7W/vxw94+Voxj09EtoqCUzNuydHBP2
# bYCxc5i0UTGOMpGSeeUL/GKq6g2NeBPmfurlaZ2jf18JSwQFINf1PTQjTb29pd84
# xIc4eVX9J7y63LSOU/8T8yml+74a73qTr7+2I+bxJ3Nt
# SIG # End signature block
