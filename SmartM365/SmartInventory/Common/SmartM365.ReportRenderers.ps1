# Report builders shared by collection and mail-only. Dot-sourcing defines functions only.

# SmartM365-ActiveDirectory-Inventory

function ConvertTo-SmartM365AdDomainSetText {
        [CmdletBinding()]
        param([AllowNull()][object[]]$DomainNames)

        return (@($DomainNames | ForEach-Object { ([string]$_).Trim().TrimEnd('.').ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique) -join ';')
    }

function Get-SmartM365AdSummaryCoverageComparison {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)]$Current, [AllowNull()]$Previous)

        if ($null -eq $Previous) {
            return [pscustomobject]@{ IsComparable = $false; Reason = 'No previous scan' }
        }
        if ((ConvertTo-SmartM365AdSummaryInt64 $Current.DomainCount) -ne (ConvertTo-SmartM365AdSummaryInt64 $Previous.DomainCount)) {
            return [pscustomobject]@{ IsComparable = $false; Reason = 'Domain count changed' }
        }
        $currentProperty = $Current.PSObject.Properties['CollectedDomains']
        $previousProperty = $Previous.PSObject.Properties['CollectedDomains']
        $currentSet = if ($currentProperty) { ConvertTo-SmartM365AdDomainSetText -DomainNames @(([string]$currentProperty.Value) -split ';') } else { '' }
        $previousSet = if ($previousProperty) { ConvertTo-SmartM365AdDomainSetText -DomainNames @(([string]$previousProperty.Value) -split ';') } else { '' }
        if ([string]::IsNullOrWhiteSpace($currentSet) -or [string]::IsNullOrWhiteSpace($previousSet)) {
            return [pscustomobject]@{ IsComparable = $false; Reason = 'Domain list unavailable' }
        }
        if (@($currentSet -split ';').Count -ne (ConvertTo-SmartM365AdSummaryInt64 $Current.DomainCount) -or
            @($previousSet -split ';').Count -ne (ConvertTo-SmartM365AdSummaryInt64 $Previous.DomainCount)) {
            return [pscustomobject]@{ IsComparable = $false; Reason = 'Domain list and count disagree' }
        }
        if ($currentSet -cne $previousSet) {
            return [pscustomobject]@{ IsComparable = $false; Reason = 'Domain set changed' }
        }
        return [pscustomobject]@{ IsComparable = $true; Reason = 'Same domains' }
    }

function Get-SmartM365AdSummaryMetricComparison {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Current,
            [AllowNull()]$Previous,
            [Parameter(Mandatory = $true)][string]$Property,
            [string]$LogicalCountBaselineAt = ''
        )

        if ($null -eq $Previous) { return [pscustomobject]@{ IsComparable = $false; Delta = 'n/a'; DeltaValue = 0; Reason = 'No previous scan' } }
        $currentProperty = $Current.PSObject.Properties[$Property]
        $previousProperty = $Previous.PSObject.Properties[$Property]
        if (-not $currentProperty -or -not $previousProperty -or
            [string]::IsNullOrWhiteSpace([string]$currentProperty.Value) -or
            [string]::IsNullOrWhiteSpace([string]$previousProperty.Value)) {
            return [pscustomobject]@{ IsComparable = $false; Delta = 'n/a'; DeltaValue = 0; Reason = 'Metric unavailable' }
        }
        if ($Property -ne 'DomainCount') {
            $coverage = Get-SmartM365AdSummaryCoverageComparison -Current $Current -Previous $Previous
            if (-not $coverage.IsComparable) {
                return [pscustomobject]@{ IsComparable = $false; Delta = 'n/a'; DeltaValue = 0; Reason = $coverage.Reason }
            }
        }
        if ($Property -in @('TotalComputers','TotalGroups','TotalContacts')) {
            $snapshotTime = [datetimeoffset]::MinValue
            $baselineTime = [datetimeoffset]::MinValue
            if ([string]::IsNullOrWhiteSpace($LogicalCountBaselineAt) -or
                -not [datetimeoffset]::TryParse([string]$Previous.GeneratedAt, [ref]$snapshotTime) -or
                -not [datetimeoffset]::TryParse($LogicalCountBaselineAt, [ref]$baselineTime) -or
                $snapshotTime -lt $baselineTime) {
                return [pscustomobject]@{ IsComparable = $false; Delta = 'n/a'; DeltaValue = 0; Reason = 'Historical count incompatible' }
            }
        }
        $deltaValue = (ConvertTo-SmartM365AdSummaryInt64 $currentProperty.Value) - (ConvertTo-SmartM365AdSummaryInt64 $previousProperty.Value)
        return [pscustomobject]@{
            IsComparable = $true
            Delta = Format-SmartM365AdSummaryDelta -Current $currentProperty.Value -Previous $previousProperty.Value -HasPrevious $true
            DeltaValue = $deltaValue
            Reason = ''
        }
    }

function Get-SmartM365AdSummaryDeltaColor {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$Property, [Parameter(Mandatory = $true)][int64]$DeltaValue)

        if ($DeltaValue -eq 0) { return '#475569' }
        if ($Property -in @('DistinctDuplicateUPNs','AffectedDuplicateUPNAccounts','DistinctDuplicateSMTPAddresses','AffectedDuplicateSMTPEntries')) {
            if ($DeltaValue -gt 0) { return '#b91c1c' }
            return '#166534'
        }
        if ($Property -eq 'DomainCount' -and $DeltaValue -lt 0) { return '#b91c1c' }
        return '#475569'
    }

function New-SmartM365AdDailySummaryKpiHtml {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Snapshot,
            [AllowNull()]$Previous,
            [string]$LogicalCountBaselineAt = ''
        )

        $upnHasIssues = (ConvertTo-SmartM365AdSummaryInt64 $Snapshot.AffectedDuplicateUPNAccounts) -gt 0
        $smtpHasIssues = (ConvertTo-SmartM365AdSummaryInt64 $Snapshot.AffectedDuplicateSMTPEntries) -gt 0
        $cards = @(
            [pscustomobject]@{ Label = 'USERS'; Property = 'TotalUsers'; Detail = 'accounts in snapshot'; Background = '#eff6ff'; Border = '#bfdbfe'; Accent = '#1d4ed8' }
            [pscustomobject]@{ Label = 'COMPUTERS'; Property = 'TotalComputers'; Detail = 'computer objects'; Background = '#f0fdfa'; Border = '#99f6e4'; Accent = '#0f766e' }
            [pscustomobject]@{ Label = 'GROUPS'; Property = 'TotalGroups'; Detail = 'groups in snapshot'; Background = '#f5f3ff'; Border = '#ddd6fe'; Accent = '#7c3aed' }
            [pscustomobject]@{ Label = 'UPN DUPLICATES'; Property = 'AffectedDuplicateUPNAccounts'; Detail = 'affected accounts'; Background = $(if ($upnHasIssues) { '#fff7ed' } else { '#ecfdf5' }); Border = $(if ($upnHasIssues) { '#fed7aa' } else { '#bbf7d0' }); Accent = $(if ($upnHasIssues) { '#c2410c' } else { '#166534' }) }
            [pscustomobject]@{ Label = 'SMTP DUPLICATES'; Property = 'AffectedDuplicateSMTPEntries'; Detail = 'affected entries'; Background = $(if ($smtpHasIssues) { '#fef2f2' } else { '#ecfdf5' }); Border = $(if ($smtpHasIssues) { '#fecaca' } else { '#bbf7d0' }); Accent = $(if ($smtpHasIssues) { '#b91c1c' } else { '#166534' }) }
        )
        $coverage = Get-SmartM365AdSummaryCoverageComparison -Current $Snapshot -Previous $Previous
        $currentDomainCount = Format-SmartM365AdSummaryNumber $Snapshot.DomainCount
        $previousDomainText = if ($Previous) {
            'Previous ({0}): {1} ({2})' -f $Previous.SnapshotDate,
                (Format-SmartM365AdSummaryNumber $Previous.DomainCount),
                (Format-SmartM365AdSummaryDelta -Current $Snapshot.DomainCount -Previous $Previous.DomainCount -HasPrevious $true)
        } else { 'Previous: n/a' }
        $coverageText = 'Domains analyzed: {0} | {1} | {2}' -f $currentDomainCount, $previousDomainText, $coverage.Reason
        $coverageBackground = if ($coverage.IsComparable) { '#ecfdf5' } else { '#fff7ed' }
        $coverageBorder = if ($coverage.IsComparable) { '#bbf7d0' } else { '#fed7aa' }
        $coverageHtml = '<div style="background:{0};border:1px solid {1};border-radius:6px;padding:9px 12px;margin:0 0 12px 0;font-size:12px;color:#334155;">{2}</div>' -f $coverageBackground, $coverageBorder, (ConvertTo-SmartM365EmailHtmlText $coverageText)
        $cells = foreach ($card in $cards) {
            $comparison = Get-SmartM365AdSummaryMetricComparison -Current $Snapshot -Previous $Previous -Property $card.Property -LogicalCountBaselineAt $LogicalCountBaselineAt
            $label = ConvertTo-SmartM365EmailHtmlText $card.Label
            $value = ConvertTo-SmartM365EmailHtmlText (Format-SmartM365AdSummaryNumber $Snapshot.($card.Property))
            $detail = ConvertTo-SmartM365EmailHtmlText $card.Detail
            $deltaColor = if ($comparison.IsComparable) { Get-SmartM365AdSummaryDeltaColor -Property $card.Property -DeltaValue $comparison.DeltaValue } else { '#64748b' }
            $previousLabel = if ($Previous) { 'vs ' + [string]$Previous.SnapshotDate } else { 'vs previous' }
            $deltaText = ConvertTo-SmartM365EmailHtmlText ('{0}: {1}' -f $previousLabel, $comparison.Delta)
            @"
<td width="20%" valign="top" style="padding:0 4px;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;">
    <tr><td style="height:108px;background:$($card.Background);border:1px solid $($card.Border);border-radius:6px;padding:10px;">
      <div style="min-height:25px;font-size:10px;line-height:12px;font-weight:700;color:#64748b;">$label</div>
      <div style="font-size:25px;line-height:30px;font-weight:700;color:$($card.Accent);">$value</div>
      <div style="font-size:10px;line-height:14px;color:#475569;">$detail</div>
      <div style="margin-top:6px;font-size:10px;line-height:13px;font-weight:700;color:$deltaColor;">$deltaText</div>
    </td></tr>
  </table>
</td>
"@
        }
        return @"
$coverageHtml
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;table-layout:fixed;margin:0 0 16px 0;">
  <tr>$($cells -join [Environment]::NewLine)</tr>
</table>
"@
    }

function New-SmartM365AdDomainSummaryHtml {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)]$Snapshot, $DomainCounts=$script:SmartM365AdDailyDomainCounts, $DomainNames=$script:SmartM365AdDailyDomainNames)

        $maps = $DomainCounts
        $domains = @(@($maps.Values | ForEach-Object { $_.Keys }) + @($DomainNames) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
        $rows = foreach ($domain in $domains) {
            $cells = foreach ($key in @('Users','Computers','Groups','OUs','Contacts','DuplicateUPN','DuplicateSMTP')) {
                $value = if ($maps[$key].ContainsKey($domain)) { $maps[$key][$domain] } else { 0 }
                '<td align="right" style="border-bottom:1px solid #eef2f7;padding:7px 5px;font-size:11px;">{0}</td>' -f (Format-SmartM365AdSummaryNumber $value)
            }
            '<tr><td style="border-bottom:1px solid #eef2f7;padding:7px 5px;font-size:11px;">{0}</td>{1}</tr>' -f (ConvertTo-SmartM365EmailHtmlText $domain), ($cells -join '')
        }
        $totals = foreach ($value in @($Snapshot.TotalUsers,$Snapshot.TotalComputers,$Snapshot.TotalGroups,$Snapshot.TotalOUs,$Snapshot.TotalContacts,$Snapshot.AffectedDuplicateUPNAccounts,$Snapshot.AffectedDuplicateSMTPEntries)) {
            '<td align="right" style="padding:8px 5px;font-size:11px;font-weight:700;background:#dbeafe;">{0}</td>' -f (Format-SmartM365AdSummaryNumber $value)
        }
        return @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
<tr style="background:#f1f5f9;font-size:10px;"><th align="left">Domain</th><th>Users</th><th>Computers</th><th>Groups</th><th>OUs</th><th>Contacts</th><th>UPN dup.</th><th>SMTP dup.</th></tr>
$($rows -join "`n")
<tr><td style="padding:8px 5px;font-size:11px;font-weight:700;background:#dbeafe;">TOTAL ($($Snapshot.DomainCount) domains)</td>$($totals -join '')</tr>
</table>
"@
    }

function ConvertTo-SmartM365AdSummaryInt64 {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0 }
        [int64]$parsed = 0
        if ([int64]::TryParse(([string]$Value).Trim(), [ref]$parsed)) { return $parsed }
        return 0
    }

function Format-SmartM365AdSummaryNumber {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        $number = ConvertTo-SmartM365AdSummaryInt64 -Value $Value
        return $number.ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture)
    }

function Format-SmartM365AdSummaryDelta {
        [CmdletBinding()]
        param(
            [AllowNull()]$Current,
            [AllowNull()]$Previous,
            [bool]$HasPrevious
        )

        if (-not $HasPrevious) { return 'n/a' }

        $delta = (ConvertTo-SmartM365AdSummaryInt64 -Value $Current) - (ConvertTo-SmartM365AdSummaryInt64 -Value $Previous)
        if ($delta -gt 0) { return ('+{0}' -f (Format-SmartM365AdSummaryNumber -Value $delta)) }
        if ($delta -lt 0) { return ('-{0}' -f (Format-SmartM365AdSummaryNumber -Value ([Math]::Abs($delta)))) }
        return '0'
    }

function New-SmartM365AdDailySummaryDiffSection {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Current,
            [AllowNull()]$Previous,
            [AllowNull()]$Week,
            [AllowNull()]$Month,
            [string]$LogicalCountBaselineAt = ''
        )

        $metrics = @(
            [pscustomobject]@{ Label = 'Domains'; Property = 'DomainCount' }
            [pscustomobject]@{ Label = 'Users'; Property = 'TotalUsers' }
            [pscustomobject]@{ Label = 'Computers'; Property = 'TotalComputers' }
            [pscustomobject]@{ Label = 'Groups'; Property = 'TotalGroups' }
            [pscustomobject]@{ Label = 'OUs'; Property = 'TotalOUs' }
            [pscustomobject]@{ Label = 'Contacts'; Property = 'TotalContacts' }
            [pscustomobject]@{ Label = 'Distinct duplicate UPNs'; Property = 'DistinctDuplicateUPNs' }
            [pscustomobject]@{ Label = 'Affected duplicate UPN accounts'; Property = 'AffectedDuplicateUPNAccounts' }
            [pscustomobject]@{ Label = 'Distinct duplicate SMTP addresses'; Property = 'DistinctDuplicateSMTPAddresses' }
            [pscustomobject]@{ Label = 'Affected duplicate SMTP entries'; Property = 'AffectedDuplicateSMTPEntries' }
        )

        $comparisons = @(
            [pscustomobject]@{ Label = 'Previous'; Snapshot = $Previous }
            [pscustomobject]@{ Label = 'J-7'; Snapshot = $Week }
            [pscustomobject]@{ Label = 'J-30'; Snapshot = $Month }
        )
        $headers = foreach ($comparison in $comparisons) {
            $dateText = if ($comparison.Snapshot) { [string]$comparison.Snapshot.SnapshotDate } else { 'n/a' }
            '<th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">{0} ({1})</th>' -f $comparison.Label, (ConvertTo-SmartM365EmailHtmlText $dateText)
        }
        $rows = foreach ($metric in $metrics) {
            $currentProperty = $Current.PSObject.Properties[$metric.Property]
            $currentValue = if ($null -ne $currentProperty) { $currentProperty.Value } else { 0 }
            $labelHtml = ConvertTo-SmartM365EmailHtmlText $metric.Label
            $currentHtml = ConvertTo-SmartM365EmailHtmlText (Format-SmartM365AdSummaryNumber -Value $currentValue)
            $comparisonCells = foreach ($comparison in $comparisons) {
                $snapshot = $comparison.Snapshot
                $metricComparison = Get-SmartM365AdSummaryMetricComparison -Current $Current -Previous $snapshot -Property $metric.Property -LogicalCountBaselineAt $LogicalCountBaselineAt
                $cell = 'n/a'
                if ($snapshot) {
                    $property = $snapshot.PSObject.Properties[$metric.Property]
                    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $color = if ($metricComparison.IsComparable) { Get-SmartM365AdSummaryDeltaColor -Property $metric.Property -DeltaValue $metricComparison.DeltaValue } else { '#64748b' }
                        $cell = '{0}<br /><strong style="color:{1};">{2}</strong>' -f
                            (ConvertTo-SmartM365EmailHtmlText (Format-SmartM365AdSummaryNumber $property.Value)),
                            $color,
                            (ConvertTo-SmartM365EmailHtmlText $metricComparison.Delta)
                    }
                }
                '<td align="right" style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;color:#334155;">{0}</td>' -f $cell
            }
            '<tr><td style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;color:#334155;">{0}</td><td align="right" style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;font-weight:700;">{1}</td>{2}</tr>' -f $labelHtml, $currentHtml, ($comparisonCells -join '')
        }
        $html = @"
<div style="font-size:12px;color:#64748b;margin-bottom:8px;">Each delta compares Current with the dated scan. Inventory and duplicate deltas require the same recorded domain set; n/a means no scan, unknown or changed coverage, or an incompatible historical count. The historical value remains visible where available.</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">Metric</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">Current ($($Current.SnapshotDate))</th>
    $($headers -join "`n")
  </tr>
  $($rows -join "`n")
</table>
"@

        return [pscustomobject]@{ Title = 'Diff against previous scans'; Html = $html }
    }

# SmartM365-ActiveDirectory-HealthCheck

function ConvertTo-IsoUtc([datetime]$d){ if($null -eq $d -or $d -eq [datetime]::MinValue){''}else{$d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)} }

function Worst($r){ if($r|Where-Object Status -eq Critical){'Critical'}elseif($r|Where-Object Status -eq Warning){'Warning'}elseif($r|Where-Object Status -eq NotMeasured){'NotMeasured'}else{'OK'} }

function Rank($s){ if($s -eq 'Critical'){0}elseif($s -eq 'Warning'){1}elseif($s -eq 'NotMeasured'){2}else{3} }

function ConvertTo-HtmlSafe($s){ [Net.WebUtility]::HtmlEncode([string]$s) }

function ConvertTo-ReportHtml([object[]]$r,[string]$status,[datetime]$started,[datetime]$ended,[string]$csv,[AllowNull()][object]$forestInfo=$null,[bool]$RemoteDcAdminChecksEnabled=$false,
    [object[]]$Facts=@($DomainFacts.ToArray()),[string]$ReportRunId=$RunId,[string]$ReportHost=$env:COMPUTERNAME){
    $r = @($r)
    $c = @{OK='#107c10';Warning='#ff8c00';Critical='#d13438';NotMeasured='#667085'}[$status]
    $criticalRows = @($r | Where-Object Status -eq Critical)
    $warningRows = @($r | Where-Object Status -eq Warning)
    $okRows = @($r | Where-Object Status -eq OK)
    $notMeasuredRows = @($r | Where-Object Status -eq NotMeasured)
    $find = @($r | Where-Object { $_.Status -in @('Critical','Warning') } | Sort-Object @{Expression={Rank $_.Status}},Domain,DC,Category,Check | Select-Object -First 200)
    $forestName = if ($forestInfo -and $forestInfo.Name) { [string]$forestInfo.Name } else { [string]($r | Where-Object Forest | Select-Object -ExpandProperty Forest -First 1) }
    $rootDomain = if ($forestInfo -and $forestInfo.RootDomain) { [string]$forestInfo.RootDomain } else { [string]($r | Where-Object Domain | Select-Object -ExpandProperty Domain -First 1) }
    if ([string]::IsNullOrWhiteSpace($forestName)) { $forestName = 'NotAvailable' }
    if ([string]::IsNullOrWhiteSpace($rootDomain)) { $rootDomain = 'NotAvailable' }
    $forestMode = if ($forestInfo -and $forestInfo.ForestMode) { [string]$forestInfo.ForestMode } else { 'NotAvailable' }
    $domains = if ($forestInfo -and $forestInfo.Domains) { @($forestInfo.Domains | Sort-Object) } else { @($r | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique | Sort-Object) }
    $domainBadges = (@($domains) | ForEach-Object { '<span class="tag">' + (ConvertTo-HtmlSafe $_) + '</span>' }) -join ' '
    $dcNames = @($r | Where-Object { $_.Category -eq 'Connectivity' -and $_.Check -eq 'Ping' -and $_.DC } | Select-Object -ExpandProperty DC -Unique | Sort-Object)
    $gcCount = if ($facts.Count -gt 0) { ($facts | Measure-Object -Property GlobalCatalogCount -Sum).Sum } else { 'NotMeasured' }
    $lockedTotal = 0
    $lockedRows = @($r | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' })
    foreach ($row in $lockedRows) {
        $value = 0.0
        if ([double]::TryParse([string]$row.NumericValue,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$value)) { $lockedTotal += [int]$value }
    }
    if ($lockedRows.Count -eq 0) { $lockedTotal = 'NotMeasured' }
    $domainsWithCritical = @($criticalRows | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique).Count
    $dcsWithCritical = @($criticalRows | Where-Object DC | Select-Object -ExpandProperty DC -Unique).Count
    $unreachableDcs = @($criticalRows | Where-Object { $_.Category -eq 'Connectivity' -and $_.DC } | Select-Object -ExpandProperty DC -Unique).Count
    $replicationCritical = @($criticalRows | Where-Object Category -eq Replication).Count
    $sysvolCritical = @($criticalRows | Where-Object { $_.Category -eq 'SYSVOL' -and $_.Check -in @('SYSVOL','NETLOGON') }).Count
    $b = [Text.StringBuilder]::new()
    [void]$b.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}.muted{color:#52657a}.small{font-size:11px}.metric{font-size:22px;font-weight:700;color:#0f172a}.tag{display:inline-block;border:1px solid #cbd8e6;background:#f8fbfe;border-radius:999px;padding:3px 8px;margin:2px;font-size:11px}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}.grid{width:100%;border-collapse:separate;border-spacing:8px}.grid td{border:1px solid #dde7f0;background:#f8fbfe;border-radius:6px;padding:10px}.riskCritical{color:#b42318;font-weight:700}.riskWarning{color:#b54708;font-weight:700}.riskNotMeasured{color:#667085;font-weight:700}.riskOK{color:#067647;font-weight:700}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowOK{background:#f3fbf3}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}.rowNotMeasured{background:#f2f4f7}</style></head><body>')
    [void]$b.AppendLine(("<div class='card'><h1>Active Directory Health Check <span class='pill' style='background:{0}'>{1}</span></h1><p class='muted'>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}</p></div>" -f $c,$status,(ConvertTo-HtmlSafe $ReportRunId),(ConvertTo-HtmlSafe $ReportHost),(ConvertTo-IsoUtc $started),(ConvertTo-IsoUtc $ended),(ConvertTo-HtmlSafe ((New-TimeSpan -Start $started -End $ended).ToString()))))
    [void]$b.AppendLine(("<div class='card'><h2>Forest recap</h2><table><tr><th>Forest</th><td>{0}</td><th>Root domain</th><td>{1}</td></tr><tr><th>Forest mode</th><td>{2}</td><th>Remote DC admin checks</th><td>{3}</td></tr><tr><th>Domains</th><td>{4}</td><th>Domain controllers</th><td>{5}</td></tr><tr><th>Global catalogs</th><td>{6}</td><th>Domains scanned</th><td>{7}</td></tr></table></div>" -f (ConvertTo-HtmlSafe $forestName),(ConvertTo-HtmlSafe $rootDomain),(ConvertTo-HtmlSafe $forestMode),$(if($RemoteDcAdminChecksEnabled){'Enabled'}else{'Disabled'}),@($domains).Count,$dcNames.Count,$gcCount,$domainBadges))
    [void]$b.AppendLine(("<div class='card'><h2>Operational recap</h2><table class='grid'><tr><td><div class='metric'>{0}</div><div>Total checks</div></td><td><div class='metric riskCritical'>{1}</div><div>Critical</div></td><td><div class='metric riskWarning'>{2}</div><div>Warning</div></td><td><div class='metric riskOK'>{3}</div><div>OK</div></td></tr><tr><td><div class='metric'>{4}</div><div>Domains with critical findings</div></td><td><div class='metric'>{5}</div><div>DCs with critical findings</div></td><td><div class='metric'>{6}</div><div>Unreachable DCs</div></td><td><div class='metric'>{7}</div><div>Locked user accounts</div></td></tr><tr><td><div class='metric'>{8}</div><div>Replication critical rows</div></td><td><div class='metric'>{9}</div><div>SYSVOL/NETLOGON critical rows</div></td><td><div class='metric'>{10}</div><div>Total DCs</div></td><td><div class='metric'>{11}</div><div>Domains</div></td></tr></table></div>" -f $r.Count,$criticalRows.Count,$warningRows.Count,$okRows.Count,$domainsWithCritical,$dcsWithCritical,$unreachableDcs,$lockedTotal,$replicationCritical,$sysvolCritical,$dcNames.Count,@($domains).Count))
    [void]$b.AppendLine(("<div class='card'><h2>Measurement coverage</h2><p><span class='riskNotMeasured'>{0} NotMeasured</span> checks were not executed or could not be measured. They are excluded from the OK count and do not imply healthy results.</p></div>" -f $notMeasuredRows.Count))
    [void]$b.AppendLine('<div class="card"><h2>FSMO roles</h2><table><tr><th>Scope</th><th>Domain</th><th>Role</th><th>Holder</th><th>Status</th></tr>')
    foreach ($x in @($r | Where-Object Category -eq FSMO | Sort-Object Domain,Check)) { $scope = if ($x.Check -in @('SchemaMaster','DomainNamingMaster')) { 'Forest' } else { 'Domain' }; [void]$b.AppendLine(("<tr class='row{0}'><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{0}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $scope),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.DC))) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Priority findings</h2><table><tr><th>Type</th><th>Name</th><th>Critical</th><th>Warning</th></tr>')
    $priorityGroups = @()
    $priorityGroups += @($r | Where-Object DC | Group-Object DC | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='DC'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    $priorityGroups += @($r | Group-Object Category | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='Category'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    foreach ($g in @($priorityGroups | Where-Object { $_.Critical -gt 0 -or $_.Warning -gt 0 } | Sort-Object @{Expression='Critical';Descending=$true},@{Expression='Warning';Descending=$true},Type,Name | Select-Object -First 20)) { [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td class='riskCritical'>{2}</td><td class='riskWarning'>{3}</td></tr>" -f (ConvertTo-HtmlSafe $g.Type),(ConvertTo-HtmlSafe $g.Name),$g.Critical,$g.Warning)) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Domain status summary</h2><table><tr><th>Domain</th><th>DCs</th><th>OK</th><th>Warning</th><th>Critical</th><th>NotMeasured</th><th>Locked users</th><th>FSMO</th><th>Replication</th><th>SYSVOL</th></tr>')
    foreach ($g in ($r | Where-Object Domain | Group-Object Domain | Sort-Object Name)) {
        $it = @($g.Group)
        # A failed domain scan can leave findings without facts or account measurements.
        $fact = $facts | Where-Object Domain -eq $g.Name | Select-Object -First 1
        $locked = $it | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' } | Select-Object -First 1
        $lockedValue = if ($locked) { $locked.NumericValue } else { 'NotMeasured' }
        $fsmoRows = @($it | Where-Object Category -eq FSMO)
        $repRows = @($it | Where-Object Category -eq Replication)
        $sysvolRows = @($it | Where-Object Category -eq SYSVOL)
        $fsmoWorst = if ($fsmoRows.Count -gt 0) { Worst $fsmoRows } else { 'NotMeasured' }
        $repWorst = if ($repRows.Count -gt 0) { Worst $repRows } else { 'NotMeasured' }
        $sysvolWorst = if ($sysvolRows.Count -gt 0) { Worst $sysvolRows } else { 'NotMeasured' }
        $dcCount = if ($fact) { $fact.DCCount } else { 'NotMeasured' }
        [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td class='risk{7}'>{7}</td><td class='risk{8}'>{8}</td><td class='risk{9}'>{9}</td></tr>" -f (ConvertTo-HtmlSafe $g.Name),$dcCount,@($it | Where-Object Status -eq OK).Count,@($it | Where-Object Status -eq Warning).Count,@($it | Where-Object Status -eq Critical).Count,@($it | Where-Object Status -eq NotMeasured).Count,(ConvertTo-HtmlSafe $lockedValue),$fsmoWorst,$repWorst,$sysvolWorst))
    }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Domain</th><th>DC</th><th>Category</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    if ($r.Count -eq 0) { [void]$b.AppendLine('<tr class="rowNotMeasured"><td colspan="8">No health measurements are available.</td></tr>') }
    elseif ($find.Count -eq 0 -and $notMeasuredRows.Count -gt 0) { [void]$b.AppendLine('<tr class="rowNotMeasured"><td colspan="8">No critical or warning findings; unmeasured checks do not imply healthy results.</td></tr>') }
    elseif ($find.Count -eq 0) { [void]$b.AppendLine('<tr class="rowOK"><td colspan="8">No critical or warning findings.</td></tr>') }
    foreach ($x in $find) { [void]$b.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.DC),(ConvertTo-HtmlSafe $x.Category),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.NumericValue),(ConvertTo-HtmlSafe $x.Threshold),(ConvertTo-HtmlSafe $x.Details))) }
    [void]$b.AppendLine('</table></div><div class="card small"><h2>Technical files</h2><p>CSV: ' + (ConvertTo-HtmlSafe $csv) + '</p></div></body></html>')
    $b.ToString()
}

# SmartM365-Exchange-Hybrid-FreeBusy-Inventory

function Format-HtmlValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return "N/A"
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-FreeBusyReportHtml {
    param(
        [AllowEmptyCollection()][object[]]$Summary = @(),
        [AllowEmptyCollection()][object[]]$Findings = @(),
        [Parameter(Mandatory)][string]$RunStatus,
        [AllowNull()][string]$RunMessage,
        [string]$CheckerVersion = $FreeBusyCheckerVersion,
        [string]$ReportRunId = $RunId
    )

    $statusColor = if ($RunStatus -eq 'Completed') { '#166534' } elseif ($RunStatus -in @('EXONotConnected', 'EXONotAuthorized')) { '#9a3412' } else { '#b91c1c' }
    $ran = $RunStatus -in @('Completed', 'EXONotConnected', 'EXONotAuthorized')
    $measured = @($Summary | Where-Object { $_.SideStatus -in @('Completed', 'NotAuthorizedPartial') })
    # Sides whose checks do not apply (DAuth when free/busy uses OAuth) are neither tested nor missing.
    $applicable = @($Summary | Where-Object { $_.SideStatus -ne 'NotApplicable' })
    $notApplicableRows = @($Summary | Where-Object { $_.SideStatus -eq 'NotApplicable' })
    $sidesText = '{0} / {1}' -f $measured.Count, $applicable.Count
    $issues = @($Findings | Where-Object { $_.Status -in @('Error', 'NotAuthorized', 'Warning') } | Sort-Object @{ Expression = { switch ($_.Status) { 'Error' { 0 } 'NotAuthorized' { 1 } default { 2 } } } }, Side, Auth, RowNumber)
    $errorCount = @($issues | Where-Object Status -eq 'Error').Count
    $warningCount = @($issues | Where-Object Status -eq 'Warning').Count
    $notAuthorizedCount = @($issues | Where-Object Status -eq 'NotAuthorized').Count
    # Counts of a check that did not run are not measured: N/A, never 0.
    $countText = { param($Value) if ($ran) { [string]$Value } else { 'N/A' } }

    $summaryRows = foreach ($row in $Summary) {
        $sideColor = switch ([string]$row.SideStatus) { 'Completed' { '#166534'; break } { $_ -in @('NotRun', 'NotApplicable') } { '#64748b'; break } default { '#9a3412' } }
        $note = if ($row.SideStatus -eq 'NotApplicable') { $row.RunMessage } else { $row.MissingRoles }
        "<tr><td>$(Format-HtmlValue $row.Side)</td><td>$(Format-HtmlValue $row.Auth)</td><td style='font-weight:700;color:$sideColor;'>$(Format-HtmlValue $row.SideStatus)</td><td class='num' style='text-align:right;'>$(Format-HtmlValue $row.ErrorCount)</td><td class='num' style='text-align:right;'>$(Format-HtmlValue $row.WarningCount)</td><td class='num' style='text-align:right;'>$(Format-HtmlValue $row.NotAuthorizedCount)</td><td$(if ($row.SideStatus -eq 'NotApplicable') { " style='color:#64748b;'" })>$(Format-HtmlValue $note)</td></tr>"
    }
    $shown = @($issues | Select-Object -First 100)
    $issueRows = foreach ($row in $shown) {
        $color = switch ([string]$row.Status) { 'Error' { '#b91c1c'; break } 'NotAuthorized' { '#9a3412'; break } default { '#a16207' } }
        "<tr><td style='font-weight:700;color:$color;'>$(Format-HtmlValue $row.Status)</td><td>$(Format-HtmlValue $row.Side)</td><td>$(Format-HtmlValue $row.Auth)</td><td>$(Format-HtmlValue $row.Category)</td><td>$(Format-HtmlValue $row.Item)</td><td>$(Format-HtmlValue $row.Value)</td></tr>"
    }
    $issuesBlock = if (-not $ran) {
        # A check that did not run found nothing: say so, never "no problem reported".
        "<div class='summary' style='border:1px solid #fecaca;border-radius:8px;background:#fef2f2;padding:16px;font-size:14px;line-height:1.55;color:#7f1d1d;margin-bottom:22px;'>The check did not run: $(Format-HtmlValue $RunMessage) No result is available for this run; the free/busy configuration is not measured.</div>"
    }
    elseif ($shown.Count -gt 0) {
        $note = if ($issues.Count -gt $shown.Count) { "<p style='font-size:12px;color:#64748b;'>Showing the first $($shown.Count) of $($issues.Count) results; all are in Exchange_Hybrid_FreeBusy_Findings.csv.</p>" } else { '' }
        "<h2 style='font-size:17px;margin:0 0 10px 0;color:#0f172a;'>Errors, not authorized checks and warnings</h2><table cellpadding='6' cellspacing='0' style='width:100%;border-collapse:collapse;font-size:12px;border:1px solid #e5e7eb;'><tr style='background:#f8fafc;'><th align='left'>Status</th><th align='left'>Side</th><th align='left'>Auth</th><th align='left'>Check</th><th align='left'>Item</th><th align='left'>Value</th></tr>$($issueRows -join "`r`n")</table>$note"
    }
    elseif ($measured.Count -lt $applicable.Count) {
        "<p>No error, warning or not authorized check was reported on the $($measured.Count) side(s) that were measured. The other sides were not run or not connected: their state is unknown, not healthy.</p>"
    }
    else { "<p>No error, warning or not authorized check was reported.</p>" }
    $notApplicableNote = if ($ran -and $notApplicableRows.Count -gt 0) { "<p style='color:#64748b;'>Not applicable: $(Format-HtmlValue $notApplicableRows[0].RunMessage) ($($notApplicableRows.Count) side(s)).</p>" } else { '' }
    $message = if ($ran -and -not [string]::IsNullOrWhiteSpace($RunMessage)) { "<div class='summary' style='border:1px solid #e5e7eb;border-radius:8px;background:#ffffff;padding:16px;font-size:14px;line-height:1.55;color:#334155;margin-bottom:22px;'>$(Format-HtmlValue $RunMessage)</div>" } else { '' }
    $card = {
        param($Label, $Value, $Detail, $Border, $Background, $Accent)
        "<td style='width:20%;vertical-align:top;border:1px solid $Border;border-radius:8px;padding:14px;background:$Background;color:#334155;'><div style='font-size:11px;text-transform:uppercase;color:#64748b;font-weight:800;'>$Label</div><div style='font-size:28px;font-weight:800;margin-top:4px;color:$Accent;'>$(Format-HtmlValue $Value)</div><div style='font-size:12px;color:#64748b;margin-top:2px;'>$Detail</div></td>"
    }
    $cards = @(
        (& $card 'Run status' $RunStatus 'Free/Busy Checker run' '#dbe3ef' '#f8fafc' $statusColor),
        (& $card 'Sides tested' $sidesText 'applicable side and authentication checks measured' '#bfdbfe' '#eff6ff' '#1d4ed8'),
        (& $card 'Errors' (& $countText $errorCount) 'configuration errors' '#fecaca' '#fef2f2' '#b91c1c'),
        (& $card 'Warnings' (& $countText $warningCount) 'items to review' '#fed7aa' '#fff7ed' '#9a3412'),
        (& $card 'Not authorized' (& $countText $notAuthorizedCount) 'Exchange Online checks the app cannot run' '#e9d5ff' '#faf5ff' '#7e22ce')
    )

    return @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Exchange Hybrid Free/Busy</title>
</head>
<body style="margin:0;padding:0;background:#f6f8fb;color:#0f172a;font-family:Segoe UI,Arial,sans-serif;">
<div class="shell" style="max-width:1180px;margin:0 auto;padding:24px;font-family:Segoe UI,Arial,sans-serif;color:#0f172a;">
  <div class="panel" style="background:#ffffff;border:1px solid #dbe3ef;border-radius:10px;overflow:hidden;">
    <div class="hero" style="background:#0f766e;color:#ffffff;padding:24px 28px;">
      <div class="eyebrow" style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;opacity:.9;font-weight:700;color:#ffffff;">Smart365 Exchange Hybrid</div>
      <h1 style="font-size:26px;line-height:1.25;margin:6px 0 0 0;color:#ffffff;">Exchange Hybrid Free/Busy Configuration</h1>
      <div class="subtitle" style="margin-top:10px;font-size:13px;opacity:.95;color:#ffffff;">Microsoft CSS-Exchange Free/Busy Checker $(Format-HtmlValue $CheckerVersion) - RunId $(Format-HtmlValue $ReportRunId)</div>
      <div class="subtitle" style="margin-top:6px;font-size:13px;color:#ffffff;">Run status: <strong>$(Format-HtmlValue $RunStatus)</strong> - sides tested $(Format-HtmlValue $sidesText)</div>
    </div>

    <div class="content" style="padding:24px 28px;">
      <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:10px;border:0;margin:0 0 22px 0;font-size:12px;">
        <tr>
          $($cards -join "`r`n          ")
        </tr>
      </table>
      $message
      $issuesBlock
      $notApplicableNote
      <h2 style="font-size:17px;margin:0 0 10px 0;color:#0f172a;">Sides and authentication</h2>
      <table cellpadding="6" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:12px;border:1px solid #e5e7eb;">
        <tr style="background:#f8fafc;"><th align="left">Side</th><th align="left">Auth</th><th align="left">Status</th><th align="right">Errors</th><th align="right">Warnings</th><th align="right">Not authorized</th><th align="left">Missing roles or note</th></tr>
        $($summaryRows -join "`r`n")
      </table>
      <div class="footer" style="margin-top:18px;padding-top:14px;border-top:1px solid #e5e7eb;font-size:12px;color:#64748b;line-height:1.5;">Sides that were not run or not connected show N/A: their state is not measured. Values that can reveal names, addresses, URLs, certificates or identifiers are masked. Not authorized checks are Exchange Online cmdlets the tenant app cannot run; they are not configuration errors.</div>
    </div>
  </div>
</div>
</body>
</html>
"@
}

# SmartM365-EXO-QuarantineMessages-Report

function Export-SmartM365QuarantineWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$DetailColumns,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][datetime]$StartReceivedDateUtc,
        [Parameter(Mandatory)][datetime]$EndReceivedDateUtc,
        [Parameter(Mandatory)][int]$EffectiveLookbackDays,
        [datetime]$GeneratedUtc = [datetime]::UtcNow
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
        [pscustomobject]@{ Metric = 'Generated UTC'; Value = $GeneratedUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') },
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

    # Daily history from the collected window: one Outlook-safe table bar per received day (UTC).
    $countsByDay = @{}
    foreach ($row in $Rows) {
        $receivedText = [string]$row.ReceivedTimeUtc
        if ($receivedText.Length -lt 10) { continue }
        $dayKey = $receivedText.Substring(0, 10)
        $countsByDay[$dayKey] = 1 + $(if ($countsByDay.ContainsKey($dayKey)) { $countsByDay[$dayKey] } else { 0 })
    }
    $historyDays = [System.Collections.Generic.List[object]]::new()
    $day = $StartReceivedDateUtc.Date
    while ($day -le $EndReceivedDateUtc.Date) {
        $dayKey = $day.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        [void]$historyDays.Add([pscustomobject]@{
            Day   = $day
            Count = $(if ($countsByDay.ContainsKey($dayKey)) { [int]$countsByDay[$dayKey] } else { 0 })
        })
        $day = $day.AddDays(1)
    }
    $maxDayCount = [int](($historyDays | Measure-Object -Property Count -Maximum).Maximum)
    $historyRows = @($historyDays | Sort-Object -Property Day -Descending | ForEach-Object {
        $dayLabel = $_.Day.ToString('ddd dd/MM', [Globalization.CultureInfo]::GetCultureInfo('fr-FR'))
        $barCell = if ($_.Count -gt 0 -and $maxDayCount -gt 0) {
            $barPercent = [math]::Max(1, [math]::Round(100 * $_.Count / $maxDayCount))
            '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="{0}%"><tr><td height="14" bgcolor="#0f4c81" style="background:#0f4c81;height:14px;font-size:1px;line-height:1px;">&nbsp;</td></tr></table>' -f $barPercent
        }
        else {
            '&nbsp;'
        }
        '<tr><td style="padding:2px 8px;white-space:nowrap;font-size:12px;color:#475569;">{0}</td><td style="padding:2px 0;width:100%;">{1}</td><td style="padding:2px 8px;text-align:right;font-size:12px;">{2}</td></tr>' -f $dayLabel, $barCell, $_.Count
    }) -join [Environment]::NewLine

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
  <h3 style="color:#0f4c81;">Blocked messages per received day (UTC)</h3>
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;max-width:720px;border:1px solid #e2e8f0;">
    $historyRows
  </table>
  <p style="font-size:12px;color:#64748b;">Daily counts cover messages still blocked at collection time; messages released since then are not counted.</p>
  <p style="font-size:12px;color:#64748b;">No message body or attachment content is collected by this report.</p>
</body>
</html>
"@
}

# SmartM365-SharePoint-OnPrem-Content-Inventory

function Get-ObservedProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return '' }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return '' }
    try { if ($null -eq $property.Value) { return '' }; return $property.Value }
    catch { throw "Cannot read $Name`: $($_.Exception.Message)" }
}

function Get-ContentMailFarmEdition {
    param([string]$BuildVersion)
    try {
        $mapping = Import-PowerShellDataFile -Path $script:BuildEditionMapPath -ErrorAction Stop
        if ($mapping.Builds.ContainsKey($BuildVersion) -and $mapping.Builds[$BuildVersion] -in @('2016','2019')) {
            return [string]$mapping.Builds[$BuildVersion]
        }
    } catch { WriteLog -Message ("Farm edition mapping unavailable for the combined email: {0}" -f $_.Exception.Message) -Level WARNING }
    return 'Unknown'
}

function New-ContentMailHtml {
    param([string]$Status, [hashtable]$Stats, [string[]]$Failures = @(), [int]$UploadFailures = 0, [string]$TenantName = '', [Nullable[timespan]]$Duration = $null, [string]$HostName = $env:COMPUTERNAME, [string]$ReportGeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'), $Infrastructure = $null, [string]$ContentStatus = '')
    if (-not $ContentStatus) { $ContentStatus = $Status }
    $specs = @(
        @('Web applications','WebApplications','#f8fafc','#dbe3ef','#0f172a'),
        @('Content databases','ContentDatabases','#eff6ff','#bfdbfe','#1d4ed8'),
        @('Site collections','SiteCollections','#f0fdf4','#bbf7d0','#166534'),
        @('Webs','Webs','#faf5ff','#e9d5ff','#7e22ce'),
        @('Coverage issues','CoverageIssues','#fff7ed','#fed7aa','#9a3412'),
        @('Read-only collections','ReadOnlyCollections','#f0fdfa','#99f6e4','#0f766e'),
        @('Unverified locks','UnverifiedLocks','#fef2f2','#fecaca','#b91c1c'),
        @('Content upload failures','UploadFailures','#eef2ff','#c7d2fe','#4338ca')
    )
    $cards = New-Object 'System.Collections.Generic.List[object]'
    foreach ($spec in $specs) {
        $value = if ($spec[1] -eq 'UploadFailures') { if ($Status -eq 'Failed') { 'n/a' } else { [string]$UploadFailures } }
                 elseif ($null -eq $Stats -or $null -eq $Stats[$spec[1]]) { 'n/a' } else { [string]$Stats[$spec[1]] }
        $detail = if ($spec[1] -eq 'UploadFailures' -and $Status -eq 'Failed') { 'not attempted' } elseif ($value -eq 'n/a') { 'not observed' } else { 'observed rows' }
        $cards.Add([pscustomobject]@{Label=$spec[0];Value=$value;Detail=$detail;Background=$spec[2];Border=$spec[3];Accent=$spec[4];Span=1})
    }
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $sections.Add([pscustomobject]@{Title='Summary';Html=(New-SmartM365EmailKpiGridHtml -Cards $cards.ToArray() -RowSizes @(4,4))})
    if ($Infrastructure) {
        $sourceRows = @(
            @('Content',$ContentStatus,'Current inventory run'),
            @('Infrastructure',[string]$Infrastructure.Status,([string]$Infrastructure.CollectedAtUtc + ' ' + [string]$Infrastructure.Detail).Trim())
        )
        $sections.Add([pscustomobject]@{Title='Source status';Html=(New-SmartM365EmailTableHtml -Headers @('Source','Status','Detail') -Rows $sourceRows)})
        if ($Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
            $infraRows = $Infrastructure.Rows
            $serverCount = @($infraRows.Servers | Where-Object { $_.Role -ne 'Invalid' }).Count
            $externalCount = @($infraRows.Servers | Where-Object { $_.Role -eq 'Invalid' }).Count
            $farmBuild = [string]$infraRows.Farms[0].BuildVersion
            $infraCards = @(
                [pscustomobject]@{Label='SharePoint servers';Value=[string]$serverCount;Detail="$externalCount external";Background='#f8fafc';Border='#dbe3ef';Accent='#0f172a';Span=1},
                [pscustomobject]@{Label='Service applications';Value=[string]$infraRows.ServiceApplications.Count;Detail='inventory rows';Background='#eff6ff';Border='#bfdbfe';Accent='#1d4ed8';Span=1},
                [pscustomobject]@{Label='Farm edition';Value=(Get-ContentMailFarmEdition -BuildVersion $farmBuild);Detail=$farmBuild;Background='#f0fdfa';Border='#99f6e4';Accent='#0f766e';Span=1}
            )
            $sections.Add([pscustomobject]@{Title='Infrastructure';Html=(New-SmartM365EmailKpiGridHtml -Cards $infraCards -RowSizes @(3))})
        }
    }
    $alerts = New-Object 'System.Collections.Generic.List[object]'
    foreach ($failure in $Failures) { $alerts.Add(@('Run failure',[string]$failure,'')) }
    if ($Stats -and $null -ne $Stats.CoverageIssues -and $Stats.CoverageIssues -gt 0) { $alerts.Add(@('Coverage incomplete',"$($Stats.CoverageIssues) issue(s)",'See CollectionCoverage.csv')) }
    if ($Stats -and $null -ne $Stats.ReadOnlyCollections -and $Stats.ReadOnlyCollections -gt 0) { $alerts.Add(@('Read-only collections',"$($Stats.ReadOnlyCollections) collection(s)",'See SiteCollections.csv')) }
    if ($Stats -and $null -ne $Stats.UnverifiedLocks -and $Stats.UnverifiedLocks -gt 0) { $alerts.Add(@('Lock state unverified',"$($Stats.UnverifiedLocks) collection(s)",'See SiteCollections.csv')) }
    if ($UploadFailures -gt 0) { $alerts.Add(@('SharePoint upload incomplete',"$UploadFailures file(s)",'')) }
    if ($Infrastructure) {
        if ($Infrastructure.Status -ne 'Qualified') {
            $alerts.Add(@('Infrastructure data unavailable',[string]$Infrastructure.Status,[string]$Infrastructure.Detail))
        } elseif ($Infrastructure.Rows) {
            foreach ($server in @($Infrastructure.Rows.Servers | Where-Object { $_.Role -ne 'Invalid' })) {
                if ($server.Status -and $server.Status -ne 'Online') { $alerts.Add(@('Server not online',[string]$server.ServerName,[string]$server.Status)) }
            }
            foreach ($service in $Infrastructure.Rows.ServiceApplications) {
                if ($service.Status -and $service.Status -notin @('Online','Started')) { $alerts.Add(@('Service application not started',[string]$service.Name,[string]$service.Status)) }
            }
            foreach ($database in $Infrastructure.Rows.ContentDatabases) {
                if ($database.Status -and $database.Status -ne 'Online') { $alerts.Add(@('Content database not online',[string]$database.Name,[string]$database.Status)) }
                if ($database.IsReadOnly -eq 'True') { $alerts.Add(@('Content database read-only',[string]$database.Name,'True')) }
                if ((Get-ObservedProperty $database 'NeedsUpgradeIncludeChildren') -eq 'True' -and (Get-ObservedProperty $database 'NeedsUpgradeStatus') -eq 'Collected') {
                    $alerts.Add(@('Content database needs upgrade',[string]$database.Name,'True'))
                }
            }
            $unverifiedUpgradeCount = @($Infrastructure.Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'NeedsUpgradeStatus') -notin @('Collected','') }).Count
            if ($unverifiedUpgradeCount -gt 0) { $alerts.Add(@('Upgrade state unavailable',"$unverifiedUpgradeCount content databases",'Check NeedsUpgradeStatus in CSV')) }
        }
    }
    if ($alerts.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Alerts';Html=(New-SmartM365EmailTableHtml -Headers @('Issue','Object','Detail') -Rows $alerts.ToArray())}) }
    if ($Infrastructure -and $Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
        $serverRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($server in @($Infrastructure.Rows.Servers | Where-Object { $_.Role -ne 'Invalid' } | Sort-Object ServerName)) {
            $serverRows.Add(@([string]$server.ServerName,[string]$server.Role,[string]$server.Status))
        }
        if ($serverRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Servers';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Role','Status') -Rows $serverRows.ToArray())}) }
        $webRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($web in @($Infrastructure.Rows.WebApplications | Sort-Object DefaultUrl)) {
            $zones = @($Infrastructure.Rows.WebApplicationZones | Where-Object { $_.WebApplicationId -eq $web.WebApplicationId } | Sort-Object Zone | ForEach-Object { '{0}: {1}' -f $_.Zone,$_.AuthenticationMode }) -join '; '
            $webRows.Add(@([string]$web.DefaultUrl,[string]$web.ApplicationPool,[string]$web.ContentDatabaseCount,[string]$zones))
        }
        if ($webRows.Count -gt 0) { $sections.Add([pscustomobject]@{Title='Web applications';Html=(New-SmartM365EmailTableHtml -Headers @('URL','Application pool','Content databases','Authentication by zone') -Rows $webRows.ToArray() -NoWrapColumns @(0,1))}) }
        $sized = @($Infrastructure.Rows.ContentDatabases | Where-Object { (Get-ObservedProperty $_ 'DatabaseSizeStatus') -eq 'Collected' -and (Get-ObservedProperty $_ 'DatabaseSizeBytes') -match '^\d+$' } | Sort-Object @{Expression={[decimal]$_.DatabaseSizeBytes};Descending=$true} | Select-Object -First 5)
        if ($sized.Count -gt 0) {
            $sizeRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($database in $sized) { $sizeRows.Add(@([string]$database.Name,[string]$database.DatabaseServer,[string]$database.DatabaseSizeBytes)) }
            $sections.Add([pscustomobject]@{Title='Top 5 largest content databases';Html=(New-SmartM365EmailTableHtml -Headers @('Name','Server','Size (bytes)') -Rows $sizeRows.ToArray())})
        }
    }
    if ($Stats -and $Stats.CoverageSamples.Count -gt 0) {
        $issueRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Stats.CoverageSamples) { $issueRows.Add(@([string]$item.Status,[string]$item.Url,[string]$item.ErrorMessage)) }
        $sections.Add([pscustomobject]@{Title='Collection coverage examples';Html=(New-SmartM365EmailTableHtml -Headers @('Status','Collection or database','Detail') -Rows $issueRows.ToArray())})
    }
    if ($Stats -and $Stats.ReadOnlySamples.Count -gt 0) {
        $readOnlyRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Stats.ReadOnlySamples) { $readOnlyRows.Add(@([string]$item.Url,[string]$item.LockState,[string]$item.ContentDatabaseIsReadOnly)) }
        $sections.Add([pscustomobject]@{Title='Read-only collection examples';Html=(New-SmartM365EmailTableHtml -Headers @('URL','Lock state','Database read-only') -Rows $readOnlyRows.ToArray() -NoWrapColumns @(0))})
    }
    if ($Stats -and $Status -in @('Qualified','Partial') -and $null -ne $Stats.SiteCollections) {
        $fileRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in @(@('SiteCollections',$Stats.SiteCollections),@('SiteAdministrators',$Stats.SiteAdministrators),@('Webs',$Stats.Webs),@('CollectionCoverage',$Stats.CollectionCoverageRows))) {
            $fileRows.Add(@("SharePoint_OnPrem_$($item[0]).csv",[string]$item[1]))
        }
        if ($Infrastructure -and $Infrastructure.Status -eq 'Qualified' -and $Infrastructure.Rows) {
            foreach ($name in @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')) {
                $fileRows.Add(@("SharePoint_OnPrem_$name.csv",[string]$Infrastructure.Rows[$name].Count))
            }
        }
        $sections.Add([pscustomobject]@{Title='Files';Html=(New-SmartM365EmailTableHtml -Headers @('CSV file','Rows') -Rows $fileRows.ToArray())})
    }
    $severity = switch ($Status) { 'Qualified' {'Success'} 'Partial' {'Warning'} default {'Error'} }
    $durationText = if ($null -ne $Duration) { ([timespan]$Duration).ToString('hh\:mm\:ss') } else { '' }
    $message = if ($Status -eq 'Failed') { 'SharePoint on-premises content inventory stopped before full publication. Values below reflect observations made before the failure.' } else { 'SharePoint on-premises content inventory summary generated from the current SmartM365 CSV outputs.' }
    return New-SmartM365EmailBody -Title 'SharePoint on-prem inventory summary' -Category 'SmartM365 SharePoint OnPrem' -Severity $severity -Tenant $TenantName -HostName $HostName -GeneratedAt $ReportGeneratedAt -StatusBadge $Status.ToUpperInvariant() -Duration $durationText -Message $message -Sections $sections.ToArray() -Footer 'This automated message was generated by SmartM365 from SharePoint on-premises inventory data.'
}

# SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory

function ConvertTo-ServersAndStorageEmailHtmlText {
    param([AllowNull()]$Value)

    if (Get-Command ConvertTo-CoreSmartM365EmailHtmlText -ErrorAction SilentlyContinue) {
        return ConvertTo-CoreSmartM365EmailHtmlText -Value $Value
    }

    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-ServersAndStorageEmailLinkHtml {
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$Url
    )

    $safeText = ConvertTo-ServersAndStorageEmailHtmlText -Value $Text
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        $safeUrl = ConvertTo-ServersAndStorageEmailHtmlText -Value $Url
        return "<a href=`"$safeUrl`" style=`"color:#2563eb;text-decoration:underline;`">$safeText</a>"
    }

    return $safeText
}

function ConvertTo-ServersAndStorageMailNumber {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [double]$Value }
    $text = ([string]$Value).Trim()
    if (-not $text) { return $null }
    $number = 0.0
    # Regenerated runs read numbers back from CSV text.
    foreach ($culture in @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::CurrentCulture)) {
        if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, $culture, [ref]$number)) { return $number }
    }
    return $null
}

function Format-ServersAndStorageMailNumber {
    param([AllowNull()]$Value, [int]$Decimals = 0, [string]$Unit)

    $number = ConvertTo-ServersAndStorageMailNumber -Value $Value
    if ($null -eq $number) { return 'N/A' }
    $text = $number.ToString(('N{0}' -f $Decimals), [Globalization.CultureInfo]::InvariantCulture)
    if ($Unit) { return "$text $Unit" }
    return $text
}

function Limit-ServersAndStorageMailText {
    param([AllowNull()]$Value, [int]$MaxLength = 140)

    $text = (([string]$Value) -replace '\s+', ' ').Trim()
    if ($text.Length -le $MaxLength) { return $text }
    return $text.Substring(0, $MaxLength - 3) + '...'
}

function ConvertTo-ServersAndStorageMailDate {
    param([AllowNull()]$Value)

    if ($Value -is [datetime]) { return $Value }
    $text = ([string]$Value).Trim()
    if (-not $text -or $text -eq 'Unknown') { return $null }
    $date = [datetime]::MinValue
    # [string] of a DateTime is invariant; values read on the server may follow its culture.
    foreach ($culture in @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::CurrentCulture)) {
        if ([datetime]::TryParse($text, $culture, [Globalization.DateTimeStyles]::None, [ref]$date)) { return $date }
    }
    return $null
}

function New-ServersAndStorageMailTableHtml {
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][int[]]$ColumnWidths,
        [int[]]$NoWrapColumns = @(),
        [switch]$BreakAll
    )

    # Shared Core table, made fixed-layout with explicit column widths so long values wrap inside the 760px frame.
    $html = [string](New-CoreSmartM365EmailTableHtml -Headers $Headers -Rows $Rows -NoWrapColumns $NoWrapColumns)
    $html = $html.Replace('<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;', '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;table-layout:fixed;')
    $headerMarker = '<th align="left" style="'
    $parts = $html.Split([string[]]@($headerMarker), [StringSplitOptions]::None)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append($parts[0])
    for ($index = 1; $index -lt $parts.Count; $index++) {
        $width = if ($index -le $ColumnWidths.Count) { $ColumnWidths[$index - 1] } else { 0 }
        if ($width -gt 0) { [void]$builder.Append(('<th align="left" width="{0}%" style="width:{0}%;' -f $width)) }
        else { [void]$builder.Append($headerMarker) }
        [void]$builder.Append($parts[$index])
    }
    $breakStyle = if ($BreakAll) { 'word-break:break-all;' } else { 'word-break:break-word;' }
    return $builder.ToString().Replace('word-break:break-word;', $breakStyle + 'overflow-wrap:anywhere;')
}

function New-ServersAndStorageMailNoteHtml {
    param([Parameter(Mandatory)][string]$Text)

    return '<div style="font-size:12px;line-height:18px;color:#475569;margin-top:8px;">{0}</div>' -f (ConvertTo-ServersAndStorageEmailHtmlText -Value $Text)
}

function New-ServersAndStorageMailBodyHtml {
    param(
        [Parameter(Mandatory)][object]$Summary,
        [AllowNull()][object[]]$PerServerSummary = @(),
        [AllowNull()][object[]]$ReadinessInventory = @(),
        [AllowNull()][object[]]$HealthCheckerSummary = @(),
        [AllowNull()][object[]]$HealthCheckerFindings = @(),
        [AllowNull()][object[]]$Files = @(),
        [string]$Tenant,
        [int]$WarningCount = 0,
        [int]$ErrorCount = 0,
        [AllowNull()]$CertificateBudgetSeconds = 120,
        [int]$CertificateExpiryDays = 60,
        [datetime]$ReferenceDate = (Get-Date),
        [string]$ReportHost = $env:COMPUTERNAME
    )

    $servers = @($PerServerSummary | Where-Object { $null -ne $_ } | Sort-Object ExchangeServerName)
    $readinessRows = @($ReadinessInventory | Where-Object { $null -ne $_ })
    $hcRows = @($HealthCheckerSummary | Where-Object { $null -ne $_ })
    $hcFindings = @($HealthCheckerFindings | Where-Object { $null -ne $_ })
    $toInt = { param($Value) $number = ConvertTo-ServersAndStorageMailNumber -Value $Value; if ($null -eq $number) { 0 } else { [int]$number } }
    $summaryValue = { param($Name) Get-ObjectPropertyValue -InputObject $Summary -Name $Name }

    $excludedServerNames = @(([string](& $summaryValue 'ExcludedServers') -split ';') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    # Health Checker: an empty summary means it was not part of this run (switch or configuration).
    $hcIncluded = $hcRows.Count -gt 0
    $hcExcludedRows = @($hcRows | Where-Object { [string]$_.RunStatus -eq 'NotRun' -and [string]$_.RunMessage -match 'excluded by configuration' })
    $hcScopedRows = @($hcRows | Where-Object { $hcExcludedRows -notcontains $_ })
    $hcCompletedRows = @($hcScopedRows | Where-Object { [string]$_.RunStatus -eq 'Completed' })
    $hcNotCompletedRows = @($hcScopedRows | Where-Object { [string]$_.RunStatus -ne 'Completed' })
    $hcVulnerabilities = 0; $hcWarnings = 0
    foreach ($row in $hcCompletedRows) { $hcVulnerabilities += & $toInt $row.VulnerabilityCount; $hcWarnings += & $toInt $row.WarningCount }
    $hcCriticalFindings = @($hcFindings | Where-Object { [string]$_.Status -eq 'Error' }).Count

    # Certificates: rows of category ExchangeCertificate written by the certificate phase (LocalFirst/RemoteRetry).
    $certificateRows = @($readinessRows | Where-Object { [string]$_.Category -eq 'ExchangeCertificate' })
    $certificateItemGroups = @($certificateRows | Where-Object { [string]$_.ObjectName -match ':' } | Group-Object ObjectName)
    $certificateFailures = @($certificateRows | Where-Object { [string]$_.Setting -eq 'Collection' -and [string]$_.Value -eq 'Failed' })
    $certificateRetries = @($certificateRows | Where-Object { [string]$_.Setting -eq 'CollectionPath' -and [string]$_.Value -eq 'RemoteRetry' })
    $certificateUnknownRows = @($certificateRows | Where-Object { [string]$_.Value -eq 'Unknown' -and [string]$_.CollectionStatus -eq 'WARNING' })
    $certificates = @(foreach ($group in $certificateItemGroups) {
        $values = @{}
        foreach ($row in $group.Group) { if (-not $values.ContainsKey([string]$row.Setting)) { $values[[string]$row.Setting] = $row.Value } }
        $serverName = if ($values['Server']) { [string]$values['Server'] } else { ([string]$group.Name -split ':')[0] }
        $notAfter = ConvertTo-ServersAndStorageMailDate -Value $values['NotAfter']
        [pscustomobject]@{
            Server   = $serverName
            Subject  = [string]$values['Subject']
            Services = [string]$values['Services']
            NotAfter = $notAfter
            DaysLeft = $(if ($null -ne $notAfter) { [int][Math]::Ceiling(($notAfter - $ReferenceDate).TotalDays) } else { $null })
            Expired  = ($null -ne $notAfter -and $notAfter -lt $ReferenceDate)
        }
    })
    $expiringCertificates = @($certificates | Where-Object { $null -ne $_.DaysLeft -and $_.DaysLeft -le $CertificateExpiryDays } | Sort-Object DaysLeft, Server)
    $undatedCertificates = @($certificates | Where-Object { $null -eq $_.DaysLeft })

    $infrastructureErrors = (& $toInt (& $summaryValue 'ComputeCollectionErrors')) + (& $toInt (& $summaryValue 'DiskDriveCollectionErrors')) + (& $toInt (& $summaryValue 'LogicalDiskCollectionErrors'))
    $readinessErrorRows = @($readinessRows | Where-Object { $_.CollectionStatus -eq 'ERROR' -or $_.Importance -eq 'Error' })
    $readinessWarningRows = @($readinessRows | Where-Object { $_.CollectionStatus -eq 'WARNING' -or $_.Importance -eq 'Warning' })
    $lowSpaceWarnings = & $toInt (& $summaryValue 'LowSpaceWarnings')

    $severity = 'Success'
    if ($ErrorCount -gt 0 -or $infrastructureErrors -gt 0 -or $readinessErrorRows.Count -gt 0) { $severity = 'Error' }
    elseif ($WarningCount -gt 0 -or $readinessWarningRows.Count -gt 0 -or $lowSpaceWarnings -gt 0 -or $hcNotCompletedRows.Count -gt 0 -or $hcVulnerabilities -gt 0) { $severity = 'Warning' }
    $statusBadge = switch ($severity) { 'Error' { 'Completed with errors'; break } 'Warning' { 'Completed with warnings'; break } default { 'Completed' } }

    $okCard = @{ Background = '#f0fdf4'; Border = '#bbf7d0'; Accent = '#166534' }
    $warnCard = @{ Background = '#fff7ed'; Border = '#fed7aa'; Accent = '#9a3412' }
    $errorCard = @{ Background = '#fef2f2'; Border = '#fecaca'; Accent = '#b91c1c' }
    $newCard = {
        param($Label, $Value, $Detail, $Colors)
        [pscustomobject]@{ Label = $Label; Value = $Value; Detail = $Detail; Background = $Colors.Background; Border = $Colors.Border; Accent = $Colors.Accent; Span = 1; FontSize = 24 }
    }
    $neutral = @{ Background = '#f8fafc'; Border = '#dbe3ef'; Accent = '#0f172a' }
    $blue = @{ Background = '#eff6ff'; Border = '#bfdbfe'; Accent = '#1d4ed8' }
    $indigo = @{ Background = '#eef2ff'; Border = '#c7d2fe'; Accent = '#4338ca' }

    $virtualMachines = @($servers | Where-Object { [string]$_.IsVirtualMachine -eq 'True' }).Count
    $serversWithCompute = @($servers | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.IsVirtualMachine) }).Count
    $hcCard = if (-not $hcIncluded) {
        & $newCard 'Health Checker' 'Not run' 'disabled for this run' $neutral
    } else {
        $colors = if ($hcNotCompletedRows.Count -gt 0 -or $hcVulnerabilities -gt 0) { $warnCard } else { $okCard }
        & $newCard 'Health Checker' ('{0}/{1}' -f $hcCompletedRows.Count, $hcScopedRows.Count) ('{0} vulnerabilities, {1} errors' -f $hcVulnerabilities, $hcCriticalFindings) $colors
    }
    $certificateCard = if ($certificateRows.Count -eq 0) {
        & $newCard 'Certificates' 'N/A' 'no certificate data' $neutral
    } else {
        $colors = if ($certificateFailures.Count -gt 0) { $errorCard } elseif ($expiringCertificates.Count -gt 0 -or $certificateUnknownRows.Count -gt 0) { $warnCard } else { $okCard }
        & $newCard 'Certificates' $certificates.Count ('{0} expiring, {1} failed' -f $expiringCertificates.Count, $certificateFailures.Count) $colors
    }
    $cards = @(
        & $newCard 'Servers' (Format-ServersAndStorageMailNumber (& $summaryValue 'ExchangeServersCount')) ('{0} excluded' -f $excludedServerNames.Count) $neutral
        & $newCard 'Virtual machines' $virtualMachines ('of {0} with compute data' -f $serversWithCompute) $blue
        & $newCard 'vCPU' (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalLogicalProcessorCount')) ('{0} physical cores' -f (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalPhysicalCoreCount'))) $blue
        & $newCard 'RAM' (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalMemoryGB') -Unit 'GB') 'total memory' $blue
        & $newCard 'Disks' (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalDiskDriveCount')) 'physical disk drives' $indigo
        & $newCard 'Provisioned storage' (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalDiskDriveSizeTB') -Decimals 2 -Unit 'TB') ('{0} used' -f (Format-ServersAndStorageMailNumber (& $summaryValue 'TotalLogicalDiskUsedTB') -Decimals 2 -Unit 'TB')) $indigo
        $hcCard
        $certificateCard
    )
    $sections = New-Object System.Collections.Generic.List[object]
    $sections.Add([pscustomobject]@{ Title = 'Key figures'; Html = (New-CoreSmartM365EmailKpiGridHtml -Cards $cards -RowSizes @(4, 4)) })

    # Per-server summary.
    $serverRows = New-Object System.Collections.Generic.List[object]
    foreach ($server in $servers) {
        $name = [string]$server.ExchangeServerName
        $failedParts = @()
        if ([string]$server.ComputeCollectionStatus -eq 'ERROR') { $failedParts += 'compute' }
        if ([string]$server.DiskDriveCollectionStatus -eq 'ERROR') { $failedParts += 'disks' }
        if ([string]$server.LogicalDiskCollectionStatus -eq 'ERROR') { $failedParts += 'volumes' }
        $collection = if ($failedParts.Count -gt 0) { 'ERROR: ' + ($failedParts -join ', ') } elseif ((& $toInt $server.LowSpaceLogicalDiskCount) -gt 0) { 'Low space' } else { 'OK' }
        $hcRow = @($hcRows | Where-Object { [string]$_.ServerName -eq $name } | Select-Object -First 1)
        $hcText = if (-not $hcIncluded) { 'Not run' }
            elseif ($hcRow.Count -eq 0) { 'No result' }
            elseif ([string]$hcRow[0].RunStatus -eq 'Completed') { '{0} vuln, {1} err' -f (& $toInt $hcRow[0].VulnerabilityCount), (& $toInt $hcRow[0].ErrorCount) }
            else { [string]$hcRow[0].RunStatus }
        $certificateText = if ($certificateRows.Count -eq 0) { 'N/A' }
            elseif (@($certificateFailures | Where-Object { [string]$_.ObjectName -eq $name }).Count -gt 0) { 'Failed' }
            else {
                $text = [string]@($certificates | Where-Object { $_.Server -eq $name }).Count
                if (@($certificateRetries | Where-Object { [string]$_.ObjectName -eq $name }).Count -gt 0) { $text += ' (RemoteRetry)' }
                $text
            }
        $serverRows.Add(@($name, (Format-ServersAndStorageMailNumber $server.LogicalProcessorCount), (Format-ServersAndStorageMailNumber $server.MemoryGB), ('{0} / {1} GB' -f (Format-ServersAndStorageMailNumber $server.DiskDriveCount), (Format-ServersAndStorageMailNumber $server.DiskDriveTotalSizeGB)), $collection, $hcText, $certificateText))
    }
    foreach ($name in $excludedServerNames) { $serverRows.Add(@($name, 'N/A', 'N/A', 'N/A', 'Excluded', 'Excluded', 'Excluded')) }
    $maxServerRows = 25
    if ($serverRows.Count -gt 0) {
        $serverHtml = New-ServersAndStorageMailTableHtml -Headers @('Server', 'vCPU', 'RAM GB', 'Disks / size', 'Collection', 'Health Checker', 'Certificates') -Rows @($serverRows | Select-Object -First $maxServerRows) -ColumnWidths @(22, 8, 10, 16, 14, 15, 15) -NoWrapColumns @(1, 2)
        if ($serverRows.Count -gt $maxServerRows) { $serverHtml += New-ServersAndStorageMailNoteHtml -Text ('Showing {0} of {1} servers; all servers are in the full report.' -f $maxServerRows, $serverRows.Count) }
    } else {
        $serverHtml = New-ServersAndStorageMailNoteHtml -Text 'No per-server data is available for this run.'
    }
    $sections.Add([pscustomobject]@{ Title = 'Per-server summary'; Html = $serverHtml })

    # Collection issues: errors first, then warnings and partial reads.
    $issues = New-Object System.Collections.Generic.List[object]
    foreach ($server in $servers) {
        $name = [string]$server.ExchangeServerName
        foreach ($pair in @(@('ComputeCollectionStatus', 'Compute', 'CPU/RAM inventory failed or returned no data.'), @('DiskDriveCollectionStatus', 'Disk drives', 'Physical disk inventory failed or returned errors.'), @('LogicalDiskCollectionStatus', 'Logical disks', 'Logical disk inventory failed or returned errors.'))) {
            if ([string](Get-ObjectPropertyValue -InputObject $server -Name $pair[0]) -eq 'ERROR') { $issues.Add([pscustomobject]@{ Rank = 0; Values = @('Infrastructure', $name, $pair[1], $pair[2]) }) }
        }
        $lowSpace = & $toInt $server.LowSpaceLogicalDiskCount
        if ($lowSpace -gt 0) { $issues.Add([pscustomobject]@{ Rank = 2; Values = @('Capacity', $name, 'Low disk space', ('{0} logical disk(s) below the free-space threshold.' -f $lowSpace)) }) }
    }
    foreach ($row in $certificateFailures) {
        $issues.Add([pscustomobject]@{ Rank = 0; Values = @('Certificates', [string]$row.ObjectName, 'Collection failed', (Limit-ServersAndStorageMailText -Value $row.ErrorMessage -MaxLength 180)) })
    }
    foreach ($row in $certificateRetries) {
        $retryServer = [string]$row.ObjectName
        $retryText = if (@($certificates | Where-Object { $_.Server -eq $retryServer }).Count -gt 0) { 'Native read failed or timed out; certificates were read through the remote session.' } else { 'Native read failed or timed out; remote retry attempted.' }
        $issues.Add([pscustomobject]@{ Rank = 3; Values = @('Certificates', $retryServer, 'RemoteRetry', $retryText) })
    }
    foreach ($group in @($certificateUnknownRows | Group-Object { ([string]$_.ObjectName -split ':')[0] } | Sort-Object Name)) {
        $settings = @($group.Group | ForEach-Object { [string]$_.Setting } | Sort-Object -Unique) -join ', '
        $certificateCount = @($group.Group | ForEach-Object { [string]$_.ObjectName } | Sort-Object -Unique).Count
        $issues.Add([pscustomobject]@{ Rank = 2; Values = @('Certificates', $group.Name, 'Unknown metadata', ('{0} certificate(s): {1} unavailable; no value inferred.' -f $certificateCount, $settings)) })
    }
    foreach ($row in @($readinessRows | Where-Object { [string]$_.Category -eq 'CollectorError' })) {
        $issues.Add([pscustomobject]@{ Rank = 1; Values = @('Readiness', [string]$row.ObjectName, 'Collector failed', (Limit-ServersAndStorageMailText -Value $row.ErrorMessage -MaxLength 180)) })
    }
    foreach ($row in $hcNotCompletedRows) {
        $issues.Add([pscustomobject]@{ Rank = 1; Values = @('Health Checker', [string]$row.ServerName, [string]$row.RunStatus, (Limit-ServersAndStorageMailText -Value $row.RunMessage -MaxLength 180)) })
    }
    $maxIssueRows = 15
    $sortedIssues = @($issues | Sort-Object Rank, @{ Expression = { $_.Values[0] } }, @{ Expression = { $_.Values[1] } })
    $issueNote = 'Readiness: {0} error(s), {1} warning(s) in total; every row is in the full report and Exchange_OnPrem_MigrationReadiness_Config.csv.' -f $readinessErrorRows.Count, $readinessWarningRows.Count
    if ($sortedIssues.Count -gt 0) {
        $issueHtml = New-ServersAndStorageMailTableHtml -Headers @('Area', 'Object', 'Issue', 'Details') -Rows @($sortedIssues | Select-Object -First $maxIssueRows | ForEach-Object { , $_.Values }) -ColumnWidths @(15, 22, 17, 46) -BreakAll
        if ($sortedIssues.Count -gt $maxIssueRows) { $issueNote = ('Showing {0} of {1} issues. ' -f $maxIssueRows, $sortedIssues.Count) + $issueNote }
        $issueHtml += New-ServersAndStorageMailNoteHtml -Text $issueNote
    } else {
        $issueHtml = New-ServersAndStorageMailNoteHtml -Text ('No collection failure was detected. ' + $issueNote)
    }
    $sections.Add([pscustomobject]@{ Title = 'Collection issues'; Html = $issueHtml })

    # Health Checker status, explicit when it did not run.
    $hcHtml = if (-not $hcIncluded) {
        New-ServersAndStorageMailNoteHtml -Text 'Health Checker was not run in this run (disabled by parameter or configuration). No Health Checker state is reported.'
    } else {
        $hcVersion = @($hcRows | ForEach-Object { [string]$_.HealthCheckerVersion } | Where-Object { $_ } | Select-Object -Unique) -join ', '
        $text = 'Health Checker {0} completed on {1} of {2} server(s). Vulnerabilities: {3}; critical findings (Status=Error): {4}; warnings: {5}.' -f $hcVersion, $hcCompletedRows.Count, $hcScopedRows.Count, $hcVulnerabilities, $hcCriticalFindings, $hcWarnings
        if ($hcNotCompletedRows.Count -gt 0) { $text += ' Not completed: {0} server(s); their state is not measured.' -f $hcNotCompletedRows.Count }
        if ($hcExcludedRows.Count -gt 0) { $text += ' Excluded by configuration: {0} server(s).' -f $hcExcludedRows.Count }
        New-ServersAndStorageMailNoteHtml -Text $text
    }
    $sections.Add([pscustomobject]@{ Title = 'Health Checker'; Html = $hcHtml })

    # Certificates expiring soon, when certificate data exists.
    if ($certificateRows.Count -gt 0) {
        $certificateText = '{0} certificate(s) read; {1} expiring within {2} days or expired; {3} server(s) failed; {4} RemoteRetry; per-server budget {5} s.' -f $certificates.Count, $expiringCertificates.Count, $CertificateExpiryDays, $certificateFailures.Count, $certificateRetries.Count, (Format-ServersAndStorageMailNumber $CertificateBudgetSeconds)
        if ($undatedCertificates.Count -gt 0) { $certificateText += ' Expiry date unavailable: {0}.' -f $undatedCertificates.Count }
        $certificateHtml = ''
        if ($expiringCertificates.Count -gt 0) {
            $maxCertificateRows = 10
            $certificateTableRows = @($expiringCertificates | Select-Object -First $maxCertificateRows | ForEach-Object {
                , @($_.Server, (Limit-ServersAndStorageMailText -Value $_.Subject -MaxLength 90), $_.NotAfter.ToString('yyyy-MM-dd'), $(if ($_.Expired) { 'Expired' } else { [string]$_.DaysLeft }), (Limit-ServersAndStorageMailText -Value $_.Services -MaxLength 40))
            })
            $certificateHtml = New-ServersAndStorageMailTableHtml -Headers @('Server', 'Subject', 'Expires', 'Days', 'Services') -Rows $certificateTableRows -ColumnWidths @(18, 40, 15, 10, 17) -NoWrapColumns @(2, 3) -BreakAll
            if ($expiringCertificates.Count -gt $maxCertificateRows) { $certificateText += ' Showing {0} of {1}.' -f $maxCertificateRows, $expiringCertificates.Count }
        }
        $sections.Add([pscustomobject]@{ Title = 'Certificates'; Html = ($certificateHtml + (New-ServersAndStorageMailNoteHtml -Text $certificateText)) })
    }

    # Files: the full HTML report and CSV files, linked to SharePoint when uploaded.
    $fileRows = foreach ($file in @($Files | Where-Object { $null -ne $_ })) {
        $locationText = if ($file.SharePointPath) { [string]$file.SharePointPath } else { [string]$file.Path }
        $locationHtml = New-ServersAndStorageEmailLinkHtml -Text $locationText -Url ([string]$file.WebUrl)
        "<tr><td width=`"30%`" style=`"width:30%;border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#334155;word-break:break-all;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value $file.Name)</td><td width=`"56%`" style=`"width:56%;border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:11px;color:#334155;word-break:break-all;`">$locationHtml</td><td width=`"14%`" align=`"right`" style=`"width:14%;border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;white-space:nowrap;`">$(ConvertTo-ServersAndStorageEmailHtmlText -Value ('{0} KB' -f $file.SizeKB))</td></tr>"
    }
    $filesHtml = if (@($fileRows).Count -gt 0) {
        "<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`" style=`"width:100%;table-layout:fixed;border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;`">$($fileRows -join "`r`n")</table>"
    } else { New-ServersAndStorageMailNoteHtml -Text 'No output file is available.' }
    $sections.Add([pscustomobject]@{ Title = 'Files'; Html = $filesHtml })

    $runId = [string](& $summaryValue 'RunId')
    $executionDate = [string](& $summaryValue 'ExecutionDate')
    return New-CoreSmartM365EmailBody `
        -Title 'Infrastructure and readiness summary' `
        -Category 'SmartM365 Exchange OnPrem' `
        -Severity $severity `
        -StatusBadge $statusBadge `
        -Tenant $Tenant `
        -HostName $ReportHost -GeneratedAt ($ReferenceDate.ToString('yyyy-MM-dd HH:mm:ss zzz')) `
        -Message ('Exchange on-premises infrastructure and readiness inventory, run {0} ({1}). The complete HTML report is unchanged and linked in Files.' -f $runId, $executionDate) `
        -Sections $sections.ToArray() `
        -Footer 'This automated message was generated by SmartM365 from Exchange on-premises infrastructure and readiness inventory data.'
}

function Get-ObjectPropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

# SmartM365-Exchange-Local-Mailboxes-Inventory

function Get-SmartM365MailboxReportCsvRowCount {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return 0 }
    try {
        $lineCount = @(Get-Content -LiteralPath $Path -ErrorAction Stop | Measure-Object -Line).Lines
        if ($lineCount -le 1) { return 0 }
        return ($lineCount - 1)
    }
    catch { return 0 }
}

function New-SmartM365ExchangeMailboxDailyDiffSection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Current, [AllowEmptyCollection()][object[]]$History = @())

    $earlier = @($History | Where-Object { $_.SnapshotDate -lt $Current.SnapshotDate } | Sort-Object SnapshotDate,GeneratedAt -Descending)
    $previous = if ($earlier.Count) { $earlier[0] } else { $null }
    $date = [datetime]::ParseExact([string]$Current.SnapshotDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $weekDate = $date.AddDays(-7).ToString('yyyy-MM-dd')
    $monthDate = $date.AddDays(-30).ToString('yyyy-MM-dd')
    $week = @($earlier | Where-Object { $_.SnapshotDate -eq $weekDate } | Select-Object -First 1)
    $month = @($earlier | Where-Object { $_.SnapshotDate -eq $monthDate } | Select-Object -First 1)
    $comparisons = @(
        [pscustomobject]@{ Label = 'Previous'; DateText = if ($previous) { $previous.SnapshotDate } else { 'n/a' }; Snapshot = $previous }
        [pscustomobject]@{ Label = 'J-7'; DateText = $weekDate; Snapshot = if ($week.Count) { $week[0] } else { $null } }
        [pscustomobject]@{ Label = 'J-30'; DateText = $monthDate; Snapshot = if ($month.Count) { $month[0] } else { $null } }
    )
    $metrics = @(
        [pscustomobject]@{ Label = 'Domains'; Property = 'DomainCount'; Size = $false }
        [pscustomobject]@{ Label = 'Total mailboxes'; Property = 'TotalMailboxCount'; Size = $false }
        [pscustomobject]@{ Label = 'Local mailboxes'; Property = 'TotalLocalMailboxCount'; Size = $false }
        [pscustomobject]@{ Label = 'Remote mailboxes'; Property = 'TotalRemoteMailboxCount'; Size = $false }
        [pscustomobject]@{ Label = 'Enabled accounts'; Property = 'EnabledAccounts'; Size = $false }
        [pscustomobject]@{ Label = 'Disabled accounts'; Property = 'DisabledAccounts'; Size = $false }
        [pscustomobject]@{ Label = 'Local mailbox size (GB)'; Property = 'TotalLocalMailboxSizeGB'; Size = $true }
    )
    $format = { param($Value,$Size) if ($Size) { return ([double]$Value).ToString('N2') }; return ([long]$Value).ToString('N0') }
    $headers = foreach ($comparison in $comparisons) {
        '<th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">{0} ({1})</th>' -f $comparison.Label,(ConvertTo-SmartM365EmailHtmlText $comparison.DateText)
    }
    $bodyRows = foreach ($metric in $metrics) {
        $currentValue = $Current.($metric.Property)
        $currentHtml = ConvertTo-SmartM365EmailHtmlText (& $format $currentValue $metric.Size)
        $cells = foreach ($comparison in $comparisons) {
            $snapshot = $comparison.Snapshot
            $comparable = ($null -ne $snapshot -and (-not $metric.Size -or ($Current.LocalSizeCoverageComplete -and $snapshot.LocalSizeCoverageComplete)))
            $cell = 'n/a'
            if ($comparable) {
                $value = $snapshot.($metric.Property)
                $delta = if ($metric.Size) { [math]::Round(([double]$currentValue - [double]$value), 2) } else { [long]$currentValue - [long]$value }
                $deltaText = if ($delta -gt 0) { '+' + (& $format $delta $metric.Size) } elseif ($delta -lt 0) { '-' + (& $format ([math]::Abs($delta)) $metric.Size) } else { & $format 0 $metric.Size }
                $color = if ($delta -gt 0) { '#1d4ed8' } elseif ($delta -lt 0) { '#b91c1c' } else { '#334155' }
                $cell = '{0}<br /><strong style="color:{1};">{2}</strong>' -f (ConvertTo-SmartM365EmailHtmlText (& $format $value $metric.Size)),$color,(ConvertTo-SmartM365EmailHtmlText $deltaText)
            }
            '<td align="right" style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;color:#334155;">{0}</td>' -f $cell
        }
        '<tr><td style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;color:#334155;">{0}</td><td align="right" style="border-bottom:1px solid #eef2f7;padding:8px 6px;font-size:11px;font-weight:700;">{1}</td>{2}</tr>' -f (ConvertTo-SmartM365EmailHtmlText $metric.Label),$currentHtml,($cells -join '')
    }
    return [pscustomobject]@{ Title = 'Raw inventory diff against previous scans'; Html = @"
<div style="font-size:12px;color:#64748b;margin-bottom:8px;">Each delta compares Current with a complete dated scan. n/a means no comparable scan on that date or unavailable local mailbox size statistics. Colors show direction, not health.</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr><th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">Metric</th><th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:8px 6px;font-size:11px;color:#475569;">Current ($($Current.SnapshotDate))</th>$($headers -join "`n")</tr>
  $($bodyRows -join "`n")
</table>
"@ }
}

function New-SmartM365LocalMailboxIssueEmailSection {
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Issues, [string]$IssuesCsvPath)

    if (@($Issues).Count -eq 0) { return }
    $rowsHtml = foreach ($group in @($Issues | Group-Object Category,Operation | Sort-Object Name)) {
        $nativeIds = @($group.Group | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ObjectGuid) } | Select-Object -ExpandProperty ObjectGuid -Unique)
        $unbound = @($group.Group | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.ObjectGuid) }).Count
        $objects = if ($unbound -gt 0) { '{0} identified; {1} unbound occurrences' -f $nativeIds.Count,$unbound } else { [string]$nativeIds.Count }
        $blocking = @($group.Group | Where-Object { [string]$_.BlocksCmdbQualification -eq 'True' }).Count
        @"
<tr>
  <td style="padding:8px;border-bottom:1px solid #fde68a;">$(ConvertTo-SmartM365EmailHtmlText $group.Group[0].Category)</td>
  <td style="padding:8px;border-bottom:1px solid #fde68a;">$(ConvertTo-SmartM365EmailHtmlText $group.Group[0].Operation)</td>
  <td style="padding:8px;border-bottom:1px solid #fde68a;">$(ConvertTo-SmartM365EmailHtmlText $objects)</td>
  <td align="right" style="padding:8px;border-bottom:1px solid #fde68a;">$($group.Count)</td>
  <td align="right" style="padding:8px;border-bottom:1px solid #fde68a;">$blocking</td>
</tr>
"@
    }
    [pscustomobject]@{ Title='Mailbox acquisition and recipient findings'; Html=@"
<p>Recipient findings do not mean objects were omitted. Unbound warnings and acquisition failures still block CMDB qualification.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#fffbeb;color:#78350f;font-size:12px;">
<tr style="background:#fef3c7;"><th align="left" style="padding:8px;">Finding</th><th align="left" style="padding:8px;">Operation</th><th align="left" style="padding:8px;">Affected objects</th><th align="right" style="padding:8px;">Occurrences</th><th align="right" style="padding:8px;">Blocking occurrences</th></tr>
$($rowsHtml -join "`r`n")
</table>
<p>Detailed evidence: $(ConvertTo-SmartM365EmailHtmlText $IssuesCsvPath)</p>
"@ }
}

function New-SmartM365ExchangeLocalMailboxReportEmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ReportRows,

        [Parameter(Mandatory = $true)]
        [string]$DailyStatsCsv,

        [string]$LatestDailyStatsCsv,

        [Parameter(Mandatory = $true)]
        [string]$SummaryCsv,

        [string]$LocalMailboxCsv,
        [string]$LatestLocalMailboxCsv,
        [string]$RemoteMailboxCsv,
        [string]$LatestRemoteMailboxCsv,

        [AllowNull()]$DailyStatsUpload,
        [AllowNull()]$SummaryUpload,
        [AllowNull()]$LocalMailboxUpload,
        [AllowNull()]$RemoteMailboxUpload,

        [object[]]$RemoteMailboxDataQualityWarnings = @(),
        [object[]]$LocalMailboxCollectionIssues = @(),
        [string]$CollectionIssuesCsvPath = '',
        [AllowNull()]$ProxyEvidence,
        [AllowNull()]$DailyDiffSection,

        [AllowNull()]$LocalMailboxWithDatabaseCount,

        [bool]$ShowMailLinks = $false,

        [Parameter(Mandatory = $true)]
        [string]$Title,
        [string]$TenantName=$Tenant,
        [string]$ReportHost=$env:COMPUTERNAME,
        [string]$ReportGeneratedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'),
        [System.Collections.IDictionary]$FileRowCounts
    )

    function Format-SmartM365ExchangeMailboxReportNumber {
        param([AllowNull()]$Value)
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '0' }
        try {
            $number = [double]$Value
            if ([math]::Abs($number % 1) -lt 0.000001) { return $number.ToString('N0') }
            return $number.ToString('N2')
        }
        catch {
            return [string]$Value
        }
    }

    function New-SmartM365ExchangeMailboxReportLinkHtml {
        param(
            [string]$Text,
            [string]$Url
        )

        $safeText = ConvertTo-SmartM365EmailHtmlText $Text
        if (-not $ShowMailLinks -or [string]::IsNullOrWhiteSpace($Url)) { return $safeText }
        $safeUrl = ConvertTo-SmartM365EmailHtmlText $Url
        return ('<a href="{0}" style="color:#075985;text-decoration:underline;">{1}</a>' -f $safeUrl, $safeText)
    }

    $rows = @($ReportRows)
    $domainCount = $rows.Count
    $totalMailboxes = ($rows | Measure-Object -Property TotalMailboxCount -Sum).Sum
    $totalLocalMailboxes = ($rows | Measure-Object -Property TotalLocalMailboxCount -Sum).Sum
    $totalRemoteMailboxes = ($rows | Measure-Object -Property TotalRemoteMailboxCount -Sum).Sum
    $totalEnabledAccounts = ($rows | Measure-Object -Property EnabledAccounts -Sum).Sum
    $totalDisabledAccounts = ($rows | Measure-Object -Property DisabledAccounts -Sum).Sum
    $totalSizeGb = ($rows | Measure-Object -Property TotalLocalMailboxSizeGB -Sum).Sum
    if ($null -eq $totalMailboxes) { $totalMailboxes = 0 }
    if ($null -eq $totalLocalMailboxes) { $totalLocalMailboxes = 0 }
    if ($null -eq $totalRemoteMailboxes) { $totalRemoteMailboxes = 0 }
    if ($null -eq $totalEnabledAccounts) { $totalEnabledAccounts = 0 }
    if ($null -eq $totalDisabledAccounts) { $totalDisabledAccounts = 0 }
    if ($null -eq $totalSizeGb) { $totalSizeGb = 0 }
    $totalSizeGb = [math]::Round([double]$totalSizeGb, 2)
    $qualifiedLocalValue = if ($null -eq $LocalMailboxWithDatabaseCount) { 'N/A' } else { Format-SmartM365ExchangeMailboxReportNumber -Value $LocalMailboxWithDatabaseCount }

    $summaryCards = @(
        [pscustomobject]@{ Label = 'Domains'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $domainCount; Detail = 'domains inventoried'; Background = '#f8fafc'; Border = '#dbe3ef'; Accent = '#0f172a'; Span = 1 }
        [pscustomobject]@{ Label = 'Inventory objects'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalMailboxes; Detail = 'raw local + remote'; Background = '#eff6ff'; Border = '#bfdbfe'; Accent = '#1d4ed8'; Span = 1 }
        [pscustomobject]@{ Label = 'Local mailboxes'; Value = $qualifiedLocalValue; Detail = 'distinct, with database'; Background = '#f0fdf4'; Border = '#bbf7d0'; Accent = '#166534'; Span = 1 }
        [pscustomobject]@{ Label = 'Remote mailboxes'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalRemoteMailboxes; Detail = 'remote recipient objects'; Background = '#faf5ff'; Border = '#e9d5ff'; Accent = '#7e22ce'; Span = 1 }
        [pscustomobject]@{ Label = 'Enabled accounts'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalEnabledAccounts; Detail = 'accounts enabled'; Background = '#f0fdfa'; Border = '#99f6e4'; Accent = '#0f766e'; Span = 1 }
        [pscustomobject]@{ Label = 'Disabled accounts'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalDisabledAccounts; Detail = 'accounts disabled'; Background = '#fff7ed'; Border = '#fed7aa'; Accent = '#9a3412'; Span = 1 }
        [pscustomobject]@{ Label = 'Local mailbox size'; Value = Format-SmartM365ExchangeMailboxReportNumber -Value $totalSizeGb; Detail = 'GB'; Background = '#eef2ff'; Border = '#c7d2fe'; Accent = '#4338ca'; Span = 2 }
    )

    $summaryCardsHtml = New-SmartM365EmailKpiGridHtml -Cards $summaryCards -RowSizes @(4,3)

    $domainRowsHtml = foreach ($row in ($rows | Sort-Object -Property @{ Expression = 'TotalMailboxCount'; Descending = $true }, DomainName)) {
        @"
<tr>
  <td style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText $row.DomainName)</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalLocalMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalRemoteMailboxCount))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.EnabledAccounts))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.DisabledAccounts))</td>
  <td align="right" style="border-bottom:1px solid #eef2f7;padding:9px 10px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $row.TotalLocalMailboxSizeGB))</td>
</tr>
"@
    }

    $domainTableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Domain</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Total</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Local</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Remote</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Enabled</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Disabled</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;color:#475569;text-transform:uppercase;">Local GB</th>
  </tr>
  $($domainRowsHtml -join "`r`n")
</table>
"@

    $fileRows = @()
    $localMailboxPathForMail = if (-not [string]::IsNullOrWhiteSpace($LatestLocalMailboxCsv)) { $LatestLocalMailboxCsv } else { $LocalMailboxCsv }
    if (-not [string]::IsNullOrWhiteSpace($localMailboxPathForMail)) {
        $fileRows += [pscustomobject]@{ Label = 'Mailbox inventory'; Path = $localMailboxPathForMail; WebUrl = if ($LocalMailboxUpload -and $LocalMailboxUpload.WebUrl) { [string]$LocalMailboxUpload.WebUrl } else { '' } }
    }
    $remoteMailboxPathForMail = if (-not [string]::IsNullOrWhiteSpace($LatestRemoteMailboxCsv)) { $LatestRemoteMailboxCsv } else { $RemoteMailboxCsv }
    if ($remoteMailboxPathForMail -and $(if ($null -ne $FileRowCounts) { $FileRowCounts.Contains($remoteMailboxPathForMail) } else { Test-Path -LiteralPath $remoteMailboxPathForMail -PathType Leaf })) {
        $fileRows += [pscustomobject]@{ Label = 'Remote mailbox inventory'; Path = $remoteMailboxPathForMail; WebUrl = if ($RemoteMailboxUpload -and $RemoteMailboxUpload.WebUrl) { [string]$RemoteMailboxUpload.WebUrl } else { '' } }
    }
    if (-not [string]::IsNullOrWhiteSpace($LatestDailyStatsCsv)) {
        $fileRows += [pscustomobject]@{ Label = 'Daily stats'; Path = $LatestDailyStatsCsv; WebUrl = if ($DailyStatsUpload -and $DailyStatsUpload.WebUrl) { [string]$DailyStatsUpload.WebUrl } else { '' } }
    }
    else {
        $fileRows += [pscustomobject]@{ Label = 'Daily stats'; Path = $DailyStatsCsv; WebUrl = if ($DailyStatsUpload -and $DailyStatsUpload.WebUrl) { [string]$DailyStatsUpload.WebUrl } else { '' } }
    }
    $fileRows += [pscustomobject]@{ Label = 'Summary'; Path = $SummaryCsv; WebUrl = if ($SummaryUpload -and $SummaryUpload.WebUrl) { [string]$SummaryUpload.WebUrl } else { '' } }
    $fileRowsHtml = foreach ($fileRow in $fileRows) {
        $rowCount = if ($null -ne $FileRowCounts) { $FileRowCounts[$fileRow.Path] } else { Get-SmartM365MailboxReportCsvRowCount -Path $fileRow.Path }
        $pathHtml = New-SmartM365ExchangeMailboxReportLinkHtml -Text $fileRow.Path -Url $fileRow.WebUrl
        @"
<tr>
  <td style="width:130px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText $fileRow.Label)</td>
  <td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;">$pathHtml</td>
  <td align="right" style="width:80px;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(ConvertTo-SmartM365EmailHtmlText (Format-SmartM365ExchangeMailboxReportNumber -Value $rowCount))</td>
</tr>
"@
    }

    $filesTableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">File</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Path / SharePoint link</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Rows</th>
  </tr>
  $($fileRowsHtml -join "`r`n")
</table>
"@

    $sections = @(
        [pscustomobject]@{ Title = 'Summary'; Html = $summaryCardsHtml }
    )
    if ($DailyDiffSection) { $sections += $DailyDiffSection }
    $sections += [pscustomobject]@{ Title = 'Domain summary (raw inventory objects)'; Html = $domainTableHtml }
    $localIssueSection = New-SmartM365LocalMailboxIssueEmailSection -Issues $LocalMailboxCollectionIssues -IssuesCsvPath $CollectionIssuesCsvPath
    if ($localIssueSection) { $sections += $localIssueSection }

    $exchangeDataQualityWarnings = @($RemoteMailboxDataQualityWarnings | Select-Object -First 100)
    if ($exchangeDataQualityWarnings.Count -gt 0) {
        $warningRowsHtml = foreach ($warningRow in $exchangeDataQualityWarnings) {
            $issueLabel = switch ([string]$warningRow.Issue) {
                'ExchangeObjectInconsistentState' { 'Inconsistent state'; break }
                'MissingPrimarySmtpAddress' { 'Missing primary SMTP'; break }
                'MissingExternalEmailAddress' { 'Missing external address'; break }
                'InvalidExternalEmailAddress' { 'Invalid external address'; break }
                'InvalidDisplayName' { 'Invalid display name'; break }
                default { [string]$warningRow.Issue }
            }
            @"
<tr>
  <td valign="top" style="width:165px;border-bottom:1px solid #fde68a;padding:10px 12px;font-size:12px;color:#92400e;font-weight:700;line-height:1.35;">$(ConvertTo-SmartM365EmailHtmlText $issueLabel)</td>
  <td valign="top" style="border-bottom:1px solid #fde68a;padding:10px 12px;font-size:12px;color:#78350f;line-height:1.4;">
    <div style="font-weight:700;margin:0 0 6px 0;color:#78350f;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.Warning)</div>
    <div style="margin:0 0 6px 0;font-family:Consolas,'Courier New',monospace;font-size:11px;color:#92400e;word-break:break-word;overflow-wrap:anywhere;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.ObjectPath)</div>
    <div style="margin:0;color:#92400e;">$(ConvertTo-SmartM365EmailHtmlText $warningRow.SuggestedAction)</div>
  </td>
</tr>
"@
        }
        $issueCountRowsHtml = foreach ($issueGroup in @($RemoteMailboxDataQualityWarnings | Group-Object Issue | Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false })) {
            $issueLabel = switch ([string]$issueGroup.Name) {
                'ExchangeObjectInconsistentState' { 'Inconsistent state'; break }
                'MissingPrimarySmtpAddress' { 'Missing primary SMTP'; break }
                'MissingExternalEmailAddress' { 'Missing external address'; break }
                'InvalidExternalEmailAddress' { 'Invalid external address'; break }
                'InvalidDisplayName' { 'Invalid display name'; break }
                default { [string]$issueGroup.Name }
            }
            @"
<tr>
  <td style="border-bottom:1px solid #fde68a;padding:8px 10px;font-size:12px;color:#78350f;font-weight:700;">$(ConvertTo-SmartM365EmailHtmlText $issueLabel)</td>
  <td align="right" style="border-bottom:1px solid #fde68a;padding:8px 10px;font-size:12px;color:#78350f;font-weight:700;">$(ConvertTo-SmartM365EmailHtmlText $issueGroup.Count)</td>
</tr>
"@
        }
        $warningsTableHtml = @"
<p style="margin:0 0 10px 0;color:#78350f;font-size:13px;">Exchange object data quality warnings captured during Get-RemoteMailbox.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;max-width:420px;border-collapse:collapse;border:1px solid #f59e0b;font-size:12px;background:#fffbeb;margin:0 0 14px 0;">
  <tr>
    <th align="left" style="background:#fef3c7;border-bottom:1px solid #f59e0b;padding:9px 10px;color:#92400e;text-transform:uppercase;">Issue</th>
    <th align="right" style="width:90px;background:#fef3c7;border-bottom:1px solid #f59e0b;padding:9px 10px;color:#92400e;text-transform:uppercase;">Count</th>
  </tr>
  $($issueCountRowsHtml -join "`r`n")
</table>
<p style="margin:0 0 10px 0;color:#78350f;font-size:13px;">Top 100 detailed warnings.</p>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #f59e0b;font-size:12px;background:#fffbeb;table-layout:fixed;">
  <tr>
    <th align="left" style="width:165px;background:#fef3c7;border-bottom:1px solid #f59e0b;padding:10px 12px;color:#92400e;text-transform:uppercase;">Issue</th>
    <th align="left" style="background:#fef3c7;border-bottom:1px solid #f59e0b;padding:10px 12px;color:#92400e;text-transform:uppercase;">Details</th>
  </tr>
  $($warningRowsHtml -join "`r`n")
</table>
"@
        $sections += [pscustomobject]@{ Title = 'Exchange object data quality warnings - Top 100'; Html = $warningsTableHtml }
    }

    if ($null -ne $ProxyEvidence) {
        if (Get-Command New-SmartM365ProxyMailSections -ErrorAction SilentlyContinue) {
            $sections += @(New-SmartM365ProxyMailSections -Evidence $ProxyEvidence -ShowMailLinks:$ShowMailLinks)
        }
        else {
            $sections += [pscustomobject]@{ Title = 'Proxy address audit'; Html = '<p>ProxyAddresses report data is unavailable.</p>' }
        }
    }
    $sections += [pscustomobject]@{ Title = 'Files'; Html = $filesTableHtml }

    return New-SmartM365EmailBody `
        -Title 'Mailbox inventory summary' `
        -Category 'SmartM365 Exchange OnPrem' `
        -Severity $(if ($ProxyEvidence -and (-not $ProxyEvidence.Available -or @($ProxyEvidence.Summary | Where-Object { $_.Summary -eq 'With expected address missing' -and [int]$_.Count -gt 0 }).Count -gt 0)) { 'Warning' } else { 'Success' }) `
        -Tenant $TenantName -HostName $ReportHost -GeneratedAt $ReportGeneratedAt `
        -Message 'Exchange on-premises mailbox inventory summary generated from the latest SmartM365 CSV outputs.' `
        -Sections $sections `
        -Footer 'This automated message was generated by SmartM365 from Exchange on-premises mailbox inventory data.'
}

# SmartM365-WinUpdate_Status_From_Intune

function Html-Encode {
    param([Parameter(Mandatory=$false)][object]$Value)
    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $s = $s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace('"',"&quot;").Replace("'","&#39;")
    return $s
}

function Get-SmartM365QuarantineReportColumns {
    'TenantKey','RunId','CollectedAtUtc','ReceivedTimeUtc','AgeDays','SenderAddress','RecipientAddress','Subject',
    'QuarantineType','PolicyName','PolicyType','Direction','ReleaseStatus','ExpiresUtc','Reported','SystemReleased','MessageId','Identity'
}

function Get-SmartM365QuarantineReportSubject {
    param([int]$Count,[string]$LocalDate=(Get-Date -Format 'yyyy-MM-dd'))
    '[SmartM365] Quarantaine EXO | {0} message(s) bloque(s) | {1}' -f $Count,$LocalDate
}

function Add-SmartM365FreeBusyTemplateMarker {
    param([string]$Html)
    if ($Html -notmatch 'SmartM365EmailTemplate:v1') { return "<!-- SmartM365EmailTemplate:v1 -->`r`n$Html" }
    $Html
}

function New-SmartM365AdInventoryMailReport {
    param($SummarySnapshot,$PreviousSnapshot,$WeekSnapshot,$MonthSnapshot,[string]$LogicalCountBaselineAt,
        [string]$SourceFolder,[string]$SummaryCsvPath,[string]$LatestSummaryPath,[object[]]$SummarySharePointUploads,
        [bool]$ShowMailLinks,$DomainCounts,[string[]]$DomainNames,[string]$Tenant,[string]$ReportHost,
        [string]$ReportGeneratedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'),[AllowNull()]$IdentityContext)
    $diffSection = New-SmartM365AdDailySummaryDiffSection -Current $summarySnapshot -Previous $previousSnapshot -Week $weekSnapshot -Month $monthSnapshot -LogicalCountBaselineAt $logicalCountBaselineAt
    $allUploads = @($summarySharePointUploads)
    if ($IdentityContext) { $allUploads += @($IdentityContext.SharePointUploads) }
    $sharePointSection = if ($ShowMailLinks) { New-SmartM365SharePointLinksSection -UploadRecords $allUploads } else { $null }
    $sections = @($diffSection)
    if ($IdentityContext) { $sections += @($IdentityContext.Sections) }
    if ($sharePointSection) { $sections += $sharePointSection }

    $hasDuplicateIdentities = ((ConvertTo-SmartM365AdSummaryInt64 -Value $summarySnapshot.AffectedDuplicateUPNAccounts) -gt 0) -or ((ConvertTo-SmartM365AdSummaryInt64 -Value $summarySnapshot.AffectedDuplicateSMTPEntries) -gt 0)
    $severity = if ($hasDuplicateIdentities) { 'Warning' } else { 'Success' }
    $actionTitle = if ($hasDuplicateIdentities) { 'Review required' } else { 'No duplicate identity conflict detected' }
    $actionHtml = if ($hasDuplicateIdentities) { 'Review duplicate UPN and SMTP counters before identity cleanup, migration, or synchronization decisions.' } else { 'Keep the generated CSV files as the daily Active Directory inventory baseline.' }

    $summaryHtml = (New-SmartM365AdDailySummaryKpiHtml -Snapshot $summarySnapshot -Previous $previousSnapshot -LogicalCountBaselineAt $logicalCountBaselineAt) + (New-SmartM365AdDomainSummaryHtml -Snapshot $summarySnapshot -DomainCounts $DomainCounts -DomainNames $DomainNames)
    if ($IdentityContext) {
        $summaryHtml += '<div style="font-size:13px;font-weight:700;margin:16px 0 8px;">Identity and mail routing issues by domain</div>' + [string]$IdentityContext.SummaryHtml
        $severity = 'Warning'
        $actionTitle = 'Action required'
        $actionHtml = 'Review the attached Excel workbook and the identity diagnostics before cleanup, migration, or synchronization decisions.'
    }

    $pathRows = @(
        [pscustomobject]@{ Label = 'Source folder'; Path = $SourceFolder }
        [pscustomobject]@{ Label = 'Daily summary'; Path = $summaryCsvPath }
    )
    if (-not [string]::IsNullOrWhiteSpace($latestSummaryPath)) {
        $pathRows += [pscustomobject]@{ Label = 'Latest summary'; Path = $latestSummaryPath }
    }

    if ($IdentityContext -and $ShowMailLinks) { $pathRows += @($IdentityContext.PathRows) }
    if (-not $ShowMailLinks) { $pathRows = @() }
    $mailTitle = if ($IdentityContext) { 'Active Directory daily summary and identity diagnostics' } else { 'Active Directory daily summary' }
    $mailSubject = 'SmartM365 ' + $mailTitle
    $mailFooter = if ($ShowMailLinks) { 'This automated message was generated by SmartM365. Use the exported CSV paths and SharePoint links as the inventory source of truth.' } elseif ($IdentityContext) { 'This automated message was generated by SmartM365. Use the exported CSV files and attached workbook as the inventory source of truth.' } else { 'This automated message was generated by SmartM365. Use the exported CSV files as the inventory source of truth.' }

    $emailBody = New-SmartM365EmailBody `
        -Title $mailTitle `
        -Category 'SmartM365 Active Directory' `
        -Severity $severity `
        -Tenant $Tenant `
        -HostName $ReportHost -GeneratedAt $ReportGeneratedAt `
        -Message 'Daily Active Directory inventory summary with previous, seven-day, and thirty-day comparisons.' `
        -ActionTitle $actionTitle `
        -ActionHtml $actionHtml `
        -SummaryHtml $summaryHtml `
        -PathRows $pathRows `
        -Sections $sections `
        -Footer $mailFooter

    $attachments = if ($IdentityContext) { @($IdentityContext.WorkbookPath) } else { @() }
    [pscustomobject]@{Subject=$mailSubject; BodyHtml=$emailBody; Attachments=@($attachments); AllowAttachments=$true; SuppressAttachmentLinks=(-not $ShowMailLinks)}
}

function New-SmartM365PreparedCompletionMailBody {
    param([System.Collections.IDictionary]$Details,[AllowEmptyCollection()][string[]]$Warnings)
    $rows=foreach($label in $details.Keys) {
        '<tr><th>{0}</th><td>{1}</td></tr>' -f [Net.WebUtility]::HtmlEncode($label),[Net.WebUtility]::HtmlEncode([string]$details[$label])
    }
    $maxListedWarnings=50
    $warningHtml=''
    if ($warnings.Count) {
        $items=@($warnings | Select-Object -First $maxListedWarnings | ForEach-Object { '<li>{0}</li>' -f [Net.WebUtility]::HtmlEncode($_) })
        $more=if($warnings.Count -gt $maxListedWarnings){'<p>{0} more warning(s) in the log.</p>' -f ($warnings.Count-$maxListedWarnings)}else{''}
        $warningHtml='<h3>Warnings ({0})</h3><ul>{1}</ul>{2}' -f $warnings.Count,($items -join ''),$more
    }
    $warningNote=if($warnings.Count){'The warnings above were logged before this notification.'}else{'No warning was logged before this notification.'}
    $body='<h2>Workplace evidence completed</h2><table>'+($rows -join '')+'</table>'+$warningHtml+'<p>'+$warningNote+' Batch retention runs after this notification; its result and any warning are in the same log. Power BI refresh and publication are separate operations.</p>'
    return $body
}

function Convert-ObjectsToHtmlTable {
    param(
        [Parameter(Mandatory=$true)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory=$true)][string[]]$Columns,
        [Parameter(Mandatory=$true)][string]$Title
    )

    if ($null -eq $Rows) { $Rows = @() }
    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<h3 style='margin:16px 0 8px 0;'>$(Html-Encode $Title)</h3><div style='font-family:Segoe UI,Arial;font-size:12px;color:#666;'>No data.</div>"
    }

    $th  = ($Columns | ForEach-Object { "<th style='padding:6px;border:1px solid #ddd;background:#f5f5f5;text-align:left;'>$(Html-Encode $_)</th>" }) -join ""
    $trs = ($Rows | ForEach-Object {
        $row = $_
        $tds = ($Columns | ForEach-Object { "<td style='padding:6px;border:1px solid #ddd;'>$(Html-Encode $row.$_)</td>" }) -join ""
        "<tr>$tds</tr>"
    }) -join "`n"

    return @"
<h3 style="margin:16px 0 8px 0;">$(Html-Encode $Title)</h3>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial;font-size:12px;">
<tr>$th</tr>
$trs
</table>
"@
}

function Get-WinUpdateCombinedVersionDistribution {
    param([Parameter(Mandatory)][object]$CombinedSummary)

    $windows11 = [int]$CombinedSummary.Windows11
    $windows10 = [int]$CombinedSummary.Windows10
    $unknown = [int]$CombinedSummary.UnknownOrOther
    $total = [int]$CombinedSummary.Total
    if ($total -le 0 -or $windows11 -lt 0 -or $windows10 -lt 0 -or $unknown -lt 0 -or
        ($windows11 + $windows10 + $unknown) -ne $total) {
        throw 'Combined Windows version distribution does not reconcile.'
    }
    [pscustomobject]@{
        Windows11 = $windows11
        Windows10 = $windows10
        UnknownOrOther = $unknown
        Total = $total
        Windows11Pct = [math]::Round(100.0 * $windows11 / $total, 2)
        Windows10Pct = [math]::Round(100.0 * $windows10 / $total, 2)
        UnknownOrOtherPct = [math]::Round(100.0 * $unknown / $total, 2)
    }
}

function New-SmartM365WindowsUpdateMailReport {
    param(
        $actionRequiredCount, $adSourceText, $adWithoutIntuneSummary, $allPolicySummary, $combinedFleetSummary,
        $completedCount, $completionPct, $computedRowCount, $count, $coverageAvailable,
        $CsvLastFinal, $durationSeconds, $EnableReadinessEnrichment, $fleetDevices, $fleetKnownOsCoveragePct,
        $fleetOsBelowTarget, $fleetOsCoveragePct, $fleetOsCovered, $fleetOsVersionUnknown, $fleetPolicyCompletionPct,
        $fleetPolicyName, $fleetTargetLabel, $fleetTargetShortLabel, $inProgressCount, $installingCount,
        $intuneFleetSummary, $intuneSourceText, $offeringCount, $OrgDomain, $otherInProgressCount,
        $pendingCount, $policyCount, $primaryCoveragePolicy, $priority0Count, $readinessMode,
        $reportHealth, $reportStatus, $RunId, $spUploadUrl, $statusAgeCoveragePct,
        $statusAgePopulatedRows, $statusAgeSourceColumn, $totalUniqueDevices, $unknownCount, $actionRows,
        $actionLimit, $summaryMatched, $summaryNotMatched, $summaryEntraOsVersion, $summaryUnavailableOsVersion,
        $summaryReadinessPct
    )
    # Rows are already canonical per policy/device, so a second device deduplication is unnecessary here.
    $policySummary = $allPolicySummary
    $policyStrips = foreach ($policy in $policySummary) {
        $policyStatusColor = switch ($policy.Status) { 'CRITICAL' {'#b91c1c'} 'WARNING' {'#b45309'} default {'#15803d'} }
        $policyStatusBackground = switch ($policy.Status) { 'CRITICAL' {'#fee2e2'} 'WARNING' {'#fef3c7'} default {'#dcfce7'} }
        $encodedPolicyName = Html-Encode $policy.PolicyName
        $encodedPolicyId = Html-Encode $policy.PolicyId
        if ($null -ne $policy.TargetBuild) {
            $encodedTargetLabel = Html-Encode $policy.TargetLabel
            $policyTargetHtml = "OS target: $encodedTargetLabel (build $($policy.TargetBuild)+)"
            $policyOsCoveredText = [string]$policy.OsCovered
            $policyOsBelowText = [string]$policy.OsBelowTarget
            $policyOsUnknownText = [string]$policy.OsVersionUnknown
            $policyOsCoverageText = "$($policy.OsCoveragePct)%"
        } else {
            $policyTargetHtml = 'OS target: not detected from policy name'
            $policyOsCoveredText = 'N/A'
            $policyOsBelowText = 'N/A'
            $policyOsUnknownText = 'N/A'
            $policyOsCoverageText = 'N/A'
        }
        @"
<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 14px 0;border:1px solid #dbe4ee;background:#ffffff;">
<tr><td colspan="7" style="padding:12px 14px;border-bottom:1px solid #dbe4ee;background:#f8fafc;">
<span style="display:inline-block;margin-right:8px;padding:4px 9px;border-radius:999px;background:$policyStatusBackground;color:$policyStatusColor;font-size:11px;font-weight:700;">$($policy.Status)</span>
<strong style="font-size:14px;color:#0f172a;">$encodedPolicyName</strong>
<div style="margin-top:4px;font-size:11px;color:#64748b;">Policy ID: $encodedPolicyId | $policyTargetHtml</div>
</td></tr>
<tr>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;"><div style="font-size:10px;color:#64748b;">DEVICES</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.Devices)</div></td>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#dcfce7;"><div style="font-size:10px;color:#166534;">OS COVERED</div><div style="font-size:20px;font-weight:700;color:#334155;">$policyOsCoveredText</div></td>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#fef3c7;"><div style="font-size:10px;color:#92400e;">OS BELOW TARGET</div><div style="font-size:20px;font-weight:700;color:#334155;">$policyOsBelowText</div></td>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#fffbeb;"><div style="font-size:10px;color:#92400e;">OS UNKNOWN</div><div style="font-size:20px;font-weight:700;color:#334155;">$policyOsUnknownText</div></td>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#f0fdf4;"><div style="font-size:10px;color:#166534;">OS COVERAGE</div><div style="font-size:20px;font-weight:700;color:#334155;">$policyOsCoverageText</div></td>
<td style="width:14.28%;padding:11px 9px;border-right:1px solid #e2e8f0;"><div style="font-size:10px;color:#64748b;">INTUNE COMPLETED</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.Completed)</div></td>
<td style="width:14.28%;padding:11px 9px;"><div style="font-size:10px;color:#64748b;">INTUNE COMPLETION</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.PolicyCompletionPct)%</div></td>
</tr>
<tr>
<td style="padding:9px;background:#eff6ff;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#1d4ed8;">OFFERING</div><div style="font-size:17px;font-weight:700;">$($policy.Offering)</div></td>
<td style="padding:9px;background:#eff6ff;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#1d4ed8;">INSTALLING</div><div style="font-size:17px;font-weight:700;">$($policy.Installing)</div></td>
<td style="padding:9px;background:#eff6ff;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#1d4ed8;">PENDING</div><div style="font-size:17px;font-weight:700;">$($policy.Pending)</div></td>
<td style="padding:9px;background:#eff6ff;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#1d4ed8;">OTHER IN PROGRESS</div><div style="font-size:17px;font-weight:700;">$($policy.OtherInProgress)</div></td>
<td style="padding:9px;background:#eff6ff;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#1d4ed8;">IN PROGRESS TOTAL</div><div style="font-size:17px;font-weight:700;">$($policy.InProgress)</div></td>
<td style="padding:9px;background:#fef2f2;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#b91c1c;">ACTION REQUIRED</div><div style="font-size:17px;font-weight:700;">$($policy.ActionRequired)</div></td>
<td style="padding:9px;background:#fffbeb;border-top:1px solid #e2e8f0;"><div style="font-size:10px;color:#92400e;">UNKNOWN</div><div style="font-size:17px;font-weight:700;">$($policy.Unknown)</div></td>
</tr></table>
"@
    }
    $policyStripsHtml = $policyStrips -join "`n"

    $actionTable = Convert-ObjectsToHtmlTable -Rows $actionRows -Columns @('ActionPriority','DeviceName','PolicyName','AggregateState','CurrentDeviceUpdateStatus','LatestAlertMessage','UpgradeEligibilityLabel','ActionCode','ActionOwner','ActionDescription') -Title ("Top {0} devices requiring action" -f $actionLimit)

    $readinessRows = @(
        [pscustomobject]@{ Metric='Enrichment enabled'; Value=$EnableReadinessEnrichment }
        [pscustomobject]@{ Metric='Mode'; Value=$readinessMode }
        [pscustomobject]@{ Metric='Matched devices'; Value=$summaryMatched }
        [pscustomobject]@{ Metric='Unmatched devices'; Value=$summaryNotMatched }
        [pscustomobject]@{ Metric='Match percentage'; Value="$summaryReadinessPct%" }
        [pscustomobject]@{ Metric='OS version from Entra fallback'; Value=$summaryEntraOsVersion }
        [pscustomobject]@{ Metric='OS version unavailable after fallback'; Value=$summaryUnavailableOsVersion }
        [pscustomobject]@{ Metric='Last-status date source'; Value=$statusAgeSourceColumn }
        [pscustomobject]@{ Metric='Rows with DaysSinceLastStatus'; Value="$statusAgePopulatedRows/$computedRowCount" }
        [pscustomobject]@{ Metric='DaysSinceLastStatus coverage'; Value="$statusAgeCoveragePct%" }
    )
    $readinessTable = Convert-ObjectsToHtmlTable -Rows $readinessRows -Columns @('Metric','Value') -Title 'Readiness data quality'

    $statusColor = switch ($reportStatus) { 'CRITICAL' {'#b91c1c'} 'WARNING' {'#b45309'} default {'#15803d'} }
    $statusBackground = switch ($reportStatus) { 'CRITICAL' {'#fee2e2'} 'WARNING' {'#fef3c7'} default {'#dcfce7'} }
    $severityCoverageText = if ($coverageAvailable) { "$fleetKnownOsCoveragePct%" } else { 'N/A' }
    $severityUnknownText = if ($coverageAvailable) { "$($reportHealth.OsVersionUnknownRatePct)%" } else { 'N/A' }
    $severityInputsHtml = "<p style='margin:0 0 18px 0;font-size:12px;color:#64748b;'>Severity inputs: actions $($reportHealth.ActionRequiredRatePct)% | P0 blockers $priority0Count/$($reportHealth.Priority0Threshold) critical threshold | known-OS coverage $severityCoverageText | OS unknown $severityUnknownText.</p>"
    if ($coverageAvailable) {
        $subject = "SMART365 - [$reportStatus] WinUpdate Feature Update - $fleetOsCovered/$fleetDevices devices on $fleetTargetShortLabel+ - $actionRequiredCount action(s)"
        $fleetCoverageSection = @"
<h2 style="margin:0 0 6px 0;font-size:20px;color:#0f172a;">Fleet OS coverage - $(Html-Encode $fleetTargetLabel)+</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Reference policy: $(Html-Encode $fleetPolicyName) | Minimum build: $($primaryCoveragePolicy.TargetBuild) | OS coverage uses the current OSVersion, independently of the Intune policy workflow state.</p>
<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;"><tr>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">TARGETED DEVICES</div><div style="font-size:24px;font-weight:700;">$fleetDevices</div></td>
<td style="padding:14px;background:#dcfce7;border:1px solid #bbf7d0;"><div style="font-size:11px;color:#166534;">OS COVERED</div><div style="font-size:24px;font-weight:700;">$fleetOsCovered</div></td>
<td style="padding:14px;background:#fef3c7;border:1px solid #fde68a;"><div style="font-size:11px;color:#92400e;">BELOW TARGET</div><div style="font-size:24px;font-weight:700;">$fleetOsBelowTarget</div></td>
<td style="padding:14px;background:#fffbeb;border:1px solid #fde68a;"><div style="font-size:11px;color:#92400e;">OS UNKNOWN</div><div style="font-size:24px;font-weight:700;">$fleetOsVersionUnknown</div></td>
<td style="padding:14px;background:#dcfce7;border:1px solid #bbf7d0;"><div style="font-size:11px;color:#166534;">OS COVERAGE</div><div style="font-size:24px;font-weight:700;">$fleetOsCoveragePct%</div></td>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">INTUNE COMPLETION</div><div style="font-size:24px;font-weight:700;">$fleetPolicyCompletionPct%</div></td>
</tr></table>
<p style="margin:6px 0 22px 0;font-size:12px;color:#64748b;">Coverage among devices with a known OS version: $fleetKnownOsCoveragePct%. Intune policy completion remains visible as a secondary workflow metric.</p>
"@
    } else {
        $subject = "SMART365 - [$reportStatus] WinUpdate Feature Update - All policies - OS coverage unavailable - $actionRequiredCount action(s)"
        $fleetCoverageSection = "<h2 style='margin:0 0 6px 0;font-size:20px;color:#0f172a;'>Fleet OS coverage unavailable</h2><p style='margin:0 0 22px 0;font-size:12px;color:#64748b;'>No supported Windows version was detected in a policy name. Policy workflow metrics remain available below.</p>"
    }
    if ($combinedFleetSummary) {
        $distribution = Get-WinUpdateCombinedVersionDistribution -CombinedSummary $combinedFleetSummary
        $invariantCulture = [Globalization.CultureInfo]::InvariantCulture
        $windows11Width = [string]::Format($invariantCulture,'{0:0.##}',$distribution.Windows11Pct)
        $windows10Width = [string]::Format($invariantCulture,'{0:0.##}',$distribution.Windows10Pct)
        $unknownOrOtherWidth = [string]::Format($invariantCulture,'{0:0.##}',$distribution.UnknownOrOtherPct)
        $windowsVersionDistributionSection = @"
<h2 style="margin:0 0 6px 0;font-size:20px;color:#0f172a;">Windows version distribution - enabled Windows devices</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Intune inventory and AD Windows computers without an exact Intune ID match | Enabled in Entra or AD | OS family from each source.</p>
<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 10px 0;"><tr>
<td width="$windows11Width%" style="width:$windows11Width%;height:22px;background:#2563eb;font-size:1px;line-height:22px;">&nbsp;</td>
<td width="$windows10Width%" style="width:$windows10Width%;height:22px;background:#f59e0b;font-size:1px;line-height:22px;">&nbsp;</td>
<td width="$unknownOrOtherWidth%" style="width:$unknownOrOtherWidth%;height:22px;background:#94a3b8;font-size:1px;line-height:22px;">&nbsp;</td>
</tr></table>
<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;"><tr>
<td style="width:33.33%;padding:12px;background:#eff6ff;border:1px solid #bfdbfe;"><div style="font-size:11px;color:#1d4ed8;">WINDOWS 11</div><div style="font-size:22px;font-weight:700;color:#0f172a;">$($distribution.Windows11)</div><div style="font-size:12px;color:#475569;">$($distribution.Windows11Pct)%</div></td>
<td style="width:33.33%;padding:12px;background:#fffbeb;border:1px solid #fde68a;"><div style="font-size:11px;color:#b45309;">WINDOWS 10</div><div style="font-size:22px;font-weight:700;color:#0f172a;">$($distribution.Windows10)</div><div style="font-size:12px;color:#475569;">$($distribution.Windows10Pct)%</div></td>
<td style="width:33.33%;padding:12px;background:#f8fafc;border:1px solid #cbd5e1;"><div style="font-size:11px;color:#64748b;">OS VERSION UNKNOWN</div><div style="font-size:22px;font-weight:700;color:#0f172a;">$($distribution.UnknownOrOther)</div><div style="font-size:12px;color:#475569;">$($distribution.UnknownOrOtherPct)%</div></td>
</tr></table>
<p style="margin:6px 0 22px 0;font-size:12px;color:#64748b;">Combined exact-ID scope: $($intuneFleetSummary.Total.Total) Intune + $($adWithoutIntuneSummary.All.Total) unmatched AD = $($distribution.Total). Fleet OS coverage below uses the reference policy.</p>
"@
    } else {
        $windowsVersionDistributionSection = "<h2 style='margin:0 0 6px 0;font-size:20px;color:#0f172a;'>Windows version distribution unavailable</h2><p style='margin:0 0 22px 0;font-size:12px;color:#64748b;'>The enabled Intune and unmatched AD sources could not be reconciled. No combined OS counts are inferred.</p>"
    }
    if ($intuneFleetSummary -and $adWithoutIntuneSummary -and $combinedFleetSummary) {
        $combinedCountryRowsHtml = foreach ($countryRow in @($combinedFleetSummary.Known) + @($combinedFleetSummary.Unknown)) {
            $countryName = if ($countryRow.Country) { Html-Encode $countryRow.Country } else { 'Country unknown' }
            $rowBackground = if ($countryRow.Country) { '#ffffff' } else { '#f8fafc' }
            "<tr style='background:$rowBackground;'><td style='padding:7px 8px;border-bottom:1px solid #e2e8f0;'>$countryName</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.Windows11)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.Windows10)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.UnknownOrOther)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;font-weight:600;'>$($countryRow.Total)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.Intune)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.AdWithoutIntune)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.ReferencePolicy)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.OtherPolicyOnly)</td><td style='padding:7px 8px;text-align:right;border-bottom:1px solid #e2e8f0;'>$($countryRow.NeitherPolicy)</td></tr>"
        }
        $combinedCountryRowsHtml = $combinedCountryRowsHtml -join [Environment]::NewLine
        $intuneCountryResolved = $intuneFleetSummary.Total.Total - $intuneFleetSummary.Unknown.Total
        $adCountryResolved = $adWithoutIntuneSummary.All.Total - $adWithoutIntuneSummary.Unknown.Total
        $referencePolicyLabel = if ($coverageAvailable) { Html-Encode $fleetPolicyName } else { 'Unavailable' }
        $combinedRowHtml = "<tr style='background:#dbeafe;font-weight:700;'><td style='padding:9px;'>COMBINED EXACT-ID SCOPE</td><td style='padding:9px;text-align:right;'>$($combinedFleetSummary.Windows11)</td><td style='padding:9px;text-align:right;'>$($combinedFleetSummary.Windows10)</td><td style='padding:9px;text-align:right;'>$($combinedFleetSummary.UnknownOrOther)</td><td style='padding:9px;text-align:right;'>$($combinedFleetSummary.Total)</td><td style='padding:9px;text-align:right;'>$($intuneFleetSummary.Total.Total)</td><td style='padding:9px;text-align:right;'>$($adWithoutIntuneSummary.All.Total)</td><td style='padding:9px;text-align:right;'>$($intuneFleetSummary.Total.ReferencePolicy)</td><td style='padding:9px;text-align:right;'>$($intuneFleetSummary.Total.OtherPolicyOnly)</td><td style='padding:9px;text-align:right;'>$($intuneFleetSummary.Total.NeitherPolicy)</td></tr>"
        $referenceReconciliationHtml = if ($coverageAvailable) {
            "<p style='margin:0 0 8px 0;font-size:12px;color:#64748b;'>Reference-policy report to combined scope: $fleetDevices - $($intuneFleetSummary.ExcludedReferencePolicy) excluded by activation - $($intuneFleetSummary.ReportOnlyReferenceIds) absent from the Intune snapshot + $($intuneFleetSummary.Total.OtherPolicyOnly) other-policy-only + $($intuneFleetSummary.Total.NeitherPolicy) neither-policy + $($adWithoutIntuneSummary.All.Total) unmatched AD = $($combinedFleetSummary.Total).</p>"
        } else { '' }
        $intuneFleetSection = @"
<h2 style="margin:0 0 6px 0;font-size:20px;color:#0f172a;">Enabled Windows devices by country - Intune and unmatched AD</h2>
<p style="margin:0 0 10px 0;font-size:12px;color:#64748b;">Included when Entra AccountEnabled or AD Enabled is true. Reference Feature Update policy: $referencePolicyLabel | Intune country comes from the primary user's Entra CountryOrRegion; unmatched AD country is inferred from the first two letters of the computer name using the selected tenant's private mapping. These are different country definitions, not verified physical locations.</p>
<table role="presentation" style="width:100%;border-collapse:collapse;font-size:11px;color:#0f172a;margin:0 0 10px 0;">
<tr style="background:#e2e8f0;"><th style="padding:8px;text-align:left;">COUNTRY</th><th style="padding:8px;text-align:right;">WINDOWS 11</th><th style="padding:8px;text-align:right;">WINDOWS 10</th><th style="padding:8px;text-align:right;">OS VERSION UNKNOWN</th><th style="padding:8px;text-align:right;">TOTAL</th><th style="padding:8px;text-align:right;">INTUNE</th><th style="padding:8px;text-align:right;">AD WITHOUT INTUNE ID MATCH</th><th style="padding:8px;text-align:right;">INTUNE REFERENCE POLICY</th><th style="padding:8px;text-align:right;">INTUNE OTHER FEATURE UPDATE POLICY ONLY</th><th style="padding:8px;text-align:right;">INTUNE NEITHER POLICY</th></tr>
$combinedCountryRowsHtml
$combinedRowHtml
</table>
<p style="margin:0 0 8px 0;font-size:12px;color:#64748b;">Intune enabled: $($intuneFleetSummary.Total.Total) of $($intuneFleetSummary.AllIntuneWindows) Windows devices; excluded: $($intuneFleetSummary.ExcludedDisabledBoth) disabled in both directories, $($intuneFleetSummary.ExcludedUnqualified) without positive enabled evidence. Country resolved: $intuneCountryResolved/$($intuneFleetSummary.Total.Total). Intune source file times (local): $(Html-Encode $intuneSourceText).</p>
<p style="margin:0 0 8px 0;font-size:12px;color:#64748b;">AD enabled without exact Intune ID match: $($adWithoutIntuneSummary.All.Total) of $($adWithoutIntuneSummary.UnmatchedTotal); of these, $($adWithoutIntuneSummary.EnabledActive45.Total) were active in the last 45 days. Country prefix resolved: $adCountryResolved/$($adWithoutIntuneSummary.All.Total); unmapped names are in Country unknown. Excluded: $($adWithoutIntuneSummary.ExcludedDisabledBoth) disabled in both directories, $($adWithoutIntuneSummary.ExcludedUnqualified) without positive enabled evidence. Ambiguous Intune IDs excluded: $($adWithoutIntuneSummary.AmbiguousIntuneIds). Source file times (local): $(Html-Encode $adSourceText).</p>
$referenceReconciliationHtml
<p style="margin:0 0 22px 0;font-size:12px;color:#64748b;">The combined total adds distinct source IDs after the exact AD ObjectGUID to Intune Azure AD Device ID check; identity differences can still make one physical PC appear twice. Policy columns apply only to the Intune count. Neither policy means absent from exported Feature Update policies, not confirmed outside Windows Autopatch. Fleet OS coverage below retains its reference-policy scope and report OSVersion source.</p>
"@
    }
    else {
        $intuneFleetSection = "<h2 style='margin:0 0 6px 0;font-size:20px;color:#0f172a;'>Enabled Windows devices - Intune countries and unmatched AD</h2><p style='margin:0 0 22px 0;font-size:12px;color:#64748b;'>Combined Intune/AD breakdown unavailable: a source is missing, stale, invalid, or could not be reconciled. No combined total is inferred.</p>"
    }
    $fileLinkHtml = if (-not [string]::IsNullOrWhiteSpace($spUploadUrl)) {
        $encodedUrl = Html-Encode $spUploadUrl
        $encodedName = Html-Encode ([IO.Path]::GetFileName($CsvLastFinal))
        "<div style='margin-top:18px;padding:14px;border:1px solid #dbe4ee;background:#f8fafc;'><h3 style='margin:0 0 8px;'>Export</h3><a href='$encodedUrl' style='color:#075985;text-decoration:underline;'>$encodedName</a></div>"
    } else { '' }

    $html = @"
<div style="font-family:Segoe UI,Arial;color:#1f2937;">
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Tenant: $(Html-Encode $OrgDomain) | Policies: $policyCount | Source rows: $count | Duration: $durationSeconds seconds | RunId: $(Html-Encode $RunId)</p>
<div style="display:inline-block;margin-bottom:18px;padding:6px 12px;border-radius:999px;background:$statusBackground;color:$statusColor;font-weight:700;">$reportStatus</div>
$severityInputsHtml
$windowsVersionDistributionSection
$intuneFleetSection
$fleetCoverageSection
<h2 style="margin:24px 0 6px 0;font-size:18px;color:#0f172a;">Operational policy state - all policies</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Each device is counted once across all policies for operational follow-up. The most severe policy state is retained here only for actions and workflow monitoring; it is not used as the fleet OS coverage KPI.</p>
<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;"><tr>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">UNIQUE DEVICES</div><div style="font-size:24px;font-weight:700;">$totalUniqueDevices</div></td>
<td style="padding:14px;background:#f0fdf4;border:1px solid #bbf7d0;"><div style="font-size:11px;color:#166534;">POLICY COMPLETED</div><div style="font-size:24px;font-weight:700;">$completedCount</div></td>
<td style="padding:14px;background:#dbeafe;border:1px solid #bfdbfe;"><div style="font-size:11px;color:#1d4ed8;">IN PROGRESS</div><div style="font-size:24px;font-weight:700;">$inProgressCount</div></td>
<td style="padding:14px;background:#fee2e2;border:1px solid #fecaca;"><div style="font-size:11px;color:#b91c1c;">ACTION REQUIRED</div><div style="font-size:24px;font-weight:700;">$actionRequiredCount</div></td>
<td style="padding:14px;background:#fef3c7;border:1px solid #fde68a;"><div style="font-size:11px;color:#92400e;">UNKNOWN</div><div style="font-size:24px;font-weight:700;">$unknownCount</div></td>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">POLICY COMPLETION</div><div style="font-size:24px;font-weight:700;">$completionPct%</div></td>
</tr></table>
<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;margin-top:2px;"><tr>
<td style="padding:11px;background:#eff6ff;border:1px solid #bfdbfe;"><div style="font-size:10px;color:#1d4ed8;">OFFERING</div><div style="font-size:19px;font-weight:700;">$offeringCount</div></td>
<td style="padding:11px;background:#eff6ff;border:1px solid #bfdbfe;"><div style="font-size:10px;color:#1d4ed8;">INSTALLING</div><div style="font-size:19px;font-weight:700;">$installingCount</div></td>
<td style="padding:11px;background:#eff6ff;border:1px solid #bfdbfe;"><div style="font-size:10px;color:#1d4ed8;">PENDING</div><div style="font-size:19px;font-weight:700;">$pendingCount</div></td>
<td style="padding:11px;background:#eff6ff;border:1px solid #bfdbfe;"><div style="font-size:10px;color:#1d4ed8;">OTHER IN PROGRESS</div><div style="font-size:19px;font-weight:700;">$otherInProgressCount</div></td>
</tr></table>
<h2 style="margin:24px 0 6px 0;font-size:18px;color:#0f172a;">Feature Update policies</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Each device is counted once within each policy. A device targeted by several policies appears in every relevant policy strip, so policy totals are not additive and may exceed the global unique-device total.</p>
$policyStripsHtml
$actionTable
$readinessTable
$fileLinkHtml
</div>
"@

    [pscustomobject]@{Subject=$subject; BodyHtml=$html}
}

function New-SmartM365SharePointLinksSection {
    [CmdletBinding()]
    param([array]$UploadRecords)

    $records = @($UploadRecords | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.WebUrl) })
    if ($records.Count -eq 0) { return $null }

    $rows = foreach ($record in $records) {
        $label = ConvertTo-SmartM365EmailHtmlText $record.Label
        if ([string]::IsNullOrWhiteSpace($label)) { $label = ConvertTo-SmartM365EmailHtmlText $record.FileName }
        $url = ConvertTo-SmartM365EmailHtmlText $record.WebUrl
        $sharePointPath = ConvertTo-SmartM365EmailHtmlText $record.SharePointPath
        "<tr><td style=`"width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;`">$label</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;`"><a href=`"$url`" style=`"color:#2563eb;text-decoration:underline;`">$sharePointPath</a></td></tr>"
    }

    $tableHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  $($rows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = 'SharePoint links'; Html = $tableHtml }
}

function New-SmartM365AdDuplicatePreviewSection {
    [CmdletBinding()]
    param(
        [array]$Rows,
        [Parameter(Mandatory = $true)]
        [ValidateSet('UPN','SMTP','REMOTE')]
        [string]$DuplicateType,
        [int]$Limit = 10
    )

    $items = @($Rows | Where-Object { $_ })
    if ($items.Count -eq 0) { return $null }

    $valueProperty = switch ($DuplicateType) {
        'UPN' { 'UserPrincipalName' }
        'SMTP' { 'SmtpAddress' }
        'REMOTE' { 'NormalizedRemoteRoutingAddress' }
    }
    $label = switch ($DuplicateType) {
        'UPN' { 'duplicate UPN' }
        'SMTP' { 'duplicate SMTP address' }
        'REMOTE' { 'duplicate remote routing address' }
    }
    $title = switch ($DuplicateType) {
        'UPN' { "Top $Limit duplicate UPN accounts" }
        'SMTP' { "Top $Limit duplicate SMTP entries" }
        'REMOTE' { "Top $Limit duplicate remote routing addresses" }
    }

    $groups = @(
        $items |
            Where-Object {
                $property = $_.PSObject.Properties[$valueProperty]
                $null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)
            } |
            Group-Object -Property $valueProperty |
            Where-Object { $_.Count -gt 1 } |
            Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, Name
    )
    if ($groups.Count -eq 0) { return $null }

    $topGroups = @($groups | Select-Object -First $Limit)
    $tableRows = foreach ($group in $topGroups) {
        $sortedRows = @($group.Group | Sort-Object DomainName, SamAccountName, UserPrincipalName)
        $firstInGroup = $true
        foreach ($row in $sortedRows) {
            $valueCell = if ($firstInGroup) { ConvertTo-SmartM365EmailHtmlText $group.Name } else { '' }
            $countCell = if ($firstInGroup) { ConvertTo-SmartM365EmailHtmlText $group.Count } else { '' }
            $firstInGroup = $false

            $domainNameShort = [string]$row.DomainNameShort
            $domainName = [string]$row.DomainName
            $samAccountName = [string]$row.SamAccountName
            $userPrincipalName = [string]$row.UserPrincipalName
            $accountText = ''
            if (-not [string]::IsNullOrWhiteSpace($domainNameShort) -and -not [string]::IsNullOrWhiteSpace($samAccountName)) { $accountText = '{0}\{1}' -f $domainNameShort, $samAccountName }
            elseif (-not [string]::IsNullOrWhiteSpace($samAccountName)) { $accountText = $samAccountName }
            elseif (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) { $accountText = $userPrincipalName }

            $enabledText = [string]$row.Enabled
            if ([string]::IsNullOrWhiteSpace($enabledText)) { $enabledText = 'Unknown' }
            $lastLogonText = [string]$row.LastLogonDate
            if ([string]::IsNullOrWhiteSpace($lastLogonText)) { $lastLogonText = 'Never / unavailable' }

            $accountHtml = ConvertTo-SmartM365EmailHtmlText $accountText
            $enabledHtml = ConvertTo-SmartM365EmailHtmlText $enabledText
            $lastLogonHtml = ConvertTo-SmartM365EmailHtmlText $lastLogonText
            $domainHtml = ConvertTo-SmartM365EmailHtmlText $domainName

            "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$valueCell</td><td align=`"right`" style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:13px;font-weight:700;color:#111827;`">$countCell</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$accountHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$enabledHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;white-space:nowrap;`">$lastLogonHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$domainHtml</td></tr>"
        }
    }

    $caption = ConvertTo-SmartM365EmailHtmlText ('Showing accounts for top {0} of {1} {2} value(s). Full details are available in the attached Excel workbook.' -f $topGroups.Count, $groups.Count, $label)
    $html = @"
<div style="font-size:13px;color:#64748b;margin-bottom:8px;">$caption</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Value</th>
    <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Count</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Account</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Enabled</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Last logon</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Domain</th>
  </tr>
  $($tableRows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = $title; Html = $html }
}

function New-SmartM365AdRemoteRoutingIssuePreviewSection {
    [CmdletBinding()]
    param(
        [array]$Rows,
        [int]$Limit = 10
    )

    $items = @($Rows | Where-Object { $_ } | Sort-Object Severity, IssueType, DomainName, SamAccountName | Select-Object -First $Limit)
    if ($items.Count -eq 0) { return $null }

    $tableRows = foreach ($row in $items) {
        $accountText = if (-not [string]::IsNullOrWhiteSpace([string]$row.DomainNameShort) -and -not [string]::IsNullOrWhiteSpace([string]$row.SamAccountName)) {
            '{0}\{1}' -f $row.DomainNameShort, $row.SamAccountName
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$row.UserPrincipalName)) { [string]$row.UserPrincipalName }
        else { [string]$row.DistinguishedName }

        $issueHtml = ConvertTo-SmartM365EmailHtmlText $row.IssueType
        $severityHtml = ConvertTo-SmartM365EmailHtmlText $row.Severity
        $accountHtml = ConvertTo-SmartM365EmailHtmlText $accountText
        $targetAddressHtml = ConvertTo-SmartM365EmailHtmlText $row.TargetAddress
        $expectedDomainHtml = ConvertTo-SmartM365EmailHtmlText $row.ExpectedRemoteRoutingDomain
        "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$issueHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#334155;`">$severityHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$accountHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;`">$targetAddressHtml</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$expectedDomainHtml</td></tr>"
    }

    $caption = ConvertTo-SmartM365EmailHtmlText ('Showing {0} of {1} remote routing issue row(s). Full details are available in the attached Excel workbook.' -f $items.Count, @($Rows).Count)
    $html = @"
<div style="font-size:13px;color:#64748b;margin-bottom:8px;">$caption</div>
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Issue</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Severity</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Account</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">targetAddress</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Expected domain</th>
  </tr>
  $($tableRows -join "`n")
</table>
"@

    return [pscustomobject]@{ Title = "Top $Limit remote routing issues"; Html = $html }
}

function Get-SmartM365AdCaseInsensitiveDistinctCount {
    param([array]$Rows, [Parameter(Mandatory = $true)][string]$PropertyName)
    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        $property = $row.PSObject.Properties[$PropertyName]
        if ($null -eq $property) { continue }
        $value = ([string]$property.Value).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) { [void]$values.Add($value) }
    }
    return $values.Count
}

function New-SmartM365AdIdentitySummaryHtml {
    [CmdletBinding()]
    param([array]$AllUsers, [array]$DuplicateUpnRows, [array]$DuplicateSmtpRows,
          [array]$DuplicateRemoteRows, [array]$RoutingIssueRows)

    $domains = @($AllUsers | ForEach-Object { if ([string]::IsNullOrWhiteSpace([string]$_.DomainName)) { 'Unknown domain' } else { [string]$_.DomainName } } | Sort-Object -Unique)
    $sets = @(
        [pscustomobject]@{ Title = 'Duplicate identities'; Metrics = @(
            [pscustomobject]@{ Label = 'UPN values'; Rows = $DuplicateUpnRows; Distinct = 'UserPrincipalName'; IssueType = '' }
            [pscustomobject]@{ Label = 'UPN accounts'; Rows = $DuplicateUpnRows; Distinct = ''; IssueType = '' }
            [pscustomobject]@{ Label = 'SMTP values'; Rows = $DuplicateSmtpRows; Distinct = 'SmtpAddress'; IssueType = '' }
            [pscustomobject]@{ Label = 'SMTP entries'; Rows = $DuplicateSmtpRows; Distinct = ''; IssueType = '' }
            [pscustomobject]@{ Label = 'Routing values'; Rows = $DuplicateRemoteRows; Distinct = 'NormalizedRemoteRoutingAddress'; IssueType = '' }
            [pscustomobject]@{ Label = 'Routing accounts'; Rows = $DuplicateRemoteRows; Distinct = ''; IssueType = '' }
        ) }
        [pscustomobject]@{ Title = 'Remote routing issues'; Metrics = @(
            [pscustomobject]@{ Label = 'Accounts'; Rows = $RoutingIssueRows; Distinct = 'DistinguishedName'; IssueType = '' }
            [pscustomobject]@{ Label = 'Missing target'; Rows = $RoutingIssueRows; Distinct = ''; IssueType = 'MissingRemoteRoutingAddress' }
            [pscustomobject]@{ Label = 'Invalid target'; Rows = $RoutingIssueRows; Distinct = ''; IssueType = 'InvalidRemoteRoutingAddress' }
            [pscustomobject]@{ Label = 'Wrong domain'; Rows = $RoutingIssueRows; Distinct = ''; IssueType = 'UnexpectedRemoteRoutingDomain' }
            [pscustomobject]@{ Label = 'Target absent in proxy'; Rows = $RoutingIssueRows; Distinct = ''; IssueType = 'RemoteRoutingAddressMissingFromProxyAddresses' }
            [pscustomobject]@{ Label = 'Tenant proxy missing'; Rows = $RoutingIssueRows; Distinct = ''; IssueType = 'MissingMailOnMicrosoftProxyAddress' }
        ) }
    )
    $tables = foreach ($set in $sets) {
        $headers = @($set.Metrics | ForEach-Object { '<th align="right" style="padding:6px 4px;">{0}</th>' -f (ConvertTo-SmartM365EmailHtmlText $_.Label) }) -join ''
        $rows = foreach ($domain in $domains) {
            $cells = foreach ($metric in $set.Metrics) {
                $items = @($metric.Rows | Where-Object {
                    $rowDomain = if ([string]::IsNullOrWhiteSpace([string]$_.DomainName)) { 'Unknown domain' } else { [string]$_.DomainName }
                    $rowDomain -eq $domain -and ([string]::IsNullOrWhiteSpace($metric.IssueType) -or [string]$_.IssueType -eq $metric.IssueType)
                })
                $value = if ($metric.Distinct) { Get-SmartM365AdCaseInsensitiveDistinctCount -Rows $items -PropertyName $metric.Distinct } else { $items.Count }
                '<td align="right" style="border-bottom:1px solid #eef2f7;padding:6px 4px;">{0}</td>' -f $value
            }
            '<tr><td style="border-bottom:1px solid #eef2f7;padding:6px 4px;">{0}</td>{1}</tr>' -f (ConvertTo-SmartM365EmailHtmlText $domain), ($cells -join '')
        }
        $totalCells = foreach ($metric in $set.Metrics) {
            $items = @($metric.Rows | Where-Object { [string]::IsNullOrWhiteSpace($metric.IssueType) -or [string]$_.IssueType -eq $metric.IssueType })
            $value = if ($metric.Distinct) { Get-SmartM365AdCaseInsensitiveDistinctCount -Rows $items -PropertyName $metric.Distinct } else { $items.Count }
            '<td align="right" style="background:#dbeafe;padding:7px 4px;font-weight:700;">{0}</td>' -f $value
        }
        '<div style="font-size:12px;font-weight:700;margin:12px 0 6px;">{0}</div><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;font-size:10px;"><tr style="background:#f1f5f9;"><th align="left">Domain</th>{1}</tr>{2}<tr><td style="background:#dbeafe;padding:7px 4px;font-weight:700;">TOTAL</td>{3}</tr></table>' -f $set.Title, $headers, ($rows -join ''), ($totalCells -join '')
    }
    return (($tables -join "`n") + '<div style="font-size:11px;color:#64748b;margin-top:6px;">Distinct values may span domains; TOTAL counts distinct values across the forest.</div>')
}

function Export-SmartM365AdDiagnosticsWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$Sources,
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$ImportExcelAlreadyLoaded
    )

    if (-not $ImportExcelAlreadyLoaded) { Initialize-SmartM365ImportExcel | Out-Null }

    $parentPath = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentPath) -and -not (Test-Path -LiteralPath $parentPath)) {
        New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    }

    $excelPackage = $null
    try {
        $excelPackage = Open-ExcelPackage -Path $Path -Create
        foreach ($source in @($Sources)) {
            $csvPath = [string]$source.CsvPath
            $worksheetName = [string]$source.WorksheetName
            $tableName = [string]$source.TableName
            if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
                throw ("Diagnostic CSV not found: {0}" -f $csvPath)
            }

            $rows = @(Import-Csv -LiteralPath $csvPath)
            if ($rows.Count -gt 0) {
                $columns = @($rows[0].PSObject.Properties.Name)
            }
            else {
                $header = [string](Get-Content -LiteralPath $csvPath -TotalCount 1 -ErrorAction Stop)
                if ([string]::IsNullOrWhiteSpace($header)) {
                    throw ("Diagnostic CSV has no header: {0}" -f $csvPath)
                }
                $separatorCount = ([regex]::Matches($header, ',')).Count
                $fakeRow = (@('x') * ($separatorCount + 1)) -join ','
                $probe = @(@($header, $fakeRow) | ConvertFrom-Csv)
                $columns = if ($probe.Count -gt 0) { @($probe[0].PSObject.Properties.Name) } else { @() }
            }
            if ($columns.Count -eq 0) {
                throw ("Diagnostic CSV columns could not be resolved: {0}" -f $csvPath)
            }

            $worksheet = $excelPackage.Workbook.Worksheets.Add($worksheetName)
            $matrix = [System.Collections.Generic.List[object[]]]::new()
            $matrix.Add([object[]]$columns)
            foreach ($row in $rows) {
                $values = foreach ($column in $columns) {
                    $property = $row.PSObject.Properties[$column]
                    if ($null -eq $property -or $null -eq $property.Value) { '' } else { [string]$property.Value }
                }
                $matrix.Add([object[]]$values)
            }
            $null = $worksheet.Cells['A1'].LoadFromArrays($matrix)
            $worksheet.View.FreezePanes(2, 1)
            $worksheet.Cells.AutoFitColumns(10, 60)

            $headerAddress = [OfficeOpenXml.ExcelCellBase]::GetAddress(1, 1, 1, $columns.Count)
            $headerRange = $worksheet.Cells[$headerAddress]
            $headerRange.Style.Font.Bold = $true
            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))

            if ($rows.Count -gt 0) {
                $tableAddress = [OfficeOpenXml.ExcelCellBase]::GetAddress(1, 1, $rows.Count + 1, $columns.Count)
                $tableRange = $worksheet.Cells[$tableAddress]
                $table = $worksheet.Tables.Add($tableRange, $tableName)
                $table.TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium2
            }
            else {
                $headerRange.AutoFilter = $true
            }
        }

        Close-ExcelPackage -ExcelPackage $excelPackage
        $excelPackage = $null
    }
    finally {
        if ($null -ne $excelPackage) {
            Close-ExcelPackage -ExcelPackage $excelPackage -NoSave
        }
    }

    $workbookFile = Get-Item -LiteralPath $Path -ErrorAction Stop
    WriteLog -Message ("AD identity and mail routing workbook created: {0} ({1:N0} bytes)" -f $workbookFile.FullName, $workbookFile.Length)
    return $workbookFile.FullName
}

function Get-SmartM365AdDiagnosticWorkbookSource {
    param([string]$Folder)
    foreach ($entry in @(
        @('AD_Users_DuplicateUPN.csv','Duplicate UPN','DuplicateUPN'),
        @('AD_Users_DuplicateSMTP.csv','Duplicate SMTP','DuplicateSMTP'),
        @('AD_Users_DuplicateRemoteRoutingAddress.csv','Duplicate RemoteRouting','DuplicateRemoteRouting'),
        @('AD_Users_RemoteRoutingIssues.csv','RemoteRouting Issues','RemoteRoutingIssues')
    )) {
        [pscustomobject]@{ Name=$entry[0]; CsvPath=(Join-Path $Folder $entry[0]); WorksheetName=$entry[1]; TableName=$entry[2] }
    }
}

function New-SmartM365AdIdentityMailSections {
    param([object[]]$DuplicateUpnRows,[object[]]$DuplicateSmtpRows,[object[]]$DuplicateRemoteRows,[object[]]$RoutingIssueRows,[int]$Limit)
    New-SmartM365AdDuplicatePreviewSection -Rows $DuplicateUpnRows -DuplicateType UPN -Limit $Limit
    New-SmartM365AdDuplicatePreviewSection -Rows $DuplicateSmtpRows -DuplicateType SMTP -Limit $Limit
    New-SmartM365AdDuplicatePreviewSection -Rows $DuplicateRemoteRows -DuplicateType REMOTE -Limit $Limit
    New-SmartM365AdRemoteRoutingIssuePreviewSection -Rows $RoutingIssueRows -Limit $Limit
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCArEZoCeiKPxqn5
# uTr1MUnGmxRM4xPCvf1zCa9zBybfIaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgOMCjSHmzYDz0
# m7XWaUdmLzhTG6+vM6CBtxWxmGiBI78wDQYJKoZIhvcNAQEBBQAEggGAkxM4g6f+
# Ck9u8VL62yo9otuXn6jOWvPl/0OHFWBwhTfASB7u11Klk/w7TMrBJNqdQv2pG7Ej
# 84e80+wkqMu9cGLLqJlLERRznD29ivWuoKlJcmrceG4yPgyfckpYpOvN24nnqUJb
# oQDFSCDgtlrICSbDV1ERC7JvGeSShxCUOjL0laO/8u5KVrRV9hhlapYFdI/29jh3
# 8TIdaVZEbnrtXZFNpha+QgIvhgTdVd7gqMK7rVwKDlSw/U853UYoPWVsb6FSD5XS
# EyQDZI5OAMDmGGEIynn0CTrm73Os/0oI9CwToLmXlCnqAxcmHY/J9IJ/0JeZlihV
# tyL7Ye+tAyjqbWweBXretqD5RktMc4USGAo5y8TJCoCknCha7UBTrfqFeSZ0wZKW
# NL2/jZwon4rEHQLYOLQUc/7aHOeyF4Txu+HI1AQ3ZhWWklkMeE144HFBJFTfghPK
# 7iBNkMRw+DxyayKeYEfhHlWNR4LHOoyw9hbCmk4eacL7V7zdvQyPDPOUoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTM3WjA/BgkqhkiG9w0BCQQxMgQw+lT/ZTXeBLW1SE5Vo14EVzGa
# 4JqTHsOoopKBEu4TAIdRnTlYUzva/dz0pTWebVJzMA0GCSqGSIb3DQEBAQUABIIC
# AEAjtKFHRkg9LfiL6Z2HzvFbn6kXi8O9Cw39qNoLyr+Pzsu6u80t6Xxbpy//fxEO
# 22jCzyvxjr38tvOVGBUnf54iKwxxJn53nUm1DzQPz7Bv1fOeVOL3aVg09u+7bfBQ
# i3HEtpSFGzoyioilQd9L5i/LINzD1D4y4PCbgdQLImvuQQDW5ywJDuiCyZRiwbCl
# W3dhWTMCzi53P+epneA/6YLBHRSpoBklbPaZXlYlD5ueEcD8ZzR2rtprp0FGAXJA
# f4oEP+gcxJO9TcN1QOs8zhtkpJ1TbsCYKN8sPrmS3ubyg5vdjMZ+CCDM43R4Clw4
# Fw6xg97MLhYfUDAFazoZ8NerKw/2IQ6x3EGMEraN2sE3iZu8RqLXsG7ieUifLI9a
# 8J056/DISSMD7WPf7ynrW9iVnv5F9A8IRfoez8ZEahxVv62niIEKoI/+sjMeeptC
# EQKLRWZ1ZEP0oMXJmbwl/R3cmezTF6vao4oQBk6a/7yM8U8kX5H42BqgCyRnwzk6
# kptWfy3xx8xpePFIYQbKqNnbguF7tAgs8bTRBllvzspZBZUXxmNl6I1iG1/9EKVW
# sPw9QyxyxeSQVQKm6FFReV4nNiOkI/0SNeeb/S//AyG15lFnJz7NV5duMMAFvDaM
# my9pu1psJoBujaYbpt8XxNOsNUQmwyQHS40kUN1DCLDr
# SIG # End signature block
