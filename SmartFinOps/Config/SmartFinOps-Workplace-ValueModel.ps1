function Get-PriceModel {
    [CmdletBinding()]
    param()

    $baseline = if ($script:SmartFinOpsPriceBaselinePath -and (Test-Path -LiteralPath $script:SmartFinOpsPriceBaselinePath -PathType Leaf)) {
        Get-Content -LiteralPath $script:SmartFinOpsPriceBaselinePath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    }
    else {
        [pscustomobject]@{
            Name = 'No price baseline'
            AsOfDate = ''
            Currency = 'EUR'
            TaxBasis = 'HT'
            CommitmentBasis = ''
            Method = ''
            MonthlyUnitPriceBySkuPartNumber = [pscustomobject]@{}
            Sources = @()
        }
    }

    $prices = [ordered]@{}
    $baselineMap = $baseline.PSObject.Properties['MonthlyUnitPriceBySkuPartNumber']
    if ($baselineMap -and $baselineMap.Value) {
        foreach ($property in $baselineMap.Value.PSObject.Properties) { $prices[$property.Name] = [decimal]$property.Value }
    }

    $override = Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'PriceModel' -DefaultValue $null
    if ($override -and $override -isnot [string]) {
        $overrideMap = $override.PSObject.Properties['MonthlyUnitPriceBySkuPartNumber']
        if ($overrideMap -and $overrideMap.Value) {
            foreach ($property in $overrideMap.Value.PSObject.Properties) { $prices[$property.Name] = [decimal]$property.Value }
        }
    }

    [pscustomobject]@{
        Name = [string]$baseline.Name
        AsOfDate = [string]$baseline.AsOfDate
        Currency = if ($override -and $override -isnot [string] -and $override.PSObject.Properties['Currency']) { [string]$override.Currency } else { [string]$baseline.Currency }
        TaxBasis = [string]$baseline.TaxBasis
        CommitmentBasis = [string]$baseline.CommitmentBasis
        Method = [string]$baseline.Method
        MonthlyUnitPriceBySkuPartNumber = [pscustomobject]$prices
        Sources = @($baseline.Sources)
    }
}

function Get-SmartFinOpsSummaryMetricValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SummaryRows,
        [Parameter(Mandatory)][string]$Metric,
        $DefaultValue = 0
    )
    $row = $SummaryRows | Where-Object { $_.Metric -eq $Metric } | Select-Object -First 1
    if ($null -eq $row) { return $DefaultValue }
    return $row.Value
}

function New-SmartFinOpsValueOpportunityRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SummaryRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$UserDecisionRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseCapacityRows,
        [Parameter(Mandatory)][AllowNull()]$PriceModel
    )

    $rows = New-Object System.Collections.Generic.List[object]
    # Product-specific add-ons are not monetized from generic M365 inactivity; dedicated telemetry is required.

    $noLicenseDecisions = @($UserDecisionRows | Where-Object {
        $_.CurrentBaseSku -in @('SPE_E3', 'SPE_F1') -and
        $_.RecommendedLicense -match '^No license' -and
        $null -ne $_.CurrentMonthlyPriceEUR
    })
    foreach ($group in ($noLicenseDecisions | Group-Object -Property CurrentBaseSku)) {
        $sample = $group.Group | Select-Object -First 1
        $highConfidence = @($group.Group | Where-Object { $_.RecommendedLicense -eq 'No license - candidate' })
        $review = @($group.Group | Where-Object { $_.RecommendedLicense -eq 'No license - review' })
        $monthlyValue = [math]::Round([decimal](($group.Group | Measure-Object -Property IndicativeMonthlyDifferenceEUR -Sum).Sum), 2)
        $highConfidenceMonthly = [math]::Round([decimal](($highConfidence | Measure-Object -Property IndicativeMonthlyDifferenceEUR -Sum).Sum), 2)
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = 1
            ValuePillar = 'Potential savings'
            Opportunity = "Validate and remove $($sample.CurrentBaseLicense) licenses with no observed need"
            SkuPartNumber = $group.Name
            Population = $group.Count
            HighConfidencePopulation = $highConfidence.Count
            ReviewPopulation = $review.Count
            MonthlyValueEUR = $monthlyValue
            AnnualValueEUR = [math]::Round([decimal]($monthlyValue * 12), 2)
            HighConfidenceMonthlyValueEUR = $highConfidenceMonthly
            Confidence = if ($review.Count -eq 0) { 'High' } else { 'Medium' }
            Decision = 'Confirm departure, retention, ownership, and all detailed activity before removing the license.'
            FinancialTreatment = 'Indicative potential based on a consolidated user decision; it becomes a realized saving only when purchased units are reduced.'
        }) | Out-Null
    }

    $e3ToF3Decisions = @($UserDecisionRows | Where-Object {
        $_.CurrentBaseSku -eq 'SPE_E3' -and
        $_.RecommendedLicense -eq 'M365 F3' -and
        $null -ne $_.IndicativeMonthlyDifferenceEUR -and
        [decimal]$_.IndicativeMonthlyDifferenceEUR -gt 0
    })
    if ($e3ToF3Decisions.Count -gt 0) {
        $monthlyValue = [math]::Round([decimal](($e3ToF3Decisions | Measure-Object -Property IndicativeMonthlyDifferenceEUR -Sum).Sum), 2)
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = 1
            ValuePillar = 'Potential savings'
            Opportunity = 'Move E3 to F3 when target persona and usage support it'
            SkuPartNumber = 'SPE_E3->SPE_F1'
            Population = $e3ToF3Decisions.Count
            HighConfidencePopulation = 0
            ReviewPopulation = $e3ToF3Decisions.Count
            MonthlyValueEUR = $monthlyValue
            AnnualValueEUR = [math]::Round([decimal]($monthlyValue * 12), 2)
            HighConfidenceMonthlyValueEUR = 0
            Confidence = 'Medium'
            Decision = "Validate with the business that desktop applications are not required and that F3 limits are acceptable."
            FinancialTreatment = "Indicative rightsizing scenario; the E3-F3 difference is realizable only after functional and contractual validation."
        }) | Out-Null
    }
    # LicenseOptimization retains add-on review findings without turning them into unsupported savings.

    foreach ($capacity in $LicenseCapacityRows) {
        $available = [decimal]$capacity.AvailableUnits
        if ($available -le 0) { continue }
        $unitPrice = Get-MonthlySkuPrice -PriceModel $PriceModel -SkuPartNumber ([string]$capacity.SkuPartNumber)
        if ($null -eq $unitPrice -or [decimal]$unitPrice -le 0) { continue }
        $monthlyValue = [math]::Round([decimal]($available * [decimal]$unitPrice), 2)
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = 2
            ValuePillar = 'Cost avoidance'
            Opportunity = "Reuse available capacity for $($capacity.SkuDisplayName) before purchasing additional licenses"
            SkuPartNumber = [string]$capacity.SkuPartNumber
            Population = [int]$available
            HighConfidencePopulation = [int]$available
            ReviewPopulation = 0
            MonthlyValueEUR = $monthlyValue
            AnnualValueEUR = [math]::Round([decimal]($monthlyValue * 12), 2)
            HighConfidenceMonthlyValueEUR = $monthlyValue
            Confidence = 'High'
            Decision = 'Use already purchased units before ordering additional licenses; adjust quantities at the next renewal.'
            FinancialTreatment = 'Cost-avoidance equivalent; do not add it to potential savings without validating the contract schedule.'
        }) | Out-Null
    }

    $nonFinancialDefinitions = @(
        [pscustomobject]@{ Priority = 2; Pillar = 'Potential savings delivery'; Metric = 'Strong shared mailbox conversion candidates'; Opportunity = 'Convert eligible disabled user mailboxes to shared mailboxes'; Decision = 'Validate ownership, retention, application dependencies, delegates, mailbox growth, and contract terms before conversion and license removal.'; Confidence = 'High'; FinancialTreatment = 'Supporting action path only; its license value is already included in no-license potential and is not added again.' },
        [pscustomobject]@{ Priority = 3; Pillar = 'Value and fit'; Metric = 'E3 personas missing E3'; Opportunity = 'Align licenses with the actual requirements of E3 target personas'; Decision = 'Balance cost, required capabilities, and user experience before any upgrade or downgrade.'; Confidence = 'Medium' },
        [pscustomobject]@{ Priority = 3; Pillar = 'Service quality'; Metric = 'F3 mailbox size non-compliance'; Opportunity = 'Bring F3 mailboxes within limits or change the license'; Decision = 'Choose between cleanup, archiving, and a license change based on total cost and usage.'; Confidence = 'High' },
        [pscustomobject]@{ Priority = 3; Pillar = 'Continuity'; Metric = 'Unlicensed F3 migration blockers'; Opportunity = 'Remove F3 migration blockers'; Decision = 'Resolve blockers before migration to avoid delays, rework, and emergency purchases.'; Confidence = 'High' },
        [pscustomobject]@{ Priority = 4; Pillar = 'Performance and lifecycle'; Metric = 'AD computers requiring Windows 11 upgrade'; Opportunity = 'Plan Windows 11 upgrades or device replacement'; Decision = 'Prioritize by compatibility, business criticality, and fleet maintenance cost.'; Confidence = 'High' },
        [pscustomobject]@{ Priority = 4; Pillar = 'Quality and security'; Metric = 'AD computers pending Entra registration'; Opportunity = 'Complete device Entra join'; Decision = 'Reduce management gaps that affect security, support, and productivity.'; Confidence = 'High' },
        [pscustomobject]@{ Priority = 4; Pillar = 'Operational risk'; Metric = 'Remote routing distinct users'; Opportunity = 'Resolve Exchange routing anomalies'; Decision = 'Resolve before migration or license changes to prevent incidents and rollback.'; Confidence = 'High' },
        [pscustomobject]@{ Priority = 4; Pillar = 'Operational risk'; Metric = 'Proxy address non-OK findings'; Opportunity = 'Resolve Exchange address anomalies'; Decision = 'Remove conflicts and inconsistencies before transformation work.'; Confidence = 'High' }
    )

    foreach ($definition in $nonFinancialDefinitions) {
        $population = [int](Get-SmartFinOpsSummaryMetricValue -SummaryRows $SummaryRows -Metric $definition.Metric -DefaultValue 0)
        if ($population -le 0) { continue }
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = $definition.Priority
            ValuePillar = $definition.Pillar
            Opportunity = $definition.Opportunity
            SkuPartNumber = ''
            Population = $population
            HighConfidencePopulation = $population
            ReviewPopulation = 0
            MonthlyValueEUR = ''
            AnnualValueEUR = ''
            HighConfidenceMonthlyValueEUR = ''
            Confidence = $definition.Confidence
            Decision = $definition.Decision
            FinancialTreatment = if ($definition.PSObject.Properties['FinancialTreatment']) { [string]$definition.FinancialTreatment } else { 'Impact is not monetized with current data; manage it as value protection, quality, performance, or risk.' }
        }) | Out-Null
    }

    return @($rows | Sort-Object Priority, @{ Expression = { if ([string]::IsNullOrWhiteSpace([string]$_.AnnualValueEUR)) { 0 } else { -[decimal]$_.AnnualValueEUR } } })
}

function Format-SmartFinOpsEuro {
    [CmdletBinding()]
    param([AllowNull()]$Value, [switch]$Compact)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 'Not quantified' }
    $amount = [decimal]$Value
    if ($Compact -and [math]::Abs([double]$amount) -ge 1000) { return ('€{0}k' -f (([double]$amount / 1000).ToString('N1', [Globalization.CultureInfo]::GetCultureInfo('en-GB')))) }
    return ('€{0}' -f ($amount.ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-GB'))))
}

function Write-SmartFinOpsHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SummaryRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$UserDecisionRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseCapacityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DeviceRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ExchangeRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DataQualityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ValueRows,
        [Parameter(Mandatory)][AllowNull()]$PriceModel
    )

    $potentialRows = @($ValueRows | Where-Object { $_.ValuePillar -eq 'Potential savings' })
    $avoidanceRows = @($ValueRows | Where-Object { $_.ValuePillar -eq 'Cost avoidance' })
    $potentialMonthly = [decimal](($potentialRows | Measure-Object -Property MonthlyValueEUR -Sum).Sum)
    $potentialAnnual = [decimal](($potentialRows | Measure-Object -Property AnnualValueEUR -Sum).Sum)
    $highConfidenceMonthly = [decimal](($potentialRows | Measure-Object -Property HighConfidenceMonthlyValueEUR -Sum).Sum)
    $avoidanceMonthly = [decimal](($avoidanceRows | Measure-Object -Property MonthlyValueEUR -Sum).Sum)
    $pricedCandidates = [int](($potentialRows | Measure-Object -Property Population -Sum).Sum)
    $highConfidenceCandidates = [int](($potentialRows | Measure-Object -Property HighConfidencePopulation -Sum).Sum)
    $reviewCandidates = [int](($potentialRows | Measure-Object -Property ReviewPopulation -Sum).Sum)
    $noLicensePricedCandidates = [int](($potentialRows | Where-Object { $_.SkuPartNumber -notmatch '->' } | Measure-Object -Property Population -Sum).Sum)
    $downgradePricedCandidates = [int](($potentialRows | Where-Object { $_.SkuPartNumber -match '->' } | Measure-Object -Property Population -Sum).Sum)
    $staleSources = @($DataQualityRows | Where-Object { $_.FreshnessStatus -eq 'Stale' }).Count
    $invalidSources = @($DataQualityRows | Where-Object { $_.ContractStatus -eq 'Invalid' -or $_.Status -eq 'Error' }).Count
    $e3Available = [int](($LicenseCapacityRows | Where-Object SkuPartNumber -eq 'SPE_E3' | Select-Object -First 1).AvailableUnits)
    $f3Available = [int](($LicenseCapacityRows | Where-Object SkuPartNumber -eq 'SPE_F1' | Select-Object -First 1).AvailableUnits)
    $maxBarValue = [decimal](($potentialRows | Measure-Object -Property MonthlyValueEUR -Maximum).Maximum)
    $e3Recommended = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'M365 E3' }).Count
    $f3Recommended = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'M365 F3' }).Count
    $noLicenseHigh = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'No license - candidate' }).Count
    $noLicenseReview = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'No license - review' }).Count
    $e3ToF3Candidates = @($UserDecisionRows | Where-Object { $_.CurrentBaseLicense -eq 'M365 E3' -and $_.RecommendedLicense -eq 'M365 F3' }).Count
    $f3LimitReviews = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'M365 E3 or remediate F3' }).Count
    $sharedMailboxStrong = @($ExchangeRows | Where-Object { $_.FindingType -eq 'SharedMailboxConversion' -and $_.Status -eq 'Strong candidate' }).Count
    $sharedMailboxCapacityReview = @($ExchangeRows | Where-Object { $_.FindingType -eq 'SharedMailboxConversion' -and $_.Status -eq 'Capacity review' }).Count

    $barItems = New-Object System.Collections.Generic.List[string]
    foreach ($row in ($potentialRows | Sort-Object { [decimal]$_.MonthlyValueEUR } -Descending | Select-Object -First 8)) {
        $width = if ($maxBarValue -gt 0) { [math]::Round(([double][decimal]$row.MonthlyValueEUR / [double]$maxBarValue) * 100, 1) } else { 0 }
        $label = if ([string]::IsNullOrWhiteSpace([string]$row.SkuPartNumber)) { $row.Opportunity } else { $row.SkuPartNumber }
        $barItems.Add((@"
        <div class="bar-row" role="img" aria-label="$(ConvertTo-HtmlEncoded $label): $(Format-SmartFinOpsEuro $row.MonthlyValueEUR) per month, $($row.Population) candidate licenses">
          <div class="bar-label"><strong>$(ConvertTo-HtmlEncoded $label)</strong><span>$($row.Population) licenses</span></div>
          <div class="bar-track"><div class="bar-fill" style="width:$width%"></div></div>
          <div class="bar-value">$(Format-SmartFinOpsEuro $row.MonthlyValueEUR)</div>
        </div>
"@)) | Out-Null
    }

    $actionRows = New-Object System.Collections.Generic.List[string]
    foreach ($row in ($ValueRows | Sort-Object Priority, @{ Expression = { if ([string]::IsNullOrWhiteSpace([string]$_.AnnualValueEUR)) { 0 } else { -[decimal]$_.AnnualValueEUR } } } | Select-Object -First 12)) {
        $value = if ([string]::IsNullOrWhiteSpace([string]$row.AnnualValueEUR)) { 'Impact to qualify' } else { "$(Format-SmartFinOpsEuro $row.AnnualValueEUR) / year" }
        $actionRows.Add("<tr><td><span class='priority p$($row.Priority)'>P$($row.Priority)</span></td><td><strong>$(ConvertTo-HtmlEncoded $row.Opportunity)</strong><div class='muted'>$(ConvertTo-HtmlEncoded $row.ValuePillar)</div></td><td class='number'>$($row.Population)</td><td class='number'>$(ConvertTo-HtmlEncoded $value)</td><td>$(ConvertTo-HtmlEncoded $row.Confidence)</td><td>$(ConvertTo-HtmlEncoded $row.Decision)</td></tr>") | Out-Null
    }

    $capacityRows = New-Object System.Collections.Generic.List[string]
    foreach ($row in ($LicenseCapacityRows | Where-Object { $_.SkuPartNumber -in @('SPE_E3', 'SPE_F1') })) {
        $capacityRows.Add("<tr><td><strong>$(ConvertTo-HtmlEncoded $row.SkuDisplayName)</strong></td><td class='number'>$($row.ConsumedUnits) / $($row.EnabledUnits)</td><td class='number'>$($row.UtilizationPercent) %</td><td class='number'>$($row.AvailableUnits)</td><td>$(ConvertTo-HtmlEncoded $row.CapacityStatus)</td></tr>") | Out-Null
    }

    $priceLinks = New-Object System.Collections.Generic.List[string]
    foreach ($source in @($PriceModel.Sources)) {
        $priceLinks.Add("<li><a href='$(ConvertTo-HtmlEncoded $source.Url)' target='_blank' rel='noreferrer'>$(ConvertTo-HtmlEncoded $source.Products)</a></li>") | Out-Null
    }

    $technicalRows = @(
        [pscustomobject]@{ Indicator = 'Sources checked'; Value = $DataQualityRows.Count },
        [pscustomobject]@{ Indicator = 'Stale sources'; Value = $staleSources },
        [pscustomobject]@{ Indicator = 'Invalid contracts or errors'; Value = $invalidSources },
        [pscustomobject]@{ Indicator = 'MAXITEMS files used'; Value = @($DataQualityRows | Where-Object { $_.FileName -match 'MAXITEMS' }).Count }
    )

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>$(ConvertTo-HtmlEncoded $ReportTitle)</title>
  <style>
    :root { --ink:#172033; --muted:#5f6b7a; --line:#dfe5ec; --paper:#ffffff; --canvas:#f4f6f8; --blue:#1665d8; --blue-soft:#eaf2ff; --gold:#b7791f; --gold-soft:#fff7e6; --orange:#c05621; --olive:#667a25; --shadow:0 10px 30px rgba(23,32,51,.08); }
    * { box-sizing:border-box; }
    html, body { max-width:100%; overflow-x:hidden; }
    body { margin:0; font-family:"Segoe UI",Arial,sans-serif; background:var(--canvas); color:var(--ink); line-height:1.55; overflow-wrap:anywhere; }
    a { color:var(--blue); }
    header, main { width:100%; max-width:1180px; margin:auto; }
    header { padding:48px 28px 24px; }
    main { padding:0 28px 56px; }
    .eyebrow { color:var(--blue); font-weight:700; letter-spacing:.08em; text-transform:uppercase; font-size:12px; }
    h1 { font-size:clamp(30px,5vw,48px); line-height:1.08; margin:10px 0 12px; max-width:900px; }
    h2 { font-size:26px; line-height:1.2; margin:0 0 14px; }
    h3 { font-size:18px; margin:0 0 8px; }
    p { margin:0 0 12px; }
    .meta, .muted { color:var(--muted); font-size:13px; }
    section { background:var(--paper); border:1px solid var(--line); border-radius:16px; padding:26px; margin-top:20px; box-shadow:var(--shadow); }
    .executive { border-top:5px solid var(--blue); }
    .executive ul { margin:8px 0 0; padding-left:22px; }
    .executive li { margin:10px 0; }
    .kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin-top:20px; }
    .kpi { background:var(--paper); border:1px solid var(--line); border-radius:14px; padding:18px; }
    .kpi .label { color:var(--muted); font-size:13px; min-height:40px; }
    .kpi .value { font-size:28px; font-weight:750; margin:4px 0; }
    .kpi .context { color:var(--muted); font-size:12px; }
    .bar-chart { margin-top:18px; }
    .bar-row { display:grid; grid-template-columns:minmax(170px,240px) 1fr 105px; align-items:center; gap:14px; margin:13px 0; }
    .bar-label { display:flex; flex-direction:column; }
    .bar-label span { color:var(--muted); font-size:12px; }
    .bar-track { height:18px; background:#edf1f5; border-radius:999px; overflow:hidden; }
    .bar-fill { height:100%; background:var(--blue); border-radius:999px; }
    .bar-value { text-align:right; font-variant-numeric:tabular-nums; font-weight:700; }
    .insight { background:var(--blue-soft); border-left:4px solid var(--blue); padding:14px 16px; border-radius:8px; margin-top:18px; }
    .tradeoffs { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; }
    .tradeoff { border:1px solid var(--line); border-radius:12px; padding:16px; }
    .tradeoff.cost { border-top:4px solid var(--blue); }
    .tradeoff.quality { border-top:4px solid var(--gold); }
    .tradeoff.performance { border-top:4px solid var(--olive); }
    .big { font-size:26px; font-weight:750; display:block; }
    .table-wrap { width:100%; max-width:100%; overflow-x:auto; }
    table { border-collapse:collapse; width:100%; font-size:13px; }
    th { text-align:left; color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.04em; border-bottom:2px solid var(--line); padding:10px 8px; }
    td { border-bottom:1px solid var(--line); padding:12px 8px; vertical-align:top; }
    .number { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
    .priority { display:inline-flex; min-width:30px; height:30px; align-items:center; justify-content:center; border-radius:50%; font-weight:750; }
    .p1 { background:var(--blue); color:white; } .p2 { background:var(--blue-soft); color:var(--blue); } .p3 { background:var(--gold-soft); color:var(--gold); } .p4 { background:#eef3df; color:var(--olive); }
    details { border-top:1px solid var(--line); padding-top:14px; margin-top:14px; }
    summary { cursor:pointer; font-weight:700; }
    .assumptions { font-size:13px; color:var(--muted); }
    .assumptions li { margin:8px 0; }
    @media (max-width:850px) { .kpis,.tradeoffs { grid-template-columns:1fr 1fr; } .bar-row { grid-template-columns:1fr; gap:5px; } .bar-value { text-align:left; } }
    @media (max-width:560px) { header,main { padding-left:16px; padding-right:16px; } section { padding:20px; } .kpis,.tradeoffs { grid-template-columns:1fr; } }
    @media print { body { background:white; } header,main { max-width:none; } section,.kpi { box-shadow:none; break-inside:avoid; } }
  </style>
</head>
<body>
  <header>
    <div class="eyebrow">SmartFinOps Workplace</div>
    <h1>$(ConvertTo-HtmlEncoded $ReportTitle)</h1>
    <p class="meta">Tenant $(ConvertTo-HtmlEncoded $Tenant) · Analysis generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') · Indicative France pricing, excluding VAT · Run $(ConvertTo-HtmlEncoded $script:RunId)</p>
  </header>
  <main>
    <section class="executive">
      <h2>Executive Summary</h2>
      <ul>
        <li><strong>An indicative potential of $(Format-SmartFinOpsEuro $potentialMonthly -Compact) per month should be reviewed.</strong> It covers $noLicensePricedCandidates licenses E3/F3 candidates for removal and $downgradePricedCandidates licenses E3 candidates for a move to F3; $(Format-SmartFinOpsEuro $potentialAnnual -Compact) per year before contractual validation.</li>
        <li><strong>$(Format-SmartFinOpsEuro $highConfidenceMonthly -Compact) per month relates to disabled or blocked accounts.</strong> These $highConfidenceCandidates cases should be reviewed first; $reviewCandidates other decisions require business confirmation.</li>
        <li><strong>$sharedMailboxStrong disabled user mailboxes have a strong shared-mailbox conversion path; $sharedMailboxCapacityReview require a capacity review.</strong> This supports license recovery already counted above and is not an additional saving.</li>
        <li><strong>Value is not limited to savings.</strong> Available capacity, E3/F3 fit, Exchange quality, and device lifecycle must be considered together to maximize the value obtained from every euro spent.</li>
      </ul>
    </section>

    <div class="kpis" aria-label="Main indicators">
      <div class="kpi"><div class="label">Annual optimization potential</div><div class="value">$(Format-SmartFinOpsEuro $potentialAnnual -Compact)</div><div class="context">Indicative, before validation and renewal</div></div>
      <div class="kpi"><div class="label">High priority</div><div class="value">$highConfidenceCandidates</div><div class="context">Disabled or blocked accounts</div></div>
      <div class="kpi"><div class="label">Available E3 + F3 capacity</div><div class="value">$($e3Available + $f3Available)</div><div class="context">$e3Available E3 and $f3Available F3 available for reuse</div></div>
      <div class="kpi"><div class="label">Reusable capacity value</div><div class="value">$(Format-SmartFinOpsEuro $avoidanceMonthly -Compact)</div><div class="context">Monthly cost-avoidance equivalent</div></div>
    </div>

    <section>
      <h2>Potential value is concentrated in a few licenses</h2>
      <p>The ranking below shows monthly potential by SKU. It is not an instruction to remove licenses: each case must be validated with the business owner and against renewal terms.</p>
      <div class="bar-chart">$($barItems -join [Environment]::NewLine)</div>
      <div class="insight"><strong>Decision:</strong> start with disabled accounts, then review accounts with no activity for $StaleUserDays days. Reducing purchased quantities at renewal turns potential into realized savings; reassigning the same capacity creates cost avoidance.</div>
    </section>

    <section>
      <h2>Recommended decisions</h2>
      <p>Actions are ranked by financial impact, confidence, and contribution to service quality or continuity.</p>
      <div class="table-wrap"><table><thead><tr><th>Priority</th><th>Opportunity</th><th class="number">Population</th><th class="number">Indicative value</th><th>Confidence</th><th>Decision</th></tr></thead><tbody>$($actionRows -join [Environment]::NewLine)</tbody></table></div>
    </section>

    <section>
      <h2>E3, F3, or no license: requirement-based decision</h2>
      <p>The recommendation combines M365LicenseTargetPersona, account state, detailed M365 activity, Office activations, storage, and an active Intune device. It never relies on a single source.</p>
      <div class="kpis" aria-label="License decisions">
        <div class="kpi"><div class="label">E3 recommended</div><div class="value">$e3Recommended</div><div class="context">M365LicenseTargetPersona E3, desktop apps, or a technical constraint</div></div>
        <div class="kpi"><div class="label">F3 recommended</div><div class="value">$f3Recommended</div><div class="context">M365LicenseTargetPersona F3 with recent activity and no E3 constraint</div></div>
        <div class="kpi"><div class="label">No license — high priority</div><div class="value">$noLicenseHigh</div><div class="context">Disabled or blocked user accounts to validate</div></div>
        <div class="kpi"><div class="label">No license — review</div><div class="value">$noLicenseReview</div><div class="context">No recent activity observed across the correlated sources</div></div>
      </div>
      <div class="table-wrap"><table><thead><tr><th>Decision</th><th>Primary evidence</th><th>Guardrail</th></tr></thead><tbody>
        <tr><td><strong>E3</strong></td><td>M365LicenseTargetPersona E3, Microsoft 365 desktop apps, or Exchange/OneDrive storage above 2 GB.</td><td>A storage limit issue can also be addressed through cleanup or archiving; $f3LimitReviews cases require this decision.</td></tr>
        <tr><td><strong>F3</strong></td><td>M365LicenseTargetPersona F3, recent web/mobile, Teams, or email activity, no desktop app, and no storage limit issue.</td><td>$e3ToF3Candidates E3-to-F3 moves remain subject to business validation.</td></tr>
        <tr><td><strong>No license</strong></td><td>Disabled/blocked user account, or no recent activity across aggregate and detailed sources.</td><td>Validate departure, long-term leave, retention, technical accounts, shared mailboxes, and regulatory requirements.</td></tr>
        <tr><td><strong>Shared mailbox conversion</strong></td><td>Licensed UserMailbox, disabled in AD or Entra, delegated, below 50 GB, with no observed archive, hold, or service-account signal.</td><td>Strong candidates stay below 45 GB; validate application access, ownership, retention, growth, and contract terms. No automatic conversion.</td></tr>
      </tbody></table></div>
      <p class="muted">Microsoft 365 F3 provides web/mobile apps and 2 GB of storage; E3 adds desktop apps and higher storage capacity. Reference: <a href="https://www.microsoft.com/fr-fr/microsoft-365/compare-microsoft-365-enterprise-plans" target="_blank" rel="noreferrer">Microsoft France comparison</a>.</p>
    </section>
    <section>
      <h2>Maximizing value requires trade-offs, not only cost reduction</h2>
      <div class="tradeoffs">
        <div class="tradeoff cost"><h3>Cost</h3><span class="big">$(Format-SmartFinOpsEuro $potentialMonthly -Compact)/month</span><p>Potential from removing unused licenses or rightsizing E3 to F3.</p></div>
        <div class="tradeoff quality"><h3>Quality and fit</h3><span class="big">$(Get-SmartFinOpsSummaryMetricValue -SummaryRows $SummaryRows -Metric 'E3 personas missing E3')</span><p>E3 target personas without E3: balance required capabilities, user experience, and cost; never downgrade automatically.</p></div>
        <div class="tradeoff performance"><h3>Performance and lifecycle</h3><span class="big">$(Get-SmartFinOpsSummaryMetricValue -SummaryRows $SummaryRows -Metric 'AD computers requiring Windows 11 upgrade')</span><p>Devices requiring upgrade or replacement planning to avoid technical debt and lost productivity.</p></div>
      </div>
    </section>

    <section>
      <h2>Capacity and forecasting: reuse before buying</h2>
      <p>E3 and F3 are close to their current capacity. Recovering unused licenses can create headroom without immediate purchases and support adjustments at renewal.</p>
      <div class="table-wrap"><table><thead><tr><th>License</th><th class="number">Consumed / purchased</th><th class="number">Utilization</th><th class="number">Available</th><th>Status</th></tr></thead><tbody>$($capacityRows -join [Environment]::NewLine)</tbody></table></div>
    </section>

    <section>
      <h2>Questions to resolve</h2>
      <ul>
        <li>Which quantities can actually be reduced at the next renewal or true-up?</li>
        <li>Which inactive accounts represent long-term leave, technical accounts, or regulatory requirements?</li>
        <li>Which strong shared-mailbox conversion candidates are confirmed former-user mailboxes with no application dependency?</li>
        <li>Which E3/F3 matrix cases require a business, regulatory, or technical exception?</li>
        <li>What internal cost should be assigned to incidents, delayed migrations, and out-of-target devices?</li>
      </ul>
    </section>

    <section>
      <h2>Assumptions and limitations</h2>
      <ul class="assumptions">
        <li>FinOps aims to maximize the business value of technology through data-driven decisions and trade-offs between cost, quality, and speed. Reference: <a href="https://www.finops.org/introduction/what-is-finops/" target="_blank" rel="noreferrer">FinOps Foundation</a>.</li>
        <li>Pricing: $(ConvertTo-HtmlEncoded $PriceModel.Name), as of $(ConvertTo-HtmlEncoded $PriceModel.AsOfDate), $(ConvertTo-HtmlEncoded $PriceModel.TaxBasis), $(ConvertTo-HtmlEncoded $PriceModel.CommitmentBasis). These prices support prioritization; contractual prices must replace them for a budget decision.</li>
        <li>Potential savings and cost avoidance are shown separately to prevent double counting.</li>
        <li>No activity for $StaleUserDays days is a review signal, not sufficient evidence to remove a license.</li>
        <li>Detailed Exchange, OneDrive, Teams, Email, Apps, and Intune device signals complement M365_Users_Activity.csv to reduce false positives.</li>
      </ul>
      <details><summary>Indicative pricing sources</summary><ul>$($priceLinks -join [Environment]::NewLine)</ul></details>
      <details><summary>Technical appendix — data quality</summary>$(Convert-RowsToHtmlTable -Rows $technicalRows -MaxRows 10)<p class="muted">Files containing MAXITEMS are excluded from the analysis.</p></details>
    </section>
  </main>
</body>
</html>
"@

    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBlUCq16jjitMW
# kQ3W33rllgV3OT0tRc39Y3Gc1tJf8qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEU4OAhLK6zUs6PwPQyhp8FQtS/pLU6KBYwvdIvjpL33MA0GCSqG
# SIb3DQEBAQUABIIBgFIZwd10mIoSVDZcRYO9RZYf1kjjPku2qxkaAhb6fm/Z93Pm
# 7YRn1kOnu9LKycpqiRAYK8tVr3bFYFw7mlqwMEMHPbNuX4nzwlBlzH40daxCwm01
# EyjUzwAZkUGD8YvXxmsYEqAV4LUKW+xqe1d/E5/wnqUXYZbQquS1HxrVNv+5y+xh
# cUGzIWMkv9Q3sbxRVM/epiq5YJNJiILe2RS5RtfllxB1t3cN7PG6OF4grkfll7YK
# uZjX6pTE3G8xArgUUeCC+hqEU4363xGZAcg0ViBCPM+gDazzD5wN7s5tPSJFKZZ8
# Bo3tX/o80cbK8/p17mI4HpzEVmMK4DDIzv5mCQtsGaAHH7oYYzQ6xSbG0zp3009k
# 653+t7AD64zbB0mIkypd+DfBS619becd0AcKlzUjvC5HNYxqBGWgRQDh/gsPILOh
# MI3ooUE3qxkYFFi8yowmiSwp6NmhelUCVluNeJGZwrlAukOgCQQdpbK1dVW9XsSU
# Ekrg7aSzd+uwETxTqaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgxOTQ1
# MzNaMC8GCSqGSIb3DQEJBDEiBCCUdV/35hQML3b/Rs95WOSAT/XNuu1Zslneltsg
# JEnokTANBgkqhkiG9w0BAQEFAASCAgAw9ZLsO9GhjJVJ8dWik56QdCKccFKzNZpo
# A0bEbpQg//RZ3ZTPfChJyNwzqIgC8x/o6viVJLlv7cW8Y/9ShBkwFo/l8SRiBhw2
# CTdSO+ss1VcyrY5xJCGP4Oh1KjAnTr6fYo9xQXRjSz+ZbnYnnSAZPoOftAxc1jt2
# x1BvYeeGU+xaVMkJvELGMQc2EHV4p1tLdjdPc5gd+8gEnUuhzVrSymuPyesJ3eOE
# QXB/kQUh0ogOXI8GVoiMbJhs1Lai0/844RgoRM5q41YjTssm4lYI5DDFwJUWIB2/
# vFuDr4o3rsp7m+Ws6h6VXt3akVuA9uExmDZ7NfzHIeTVtkezr2A2QqW8XLfBp1bZ
# YPF4Srdon2V3R6EWa/LJ7HE6zL/hKjMIgpXCOHINp13c/bntqseY2ehiF6pWsMoI
# z6DmwwSKZ8GqC8vG9lBJWbcIwxHukuQKa1Oy7ms6MXqdK2Fz51e5r13lHaSSYbUX
# apVaC/0XyKT5+Ngc2qRpnmgY7oZ8N2xIf4KJW3cJ49XJ2I3Q4IDvcYnwbUyb3Dsu
# BtGjQ6J371hZz+wFPhWDGWJgUE4F2uo2yPdUe3VL2mbUTec4MhwspUth548iCpD2
# vQMP5BJXwGl6bAYPD2pgy/ipgwaHqqP4vy16POgWPi3QvC70/jfX985q8vBQVlKG
# xku9LsOb1Q==
# SIG # End signature block
