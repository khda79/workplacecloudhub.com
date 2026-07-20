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

function Get-SmartFinOpsDecimalSum {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Property
    )

    [decimal]$total = 0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $valueProperty = $row.PSObject.Properties[$Property]
        if ($null -eq $valueProperty -or $null -eq $valueProperty.Value -or [string]::IsNullOrWhiteSpace([string]$valueProperty.Value)) { continue }
        try { $total += [decimal]$valueProperty.Value } catch { continue }
    }
    return $total
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
    foreach ($group in ($noLicenseDecisions | Group-Object -Property CurrentBaseSku, DecisionClass)) {
        $sample = $group.Group | Select-Object -First 1
        $highConfidence = @($group.Group | Where-Object { $_.DecisionClass -eq 'Recommended' })
        $review = @($group.Group | Where-Object { $_.DecisionClass -eq 'Review' })
        $monthlyValue = [math]::Round((Get-SmartFinOpsDecimalSum -Rows $group.Group -Property 'IndicativeMonthlyDifferenceEUR'), 2)
        $opportunityClass = if ($highConfidence.Count -gt 0) { 'Recommended' } else { 'Review' }
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = 1
            ValuePillar = 'Potential savings'
            OpportunityClass = $opportunityClass
            Opportunity = if ($opportunityClass -eq 'Recommended') { "Remove $($sample.CurrentBaseLicense) licenses from disabled or blocked accounts" } else { "Review $($sample.CurrentBaseLicense) licenses with no recent activity" }
            SkuPartNumber = [string]$sample.CurrentBaseSku
            Population = $group.Count
            HighConfidencePopulation = $highConfidence.Count
            ReviewPopulation = $review.Count
            MonthlyValueEUR = $monthlyValue
            AnnualValueEUR = [math]::Round([decimal]($monthlyValue * 12), 2)
            HighConfidenceMonthlyValueEUR = if ($opportunityClass -eq 'Recommended') { $monthlyValue } else { 0 }
            Confidence = if ($opportunityClass -eq 'Recommended') { 'High' } else { 'Medium' }
            Decision = 'Confirm departure, retention, ownership, and all detailed activity before removing the license.'
            FinancialTreatment = 'Estimated savings potential. Realization depends on operational execution and the ability to reduce purchased quantities.'
        }) | Out-Null
    }
    $e3ToF3Decisions = @($UserDecisionRows | Where-Object {
        $_.CurrentBaseSku -eq 'SPE_E3' -and
        $_.RecommendedLicense -eq 'Potential M365 F3 - Frontline eligibility required' -and
        $null -ne $_.IndicativeMonthlyDifferenceEUR -and
        [decimal]$_.IndicativeMonthlyDifferenceEUR -gt 0
    })
    if ($e3ToF3Decisions.Count -gt 0) {
        $monthlyValue = [math]::Round((Get-SmartFinOpsDecimalSum -Rows $e3ToF3Decisions -Property 'IndicativeMonthlyDifferenceEUR'), 2)
        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            Priority = 1
            ValuePillar = 'Potential savings'
            OpportunityClass = 'Conditional'
            Opportunity = 'Review E3-to-F3 candidates identified by M365LicenseTargetPersona'
            SkuPartNumber = 'SPE_E3->SPE_F1'
            Population = $e3ToF3Decisions.Count
            HighConfidencePopulation = 0
            ReviewPopulation = $e3ToF3Decisions.Count
            MonthlyValueEUR = $monthlyValue
            AnnualValueEUR = [math]::Round([decimal]($monthlyValue * 12), 2)
            HighConfidenceMonthlyValueEUR = 0
            Confidence = 'Medium'
            Decision = 'Confirm documented Frontline eligibility before execution; desktop and 2 GB Exchange/OneDrive technical guardrails are already checked by SmartFinOps.'
            FinancialTreatment = 'Conditional estimated savings potential; no downgrade is executed automatically.'
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
            OpportunityClass = 'Recommended'
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
        [pscustomobject]@{ Priority = 3; Pillar = 'Value and fit'; Metric = 'F3 to E3 capability reviews'; Opportunity = 'Review F3 users whose target persona is E3'; Decision = 'Validate capability, quality, and user-experience requirements before any upgrade; persona alone is not an automatic action.'; Confidence = 'Medium' },
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
            OpportunityClass = 'Review'
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
        [Parameter(Mandatory)][AllowNull()]$PriceModel,
        [int]$StaleUserDays = 90
    )

    $potentialRows = @($ValueRows | Where-Object { $_.ValuePillar -eq 'Potential savings' })
    $recommendedRows = @($potentialRows | Where-Object { $_.OpportunityClass -eq 'Recommended' })
    $conditionalRows = @($potentialRows | Where-Object { $_.OpportunityClass -eq 'Conditional' })
    $reviewFinancialRows = @($potentialRows | Where-Object { $_.OpportunityClass -eq 'Review' })
    $potentialMonthly = (Get-SmartFinOpsDecimalSum -Rows $potentialRows -Property 'MonthlyValueEUR')
    $potentialAnnual = (Get-SmartFinOpsDecimalSum -Rows $potentialRows -Property 'AnnualValueEUR')
    $recommendedMonthly = (Get-SmartFinOpsDecimalSum -Rows $recommendedRows -Property 'MonthlyValueEUR')
    $conditionalMonthly = (Get-SmartFinOpsDecimalSum -Rows $conditionalRows -Property 'MonthlyValueEUR')
    $reviewMonthly = (Get-SmartFinOpsDecimalSum -Rows $reviewFinancialRows -Property 'MonthlyValueEUR')
    $pricedCandidates = [int](Get-SmartFinOpsDecimalSum -Rows $potentialRows -Property 'Population')
    $recommendedCandidates = [int](Get-SmartFinOpsDecimalSum -Rows $recommendedRows -Property 'Population')
    $conditionalCandidates = [int](Get-SmartFinOpsDecimalSum -Rows $conditionalRows -Property 'Population')
    $reviewCandidates = [int](Get-SmartFinOpsDecimalSum -Rows $reviewFinancialRows -Property 'Population')
    $staleSources = @($DataQualityRows | Where-Object { $_.FreshnessStatus -eq 'Stale' }).Count
    $invalidSources = @($DataQualityRows | Where-Object { $_.ContractStatus -eq 'Invalid' -or $_.Status -eq 'Error' }).Count
    $maxBarValue = if ($potentialRows.Count -gt 0) { [decimal](($potentialRows | Measure-Object -Property MonthlyValueEUR -Maximum).Maximum) } else { [decimal]0 }
    $e3ToF3Candidates = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'Potential M365 F3 - Frontline eligibility required' }).Count
    $e3ToF3ActivityReviews = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'Potential M365 F3 - activity and eligibility review' }).Count
    $e3ToF3Blockers = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'Keep M365 E3 - F3 technical blocker' }).Count
    $noLicenseHigh = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'No license - candidate' }).Count
    $noLicenseReview = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'No license - review' }).Count
    $personaNoneActiveConflicts = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'Target persona conflict - active user review' }).Count
    $f3ToE3Reviews = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'M365 E3 capability review' }).Count
    $f3TechnicalConflicts = @($UserDecisionRows | Where-Object { $_.RecommendedLicense -eq 'M365 F3 technical conflict - review' }).Count
    $unusedE3F3Rows = @($UserDecisionRows | Where-Object { $_.IsUnusedE3F3License -eq $true })
    $possiblyUnusedE3F3Rows = @($UserDecisionRows | Where-Object { $_.IsPossiblyUnusedE3F3License -eq $true })
    $unusedOrPossiblyUnusedE3F3Rows = @($unusedE3F3Rows + $possiblyUnusedE3F3Rows)
    $unusedE3F3 = $unusedE3F3Rows.Count
    $possiblyUnusedE3F3 = $possiblyUnusedE3F3Rows.Count
    $unusedOrPossiblyUnusedE3F3 = $unusedOrPossiblyUnusedE3F3Rows.Count
    $unusedOrPossiblyUnusedE3 = @($unusedOrPossiblyUnusedE3F3Rows | Where-Object { $_.CurrentBaseSku -eq 'SPE_E3' }).Count
    $unusedOrPossiblyUnusedF3 = @($unusedOrPossiblyUnusedE3F3Rows | Where-Object { $_.CurrentBaseSku -eq 'SPE_F1' }).Count
    $e3WithoutObservedE3Capabilities = @($UserDecisionRows | Where-Object { $_.IsE3WithoutObservedE3Capabilities -eq $true }).Count

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
        $actionRows.Add("<tr><td><span class='priority p$($row.Priority)'>P$($row.Priority)</span></td><td><strong>$(ConvertTo-HtmlEncoded $row.Opportunity)</strong><div class='muted'>$(ConvertTo-HtmlEncoded $row.ValuePillar)</div></td><td>$(ConvertTo-HtmlEncoded $row.OpportunityClass)</td><td class='number'>$($row.Population)</td><td class='number'>$(ConvertTo-HtmlEncoded $value)</td><td>$(ConvertTo-HtmlEncoded $row.Confidence)</td><td>$(ConvertTo-HtmlEncoded $row.Decision)</td></tr>") | Out-Null
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
    .license-banner { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px; margin-top:24px; }
    .license-signal { background:var(--paper); border:1px solid var(--line); border-radius:16px; padding:20px 22px; box-shadow:var(--shadow); }
    .license-signal.unused { border-top:5px solid var(--orange); }
    .license-signal.e3-fit { border-top:5px solid var(--gold); }
    .license-signal .signal-label { color:var(--muted); font-size:13px; font-weight:650; }
    .license-signal .signal-value { display:block; font-size:clamp(34px,5vw,52px); font-weight:800; line-height:1; margin:8px 0 10px; }
    .license-signal .signal-context { color:var(--muted); font-size:13px; }
    .license-banner-note { grid-column:1/-1; color:var(--muted); font-size:12px; margin:0 2px; }
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
    @media (max-width:560px) { header,main { padding-left:16px; padding-right:16px; } section { padding:20px; } .kpis,.tradeoffs,.license-banner { grid-template-columns:1fr; } }
    @media print { body { background:white; } header,main { max-width:none; } section,.kpi { box-shadow:none; break-inside:avoid; } }
  </style>
</head>
<body>
  <header>
    <div class="eyebrow">SmartFinOps Workplace</div>
    <h1>$(ConvertTo-HtmlEncoded $ReportTitle)</h1>
    <p class="meta">Tenant $(ConvertTo-HtmlEncoded $Tenant) · Analysis generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') · Indicative France pricing, excluding VAT · Run $(ConvertTo-HtmlEncoded $script:RunId)</p>
    <div class="license-banner" aria-label="License utilization signals">
      <div class="license-signal unused">
        <div class="signal-label">Unused or possibly unused E3/F3 licenses</div>
        <span class="signal-value">$unusedOrPossiblyUnusedE3F3</span>
        <div class="signal-context">$unusedE3F3 unused + $possiblyUnusedE3F3 possibly unused · $unusedOrPossiblyUnusedE3 E3 + $unusedOrPossiblyUnusedF3 F3</div>
      </div>
      <div class="license-signal e3-fit">
        <div class="signal-label">E3 licenses without observed E3 capability usage</div>
        <span class="signal-value">$e3WithoutObservedE3Capabilities</span>
        <div class="signal-context">No observed Microsoft 365 Apps desktop activation · measured mailbox below 100 GB</div>
      </div>
      <p class="license-banner-note">Unused means no activity observed for $StaleUserDays days, regardless of mailbox size or measurement. Possibly unused means technical presence only, no observed M365 service usage, and a measured mailbox at or below 100 MB. The E3 capability population can overlap with these signals and must not be added to them.</p>
    </div>
  </header>
  <main>
    <section class="executive">
      <h2>Executive Summary</h2>
      <ul>
        <li><strong>An estimated optimization potential of $(Format-SmartFinOpsEuro $potentialMonthly -Compact) per month has been identified across $pricedCandidates licenses.</strong> This is $(Format-SmartFinOpsEuro $potentialAnnual -Compact) per year before contract and execution checks.</li>
        <li><strong>$recommendedCandidates high-confidence opportunities represent $(Format-SmartFinOpsEuro $recommendedMonthly -Compact) per month.</strong> They concern disabled or blocked accounts whose target persona is None.</li>
        <li><strong>$conditionalCandidates conditional E3-to-F3 opportunities represent $(Format-SmartFinOpsEuro $conditionalMonthly -Compact) per month.</strong> Frontline eligibility must be confirmed before execution.</li>
        <li><strong>$reviewCandidates no-license opportunities representing $(Format-SmartFinOpsEuro $reviewMonthly -Compact) per month require further evidence.</strong> Value also includes capacity reuse, service quality, Exchange, and device lifecycle decisions.</li>
      </ul>
    </section>

    <div class="kpis" aria-label="Main indicators">
      <div class="kpi"><div class="label">Annual optimization potential</div><div class="value">$(Format-SmartFinOpsEuro $potentialAnnual -Compact)</div><div class="context">Indicative, before contract and execution checks</div></div>
      <div class="kpi"><div class="label">Recommended</div><div class="value">$recommendedCandidates</div><div class="context">$(Format-SmartFinOpsEuro $recommendedMonthly -Compact) per month, high confidence</div></div>
      <div class="kpi"><div class="label">Conditional</div><div class="value">$conditionalCandidates</div><div class="context">$(Format-SmartFinOpsEuro $conditionalMonthly -Compact) per month, prerequisite required</div></div>
      <div class="kpi"><div class="label">Review</div><div class="value">$reviewCandidates</div><div class="context">$(Format-SmartFinOpsEuro $reviewMonthly -Compact) per month, more evidence required</div></div>
    </div>

    <section>
      <h2>Estimated value by recommendation class</h2>
      <p>The ranking below separates recommendations generated automatically by SmartFinOps from conditional opportunities and cases requiring further review.</p>
      <div class="bar-chart">$($barItems -join [Environment]::NewLine)</div>
      <div class="insight"><strong>How to read this:</strong> Recommended means the available evidence is strong, Conditional means a prerequisite remains, and Review means the evidence is insufficient for a direct recommendation. Realized savings still depend on execution and contract quantities.</div>
    </section>

    <section>
      <h2>Recommended decisions</h2>
      <p>Actions are ranked by financial impact, confidence, and contribution to service quality or continuity.</p>
      <div class="table-wrap"><table><thead><tr><th>Priority</th><th>Opportunity</th><th>Class</th><th class="number">Population</th><th class="number">Indicative value</th><th>Confidence</th><th>Decision</th></tr></thead><tbody>$($actionRows -join [Environment]::NewLine)</tbody></table></div>
    </section>

    <section>
      <h2>E3, F3, or no license: persona-led, evidence-guarded decisions</h2>
      <p>M365LicenseTargetPersona identifies the population to review. Account state, detailed M365 activity, Office activations, storage, and active Intune devices then confirm, block, or qualify the opportunity. No license is changed automatically.</p>
      <div class="kpis" aria-label="License decisions">
        <div class="kpi"><div class="label">E3 to F3 — conditional</div><div class="value">$e3ToF3Candidates</div><div class="context">Persona F3 and no observed technical blocker; Frontline eligibility is mandatory before execution</div></div>
        <div class="kpi"><div class="label">E3 to F3 — blocked or incomplete</div><div class="value">$($e3ToF3Blockers + $e3ToF3ActivityReviews)</div><div class="context">$e3ToF3Blockers technical blockers and $e3ToF3ActivityReviews activity reviews</div></div>
        <div class="kpi"><div class="label">No license — high priority</div><div class="value">$noLicenseHigh</div><div class="context">Persona None plus disabled or blocked account</div></div>
        <div class="kpi"><div class="label">No license — review</div><div class="value">$noLicenseReview</div><div class="context">No recent activity; $personaNoneActiveConflicts active persona conflicts stay separate</div></div>
      </div>
      <div class="table-wrap"><table><thead><tr><th>Decision</th><th>Primary evidence</th><th>Guardrail</th></tr></thead><tbody>
        <tr><td><strong>E3 → F3</strong></td><td>Current E3, M365LicenseTargetPersona F3, recent activity, no desktop activation, and Exchange/OneDrive storage within 2 GB.</td><td>Documented Frontline eligibility is mandatory. $e3ToF3Candidates cases are conditional opportunities, never automatic downgrades.</td></tr>
        <tr><td><strong>E3/F3 → no license</strong></td><td>Current E3 or F3 with M365LicenseTargetPersona None, corroborated by disabled/blocked state or no recent activity.</td><td>Active conflicts are excluded. Validate departure, leave, ownership, retention, technical accounts, shared mailboxes, and regulatory requirements.</td></tr>
        <tr><td><strong>F3 → E3 review</strong></td><td>Current F3 with M365LicenseTargetPersona E3, optionally reinforced by a desktop or 2 GB storage conflict.</td><td>$f3ToE3Reviews cases are quality and capability reviews, not assumed upgrades; $f3TechnicalConflicts other F3 technical conflicts also require action.</td></tr>
        <tr><td><strong>Shared mailbox conversion</strong></td><td>Licensed UserMailbox, disabled in AD or Entra, delegated, below 50 GB, with no observed archive, hold, or service-account signal.</td><td>Strong candidates stay below 45 GB; validate application access, ownership, retention, growth, and contract terms. No automatic conversion.</td></tr>
      </tbody></table></div>
      <p class="muted">Microsoft 365 F3 provides web/mobile apps with a 2 GB Exchange mailbox and a 2 GB OneDrive limit in this decision model; E3 adds desktop apps and higher storage capacity. Reference: <a href="https://www.microsoft.com/fr-fr/microsoft-365/compare-microsoft-365-enterprise-plans" target="_blank" rel="noreferrer">Microsoft France comparison</a>.</p>
    </section>
    <section>
      <h2>Maximizing value requires trade-offs, not only cost reduction</h2>
      <div class="tradeoffs">
        <div class="tradeoff cost"><h3>Cost</h3><span class="big">$(Format-SmartFinOpsEuro $potentialMonthly -Compact)/month</span><p>Potential from removing unused licenses or rightsizing E3 to F3.</p></div>
        <div class="tradeoff quality"><h3>Quality and fit</h3><span class="big">$f3ToE3Reviews</span><p>F3 users whose target persona is E3: validate capability, user experience, and service risk before deciding.</p></div>
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
        <li>Which E3-to-F3 candidates have documented Frontline eligibility, and which persona mismatches require a business, regulatory, or technical exception?</li>
        <li>What internal cost should be assigned to incidents, delayed migrations, and out-of-target devices?</li>
      </ul>
    </section>

    <section>
      <h2>Assumptions and limitations</h2>
      <ul class="assumptions">
        <li>FinOps aims to maximize the business value of technology through data-driven decisions and trade-offs between cost, quality, and speed. Reference: <a href="https://www.finops.org/introduction/what-is-finops/" target="_blank" rel="noreferrer">FinOps Foundation</a>.</li>
        <li>Pricing: $(ConvertTo-HtmlEncoded $PriceModel.Name), as of $(ConvertTo-HtmlEncoded $PriceModel.AsOfDate), $(ConvertTo-HtmlEncoded $PriceModel.TaxBasis), $(ConvertTo-HtmlEncoded $PriceModel.CommitmentBasis). These prices support prioritization; contractual prices must replace them for a budget decision.</li>
        <li>Potential savings and cost avoidance are shown separately to prevent double counting.</li>
        <li>Decision classes are generated automatically from persona, account state, activity, technical usage, storage, and device evidence.</li>
        <li>Recommended does not mean automatically executed. Contract quantities, renewal timing, service ownership, retention, and implementation must still be confirmed.</li>
        <li>M365LicenseTargetPersona identifies the population to review. It never changes a license automatically.</li>
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1+6dhATzRW4I+
# YPXxvuOjbrN5RcyWSP+hLANt67LLmaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINhUaVkNfeiRFCM9sAyb3xN20s37ZVp2YrqAhYEtRh4hMA0GCSqG
# SIb3DQEBAQUABIIBgHonY6t9W/DbRW7kpOy4GxVDeh3f4/NVuO8vorR2IHLa7vgK
# PuFQw1dG+L0S44U4w/qckseXxhf40/iHKZAEXwwx2Y2aHLLot1/ObmUgv95S/7j3
# R+6t2LOkvnZ2ulRTjIL2HCKxbE/9GOa+kE7eUw4HxpWKJnc6SvKtjerM31+jxgwB
# d3WAWT+YlYmCF1ADEe5eVsN90wyEUhOaR7/+sRcnhlglLmZVptiPtVnKgiiP4E+0
# HwYmMo8CA+AkFFMBirhfBrRz6xoD8ANT4Z+AVPoUFE7iQD7FAqBia0iqAM8uSQ9/
# OZt+QpUGwkxcmkvGAXQekPzQVOcj0sR4XFfYGrUSyAQd0BZZvy8Kqknda/zrAcA7
# wMpN5efE0c9o7EiO8EJ/dHVCH48wtoTMDsMlSMFA9DLWi0wpBGFVA9pkwh96Z60W
# pdlI33Usp6unyA3GP89w4eRhavq+rtqVOzcbuuILKjStEMoslm2BACoDf1iy0upD
# ttKQz/qKVcJdAXYiSaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAxMTI1
# MDdaMC8GCSqGSIb3DQEJBDEiBCCmLVpC5N/plD2e0FJtVXCDYwzj+PYQG19dVbLk
# /0yi7TANBgkqhkiG9w0BAQEFAASCAgB3mmHgSRMAj1TkxEQuOY1ta3fH0rP+5h3Z
# snQL7EpxAbI7LTbqrLXNEH3KlYNoZsfJy8eSSLsFWJT49xdfkM085ZRKhX3sWpjP
# HM/uoeeHr/8/Shztr4OYtXgJy+eiCpB18rjoL65egkETCDHZP34oe0Q/yPxFnOV0
# 7340UzVaZdF1x55R8yWGOBGAle+w+ookq77Zf/nX5cGBl9DDf3TlfLkGEaeA6gJr
# psvwj50tEZ8cokHeoCIWERMCP/wXYHIXkdgzrEKSLzb332NfMvbiHhnbyQejSfyK
# GVygy0IpjSlMQdAUquM0gzbkhiTvEdlA0UcKEPNjhtBgN+OR52qFYv9kv5GjeLfm
# PuopLmXfCblhp8MkkKxuYuIh4WouuflWAQg+KbunHQhqbr85VutGh/1UzgKd8PMa
# EIQ/rTVkhbv6o16YxgPMJC8YJKFyaKuG1MlLVB4BGqGW7lEFKTt/7cUwV1baIhvf
# LZHBynV66ZO7x+o2jeQXf+vq5psJeDx/p0asY5QtgllOt7QHN/34HSLmB7K2vKZe
# Elia0lrc15QqaVYgHOlvG0kwtTB+rQQaZ7OT2H0eUdlsUIhsXLOhrK10qmYeVkcI
# KOGRjbgOj8HslpCBrJ96/t4QqvZWdH+Z3iwtuZAvQ34L6W4l0Fy3Uco2LIDF7bTP
# vl0pJSrxbQ==
# SIG # End signature block
