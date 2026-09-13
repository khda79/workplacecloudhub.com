# Smart Workplace CMDB — DSI cockpit contract (V1)

`report_cockpit.py` maintains the reviewed PBIR definition in place. It does
not rebuild the semantic model. The approved brownfield transformation reduces
the report to seven decision pages plus three visible 360 detail and
drill-through pages. Stable page identifiers and drill-through bindings are retained; six
redundant presentation pages are retired only after a targeted backup.

## Navigation and page jobs

1. Executive Overview: inventory, Microsoft 365 license synthesis, compliance,
   management, ownership, Windows 11, device type and country ratios.
2. Workplace Health: compliance, Windows Update alerts and CMDB-quality exceptions.
3. Transformation & Lifecycle: observed OS/management state, Autopilot,
   Endpoint Analytics and source readiness, without inventing migration or lifecycle progress.
4. Licensing & Assignments: SKU-level capacity, assignment and service-plan evidence.
5. Fleet & Hardware: reconciled asset inventory and hardware coverage.
6. People & Messaging: accounts, sign-in-derived activity, assignments, device links and mailboxes.
7. Services & Impact: bounded current relationships and explicit absence of an
   authoritative business-service/application dependency graph.

Device 360, User 360 and Group 360 remain visible physical report pages and
valid drill-through targets. Hardware detail is integrated into Device 360.
The retired page directories are `quality`, `transformation`, `impact`,
`mailboxes`, `hardwarefleet` and `hardware`; their high-value content is carried
into the consolidated pages above.

The report opens on Executive Overview. Bottom tabs expose the complete
ten-page path, including the three unchanged 360 IDs.

## Canonical design contract

```yaml
Design Brief:
  generated_by: powerbi-report-design
  contract_version: 1
  mode: brownfield
  design_identity:
    current_tone: light technical cockpit with repeated domain pages
    current_signature: off-white canvas, white rounded evidence panels, dark-navy table headers
    tone: calm executive Workplace cockpit with progressive disclosure
    signature: one decision question per page, with focused evidence available on visible 360 pages
  canvas: {width: 1280, height: 900, margin: 24, gutter: 16, snap: 8}
  pages:
    - {name: Executive Overview, role: landing, archetype: Executive, layout_variant: KPI-and-evidence, variant_rationale: "The DSI needs a ten-second tenant snapshot before domain analysis.", regions: [header, filters, kpis, distributions, licensing, footer]}
    - {name: Workplace Health, role: detail, archetype: Operational, layout_variant: exceptions-and-evidence, variant_rationale: "Compliance states and quality findings are exception populations with a shared review cadence.", regions: [header, four_filters, exception_kpis, state_charts, findings_table, footer]}
    - {name: Transformation & Lifecycle, role: detail, archetype: Narrative, layout_variant: readiness-evidence, variant_rationale: "The model supports estate/readiness evidence but not migration progress or lifecycle age.", regions: [header, filters, coverage_kpis, estate_charts, source_table, footer]}
    - {name: Licensing & Assignments, role: detail, archetype: Comparative, layout_variant: sku-comparison, variant_rationale: "F1, F3, E3, E5 and Copilot must remain individually comparable.", regions: [header, filters, sku_kpis, assignment_charts, assignment_table, footer]}
    - {name: Fleet & Hardware, role: detail, archetype: Analytical, layout_variant: filter-rail-with-evidence, variant_rationale: "Device and hardware attributes share the reconciled device grain and a common investigation path.", regions: [header, filters, coverage_kpis, state_and_manufacturer_charts, fleet_table, footer]}
    - {name: People & Messaging, role: detail, archetype: Analytical, layout_variant: paired-domains, variant_rationale: "Accounts and mailboxes are linked but require separate evidence tables.", regions: [header, filters, relationship_kpis, user_and_mailbox_charts, paired_tables, footer]}
    - {name: Services & Impact, role: detail, archetype: Analytical, layout_variant: bounded-paths, variant_rationale: "Only user-device and license paths are authoritative; service readiness stays an explicit limitation.", regions: [header, filters, relationship_kpis, path_charts, evidence_table, footer]}
    - {name: Device 360, role: detail_and_drillthrough, archetype: Analytical, layout_variant: two-column-evidence, variant_rationale: "Identity/activity and hardware evidence must remain side by side for one selected device.", regions: [header, selector, identity, source_and_hardware, activity_and_hardware_dates, findings, footer]}
    - {name: User 360, role: detail_and_drillthrough, archetype: Analytical, layout_variant: existing-detail, variant_rationale: "Focused account evidence remains directly accessible and drill-through capable.", regions: [header, selector, account, relationships, assignments, findings, footer]}
    - {name: Group 360, role: detail_and_drillthrough, archetype: Analytical, layout_variant: existing-detail, variant_rationale: "Focused group evidence remains directly accessible and drill-through capable.", regions: [header, selector, group, assignment_paths, findings, footer]}
  interaction_pattern:
    drill_targets: [Device 360, User 360, Group 360]
    cross_filter_rules: preserve existing filter behavior and exact drill-through bindings
  accessibility:
    alt_text_strategy: preserve existing visual descriptions; titles state question and caveat
    contrast_notes: preserve validated navy, teal, white and off-white pairs
  theme:
    base: existing Smart Workplace light theme preserved
    user_overrides: no palette, typography or semantic-color replacement
```

## Measure boundaries

- Managed/unmanaged and corporate/personal ratios classify only explicit source
  states; unknown values remain separate.
- Windows 11 ratio uses the filtered Windows device population.
- Country ratios use reported user country and primary-user-derived device
  country only; no location is inferred from hostname, IP address or office.
- Microsoft 365 F1, F3, E3, E5 and Copilot are distinct SKU families; assignment
  coverage is not utilization or activity.
- Compliance, user-link, mailbox-link, source-coverage and hardware-gap rates
  each retain their own documented population and are never combined into a
  composite health score.

## Deliberately unavailable metrics

V1 has no authoritative business-service/application graph, service owner,
criticality, acquisition date, warranty, support end date, cost, savings,
migration plan, VDI or Golden Image source. These values stay
unavailable; zero is never substituted.

The current Autopilot and Endpoint Analytics tables are collected evidence,
not proof of deployment readiness or user experience. Windows Update alerts
are records, not a distinct-device count unless the corresponding explicit
device measure is used. Hardware remains the last strictly validated snapshot
when a fresh hardware state-evidence file is unavailable; its coverage and
source date stay visible instead of being silently rebuilt from unqualified raw data.

The report is a private V1 artifact. It does not by itself qualify tenant,
gateway, endpoint or production behavior and must not be published without
separate authority.
