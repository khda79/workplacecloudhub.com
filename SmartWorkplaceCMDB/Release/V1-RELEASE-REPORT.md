# Smart Workplace CMDB V1 — release report

Report date: 2026-09-14

1. **Smart Workplace CMDB Version:** 1.1.3, stable.
2. **Release status:** Stable V1 release. Private environment qualification evidence is intentionally excluded from Git and the release package.
3. **Project path:** `SmartWorkplaceCMDB` in the main SmartIntune repository.
4. **Canonical PBIP path:** private canonical `PowerBI/CMDB-REPORTS/SmartWorkplaceCMDB.pbip`.
5. **Supported sources:** Active Directory, Entra ID, Intune, Microsoft 365 and Exchange Online; Azure is authorized but no unbounded collector is shipped.
6. **Collectors included:** AD domains/users/groups/computers/organizational units/direct memberships; Entra users/groups/devices/verified domains; Intune managed devices, hardware, Autopilot, detected applications, configuration/update policies and Endpoint Analytics; Microsoft 365 SKUs, assignments, user activity, SharePoint sites, Teams and membership evidence; Exchange Online mailboxes.
7. **Collection extensions added for V1:** Intune encryption, Windows upgrade eligibility and operational inventory; bounded shared Graph paging/retry handling on every Graph-native collector; transactional last-valid publication for every source, including multi-table AD.
8. **Main CMDB capabilities:** tenant isolation, typed entities, safe identity correlation, relationships, CI/source lineage, optional governed context, freshness, deterministic data quality, timestamped lifecycle console output, per-step structured logs and PowerShell transcripts with bounded retention.
9. **Power BI pages:** 10 consolidated pages: Executive Overview, Workplace Health, Transformation & Lifecycle, Licensing & Assignments, Fleet & Hardware, People & Messaging, Services & Impact, Device 360, User 360 and Group 360.
10. **Power BI model:** 39 tables, 191 measures, 25 relationships and 228 visuals in the saved canonical model.
11. **Current validation:** the release ZIP was extracted and validated independently. Tests cover reports, package closure, audit regressions, CI registry, Intune hardware/operations/analytics, orchestration, SharePoint, Active Directory pagination, transient retry, last-valid rollback and aggregate collection-summary deltas.
12. **Tests passed:** 15/15 master PowerShell checks, 98/98 Python tests and 24/24 independent PowerShell suites, including a complete 34-step synthetic pipeline and retrieval of more than 5,000 direct AD group members.
13. **Tests failed:** 0 in the current refresh-validation set.
14. **PBIR validation:** Passed with 0 errors; one non-model warning because external Fabric JSON schema 2.12 could not be fetched.
15. **DAX validation:** A private full XMLA refresh and KPI read-back completed without query errors; tenant-derived values are intentionally omitted.
16. **Desktop validation:** Passed on the exact canonical PBIP: all 10 pages and visuals persisted after close/reopen; Device 360 search, Country filtering and the default Corporate ownership selection were operator-validated; page names and visible titles contain no BETA label.
17. **Operational validation:** A completed private collection was finalized successfully; SharePoint publication and branded summary-mail delivery completed without warning. Tenant values and operational evidence remain private.
18. **Known limitations:** no Power BI RLS or Power BI Service publication claim; no Azure resource inventory, authoritative Business Services/criticality, recursive dependency graph, application lifecycle catalog or financial data.
19. **V2 backlog:** documented in `Release/V2-BACKLOG.md`, including a future allowlist-scoped Azure Workplace collector.
20. **Documentation updated:** README, Power BI/DSI/hardware contracts, source/schema inventory, operations, release notes and V2 backlog.
21. **Files modified:** scoped exclusively to `SmartWorkplaceCMDB`; exact Git status is retained as release evidence.
22. **Git status:** publication is performed only through an explicitly approved, source-only Git delta; tenant data and private artifacts remain outside that authority.
23. **External publication status:** the configured SharePoint upload and summary-mail paths were privately validated; no tenant data, Power BI artifact, Power BI Service item, Azure resource, or website content is included in this Git publication.
24. **Private-artifact boundary:** runtime data, local configuration, PBIP/PBIX files, copied ReportData, semantic-model caches and screenshots remain excluded by `.gitignore` and the release allowlist.

`Smart Workplace CMDB V1 is ready for release`
