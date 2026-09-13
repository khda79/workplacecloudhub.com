# Smart Workplace CMDB V1 — release report

Report date: 2026-09-12

1. **Smart Workplace CMDB Version:** 1.1.3, stable patch candidate.
2. **Release status:** Stable V1 code candidate. Private environment qualification evidence is intentionally excluded from Git and the release package.
3. **Project path:** `SmartWorkplaceCMDB` in the main SmartIntune repository.
4. **Canonical PBIP path:** private canonical `PowerBI/CMDB-REPORTS/SmartWorkplaceCMDB.pbip`.
5. **Supported sources:** Active Directory, Entra ID, Intune, Microsoft 365 and Exchange Online; Azure is authorized but no unbounded collector is shipped.
6. **Collectors included:** AD domains/users/groups/computers/organizational units/direct memberships; Entra users/groups/devices; Intune managed devices/hardware; Microsoft 365 SKUs/assignments; Exchange Online mailboxes.
7. **Collection extensions added for V1:** Intune encryption and list-only hardware; bounded shared Graph paging/retry handling on every Graph-native collector; transactional last-valid publication for every source, including multi-table AD.
8. **Main CMDB capabilities:** tenant isolation, typed entities, safe identity correlation, relationships, CI/source lineage, optional governed context, freshness, deterministic data quality, timestamped lifecycle console output, per-step structured logs and PowerShell transcripts with bounded retention.
9. **Power BI pages:** 10 consolidated pages: Executive Overview, Workplace Health, Transformation & Lifecycle, Licensing & Assignments, Fleet & Hardware, People & Messaging, Services & Impact, Device 360, User 360 and Group 360.
10. **Power BI model:** 38 tables, 111 measures and 34 relationships in the canonical model.
11. **Current refresh validation:** Python report tests and the complete PowerShell test set cover reports, audit regressions, CI registry, Intune hardware, orchestration, SharePoint, Active Directory pagination, transient retry, last-valid rollback and aggregate collection-summary deltas.
12. **Tests passed:** all current offline validation tests passed, including synthetic retrieval of more than 5,000 direct AD group members and rejection of partial AD publication.
13. **Tests failed:** 0 in the current refresh-validation set.
14. **PBIR validation:** Passed with 0 errors; one non-model warning because external Fabric JSON schema 2.12 could not be fetched.
15. **DAX validation:** A private full XMLA refresh and KPI read-back completed without query errors; tenant-derived values are intentionally omitted.
16. **Desktop validation:** Passed on the exact canonical PBIP: all 10 pages were captured after refresh, no capture failed, page names and visible titles contain no BETA label, and Desktop confirms `hasUnsavedChanges: false` after the operator save.
17. **Non-blocking warnings:** native click/scroll/cross-filter/drill-through behavior is not proved by screenshots; completion of the patched Active Directory live collection and a real summary-mail delivery remain pending on the domain-reachable host; PBIR has no RLS.
18. **Known limitations:** no Autopilot, Intune app/policy inventory, cloud membership expansion, effective plan catalog, AD-Entra reconciliation, authoritative Entity/Site, Business Services/criticality, advanced lifecycle, recursive graph or financial data.
19. **V2 backlog:** documented in `Release/V2-BACKLOG.md`, including a future allowlist-scoped Azure Workplace collector.
20. **Documentation updated:** README, Power BI/DSI/hardware contracts, source/schema inventory, operations, release notes, audit and V2 backlog.
21. **Files modified:** scoped exclusively to `SmartWorkplaceCMDB`; exact Git status is retained as release evidence.
22. **Git status:** publication is performed only through a separately reviewed, source-only Git delta; this report does not authorize tenant data or artifact publication.
23. **External publication status:** the configured SharePoint upload path was privately validated; no tenant data, Power BI artifact, Power BI Service item, Azure resource, or website content is included in this Git publication.
24. **Private-artifact boundary:** runtime data, local configuration, PBIP/PBIX files, copied ReportData, semantic-model caches and screenshots remain excluded by `.gitignore` and the release allowlist.

`Smart Workplace CMDB V1 is ready for release`
