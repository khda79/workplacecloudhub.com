# PBIP handoff — BETA 3.8.0-beta.1

1. Validate ZIP hash and extract the entire source package.
2. Run the offline validator with PowerShell 7; Node.js must be available.
3. In a separately authorized isolated test environment, open pbip/SmartWorkplaceDashboard.pbip in Power BI Desktop with PBIP/PBIR support. Configure one source root and ExpectedTenantKey.
4. Refresh and reconcile all 19 pages, 189 measures and country-filter coverage against known test inputs. Review blank/invalid/partial sources, identity errors and historical observations.
5. Test CSV/Excel/PDF exports and verify privacy/permissions before distribution.
6. Produce a PBIT/PBIX only after that refresh and visual/export review. Keep BETA. Never include cached customer data in the public source package.

Offline evidence and remaining qualification boundaries are in VALIDATION.md and KNOWN-LIMITATIONS.md. No tenant/service publication is authorized by this handoff.
