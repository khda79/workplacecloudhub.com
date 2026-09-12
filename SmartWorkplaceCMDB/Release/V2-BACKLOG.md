# Smart Workplace CMDB — V2 backlog

V1 is frozen. The following items are intentionally excluded until their data,
scope, permissions and operating cost are reviewed.

| Item | Expected value | Main dependency | Priority |
|---|---|---|---|
| Autopilot identities, group tags and profiles | Transformation readiness | Intune Graph contract | High |
| Intune applications, policies and assignments | Deployment governance | New CI and relationship grains | High |
| Entra group owners and memberships | Identity and assignment graph | Batching and scope rules | High |
| AD to Entra identity reconciliation | Hybrid identity | Reviewed immutable-key policy | High |
| Entity and Site hierarchy | Organization reporting | Authoritative mapping source | High |
| Scoped Azure Workplace inventory | Workplace cloud dependencies | Reviewed subscription/RG/tag allowlist | High |
| Effective named service plans | Licensing detail | Microsoft SKU/plan catalog semantics | Medium |
| Advanced lifecycle and vendor EOL | Lifecycle planning | Governed external lifecycle source | Medium |
| Business Services and criticality | Service management | Authoritative catalog and owners | Medium |
| Recursive impact graph | Change and incident analysis | Richer validated relationships | Medium |
| Financial cost and savings | FinOps | Procurement/finance source outside V1 | Low |
| Enterprise application catalog | Architecture | Additional authoritative sources | Low |

No V2 row is implicitly authorized by the V1 release.
