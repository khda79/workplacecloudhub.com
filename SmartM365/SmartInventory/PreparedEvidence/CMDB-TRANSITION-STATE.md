# CMDB publication state

The current CMDB cohort remains exactly 46 CSVs plus current.json.txt. Hash,
identity, completeness and acquisition-age checks are unchanged.

CMDB publication directs the shared SharePoint JSON transition guard and durable
recovery journal to the configured LOG-ALL/Publication/CMDB/SharePointTransition
directory. No CSV or manifest copy is required. Other inventory uploads retain
their existing default behavior when the optional state directory is omitted.

Before real preparation or publication, legacy current.json.txt transition
artifacts are relocated to that directory under the CMDB publication lock.
The manifest must have the expected owner and full tenant identity. Each legacy
file must be unlinked and exclusively available for relocation; the legacy lock
must be empty. Files are preserved, not deleted. Existing destinations are never
overwritten. Active locks, unknown cohort files, identity mismatches or conflicting
recovery state fail closed. ValidateOnly does not relocate files.

Run Test-SmartM365CmdbTransitionStateOffline.ps1 for synthetic fault checks and
two successive real generation/mock publication cycles. These tests do not
qualify live SharePoint publication, synchronization or Power BI refresh.
