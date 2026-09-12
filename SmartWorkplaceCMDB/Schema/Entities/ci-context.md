# Device and organization context — V1

`-IncludeContext` creates typed user/device context outputs without changing the
core CMDB CSVs. Department and job title remain source observations. A device
may expose its resolved primary user's context with a separate source date; this
does not declare device ownership or physical location.

An optional reviewed organization reference can map Country → Entity → Site.
Without that reference and a validated governance event, no entity or site is
assigned. Department, office, hostname and network values are never converted
into location by inference.

Raw device context retains every source candidate plus enrollment, management
and activity evidence. Timezone-qualified dates are normalized to UTC;
ambiguous text remains raw and unqualified. See
`SmartWorkplaceCMDB.ci.context.json` for exact columns and statuses.
