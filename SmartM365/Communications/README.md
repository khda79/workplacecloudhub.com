# SmartM365 Communications

Operational user communication scripts for Microsoft 365 scenarios.

This area is separate from `SmartInventory` because these scripts can send user-facing email. Inventory scripts collect and publish data; Communications scripts use data, templates, and campaign configuration to notify users.

## Organization

- `ExchangeUserNotifications/`: Exchange migration, archive mailbox, and mailbox reduction notification campaigns.

## Configuration

Shared tenant values such as Graph app-only authentication, `From`, Teams webhooks, and output roots still come from the standard SmartM365 tenant configuration:

```text
SmartM365.global.local.json
Config/Tenants/<TenantKey>.local.json
```

Communication-specific values live under this area and must remain local:

```text
Communications/ExchangeUserNotifications/Config/Communications.local.json
Communications/ExchangeUserNotifications/Config/Campaigns/<Campaign>.local.json
```

Only `*.template` files are committed. Real `*.local.json` files are ignored by Git.

## Safety

Use `-WhatIf` before every live run. Campaign scripts maintain a sent registry per tenant and campaign so users are not notified twice unless `-ForceSend` is used.

Generated recipients CSVs, logs, sent registries, and campaign outputs must stay under `SmartM365/Data` or another ignored local runtime path.
