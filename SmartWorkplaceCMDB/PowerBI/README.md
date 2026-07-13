# SmartWorkplaceCMDB Power BI

Power BI is a native deliverable of SmartWorkplaceCMDB.

The model should consume curated CMDB and Power BI-ready CSV tables from:

```text
SmartWorkplaceCMDB/Data/Tenants/<ProfileKey>/DATA-LAST/PowerBI/
```

## Multi-tenant CSV contract

Every curated table except `DimDate.csv` starts with `TenantKey`, `OrganizationKey`, `EnvironmentKey`, and `TenantId`. Relationship-bearing tables also expose hidden-key inputs such as `TenantUserKey`, `TenantDeviceKey`, and `TenantSkuKey`, built as `TenantKey|SourceIdentifier`. `DimDate.csv` is the explicit tenant-neutral exception.

## Priority

1. Direction and executive overview.
2. Technician and detailed inventory.

## Initial Pages

- Executive Overview
- Users and Identity
- Devices and Compliance
- Licenses and Mailboxes
- Data Quality

## Model Direction

Use a star-schema oriented model:

- Dimensions: tenant, date, user, device, group, license SKU.
- Facts: user license, device compliance, user-device relationship, mailbox, data quality.

Power Query should load the curated tables with minimal transformation. Relationship scoring, source reconciliation, and CMDB logic belong in the PowerShell build step.
