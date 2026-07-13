# Power Query Notes

Power BI should load CSV files from the tenant Power BI output folder:

```text
SmartWorkplaceCMDB/Data/Tenants/<ProfileKey>/DATA-LAST/PowerBI/
```

Recommended parameter:

- `PowerBIDataPath`: folder path to the `PowerBI` output directory.

Each table query should read one CSV file from that folder, require `TenantKey`, `OrganizationKey`, and `EnvironmentKey` for tenant-scoped tables (`TenantId` is optional), and apply only type conversion and display-friendly column naming. `DimDate.csv` is the explicit tenant-neutral exception.
