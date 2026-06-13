# Active Directory Inventory

Active Directory inventory and reporting scripts.

## Scripts

- `SmartM365-ActiveDirectory-Inventory.ps1`: full AD inventory across domains.
- `SmartM365-ActiveDirectory-Report.ps1`: daily AD computer/user reporting. It first reuses fresh `AD_Computers_AllDomains.csv` and `AD_Users_AllDomains.csv` files from `LatestCsvFolderPath` when available, then falls back to live AD reads.

## Report Source

`SmartM365-ActiveDirectory-Report.ps1` can avoid a second AD scan when the full inventory has just run.

- `UseLatestInventoryCsvForReport`: defaults to `true`.
- `LatestInventoryCsvMaxAgeMinutes`: defaults to `720`; set to `0` or less to disable the freshness check.
