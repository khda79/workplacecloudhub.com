# Migration Template

Copy this folder to create a new migration.

Example:

```powershell
Copy-Item -Recurse .\Migrations\_Template .\Migrations\Spain
```

Then update:

- `migration.config.psd1`
- `migration.config.source.txt`
- `migration.config.target.txt`

The `.cmd` files in `launchers` automatically detect the migration folder name.
If the copied folder is `Migrations\Spain`, they run with `-MigrationName Spain`.

Runtime logs are written to the migration `logs` folder by default. Scan CSVs,
comparison CSVs, Excel files, and generated operation scripts stay in their
dedicated output folders.
