// SmartInventory latest table query examples
Raw_M365ActiveUsers = fnLoadCsv(SmartInventoryLatestPath, "M365_Users_Active.csv");
Raw_EntraDevices = fnLoadCsv(SmartInventoryLatestPath, "M365_Entra_Devices.csv");
Raw_IntuneDevices = fnLoadCsv(SmartInventoryLatestPath, "Intune_Devices_Inventory.csv");
Raw_EXOMailboxes = fnLoadCsv(SmartInventoryLatestPath, "Exchange_EXO_Mailboxes_AllDomains.csv");
Raw_OnPremMailboxes = fnLoadCsv(SmartInventoryLatestPath, "Exchange_OnPrem_Mailboxes_AllDomains.csv");
Raw_LicenseUsers = fnLoadCsv(SmartInventoryLatestPath, "M365_Licenses_Users.csv");
Raw_BackupProtectedMailboxes = fnLoadCsv(SmartInventoryLatestPath, "M365_Backup_ProtectedMailboxes.csv");
Raw_SyncHealth = fnLoadCsv(SmartInventoryLatestPath, "M365_Entra_AzureADConnect_SyncHealth.csv");
