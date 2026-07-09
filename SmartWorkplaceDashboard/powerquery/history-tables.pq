// SmartInventory history table query examples
HistoryMailboxSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\Exchange\\EXO\\Mailboxes\\WeeklyHistory", null);
HistoryOnPremMailboxSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\Exchange\\OnPrem\\Mailboxes\\WeeklyHistory", null);
HistoryDeviceSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\Intune\\Devices\\Inventory\\WeeklyHistory", null);
HistoryLicenseSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\M365\\Licensing\\Licenses\\WeeklyHistory", null);
HistoryUsageSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\M365\\Usage\\WeeklyHistory", null);
HistorySyncHealthSnapshots = fnLoadHistoryFolder(SmartInventoryHistoryPath & "\\M365\\SyncHealth\\AzureADConnect\\WeeklyHistory", null);
