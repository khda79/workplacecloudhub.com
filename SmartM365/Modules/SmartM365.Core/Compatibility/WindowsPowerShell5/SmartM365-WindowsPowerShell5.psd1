@{
    RootModule = "SmartM365-WindowsPowerShell5.psm1"
    ModuleVersion = "1.0.71"
    GUID = "dde15961-7933-412f-8d09-e1dd8a889b65"
    Author = "Internal automation team"
    CompanyName = "Internal"
    Copyright = "(c) 2025 Internal automation team. All rights reserved."
    Description = "Windows PowerShell 5.1 compatibility module for SmartM365 initialization, logging, cleanup, and cloud session helpers."
    PowerShellVersion = "5.1"
    FunctionsToExport = @(
        "Start-SmartM365InventoryRun",
        "Complete-SmartM365InventoryRun",
        "Get-SmartM365InventoryCompletionStatus",
        "Disable-SmartM365InterruptedPublication",
        "Invoke-SmartM365RunCleanup",
        "Start-SmartM365RunTranscript",
        "Stop-SmartM365RunTranscript",
        "Set-SmartM365InventoryConnectionOwnership",
        "Close-SmartM365InventoryConnections",
        "Test-SmartM365InterruptedRun",
        "InitializeScriptEnvironment",
        "RemoveOldFiles",
        "Remove-SmartM365TimestampedFilesOlderThan",
        "Remove-SmartM365TimestampedDirectoriesOlderThan",
        "Remove-OldFiles",
        "EnsureExchangePSSnapinLoaded",
        "Format-SmartM365LogLine",
        "Update-SmartM365TimestampedTranscript",
        "Get-SmartM365ModuleDiagnosticText",
        "Write-SmartM365LoadedModuleVersions",
        "Complete-SmartM365ExecutionContext",
        "Start-SmartM365CmdbSourceReceipt",
        "Set-SmartM365CmdbSourceScope",
        "Start-SmartM365SourceReceipt",
        "Complete-SmartM365SourceReceipt",
        "Assert-SmartM365SourcePublication",
        "Write-SmartM365CompletionBanner",
        "WriteLog",
        "Set-SmartM365CoreContext",
        "Get-SmartM365MaxItemsValue",
        "Test-SmartM365MaxItemsMode",
        "Get-SmartM365MaxItemsSuffix",
        "Set-SmartM365MaxItemsMode",
        "Add-SmartM365MaxItemsSuffixToCsvPath",
        "Add-SmartM365MaxItemsSuffixToBaseName",
        "Add-SmartM365MaxItemsMailBanner",
        "Add-SmartM365MaxItemsSubjectPrefix",
        "Get-SmartM365MailTenantName",
        "Format-SmartM365MailSubject",
        "Get-SmartM365MailScriptContext",
        "Add-SmartM365MailExecutionFooter",
        "Limit-SmartM365RowsForMaxItems",
        "Get-SmartM365CsvValidationBaseName",
        "Get-SmartM365CsvValidationRule",
        "Assert-SmartM365CsvDataCompleteness",
        "Add-SmartM365CsvValidationRule",
        "Initialize-SmartM365DefaultCsvValidationRules",
        "Add-SmartM365TenantKey",
        "Repair-SmartM365CsvTenantKeySchema",
        "Write-SmartM365CsvAtomically",
        "Add-SmartM365CsvRowsAtomically",
        "Copy-SmartM365FileAtomically",
        "Write-SmartM365TextAtomically",
        "Publish-SmartM365Csv",
        "Export-SmartM365Csv",
        "Invoke-SmartM365Preflight",
        "Initialize-SmartM365ImportExcel",
        "Save-SmartM365WeeklyInventoryHistory",
        "Add-SmartM365WeeklyHistory",
        "Send-SmartM365TeamsNotification",
        "SendFileListEmailReport",
        "NewTableFilesEmailBody",
        "ConvertTo-SmartM365EmailHtmlText",
        "ConvertTo-SmartM365ConfigBoolean",
        "Get-SmartM365MailBrandingConfig",
        "ConvertTo-SmartM365MailLogoDataUri",
        "Add-SmartM365MailBranding",
        "New-SmartM365EmailBody",
        "New-SmartM365EmailKpiGridHtml",
        "New-SmartM365EmailTableHtml",
        "ConvertTo-SmartM365EmailBody",
        "ExportAndCopyCsv",
        "ExportAndCopyCsvFromConvert",
        "ConvertTo-SmartM365SharePointDataRootPath",
        "Get-SmartM365SharePointRelativeFilePath",
        "Invoke-SmartM365SharePointCsvUpload",
        "Remove-SmartM365SharePointFile",
        "Remove-SmartM365SharePointTimestampedCsvOlderThan",
        "NewRemoteScheduledTaskAndWait",
        "SendEmailHtmlReport",
        "NewSimpleEmailBody",
        "NewTableEmailBody",
        "GetFileList",
        "Connect-SmartM365CloudSession",
        "Disconnect-SmartM365CloudSession"
    )
    VariablesToExport = @()
    AliasesToExport = @()
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7AZ9LucsrlWMp
# 36bOdcCDfp4H0m5twgnvgxTNDvVKNaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCxJD4WH4tlrRP+OkD9dWsC
# IKcGbMMIKZPDAdG6qWKCGjANBgkqhkiG9w0BAQEFAASCAYAWf3pmx8zew1BHiSWR
# BynQZ9GHwW2F3f9Rh7/GAZXLCJsYaeRLue6HigAVAn+QgMhiP+f07lq4BFlU3+0O
# 2qcwSS0o4SVdHWzsc3xgBYYeXd1CA2vc6BalEK4SpjYKTRecJyiD734nOVyewHZ4
# sn1MQJJTw5rS/DXKDTsRd3NHjdcF8QzTMeyuel9xO7G1+UJRFgewyTfm3zSWjbPg
# 2k5rfkSTMvV0+W5JJJYLgV2JouJaR2OURfDppvgskX1De4nJIzJYcRZMuQFCQhQH
# 7Bcc9iua4DhQIYdTAPVcw9C69XEB67znhknSi5p90SNDXNjCUKyiz+8j9TTZaNc/
# 5gRAawD5+9Vffiik23q3bbvwN0k1IjZGx0CEnVDE/SyrSYMWbfcJnY1as4H724TD
# Nlibqnk+ayP2koRecpIT38XAkpngvMfGh9fESNd+HHnePQnDjFaQFOPoY7jUotw2
# Egw0EFZiXRZ+ykUJmBl9Z1NDZ5av6nGbamFvAsE9tkNHIIA=
# SIG # End signature block
