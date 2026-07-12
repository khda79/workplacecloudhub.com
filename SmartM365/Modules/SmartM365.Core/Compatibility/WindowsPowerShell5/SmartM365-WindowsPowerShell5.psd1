@{
    RootModule = "SmartM365-WindowsPowerShell5.psm1"
    ModuleVersion = "1.0.22"
    GUID = "dde15961-7933-412f-8d09-e1dd8a889b65"
    Author = "Internal automation team"
    CompanyName = "Internal"
    Copyright = "(c) 2025 Internal automation team. All rights reserved."
    Description = "Windows PowerShell 5.1 compatibility module for SmartM365 initialization, logging, cleanup, and cloud session helpers."
    PowerShellVersion = "5.1"
    FunctionsToExport = @(
        "InitializeScriptEnvironment",
        "RemoveOldFiles",
        "Remove-OldFiles",
        "EnsureExchangePSSnapinLoaded",
        "Format-SmartM365LogLine",
        "Update-SmartM365TimestampedTranscript",
        "Get-SmartM365ModuleDiagnosticText",
        "Write-SmartM365LoadedModuleVersions",
        "Complete-SmartM365ExecutionContext",
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
        "Limit-SmartM365RowsForMaxItems",
        "Get-SmartM365CsvValidationBaseName",
        "Get-SmartM365CsvValidationRule",
        "Assert-SmartM365CsvDataCompleteness",
        "Add-SmartM365CsvValidationRule",
        "Initialize-SmartM365DefaultCsvValidationRules",
        "Write-SmartM365CsvAtomically",
        "Publish-SmartM365Csv",
        "Export-SmartM365Csv",
        "Invoke-SmartM365Preflight",
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
        "ConvertTo-SmartM365EmailBody",
        "ExportAndCopyCsv",
        "ExportAndCopyCsvFromConvert",
        "ConvertTo-SmartM365SharePointDataRootPath",
        "Get-SmartM365SharePointRelativeFilePath",
        "Invoke-SmartM365SharePointCsvUpload",
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCLukXr7qrzVEz
# y3qu2Z1+SyWNmtBdEJsiCMgFavfkSKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAPJSEB/zDI8XRTAW0WMs1clDPjQwx2ta+l1A+4VhbAGjANBgkqhkiG9w0B
# AQEFAASCAYBZfFSIVMcZ38iDvDkcmO+/BbIgt7JwzAs3lw7HEnLX/GVErUYA9pb3
# e+DZtlOuNAJSQCriuoQOg46RWud7nhFNWcWpCL8f20HeUMGrJlka2Nd2I/iSUiyv
# 0+i4y0FFnsurzF6FAr/9JWPbUQsEMJ19z/CPIHSXoLrEH39kqLWG9LpDnZeIqagC
# iAqQcEEnVRJCVGKUtNDV5O+oYRZSqpB7UlqDGCFd3shhZkRNgIzno0lWN8GSCOaZ
# kzcLF+td8VKQ2DBA2350djdHcn0RzeG9DUfI8SZV8YpHgCbR+N42ohTXhVFVGuUC
# 2xJ/Y1qP720Fue3h2NZYjqIKk1v+h42Dy74u+m4+eeXjqBMEV8WW9vouRBPUoXB9
# 3aPpHJ5xy5/PwwdLlqH6VZ1YqETZ6zatEgMuLexOleOIASlC5WDLKNCkZGCCEU5d
# gZAmV+yNdfZssErapjyW22dUVYnp4ZWJPP524mVy6v46Yai6Zd8gdrwiaeVQkvXn
# W27vay4+Du4=
# SIG # End signature block
