@{
    RootModule = "SmartM365-WindowsPowerShell5.psm1"
    ModuleVersion = "1.0.18"
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfhjZps+SI3Zmq
# ifIYH64FL+JIxkP8ElpBpD/7dyaCnaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDCsWh2CP6JTBnZGbSpIvIvm8Nf1m3BmW06GBaqBv9UZTANBgkqhkiG9w0B
# AQEFAASCAYAl9BYxPqgUO/uGjy2lD8De85AcRcKQ47WydGox8P8D6niezY9EoBkl
# 80OQ8BAYGhoXXiMDZ01nuFRU06OWU7ktHhALxGYHH7zsLvdZaV+CfQWf5VYSDb9W
# dH7dLUBEgI/aLGTPTpz89iB7M8AD4NOeGYKJMQbQ0eyE/ezeBSD5qT3+TLACbxi2
# z3H56WPMbcNE/Jqct3tOArab4JWnJEcuL5HVY+RZ6K2SdbIzJoFRaUL2PQV+IPp1
# wCwFlsgweiutRqU1H/Cx3AGRHfuQIHr1EOHBzbHmCugPnbAzI09YJ9vlU/ZkVJE8
# CbD+OtIv6HkcGIkJawI6I75sHAT3M8kc/ynOtl0zBXACf8EsDcgeZtDG66TJLSvU
# 5J+THHwUlkA5U1fsorQ4AZPJS/Y2zVZFWr2ckoBEGnkj5gjxwhnUA5d5Yu0BUxjZ
# uen7fxiP0A3ENrxgmDamEQ9IYO8XhgdAhO1WHVNklwDhemiHsqTTxc8s+7vzTHJA
# qs/HcIiYty0=
# SIG # End signature block
