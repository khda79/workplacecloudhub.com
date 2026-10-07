function Get-SmartM365ProxyMailEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LatestCsvFolderPath,
        [Parameter(Mandatory)][string]$TenantKey,
        [datetime]$Now = (Get-Date)
    )

    $receiptPath = Join-Path $LatestCsvFolderPath 'SmartInventory_SmartM365-Check-ProxyAddresses-Exchange.current.json.txt'
    $detailPath = Join-Path $LatestCsvFolderPath 'Exchange_OnPrem_ProxyAddresses_Check.csv'
    $summaryPath = Join-Path $LatestCsvFolderPath 'Exchange_OnPrem_ProxyAddresses_Summary.csv'
    $workbookPath = Join-Path $LatestCsvFolderPath 'Exchange_OnPrem_ProxyAddresses.xlsx'
    $linksPath = Join-Path $LatestCsvFolderPath 'Exchange_OnPrem_ProxyAddresses_MailLinks.csv'
    try {
        if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'ProxyAddresses completion receipt is missing.' }
        $receiptHash=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash
        $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
        if ([string]$receipt.Producer -ne 'SmartM365-Check-ProxyAddresses-Exchange.ps1' -or [string]$receipt.Status -ne 'Completed') {
            throw 'ProxyAddresses completion receipt is not successful.'
        }
        if ([string]$receipt.TenantKey -ne $TenantKey) { throw 'ProxyAddresses receipt belongs to another tenant.' }
        if ([int]$receipt.Errors -ne 0) { throw 'ProxyAddresses receipt contains errors.' }
        $completedAt = [DateTimeOffset]::Parse([string]$receipt.CompletedAtUtc, [Globalization.CultureInfo]::InvariantCulture).ToLocalTime()
        $startedAtUtc = [DateTimeOffset]::Parse([string]$receipt.StartedAtUtc, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime
        if ($completedAt.Date -ne $Now.Date) { throw ('Last ProxyAddresses audit was completed on {0:yyyy-MM-dd}, not today.' -f $completedAt) }

        foreach ($fileName in @('Exchange_OnPrem_ProxyAddresses_Check.csv', 'Exchange_OnPrem_ProxyAddresses_Summary.csv')) {
            $filePath = Join-Path $LatestCsvFolderPath $fileName
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw ("ProxyAddresses output is missing: {0}" -f $fileName) }
            $records = @($receipt.Files | Where-Object { [string]$_.File -eq $fileName -and [string]$_.RunId -eq [string]$receipt.RunId })
            if ($records.Count -ne 1 -or [string]$records[0].Status -ne 'Success' -or [string]::IsNullOrWhiteSpace([string]$records[0].SHA256)) {
                throw ("ProxyAddresses output is not qualified by the receipt: {0}" -f $fileName)
            }
            $actualHash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256 -ErrorAction Stop).Hash
            if ($actualHash -ine [string]$records[0].SHA256) { throw ("ProxyAddresses output hash differs from the receipt: {0}" -f $fileName) }
        }

        $summary = @(Import-Csv -LiteralPath $summaryPath -ErrorAction Stop)
        $detail = @(Import-Csv -LiteralPath $detailPath -ErrorAction Stop)
        if ($summary.Count -eq 0 -or $detail.Count -eq 0) { throw 'ProxyAddresses CSV output is empty.' }
        foreach ($pair in @(@('Exchange_OnPrem_ProxyAddresses_Check.csv', $detail.Count), @('Exchange_OnPrem_ProxyAddresses_Summary.csv', $summary.Count))) {
            $record = @($receipt.Files | Where-Object { [string]$_.File -eq [string]$pair[0] })[0]
            if ([int]$record.Rows -ne [int]$pair[1]) { throw ("ProxyAddresses receipt row count differs from CSV: {0}" -f $pair[0]) }
        }
        if (@($summary | Where-Object { [string]$_.TenantKey -ne $TenantKey }).Count -gt 0 -or
            @($detail | Where-Object { [string]$_.TenantKey -ne $TenantKey }).Count -gt 0) {
            throw 'ProxyAddresses CSV output contains another tenant.'
        }
        foreach ($metric in @('With expected address present', 'With expected address missing', 'Planned address additions if Write is enabled', 'On-premises mailboxes processed', 'Remote mailboxes processed')) {
            $values = @($summary | Where-Object { $_.Summary -eq $metric })
            $parsedCount = 0
            if ($values.Count -ne 1 -or -not [int]::TryParse([string]$values[0].Count, [ref]$parsedCount)) {
                throw ("ProxyAddresses summary metric is missing or invalid: {0}" -f $metric)
            }
        }
        $totals = @($summary | Where-Object { $_.Summary -eq 'Total recipients processed' })
        if ($totals.Count -ne 1 -or [int]$totals[0].Count -ne $detail.Count) { throw 'ProxyAddresses summary and detail populations disagree.' }
        $mode = 'Unknown'
        if ($receipt.ScopeParameters -and $null -ne $receipt.ScopeParameters.AddMissingAddress) {
            $mode = if ([bool]$receipt.ScopeParameters.AddMissingAddress) { 'Write mode' } else { 'Read-only mode' }
        }
        $scope = if ($receipt.ScopeParameters -and [bool]$receipt.ScopeParameters.AllOrganizationalUnit) {
            'ALL (entire forest)'
        }
        elseif ($receipt.ScopeParameters) {
            (@($receipt.ScopeParameters.OrganizationalUnit) | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join '; '
        }
        else { 'Unknown' }
        $suffixes = @($detail | Where-Object { $_.ExpectedAddressSource -eq 'Alias' -and $_.ExpectedAddress -match '@' } |
            ForEach-Object { ([string]$_.ExpectedAddress -split '@')[-1].ToLowerInvariant() } | Sort-Object -Unique)
        $expectedSuffix = if ($suffixes.Count -eq 1) { $suffixes[0] } else { 'Unavailable from CSV' }
        $workbook = if ((Test-Path -LiteralPath $workbookPath -PathType Leaf) -and
            (Get-Item -LiteralPath $workbookPath).LastWriteTimeUtc -ge $startedAtUtc) { $workbookPath } else { '' }
        $links = @()
        if ((Test-Path -LiteralPath $linksPath -PathType Leaf) -and
            (Get-Item -LiteralPath $linksPath).LastWriteTimeUtc -ge $startedAtUtc) {
            $linkRecords = @($receipt.Files | Where-Object { [string]$_.File -eq 'Exchange_OnPrem_ProxyAddresses_MailLinks.csv' -and [string]$_.RunId -eq [string]$receipt.RunId })
            if ($linkRecords.Count -eq 1 -and [string]$linkRecords[0].Status -eq 'Success' -and
                (Get-FileHash -LiteralPath $linksPath -Algorithm SHA256 -ErrorAction Stop).Hash -ieq [string]$linkRecords[0].SHA256) {
                $links = @(Import-Csv -LiteralPath $linksPath -ErrorAction Stop | Where-Object { [string]$_.TenantKey -eq $TenantKey })
            }
        }
        Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
        if ((Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash -cne $receiptHash) { throw 'ProxyAddresses receipt changed during read.' }
        foreach ($record in @($receipt.Files)) {
            if ($record.File -in @('Exchange_OnPrem_ProxyAddresses_Check.csv','Exchange_OnPrem_ProxyAddresses_Summary.csv','Exchange_OnPrem_ProxyAddresses_MailLinks.csv') -and
                (Get-FileHash -LiteralPath (Join-Path $LatestCsvFolderPath $record.File) -Algorithm SHA256).Hash -ine $record.SHA256) { throw 'ProxyAddresses output changed during read.' }
        }
        return [pscustomobject]@{
            Available = $true; Reason = ''; CompletedAt = $completedAt; Version = [string]$receipt.ScriptVersion
            Mode = $mode; Scope = $scope; ExpectedSuffix = $expectedSuffix
            Summary = $summary; Detail = $detail; DetailPath = $detailPath; SummaryPath = $summaryPath; WorkbookPath = $workbook; Links = $links
        }
    }
    catch {
        return [pscustomobject]@{ Available = $false; Reason = $_.Exception.Message; CompletedAt = $null; Summary = @(); Detail = @() }
    }
}

function New-SmartM365ProxyMailSections {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Evidence, [bool]$ShowMailLinks = $false)

    $encode = { param($value) [System.Net.WebUtility]::HtmlEncode([string]$value) }
    if (-not $Evidence.Available) {
        $reason = & $encode $Evidence.Reason
        return @([pscustomobject]@{ Title = 'Check ProxyAddresses Report'; Html = "<p style='color:#92400e'>Today's ProxyAddresses audit is unavailable: $reason</p>" })
    }

    $summaryMap = @{}
    foreach ($row in $Evidence.Summary) { $summaryMap[[string]$row.Summary] = [string]$row.Count }
    $missingCount = [int]$summaryMap['With expected address missing']
    $message = if ($missingCount -gt 0) {
        'Review missing proxy addresses before remediation. Recipients managed by an email address policy, duplicate expected addresses, addresses already assigned to another mail-enabled recipient, and remote mailboxes without a usable RemoteRoutingAddress are always skipped in write mode.'
    }
    else { 'The audit found no missing expected proxy addresses.' }
    $contextHtml = "<p><strong>Audit completed:</strong> $(& $encode $Evidence.CompletedAt.ToString('yyyy-MM-dd HH:mm:ss zzz')) &nbsp; <strong>Version:</strong> $(& $encode $Evidence.Version) &nbsp; <strong>Mode:</strong> $(& $encode $Evidence.Mode)</p>"
    $contextHtml += "<p><strong>Scope:</strong> $(& $encode $Evidence.Scope) &nbsp; <strong>Expected suffix:</strong> $(& $encode $Evidence.ExpectedSuffix)</p>"
    $contextHtml += "<p style='color:#92400e'>$(& $encode $message)</p>"
    $sections = @([pscustomobject]@{ Title = 'Check ProxyAddresses Report'; Html = $contextHtml })

    $metricRows = foreach ($row in $Evidence.Summary) {
        '<tr><td style="border-bottom:1px solid #eef2f7;padding:7px">{0}</td><td style="border-bottom:1px solid #eef2f7;padding:7px;text-align:right">{1}</td></tr>' -f (& $encode $row.Summary), (& $encode $row.Count)
    }
    $sections += [pscustomobject]@{ Title = 'Proxy address summary'; Html = '<table style="width:100%;border-collapse:collapse">{0}</table>' -f ($metricRows -join "`n") }

    $duplicateGroups = @($Evidence.Detail | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Alias) } |
        Group-Object Alias | Where-Object { $_.Count -gt 1 } | Sort-Object @{ Expression = 'Count'; Descending = $true }, Name)
    $duplicateRows = @($duplicateGroups | ForEach-Object { $group = $_; $group.Group | Sort-Object DisplayName, Identity | ForEach-Object {
        [pscustomobject]@{ Alias = $group.Name; Count = $group.Count; Record = $_ }
    } } | Select-Object -First 50)
    if ($duplicateRows.Count -gt 0) {
        $rows = foreach ($item in $duplicateRows) {
            '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f
                (& $encode $item.Alias), (& $encode $item.Count), (& $encode $item.Record.DisplayName),
                (& $encode $item.Record.PrimaryAddress), (& $encode $item.Record.ExpectedAddress),
                (& $encode $item.Record.SuggestedUniqueAddress), (& $encode $item.Record.Status)
        }
        $sections += [pscustomobject]@{ Title = 'Top 50 duplicate aliases'; Html = '<table style="width:100%;border-collapse:collapse"><tr><th>Alias</th><th>Count</th><th>Display name</th><th>Primary SMTP</th><th>Expected proxy</th><th>Suggested proxy</th><th>Status</th></tr>{0}</table>' -f ($rows -join "`n") }
    }

    $missingRows = @($Evidence.Detail | Where-Object { $_.Status -like 'Missing*' } | Sort-Object Status, DisplayName | Select-Object -First 50)
    if ($missingRows.Count -gt 0) {
        $rows = foreach ($row in $missingRows) {
            '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f
                (& $encode $row.DisplayName), (& $encode $row.Alias), (& $encode $row.PrimaryAddress),
                (& $encode $row.ExpectedAddress), (& $encode $row.Status)
        }
        $sections += [pscustomobject]@{ Title = 'Top 50 missing proxy addresses'; Html = '<table style="width:100%;border-collapse:collapse"><tr><th>Display name</th><th>Alias</th><th>Primary SMTP</th><th>Expected proxy</th><th>Status</th></tr>{0}</table>' -f ($rows -join "`n") }
    }

    $paths = @($Evidence.DetailPath, $Evidence.SummaryPath, $Evidence.WorkbookPath) | Where-Object { $_ }
    $pathRows = foreach ($path in $paths) {
        $fileName = Split-Path -Path $path -Leaf
        $link = @($Evidence.Links | Where-Object { [string]$_.FileName -eq $fileName -and [string]$_.WebUrl -match '^https://' } | Select-Object -First 1)
        if ($ShowMailLinks -and $link.Count -eq 1) {
            '<tr><td style="word-break:break-all"><a href="{0}">{1}</a></td></tr>' -f (& $encode $link[0].WebUrl), (& $encode $path)
        }
        else { '<tr><td style="word-break:break-all">{0}</td></tr>' -f (& $encode $path) }
    }
    $sections += [pscustomobject]@{ Title = 'Proxy address files'; Html = '<table style="width:100%;border-collapse:collapse">{0}</table>' -f ($pathRows -join "`n") }
    return $sections
}

Export-ModuleMember -Function Get-SmartM365ProxyMailEvidence, New-SmartM365ProxyMailSections

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIj8RGw6XmB/wc
# OPBqUIG4rjMbzaOhgjpP+lX3MVSSjKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIKMJxMKr9+wwqhgfqhpd1X8iD6gqw44/9XNaTbcVGRZHMA0GCSqG
# SIb3DQEBAQUABIIBgFf0RfeD8oI5QBfOSR1LY/WZpwqCH5apEmlL3GaE89LBgYfw
# RovBCZIDRYDxZ1YEkkhaMjvphyHKYBYcfXE+30Ta7NvoRlrQnAe+2wz0Ya7RweKQ
# ZOnsKacAzH7anqc6UpL7P32KXsKO486vK1dvKhbYjOkqesrIPX0oNV+JZavLKkyv
# 1Heu2a6k6DvR28u20xee6oQWKEy5tvZAmVVtZvkhW5K57VQnkmMjiTgB+Jd4712V
# pcgmG6oefxWciGOSjnHFCkaAlTTzDY4jFwTXEBXya7QCyWqZaA2dkc2JRodcR59e
# k/t+oYyhQp3NMBpm7zSsrKA5X7ZSkGkFHD/1wO9TiakHT1uR83W/wELD9wF3Li1r
# qb4YSl+IRr+FYH01HyYuEmfZ6lyUzdxT1rqehB4ws43oBHlYK/RpypmhiUVzk/2f
# SPeJfdp2nREgpNtjxnj333yZ7eh0MrCFnsXrhb7s9aS6rha08mlO5XO39rWZkqhl
# FGzv+VHcpZxIGnpQiaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxMDMx
# MjNaMC8GCSqGSIb3DQEJBDEiBCBPCxVttmbZ20pGmkgzAUrUJamvudUTfTdEVYwk
# CG7lpDANBgkqhkiG9w0BAQEFAASCAgCZ+cgZGaJ/nc//lKVOgYif13ajQZJGJlQy
# WYqlp0t1Gd4qqokIi3TC5Wnks2xO6wrSzrOgxItjuBO7gr/Fla10o2FPHvlIEn6f
# Gdd7HINYI1rS3mg01Z3slYOYOIVt/5gMxc3OcyRelTcWBP+cRB2SM9SAc+apImjh
# SFcd07zR4ZP5hbzAe1cp8OjA2/zBcgpU9Tf5l8NovvWqVDTtT/skRSvztrgzq9OL
# P31r61p3Rb6xF5BmJ2ltLUkoX9BlAcC2F//pDnUyZ+DnJbzDvtUGPSwyY7KF3K8q
# ihcHxq+WnkbBW9VD6vywVk774pMktUMlaIL608k7NzRFJ633v8sTuwNslBWbAnsy
# bfT/03AW4wgIGHMYSoz3ZaDGM6PitKW/zIiwOgMOWpRexLGiimVnNd2og7P8ZFAK
# 7PeqsxE5cJ6NatuwXGci7j+EBGC4joQ8b4SpKmb4hyfIDnmY6tfTiyIU33zMn8zM
# z7/XbTS6aBafdp/Mrj6V0fmKN+F33HJZ6SfBRJrXBnYgCD3MtHqGqqheiufP7U8S
# fH9DOuymvOzva7jL99TnsdA7rEb1gvt3Hdjn3/y58w+x3RXfxF5PEjeEQds4LLDz
# sMGlxSaI9eTLp4BcYhrvXK/2OkxpTb9gN6TpKOzoQfP5wJRNXkvyKRjSnBd2DF9L
# JESTRf9zeA==
# SIG # End signature block
