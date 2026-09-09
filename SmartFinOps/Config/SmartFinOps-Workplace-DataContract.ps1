Set-StrictMode -Version 2.0

function Test-SmartFinOpsExcludedCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FileName)

    foreach ($pattern in @($script:SmartFinOpsExcludedFileNamePatterns)) {
        if ([string]::IsNullOrWhiteSpace([string]$pattern)) { continue }
        if ($FileName -match [regex]::Escape([string]$pattern)) { return $true }
    }
    return $false
}

function Get-SmartFinOpsSourceContract {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Key)

    $matches = @($script:SmartFinOpsSourceContract.Sources | Where-Object { $_.Key -eq $Key })
    if ($matches.Count -ne 1) {
        throw "Expected exactly one SmartFinOps source contract for key '$Key'; found $($matches.Count)."
    }
    return $matches[0]
}

function Get-SmartFinOpsCsvHeaderColumns {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($firstLine)) { return @() }

    $delimiter = Get-SmartFinOpsCsvDelimiter -Path $Path

    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string]$delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        return @($parser.ReadFields())
    }
    finally {
        $parser.Dispose()
    }
}

function Test-SmartFinOpsCsvHasDataRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($firstLine)) { return $false }

    $delimiter = Get-SmartFinOpsCsvDelimiter -Path $Path

    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string]$delimiter)
        $parser.HasFieldsEnclosedInQuotes = $true
        [void]$parser.ReadFields()
        while (-not $parser.EndOfData) {
            if ($null -ne $parser.ReadFields()) { return $true }
        }
        return $false
    }
    finally {
        $parser.Dispose()
    }
}

function Get-SmartFinOpsCsvReportRefreshDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$HeaderColumns
    )

    $candidateColumns = @('ReportRefreshDate', 'Report Refresh Date', 'Report refresh date')
    $reportRefreshColumn = @($candidateColumns | Where-Object { $_ -in $HeaderColumns } | Select-Object -First 1)
    if ($reportRefreshColumn.Count -eq 0) { return $null }

    $oldest = $null
    $invalidDate = $false
    Import-Csv -LiteralPath $Path -Delimiter (Get-SmartFinOpsCsvDelimiter -Path $Path) | ForEach-Object {
        $date = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $_ -Names @($reportRefreshColumn[0]))
        if ($null -eq $date -or $date -gt (Get-Date)) { $invalidDate = $true }
        elseif ($null -eq $oldest -or $date -lt $oldest) { $oldest = $date }
    }
    if ($invalidDate) { return $null }
    return $oldest
}

function Resolve-FirstExistingCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string[]]$FileNames
    )

    foreach ($fileName in $FileNames) {
        if (Test-SmartFinOpsExcludedCsv -FileName $fileName) { continue }
        $path = Join-Path -Path $FolderPath -ChildPath $fileName
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    return ''
}

function Import-SmartFinOpsSourceCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string[]]$FileNames,
        [string]$SemanticRole = '',
        [string[]]$RequiredColumns = @(),
        [switch]$Optional,
        [string]$BlockedReason = '',
        [switch]$ValidationOnly,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DataQualityRows
    )

    $eligibleFileNames = @($FileNames | Where-Object { -not (Test-SmartFinOpsExcludedCsv -FileName $_) })
    if ($eligibleFileNames.Count -eq 0) {
        Write-SmartFinOpsLog -Message ("Skipped source {0}; all candidates match the excluded filename patterns." -f $SourceName)
        return @()
    }

    $path = Resolve-FirstExistingCsv -FolderPath $script:SmartM365LatestCsvFolderPath -FileNames $eligibleFileNames
    if ([string]::IsNullOrWhiteSpace($path)) {
        $missingStatus = if (-not [string]::IsNullOrWhiteSpace($BlockedReason)) {
            'Blocked'
        }
        elseif ($Optional) {
            'Missing optional'
        }
        else {
            'Missing'
        }
        $missingNotes = if ($missingStatus -eq 'Blocked') {
            $BlockedReason
        }
        elseif ($Optional) {
            'Optional source CSV not found in SmartM365 DATA-LAST.'
        }
        else {
            'Source CSV not found in SmartM365 DATA-LAST.'
        }
        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            SourceRequirement = if ($Optional) { 'Optional' } else { 'Required' }
            Status = $missingStatus
            ContractStatus = 'NotChecked'
            FreshnessStatus = 'NotChecked'
            FreshnessBasis = 'NotChecked'
            Path = ($eligibleFileNames -join ' | ')
            FileName = ''
            RowCount = 0
            ColumnCount = 0
            ReportRefreshDate = ''
            LastWriteTime = ''
            AgeHours = ''
            FileAgeHours = ''
            RequiredColumnsMissing = ''
            BlockedReason = $BlockedReason
            Notes = $missingNotes
        }) | Out-Null
        return @()
    }

    try {
        $item = Get-Item -LiteralPath $path
        $fileAgeHours = [math]::Round(((Get-Date) - $item.LastWriteTime).TotalHours, 1)
        $headerColumns = @(Get-SmartFinOpsCsvHeaderColumns -Path $path)
        $reportRefreshDate = Get-SmartFinOpsCsvReportRefreshDate -Path $path -HeaderColumns $headerColumns
        $freshnessBasis = if ($reportRefreshDate) { 'ReportRefreshDate' } else { 'LastWriteTime' }
        $ageHours = if ($reportRefreshDate) {
            [math]::Round(((Get-Date) - $reportRefreshDate).TotalHours, 1)
        }
        else { $fileAgeHours }
        $hasRefreshColumn = @($headerColumns | Where-Object { $_ -in @('ReportRefreshDate', 'Report Refresh Date') }).Count -gt 0
        $freshnessStatus = if (($hasRefreshColumn -and $null -eq $reportRefreshDate) -or $ageHours -lt 0) { 'Unknown' } elseif ($ageHours -gt $script:SmartFinOpsMaxSourceAgeHours) { 'Stale' } else { 'Fresh' }
        if ($hasRefreshColumn -and $null -eq $reportRefreshDate) { $freshnessBasis = 'InvalidReportRefreshDate'; $ageHours = '' }
        $missingColumns = @($RequiredColumns | Where-Object { $_ -notin $headerColumns })
        $contractStatus = if ($missingColumns.Count -eq 0) { 'Valid' } else { 'Invalid' }
        $hasDataRow = Test-SmartFinOpsCsvHasDataRow -Path $path
        $rows = @(if (-not $ValidationOnly -and $contractStatus -eq 'Valid') { Import-Csv -LiteralPath $path -Delimiter (Get-SmartFinOpsCsvDelimiter -Path $path) })
        $uniqueColumns = switch ($SourceName) {
            'M365 active users' { @('User principal name', 'UserPrincipalName') }
            'Active Directory users - canonical enriched' { @('UserPrincipalName') }
            'M365 tenant licenses' { @('TenantSkuPartNumber') }
            'Exchange Online mailboxes' { @('UserPrincipalName') }
        }
        $identityIssue = $false
        if ($uniqueColumns -and -not $ValidationOnly) {
            $keys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($row in $rows) {
                $key = ([string](Get-RowPropertyValue -Row $row -Names $uniqueColumns)).Trim()
                if (-not $key -or -not $keys.Add($key)) { $identityIssue = $true }
            }
            if ($identityIssue) { $contractStatus = 'Invalid'; $rows = @() }
        }
        $rowCount = if ($ValidationOnly) { '' } else { $rows.Count }
        $status = if ($contractStatus -eq 'Invalid') {
            'Invalid schema'
        }
        elseif (-not $hasDataRow) {
            'Empty'
        }
        else {
            'Loaded'
        }
        $notes = New-Object System.Collections.Generic.List[string]
        if ($identityIssue) { $notes.Add('Duplicate or blank identity in a single-grain source; source excluded from calculations.') | Out-Null }
        if ($freshnessStatus -eq 'Unknown') { $notes.Add('Missing, invalid or future source timestamp; freshness cannot be established.') | Out-Null }
        if ($freshnessStatus -eq 'Stale') {
            $notes.Add(("Source is older than {0} hours." -f $script:SmartFinOpsMaxSourceAgeHours)) | Out-Null
        }
        if ($missingColumns.Count -gt 0) { $notes.Add('Required columns are missing.') | Out-Null }
        if (-not $hasDataRow) { $notes.Add('CSV contains a header but no data rows.') | Out-Null }
        if ($ValidationOnly) { $notes.Add('Header and freshness validation only; source rows were not imported.') | Out-Null }

        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            SourceRequirement = if ($Optional) { 'Optional' } else { 'Required' }
            Status = $status
            ContractStatus = $contractStatus
            FreshnessStatus = $freshnessStatus
            FreshnessBasis = $freshnessBasis
            Path = $path
            FileName = $item.Name
            RowCount = $rowCount
            ColumnCount = $headerColumns.Count
            ReportRefreshDate = if ($reportRefreshDate) { $reportRefreshDate.ToString('yyyy-MM-dd') } else { '' }
            LastWriteTime = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            AgeHours = $ageHours
            FileAgeHours = $fileAgeHours
            RequiredColumnsMissing = ($missingColumns -join ' | ')
            BlockedReason = ''
            Notes = ($notes -join ' ')
        }) | Out-Null

        if ($missingColumns.Count -gt 0) {
            Write-SmartFinOpsLog -Level WARN -Message ("Source contract invalid for {0}: missing {1}" -f $SourceName, ($missingColumns -join ', '))
        }
        if ($freshnessStatus -eq 'Stale') {
            Write-SmartFinOpsLog -Level WARN -Message ("Source is stale: {0}, age={1}h" -f $item.Name, $ageHours)
        }
        if (-not $ValidationOnly) {
            Write-SmartFinOpsLog -Message ("Loaded source {0}: {1} row(s)" -f $SourceName, $rows.Count)
        }
        return $rows
    }
    catch {
        $DataQualityRows.Add([pscustomobject]@{
            RunId = $script:RunId
            SourceName = $SourceName
            SemanticRole = $SemanticRole
            SourceRequirement = if ($Optional) { 'Optional' } else { 'Required' }
            Status = 'Error'
            ContractStatus = 'Error'
            FreshnessStatus = 'NotChecked'
            FreshnessBasis = 'NotChecked'
            Path = $path
            FileName = [System.IO.Path]::GetFileName($path)
            RowCount = 0
            ColumnCount = 0
            ReportRefreshDate = ''
            LastWriteTime = ''
            AgeHours = ''
            FileAgeHours = ''
            RequiredColumnsMissing = ''
            BlockedReason = $BlockedReason
            Notes = $_.Exception.Message
        }) | Out-Null
        Write-SmartFinOpsLog -Level WARN -Message ("Failed to load source {0}: {1}" -f $SourceName, $_.Exception.Message)
        return @()
    }
}

function Import-SmartFinOpsContractSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Key,
        [switch]$ValidationOnly,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DataQualityRows
    )

    $contract = Get-SmartFinOpsSourceContract -Key $Key
    $contractValidationOnly = $contract.PSObject.Properties['ValidationOnly'] -and [bool]$contract.ValidationOnly
    $contractOptional = $contract.PSObject.Properties['Optional'] -and [bool]$contract.Optional
    $contractBlockedReason = if ($contract.PSObject.Properties['BlockedReason']) { [string]$contract.BlockedReason } else { '' }
    return Import-SmartFinOpsSourceCsv `
        -SourceName ([string]$contract.DisplayName) `
        -FileNames @($contract.FileNames) `
        -SemanticRole ([string]$contract.SemanticRole) `
        -RequiredColumns @($contract.RequiredColumns) `
        -Optional:$contractOptional `
        -BlockedReason $contractBlockedReason `
        -ValidationOnly:($ValidationOnly -or $contractValidationOnly) `
        -DataQualityRows $DataQualityRows
}

function Get-SmartFinOpsCsvDelimiter {
    param([Parameter(Mandatory)][string]$Path)
    $line = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    # Ignore separators inside quoted header fields.
    $unquoted = [regex]::Replace([string]$line, '"(?:[^"]|"")*"', '')
    if (($unquoted.ToCharArray() | Where-Object { $_ -eq ';' } | Measure-Object).Count -gt ($unquoted.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count) { return [char]';' }
    return [char]','
}

function Test-SmartFinOpsDecisionSources {
    param([AllowEmptyCollection()][object[]]$DataQualityRows)
    $required = @('M365 active users', 'M365 user activity', 'M365 mailbox usage', 'M365 OneDrive usage', 'M365 Apps activations', 'M365 Teams user activity', 'M365 email activity', 'M365 license user assignments', 'M365 tenant licenses', 'Active Directory users - canonical enriched', 'Intune devices')
    foreach ($name in $required) {
        $match = @($DataQualityRows | Where-Object { $_.SourceName -eq $name })
        if ($match.Count -ne 1 -or $match[0].Status -ne 'Loaded' -or $match[0].FreshnessStatus -ne 'Fresh') { return $false }
    }
    # An optional activity source may be absent; a present but unusable source cannot prove inactivity.
    foreach ($row in $DataQualityRows) {
        if ($row.SourceName -in @('M365 SharePoint user activity','M365 Teams device usage','M365 Copilot user usage','M365 Teams Phone user usage') -and $row.Status -ne 'Missing optional' -and ($row.Status -ne 'Loaded' -or $row.FreshnessStatus -ne 'Fresh')) { return $false }
    }
    return $true
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCbjy9mF17haLrh
# 5v3QCcRzgAxGwVi8qicc6OPta81QDqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID0BXqc9OfGQLkEeOf2V4thPFO/2xg0BbWPv42mCm5hzMA0GCSqG
# SIb3DQEBAQUABIIBgEKfCmFvvt3eI5UULI6mvzCfz22sXBsUs9GLbJAvrDZasaA9
# 8Fq/IDYCbJY2/z6yDP/UJ6XQjYsF3Duz3jYJ4l7lUH7AKWSOk0qYnEKPzLvwqr/R
# u9spcRsCQ0gOVHZZIwxmDRjYqT7FUt8S3+wdjnm4zEA3xIv1o1kxELR3YyX+Co+c
# hDBSpF98y0b8FvmMTAVn+f9PuVIxhophOrJLNABvwbEoj8yDvkeAvHPSZNi52svb
# JmPgRcAH+/UlpjztirM9zP4JK0AbEJXvJqwQ3ThoJ+Nh6mOTuopXifi5KvLI1INX
# ivWI1Ka0kpz3bI0cZfVSe7EkJ75K5OOAzObimoDeEWcIusVyVZeuTUMUXU7COFHH
# 6Sj4KnjN0llAdhoOO726oEtC79FidvJCPl9QfIQkvKEtbj0HSezyRjve9rVdixVH
# lH1mrA0uJin4/ixaYJNlTv6eJDDJ0TeoJ5fTuqxKtHbEifu8wWy2bRS9g6pEB/Ss
# 6M8XyCkPUM1OlJekW6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkwODA2
# MDJaMC8GCSqGSIb3DQEJBDEiBCAi4AQU1PmXqObbMxEOKeDpAP/capt9NiWU6AqC
# 8c1iDzANBgkqhkiG9w0BAQEFAASCAgCqAL2o6/bKbC9ztxoitAXAAB2hqY32vn6k
# 1wSaY4ThKTbqYYMHtBFPqSHMWcDwH2+j7v5inRxGeBi/jaMRHGsUB0BBA4wEqyhE
# yO6SQoHjJebYFIcnCypGiRO1qGQAcYL+CT6cbfgGtuhQZJ/pkN4L/vcd0WFb+f8I
# A0/IRlTUIzuDBBR70PKRFu833nPHQjXJao2NpM/XoqG1zYpp28UOIJWjlR39rFEi
# xaSg9P0949RhggIwKQGQGEhknppREiwD6VVMlRGlS5JF/mg+spShOvIR270Gczzc
# rFM/9qb3SGyVb0Qgd8SHjysVI9ak/eLgID6OQgss4fq69z+fBA7ltw21NlvfbQ2u
# JiNOaxPIa4t5N+Z2qBYKXsUkLdLUCc7SEx0j6tey+5VigxdjIgXV1F7vhfLk9OTm
# Dp2g+bR2EsMpdTd0mG3cv+5T9hVLwMsFkscEeDvAo4tz61bvR0jFfYfHV/z9nAFA
# LCK28MUJtGg6hu9hDxA5QZDcuBj+SwUukP/IA35tHliBqvSe/D16hq02XxMGr+1B
# gkegHegiWFzQcLCBl81BKXzGHtxs2Z6MwuG/qKkzjsLFPc7hUcegCzYtfBXX5kcT
# ExOCjCwUFwa2qcvSeFeh34hYydbMoqhtQePfxW3ZyqCnvMHisn9ILEqBSd5HFpXB
# AziHguSO+Q==
# SIG # End signature block
