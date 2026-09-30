# Version: 1.1.0
$script:SmartWorkplaceCMDBCollectionVersion = '1.1.0'

function Read-SmartWorkplaceCMDBCollectionFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $document = Read-SmartWorkplaceCMDBJsonFile -Path $Path
    if ($document -is [array]) { return $document }
    if ($null -ne $document -and $null -ne $document.PSObject.Properties['value'] -and $document.value -is [array]) {
        return $document.value
    }
    throw "Offline collection JSON must be an array or contain a value array: $Path"
}

function Get-SmartWorkplaceCMDBSourceStatePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $current = $Path + '.status.json.txt'
    if (Test-Path -LiteralPath $current -PathType Leaf) { return $current }
    return ($Path + '.status.json')
}

function Resolve-SmartWorkplaceCMDBCollectionPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [switch]$Fixture, [int]$MaxItems,
        [switch]$ExplicitDataRoot, [switch]$NoWrite)
    $kind = if ($Fixture) { 'Fixture' } else { 'Live' }
    if ($MaxItems -gt 0) { $kind += '-Bounded' }
    $isTest = $Fixture -or $MaxItems -gt 0
    $root = $Paths.DataRootPath
    $markerPath = Join-Path $root '.collection-root.json.txt'
    if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
        $markerPath = Join-Path $root '.collection-root.json'
    }
    $marker = if (Test-Path -LiteralPath $markerPath) { Read-SmartWorkplaceCMDBJsonFile $markerPath } else { $null }
    $matches = $null -ne $marker
    if ($matches) {
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            if ([string]$marker.$name -ne [string]$Paths.$name) { $matches = $false }
        }
        if ([string]$marker.Kind -ne $kind) { $matches = $false }
    }
    if (-not $isTest -and $null -ne $marker -and -not $matches) {
        throw 'Collection root identity or mode mismatch. Choose a dedicated live data root.'
    }
    if ($isTest) {
        $occupied = (Test-Path -LiteralPath $root) -and @(Get-ChildItem -LiteralPath $root -Force | Select-Object -First 1).Count -gt 0
        if (-not $matches -and (-not $ExplicitDataRoot -or $occupied)) {
            $base = if ($ExplicitDataRoot) { $root } else { Join-Path $Paths.ProjectRootPath 'Data' }
            $root = Join-Path $base ('TestRuns\{0}_{1}' -f $kind, [guid]::NewGuid().ToString('N'))
        }
        $Paths = Resolve-SmartWorkplaceCMDBTenantPath -Tenant $Paths.ProfileKey -OrganizationKey $Paths.OrganizationKey -EnvironmentKey $Paths.EnvironmentKey -TenantKey $Paths.TenantKey -TenantId $Paths.TenantId -DataRootPath $root
    }
    if (-not $NoWrite) {
        $document = [ordered]@{Version=1; Kind=$kind}
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) { $document[$name]=$Paths.$name }
        Write-SmartWorkplaceCMDBJsonAtomically -InputObject $document -Path (Join-Path $root '.collection-root.json.txt')
    }
    return $Paths
}

function Start-SmartWorkplaceCMDBSourceCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [Parameter(Mandatory)][string[]]$RawPath,
        [switch]$Fixture, [int]$MaxItems, [switch]$Scoped, [switch]$NoWrite)
    $coverage = if ($MaxItems -gt 0) { 'Bounded' } elseif ($Fixture) { 'Fixture' } elseif ($Scoped) { 'Scoped' } else { 'Complete' }
    $run = [pscustomobject]@{Paths=$Paths; RawPath=$RawPath; Fixture=[bool]$Fixture; Coverage=$coverage; MaxItems=$MaxItems; RunId=[guid]::NewGuid().ToString('N'); StartedUtc=[datetime]::UtcNow.ToString('o'); NoWrite=[bool]$NoWrite; Locks=(New-Object 'System.Collections.Generic.List[object]'); PreviousState=@{}}
    if ($NoWrite) { return $run }
    try {
        foreach ($path in $RawPath) {
            New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force | Out-Null
            $run.Locks.Add([IO.File]::Open(($path + '.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None))
        }
        foreach ($path in $RawPath) {
            $statePath = Get-SmartWorkplaceCMDBSourceStatePath -Path $path
            $previous = [pscustomobject]@{
                Exists = $false
                Bytes = $null
                IsUsable = $false
                SHA256 = ''
            }
            if (Test-Path -LiteralPath $statePath -PathType Leaf) {
                $previous.Exists = $true
                $previous.Bytes = [IO.File]::ReadAllBytes($statePath)
                try {
                    $state = Read-SmartWorkplaceCMDBJsonFile -Path $statePath
                }
                catch {
                    $state = $null
                }
                if ($null -ne $state -and $state.Status -eq 'Completed' -and
                    (Test-Path -LiteralPath $path -PathType Leaf)) {
                    $previous.SHA256 = [string]$state.SHA256
                    $previous.IsUsable = (
                        [string]$state.SHA256 -eq
                        [string](Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
                    )
                }
            }
            $previous | Add-Member -NotePropertyName Path -NotePropertyValue $statePath
            $run.PreviousState[$path] = $previous
            Write-SmartWorkplaceCMDBSourceState -Run $run -Path $path -Status 'InProgress'
        }
    } catch { foreach ($handle in $run.Locks) { $handle.Dispose() }; throw }
    return $run
}

function Write-SmartWorkplaceCMDBSourceState {
    param($Run, [string]$Path, [string]$Status, [string[]]$NotCollectedPath = @())
    $state = [ordered]@{Version=1; SourceName=[IO.Path]::GetFileName($Path); RunId=$Run.RunId; Status=$Status; Coverage=$Run.Coverage; Mode=$(if ($Run.Fixture) {'Fixture'} else {'Live'}); MaxItems=$Run.MaxItems; StartedUtc=$Run.StartedUtc; CompletedUtc=''; RowCount=$null; SHA256=''}
    foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) { $state[$name]=$Run.Paths.$name }
    if ($Status -eq 'Completed') {
        $state.CompletedUtc=[datetime]::UtcNow.ToString('o')
        $state.RowCount=(Import-Csv -LiteralPath $Path | Measure-Object).Count
        $state.SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        if ($Path -in $NotCollectedPath) { $state.Coverage='NotCollected' }
    }
    Write-SmartWorkplaceCMDBJsonAtomically -InputObject $state -Path ($Path + '.status.json.txt')
}

function Complete-SmartWorkplaceCMDBSourceCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Run, [switch]$Failed, [string[]]$NotCollectedPath=@())
    if ($Run.NoWrite) { return }
    try {
        foreach ($path in $Run.RawPath) {
            $restored = $false
            if ($Failed -and $Run.PSObject.Properties['PreviousState'] -and
                $Run.PreviousState.ContainsKey($path)) {
                $previous = $Run.PreviousState[$path]
                if ($previous.Exists -and $previous.IsUsable -and
                    (Test-Path -LiteralPath $path -PathType Leaf)) {
                    try {
                        $snapshotMatches = (
                            [string]$previous.SHA256 -eq
                            [string](Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
                        )
                        if ($snapshotMatches) {
                            $statePath = $path + '.status.json.txt'
                            $temporaryStatePath = $statePath + '.restore.' + [guid]::NewGuid().ToString('N')
                            [IO.File]::WriteAllBytes($temporaryStatePath, $previous.Bytes)
                            Move-Item -LiteralPath $temporaryStatePath -Destination $statePath -Force
                            $restored = $true
                        }
                    }
                    catch {
                        $restored = $false
                    }
                }
            }
            if (-not $restored) {
                Write-SmartWorkplaceCMDBSourceState -Run $Run -Path $path -Status $(if ($Failed) {'Failed'} else {'Completed'}) -NotCollectedPath $NotCollectedPath
            }
        }
    } finally { foreach ($handle in $Run.Locks) { $handle.Dispose() } }
}

function Publish-SmartWorkplaceCMDBSourceCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(ParameterSetName = 'Objects', Mandatory)][AllowEmptyCollection()][object[]]$InputObject,
        [Parameter(ParameterSetName = 'CsvFile', Mandatory)][string]$InputCsvPath,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$HistoryPath,
        [Parameter(Mandatory)][string]$LatestPath,
        [Parameter(Mandatory)][string]$ContractPath,
        [Parameter(Mandatory)][string]$ContractTableName
    )

    if ($Run.NoWrite) {
        throw 'Publish-SmartWorkplaceCMDBSourceCsv cannot publish a NoWrite collection run.'
    }

    $contract = Get-SmartWorkplaceCMDBTableContract -Path $ContractPath
    $table = @($contract.tables | Where-Object name -eq $ContractTableName)
    if ($table.Count -ne 1) {
        throw "Source contract table '$ContractTableName' was not found exactly once."
    }

    $transactionRoot = Join-Path $Run.Paths.DataRootPath (
        '.staging\Sources\{0}' -f $Run.RunId
    )
    $stagedLatestRoot = Join-Path $transactionRoot 'DATA-LAST'
    $stagedPath = Join-Path $stagedLatestRoot (
        Join-Path ([string]$table[0].area) ([string]$table[0].name)
    )
    $backupRoot = Join-Path $transactionRoot 'backups'
    $promotions = @(
        [pscustomobject]@{
            Source = $stagedPath
            Destination = [IO.Path]::GetFullPath($HistoryPath)
            Backup = Join-Path $backupRoot ('history-' + [IO.Path]::GetFileName($HistoryPath))
        },
        [pscustomobject]@{
            Source = $stagedPath
            Destination = [IO.Path]::GetFullPath($LatestPath)
            Backup = Join-Path $backupRoot ('latest-' + [IO.Path]::GetFileName($LatestPath))
        }
    )
    $completedPromotions = New-Object System.Collections.Generic.List[object]

    try {
        if ($PSCmdlet.ParameterSetName -eq 'CsvFile') {
            $resolvedInputCsvPath = [IO.Path]::GetFullPath($InputCsvPath)
            if (-not (Test-Path -LiteralPath $resolvedInputCsvPath -PathType Leaf)) {
                throw "Prepared source CSV was not found: $resolvedInputCsvPath"
            }

            $sourceHeader = Get-Content -LiteralPath $resolvedInputCsvPath -TotalCount 1 -ErrorAction Stop
            $actualColumns = if ([string]::IsNullOrWhiteSpace($sourceHeader)) {
                @()
            }
            else {
                @($sourceHeader.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
            }
            $expectedColumns = @($Columns)
            if (($actualColumns -join [char]31) -cne ($expectedColumns -join [char]31)) {
                throw "Prepared source CSV '$resolvedInputCsvPath' does not use the exact contract column order for '$ContractTableName'."
            }

            $stagedFolder = Split-Path -Path $stagedPath -Parent
            New-Item -ItemType Directory -Path $stagedFolder -Force | Out-Null
            Copy-Item -LiteralPath $resolvedInputCsvPath -Destination $stagedPath -Force
        }
        else {
            Export-SmartWorkplaceCMDBCsv `
                -InputObject @($InputObject) `
                -Columns $Columns `
                -Path $stagedPath `
                -TenantKey $Run.Paths.TenantKey `
                -OrganizationKey $Run.Paths.OrganizationKey `
                -EnvironmentKey $Run.Paths.EnvironmentKey `
                -TenantId $Run.Paths.TenantId
        }

        $stagedResults = @(Test-SmartWorkplaceCMDBCsvContract `
                -LatestOutputRootPath $stagedLatestRoot `
                -ContractPath $ContractPath)
        $stagedTable = @($stagedResults | Where-Object Name -eq $ContractTableName)
        if ($stagedTable.Count -ne 1 -or $stagedTable[0].Status -ne 'Valid') {
            throw "Staged source CSV '$ContractTableName' does not satisfy its contract."
        }

        foreach ($promotion in $promotions) {
            $destinationFolder = Split-Path $promotion.Destination -Parent
            New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
            $hadPrevious = Test-Path -LiteralPath $promotion.Destination -PathType Leaf
            if ($hadPrevious) {
                New-Item -ItemType Directory -Path (Split-Path $promotion.Backup -Parent) -Force | Out-Null
                Copy-Item -LiteralPath $promotion.Destination -Destination $promotion.Backup -Force
            }
            $candidate = $promotion.Destination + '.candidate.' + $Run.RunId
            Copy-Item -LiteralPath $promotion.Source -Destination $candidate -Force
            Move-Item -LiteralPath $candidate -Destination $promotion.Destination -Force
            $completedPromotions.Add([pscustomobject]@{
                    Destination = $promotion.Destination
                    Backup = $promotion.Backup
                    HadPrevious = $hadPrevious
                })
        }

        $defaultLatestPath = Join-Path $Run.Paths.LatestOutputRootPath (
            Join-Path ([string]$table[0].area) ([string]$table[0].name)
        )
        if ([IO.Path]::GetFullPath($defaultLatestPath) -eq [IO.Path]::GetFullPath($LatestPath)) {
            $publishedResults = @(Test-SmartWorkplaceCMDBCsvContract `
                    -LatestOutputRootPath $Run.Paths.LatestOutputRootPath `
                    -ContractPath $ContractPath)
            $publishedTable = @($publishedResults | Where-Object Name -eq $ContractTableName)
            if ($publishedTable.Count -ne 1 -or $publishedTable[0].Status -ne 'Valid') {
                throw "Published source CSV '$ContractTableName' does not satisfy its contract."
            }
        }

        Complete-SmartWorkplaceCMDBSourceCollection -Run $Run
        return [pscustomobject]@{
            HistoryPath = [IO.Path]::GetFullPath($HistoryPath)
            LatestPath = [IO.Path]::GetFullPath($LatestPath)
            ContractTableName = $ContractTableName
        }
    }
    catch {
        for ($index = $completedPromotions.Count - 1; $index -ge 0; $index--) {
            $promotion = $completedPromotions[$index]
            if ($promotion.HadPrevious -and
                (Test-Path -LiteralPath $promotion.Backup -PathType Leaf)) {
                Copy-Item -LiteralPath $promotion.Backup -Destination $promotion.Destination -Force
            }
            elseif (Test-Path -LiteralPath $promotion.Destination -PathType Leaf) {
                Remove-Item -LiteralPath $promotion.Destination -Force
            }
        }
        throw
    }
    finally {
        if (Test-Path -LiteralPath $transactionRoot) {
            Remove-Item -LiteralPath $transactionRoot -Recurse -Force
        }
    }
}

function Import-SmartWorkplaceCMDBSourceCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LiteralPath, [Parameter(Mandatory)]$Paths)
    $statePath = Get-SmartWorkplaceCMDBSourceStatePath -Path $LiteralPath
    if (Test-Path -LiteralPath $statePath) {
        $state = Read-SmartWorkplaceCMDBJsonFile $statePath
        foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
            if ([string]$state.$name -ne [string]$Paths.$name) { throw 'Source evidence identity mismatch.' }
        }
        if ($state.Status -ne 'Completed' -or
            $state.SHA256 -ne (Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).Hash) {
            throw "Unusable source snapshot: '$LiteralPath'. Recollect before normalization."
        }
    }
    Import-Csv -LiteralPath $LiteralPath -ErrorAction Stop
}

function Assert-SmartWorkplaceCMDBCollectionPage {
    [CmdletBinding()]
    param([AllowNull()]$Response)
    $value = $null
    if ($Response -is [System.Collections.IDictionary]) {
        if ($Response.Contains('value')) { $value = $Response['value'] }
    } elseif ($null -ne $Response -and $null -ne $Response.PSObject.Properties['value']) {
        $value = $Response.value
    }
    if ($null -eq $value -or $value -is [string] -or $value -isnot [System.Collections.IList]) {
        throw 'Invalid collection page: expected a value array. An unavailable response is not an empty snapshot.'
    }
}

function Get-SmartWorkplaceCMDBSourceHealth {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [datetimeoffset]$ReferenceDateTime=[datetimeoffset]::UtcNow,
        [int]$WarningHours=48, [int]$CriticalHours=168)
    $contract = Get-SmartWorkplaceCMDBTableContract (Join-Path $Paths.ProjectRootPath 'Schema/SmartWorkplaceCMDB.raw.tables.json')
    foreach ($table in $contract.tables) {
        $path=Join-Path $Paths.LatestOutputRootPath (Join-Path $table.area $table.name)
        $statePath=Get-SmartWorkplaceCMDBSourceStatePath -Path $path
        $health='Unknown'; $coverage='Unknown'; $completed=''; $severity='Warning'; $rowCount=$null
        if (Test-Path -LiteralPath $statePath) {
            try {
                $state=Read-SmartWorkplaceCMDBJsonFile $statePath
                foreach ($name in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId')) {
                    if ([string]$state.$name -ne [string]$Paths.$name) { throw 'Source evidence identity mismatch.' }
                }
                $health=[string]$state.Status; $coverage=[string]$state.Coverage; $rowCount=$state.RowCount
                # PS7 can materialize JSON dates as DateTime; a plain string cast
                # loses their offset and fractional seconds. Round-trip typed
                # dates while retaining the string path used by Windows PS5.1.
                $completed = if ($state.CompletedUtc -is [datetime] -or $state.CompletedUtc -is [datetimeoffset]) {
                    $state.CompletedUtc.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                } else { [string]$state.CompletedUtc }
                if ($health -eq 'Completed') {
                    $snapshotMatches=(Test-Path -LiteralPath $path) -and $state.SHA256 -eq (Get-FileHash -LiteralPath $path).Hash
                    $health=$coverage
                    $date=[datetimeoffset]::MinValue
                    if (-not $snapshotMatches) { $health='SnapshotMismatch'; $severity='Critical' }
                    elseif (-not [datetimeoffset]::TryParse($completed,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$date)) { $health='InvalidDate' }
                    elseif ($date -gt $ReferenceDateTime.AddMinutes(5)) { $health='FutureDate' }
                    elseif (($ReferenceDateTime-$date).TotalHours -gt $WarningHours) {
                        $health='Stale'; $severity=if (($ReferenceDateTime-$date).TotalHours -gt $CriticalHours) {'Critical'} else {'Warning'}
                    }
                    elseif ($coverage -eq 'Complete') { $severity='Information' }
                }
            } catch { $health='InvalidEvidence'; $severity='Critical' }
        }
        [pscustomobject]@{SourceName=$table.name; Status=$health; Coverage=$coverage; CompletedUtc=$completed; RowCount=$rowCount; Severity=$severity; HasEvidence=(Test-Path -LiteralPath $statePath)}
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCrnWwpG2VvnKo2
# P+urhACd3rRpPQWtOqyNoJ8BnJZkr6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBYa8FAGz6zvaAbJ2EYiTtN
# cui7tbXM7BBBBqwKCQ88njANBgkqhkiG9w0BAQEFAASCAYA2a8PIRyyi1oJiJLRw
# S4Ph95XbHGiBxPixkmrqe+P4ecvmiahLEUHiCpRtx1QDquJWQxukWWyuKb1c4uwT
# 7PpAb89KoIGiNxAZ6JMPTX6N3ga3LCof7mlX46sF2edD36lEYyIU7VNVUNcf7Apn
# FrmsOHXFbLWGpZh81zUAXFPIPIRkxfIOMZj/NCgjG+o/laeXQRxh6TB/aZz0Lbh1
# 6BWFBwuB162SNHuyBy/5DJjRrdP+WCCN9KskzL23oNLdVgpCWXvHAX6wOTswUO1w
# kVAPYopQPo6yBheLczK3eE1OHBUPl2J+AuyTcPrKhE4jvtt6uu18I1LwcjdZ5aUX
# PvrpSDxpYdgkxyQUwfLP3ZEMrM4xboP68RqgohoMSZqEZANEld6P13XW/wyd9IBI
# dCiGUmN3n8dFpKqcpxTLGccdESa6v28mIdsNl9hCeFOvxOYgnFJr3MsmX4LJ/8+x
# EDSA2mQQ5Wa5J/2/kP2K6TfDzPA0V7k3j0UsQ/7BPH4yGB8=
# SIG # End signature block
