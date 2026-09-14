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

function Resolve-SmartWorkplaceCMDBCollectionPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Paths, [switch]$Fixture, [int]$MaxItems,
        [switch]$ExplicitDataRoot, [switch]$NoWrite)
    $kind = if ($Fixture) { 'Fixture' } else { 'Live' }
    if ($MaxItems -gt 0) { $kind += '-Bounded' }
    $isTest = $Fixture -or $MaxItems -gt 0
    $root = $Paths.DataRootPath
    $markerPath = Join-Path $root '.collection-root.json'
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
        Write-SmartWorkplaceCMDBJsonAtomically -InputObject $document -Path (Join-Path $root '.collection-root.json')
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
            $statePath = $path + '.status.json'
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
    Write-SmartWorkplaceCMDBJsonAtomically -InputObject $state -Path ($Path + '.status.json')
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
                            $statePath = $path + '.status.json'
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
    $statePath = $LiteralPath + '.status.json'
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
        $statePath=$path+'.status.json'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBT+D9b5F/OzRFL
# PGmrqK+24Zd60oEg4S59wcqhev+qB6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEcSCe23kbeIMjxdehrENjRY26/Pl0HZ2D5OuH0ct6LlMA0GCSqG
# SIb3DQEBAQUABIIBgDDyjATM/oGFYFeynLw+ogYz79bVbTGBeqUZFAwS5gy0pPZ1
# xd+jTHMzcoR/K3sCY6xwV1lC0G5OOBQ1gzLp5/53kwcErOKiHN9gPoCyk7vkhPt9
# NXjJX5bO4zc2D2Ds/OTAuaRR0W+pQ33KJLK5kv4LM1sTmwkeXp0aPfwSEowW5z3O
# VEQBbAX4wMJZu/6+mz1X138d3u2ONGi383wPWX9AZDtcnHec/vwRn1TV4Jic5E5e
# 4EJ92tqeXPoNZsncnec4+ojT4siU++D3H2b5Rb6cLaPMfpsbEYGvtIjkRvGC92+Q
# ZCIS2UTyWwPysFsyJPT9QpKYGAvx0e2U/I9bUrEhx3a85cHw+6An92EJWI7PanD7
# GCJCdOBITDTLHyXWmEjG7VHff2XPUfGqtWUf0bocJhdwKgNcWnHngDnTse1GPeKq
# qK7WzXlokgN0foch5H9MxIWUtpUS2Z6fvHk7G7hR7es2oAbHFWd/XLZpNNSUVdz/
# sx/+iWqLFAvUX8v6B6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTQwNDI0
# NTZaMC8GCSqGSIb3DQEJBDEiBCAoy8vwkMezvhRZ3Lk97Lfc/ThLpY9Gi7G6kz9+
# NcmSWjANBgkqhkiG9w0BAQEFAASCAgAaQHTrrmTZ4KiqN32DD5ke9/DOdV+LIMdV
# Es5QJ1Nj1nTejD+DHGgM6TDNB6zQTOFyC945+bAqQG34VS+/h0ua+zYVftUqH99j
# ATMBK0BmDeaXI05BVQ63L+Rs8FOP1nFMfSTodt6EI8dzFyRvboctRwc4X0hGYvuT
# MV2orKUQJ76nv6WSRODBzjwIgUAAoRATt6uxkI5yg+IW/x3CFE9xLUzqOgjOXgsv
# VJixeUEqUzb5WCPQmLfloE+CyGF1w/Sc0V7xAl2+pxK6hVkyo16PhtS1rKNNEFeE
# JiQkxZPQitZhkpg2xqxHjwha5LejCkcBf6TfvgHbE9JNkJoPJgrHpYnhzY1MlixU
# Zsul8Ch4DfYfBzaYwewLVtwWXFy6Wi4QhOXu+9q66pg2v/XTBNH4ThaCB2kx/Ees
# d9C+BWEqAg1dDkZfziS7Sj+UhhPM8Q37FKLZmEyZLHJgfzkhVchhKh1SieA+OCaA
# Hf5BXDfRJ9mm9CUeplOTUH1ZrwMCACQZZFLVLV97BctwzAc9uU5aAkmHBxGqx9pH
# sEJsDu6qSKnh8CQPFtJoSVx/Ec08ubFjUcJpLXSQhojENcxFUZne6YOmSPGXDCS+
# YqRyVJwOFZR8vjqh8ZmyXEMflA80YFxfBjLRSKewxJANwXUj16KK3a6kSYmwrzlC
# R8RnoJDAZQ==
# SIG # End signature block
