# Optional stable V1 projection. Source-reported hardware never defines CI identity.
function Export-CIHardwareContext {
    param([string]$Path, $Identity, $Mapping, [string]$ProjectRoot, [string]$Stage, [bool]$ValidateOnly)
    $hashes=@{}
    $schemaPath=Join-Path $ProjectRoot 'Schema\SmartWorkplaceCMDB.ci.hardware.json'
    $rawSchemaPath=Join-Path $ProjectRoot 'Schema\SmartWorkplaceCMDB.hardware.tables.json'
    $statePath=$Path+'.status.json.txt'
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { $statePath=$Path+'.status.json' }
    foreach ($file in @($Path,$statePath,$schemaPath,$rawSchemaPath)) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw 'Missing hardware input, state or contract.' }
        $hashes[$file]=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    }
    $schema=Get-Content -LiteralPath $schemaPath -Raw | ConvertFrom-Json
    $rawSchema=Get-Content -LiteralPath $rawSchemaPath -Raw | ConvertFrom-Json
    foreach ($definition in @($schema,$rawSchema)) {
        if ($definition.channel -cne 'stable' -or $definition.contractVersion -cne '1.0.0') { throw 'Hardware contracts must be the frozen V1 stable contracts.' }
    }
    $rawColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','SourceSystem','ManagedDeviceId','AzureAdDeviceId','SerialNumber','SerialNumberStatus','Manufacturer','ManufacturerStatus','Model','ModelStatus','TotalStorageSpaceInBytes','StorageStatus','SourceCollectedDateTime')
    $columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','CI_ID','SourceSystem','ManagedDeviceId','AzureAdDeviceId','SerialNumber','SerialNumberStatus','Manufacturer','ManufacturerStatus','Model','ModelStatus','TotalStorageSpaceInBytes','StorageStatus','HardwareCollectedDateTime','InventoryCollectedDateTime','CollectionCoverage','CollectionMode')
    if ($schema.name -cne 'CMDB_CIDeviceHardware.csv' -or ($schema.columns -join ',') -cne ($columns -join ',') -or
        @($rawSchema.tables).Count -ne 1 -or ($rawSchema.tables[0].columns -join ',') -cne ($rawColumns -join ',')) { throw 'Unsupported hardware column contract.' }
    Test-CIExactHeader -Path $Path -Columns $rawColumns
    $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    foreach ($field in $Identity.Keys) {
        if ([string]$state.$field -cne [string]$Identity[$field]) { throw 'Hardware state tenant mismatch.' }
    }
    if ($state.Version -ne 1 -or $state.SourceName -cne [IO.Path]::GetFileName($Path) -or
        [string]::IsNullOrWhiteSpace([string]$state.RunId) -or $state.Status -cne 'Completed' -or
        [string]$state.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $state.SHA256 -ine $hashes[$Path]) { throw 'Hardware state is not a matching Completed snapshot.' }
    $expectedCount=[long]0; $limit=[long]0
    if ([string]$state.RowCount -cnotmatch '^\d+$' -or -not [long]::TryParse([string]$state.RowCount,[ref]$expectedCount) -or
        [string]$state.MaxItems -cnotmatch '^\d+$' -or -not [long]::TryParse([string]$state.MaxItems,[ref]$limit)) { throw 'Invalid hardware state row count or limit.' }
    if (($state.Mode -ceq 'Fixture' -and $state.Coverage -cnotin @('Fixture','Bounded')) -or
        ($state.Mode -ceq 'Live' -and $state.Coverage -cnotin @('Complete','Bounded')) -or
        $state.Mode -cnotin @('Fixture','Live') -or
        ($state.Coverage -ceq 'Bounded' -and ($limit -le 0 -or $expectedCount -gt $limit)) -or
        ($state.Coverage -cne 'Bounded' -and $limit -ne 0)) { throw 'Invalid hardware coverage or mode.' }
    $times=@{}
    foreach ($field in @('StartedUtc','CompletedUtc')) {
        $value=$state.$field
        $text=if ($value -is [datetime] -or $value -is [datetimeoffset]) { $value.ToString('o',[Globalization.CultureInfo]::InvariantCulture) } else { [string]$value }
        $qualified=ConvertTo-CIQualifiedTime -Value $text
        if ($qualified.Status -cne 'UTCQualified') { throw 'Invalid hardware state timestamp.' }
        $times[$field]=[datetimeoffset]::Parse($qualified.Utc,[Globalization.CultureInfo]::InvariantCulture)
    }
    $validatedUtc=[datetimeoffset]::UtcNow
    if ($times.StartedUtc -gt $times.CompletedUtc -or $times.CompletedUtc -gt $validatedUtc.AddMinutes(5)) { throw 'Hardware state timestamps are reversed or in the future.' }
    $seen=@{}; $ciSeen=@{}; $serials=@{}; $missing=[ordered]@{SerialNumber=0;Manufacturer=0;Model=0;Storage=0}
    $count=0; $zero=0; $writer=$null; $rowDate=''
    try {
        if (-not $ValidateOnly) {
            $writer=[IO.StreamWriter]::new((Join-Path $Stage $schema.name),$false,[Text.UTF8Encoding]::new($true))
            $writer.WriteLine($columns -join ',')
        }
        Import-Csv -LiteralPath $Path | ForEach-Object {
            $row=$_
            foreach ($field in $Identity.Keys) { if ([string]$row.$field -cne [string]$Identity[$field]) { throw 'Hardware row tenant mismatch.' } }
            if ($row.SourceSystem -cne 'MicrosoftIntune') { throw 'Unexpected hardware SourceSystem.' }
            $native=([string]$row.ManagedDeviceId).Trim().ToLowerInvariant()
            if (-not $native -or $seen.ContainsKey($native)) { throw 'Missing or duplicate hardware native ID.' }
            if (-not $Mapping.ContainsKey($native)) { throw 'Unmapped hardware native ID.' }
            $source=$Mapping[$native]
            $correlation=([string]$row.AzureAdDeviceId).Trim().ToLowerInvariant()
            if (-not $correlation -or $correlation -ceq '00000000-0000-0000-0000-000000000000') { $correlation='intune:'+$native }
            if ($correlation -cne $source.Correlation) { throw 'Hardware correlation differs from mapped inventory.' }
            foreach ($attribute in @('SerialNumber','Manufacturer','Model')) {
                $status=[string]$row.PSObject.Properties[$attribute+'Status'].Value; $value=[string]$row.$attribute
                if (($status -ceq 'Missing' -and $value -cne '') -or
                    ($status -ceq 'Reported' -and [string]::IsNullOrWhiteSpace($value)) -or
                    $status -cnotin @('Missing','Reported')) { throw 'Inconsistent hardware text qualification.' }
                if ($status -ceq 'Missing') { $missing[$attribute]++ }
            }
            $capacity=[long]0; $storage=[string]$row.TotalStorageSpaceInBytes
            if ($row.StorageStatus -ceq 'Missing') {
                if ($storage -cne '') { throw 'Missing hardware storage has a value.' }
                $missing.Storage++
            } else {
                if ($storage -cnotmatch '^\d+$' -or -not [long]::TryParse($storage,[ref]$capacity) -or
                    ($row.StorageStatus -ceq 'Reported' -and $capacity -le 0) -or
                    ($row.StorageStatus -ceq 'ZeroReported' -and $capacity -ne 0) -or
                    $row.StorageStatus -cnotin @('Reported','ZeroReported')) { throw 'Invalid hardware storage value or qualification.' }
                if ($capacity -eq 0) { $zero++ }
            }
            $date=ConvertTo-CIQualifiedTime -Value ([string]$row.SourceCollectedDateTime)
            if ($date.Status -cne 'UTCQualified' -or
                [datetimeoffset]::Parse($date.Utc,[Globalization.CultureInfo]::InvariantCulture) -gt $times.CompletedUtc -or
                ($count -gt 0 -and $date.Utc -cne $rowDate)) { throw 'Invalid or mixed hardware collection timestamp.' }
            $rowDate=$date.Utc
            $seen[$native]=$true; $ciSeen[$source.CI_ID]=$true; $count++
            if ($row.SerialNumberStatus -ceq 'Reported') {
                $serial=([string]$row.SerialNumber).Trim().ToLowerInvariant()
                if (-not $serials.ContainsKey($serial)) { $serials[$serial]=0 }; $serials[$serial]++
            }
            $record=[ordered]@{}
            foreach ($field in $Identity.Keys) { $record[$field]=$Identity[$field] }
            $record.CI_ID=$source.CI_ID
            foreach ($field in $rawColumns[4..14]) { $record[$field]=[string]$row.$field }
            $record.HardwareCollectedDateTime=[string]$row.SourceCollectedDateTime
            $record.InventoryCollectedDateTime=$source.Collected
            $record.CollectionCoverage=[string]$state.Coverage; $record.CollectionMode=[string]$state.Mode
            if ($null -ne $writer) { $writer.WriteLine((@([pscustomobject]$record | ConvertTo-Csv -NoTypeInformation)[1])) }
        }
        if ($count -ne $expectedCount) { throw 'Hardware state row count mismatch.' }
    } finally { if ($null -ne $writer) { $writer.Dispose() } }
    foreach ($file in $hashes.Keys) { if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -cne $hashes[$file]) { throw 'Hardware input or contract changed during build.' } }
    return [ordered]@{
        Status='Validated';ContractVersion=$schema.contractVersion;RowCount=$count;CIsWithHardware=$ciSeen.Count
        InventoryNativeDeviceCount=$Mapping.Count;InventoryNativeDevicesWithoutHardware=$Mapping.Count-$count
        MissingAttributes=$missing;ZeroReportedStorageCount=$zero
        DuplicateSerialValueCount=@($serials.Values | Where-Object { $_ -gt 1 }).Count
        Coverage=[string]$state.Coverage;Mode=[string]$state.Mode;RunId=[string]$state.RunId
        HardwareCollectedUtc=$rowDate;CompletedUtc=$times.CompletedUtc.UtcDateTime.ToString('o')
        ValidatedUtc=$validatedUtc.UtcDateTime.ToString('o');AttributeFreshness='NotAssessed';InputHashes=$hashes
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBKoaFiOF5cE15c
# Bby1ODOwZx2u4vvmetx8G5wgYceYuKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAT/jybBoDz1uaDr2v7/gCc
# e/4/V72P8v19Vk+P2mm+fzANBgkqhkiG9w0BAQEFAASCAYAbDwtgY9KHylo1tZ0s
# 6YqAEGXKn5+lICoCR5APUE6naqGG8O1rOWhD4L5vMCag0DwTGofS+ndUqEnwOV5n
# p45HsKMhmBy0hIJfAR4sBw225jLUUCRKxyXJKNPp8bLCcUJeD/ihAeelrSvPX/c9
# uUtH7PeIF4yp3joLkXinLRt2l7EjxT4oY+yU+WBZ0Cg5RtHnHcEpwxPumsk4BT48
# 38phGJi3Y6FmOWbFR0XPU1FayuxqwJJfGui8jClozo26fcQ/3fQVAeOuqUVpMC+7
# HRbvd2RBsMR2grvTRTSteA/IngIHSSOmnnyaeVwjjCjd5x13jLMcOCOZTb8K1kAl
# RHaOBL+HNlt9tnbGkWLVQ0n60QoWW5t7okEdiZUSLQD/MrWCAnQ1ApOSG8FMNNEi
# ZgqE8JDAxdn28IudgUFOQ/ozaHBbyR5+V2+Fc4TMPWAozD83mo19muHd18SAYhSY
# uJ5yly4pqowAdudJHUjtc0rePAqiC9OdIDKp0t5Wdj7g3jg=
# SIG # End signature block
