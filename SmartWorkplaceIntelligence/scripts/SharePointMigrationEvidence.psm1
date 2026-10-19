# Private input capture and predecessor qualification only. No publication.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Assert-SharePointMigrationFileName([string]$Name,[switch]$AllowEmpty){
    if($AllowEmpty -and -not $Name){return}
    if(-not $Name -or $Name -in '.','..' -or [IO.Path]::GetFileName($Name) -cne $Name -or $Name -match '[/\\]'){
        throw 'Migration input must be a safe file name below its configured data root.'
    }
}

function Get-SharePointMigrationSourceItems {
    param([Parameter(Mandatory)][string]$DataRoot)
    $contract=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/sharepoint-migration-contract.json.txt') -Raw | ConvertFrom-Json
    foreach($producer in $contract.producers){
        Assert-SharePointMigrationFileName $producer.receipt
        $receiptPath=Join-Path (Join-Path $DataRoot 'DATA-LAST') $producer.receipt
        $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
        if($receipt.Owner -cne 'SmartInventory-SourceReceipt' -or $receipt.Producer -cne ($producer.producer+'.ps1') -or $receipt.Status -notin @('Completed','CompletedWithWarnings')){throw 'Invalid SharePoint migration source producer receipt.'}
        $names=@($receipt.Files.File)
        if(@($names | Sort-Object -Unique).Count -ne $names.Count){throw 'Duplicate SharePoint receipt payload file.'}
        foreach($required in $producer.requiredFiles){if($required -cnotin $names){throw 'SharePoint migration receipt is missing a required inventory.'}}
        @{Relative=('DATA-LAST/'+$producer.receipt);Kind='Receipt'}
        foreach($entry in $receipt.Files){
            Assert-SharePointMigrationFileName $entry.File
            if($entry.File -notlike '*.csv'){throw 'Unexpected non-CSV SharePoint receipt payload.'}
            @{Relative=('DATA-LAST/'+$entry.File);Kind='NativeCurrent'}
        }
    }
}

function Get-SharePointMigrationHistoryCapture {
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey)
    $pointerPath=Join-Path $OutputRoot 'current.json.txt'
    if(-not (Test-Path -LiteralPath $pointerPath -PathType Leaf)){
        if(Test-Path -LiteralPath (Join-Path $OutputRoot 'current.json')){throw 'Migrate legacy pointer transport before enabling migration history.'}
        return [pscustomobject]@{BatchId='__NONE__';Entry=$null}
    }
    $pointer=Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
    if($pointer.TenantKey -cne $TenantKey -or $pointer.BatchId -notmatch '^[A-Za-z0-9_-]+$'){throw 'Invalid predecessor identity.'}
    $batch=Join-Path (Join-Path $OutputRoot 'batches') $pointer.BatchId
    $manifestPath=Join-Path $batch 'batch.json.txt'
    if((Get-FileHash -LiteralPath $manifestPath).Hash -cne [string]$pointer.ManifestSHA256){throw 'Migration predecessor manifest fingerprint mismatch.'}
    $manifest=Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $validationPath=Join-Path $batch 'validation.json.txt'
    $validation=Get-Content -LiteralPath $validationPath -Raw | ConvertFrom-Json
    if($manifest.BatchId -cne $pointer.BatchId -or $manifest.TenantKey -cne $TenantKey -or $validation.Passed -ne $true -or $validation.SchemaOnly -ne $false){throw 'Predecessor is not a fully validated tenant batch.'}
    $entries=@($manifest.Files | Where-Object File -CEQ 'SharePointMigrationHistoryEvidence.csv')
    if($entries.Count -gt 1){throw 'Ambiguous migration history predecessor.'}
    if(-not $entries.Count){return [pscustomobject]@{BatchId=$pointer.BatchId;Entry=$null}}
    $entry=$entries[0]
    $path=Join-Path $batch $entry.File
    if((Get-Item -LiteralPath $path).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $path).Hash -cne [string]$entry.SHA256){throw 'Migration predecessor payload mismatch.'}
    $validated=@($validation.Files | Where-Object File -CEQ $entry.File)
    if($validated.Count -ne 1 -or $validated[0].SHA256 -cne $entry.SHA256){throw 'Predecessor history is missing its matching validation record.'}
    [pscustomobject]@{BatchId=$pointer.BatchId;Entry=[pscustomobject]@{Relative='MigrationHistory/SharePointMigrationHistoryEvidence.csv';SourcePath=$path;Kind='History';Bytes=$entry.Bytes;ModifiedUtc=(Get-Item -LiteralPath $path).LastWriteTimeUtc.ToString('o');SHA256=$entry.SHA256;IsOptional=$false;MaxAgeHours=$null}}
}

Export-ModuleMember -Function Assert-SharePointMigrationFileName,Get-SharePointMigrationHistoryCapture,Get-SharePointMigrationSourceItems

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCGsMf5l7nTeoi5
# RuSIqCdwUWXZofm3NRAdemvmMR6K4KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCClhqx11VM8qCGWe1ixjtFj
# 9tBQtVsdOzi4fch2I5MLZzANBgkqhkiG9w0BAQEFAASCAYAY9R988XUtUC8h5dly
# nYCIz/YfCtFKgbotzi/IsOHNdZK3LwCIW602RbsiI8BkEl1DJHxNC6O/GZZ+G3nq
# VUOEBKN5C75o6HmruZv81b3S4TCgwNYe/pP/7Mn5zOB4XoLxcfKlCuKnTDm/Kjqf
# WmFnB+lnyjdvo/gXCNY/7nKTfKPYs69VZHNgNW4YNABPZL5LBKxpnE3ZPuuEh/jN
# HYwvTotlSwJZjbUV2aNJ7c/CgJ4onxvBtFgxGz8jrE4dt4aXQxOnNEdzcXnV7S8w
# g1jyTbbrKNuDg/kvF4WtIbj+3qgqLi3sheUpMrD+FUPccf/QSpYS6Bxa87EDp4o4
# E7yWo0kYEny5bsbZHNn6K+MLTh9RlVnzSX/2UCA7HXW0rBZSb82irAofsu8g9S9A
# 4+8ouJ+whEMwj/C0xZkHaMayhhRQ2cUI2pE8PEvawhfaNI6veEW2ISZ+TS+7jiVs
# HtzMY6qON3TotM1e4kWxovHbJYKeTIXrOrXbRE4+uox1Ohc=
# SIG # End signature block
