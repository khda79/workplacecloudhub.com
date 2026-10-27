# Version 1.1.0: explicit historical artifact cohorts.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Get-PreparedArtifactContract {
    Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json
}

function Get-PreparedLegacyArtifactNames {
    param([string]$ContractSHA256)
    if (-not $ContractSHA256) { return @() }
    $contract=Get-PreparedArtifactContract
    if (-not (Test-PreparedArtifactContractCompatible $ContractSHA256)) { throw 'Artifact contract mismatch.' }
    if ($contract.PSObject.Properties['legacyCurrentFiles']) {
        foreach ($entry in $contract.legacyCurrentFiles) {
            if ($ContractSHA256.ToUpperInvariant() -cin @($entry.contractSHA256)) { 'Source-'+$entry.file }
        }
    }
}

function Get-PreparedOptionalArtifactFiles {
    param([Parameter(Mandatory)]$Contract)
    if(-not $Contract.PSObject.Properties['optionalCurrentFiles']){return @()}
    $names=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($entry in @($Contract.optionalCurrentFiles)){
        if([string]::IsNullOrWhiteSpace($entry.file) -or [IO.Path]::GetFileName([string]$entry.file) -cne [string]$entry.file -or
           -not $names.Add($entry.file) -or $entry.file -cin @($Contract.currentFiles) -or [int]$entry.maxAgeHours -lt 1 -or [int]$entry.maxAgeHours -gt 8760){
            throw 'Invalid or duplicate optional artifact file.'
        }
        if($entry.PSObject.Properties['delimiter'] -and [string]$entry.delimiter -cnotin @(',', ';')){
            throw 'Invalid optional artifact delimiter.'
        }
    }
    @($Contract.optionalCurrentFiles)
}

function Assert-PreparedArtifactTenant {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)]$Specification)
    $delimiter=if($Specification.PSObject.Properties['delimiter']){[string]$Specification.delimiter}else{','}
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try{
        $parser.SetDelimiters($delimiter)
        $parser.HasFieldsEnclosedInQuotes=$true
        $parser.TrimWhiteSpace=$false
        $header=$parser.ReadFields()
        if(-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or
           @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count){throw 'Empty or duplicate optional artifact header.'}
        $tenantIndex=[array]::IndexOf($header,'TenantKey')
        if($tenantIndex -lt 0){throw "Optional artifact TenantKey missing: $($Specification.file)"}
        while(-not $parser.EndOfData){
            $fields=$parser.ReadFields()
            if($fields.Length -ne $header.Length){throw "Malformed optional artifact record: $($Specification.file)"}
            if([string]::IsNullOrWhiteSpace($fields[$tenantIndex]) -or $fields[$tenantIndex] -cne $TenantKey){
                throw "Optional artifact tenant mismatch: $($Specification.file)"
            }
        }
    }finally{$parser.Dispose()}
}

function Get-PreparedArtifactSources {
    param([Parameter(Mandatory)][string]$DataRoot)
    $contract=Get-PreparedArtifactContract
    foreach($name in $contract.currentFiles){@{Relative="DATA-LAST/$name";Kind='Current'}}
    # Optional current files: missing or stale inputs are optional source issues, never a failed capture.
    foreach($entry in @(Get-PreparedOptionalArtifactFiles $contract)){@{Relative="DATA-LAST/$($entry.file)";Kind='Current';IsOptional=$true;MaxAgeHours=[int]$entry.maxAgeHours}}
    foreach($spec in $contract.history){
        $root=Join-Path $DataRoot $spec.root
        $weeks=@(Get-ChildItem -LiteralPath $root -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$' | Sort-Object Name)
        if(-not $weeks.Count){throw "Required migration history missing: $($spec.kind)"}
        foreach($week in $weeks){
            if([int]$week.Name.Substring(6) -notin 1..53){throw 'Invalid history week.'}
            $relative="$($spec.root)/$($week.Name)"
            @{Relative="$relative/$($spec.file)";Kind='History'}
            $metadata=if(Test-Path -LiteralPath (Join-Path $week.FullName 'manifest.json.txt')){'manifest.json.txt'}else{'manifest.json'}
            @{Relative="$relative/$metadata";Kind='History'}
            if($spec.PSObject.Properties['optionalFile'] -and (Test-Path -LiteralPath (Join-Path $week.FullName $spec.optionalFile))){
                @{Relative="$relative/$($spec.optionalFile)";Kind='History'}
            }
        }
    }
}

function Test-PreparedArtifactName {
    param([Parameter(Mandatory)][string]$Name,[string]$ContractSHA256)
    $contract=Get-PreparedArtifactContract
    if ($Name -cin @(Get-PreparedLegacyArtifactNames $ContractSHA256)) { return $true }
    if($Name -ceq 'PreparedHistoryIndex.json.txt'){return $true}
    if($Name -cin @($contract.currentFiles | ForEach-Object {"Source-$_"})){return $true}
    if($Name -cin @(Get-PreparedOptionalArtifactFiles $contract | ForEach-Object {"Source-$($_.file)"})){return $true}
    foreach($spec in $contract.history){
        $prefix='History-'+$spec.kind+'-\d{4}-W(?:0[1-9]|[1-4][0-9]|5[0-3])-'
        $names=@($spec.file,'manifest.json.txt')
        if($spec.PSObject.Properties['optionalFile']){$names+=@($spec.optionalFile)}
        foreach($leaf in $names){if($Name -cmatch ('^'+$prefix+[regex]::Escape($leaf)+'$')){return $true}}
    }
    return $false
}

function Copy-PreparedArtifacts {
    # Only the already verified capture is read. No second live capture or publication.
    param([Parameter(Mandatory)][string]$SnapshotRoot,[Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$TenantKey)
    $contract=Get-PreparedArtifactContract
    $entries=[Collections.Generic.List[object]]::new()
    $index=[Collections.Generic.List[object]]::new()
    function Copy-One([string]$Relative,[string]$Name){
        if(-not (Test-PreparedArtifactName $Name)){throw "Unapproved artifact: $Name"}
        $source=Join-Path $SnapshotRoot $Relative
        $target=Join-Path $Destination $Name
        if(Test-Path -LiteralPath $target){throw "Artifact output already exists: $Name"}
        $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        Copy-Item -LiteralPath $source -Destination $target -ErrorAction Stop
        if((Get-FileHash -LiteralPath $target).Hash -ne $hash -or (Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Captured artifact changed during publication.'}
        $entries.Add([pscustomobject]@{File=$Name;Bytes=(Get-Item -LiteralPath $target).Length;SHA256=$hash;SourceRelative=$Relative})
    }
    $snapshot=Get-Content -LiteralPath (Join-Path $SnapshotRoot 'DATA-LAST/M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
    if($snapshot.TenantKey -cne $TenantKey -or $snapshot.SchemaVersion -notin @(1,2) -or $snapshot.SnapshotId -notmatch '^[a-fA-F0-9]{32}$'){throw 'Invalid license report snapshot identity.'}
    # Match the existing Power Query reader before committing the unified pointer.
    if($snapshot.Products -isnot [array] -or $snapshot.Products.Count -ne 4 -or
       $snapshot.OtherProducts -isnot [array] -or $snapshot.OtherProducts.Count -ne 3){throw 'Invalid license report snapshot structure.'}
    if($snapshot.SchemaVersion -eq 2){
        $issues=$snapshot.AssignmentIssues
        if($issues -isnot [pscustomobject] -or $issues.Available -isnot [bool] -or $issues.Rows -isnot [array] -or
           ($issues.Available -and $issues.Counts -isnot [pscustomobject]) -or
           (-not $issues.Available -and $issues.Rows.Count -ne 0)){throw 'Invalid license assignment snapshot structure.'}
    }
    foreach($name in $contract.currentFiles){
        if($name -like 'M365_Licenses_Report*.csv'){
            # Another tenant is fatal. A row of another snapshot is a lineage defect: the report is
            # copied and declared Not qualified by Get-PreparedLicenseReportQualification.
            foreach($row in Import-Csv -LiteralPath (Join-Path $SnapshotRoot "DATA-LAST/$name")){
                if($row.TenantKey -cne $TenantKey){throw "License report tenant mismatch: $name"}
            }
        }
        Copy-One "DATA-LAST/$name" "Source-$name"
    }
    foreach($entry in @(Get-PreparedOptionalArtifactFiles $contract)){
        $relative="DATA-LAST/$($entry.file)"
        if(-not (Test-Path -LiteralPath (Join-Path $SnapshotRoot $relative) -PathType Leaf)){continue}
        Assert-PreparedArtifactTenant -Path (Join-Path $SnapshotRoot $relative) -TenantKey $TenantKey -Specification $entry
        Copy-One $relative "Source-$($entry.file)"
    }
    foreach($spec in $contract.history){
        foreach($week in Get-ChildItem -LiteralPath (Join-Path $SnapshotRoot $spec.root) -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$' | Sort-Object Name){
            $relative="$($spec.root)/$($week.Name)"
            $prefix="History-$($spec.kind)-$($week.Name)-"
            $metadata=if(Test-Path -LiteralPath (Join-Path $week.FullName 'manifest.json.txt')){'manifest.json.txt'}else{'manifest.json'}
            $manifest=Get-Content -LiteralPath (Join-Path $week.FullName $metadata) -Raw | ConvertFrom-Json
            if($manifest.PSObject.Properties['Week'] -and $manifest.Week -cne $week.Name){throw 'History manifest week mismatch.'}
            Copy-One "$relative/$metadata" "${prefix}manifest.json.txt"
            Copy-One "$relative/$($spec.file)" "$prefix$($spec.file)"
            $issues=$null
            if($spec.PSObject.Properties['optionalFile'] -and (Test-Path -LiteralPath (Join-Path $week.FullName $spec.optionalFile))){
                $issues="$prefix$($spec.optionalFile)"; Copy-One "$relative/$($spec.optionalFile)" $issues
            }
            $index.Add(@{Kind=$spec.kind;Week=$week.Name;SourceFile=$spec.file;File="$prefix$($spec.file)";Manifest="${prefix}manifest.json.txt";Issues=$issues})
        }
    }
    $indexPath=Join-Path $Destination 'PreparedHistoryIndex.json.txt'
    if(Test-Path -LiteralPath $indexPath){throw 'History index already exists.'}
    @{SchemaVersion=1;TenantKey=$TenantKey;Entries=$index.ToArray()} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $indexPath -Encoding utf8NoBOM
    $entries.Add([pscustomobject]@{File='PreparedHistoryIndex.json.txt';Bytes=(Get-Item -LiteralPath $indexPath).Length;SHA256=(Get-FileHash -LiteralPath $indexPath).Hash;SourceRelative='GeneratedFromVerifiedCapture'})
    $entries.ToArray()
}

function Assert-PreparedArtifacts {
    param([Parameter(Mandatory)][string]$BatchRoot,[Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Entries,
        [Parameter(Mandatory)][string]$TenantKey,[string]$ContractSHA256)
    $legacy=@(Get-PreparedLegacyArtifactNames $ContractSHA256)
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $optional=@(Get-PreparedOptionalArtifactFiles (Get-PreparedArtifactContract))
    foreach($entry in $Entries){
        if(-not (Test-PreparedArtifactName ([string]$entry.File) -ContractSHA256 $ContractSHA256) -or -not $seen.Add($entry.File)){throw 'Unsafe or duplicate prepared artifact.'}
        $path=Join-Path $BatchRoot $entry.File
        if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Artifact links are forbidden.'}
        if($entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or (Get-Item -LiteralPath $path).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $entry.SHA256){throw "Artifact integrity mismatch: $($entry.File)"}
        $specification=@($optional | Where-Object { "Source-$($_.file)" -ceq $entry.File })
        if($specification.Count -eq 1){Assert-PreparedArtifactTenant -Path $path -TenantKey $TenantKey -Specification $specification[0]}
    }
    $required=@('PreparedHistoryIndex.json.txt')+$legacy+@((Get-PreparedArtifactContract).currentFiles | ForEach-Object {"Source-$_"})
    foreach($name in $required){if(-not $seen.Contains($name)){throw "Required artifact missing: $name"}}
    $index=Get-Content -LiteralPath (Join-Path $BatchRoot 'PreparedHistoryIndex.json.txt') -Raw | ConvertFrom-Json
    if($index.SchemaVersion -ne 1 -or $index.TenantKey -cne $TenantKey){throw 'Invalid prepared history index.'}
    $referenced=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($name in $required){$null=$referenced.Add($name)}
    # Optional current files are declared by the contract, not by the history index.
    foreach($entry in @(Get-PreparedOptionalArtifactFiles (Get-PreparedArtifactContract))){if($seen.Contains("Source-$($entry.file)")){$null=$referenced.Add("Source-$($entry.file)")}}
    $weeks=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($entry in $index.Entries){
        if($entry.Kind -cnotin @('Mailbox','Windows') -or $entry.Week -notmatch '^\d{4}-W(?:0[1-9]|[1-4][0-9]|5[0-3])$' -or -not $weeks.Add("$($entry.Kind)/$($entry.Week)")){throw 'Invalid or duplicate history index entry.'}
        $spec=@((Get-PreparedArtifactContract).history | Where-Object kind -CEQ $entry.Kind)[0]
        $prefix="History-$($entry.Kind)-$($entry.Week)-"
        if($entry.SourceFile -cne $spec.file -or $entry.File -cne "$prefix$($spec.file)" -or $entry.Manifest -cne "${prefix}manifest.json.txt"){throw 'History index scope mismatch.'}
        if($entry.Issues -and (-not $spec.PSObject.Properties['optionalFile'] -or $entry.Issues -cne "$prefix$($spec.optionalFile)")){throw 'Invalid history issue artifact.'}
        foreach($name in @($entry.File,$entry.Manifest,$entry.Issues) | Where-Object {$_}){
            if(-not $seen.Contains($name)){throw 'History index references an undeclared artifact.'}
            $null=$referenced.Add($name)
        }
    }
    foreach($kind in 'Mailbox','Windows'){if(-not @($index.Entries | Where-Object Kind -CEQ $kind).Count){throw "Required indexed history missing: $kind"}}
    if($referenced.Count -ne $seen.Count){throw 'Unindexed historical artifact.'}
}

function Get-PreparedLicenseReportQualification {
    # The license report is consumed only when the current licensing receipt records it:
    # completed receipt, same tenant and producer, every report CSV hashed in Files and declared
    # Published, snapshot ReceiptRun.RunId equal to the receipt RunId. A lineage defect yields
    # "Not qualified" (empty tables in the report), never a failed batch or older values.
    param([Parameter(Mandatory)][string]$SnapshotRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Artifacts)
    $spec=(Get-PreparedArtifactContract).licenseReport
    $last=Join-Path $SnapshotRoot 'DATA-LAST'
    $hashes=@{}
    foreach($entry in $Artifacts){$hashes[[string]$entry.File]=[string]$entry.SHA256}
    $snapshot=Get-Content -LiteralPath (Join-Path $last $spec.snapshot) -Raw | ConvertFrom-Json
    $result=[ordered]@{Version=1;State='Qualified';Reason='';SnapshotId=[string]$snapshot.SnapshotId;ReceiptRunId='';ReceiptSHA256='';Files=@()}
    $receipt=$null;$receiptReason=''
    $receiptPath=Join-Path $last $spec.receipt
    if(-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)){$receiptReason='Licensing receipt is not part of the capture.'}
    else{
        try{
            $result.ReceiptSHA256=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash
            $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
            $result.ReceiptRunId=[string]$receipt.RunId
            if([string]$receipt.Producer -cne $spec.producer -or [string]$receipt.TenantKey -cne $TenantKey){$receiptReason='Licensing receipt belongs to another producer or tenant.'}
            elseif([string]$receipt.Status -cne 'Completed'){$receiptReason=('Licensing receipt is {0}, not Completed.' -f $receipt.Status)}
            elseif(-not $snapshot.PSObject.Properties['ReceiptRun'] -or $null -eq $snapshot.ReceiptRun){$receiptReason='Report snapshot names no licensing receipt run (produced outside a receipt run).'}
            elseif([string]$snapshot.ReceiptRun.Producer -cne $spec.producer -or [string]$snapshot.ReceiptRun.RunId -cne [string]$receipt.RunId){$receiptReason='Report snapshot belongs to another licensing receipt run.'}
        }catch{$receiptReason=('Licensing receipt is unreadable: {0}' -f $_.Exception.Message)}
    }
    $files=[Collections.Generic.List[object]]::new()
    foreach($name in @($spec.reportFiles)+@($spec.decisionFiles)){
        $decision=$name -cin @($spec.decisionFiles)
        $artifact="Source-$name"
        $state='Qualified';$reason=''
        if(-not $hashes.ContainsKey($artifact)){$state='Absent';$reason='Not part of the licensing output used by this batch.'}
        elseif($receiptReason){$state='Not qualified';$reason=$receiptReason}
        else{
            $record=@($receipt.Files | Where-Object {[string]$_.File -ceq $name})
            $declared=@($receipt.OptionalFiles | Where-Object {[string]$_.File -ceq $name -and $_.Published -eq $true})
            if($record.Count -ne 1 -or $declared.Count -ne 1){$state='Not qualified';$reason='File is not recorded as published by the current licensing receipt.'}
            elseif([string]$record[0].SHA256 -ine $hashes[$artifact]){$state='Not qualified';$reason='File hash differs from the current licensing receipt.'}
            else{
                foreach($row in Import-Csv -LiteralPath (Join-Path $last $name)){
                    if([string]$row.SnapshotId -cne [string]$snapshot.SnapshotId){$state='Not qualified';$reason='Rows belong to another report snapshot.';break}
                    if($decision -and [string]$row.ReceiptRunId -cne [string]$receipt.RunId){$state='Not qualified';$reason='Rows belong to another licensing receipt run.';break}
                }
            }
        }
        $files.Add([ordered]@{File=$name;Kind=$(if($decision){'Decision'}else{'Report'});State=$state;Reason=$reason;SHA256=$(if($hashes.ContainsKey($artifact)){$hashes[$artifact]}else{''})})
    }
    $core=@($files | Where-Object {$_.Kind -eq 'Report' -and $_.State -ne 'Qualified'})
    if($core.Count -gt 0){
        $result.State='Not qualified'
        $result.Reason=$(if($receiptReason){$receiptReason}else{'{0}: {1}' -f $core[0].File,$core[0].Reason})
        # Decision files are never shown without their qualified report.
        foreach($file in $files){if($file.Kind -eq 'Decision' -and $file.State -eq 'Qualified'){$file.State='Not qualified';$file.Reason='License report is not qualified.'}}
    }
    $result.Files=$files.ToArray()
    [pscustomobject]$result
}

function Test-PreparedArtifactContractCompatible {
    # A batch built with a previous artifact contract stays transferable when the current
    # contract lists the exact previous hash; retired files use a separate read-only cohort.
    param([Parameter(Mandatory)][string]$ContractSHA256)
    $path=Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt'
    if($ContractSHA256 -ieq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash){return $true}
    $contract=Get-PreparedArtifactContract
    return $contract.PSObject.Properties['compatibleContractSHA256'] -and $ContractSHA256 -cin @($contract.compatibleContractSHA256 | ForEach-Object {([string]$_).ToUpperInvariant()})
}

Export-ModuleMember -Function Get-PreparedArtifactSources,Copy-PreparedArtifacts,Assert-PreparedArtifacts,Test-PreparedArtifactName,Get-PreparedLicenseReportQualification,Test-PreparedArtifactContractCompatible

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD2qTj/JgOH4qzM
# q6JcBQCWW2vf5Rhx+BmqSs6YJRZWMqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKbBnj08U1PXQkC+ZK7NzxxbMDa8zOXdA9+2jcNoMSRqMA0GCSqG
# SIb3DQEBAQUABIIBgAIuwoeUyxB9LpDpM+oFKy3IAoQEEZ+vCz5CxDx8nZLnbO1Y
# 0tj/fRyzYhaBol/WHMk0FNaNST4CZnqf+cVGdWP/2n2roQLqi5WopSL/bVmUi+xp
# ISZs/EW7qVSBpWIf4+FjXyGu0rT9afek4AUkJUO6bg41N0DEA+svheuYfXLKUy/t
# e6LeaLThiPRx9bEHskp6E05QRpsy0+rtbiLo9K5BjHc9roCHfmh2YSzl9BXRBLE8
# oU+OEtnxQ4ITBF9JoHfRbzqLv1tMG9+ZlFFQ+qtiXZRniQU1bzIWnZqeMKGYJNCl
# qUUGAJ3QnOXV0VNGoROkqZ3LXE0HyKCrZ5klU8MA0ACs1eqlWZLA1D9OQMzL5Tal
# n7rUCjHBmhhYuvRcyNMGbQqNf2aQLIEPyc1lzmmwGLkngRVrBRiaKS9gcD6PhswR
# wijr/FqmGW0zHSqB6hYOhVCT+AsQU5Puw489Zaok+Q5UiTseSvGapiGqgBI9EjZQ
# 4z2/38Rr0F87NPEdDaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcwOTQx
# NTNaMC8GCSqGSIb3DQEJBDEiBCCBA/pyFBWQfrNc8U+tJG85LlEJDbp7ntziWBdM
# +S0JxjANBgkqhkiG9w0BAQEFAASCAgAjpfZ6ddN0xqQBZnmLAiJXQA+BMj9YPHUz
# OvLGHUW8+V6J5vsOZBAw6IjXYIq2cASWDD6VHEZzU4SN2dRBmeQp5SECrmv86KY4
# zjO3idSucRHsP9+33TwJKctKbijKMuGbPYhzZFtUdaYMwHwY4+rgT7eRnkrD8itN
# 4JNLhHN2P3Yo0iJf7Mm7atgl4R0ZSIkNwhRJAFqQI6f551p6JGu42aPq+24l/a8L
# OB3wiQ5zOk9KKW3cejrF/wdJZfGy6LEg76s06YYtf35R2CnN5Mt/X7Do9dv6y1XN
# WnBO8twGvZFo3ZjFoOcbADQ90w29aZ5lIBJ7fOcyymivWEiCsxDwkV6f4B2/2hWj
# nAkKPSUuqD1J5lIw2Jj+VviJSDIEAftDRmJOBGJ08p0WUAc4qS6hT268TfaKF763
# qZ9jYjUq8qLXR1TrF8ift/TXxvSNIhFnORuO6WXt+I7S9fN1E2mRgG1LUPY6KJ2U
# TJEy40bPEnNgpA8oCU4BHUezUTt68ZkRInY/hKImn+bFSXAJOG2kFgOIkhKJJJSI
# yw1RxKBdzv9PmASZLHrUWHpmHyHXCnH2EZDsa0IauDCpfoi/h9AqpJiZPlIMeE5D
# xa0Dt2SnCNGYMrASAsWsgDEjtCSJsfMDOneLE9o+R4meaoAdjObbiiJZ/HFn9x1D
# pCzpDtjIzA==
# SIG # End signature block
