#Requires -Version 7.0
<#
.SYNOPSIS
Offline byte equivalence, validation, atomic-copy and bounded performance tests.
.VERSION
1.0.0
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core globals in an isolated offline test process.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Only unique temporary fixtures and module-scoped mocks are changed.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Module-scoped copy and serialization mocks; real cmdlets are explicitly qualified.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline mocks retain the real APIs but record only the fields required by each assertion.')]
param([ValidateRange(0,100000)][int]$BenchmarkRows=0)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CsvParity-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$results=[Collections.Generic.List[object]]::new()
$checks=0
function Assert-Csv([bool]$Condition,[string]$Message) { if (-not $Condition) { throw $Message }; $script:checks++ }
function Test-CsvCase([string]$Name,[scriptblock]$Body) {
    try { & $Body; $results.Add([pscustomobject]@{Name=$Name;Passed=$true;Error=''}) }
    catch { $results.Add([pscustomobject]@{Name=$Name;Passed=$false;Error=$_.Exception.Message}) }
}
$names=@('Get-SmartM365CoreContextValue','Add-SmartM365TenantKeyToCsvData',
    'Get-SmartM365CsvValidationRule','Get-SmartM365CsvValidationBaseName','Assert-SmartM365CsvDataCompleteness',
    'Write-SmartM365CsvAtomically','Write-SmartM365PreparedCsvAtomically','Copy-SmartM365FileAtomically',
    'Publish-SmartM365Csv','Export-SmartM365Csv',
    'Get-SmartM365MaxItemsValue','Test-SmartM365MaxItemsMode','Get-SmartM365MaxItemsSuffix',
    'Limit-SmartM365RowsForMaxItems','Add-SmartM365MaxItemsSuffixToCsvPath')
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psm1'),[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$definitions=foreach($name in $names) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if ($null -eq $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
$module=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
try {
    & $module {
        function script:WriteLog { param($Message,$Level) }
        function script:RemoveOldFiles { param($Path,$Filter,$KeepCount,$LogFile) $script:RetentionCalls++ }
        function script:Invoke-SmartM365SharePointCsvUpload {
            param($LocalFilePath)
            $script:Uploads.Add($LocalFilePath)
            [pscustomobject]@{LocalFilePath=$LocalFilePath}
        }
        function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param($SourceFiles,$TimestampedPath) $script:WeeklyCalls++ }
        function script:Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath,$RetentionDays) $script:RemoteRetentionCalls++ }
        function script:Copy-Item {
            param($LiteralPath,$Destination,[switch]$Force)
            if ($script:ProbeSourceMutation) {
                $guarded=$false
                try { $writer=[IO.File]::Open($LiteralPath,'Open','Write','ReadWrite');$writer.Dispose() }
                catch { $guarded=$true }
                if (-not $guarded) { throw 'Validated source was writable during verified copy.' }
            }
            Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force -ErrorAction Stop
            if ($script:CorruptCopy) { Add-Content -LiteralPath $Destination -Value 'synthetic corruption' -Encoding utf8 }
        }
        function script:Export-Csv {
            param($Path,$Encoding,$Delimiter,[switch]$NoTypeInformation,[Parameter(ValueFromPipeline)]$InputObject)
            begin { $script:Serializations++;$rows=[Collections.Generic.List[object]]::new() }
            process { $rows.Add($InputObject) }
            end {
                if ($script:FailSerialization) { Set-Content -LiteralPath $Path -Value 'synthetic partial write';throw 'Synthetic serialization failure.' }
                $rows.ToArray() | Microsoft.PowerShell.Utility\Export-Csv -LiteralPath $Path -Encoding $Encoding -Delimiter $Delimiter -NoTypeInformation:$NoTypeInformation -ErrorAction Stop
            }
        }
        $script:SmartM365CoreTenantKey='contoso-test'
        $script:SmartM365CoreOrganizationKey='contoso';$script:SmartM365CoreEnvironmentKey='test'
        $script:SmartM365CoreTenantId='00000000-0000-0000-0000-000000000001'
        $script:Serializations=0;$script:RetentionCalls=0;$script:WeeklyCalls=0;$script:RemoteRetentionCalls=0
        $script:Uploads=[Collections.Generic.List[string]]::new()
        $script:CorruptCopy=$false;$script:ProbeSourceMutation=$false;$script:FailSerialization=$false
        $global:SmartM365MaxItems=0;$global:SmartM365TestMaxItems=0;$global:SmartM365IsMaxItemsRun=$false
        $global:csvGeneratedPaths=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $global:LogTextFile='';$global:RetentionMaxCSV=0
        $rule=@{Name='Synthetic_Members';CriticalFields=@('GroupId','MemberId','MembershipKind','CollectionStatus','RunId','CollectedAtUtc')
            RequiredColumns=@('Label');UniqueFields=@('GroupId','MemberId');AllowEmptyDataset=$true
            CriticalMissingFailMinRows=1;CriticalMissingFailPercent=0}
        $global:SmartM365CsvValidationRules=@{Synthetic_Members=$rule}
        $global:SmartM365RequireCsvValidationRules=$true
    }
    $columns=@('GroupId','MemberId','MembershipKind','CollectionStatus','RunId','CollectedAtUtc','Label')
    $rows=@(
        [pscustomobject][ordered]@{GroupId='ABC';MemberId='001';MembershipKind='Direct';CollectionStatus='Collected';RunId='current-run';CollectedAtUtc='2026-01-01T12:00:00.1234567Z';Label="accent $([char]0xE9), quote `"x`"`r`nsecond line"},
        [pscustomobject][ordered]@{GroupId='ABC';MemberId='002';MembershipKind='Direct';CollectionStatus='Collected';RunId='current-run';CollectedAtUtc='2026-01-01T12:00:00.1234567Z';Label='=literal;not a formula'}
    )
    $before=ConvertTo-Json -InputObject $rows -Depth 10 -Compress
    foreach($encoding in @('UTF8','utf8BOM','Unicode','ASCII')) {
        foreach($delimiter in @(',',';')) {
            Test-CsvCase "Byte parity: $encoding / $delimiter" {
                $folder=Join-Path $testRoot ([guid]::NewGuid().ToString('N'))
                New-Item -ItemType Directory -Path $folder | Out-Null
                $old=Join-Path $folder 'old/Synthetic_Members.csv';$oldLast=Join-Path $folder 'old-last/Synthetic_Members.csv'
                $new=Join-Path $folder 'new/Synthetic_Members.csv';$newLast=Join-Path $folder 'new-last/Synthetic_Members.csv'
                $counts=& $module {
                    param($Data,$Columns,$Old,$OldLast,$New,$NewLast,$Encoding,$Delimiter)
                    $script:Serializations=0
                    Publish-SmartM365Csv -Data $Data -Columns $Columns -TimestampedPath $Old -LatestPath $OldLast -Encoding $Encoding -Delimiter $Delimiter -NoSharePointUpload -NoWeeklyHistory | Out-Null
                    $legacy=$script:Serializations;$script:Serializations=0
                    Publish-SmartM365Csv -Data $Data -Columns $Columns -TimestampedPath $New -LatestPath $NewLast -Encoding $Encoding -Delimiter $Delimiter -NoSharePointUpload -NoWeeklyHistory -SingleSerialization | Out-Null
                    [pscustomobject]@{Legacy=$legacy;Optimized=$script:Serializations}
                } $rows $columns $old $oldLast $new $newLast $encoding $delimiter
                $hash=(Get-FileHash -LiteralPath $old -Algorithm SHA256).Hash
                foreach($path in @($oldLast,$new,$newLast)) { Assert-Csv ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ceq $hash) 'Legacy and optimized bytes differ.' }
                Assert-Csv ($counts.Legacy -eq 2 -and $counts.Optimized -eq 1) 'Serialization count changed unexpectedly.'
                $parsed=@(Import-Csv -LiteralPath $newLast -Delimiter $delimiter -Encoding $encoding)
                Assert-Csv ($parsed.Count -eq 2 -and $parsed[0].MemberId -ceq '001') 'Row count or leading zeros changed.'
                Assert-Csv ($parsed[0].CollectedAtUtc -ceq $rows[0].CollectedAtUtc -and $parsed[1].Label -ceq $rows[1].Label) 'Native dates or text changed.'
                Assert-Csv (($parsed[0].PSObject.Properties.Name -join ',') -ceq ('TenantKey,OrganizationKey,EnvironmentKey,TenantId,'+($columns -join ','))) 'Column order changed.'
            }
        }
    }
    Test-CsvCase 'Schema-only empty export is byte-identical' {
        $old=Join-Path $testRoot 'empty-old/Synthetic_Members.csv';$new=Join-Path $testRoot 'empty-new/Synthetic_Members.csv';$last=Join-Path $testRoot 'empty-last/Synthetic_Members.csv'
        & $module {param($c,$o,$n,$l)
            Publish-SmartM365Csv -Data @() -Columns $c -TimestampedPath $o -NoSharePointUpload -NoWeeklyHistory | Out-Null
            Publish-SmartM365Csv -Data @() -Columns $c -TimestampedPath $n -LatestPath $l -NoSharePointUpload -NoWeeklyHistory -SingleSerialization | Out-Null
        } $columns $old $new $last
        Assert-Csv ((Get-FileHash $old).Hash -ceq (Get-FileHash $last).Hash) 'Empty schema bytes changed.'
        Assert-Csv (@(Import-Csv $last).Count -eq 0) 'Empty schema invented rows.'
    }
    foreach($invalid in @('Duplicate','MissingField','MissingColumn','TenantConflict','EmptyNotAllowed','MissingRule')) {
        Test-CsvCase "Quality gate preserves both previous files: $invalid" {
            $folder=Join-Path $testRoot $invalid;New-Item -ItemType Directory -Path $folder | Out-Null
            $stamp=Join-Path $folder 'Synthetic_Members.csv';$last=Join-Path $folder 'latest/Synthetic_Members.csv'
            New-Item -ItemType Directory -Path (Split-Path $last) | Out-Null
            Set-Content -LiteralPath $stamp -Value 'OLD HISTORY';Set-Content -LiteralPath $last -Value 'OLD LATEST'
            $bad=@($rows | ForEach-Object { $_.PSObject.Copy() });$badColumns=@($columns)
            switch($invalid) {
                Duplicate { $bad=@($bad[0],$bad[0]) }
                MissingField { $bad[0].MemberId='' }
                MissingColumn { $badColumns=@($columns | Where-Object { $_ -ne 'Label' });foreach($row in $bad){$row.PSObject.Properties.Remove('Label')} }
                TenantConflict { $bad[0] | Add-Member TenantKey 'another-tenant' }
                EmptyNotAllowed { $bad=@() }
                MissingRule { }
            }
            $rejected=& $module {param($d,$c,$s,$l,$case)
                $rule=$global:SmartM365CsvValidationRules.Synthetic_Members
                $previousEmpty=$rule.AllowEmptyDataset
                if($case -eq 'EmptyNotAllowed'){$rule.AllowEmptyDataset=$false}
                if($case -eq 'MissingRule'){$global:SmartM365CsvValidationRules=@{}}
                try {Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $s -LatestPath $l -SingleSerialization -NoSharePointUpload -NoWeeklyHistory | Out-Null;return $false}
                catch{return $true}
                finally {$rule.AllowEmptyDataset=$previousEmpty;$global:SmartM365CsvValidationRules=@{Synthetic_Members=$rule}}
            } $bad $badColumns $stamp $last $invalid
            Assert-Csv $rejected 'Invalid dataset was accepted.'
            Assert-Csv ((Get-Content $stamp -Raw).Trim() -ceq 'OLD HISTORY' -and (Get-Content $last -Raw).Trim() -ceq 'OLD LATEST') 'Invalid rows replaced a valid CSV.'
        }
    }
    foreach($case in @('LockedDestination','CorruptedCopy','SourceGuard')) {
        Test-CsvCase "Verified atomic copy: $case" {
            $source=Join-Path $testRoot "$case-source.csv";$target=Join-Path $testRoot "$case-target.csv"
            Set-Content -LiteralPath $source -Value 'NEW BYTES';Set-Content -LiteralPath $target -Value 'OLD BYTES'
            $lock=$null
            if($case -eq 'LockedDestination'){$lock=[IO.File]::Open($target,'Open','Read','Read')}
            try {
                $rejected=& $module {param($s,$d,$case)
                    $script:CorruptCopy=$case -eq 'CorruptedCopy';$script:ProbeSourceMutation=$case -eq 'SourceGuard'
                    try {Copy-SmartM365FileAtomically -SourcePath $s -DestinationPath $d -VerifyHash;return $false}
                    catch {return $true}
                    finally {$script:CorruptCopy=$false;$script:ProbeSourceMutation=$false}
                } $source $target $case
                Assert-Csv ($rejected -eq ($case -ne 'SourceGuard')) 'Incorrect verified copy outcome.'
            }
            finally {if($null -ne $lock){$lock.Dispose()}}
            if($case -eq 'SourceGuard'){Assert-Csv ((Get-FileHash $source).Hash -ceq (Get-FileHash $target).Hash) 'Guarded copy changed bytes.'}
            else {Assert-Csv ((Get-Content $target -Raw).Trim() -ceq 'OLD BYTES') 'Failed copy damaged latest output.'}
            Assert-Csv (@(Get-ChildItem $testRoot -Filter '*.tmp').Count -eq 0) 'Verified copy retained temporary files.'
        }
    }
    foreach($fault in @('Serialization','Copy')) {
        Test-CsvCase "Publication failure is blocking and keeps latest: $fault" {
            $folder=Join-Path $testRoot "publisher-$fault";New-Item -ItemType Directory -Path $folder | Out-Null
            $stamp=Join-Path $folder 'Synthetic_Members.csv';$last=Join-Path $folder 'latest/Synthetic_Members.csv'
            New-Item -ItemType Directory -Path (Split-Path $last) | Out-Null
            Set-Content $stamp 'OLD HISTORY';Set-Content $last 'OLD LATEST'
            $observed=& $module {param($d,$c,$s,$l,$fault)
                $script:Uploads.Clear();$script:WeeklyCalls=0;$script:RetentionCalls=0
                $script:CorruptCopy=$fault -eq 'Copy';$script:FailSerialization=$fault -eq 'Serialization'
                $rejected=$false
                try {Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $s -LatestPath $l -SingleSerialization | Out-Null}
                catch {$rejected=$true}
                finally {$script:CorruptCopy=$false;$script:FailSerialization=$false}
                [pscustomobject]@{Rejected=$rejected;Uploads=$script:Uploads.Count;Weekly=$script:WeeklyCalls;Retention=$script:RetentionCalls}
            } $rows $columns $stamp $last $fault
            Assert-Csv ($observed.Rejected -and $observed.Uploads -eq 0 -and $observed.Weekly -eq 0 -and $observed.Retention -eq 0) 'Failed pair was published or triggered retention.'
            Assert-Csv ((Get-Content $last -Raw).Trim() -ceq 'OLD LATEST') 'Failed pair replaced latest output.'
            if($fault -eq 'Serialization'){Assert-Csv ((Get-Content $stamp -Raw).Trim() -ceq 'OLD HISTORY') 'Failed serialization replaced timestamped output.'}
            else{Assert-Csv (@(Import-Csv $stamp).Count -eq 2) 'Copy failure lost the valid timestamped evidence.'}
            Assert-Csv (@(Get-ChildItem $folder -Recurse -Filter '*.tmp').Count -eq 0 -and @(Get-ChildItem $folder -Filter 'Synthetic_Members.*.csv').Count -eq 0) 'Publication retained a partial temporary file.'
        }
    }
    Test-CsvCase 'MAXITEMS isolates canonical output' {
        $stamp=Join-Path $testRoot 'max-stamp/Synthetic_Members.csv';$last=Join-Path $testRoot 'max-last/Synthetic_Members.csv'
        New-Item -ItemType Directory -Path (Split-Path $last) | Out-Null;Set-Content $last 'CANONICAL'
        $published=& $module {param($d,$c,$s,$l)
            $global:SmartM365MaxItems=1
            try {Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $s -LatestPath $l -SingleSerialization -NoSharePointUpload -NoWeeklyHistory}
            finally {$global:SmartM365MaxItems=0}
        } $rows $columns $stamp $last
        Assert-Csv ($published.LatestPath -like '*_MAXITEMS-1.csv' -and @((Import-Csv $published.LatestPath)).Count -eq 1) 'Partial output isolation failed.'
        Assert-Csv ((Get-Content $last -Raw).Trim() -ceq 'CANONICAL') 'MAXITEMS overwrote canonical output.'
    }
    Test-CsvCase 'SharePoint, history, retention and output tracking preserved' {
        $stamp=Join-Path $testRoot 'tracked/Synthetic_Members_20260101_120000.csv';$last=Join-Path $testRoot 'tracked-last/Synthetic_Members.csv'
        $observed=& $module {param($d,$c,$s,$l)
            $script:Uploads.Clear();$script:RetentionCalls=0;$script:WeeklyCalls=0;$script:RemoteRetentionCalls=0
            $published=Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $s -LatestPath $l -SingleSerialization -RetentionMaxCsv 2
            [pscustomobject]@{Published=$published;Uploads=$script:Uploads.Count;Retention=$script:RetentionCalls;Weekly=$script:WeeklyCalls;RemoteRetention=$script:RemoteRetentionCalls;Tracked=$global:csvGeneratedPaths.Contains($s) -and $global:csvGeneratedPaths.Contains($l)}
        } $rows $columns $stamp $last
        Assert-Csv ($observed.Uploads -eq 2 -and $observed.Published.SharePointUploads.Count -eq 2) 'Both SharePoint uploads were not preserved.'
        Assert-Csv ($observed.Retention -eq 1 -and $observed.Weekly -eq 1 -and $observed.RemoteRetention -eq 1 -and $observed.Tracked) 'History, retention or receipt tracking changed.'
    }
    Test-CsvCase 'Same-file publication and tenant-neutral export' {
        $path=Join-Path $testRoot 'same/Synthetic_Members.csv'
        & $module {param($d,$c,$p)Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $p -LatestPath $p -SingleSerialization -NoSharePointUpload -NoWeeklyHistory -NoTenantKey | Out-Null} $rows $columns $path
        $row=@(Import-Csv $path)[0]
        Assert-Csv (($row.PSObject.Properties.Name -join ',') -ceq ($columns -join ',')) 'Neutral export gained tenant columns.'
    }
    Assert-Csv ((ConvertTo-Json -InputObject $rows -Depth 10 -Compress) -ceq $before) 'Source objects were mutated.'
    $collector=Get-Content (Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1') -Raw
    Assert-Csv ($collector -match '-NoWeeklyHistory -SingleSerialization' -and $collector -match "-MinimumVersion '1.0.71'") 'WorkplaceScope optimization guard missing.'
    if($BenchmarkRows -gt 0) {
        $timings=& $module {param($count,$folder,$c)
            $d=@(for($i=0;$i -lt $count;$i++){[pscustomobject]@{GroupId='group';MemberId=[string]$i;MembershipKind='Direct';CollectionStatus='Collected';RunId='synthetic';CollectedAtUtc='2026-01-01T00:00:00Z';Label='synthetic'}})
            $measurements=foreach($optimized in @($false,$true)) {
                $s=Join-Path $folder ("benchmark-$optimized/Synthetic_Members.csv");$l=Join-Path $folder ("benchmark-$optimized-last/Synthetic_Members.csv")
                $watch=[Diagnostics.Stopwatch]::StartNew()
                Publish-SmartM365Csv -Data $d -Columns $c -TimestampedPath $s -LatestPath $l -SingleSerialization:$optimized -NoSharePointUpload -NoWeeklyHistory | Out-Null
                $watch.Stop()
                [pscustomobject]@{Optimized=$optimized;Rows=$count;Milliseconds=[math]::Round($watch.Elapsed.TotalMilliseconds);SHA256=(Get-FileHash $l).Hash}
            }
            return @($measurements)
        } $BenchmarkRows $testRoot $columns
        Assert-Csv ($timings[0].SHA256 -ceq $timings[1].SHA256) 'Benchmark outputs differ.'
        $timings | Format-Table -AutoSize
    }
    $results | Format-Table -AutoSize
    if(@($results | Where-Object {-not $_.Passed}).Count){throw 'Single serialization regression failed.'}
    Write-Output ("PASS: {0} cases, {1} checks; synthetic files and mocked external actions only." -f $results.Count,$checks)
}
finally {
    Remove-Module $module -Force
    $resolved=[IO.Path]::GetFullPath($testRoot);$parent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($parent,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe test cleanup path.'}
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLmJQ3G7vpVX8M
# APPZjggzC5ch1TMPPKVyON+PPTsqB6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBQE/f9GxZR7JH/E9T3D+JjQaOHkKr6Kl0i8yPGCepqLMA0GCSqG
# SIb3DQEBAQUABIIBgKLXF07tAs+4rLLXi0oMnq81qZshUofPr11pkVjh3AsICkap
# UmVGfavy29SsXbaJj33KAENZ/P0JNX4go76Dbc4ly1htSKrWqNgTcpoW7C1q6ONo
# O1qbw4UfAvkyUt580UyAd9FJt0mrAD9j58Ivud9HKTR7rbFJ8HPgpuovcRPk84dF
# uHdXNyaFSX3qfSj7nhMobVg8t+PIDqjya985Xko8IJQhJH6tXO3rIxYtB+vbwaBA
# 0OIZaxJjvHLwqGmhLv85j8c59IFYJruu4rJcBreJ1wWyrgOqXSTRxhyodx4GNwZA
# nP+WtW7ENoho0vRnjPMxPFZu7tXfh2lYLB8XYbh+E3R1Jh3x0aurLmey0sq/0W+8
# su8qoXH9G3O1zNy60BwAPJqQamtORMkceY5qHl/Y47U0FiVFqn8CjxlNs1tp+J7l
# cBKXm8eorIG6eD2cb9UNBGMx2etVwEMPFgThpCnClA9SewHbG4SlZ2ckn5NYHwkT
# 1WqhapDGa3TPfoF2TqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQyMjAw
# NDhaMC8GCSqGSIb3DQEJBDEiBCCRAvUzRwd3K5cnFPxsJvcH4hzruT15H+EQB1ok
# pd5i+TANBgkqhkiG9w0BAQEFAASCAgC1feece1osmxA1NDYm2HTfOGJRO9Keefil
# GYm37gFjVwhFtnf3S4+ntyTJMxe1zygT29nKsF73SMN1p/kqaIJec6brR+N9cxSU
# mLxsUon4mBxkHAAo02SUl7x9aOxFDTrk4yp5/JP2F/gHQPk3fPYQ0yWqWrMTM1Qv
# 2qcpRx7/Gn6bLLLf3l9P+Q20cxUqhqBgkqP8LflDh41CnRvg7j4U6HgOg5q0LVCQ
# 98LcHaZKIzpaNbbhIbruOFAYwvKu1qAAoTidgdjghY97THr9c9NvC5SPug8bPJk9
# /gC6B5XUxM9Z5X6khlk1JnRwb/9UQ8D6bepPxmrphNkfsp0DVixxIBEiidZwNl9G
# tK9wCmMT1d71w/le0utE+8Gt/PBD7nVh4n5Er8Hsb8Nop4EiVAaOSt0DgGHbbk4Z
# 7mrZ509nxIRycD3y4goVzGyacISqHnRHqE3a4RFQe1SXjR0QJI7CsN3c/5cFt7kv
# 2Nd547v/jSb7diP33AJzYRG/8Nat91i4kli13qmis4nwTzNa2NBzJG3tjj3V4uIG
# MqQd1yZ4by9z8PsVBz3K6blxZqsBRz/ht3lokHiIRBW/DPWU+5KmDYAzRVWk3D01
# jyN3eZ7wvQML3FVOlp9O6PcQ5TLtGMMcmXzjQvY7d8CuYVhTI2RCQE1G1ELO8v9J
# YR5WWP5rag==
# SIG # End signature block
