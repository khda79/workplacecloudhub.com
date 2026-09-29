[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Root,
    [string]$ContractPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'config\prepared-evidence-contract.json'),
    [switch]$SchemaOnly,
    [string[]]$Tables,
    [string]$ReportPath
)
$ErrorActionPreference='Stop'
$contract=Get-Content -LiteralPath $ContractPath -Raw | ConvertFrom-Json
$results=[Collections.Generic.List[object]]::new()
$failures=[Collections.Generic.List[string]]::new()
foreach ($requested in $Tables) {
    if ($requested -notin $contract.tables.table) { $failures.Add("Unknown table requested: $requested") }
}
foreach($table in $contract.tables) {
    if($Tables -and $table.table -notin $Tables) {continue}
    $path=Join-Path $Root $table.file
    if(-not(Test-Path -LiteralPath $path -PathType Leaf)) { $failures.Add("Missing file: $($table.file)"); continue }
    $reader=[IO.File]::OpenText($path)
    try { $header=$reader.ReadLine() } finally { $reader.Dispose() }
    if(-not $header) { $failures.Add("No header: $($table.file)"); continue }
    $sample=($header + "`n" + $header | ConvertFrom-Csv)[0]
    $headers=@($sample.PSObject.Properties.Name)
    $missing=@($table.columns.name | Where-Object { $_ -notin $headers })
    if($missing.Count) { $failures.Add("Missing columns in $($table.file): $($missing -join ', ')"); continue }
    $state=@{Rows=0;DuplicateKeys=0;InvalidValues=0}
    $keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if(-not $SchemaOnly) {
        Import-Csv -LiteralPath $path | ForEach-Object {
            $row=$_; $state.Rows++
            if($table.uniqueKey.Count) {
                $values=@($table.uniqueKey | ForEach-Object { ([string]$row.$_).Trim() })
                if(@($values | Where-Object { $_ -eq '' }).Count) { $state.InvalidValues++ }
                $key=($values | ForEach-Object { $_.Length.ToString()+':'+$_ }) -join '|'
                if(-not $keys.Add($key)) { $state.DuplicateKeys++ }
            }
            foreach($column in $table.columns) {
                $value=[string]$row.($column.name)
                if([string]::IsNullOrWhiteSpace($value) -or $column.type -eq 'string') { continue }
                try {
                    switch($column.type) {
                        'int64' { $null=[long]::Parse($value,[Globalization.CultureInfo]::InvariantCulture) }
                        'double' { $n=[double]::Parse($value,[Globalization.CultureInfo]::InvariantCulture); if([double]::IsInfinity($n) -or [double]::IsNaN($n)) {throw 'Non-finite'} }
                        'decimal' { $null=[decimal]::Parse($value,[Globalization.CultureInfo]::InvariantCulture) }
                        'dateTime' { $null=[datetime]::Parse($value,[Globalization.CultureInfo]::InvariantCulture) }
                        'boolean' { $null=[bool]::Parse($value) }
                    }
                } catch { $state.InvalidValues++ }
            }
        }
        if($state.DuplicateKeys -or $state.InvalidValues) { $failures.Add("Invalid $($table.file): duplicate keys=$($state.DuplicateKeys), invalid values=$($state.InvalidValues)") }
    }
    $results.Add([pscustomobject]@{File=$table.file;Rows=if($SchemaOnly){$null}else{$state.Rows};Columns=$headers.Count;SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash})
}
$json=[pscustomobject]@{Passed=($failures.Count -eq 0);SchemaOnly=[bool]$SchemaOnly;Files=$results;Errors=@($failures)} | ConvertTo-Json -Depth 8
if($ReportPath) {[IO.File]::WriteAllText([IO.Path]::GetFullPath($ReportPath),$json,[Text.UTF8Encoding]::new($false))}
$json
if($failures.Count) {exit 1}
