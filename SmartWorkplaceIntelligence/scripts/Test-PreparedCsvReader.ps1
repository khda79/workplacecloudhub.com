[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic fixtures only. The compiled strict reader must either return exactly what the
# reference parsers return (TextFieldParser tenant checks, Import-Csv column selection) or
# return null so that the reference parser is used.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
$module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
$script:checks = 0
function Check([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }; $script:checks++
}
function WriteFixture([string]$Name, [string]$Text, [bool]$Bom = $true) {
    $path = Join-Path $fixture $Name
    [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new($Bom))
    $path
}
# Reference tenant check: the TextFieldParser implementation used before the strict reader.
function ReferenceTenant([string]$Path, [string]$TenantKey, [string]$Delimiter, [bool]$AllowMissingTenantKey) {
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.SetDelimiters($Delimiter); $parser.HasFieldsEnclosedInQuotes = $true; $parser.TrimWhiteSpace = $false
        $header = $parser.ReadFields()
        if (-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count) { throw 'Empty or duplicate CSV header.' }
        $index = [array]::IndexOf($header,'TenantKey')
        if ($index -lt 0 -and -not $AllowMissingTenantKey) { throw 'TenantKey column missing.' }
        $rows = 0L
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields(); $rows++
            if ($fields.Length -ne $header.Length) { throw "Malformed CSV at record $rows." }
            if ($index -ge 0 -and ([string]::IsNullOrWhiteSpace($fields[$index]) -or $fields[$index] -ne $TenantKey)) { throw "Empty or incompatible TenantKey at record $rows." }
        }
        "Rows=$rows;HasTenantKey=$($index -ge 0)"
    } finally { $parser.Dispose() }
}
function Outcome([scriptblock]$Action) {
    try { & $Action } catch { 'Error: ' + $_.Exception.Message }
}
$tenant = 'synthetic-test'
$t = '"synthetic-test"'
$cases = @(
    @{Name='quoted'; Text="`"TenantKey`",`"Name`",`"Note`"`r`n$t,`"a, b`",`"say `"`"hi`"`"`"`r`n$t,`"line1`r`nline2`",`"`u{00E9}t`u{00E9}`"`r`n"; Strict=$true; Columns=@('TenantKey','Note')},
    @{Name='unquoted-lf'; Text="TenantKey,Name,Note`nsynthetic-test,alpha,`nsynthetic-test,,omega"; Strict=$true; Columns=@('Name','Note')},
    @{Name='mixed'; Text="`"TenantKey`",Name,`"Note`"`r`n$t,plain value,`"`"`r`n$t,`"`",x`r`n"; Strict=$true; Columns=@('Name')},
    @{Name='header-only'; Text="`"TenantKey`",`"Name`"`r`n"; Strict=$true; Columns=@('Name')},
    @{Name='no-bom'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Bom=$false; Strict=$true; Columns=@('name')},
    @{Name='blank-line'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n`r`n$t,`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='white-line'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n   `r`n$t,`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='lone-cr'; Text="`"TenantKey`",`"Name`"`r$t,`"x`"`r"; Strict=$false; Columns=@('Name')},
    @{Name='inner-quote'; Text="`"TenantKey`",`"Name`"`r`n$t,ab`"c`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='after-quote'; Text="`"TenantKey`",`"Name`"`r`n$t,`"ab`"c`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='space-before-quote'; Text="`"TenantKey`",`"Name`"`r`n$t, `"x`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='edge-space-value'; Text="`"TenantKey`",Name`r`n$t, padded `r`n"; Strict=$true; Columns=@('Name'); ColumnsStrict=$false},
    @{Name='tenant-edge-space'; Text="TenantKey,Name`r`n synthetic-test,x`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-case'; Text="`"TenantKey`",`"Name`"`r`n`"Synthetic-Test`",`"x`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-wrong'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n`"other`",`"y`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-empty'; Text="`"TenantKey`",`"Name`"`r`n`"`",`"x`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='duplicate-header'; Text="`"TenantKey`",`"Name`",`"name`"`r`n$t,`"x`",`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='missing-tenant'; Text="`"Name`",`"Note`"`r`n`"x`",`"y`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='short-row'; Text="`"TenantKey`",`"Name`"`r`n$t`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='unclosed'; Text="`"TenantKey`",`"Name`"`r`n$t,`"open"; Strict=$false; Columns=@('Name')},
    @{Name='type-line'; Text="#TYPE Synthetic`r`n`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='missing-column'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Strict=$true; Columns=@('Absent'); ColumnsStrict=$false},
    @{Name='multiline-selected'; Text="`"TenantKey`",`"Name`"`r`n$t,`"a`r`nb`"`r`n"; Strict=$true; Columns=@('Name'); ColumnsStrict=$false}
)
try {
    foreach ($case in $cases) {
        $bom = if ($case.ContainsKey('Bom')) { $case.Bom } else { $true }
        foreach ($delimiter in ',', ';') {
            $name = if ($delimiter -eq ';') { "$($case.Name)_DailyStats.csv" } else { "$($case.Name).csv" }
            $path = WriteFixture $name ($case.Text.Replace(',', $delimiter)) $bom
            $reference = Outcome { ReferenceTenant $path $tenant $delimiter $false }
            $fast = [Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]$delimiter, $tenant, $false)
            Check (($null -ne $fast) -eq $case.Strict) "Strict tenant path expectation failed: $name"
            if ($null -ne $fast) { Check ("Rows=$($fast.Rows);HasTenantKey=$($fast.HasTenantKey)" -ceq $reference) "Strict tenant result differs: $name" }
            $pipeline = Outcome { $r = & $module { param($p, $k) Test-PreparedTenant -Path $p -TenantKey $k } $path $tenant; "Rows=$($r.Rows);HasTenantKey=$($r.HasTenantKey)" }
            Check ($pipeline -ceq $reference) "Pipeline tenant outcome differs: $name ($pipeline / $reference)"
            $allowed = Outcome { ReferenceTenant $path $tenant $delimiter $true }
            $fastAllowed = [Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]$delimiter, $tenant, $true)
            if ($null -ne $fastAllowed) { Check ("Rows=$($fastAllowed.Rows);HasTenantKey=$($fastAllowed.HasTenantKey)" -ceq $allowed) "Strict legacy result differs: $name" }
        }
        $path = WriteFixture "$($case.Name)-columns.csv" $case.Text $bom
        $columnsStrict = if ($case.ContainsKey('ColumnsStrict')) { $case.ColumnsStrict } else { $case.Strict }
        $fastRows = [Workplace.Prepared.StrictCsv]::ReadColumns($path, [char]',', [string[]]$case.Columns)
        Check (($null -ne $fastRows) -eq $columnsStrict) "Strict column path expectation failed: $($case.Name)"
        if ($null -ne $fastRows) {
            $referenceRows = @(Import-Csv -LiteralPath $path | Select-Object -Property $case.Columns)
            Check ($fastRows.Length -eq $referenceRows.Count) "Column row count differs: $($case.Name)"
            for ($i = 0; $i -lt $referenceRows.Count; $i++) {
                foreach ($column in $case.Columns) {
                    Check ([string]$fastRows[$i].PSObject.Properties[$column].Value -ceq [string]$referenceRows[$i].PSObject.Properties[$column].Value) "Column value differs: $($case.Name) row $i $column"
                    Check ($null -ne $fastRows[$i].PSObject.Properties[$column].Value) "Column value is null: $($case.Name)"
                }
            }
        }
    }
    $before = (Get-FileHash -LiteralPath $path).Hash
    [void][Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]',', $tenant, $false)
    Check ((Get-FileHash -LiteralPath $path).Hash -ceq $before) 'Reader changed source bytes'
    Write-Host "PASS: $script:checks strict CSV reader checks against TextFieldParser and Import-Csv references. Synthetic data only."
} finally { Remove-Module $module; Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue }
