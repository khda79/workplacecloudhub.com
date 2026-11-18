<#
.SYNOPSIS
    Offline checks for User Principal Names listed more than once in the active user detail report.
.DESCRIPTION
    Runs the collector's actual normalization, ambiguity filter and export functions on a
    synthetic Graph CSV. Export-SmartM365Csv is mocked: no Graph, tenant collection, or
    publication is performed.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$collector = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-M365UserActivity-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($name in @('ConvertTo-DateOrNull','Get-SourcePropertyValue','ConvertTo-ReportBool','Get-LatestActivity','ConvertFrom-M365UserActivityReport','Remove-M365UserActivityAmbiguousRows','Get-M365UsageReportDefinition','Export-M365UsageReport')) {
    $definition = @($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true))
    if ($definition.Count -ne 1) { throw "Expected one $name function." }
    Invoke-Expression $definition[0].Extent.Text
}

$script:warnings = [System.Collections.Generic.List[string]]::new()
$script:exported = $null
function WriteLog { param([string]$Message,[string]$Level) if ($Level -eq 'WARNING') { $script:warnings.Add($Message) | Out-Null } }
function Export-SmartM365Csv {
    param($BaseFileName,$OutputPath,$GlobalPath,$Data,[switch]$NoWeeklyHistory,$Columns)
    $script:exported = @($Data)
    [pscustomobject]@{ LatestPath = 'latest'; PublishedPath = 'published'; TimestampedPath = 'timestamped' }
}
function Assert-Case { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }

$runId = '20261118_044507'
$Period = 'D180'
$runOutputRoot = 'unused'
$LatestCsvFolderPath = 'unused'
$header = 'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Has Exchange License,Exchange Last Activity Date,Assigned Products'
function New-RawLine { param([string]$Upn,[string]$Name,[string]$Licensed,[string]$Activity,[string]$Products)
    '2026-11-14,{0},{1},False,,{2},{3},{4}' -f $Upn,$Name,$Licensed,$Activity,$Products
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-UsageDuplicateUpn-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $definition = (Get-M365UsageReportDefinition)['Office365ActiveUserDetail']
    $rawPath = Join-Path $tempRoot 'raw.csv'

    # Unique UPNs: everything is published, no warning, no ambiguity reported.
    $lines = @($header) + @(1..150 | ForEach-Object { New-RawLine "user$_@contoso.example" "User $_" 'True' '2026-11-10' 'MICROSOFT 365 E3' })
    Set-Content -LiteralPath $rawPath -Value $lines -Encoding UTF8
    $script:warnings.Clear()
    $result = Export-M365UsageReport -ReportDefinition $definition -RawPath $rawPath
    Assert-Case ($result.Rows -eq 150 -and $script:exported.Count -eq 150) "Unique report lost rows: $($result.Rows)."
    Assert-Case (@($result.AmbiguousUserPrincipalNames).Count -eq 0 -and $script:warnings.Count -eq 0) 'Unique report reported ambiguity.'

    # A deleted licensed account and the new account reusing its UPN (case differs): both rows are left out.
    $lines = @($header) + @(1..300 | ForEach-Object { New-RawLine "user$_@contoso.example" "User $_" 'True' '2026-11-10' 'MICROSOFT 365 E3' }) +
        @((New-RawLine 'quality.site@contoso.example' 'Quality old account' 'True' '2026-11-10' 'MICROSOFT 365 E3'),
          (New-RawLine 'Quality.Site@contoso.example' 'Quality new account' 'False' '' ''))
    Set-Content -LiteralPath $rawPath -Value $lines -Encoding UTF8
    $script:warnings.Clear()
    $result = Export-M365UsageReport -ReportDefinition $definition -RawPath $rawPath
    Assert-Case ($result.Rows -eq 300 -and $script:exported.Count -eq 300) "Ambiguous rows were not left out: $($result.Rows)."
    Assert-Case (@($script:exported | Where-Object { $_.UserPrincipalName -match 'quality' }).Count -eq 0) 'An ambiguous UPN row was published.'
    Assert-Case ((@($result.AmbiguousUserPrincipalNames) -join ',') -eq 'quality.site@contoso.example') "Unexpected ambiguous list: $(@($result.AmbiguousUserPrincipalNames) -join ',')."
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match '1 User Principal Name\(s\) more than once; their 2 row\(s\)' -and $script:warnings[0] -match 'quality\.site@contoso\.example' -and $script:warnings[0] -match [regex]::Escape($rawPath)) "Ambiguity warning missing or incomplete: $($script:warnings -join ' | ')"
    $keys = @($script:exported | ForEach-Object { $_.UserPrincipalName.ToLowerInvariant() })
    Assert-Case ($keys.Count -eq @($keys | Sort-Object -Unique).Count) 'Published CSV still has a duplicate UPN.'

    # Above 1 % of the report the run fails before any publication.
    $lines = @($header) + @(1..10 | ForEach-Object { New-RawLine "user$_@contoso.example" "User $_" 'True' '2026-11-10' 'MICROSOFT 365 E3' }) +
        @((New-RawLine 'user1@contoso.example' 'Duplicate' 'False' '' ''))
    Set-Content -LiteralPath $rawPath -Value $lines -Encoding UTF8
    $script:exported = $null
    $caught = $false
    try { Export-M365UsageReport -ReportDefinition $definition -RawPath $rawPath | Out-Null } catch { $caught = $_.Exception.Message -match 'above the 1 % limit\. Nothing was published' }
    Assert-Case ($caught -and $null -eq $script:exported) 'Mass duplication was published or not rejected.'

    # Rows without a UPN are not treated as one ambiguous key.
    $filtered = Remove-M365UserActivityAmbiguousRows -Rows @([pscustomobject]@{UserPrincipalName=''},[pscustomobject]@{UserPrincipalName=' '},[pscustomobject]@{UserPrincipalName='a@contoso.example'})
    Assert-Case ($filtered.RemovedRows -eq 0 -and @($filtered.Rows).Count -eq 3) 'Blank UPNs were treated as duplicates.'

    'PASS: M365 user activity duplicate UPN checks'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
