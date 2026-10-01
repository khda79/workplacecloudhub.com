<#
.SYNOPSIS
    Validates the shared site and persona classification module with synthetic workbooks and,
    when given, the TestCases worksheets of the private SmartWorkplaceIntelligence workbooks.
.VERSION
1.0
#>

[CmdletBinding()]
param(
    [string]$PersonaClassificationPath = '',
    [string]$SiteClassificationPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartInventoryRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory'
Import-Module (Join-Path $smartInventoryRoot 'Common\SmartM365.WorkplaceClassification.psd1') -MinimumVersion '1.0.0' -Force -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
$labels = Get-SmartM365WorkplacePersonaLabel
$names = Get-SmartM365WorkplaceClassificationWorkbookName

function Assert-Classification {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAILED: $Message" }
    Write-Output "PASS: $Message"
}

function Assert-Rejected {
    param([scriptblock]$Action, [string]$Message, [string]$Like = '*')
    $rejected = $false
    try { & $Action | Out-Null } catch { $rejected = $_.Exception.Message -like $Like }
    Assert-Classification $rejected $Message
}

function Export-SyntheticWorkbook {
    param([string]$Path, [hashtable]$Sheets)
    # Site codes stay text ('000001'), as in the governed workbooks.
    foreach ($sheet in $Sheets.Keys) { $Sheets[$sheet] | Export-Excel -Path $Path -WorksheetName $sheet -NoNumberConversion '*' }
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('WorkplaceClassification-{0}' -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    $sitePath = Join-Path $root $names.Site
    Export-SyntheticWorkbook -Path $sitePath -Sheets ([ordered]@{
        Settings = @([pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' }, [pscustomobject]@{ Parameter = 'Headquarters Persona'; Value = 'Headquarters staff' })
        SiteTypes = @('HQ','CLINIC','NURSING HOME','OTHER','UNKNOWN' | ForEach-Object { [pscustomobject]@{ 'Site Type Code' = $_ } })
        Sites = @(
            [pscustomobject]@{ 'Directory Site Code' = '000001'; 'Site Display Name' = 'Head office'; 'Site Type' = 'HQ'; 'Integration Status' = 'INTEGRATED'; 'Use for Classification' = 'Yes' }
            [pscustomobject]@{ 'Directory Site Code' = '000002'; 'Site Display Name' = 'Clinic A'; 'Site Type' = 'CLINIC'; 'Integration Status' = 'NOT-INTEGRATED'; 'Use for Classification' = 'Yes' }
            [pscustomobject]@{ 'Directory Site Code' = '000003'; 'Site Display Name' = 'Retired site'; 'Site Type' = 'OTHER'; 'Integration Status' = 'INTEGRATED'; 'Use for Classification' = 'No' }
        )
    })
    $personaPath = Join-Path $root $names.Persona
    Export-SyntheticWorkbook -Path $personaPath -Sheets ([ordered]@{
        Personas = @(
            [pscustomobject]@{ 'Persona ID' = 'HQ'; 'Persona name' = 'Headquarters staff'; Enabled = 'Yes' }
            [pscustomobject]@{ 'Persona ID' = 'ADMIN'; 'Persona name' = 'Management and Administrative staff'; Enabled = 'Yes' }
            [pscustomobject]@{ 'Persona ID' = 'CARE'; 'Persona name' = 'Healthcare staff'; Enabled = 'Yes' }
        )
        Rules = @(
            [pscustomobject]@{ 'Rule ID' = 'R1'; Keyword = 'accountant'; Country = 'ALL'; 'Persona ID' = 'ADMIN'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R2'; Keyword = 'nurse'; Country = 'ALL'; 'Persona ID' = 'CARE'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R3'; Keyword = 'manager'; Country = 'FR'; 'Persona ID' = 'ADMIN'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
            [pscustomobject]@{ 'Rule ID' = 'R4'; Keyword = 'infirmiere'; Country = 'FR'; 'Persona ID' = 'CARE'; 'Match type' = 'Contains phrase'; Priority = 100; Enabled = 'Yes' }
        )
        Exclusions = @([pscustomobject]@{ 'Rule ID' = 'X1'; Keyword = 'scan to mail'; Country = 'ALL'; Enabled = 'Yes' })
    })

    $site = Read-SmartM365WorkplaceSiteClassification -Path $sitePath
    $persona = Read-SmartM365WorkplacePersonaClassification -Path $personaPath -HeadquartersPersona $site.HeadquartersPersona
    $hq = Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode ' 000001 '
    Assert-Classification ($hq.Type -eq 'HQ' -and $hq.Name -eq 'Head office' -and $hq.IntegrationStatus -eq 'INTEGRATED') 'Mapped site code returns its type, name and integration status.'
    Assert-Classification ($null -eq (Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode '000003')) 'A site with Use for Classification = No is not mapped.'
    Assert-Classification ($null -eq (Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode '')) 'A blank site code is not mapped.'

    $cases = @(
        @{ Title = 'Accountant'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = 'Management and Administrative staff'; Name = 'Job title keyword' }
        @{ Title = 'Accountant'; Description = ''; Country = 'FR'; Site = 'HQ'; Expected = 'Headquarters staff'; Name = 'HQ site has priority over job title' }
        @{ Title = 'Scan to mail'; Description = ''; Country = 'FR'; Site = 'HQ'; Expected = $labels.NonHuman; Name = 'Non-human exclusion has priority over the HQ site' }
        @{ Title = 'Nurse manager'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.Ambiguous; Name = 'Equal-priority matches across personas stay ambiguous' }
        @{ Title = 'Nurse manager'; Description = ''; Country = 'DE'; Site = 'CLINIC'; Expected = 'Healthcare staff'; Name = 'Country-specific rule applies only to its country' }
        @{ Title = 'Infirmiere'; Description = ''; Country = 'France'; Site = 'CLINIC'; Expected = 'Healthcare staff'; Name = 'Country name is converted to its code' }
        @{ Title = 'Unmapped position'; Description = 'Site - Accountant'; Country = 'FR'; Site = 'CLINIC'; Expected = 'Management and Administrative staff'; Name = 'AD description is used when the title has no match' }
        @{ Title = 'Unmapped position'; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.NoMatch; Name = 'No match stays unclassified' }
        @{ Title = ''; Description = ''; Country = 'FR'; Site = 'CLINIC'; Expected = $labels.MissingTitle; Name = 'Missing title stays unclassified' }
    )
    foreach ($case in $cases) {
        $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle $case.Title -Description $case.Description -Country $case.Country -SiteType $case.Site
        Assert-Classification ($result.Persona -eq $case.Expected) ("{0} ({1})." -f $case.Name, $result.Persona)
    }

    $broken = Join-Path $root 'broken.xlsx'
    [pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' } | Export-Excel -Path $broken -WorksheetName 'Settings'
    $brokenSite = Join-Path -Path $root -ChildPath 'site-broken' -AdditionalChildPath $names.Site
    New-Item -ItemType Directory -Path (Split-Path $brokenSite) -Force | Out-Null
    Copy-Item -LiteralPath $broken -Destination $brokenSite
    Assert-Rejected -Action { Read-SmartM365WorkplaceSiteClassification -Path $brokenSite } -Message 'A workbook without the required worksheets is rejected.' -Like '*Required worksheet missing*'
    $duplicateSite = Join-Path -Path $root -ChildPath 'site-duplicate' -AdditionalChildPath $names.Site
    New-Item -ItemType Directory -Path (Split-Path $duplicateSite) -Force | Out-Null
    Export-SyntheticWorkbook -Path $duplicateSite -Sheets ([ordered]@{
        Settings = @([pscustomobject]@{ Parameter = 'Directory Site Attribute'; Value = 'extensionAttribute13' }, [pscustomobject]@{ Parameter = 'Headquarters Persona'; Value = 'Headquarters staff' })
        SiteTypes = @('HQ','UNKNOWN' | ForEach-Object { [pscustomobject]@{ 'Site Type Code' = $_ } })
        Sites = @('000001','000001' | ForEach-Object { [pscustomobject]@{ 'Directory Site Code' = $_; 'Site Display Name' = 'Duplicate'; 'Site Type' = 'HQ'; 'Use for Classification' = 'Yes' } })
    })
    Assert-Rejected -Action { Read-SmartM365WorkplaceSiteClassification -Path $duplicateSite } -Message 'A duplicate enabled site code is rejected.' -Like '*duplicate*'
    Assert-Rejected -Action { Read-SmartM365WorkplacePersonaClassification -Path $personaPath -HeadquartersPersona 'Disabled persona' } -Message 'An HQ persona missing from the enabled personas is rejected.' -Like '*Headquarters Persona*'
    $receiveFolder = Join-Path $root 'received'
    Assert-Rejected -Action { Receive-SmartM365WorkplaceClassificationWorkbook -Folder $receiveFolder -DownloadFile { param($destination, $name) Write-Verbose ("Simulated failed download: {0} -> {1}" -f $name, $destination) } } -Message 'A failed download stops without any cached fallback.' -Like '*No cached fallback*'
    $received = Receive-SmartM365WorkplaceClassificationWorkbook -Folder $receiveFolder -DownloadFile { param($destination, $name) Copy-Item -LiteralPath (Join-Path $root $name) -Destination $destination; Get-Item -LiteralPath $destination }.GetNewClosure()
    Assert-Classification ((Test-Path -LiteralPath $received.PersonaPath) -and (Test-Path -LiteralPath $received.SitePath)) 'Both downloaded workbooks are validated and returned.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

# Private workbooks: their own TestCases worksheets.
if ($PersonaClassificationPath -or $SiteClassificationPath) {
    if (-not $PersonaClassificationPath -or -not $SiteClassificationPath) { throw 'Give both PersonaClassificationPath and SiteClassificationPath.' }
    $site = Read-SmartM365WorkplaceSiteClassification -Path $SiteClassificationPath
    $persona = Read-SmartM365WorkplacePersonaClassification -Path $PersonaClassificationPath -HeadquartersPersona $site.HeadquartersPersona
    foreach ($case in @(Import-Excel -Path $PersonaClassificationPath -WorksheetName 'TestCases')) {
        $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle ([string]$case.'Job title') -Description '' -Country ([string]$case.Country) -SiteType 'UNKNOWN'
        Assert-Classification ($result.Persona -eq ([string]$case.'Expected result').Trim()) ("Persona TestCase '{0}' ({1})." -f $case.'Job title', $result.Persona)
    }
    foreach ($case in @(Import-Excel -Path $SiteClassificationPath -WorksheetName 'TestCases')) {
        $mapped = Get-SmartM365WorkplaceSite -SiteClassification $site -SiteCode ([string]$case.'Directory Site Code')
        $type = if ($null -ne $mapped) { $mapped.Type } else { 'UNKNOWN' }
        Assert-Classification ($type -eq ([string]$case.'Expected Site Type').Trim()) ("Site TestCase {0} ({1})." -f $case.'Directory Site Code', $type)
        $expectedPersona = ([string]$case.'Expected Persona').Trim()
        if ($expectedPersona) {
            $result = Get-SmartM365WorkplacePersona -PersonaClassification $persona -JobTitle 'Unmapped position' -Description '' -Country '' -SiteType $type
            Assert-Classification ($result.Persona -eq $expectedPersona) ("Site TestCase {0} persona ({1})." -f $case.'Directory Site Code', $result.Persona)
        }
    }
}
else { Write-Output 'Private classification workbooks not given; only the synthetic workbooks were validated.' }
