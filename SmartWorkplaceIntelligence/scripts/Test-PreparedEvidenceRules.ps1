[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
function Load-Function([string]$File,[string]$Name) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $File),[ref]$tokens,[ref]$errors)
    if($errors.Count) {throw ($errors.Message -join '; ')}
    $node=$ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$true) | Where-Object Name -eq $Name
    if(@($node).Count -ne 1) {throw "Function lookup failed: $Name"}
    [scriptblock]::Create($node.Extent.Text)
}
. (Load-Function 'New-LicensingEvidence.ps1' 'Test-OfficeDesktopExplicitNoUse')
. (Load-Function 'New-WorkforceIdentityEvidence.ps1' 'Get-OfficeDesktopUsageState')
$fields='Outlook (Windows)','Word (Windows)','Excel (Windows)','PowerPoint (Windows)','OneNote (Windows)'
$row=[ordered]@{};foreach($f in $fields){$row[$f]='False'};$row['Word (Mac)']='True'
if(-not(Test-OfficeDesktopExplicitNoUse ([pscustomobject]$row))) {throw 'Mac must not change the Windows-only criterion'}
if((Get-OfficeDesktopUsageState ([pscustomobject]$row)) -ne 'No PC app use in 30D') {throw '30D Windows-only no-use failed'}
$row['Word (Windows)']='True'
if(Test-OfficeDesktopExplicitNoUse ([pscustomobject]$row)) {throw 'Windows usage must exclude candidate'}
if((Get-OfficeDesktopUsageState ([pscustomobject]$row)) -ne 'Used on PC in 30D') {throw 'Observed Windows use failed'}
$row['Word (Windows)']=''
if(Test-OfficeDesktopExplicitNoUse ([pscustomobject]$row)) {throw 'Unknown must not mean no use'}
$row.Remove('Word (Windows)')
if(Test-OfficeDesktopExplicitNoUse ([pscustomobject]$row)) {throw 'Missing must not mean no use'}
if(Test-OfficeDesktopExplicitNoUse $null) {throw 'Missing report must exclude candidate'}
foreach($file in Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1') {
    $tokens=$null;$errors=$null
    $null=[Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
    if($errors.Count) {throw "Syntax error in $($file.Name): $($errors.Message -join '; ')"}
}
'PASS: Windows-only usage, explicit no-use, unknown/missing evidence, PowerShell syntax.'
