<#
.SYNOPSIS
    Verify conservative same-tenant attribution of ShareGate 401 traces.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGateReportReader.ps1')
$source='https://tenant.sharepoint.com/sites/source'
$target='https://tenant.sharepoint.com/sites/target'
function Check-Side([string]$failed,[string]$expected,[string]$src=$source,[string]$dst=$target){
 $actual=Get-SmartM365ShareGateFailingUrlSide -Text ("401 Unauthorized. WebUri : '{0}'" -f $failed) -SourceUrl $src -DestinationUrl $dst
 if($actual -ne $expected){throw "Expected $expected but got $actual for $failed"}
}
Check-Side ($source+'/_vti_bin/client.svc/ProcessQuery') 'Source'
Check-Side ($target+'/_api/web') 'Destination'
Check-Side 'https://tenant.sharepoint.com/sites/source2/_api' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com/_vti_bin/client.svc' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com/sites/source%2ftarget/_api' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com/sites/source/../target/_api' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com/sites/source/%2e%2e/target/_api' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com/sites/source/%252e%252e/target/_api' 'Undetermined'
Check-Side 'https://tenant.sharepoint.com:444/sites/source/_api' 'Undetermined'
Check-Side 'https://evil.example/sites/source/_api' 'Undetermined'
Check-Side ($source+'/_api') 'Undetermined' $source ($source+'/subweb')
$both="WebUri : '$source/_api' WebUri : '$target/_api'"
if((Get-SmartM365ShareGateFailingUrlSide -Text $both -SourceUrl $source -DestinationUrl $target) -ne 'Undetermined'){throw 'Conflicting failing URLs were attributed.'}
if((Get-SmartM365ShareGateFailingUrlSide -Text "WebUri : 'https://old.example/sites/a/_api'" -SourceUrl 'https://old.example/sites/a' -DestinationUrl 'https://new.example/sites/b') -ne 'Source'){throw 'Different-host source attribution changed.'}
$aliases=@{Errors=@('Errors');Warnings=@('Warnings');Messages=@('Messages');Details=@('Details');Status=@('Status');SourceItemId=@('Source ID')}
$rows=@([pscustomobject]@{Errors="401 Unauthorized. WebUri : '$source/_vti_bin/client.svc'";Warnings='';Messages='';Details='';Status='Error';'Source ID'='42'})
$assessment=Get-SmartM365ShareGatePrecheckAssessment -Rows $rows -SourceItemId 42 -SourceUrl $source -DestinationUrl $target -Aliases $aliases
if($assessment.Status -ne 'Source 401 observed'){throw 'Precheck did not use same-tenant attribution.'}
Write-Host 'PASS: same-tenant URL attribution and ambiguous cases; no site access.'
