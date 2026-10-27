# Version 1.0.0
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1')

function Get-WorkforceDisabledLicenseEvidence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot,[Parameter(Mandatory)][string]$TenantKey)
    $result=[ordered]@{Count=$null;State='Not measured';Reason='License decision evidence is unavailable.'}
    try {
        $spec=(Get-Content (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json).licenseReport
        $last=Join-Path $DataRoot 'DATA-LAST'
        $entries=@(foreach ($name in @($spec.reportFiles)+@($spec.decisionFiles)) {
            $path=Join-Path $last $name
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                [pscustomobject]@{File='Source-'+$name;SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
            }
        })
        $qualified=Get-PreparedLicenseReportQualification -SnapshotRoot $DataRoot -TenantKey $TenantKey -Artifacts $entries
        $file='M365_Licenses_ReportSkuCandidates.csv'
        $state=@($qualified.Files | Where-Object File -ceq $file)
        if ($qualified.State -cne 'Qualified' -or $state.Count -ne 1 -or $state[0].State -cne 'Qualified') { return [pscustomobject]$result }
        $snapshot=Get-Content (Join-Path $last $spec.snapshot) -Raw | ConvertFrom-Json
        if ([string]$snapshot.TenantKey -cne $TenantKey) { throw 'License snapshot tenant mismatch.' }
        $path=Join-Path $last $file
        $columns=@('TenantKey','SnapshotId','ReceiptRunId','UserId','CandidateType','Recommendation')
        # A qualified empty file is measured zero only with the complete CSV schema.
        $header=@(Get-Content -LiteralPath $path -TotalCount 1)
        $schema=@(($header[0]+"`n"+(($columns | ForEach-Object {'""'}) -join ',')) | ConvertFrom-Csv)
        foreach ($name in $columns) { if (-not $schema[0].PSObject.Properties[$name]) { throw 'License decision schema mismatch.' } }
        $accounts=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in Import-Csv -LiteralPath $path) {
            if ([string]$row.TenantKey -cne $TenantKey -or [string]$row.SnapshotId -cne $qualified.SnapshotId -or [string]$row.ReceiptRunId -cne $qualified.ReceiptRunId) { throw 'License decision identity mismatch.' }
            if ($row.CandidateType -ceq 'Disabled account' -and $row.Recommendation -in @('Recommended','Review')) {
                if ([string]::IsNullOrWhiteSpace([string]$row.UserId)) { throw 'License candidate account key missing.' }
                $null=$accounts.Add($TenantKey+'|'+([string]$row.UserId).Trim())
            }
        }
        if ((Get-FileHash -LiteralPath $path).Hash -ine $state[0].SHA256) { throw 'License decisions changed during reading.' }
        $result.Count=$accounts.Count;$result.State='Observed';$result.Reason='Qualified License Report; distinct disabled-account candidates across paid SKUs, including review cases.'
    } catch {
        Write-Verbose $_.Exception.Message
        # Unqualified input is explicit and non-blocking; never infer zero or reuse older issues.
        $result.Count=$null;$result.State='Not measured';$result.Reason='License decision evidence failed qualification.'
    }
    [pscustomobject]$result
}
Export-ModuleMember -Function Get-WorkforceDisabledLicenseEvidence
