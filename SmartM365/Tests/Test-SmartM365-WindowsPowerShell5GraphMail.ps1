<#
.SYNOPSIS
    Offline check of the Windows PowerShell 5 Graph mail body: recipients and attachments
    are always JSON arrays, also for a single value. Graph is mocked; nothing is sent.
.VERSION
1.0
#>

[CmdletBinding()]
param(
    [string]$ModulePath = ''
)

$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 leaves $PSScriptRoot empty in param() defaults.
if (-not $ModulePath) { $ModulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1' }
$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GraphMailTest-{0}' -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
# A synthetic global configuration: the module must not look for or create a tenant profile.
$global:SmartM365GlobalConfig = [pscustomobject]@{ ProfileKey = 'synthetic'; TenantKey = 'synthetic'; OrganizationKey = 'contoso'; MailTenantName = 'CONTOSO' }
$global:LogTextFile = Join-Path $workRoot 'test.log'
Import-Module $ModulePath -Force
$module = Get-Module SmartM365-WindowsPowerShell5

function Assert-GraphMail {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAILED: $Message" }
    Write-Output "PASS: $Message"
}

try {
    & $module {
        Set-Item -Path function:script:Connect-SmartM365GraphAppOnly -Value { param($AppId, $TenantId, $Thumbprint, $Purpose) $true }
    }
    Set-Item -Path function:global:Invoke-MgGraphRequest -Value { param($Method, $Uri, $Body, $ContentType) $global:SmartM365CapturedGraphBody = $Body }
    $attachment = Join-Path $workRoot 'report.txt'
    Set-Content -LiteralPath $attachment -Value 'synthetic attachment'

    $cases = @(
        @{ Name = 'one recipient'; To = 'one@contoso.com'; Cc = ''; Attachments = @(); ToCount = 1; CcCount = 0; AttachmentCount = 0 }
        @{ Name = 'two recipients and one Cc'; To = 'one@contoso.com;two@contoso.com'; Cc = 'cc@contoso.com'; Attachments = @(); ToCount = 2; CcCount = 1; AttachmentCount = 0 }
        @{ Name = 'one allowed attachment'; To = 'one@contoso.com'; Cc = ''; Attachments = @($attachment); ToCount = 1; CcCount = 0; AttachmentCount = 1 }
    )
    foreach ($case in $cases) {
        $global:SmartM365CapturedGraphBody = $null
        & $module {
            param($Case)
            Send-SmartM365GraphMail -From 'sender@contoso.com' -To $Case.To -Cc $Case.Cc -Subject 'Synthetic test' -BodyHtml '<p>Synthetic</p>' -Attachments $Case.Attachments -AllowAttachments -SkipHtmlCopy -AppId 'app' -TenantId 'tenant' -Thumbprint 'thumb'
        } $case
        $json = [string]$global:SmartM365CapturedGraphBody
        Assert-GraphMail ($json -match '"toRecipients":\s*\[') ("{0}: toRecipients is a JSON array" -f $case.Name)
        $document = $json | ConvertFrom-Json
        Assert-GraphMail (@($document.message.toRecipients).Count -eq $case.ToCount) ("{0}: {1} recipient(s)" -f $case.Name, $case.ToCount)
        if ($case.CcCount -gt 0) { Assert-GraphMail ($json -match '"ccRecipients":\s*\[') ("{0}: ccRecipients is a JSON array" -f $case.Name) }
        if ($case.AttachmentCount -gt 0) {
            Assert-GraphMail ($json -match '"attachments":\s*\[') ("{0}: attachments is a JSON array" -f $case.Name)
            Assert-GraphMail (@($document.message.attachments).Count -eq 1 -and $document.message.attachments[0].name -eq 'report.txt') ("{0}: attachment name preserved" -f $case.Name)
        }
    }
}
finally {
    Remove-Item -Path function:global:Invoke-MgGraphRequest -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}
