# Read-only view of the SmartM365 configuration files for the Orchestrator GUI Configuration tab.
# Reads through JsonTransport naming (.json.txt or legacy .json) and waits for the transport lock;
# it never creates, renames, journals or rewrites a file. No WPF code here.
# Version: 1.0.0
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop

. (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ConfigView.Functions.ps1')

function Read-SmartM365OrchestratorConfigFile {
    <# Reads one configuration file while no writer holds its transport lock. The lock file is
       opened only when it already exists, exclusively and without writing, for the read duration. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [ValidateRange(0, 120)][int]$LockTimeoutSeconds = 10)
    $names = Get-SmartM365JsonNames $Path
    $selected = Get-SmartM365JsonReadPath $Path -Optional
    if (-not $selected) { return [pscustomobject]@{ Status = 'Missing'; Path = $names.Preferred; Document = $null; SHA256 = ''; Error = '' } }
    $lock = $null
    $deadline = [datetime]::UtcNow.AddSeconds($LockTimeoutSeconds)
    while (Test-Path -LiteralPath $names.Lock -PathType Leaf) {
        try { $lock = [IO.File]::Open($names.Lock, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None); break }
        catch [IO.FileNotFoundException] { break }
        catch [IO.IOException] {
            if ([datetime]::UtcNow -ge $deadline) {
                return [pscustomobject]@{ Status = 'Busy'; Path = $selected; Document = $null; SHA256 = ''; Error = "Transport lock held by another process: $($names.Lock)" }
            }
            Start-Sleep -Milliseconds 100
        }
    }
    try {
        $read = Read-SmartM365JsonDocument $Path -Validate { param($document) if ($document -isnot [System.Management.Automation.PSCustomObject]) { throw 'Configuration must be a JSON object.' } }
        return [pscustomobject]@{ Status = 'OK'; Path = $read.Path; Document = $read.Document; SHA256 = $read.SHA256; Error = '' }
    }
    catch { return [pscustomobject]@{ Status = 'Error'; Path = $selected; Document = $null; SHA256 = ''; Error = $_.Exception.Message } }
    finally { if ($null -ne $lock) { $lock.Dispose() } }
}

function Get-SmartM365OrchestratorConfigurationView {
    <# Reads the global, tenant and mail configuration of a profile with their templates, read-only. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootPath, [Parameter(Mandatory)][string]$ProfileKey,
        [string]$DescriptionsPath = '', [ValidateRange(0, 120)][int]$LockTimeoutSeconds = 10
    )
    $profileName = $ProfileKey.Trim().ToLowerInvariant()
    if ($profileName -notmatch '^[a-z0-9][a-z0-9._-]*$') { throw "Invalid tenant profile name: $ProfileKey" }
    $configFolder = Join-Path $RootPath 'Config'
    if (-not $DescriptionsPath) { $DescriptionsPath = Join-Path $configFolder 'SmartM365.ConfigDescriptions.json' }
    $descriptions = $null
    if (Test-Path -LiteralPath $DescriptionsPath -PathType Leaf) { $descriptions = Get-Content -LiteralPath $DescriptionsPath -Raw -Encoding UTF8 | ConvertFrom-Json }
    $definitions = @(
        @{ Kind = 'Global'; Title = 'SmartM365.global.local.json'; Path = (Join-Path $configFolder 'SmartM365.global.local.json'); Template = (Join-Path $configFolder 'SmartM365.global.local.json.template'); DescriptionKey = 'SmartM365.global.local.json'; SourceName = 'global' }
        @{ Kind = 'Tenant'; Title = "Tenants\$profileName.local.json"; Path = (Join-Path $configFolder "Tenants\$profileName.local.json"); Template = (Join-Path $configFolder 'Tenants\tenant.local.json.template'); DescriptionKey = 'Tenants/tenant.local.json'; SourceName = 'tenant' }
        @{ Kind = 'Mail'; Title = "Tenants\$profileName.mail.local.json"; Path = (Join-Path $configFolder "Tenants\$profileName.mail.local.json"); Template = (Join-Path $configFolder 'Tenants\tenant.mail.local.json.template'); DescriptionKey = 'Tenants/tenant.mail.local.json'; SourceName = 'mail' }
    )
    $files = foreach ($definition in $definitions) {
        $read = Read-SmartM365OrchestratorConfigFile -Path $definition.Path -LockTimeoutSeconds $LockTimeoutSeconds
        $template = $null; $templateError = ''
        if (Test-Path -LiteralPath $definition.Template -PathType Leaf) {
            try { $template = Get-Content -LiteralPath $definition.Template -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $templateError = $_.Exception.Message }
        }
        [pscustomobject]@{
            Kind = $definition.Kind; Title = $definition.Title; Path = $read.Path; Status = $read.Status
            Error = (@($read.Error, $(if ($templateError) { 'Template unreadable: ' + $templateError })) | Where-Object { $_ }) -join ' '
            Document = $read.Document; Template = $template; DescriptionKey = $definition.DescriptionKey; SourceName = $definition.SourceName
        }
    }
    $view = New-SmartM365OrchestratorConfigurationView -ProfileKey $profileName -Files @($files) -Descriptions $descriptions
    $view | Add-Member -NotePropertyName ConfigFolder -NotePropertyValue $configFolder
    $view | Add-Member -NotePropertyName ReadAt -NotePropertyValue (Get-Date)
    return $view
}

Export-ModuleMember -Function Read-SmartM365OrchestratorConfigFile, ConvertTo-SmartM365OrchestratorConfigEntries, Get-SmartM365OrchestratorConfigDescription,
    Test-SmartM365OrchestratorConfigSensitive, New-SmartM365OrchestratorConfigurationView, Get-SmartM365OrchestratorConfigurationView, Select-SmartM365OrchestratorConfigRows
