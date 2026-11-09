<#
.SYNOPSIS
Installs PowerShell and Python prerequisites for SmartM365 SmartInventory scripts.

.DESCRIPTION
Installs the PowerShell modules required by SmartM365 SmartInventory scripts from PSGallery,
then validates that cloud modules can be imported.

ActiveDirectory is a Windows RSAT / Windows Server feature, not a PSGallery module. This script
checks for it and prints the appropriate install hint when it is missing.

.PARAMETER Scope
PowerShellGet installation scope. Defaults to CurrentUser.

.PARAMETER Repository
PowerShell repository name. Defaults to PSGallery.

.PARAMETER TrustRepository
Sets the repository installation policy to Trusted before installing modules.

.PARAMETER Force
Passes -Force to Install-Module.

.PARAMETER AllowClobber
Passes -AllowClobber to Install-Module.

.PARAMETER SkipPublisherCheck
Passes -SkipPublisherCheck to Install-Module.

.PARAMETER SkipMicrosoftGraphMetaModule
Skips the Microsoft.Graph rollup module. Use only when Graph submodules are managed separately.

.PARAMETER WhatIf
Shows what would be installed without installing modules.

.PARAMETER SkipImportValidation
Skips Import-Module validation after installation.

.PARAMETER AuditOnly
Inspects module availability and Graph version alignment without installation,
repository access, imports, or changes to repository trust.

.NOTES
Version: 1.8
Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [string]$Repository = 'PSGallery',
    [switch]$TrustRepository,
    [switch]$Force = $true,
    [switch]$AllowClobber = $true,
    [switch]$SkipPublisherCheck,
    [switch]$SkipMicrosoftGraphMetaModule,
    [switch]$SkipImportValidation,
    [switch]$AuditOnly
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Stop'
function Get-SmartM365PythonPrerequisite {
    [CmdletBinding()]
    param([switch]$MachineOnly)
    $candidates=[Collections.Generic.List[string]]::new()
    if (-not $MachineOnly) {
        $command=Get-Command -Name 'python.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command -and $command.Source -notmatch '(?i)\\WindowsApps\\') { $candidates.Add([string]$command.Source) }
    }
    $pathValues=@([Environment]::GetEnvironmentVariable('Path','Machine'))
    if (-not $MachineOnly) { $pathValues+=@([Environment]::GetEnvironmentVariable('Path','User')) }
    foreach ($entry in @($pathValues | ForEach-Object { $_ -split ';' })) {
        if ([string]::IsNullOrWhiteSpace($entry) -or $entry -match '(?i)\\WindowsApps(?:\\|$)') { continue }
        $candidate=Join-Path ([Environment]::ExpandEnvironmentVariables($entry.Trim(' ', '"'))) 'python.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $candidates.Add($candidate) }
    }
    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        try {
            $info=[Diagnostics.ProcessStartInfo]::new($candidate,'--version')
            $info.UseShellExecute=$false; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true; $info.CreateNoWindow=$true
            $process=[Diagnostics.Process]::Start($info)
            try {
                $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
                if (-not $process.WaitForExit(15000)) { try {$process.Kill($true)} catch {}; continue }
                $versionText=([string]$stdout.GetAwaiter().GetResult())+' '+([string]$stderr.GetAwaiter().GetResult())
                if ($process.ExitCode -eq 0 -and $versionText -match 'Python\s+(\d+\.\d+\.\d+)' -and [version]$Matches[1] -ge [version]'3.10') {
                    return [pscustomobject]@{Available=$true;Version=$Matches[1];Path=$candidate}
                }
            } finally {$process.Dispose()}
        } catch { continue }
    }
    return [pscustomobject]@{Available=$false;Version='';Path=''}
}
function Invoke-SmartM365WindowsPowerShellPrerequisites {
    param([ValidateSet('Audit','Install')][string]$Mode,[ValidateSet('CurrentUser','AllUsers')][string]$ModuleScope,[bool]$Trust,[string]$RepositoryName='PSGallery')
    $legacy=Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $legacy -PathType Leaf)) { throw 'Windows PowerShell 5.1 is required for on-premises inventory.' }
    $child=@'
$ErrorActionPreference='Stop'
$scope='__SCOPE__'; $mode='__MODE__'; $trust=__TRUST__; $repository=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__REPOSITORY__'))
$roots=@((Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'),(Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules'))
if($scope -eq 'CurrentUser') {$roots=@((Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'))+$roots}
$env:PSModulePath=$roots -join [IO.Path]::PathSeparator
$missing=@()
foreach($name in @('Microsoft.Graph.Authentication','ImportExcel')) {
    $module=Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
    if(-not $module -and $mode -eq 'Install') {
        if(-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $scope -Force | Out-Null}
        if($trust) {Set-PSRepository -Name $repository -InstallationPolicy Trusted -ErrorAction Stop}
        $installOptions=@{Name=$name;Repository=$repository;Scope=$scope;Force=$true;AllowClobber=$true;ErrorAction='Stop'}
        if($name -eq 'Microsoft.Graph.Authentication') {$installOptions.RequiredVersion='2.41.2'}
        Install-Module @installOptions
        $module=Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
    }
    if(-not $module) {Write-Host ("MISSING WindowsPowerShell5 {0}" -f $name);$missing+=@($name);continue}
    if($mode -eq 'Install') {Import-Module -Name $module.Path -ErrorAction Stop}
    Write-Host ("AVAILABLE WindowsPowerShell5 {0} {1} {2}" -f $name,$module.Version,$module.Path)
}
if($missing.Count){exit 1}
exit 0
'@
    $repositoryEncoded=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($RepositoryName))
    $child=$child.Replace('__SCOPE__',$ModuleScope).Replace('__MODE__',$Mode).Replace('__TRUST__',$(if($Trust){'$true'}else{'$false'})).Replace('__REPOSITORY__',$repositoryEncoded)
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($child))
    & $legacy -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encoded
    if($LASTEXITCODE -ne 0) {throw 'Windows PowerShell 5.1 prerequisite check or installation failed.'}
}

$requiredModules = @()
if (-not $SkipMicrosoftGraphMetaModule) {
    $requiredModules += 'Microsoft.Graph'
}

$requiredModules += @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Sites',
    'Microsoft.Graph.Teams',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.DeviceManagement.Administration',
    'Microsoft.Graph.DeviceManagement.Enrollment',
    'Microsoft.Graph.Beta.Groups',
    'Microsoft.Graph.Beta.DeviceManagement',
    'Microsoft.Graph.Beta.Reports',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Reports',
    'MicrosoftTeams',
    'MicrosoftPowerBIMgmt.Profile',
    'ExchangeOnlineManagement',
    'ImportExcel',
    'PnP.PowerShell',
    'MSAL.PS'
)

$requiredModules = @($requiredModules | Select-Object -Unique)

if ($AuditOnly) {
    $missing = @()
    $graphVersions = @{}
    foreach ($name in $requiredModules) {
        $module = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if ($module) {
            Write-Host ("AVAILABLE {0} {1} {2}" -f $name, $module.Version, $module.Path)
            if ($name -like 'Microsoft.Graph.*') { $graphVersions[$name] = $module.Version }
        } else {
            Write-Host ("MISSING {0}" -f $name)
            $missing += $name
        }
    }
    $authVersion = $graphVersions['Microsoft.Graph.Authentication']
    $mismatched = @($graphVersions.Keys | Where-Object { $authVersion -and $graphVersions[$_] -ne $authVersion })
    foreach ($name in $mismatched) { Write-Host ("GRAPH VERSION MISMATCH {0} {1}; Authentication {2}" -f $name, $graphVersions[$name], $authVersion) }
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { Write-Host 'OPTIONAL ActiveDirectory RSAT missing; required for AD workloads.' }
    $python=Get-SmartM365PythonPrerequisite -MachineOnly:($Scope -eq 'AllUsers')
    if ($python.Available) { Write-Host ("AVAILABLE Python {0} {1}" -f $python.Version,$python.Path) }
    else { Write-Host 'MISSING Python 3.10+ (required for evidence preparation).'; $missing += 'Python' }
    try { Invoke-SmartM365WindowsPowerShellPrerequisites -Mode Audit -ModuleScope $Scope -Trust $false -RepositoryName $Repository } catch { Write-Host $_.Exception.Message; $missing += 'WindowsPowerShell5' }
    Write-Host 'Audit applies only to this PowerShell edition, account and PSModulePath; it does not test imports or service-account access.'
    if ($missing.Count -or $mismatched.Count) { exit 1 }
    exit 0
}

function Write-SmartM365SetupLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    Write-Host ("{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
}

function Ensure-PackageProvider {
    [CmdletBinding()]
    param()

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        if ($PSCmdlet.ShouldProcess('NuGet package provider', 'Install')) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $Scope -Force | Out-Null
        }
    }
}

function Ensure-RepositoryTrust {
    [CmdletBinding()]
    param([string]$Name)

    $repository = Get-PSRepository -Name $Name -ErrorAction Stop
    if ($TrustRepository -and $repository.InstallationPolicy -ne 'Trusted') {
        if ($PSCmdlet.ShouldProcess($Name, 'Set PSGallery installation policy to Trusted')) {
            Set-PSRepository -Name $Name -InstallationPolicy Trusted
        }
    }
}

function Install-RequiredModule {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($installed -and -not $Force) {
        Write-SmartM365SetupLog -Level SUCCESS -Message ("Module already installed: {0} {1}" -f $Name, $installed.Version)
        return
    }

    $installParams = @{
        Name       = $Name
        Scope      = $Scope
        Repository = $Repository
        ErrorAction = 'Stop'
    }
    if ($Force) { $installParams['Force'] = $true }
    if ($AllowClobber) { $installParams['AllowClobber'] = $true }
    if ($SkipPublisherCheck) { $installParams['SkipPublisherCheck'] = $true }

    if ($PSCmdlet.ShouldProcess($Name, "Install-Module -Scope $Scope")) {
        Write-SmartM365SetupLog -Message ("Installing module: {0}" -f $Name)
        Install-Module @installParams
    }
}

function Test-ActiveDirectoryPrerequisite {
    [CmdletBinding()]
    param()

    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Write-SmartM365SetupLog -Level SUCCESS -Message 'ActiveDirectory module is available.'
        return
    }

    Write-SmartM365SetupLog -Level WARN -Message 'ActiveDirectory module is missing. It is installed through RSAT or Windows Server features, not PSGallery.'
    Write-Host ''
    Write-Host 'Windows 10/11 client hint:'
    Write-Host '  Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
    Write-Host ''
    Write-Host 'Windows Server hint:'
    Write-Host '  Install-WindowsFeature RSAT-AD-PowerShell'
    Write-Host ''
}

Write-SmartM365SetupLog -Message 'Installing SmartM365 SmartInventory prerequisites.'
$python=Get-SmartM365PythonPrerequisite -MachineOnly:($Scope -eq 'AllUsers')
if (-not $python.Available) {
    if ($PSCmdlet.ShouldProcess('Python 3.14','Install using winget')) {
        $winget=Get-Command -Name 'winget.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $winget) { throw 'Python 3.10+ is missing and winget is unavailable. Install Python for all users and add it to the machine PATH, then rerun prerequisites.' }
        $pythonScope=if($Scope -eq 'AllUsers'){'machine'}else{'user'}
        $pythonOverride=if($Scope -eq 'AllUsers'){'/quiet InstallAllUsers=1 PrependPath=1 Include_test=0'}else{'/quiet InstallAllUsers=0 PrependPath=1 Include_test=0'}
        & $winget.Source install --id Python.Python.3.14 --exact --source winget --scope $pythonScope --override $pythonOverride --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) {throw "Python installation failed with exit code $LASTEXITCODE."}
        $python=Get-SmartM365PythonPrerequisite -MachineOnly:($Scope -eq 'AllUsers')
        if (-not $python.Available) {throw 'Python was installed but is not yet visible in PATH for the selected scope. Reopen the session and run -AuditOnly.'}
    }
} else { Write-SmartM365SetupLog -Level SUCCESS -Message ("Python {0} available: {1}" -f $python.Version,$python.Path) }
Ensure-PackageProvider
Ensure-RepositoryTrust -Name $Repository

foreach ($moduleName in $requiredModules) {
    Install-RequiredModule -Name $moduleName
}

if (-not $SkipImportValidation -and -not $WhatIfPreference) {
    foreach ($moduleName in $requiredModules) {
        Write-SmartM365SetupLog -Message ("Validating module import: {0}" -f $moduleName)
        Import-Module -Name $moduleName -ErrorAction Stop
        $loadedModule = Get-Module -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $loadedModule) {
            $loadedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        }
        if ($loadedModule) {
            Write-SmartM365SetupLog -Level SUCCESS -Message ("Module ready: {0} {1}; Path={2}" -f $loadedModule.Name, $loadedModule.Version, $loadedModule.Path)
        }
        else {
            throw ("Module import reported success but module version could not be resolved: {0}" -f $moduleName)
        }
    }
}

Test-ActiveDirectoryPrerequisite
if (-not $WhatIfPreference) { Invoke-SmartM365WindowsPowerShellPrerequisites -Mode Install -ModuleScope $Scope -Trust ([bool]$TrustRepository) -RepositoryName $Repository }
else { Write-SmartM365SetupLog -Message 'Preview: Windows PowerShell 5.1 Graph Authentication and ImportExcel prerequisites.' }

if ($WhatIfPreference) {
    Write-SmartM365SetupLog -Level INFO -Message 'Preview completed without installation or import validation. Use -AuditOnly to inspect installed modules.'
}
else {
    Write-SmartM365SetupLog -Level INFO -Message 'Module installation completed. Run -AuditOnly and the server acceptance checks under the actual service account; other host and workload prerequisites are separate.'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDf18ItBeFcNPSu
# Tv8gWVszIm8VqF/k/2cjX2QfGk0SCaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIECuMO/Y/bNx24irlNGqPLMblRElmDXIIBYB/a7hTWm/MA0GCSqG
# SIb3DQEBAQUABIIBgIaIuPnJ0VG1Bi9D3nZMEr9e2tqSXVoIXCDIsi9oLcpUGC0A
# rvckSZcxZp3IF/r5diYncE1jGLPAgWx3Ha4qMwlKM/txhTNMlA2+aq5miBfAvLDY
# xiynu/CmqFj+aIKd45idlx6r63P+ZUBuHCfQ0FTU2XE/rHEqGCecKZYhT/O0G51h
# s8Ds0BOcKRzAccai1MaJUYNJ840oYherPLYes6kfKKFVS2wFUWz7o1YXA2RqTuKK
# GeSFzNEXLg0jC2fSNHbkEsCRkCj3y/rEron+6AewBDdwJYt3ha2rjDbyd45k0sl3
# gvS0tn/xzuGYoLHiJq7QDwBx4xpHHlYoj7v+AvNDwr93mrEyMtEyRTmZx5kk0/aE
# 9wsWB2sUz7St52SmdpLB1xnMh2e4jYNqdbT/yjhyQgAZjAevPYWTiYTEalWGLUT6
# v3sm07f1ORp2KWZYz13VjhL8HgM+OtvZAeFwZZNw7uDyOiNKDVRfy012lnkwXKik
# 3Ysz3Pnr7P3T9Ed3wqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkxNDEx
# MjVaMC8GCSqGSIb3DQEJBDEiBCDGPun6Qqo6hGgfR6Y8PZh7QPm4EYR1gz70n97n
# 1LJzEDANBgkqhkiG9w0BAQEFAASCAgC0Ke6B6Ltyv7XbTgdYRa8kDsBti4B3gvin
# +/VTPoVSLnnIGsq/g49ghhMaPFe8p/KUjb8RmpaC12H3k1aU5AFCq+MaIL783wjU
# uMeNKxmg3F3cTsfs4bPadQ046C8Gz+bmWCosIC2iXZ216cOUSp6vbPhd+NodyM+U
# JFT0HyhZyX5HIoHs3k4XqdX1i0o+1TKGy9O+WdQfi30dCp/sj7ISEJQb2u0FBcxu
# y/Dah6jwxCKkMggMeuVPy0L0Imj/gDvZHOicgK0FnAuxeATNtEmSoAzsKpFudy+A
# bKrGkJyev+RmwLmg1S8leNvUQMX9yiIoc3PBQIC2miHMv/0wZTAazm7P4RyaCbSo
# JHF82aKg9yYjhCatl8ucjeJVFjfxQd2X7WhG+DUtxJsWHOxfbbgZUa9YmTmBjkd3
# wY8yxbaI9Ge6880hmax4WjuZtSY9tlrwZ4rxNt00PZJZ/1RfVrB9sojgfo//LO32
# yQkfsxkKoLSC/A0ffZ9ZxgYnPsSnLWntM2H1kPBhsHwVb1HIma8v2U1Ozzvdrjf0
# NnxmXOcAwwk3abElhLbN3A6T+b7TGyPZiYDGqFpeMSNHOl4HJFl/DG0mUzkoyTQz
# UZYmJT3AOu2JAEojonaALfocVg4813vvAvu3pGsSnjsHB0kTOfmZu6VsJL7nTfX2
# 7xDrZeOZsA==
# SIG # End signature block
