[CmdletBinding()]
param([Parameter(Mandatory)][string]$PackageRoot,[string]$ExistingRepositoryRoot)
# Read-only verifier. No dot-sourcing of tenant profiles or remote authentication.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath $PackageRoot).Path
$manifest=Get-Content -LiteralPath (Join-Path $root 'package-manifest.json') -Raw | ConvertFrom-Json
if ($manifest.SchemaVersion -ne 1 -or @($manifest.Files).Count -eq 0) {throw 'Invalid manifest.'}
$seen=@{}
foreach($file in $manifest.Files) {
    $relative=[string]$file.Path
    if ([IO.Path]::IsPathRooted($relative) -or $relative -match '(^|[\\/])\.\.([\\/]|$)' -or $seen.ContainsKey($relative)) {throw 'Unsafe/duplicate manifest path.'}
    $seen[$relative]=$true
    $path=Join-Path $root $relative
    $item=Get-Item -LiteralPath $path
    if ($item.Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $file.SHA256) {throw "Package mismatch: $relative"}
    if ($item.Extension -in '.ps1','.psm1','.psd1') {
        $tokens=$null;$parseErrors=$null
        $null=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$parseErrors)
        if($parseErrors.Count) {throw "PowerShell syntax error: $relative"}
    }
}
foreach($item in Get-ChildItem -LiteralPath $root -File -Recurse) {
    $relative=[IO.Path]::GetRelativePath($root,$item.FullName).Replace('\','/')
    if($relative -ne 'package-manifest.json' -and -not $seen.ContainsKey($relative)) {throw "Unexpected package file: $relative"}
}
if($ExistingRepositoryRoot) {
    if ($PSVersionTable.PSVersion.Major -lt 7 -or -not [Environment]::Is64BitProcess) {throw 'PowerShell 7 x64 is required.'}
    foreach($dependency in $manifest.ExistingPrerequisites) {
        $path=Join-Path $ExistingRepositoryRoot $dependency.Path
        if(-not(Test-Path -LiteralPath $path -PathType Leaf)) {throw "Missing installed prerequisite: $($dependency.Path)"}
        if((Get-FileHash -LiteralPath $path).Hash -ne $dependency.ReferenceSHA256) {
            throw "Installed prerequisite differs from qualified reference; review before deployment: $($dependency.Path)"
        }
    }
    if(-not(Get-Module -ListAvailable ImportExcel)) {throw 'ImportExcel must be available to this PowerShell 7 account.'}
}
Write-Host "Verified $(@($manifest.Files).Count) package files. No installation or collection performed."
