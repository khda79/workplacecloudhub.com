#Requires -Version 7.0
# Version: 1.0.0
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot,[switch]$SkipLoopbackShare)
# Synthetic files only: proves that the 1 MB-block SHA256 equals Get-FileHash, byte for byte, including
# the parallel hash map of the SharePoint transfer. No tenant, share or cloud connection.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
# After the transfer module: its own forced nested import would otherwise unload this global copy.
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$script:checks=0
function Check($condition,[string]$message){if(-not $condition){throw $message};$script:checks++}
function Reject([scriptblock]$action,[string]$message){$failed=$false;try{& $action | Out-Null}catch{$failed=$true};Check $failed $message}
try {
    # Sizes around the 4 KB and 1 MB block boundaries, plus multi-block files.
    $sizes=@(0,1,4095,4096,4097,65536,1048575,1048576,1048577,3145745,8388608)
    $random=[Random]::new(20261111)
    $paths=[Collections.Generic.List[string]]::new()
    foreach($size in $sizes){
        $bytes=[byte[]]::new($size);$random.NextBytes($bytes)
        $path=Join-Path $root ('file-{0}.bin' -f $size);[IO.File]::WriteAllBytes($path,$bytes);$paths.Add($path)
    }
    $csv=Join-Path $root 'synthetic.csv'
    $lines=foreach($i in 1..60000){'"TenantKey","Row{0}","{1}"' -f $i,('e' + [char]0x301 + ' ' + $i)}
    [IO.File]::WriteAllLines($csv,[string[]]$lines,[Text.UTF8Encoding]::new($true));$paths.Add($csv)
    $expected=@{}
    foreach($path in $paths){
        $expected[$path]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $actual=Get-PreparedFileSHA256 -LiteralPath $path
        Check ($actual -ceq $expected[$path]) "SHA256 differs from Get-FileHash: $(Split-Path $path -Leaf)"
        Check ($actual -cmatch '^[0-9A-F]{64}$') "SHA256 format differs: $(Split-Path $path -Leaf)"
    }
    # Same bytes as the published batch manifests: the empty-file hash is the SHA256 standard value.
    Check ((Get-PreparedFileSHA256 -LiteralPath $paths[0]) -ceq 'E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855') 'Empty file SHA256'
    # Relative paths resolve against the PowerShell location, like Get-FileHash, not the process directory.
    Push-Location -LiteralPath $root
    try {Check ((Get-PreparedFileSHA256 -LiteralPath '.\synthetic.csv') -ceq $expected[$csv]) 'Relative path resolution'} finally {Pop-Location}
    # Wildcard characters are literal.
    $bracket=Join-Path $root 'name[1].csv';Copy-Item -LiteralPath $csv -Destination $bracket
    Check ((Get-PreparedFileSHA256 -LiteralPath $bracket) -ceq $expected[$csv]) 'Literal path with brackets'
    Reject {Get-PreparedFileSHA256 -LiteralPath (Join-Path $root 'missing.csv')} 'Missing file must fail'
    Reject {Get-PreparedFileSHA256 -LiteralPath $root} 'Directory must fail'
    # A writer holding the file without read sharing blocks both methods: no silent read of a file being written.
    $writer=[IO.File]::Open($paths[5],[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try {
        Reject {Get-PreparedFileSHA256 -LiteralPath $paths[5]} 'Locked file must fail'
        Reject {Get-FileHash -LiteralPath $paths[5] -ErrorAction Stop} 'Reference: Get-FileHash fails on a locked file'
    } finally {$writer.Dispose()}
    # A changed byte changes the hash (the verification still detects any modification).
    $changed=Join-Path $root 'changed.bin';Copy-Item -LiteralPath $paths[8] -Destination $changed
    $stream=[IO.File]::Open($changed,'Open','ReadWrite','None');try{$stream.Position=1048576;$stream.WriteByte(([byte](($stream.ReadByte()+1)%256)))}finally{$stream.Dispose()}
    Check ((Get-PreparedFileSHA256 -LiteralPath $changed) -cne $expected[$paths[8]]) 'One changed byte past the first block must change the hash'
    Check ((Get-PreparedFileSHA256 -LiteralPath $changed) -ceq (Get-FileHash -LiteralPath $changed).Hash) 'Changed file equals Get-FileHash'
    # Parallel transfer hash map: same values, every path, caller order, read errors rethrown.
    $transfer=Get-Module PreparedSharePointTransfer
    $map=& $transfer {param($p) Get-PreparedFileHashMap $p} ([string[]]$paths)
    foreach($path in $paths){Check ($map[$path] -ceq $expected[$path]) "Hash map differs: $(Split-Path $path -Leaf)"}
    Check ($map.Count -eq $paths.Count) 'Hash map count'
    Reject {& $transfer {param($p) Get-PreparedFileHashMap $p} ([string[]]@($paths[1],(Join-Path $root 'missing.csv')))} 'Hash map must rethrow a read error'
    # Optional: the same comparison through the SMB client (loopback administrative share).
    if(-not $SkipLoopbackShare -and $root -match '^([A-Za-z]):\\(.*)$'){
        $unc='\\localhost\{0}${1}' -f $Matches[1],('\'+$Matches[2])
        if(Test-Path -LiteralPath $unc){
            foreach($path in $paths){
                $remote=Join-Path $unc (Split-Path $path -Leaf)
                Check ((Get-PreparedFileSHA256 -LiteralPath $remote) -ceq (Get-FileHash -LiteralPath $remote).Hash) "SMB SHA256 differs: $(Split-Path $path -Leaf)"
            }
            $uncMap=& $transfer {param($p) Get-PreparedFileHashMap $p} ([string[]]@($paths | ForEach-Object {Join-Path $unc (Split-Path $_ -Leaf)}))
            foreach($path in $paths){Check ($uncMap[(Join-Path $unc (Split-Path $path -Leaf))] -ceq $expected[$path]) "SMB hash map differs: $(Split-Path $path -Leaf)"}
        } else {Write-Host 'Loopback share not reachable: SMB comparison skipped.'}
    }
    Write-Host ('Test-PreparedFileSHA256: PASS ({0} checks)' -f $script:checks)
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
