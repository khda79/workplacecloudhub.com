# Exchange Free/Busy Checker - provenance

This folder vendors Microsoft's Free/Busy Checker script unchanged, so the Exchange hybrid
free/busy inventory runs a fixed, reviewed version.

| Item | Value |
|---|---|
| Upstream project | [microsoft/CSS-Exchange](https://github.com/microsoft/CSS-Exchange) |
| Documentation | https://microsoft.github.io/CSS-Exchange/Diagnostics/FreeBusyChecker/FreeBusyChecker/ |
| Release | [v26.10.19.1807](https://github.com/microsoft/CSS-Exchange/releases/tag/v26.10.19.1807) (published 2026-10-19, commit `5468126b85bffb8fd541ea29f2f4d84b4d611a36`) |
| Asset | https://github.com/microsoft/CSS-Exchange/releases/download/v26.10.19.1807/FreeBusyChecker.ps1 |
| Script version | 26.06.06.1238 |
| Size | 220759 bytes |
| SHA256 | `5E9BCF887AB3A06F29208EB0C3CAD87D0B38C30019F18BB8DE27791038EEEC34` (matches the release asset digest) |
| Authenticode | Valid, signed by CN=Microsoft Corporation, Microsoft timestamp |
| License | MIT, see [LICENSE](LICENSE) (upstream `LICENSE` at the same release) |
| Verified on | 2026-10-21 |

## Rules

- The file is stored byte for byte as published (UTF-8 with BOM, CRLF). `.gitattributes` marks
  `ThirdParty` folders `-text` so Git never converts line endings, which would break the Microsoft
  signature.
- The repository signing script and pre-push hook never sign files under a `ThirdParty` folder:
  they only check that the publisher signature is still valid and refuse the push otherwise.
- The collector checks the SHA256 above and the Microsoft signature before every run, on the copy
  it executes from a protected local work folder, and skips Free/Busy Checker when either check
  fails.
- It always runs the script with `-SkipVersionCheck` and no other parameter: any other parameter
  selects a parameter set that lets the script download a newer version and replace itself.
- The script's HTML report and transcript stay in that local work folder and are never published.

## Updating

1. Download `FreeBusyChecker.ps1` from a newer CSS-Exchange release.
2. Check `Get-AuthenticodeSignature` (Valid, Microsoft Corporation) and compare its SHA256 with the
   release asset digest.
3. Replace the file, update this table, and update the expected SHA256 in the collector in the same
   commit. Re-check the HTML report format: the collector parses it.
