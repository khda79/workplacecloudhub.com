# Exchange Server Health Checker - provenance

This folder vendors Microsoft's Exchange Server Health Checker script unchanged, so the
Exchange on-premises infrastructure inventory runs a fixed, reviewed version.

| Item | Value |
|---|---|
| Upstream project | [microsoft/CSS-Exchange](https://github.com/microsoft/CSS-Exchange) |
| Documentation | https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/ |
| Release | [v26.10.19.1807](https://github.com/microsoft/CSS-Exchange/releases/tag/v26.10.19.1807) (published 2026-10-19, commit `5468126b85bffb8fd541ea29f2f4d84b4d611a36`) |
| Asset | https://github.com/microsoft/CSS-Exchange/releases/download/v26.10.19.1807/HealthChecker.ps1 |
| Script version | 26.10.13.1644 |
| Size | 1048730 bytes |
| SHA256 | `BF8758E96E8760226E3094D45A84909AFFA9F91167DE11F739D8E3D9343229DC` (matches the release asset digest) |
| Authenticode | Valid, signed by CN=Microsoft Corporation (issuer Microsoft Code Signing PCA 2024), Microsoft timestamp |
| License | MIT, see [LICENSE](LICENSE) (upstream `LICENSE` at the same release) |
| Verified on | 2026-10-21 |

## Rules

- The file is stored byte for byte as published (UTF-8 with BOM, CRLF). `.gitattributes` marks
  `ThirdParty` folders `-text` so Git never converts line endings, which would break the Microsoft
  signature.
- The repository signing script and pre-push hook never sign files under a `ThirdParty` folder:
  they only check that the publisher signature is still valid and refuse the push otherwise.
- The inventory collector checks the SHA256 above and the Microsoft signature before every run and
  skips Health Checker when either check fails. It always runs it with `-SkipVersionCheck`, so the
  script never updates itself.

## Updating

1. Download `HealthChecker.ps1` from a newer CSS-Exchange release.
2. Check `Get-AuthenticodeSignature` (Valid, Microsoft Corporation) and compare its SHA256 with the
   release asset digest.
3. Replace the file, update this table, and update the expected SHA256 in the collector in the same
   commit.
