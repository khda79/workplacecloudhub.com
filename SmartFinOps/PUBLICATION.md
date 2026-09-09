# SmartFinOps Workplace Beta 1.5.0-beta.1 — publication handoff

Status: published on 2026-09-09 after explicit user validation. GitHub prerelease=true, latest=false; the 25-file site delta was published to OVH cluster129. Public package hashes and all 25 site files were read back and verified. Preserve the beta channel. No tenant or LinkedIn action was performed. The workflow below is retained for traceability and future releases; it is not an instruction to repeat this publication.

## Candidate assets

- `SmartFinOps-Workplace-1.5.0-beta.1.zip`
- `SmartFinOps-Workplace-1.5.0-beta.1.zip.sha256`
- `package-manifest.json` with `channel: beta`, `prerelease: true` and the packaged file hashes.
- Release notes and the source release manifest.

The package builder uses an explicit public-file allowlist, rejects unsigned/invalid PowerShell inputs, requires the beta prerelease manifest and refuses to overwrite an existing ZIP. It includes public certificate and license files. It excludes local JSON, exports, logs, reports, work folders, credentials and private keys.

```powershell
.\SmartFinOps-Workplace-BuildBetaPackage.ps1 -DestinationRoot 'C:\ReleaseCandidates\SmartFinOps-1.5.0-beta.1'
Get-FileHash 'C:\ReleaseCandidates\SmartFinOps-1.5.0-beta.1\SmartFinOps-Workplace-1.5.0-beta.1.zip' -Algorithm SHA256
```

## After final validation only

1. Recheck the exact source/package hashes, local and remote Git states, and concurrent work. Stage only the file allowlist in the local dossier; never use broad staging. Review hooks before push; avoid unrelated signature changes.
2. Commit the validated SmartFinOps changes, then push the approved commit. Create and push tag `smartfinops-workplace-v1.5.0-beta.1` pointing at that commit.
3. Create the release as a **prerelease**, not latest. For example, with the actual validated asset paths and release-notes body file:

```powershell
gh release create smartfinops-workplace-v1.5.0-beta.1 --verify-tag --prerelease --latest=false --title 'SmartFinOps Workplace 1.5.0-beta.1 — Beta' --notes-file RELEASE-NOTES.md SmartFinOps-Workplace-1.5.0-beta.1.zip SmartFinOps-Workplace-1.5.0-beta.1.zip.sha256 package-manifest.json
```

4. Read back GitHub release JSON and require `prerelease=true` and the correct tag/target. Download the public assets and verify hashes, manifest beta fields, signatures and ZIP contents.
5. Reconcile the two site-source patches with current `tools/build_site.py` and `content/manual-translations.json`. They were prepared from the existing local source including concurrent work; do not replace entire shared files. Check the patch baseline first.
6. Recheck the public hashes of all existing HTML/sitemap paths and the absence of the six new routes. If anything changed, rebuild the bounded SmartFinOps delta against that current public baseline. Publish only the 25 paths from the local site manifest to the configured OVH `cluster129` target after GitHub assets are available. Use Windows PowerShell 5.1 for the existing OVH publisher.
7. Read back all 25 public paths and compare bytes, then verify the six product languages, canonical/hreflang, beta markers and working GitHub asset links. Submit only the affected URLs to IndexNow if included in the approved publication. Submission receipt does not prove indexing.

The site download link now targets the published beta prerelease. For future releases, publish site links only after the approved prerelease exists. Do not publish an entire generated site directory: it contains unrelated concurrent work and existing translation debt.
