# Beta 1.5.0-beta.1 validation

Candidate prepared on 2026-09-09. Scope: local code, synthetic evidence, standalone packaging and multilingual site preparation. The initial preparation made no tenant connection/write, Git publication, OVH deployment or LinkedIn action. After explicit user validation, Git commit/push/tag, the GitHub prerelease and the 25-file OVH delta were published and publicly verified on 2026-09-09. No tenant or LinkedIn action was performed.

## Reproducible checks

Run both scripts under Windows PowerShell 5.1 and PowerShell 7 on Windows:

```powershell
.\Tests\SmartFinOps-Workplace-Beta.Tests.ps1
.\Tests\SmartFinOps-Workplace-Integration.Tests.ps1
```

The focused suite covers 48 assertions: comma/semicolon parsing, oldest refresh dates, malformed/future dates, invalid schemas, missing and empty CSVs, duplicate normalized identities, MAXITEMS exclusion, missing/invalid/negative numeric data, financial arithmetic, duplicate assignments, missing and stale evidence, directory conflicts, F3 storage boundaries and desktop activation, special accounts, separate capacity reuse, mailbox size bands, unknown/active holds, and no second value for mailbox conversion.

The seven integration scenarios run the complete analyzer from an isolated public-file copy: complete synthetic inputs, validation-only, duplicate assignments, duplicate tenant capacity, stale activity, nonexistent source directory, and validation-only with no source directory. Each financial scenario checks the CSV result and the beta HTML marker; the HTML must not expose the synthetic user identity. These checks passed on PowerShell 5.1 and 7. The complete case produces EUR 31.20 monthly potential; duplicate assignment retains EUR 31.20; stale activity and duplicate capacity suppress that potential. This is a fixture result, not customer savings.

Release preparation additionally checks PowerShell syntax, Authenticode signatures, JSON parsing, explicit package contents, per-file SHA-256 hashes, extracted ZIP behavior and Git whitespace. Detailed host results and the package checksum are retained in the local final-validation dossier, outside public runtime output folders.

## Site validation

Six product guides: EN/FR/IT/ES/DE/AR. Required checks: beta/version marker, translated product content, canonical URL, hreflang, RTL Arabic and GTM preservation. No missing translations remain in the new product guides. Existing unrelated translation debt on general site pages is excluded from this product change.

The site delta is based on downloaded public HTML: 18 existing SmartFinOps cards across home, tools and repository map; six new product guides; six new URLs added to the existing sitemap. Everything outside each existing SmartFinOps card is preserved byte-for-byte. Source modifications are delivered as two separate patches against the captured local site source baseline.

## Limits of the evidence

Offline checks do not establish tenant correctness, current contractual prices, product entitlements or production readiness. Large customer datasets, anonymized/renamed identities and simultaneous writes to one current output directory are not certified. A successful validation-only run does not mean the data is complete or fresh. Review [KNOWN-LIMITATIONS.md](KNOWN-LIMITATIONS.md) before final approval.
