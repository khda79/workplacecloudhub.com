# Local release safety

Generated **DemoAssets are local-only for every tool**. Never version them or upload them to GitHub, separately or inside application archives. Offline demo source/profile/launchers and normal runtime branding remain distributable.

`New-WorkplaceCloudHubReleaseZip.ps1` accepts an explicit file list, excludes DemoAssets directories and media archive names at every depth, and rejects embedded ZIP/NuGet archives containing DemoAssets. Preview is the default; `-Build` creates a new ZIP without replacing existing files.

`Publish-WorkplaceCloudHubRelease.ps1` checks every explicit asset, nested ZIP/NuGet content and release-note archive references before invoking GitHub CLI. Without `-Publish`, it only validates locally and does not run Git or GitHub CLI. An explicit repository, tag and full commit SHA are required. Existing releases are never clobbered by this wrapper.

Standalone product packagers embed the same two functions from `WorkplaceCloudHub.ReleaseSafety.ps1`, avoiding a new runtime dependency in distributed tools. Regression tests compare those function bodies and exercise contaminated inputs. Archive inspection fails closed on corruption, depth greater than eight, more than 100,000 entries or more than 512 MiB of inspected uncompressed data. No archive is extracted by the guard.

Build and upload through these guarded entrypoints; do not use wildcard uploads or attach whole local output directories. This code does not prevent a person from bypassing it with a direct `gh` command; the repository policy also applies to manual publication.
