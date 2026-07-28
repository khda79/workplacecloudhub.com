# Contributing to WorkplaceCloudHub

Thank you for considering a contribution. WorkplaceCloudHub contains independent
PowerShell tools and applications for Microsoft 365, Intune, Entra ID, Exchange,
Azure, Windows endpoints, and modern workplace operations.

## Before You Start

1. Read the root `README.md` and the README in the project directory you intend
   to change.
2. Search existing issues to avoid duplicate work.
3. Open a feature request before starting a large change or a change that affects
   authentication, permissions, output schemas, tenant operations, or compatibility.
4. Never include credentials, certificates, tenant identifiers, customer data,
   production exports, logs, or other sensitive information.

## Preparing a Change

- Keep each contribution focused on one project and one logical purpose.
- Preserve the existing architecture, parameter conventions, safety gates, and
  read-only defaults of the affected project.
- Use placeholders in examples and documentation.
- Do not commit generated CSV files, reports, logs, support bundles, build output,
  local configuration, or test artifacts.
- Update the relevant README when behavior, prerequisites, permissions, parameters,
  or outputs change.

## Validation

Run the checks documented by the affected project. Depending on the project, this
may include:

- parsing under Windows PowerShell 5.1 and/or PowerShell 7;
- PSScriptAnalyzer;
- the project's local or synthetic tests;
- a bounded or preview-only execution that cannot overwrite production outputs;
- verification that no secrets or tenant-specific values are present.

Document the exact validation performed in the pull request. If a check could not
be run, explain why.

## Pull Requests

- Create a branch in your fork and open a focused pull request.
- Explain what changed, why it is needed, and which project paths are affected.
- Identify any authentication, permission, compatibility, schema, or operational
  impact.
- Include sanitized screenshots or output samples only when they materially help
  the review.
- Keep unrelated formatting or refactoring out of the pull request.

By submitting a contribution, you agree that it may be distributed under the
repository's `GPL-3.0-only` license.

## Security Issues

Do not open a public issue for a suspected vulnerability. Follow the instructions
in `.github/SECURITY.md`.

## Conduct

Participation in this project is governed by `CODE_OF_CONDUCT.md`.
