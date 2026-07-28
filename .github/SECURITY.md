# Security Policy

## Supported Version

Security fixes are applied to the latest version available on the default `main`
branch. Older copies, forks, and locally modified versions are not actively
maintained by WorkplaceCloudHub.

## Reporting a Vulnerability

Do not create a public GitHub issue for a suspected vulnerability.

Email [contact@workplacecloudhub.com](mailto:contact@workplacecloudhub.com) with
the subject `[SECURITY] WorkplaceCloudHub vulnerability`. Include:

- the affected project and file path;
- the version, release, or commit tested;
- a clear description of the potential impact;
- reproducible steps or a minimal proof of concept;
- any suggested remediation or mitigating control.

Remove credentials, tokens, certificates, tenant identifiers, customer data, and
other sensitive information before sending evidence. If sensitive evidence is
required, first request a suitable private exchange method.

The report will be reviewed privately. When the issue is confirmed, remediation
and disclosure will be coordinated with the reporter when practical. No response
or resolution time is guaranteed.

## Operational Safety

Many projects interact with enterprise platforms. A script that supports an
explicit write or repair operation must still be tested with preview, bounded,
or non-production inputs before use. A security report is not authorization to
run a production action.
