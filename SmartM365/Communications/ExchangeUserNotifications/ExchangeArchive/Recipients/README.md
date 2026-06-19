# Exchange Archive Recipients

Place campaign recipient CSV files in this folder.

Start from `ExchangeArchive-Recipients.csv.template`, copy it to a new `.csv` file, then replace the example rows with the real recipients for the run.

Expected columns:

- `PrimarySmtpAddress`: recipient mailbox address.
- `LanguageTag`: optional language tag such as `en`, `fr`, or `fr-FR`.
- `EffectiveDate`: optional archive activation date used by the template and skip logic. It can stay empty when `Force effective date` is selected in the GUI run options. When filled in CSV, use `yyyy-MM-dd` only, for example `2026-07-01`; avoid regional formats such as `01/07/2026`.
- `MailboxTotalGb`: optional mailbox size in GB when live Exchange usage checks are not used.
- `MailboxUsagePercent`: optional mailbox usage percentage when live Exchange usage checks are not used.

When `Force effective date` is set in the GUI, it overrides CSV `EffectiveDate` and `Date` values for every recipient.

Real recipient CSV files are local operational data and must not be committed.
