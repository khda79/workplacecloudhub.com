# Migration Mailbox Size Reduction Recipients

Place optional manual recipient CSV files in this folder for `FromList` runs.

Start from `ExchangeMigrationMailboxSizeReduction-List.csv.template`, copy it to `ExchangeMigrationMailboxSizeReduction-List.csv` or another `.csv` file, then replace the example rows with the real recipients for the run.

Expected columns:

- `PrimarySmtpAddress`: recipient mailbox address.
- `UserName`: display name used in the email template.
- `LanguageTag`: optional language tag such as `en`, `fr`, or `fr-FR`.
- `TargetSkuPartNumber`: license SKU used to resolve the configured mailbox quota, such as `SPE_E3` or `SPE_F1`.
- `MailboxSizeMB`: optional mailbox size in MB when live Exchange usage checks are not used.
- `MailboxType`: optional mailbox type, usually `UserMailbox`.

Real recipient CSV files are local operational data and must not be committed.
