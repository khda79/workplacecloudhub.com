// Differential reference: unchanged partition from a9279196. Synthetic tests only.
let
    Key = (v as any) as text => if v = null then "" else Text.Lower(Text.Trim(Text.From(v))),
    Num = (v as any) as nullable number => if v = null or Text.Trim(Text.From(v)) = "" then null else try Number.FromText(Text.From(v), "en-US") otherwise null,
    Yes = (v as any) as logical => List.Contains({"true", "yes", "1"}, Key(v)),
    Field = (r as record, n as text) as any => Record.FieldOrDefault(r, n, null),
    Has = (v as any, terms as list) as logical => List.AnyTrue(List.Transform(terms, each Text.Contains(Key(v), _))),
    Prepared = ReadMigrationCsv(GetPreparedEvidenceFile("MailboxEvidence.csv"), {"TenantKey", "Primary SMTP Address", "Country", "Exchange Version"}),
    Tenants = List.Distinct(Prepared[TenantKey]),
    Tenant = if List.Count(Tenants) = 1 and Tenants{0} <> null and Text.Trim(Tenants{0}) <> "" then Tenants{0} else error "Prepared mailbox evidence must have one non-empty tenant",
    Read = (name as text, columns as list) as table =>
        let Source = ReadMigrationCsv(GetPreparedEvidenceFile("Source-" & name), {"TenantKey"} & columns)
        in if List.AllTrue(List.Transform(Source[TenantKey], each _ = Tenant)) then Table.Buffer(Source) else error Error.Record("Migration tenant", "Raw and prepared source tenant differ; keep previous model", name),
    GroupIndex = (source as table, column as text, unique as logical) as record =>
        let Keys = Table.AddColumn(source, "__key", each Key(Record.Field(_, column)), type text),
            Nonblank = Table.SelectRows(Keys, each [__key] <> ""),
            Groups = Table.Group(Nonblank, {"__key"}, {{"Rows", each Table.ToRecords(Table.RemoveColumns(_, {"__key"})), type list}}),
            Values = List.Transform(Groups[Rows], each if unique then (if List.Count(_) = 1 then _{0} else null) else _)
        in Record.FromList(Values, Groups[__key]),
    Mail = Read("Exchange_OnPrem_Mailboxes_AllDomains.csv", {"ExchangeGuid", "ObjectGUID", "RecipientType", "PrimarySMTPaddress", "DisplayName", "TotalItemSize-In-MB", "ArchiveTotalItemSize-In-MB", "FullAccessCount", "SendAsCount", "DomainName"}),
    Identity = Read("Exchange_HybridIdentity_Issues.csv", {"ObjectGUID", "Potential_Issue", "IssueCategory"}),
    Permissions = Read("Exchange_Mailboxes_AllSources_PermissionsByUser.csv", {"PrimarySMTPaddress", "HasCrossPremisesPermissions"}),
    Jobs = Read("Exchange_EXO_MigrationJobs.csv", {"EmailAddress", "MigrationUser", "BatchName", "BatchStatus", "UserStatus"}),
    JobKeys = Table.AddColumn(Jobs, "__user", each if Key([EmailAddress]) <> "" then [EmailAddress] else [MigrationUser], type text),
    PreparedIndex = GroupIndex(Prepared, "Primary SMTP Address", true),
    IdentityIndex = GroupIndex(Identity, "ObjectGUID", false),
    PermissionIndex = GroupIndex(Permissions, "PrimarySMTPaddress", true),
    JobIndex = GroupIndex(JobKeys, "__user", false),
    Local = Table.SelectRows(Mail, each not Text.StartsWith(Key([RecipientType]), "remote")),
    Qualified = Table.AddColumn(Local, "__id", each if Key([ExchangeGuid]) <> "" then Key([ExchangeGuid]) else Key([ObjectGUID]), type text),
    Ids = Qualified[__id],
    Checked = if not List.Contains(Ids, "") and List.Count(Ids) = List.Count(List.Distinct(Ids)) then Qualified else error "Missing or duplicate mailbox identity; previous inventory must be retained",
    Project = (r as record) as record =>
        let
            Smtp = Key(r[PrimarySMTPaddress]),
            P = Record.FieldOrDefault(PreparedIndex, Smtp, null),
            I = Record.FieldOrDefault(IdentityIndex, Key(r[ObjectGUID]), {}),
            Q = Record.FieldOrDefault(PermissionIndex, Smtp, null),
            J = Record.FieldOrDefault(JobIndex, Smtp, {}),
            Job = if List.Count(J) = 1 then J{0} else null,
            Size = Num(r[#"TotalItemSize-In-MB"]), Archive = Num(r[#"ArchiveTotalItemSize-In-MB"]),
            FA = Num(r[FullAccessCount]), SA = Num(r[SendAsCount]),
            Holds = Key(Field(r, "InPlaceHolds")),
            Hold = Yes(Field(r, "LitigationHoldEnabled")) or Yes(Field(r, "RetentionHoldEnabled")) or not List.Contains({"", "[]", "false", "none"}, Holds),
            Large = Num(Field(r, "LargeItemCount-Over-35MB")),
            Critical = List.AnyTrue(List.Transform(I, each Has(Field(_, "IssueCategory"), {"critical"}))),
            Failed = Job <> null and (Has(Field(Job, "UserStatus"), {"failed", "error"}) or Has(Field(Job, "BatchStatus"), {"failed", "error"})),
            Risk = List.RemoveNulls({if Critical then "Identity issue" else null, if Size = null then "Size unknown" else null, if Size <> null and Size / 1024 >= 45 then "Large mailbox - target quota review" else null, if Hold then "Hold or retention dependency" else null, if (FA <> null and FA <> 0) or (SA <> null and SA <> 0) or Key(Field(r, "GrantSendOnBehalfTo")) <> "" then "Delegation dependency" else null, if Large <> null and Large > 0 then "Large items - investigate" else null, if List.Count(J) > 1 then "Ambiguous batch association" else null, if Failed then "Migration failure" else null}),
            Blocker = Critical or Failed,
            Nonempty = (v as any, fallback as text) as text => if v = null or Text.Trim(Text.From(v)) = "" then fallback else Text.From(v),
            Issues = Text.Combine(List.Transform(I, each Nonempty(Field(_, "Potential_Issue"), "")), "; "),
            Collected = Nonempty(Field(r, "CollectedAtUtc"), "Unknown")
        in [#"Mailbox GUID"=Nonempty(r[ExchangeGuid], Text.From(r[ObjectGUID])), #"Object GUID"=r[ObjectGUID], #"Primary SMTP Address"=Smtp, #"Display Name"=r[DisplayName], Country=if P = null then "Unknown" else Nonempty(Field(P, "Country"), "Unknown"), #"Country Basis"="Prepared identity-country match; no inference from email suffix", #"Exchange Version"=if P = null then "Not observed" else Nonempty(Field(P, "Exchange Version"), "Not observed"), #"Mailbox Type"=r[RecipientType], #"Mailbox Size GB"=if Size = null then null else Size / 1024, #"Archive Size GB"=if Archive = null then null else Archive / 1024, #"Archive State"=Nonempty(Field(r, "ArchiveStatus"), Nonempty(Field(r, "ArchiveState"), "Unknown")), #"Full Access Count"=FA, #"Send As Count"=SA, #"Send On Behalf"=Nonempty(Field(r, "GrantSendOnBehalfTo"), ""), #"Cross-premises Permissions"=if Q = null then "Unknown" else if Yes(Field(Q, "HasCrossPremisesPermissions")) then "Observed" else "None observed", #"Hold and Retention"=if Hold then "Review - preserve settings" else "None observed", #"Large Item Count"=if Large = null then null else Text.From(Large, "en-US"), #"Large Item Evidence"=Nonempty(Field(r, "LargeItemCollectionStatus"), "Unknown"), #"Identity Issues"=if Issues = "" then "None in issue export - not identity readiness proof" else Issues, #"Batch Name"=if List.Count(J) > 1 then "Multiple - review" else if Job = null then "Not observed" else Nonempty(Field(Job, "BatchName"), "Not observed"), #"Migration User Status"=if Job = null then "Not observed" else Nonempty(Field(Job, "UserStatus"), "Not observed"), #"Risk Reasons"=if List.IsEmpty(Risk) then "Readiness evidence incomplete" else Text.Combine(Risk, "; "), #"Migration Status"=if Blocker then "Observed blocker" else if not List.IsEmpty(Risk) then "Risk review required" else "Readiness unknown", #"Next Action"=if Blocker then "Resolve identity or migration errors before scheduling" else "Validate quota, archives, permission migration and user readiness; schedule dependency-aware waves", #"Evidence Date"=if Collected = "Unknown" then Collected else Text.Start(Collected, 10), #"Source Domain"=r[DomainName]],
    Rows = Table.FromRecords(List.Transform(Table.ToRecords(Checked), Project), {"Mailbox GUID", "Object GUID", "Primary SMTP Address", "Display Name", "Country", "Country Basis", "Exchange Version", "Mailbox Type", "Mailbox Size GB", "Archive Size GB", "Archive State", "Full Access Count", "Send As Count", "Send On Behalf", "Cross-premises Permissions", "Hold and Retention", "Large Item Count", "Large Item Evidence", "Identity Issues", "Batch Name", "Migration User Status", "Risk Reasons", "Migration Status", "Next Action", "Evidence Date", "Source Domain"}, MissingField.Error),
    Typed = Table.TransformColumnTypes(Rows, {{"Mailbox Size GB", type number}, {"Archive Size GB", type number}, {"Full Access Count", type number}, {"Send As Count", type number}}, "en-US")
in
    Typed
