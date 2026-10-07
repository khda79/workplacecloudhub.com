"""Flatten the SmartInventory mail snapshot without reclassifying licenses.

The JSON is the calculation authority. This module only validates its totals,
preserves unavailable values as blank, and links candidate users to CMDB IDs.
"""

import json
from pathlib import Path


SUMMARY_COLUMNS = (
    "SnapshotId", "TenantKey", "GeneratedAtUtc", "Product", "Category",
    "EnabledUnits", "ConsumedUnits", "AssignedUsers", "RecoveryCandidates",
    "RecoveryUnknown", "DisabledUsers", "NoM365Activity", "MultipleTargetSuites",
    "MultipleAssignedSkus", "NoAdOrEntraActivity", "NoMailboxActivity",
    "NoLocalAppsUse", "CandidatesPrimaryIntunePc", "SharedMailboxesLicensed",
    "SharedUnder50Gb", "SharedRemovalCandidates", "E3ToF3ReviewCandidates",
    "E3ToF3ReviewUnknown", "EvidenceStatus",
)
CANDIDATE_COLUMNS = (
    "SnapshotId", "TenantKey", "CandidateType", "Product", "UserId",
    "TenantUserKey", "IdentityJoinStatus", "UserPrincipalName", "DisplayName",
    "Reason", "LastAdActivityDate", "LastM365ActivityDate",
    "PrimaryOnIntuneWindowsPc", "TargetSuites", "MailboxSizeGB", "OneDriveUsedGB",
)
GAP_COLUMNS = ("SnapshotId", "TenantKey", "Section", "Metric", "Value", "EvidenceStatus")
GAP_METRICS = {
    "UserMailboxes": ("Total", "Universe", "OtherSkus", "NoSkus", "Unknown",
                      "EntraEnabled", "EntraDisabled", "EntraStateUnknown"),
    "NoUserMailbox": ("Total", "Universe", "OtherSkus", "NoSkus", "Unknown",
                      "Guests", "MemberEnabled", "MemberDisabled"),
    "AdGapActivity": ("Members", "AdObserved", "NoObservedMatch", "Ambiguous",
                      "AdEnabledRecent", "AdEnabledInactive", "AdEnabledNoDate",
                      "AdDisabled", "AdEnabledUnknown"),
}


def load_snapshot(path: Path, tenant_key: str) -> dict:
    if not path.is_file():
        raise ValueError("SmartInventory license report snapshot is missing: " + str(path))
    snapshot = json.loads(path.read_text(encoding="utf-8-sig"))
    if (snapshot.get("SchemaVersion") != 1 or snapshot.get("TenantKey") != tenant_key
            or len(snapshot.get("SnapshotId", "")) != 32
            or len(snapshot.get("Products", [])) != 4
            or len(snapshot.get("OtherProducts", [])) != 3):
        raise ValueError("SmartInventory license report snapshot has an invalid contract or tenant")
    if [row.get("Product") for row in snapshot["Products"]] != [
            "Microsoft 365 F1", "Microsoft 365 F3", "Microsoft 365 E3", "Microsoft 365 E5"]:
        raise ValueError("SmartInventory license report snapshot has unexpected target products")
    if len({row.get("Product") for row in snapshot["OtherProducts"]}) != 3:
        raise ValueError("SmartInventory license report snapshot has duplicate other products")
    recovery = snapshot.get("RecoveryCandidates", [])
    downgrade = snapshot.get("DowngradeCandidates", [])
    if not isinstance(recovery, list) or not isinstance(downgrade, list):
        raise ValueError("SmartInventory license report candidate lists are malformed")
    for product in snapshot["Products"]:
        counts = product.get("Counts")
        if product.get("UsageAvailable"):
            if not isinstance(counts, dict) or counts.get("RecoveryCandidates") != sum(
                    row.get("License") == product["Product"] for row in recovery):
                raise ValueError("Recovery details differ from the snapshot KPI")
        elif counts is not None:
            raise ValueError("Unqualified recovery metrics must be null")
    review = snapshot.get("E3ToF3Review")
    if review and review.get("Available") and review.get("Candidates") != len(downgrade):
        raise ValueError("E3 to F3 details differ from the snapshot KPI")
    return snapshot


def flatten(snapshot: dict, cmdb_users=()):
    def date_or_blank(value):
        # SmartInventory renders unavailable dates as N/D in the workbook.
        # Power BI needs a blank cell to type this column as a date.
        return "" if value in (None, "", "N/D") else value

    users_by_id = {}
    for user in cmdb_users:
        user_id = (user.get("SourceUserId") or "").strip().casefold()
        if user_id:
            if user_id in users_by_id:
                raise ValueError("CMDB user SourceUserId is duplicated")
            users_by_id[user_id] = user.get("TenantUserKey") or ""
    provisional = any(source.get("Provisional") for source in snapshot.get("Sources", []))
    summary = []
    for category, products in (("Suite", snapshot["Products"]),
                               ("Other paid product", snapshot["OtherProducts"])):
        for product in products:
            counts = product.get("Counts")
            value = lambda key: counts.get(key) if counts is not None else None
            review = snapshot.get("E3ToF3Review") if product["Product"] == "Microsoft 365 E3" else None
            summary.append({
                "SnapshotId": snapshot["SnapshotId"], "TenantKey": snapshot["TenantKey"],
                "GeneratedAtUtc": snapshot["GeneratedAtUtc"], "Product": product["Product"],
                "Category": category, "EnabledUnits": product["Enabled"],
                "ConsumedUnits": product["Consumed"], "AssignedUsers": value("Assigned"),
                "RecoveryCandidates": value("RecoveryCandidates"),
                "RecoveryUnknown": value("RecoveryUnknown"), "DisabledUsers": value("Disabled"),
                "NoM365Activity": value("M365Inactive"), "MultipleTargetSuites": value("Multiple"),
                "MultipleAssignedSkus": value("MultipleAll"),
                "NoAdOrEntraActivity": value("AdEntraInactive"),
                "NoMailboxActivity": value("MailboxInactive"),
                "NoLocalAppsUse": value("LocalAppsInactive"),
                "CandidatesPrimaryIntunePc": value("RecoveryPrimaryPc"),
                "SharedMailboxesLicensed": value("SharedLicensed"),
                "SharedUnder50Gb": value("SharedUnder50"),
                "SharedRemovalCandidates": value("SharedEligible"),
                "E3ToF3ReviewCandidates": review.get("Candidates") if review and review.get("Available") else "",
                "E3ToF3ReviewUnknown": review.get("Unknown") if review and review.get("Available") else "",
                "EvidenceStatus": "Capacity only" if category != "Suite" else
                    "N/D" if not product.get("UsageAvailable") else
                    "Provisional" if provisional else "Qualified",
            })
    candidates = []
    for candidate_type, details in (("Recovery", snapshot["RecoveryCandidates"]),
                                    ("E3 to F3 review", snapshot["DowngradeCandidates"])):
        for detail in details:
            user_id = (detail.get("UserId") or "").strip()
            cmdb_key = users_by_id.get(user_id.casefold(), "")
            candidates.append({
                "SnapshotId": snapshot["SnapshotId"], "TenantKey": snapshot["TenantKey"],
                "CandidateType": candidate_type,
                "Product": detail.get("License") if candidate_type == "Recovery" else "Microsoft 365 E3",
                "UserId": user_id, "TenantUserKey": cmdb_key,
                "IdentityJoinStatus": "Resolved" if cmdb_key else "Not observed in CMDB snapshot",
                "UserPrincipalName": detail.get("UserPrincipalName") or "",
                "DisplayName": detail.get("DisplayName") or "",
                "Reason": detail.get("RecoveryReason") if candidate_type == "Recovery" else "Manual downgrade review",
                "LastAdActivityDate": date_or_blank(detail.get("LastAdActivityDate")) if candidate_type == "Recovery" else "",
                "LastM365ActivityDate": date_or_blank(detail.get("LastM365ActivityDate")) if candidate_type == "Recovery" else "",
                "PrimaryOnIntuneWindowsPc": detail.get("PrimaryOnIntuneWindowsPc", "") if candidate_type == "Recovery" else "",
                "TargetSuites": detail.get("TargetSuites", "") if candidate_type == "Recovery" else "Microsoft 365 E3",
                "MailboxSizeGB": detail.get("MailboxSizeGB", "") if candidate_type != "Recovery" else "",
                "OneDriveUsedGB": detail.get("OneDriveUsedGB", "") if candidate_type != "Recovery" else "",
            })
    return summary, candidates


def flatten_gaps(snapshot: dict):
    result = []
    gaps = snapshot.get("MailboxGap") or {}
    for section, names in GAP_METRICS.items():
        record = snapshot.get("AdGapActivity") if section == "AdGapActivity" else gaps.get(section)
        available = bool(record and record.get("Available"))
        for name in names:
            value = record.get(name) if available else None
            result.append({"SnapshotId": snapshot["SnapshotId"], "TenantKey": snapshot["TenantKey"],
                           "Section": section, "Metric": name, "Value": value,
                           "EvidenceStatus": "N/D" if value is None else
                               "Provisional" if record.get("Provisional") else "Qualified"})
    return result
