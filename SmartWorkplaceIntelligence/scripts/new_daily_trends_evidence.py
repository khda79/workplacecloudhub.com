"""Prepare domain and tenant daily trends from governed DATA-ALL daily exports.

This command reads only source CSVs and writes only a caller-selected staging folder.
It never promotes, refreshes, collects, or publishes evidence.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path


IDENTITY = ("TenantKey", "OrganizationKey", "EnvironmentKey", "TenantId")
COMMON = ("Date", "DomainName")
SOURCES = {
    "AD Users": {
        "file": "ActiveDirectory/Inventory/AD_Users_DailyStats.csv",
        "output": "DailyDirectoryTrends.csv",
        "metrics": ("TotalUsers", "EnabledUsers", "DisabledUsers", "Inactive90DaysUsers"),
        "core": ("TotalUsers", "EnabledUsers", "DisabledUsers"),
        "reference": "TotalUsers",
    },
    "AD Computers": {
        "file": "ActiveDirectory/Inventory/AD_Computers_DailyStats.csv",
        "output": "DailyDirectoryTrends.csv",
        "metrics": (
            "TotalComputers", "EnabledAccounts", "DisabledAccounts",
            "Windows 7 Enabled", "Windows 7 Disabled", "Windows 8 Enabled",
            "Windows 8 Disabled", "Windows 10 Enabled", "Windows 10 Disabled",
            "Windows 11 Enabled", "Windows 11 Disabled",
        ),
        "core": ("TotalComputers", "EnabledAccounts", "DisabledAccounts", "Windows 10 Enabled", "Windows 11 Enabled"),
        "reference": "TotalComputers",
    },
    "Exchange On-premises": {
        "file": "Exchange/OnPrem/MailboxReports/Exchange_OnPrem_Mailboxes_DailyStats.csv",
        "output": "DailyMailboxTrends.csv",
        "metrics": (
            "TotalMailboxCount", "TotalLocalMailboxCount", "TotalRemoteMailboxCount",
            "EnabledAccounts", "DisabledAccounts", "TotalLocalMailboxSizeGB",
            "EquipmentMailbox", "RemoteSharedMailbox", "RemoteUserMailbox",
            "RoomMailbox", "SharedMailbox", "UserMailbox",
        ),
        "core": ("TotalMailboxCount", "TotalLocalMailboxCount", "TotalRemoteMailboxCount", "EnabledAccounts", "DisabledAccounts"),
        "reference": "TotalMailboxCount",
    },
}
FIELDS = (
    *IDENTITY, "Date Key", "Date", "Source Name", "Domain Key", "Domain Name",
    "Metric Name", "Metric Unit", "Metric Value", "Reading At", "Metric Status",
    "Source Day Status", "Quality Event", "Quality Reason", "Observed Domain Count",
    "Absent Domain Count", "Persistent Absent Domain Count", "Absent Prior Share (%)",
    "Partial Threshold (%)", "Reference Complete Day",
)


def parse_date(value: str) -> datetime:
    value = value.strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(value.rstrip("Z"), fmt)
        except ValueError:
            pass
    raise ValueError(f"Invalid daily reading timestamp: {value!r}")


def domain_key(identity: tuple[str, ...], domain: str) -> str:
    payload = "\x1f".join((*identity, domain.casefold().strip())).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()[:24]


def numeric(value: str) -> Decimal | None:
    if not value.strip():
        return None
    try:
        return Decimal(value.strip().replace(" ", "").replace(",", "."))
    except InvalidOperation as exc:
        raise ValueError("Invalid numeric daily metric") from exc


def load_qualification(path: Path | None) -> dict:
    result = {"manualDays": [], "persistentDomains": []}
    if path is None:
        return result
    with path.open(encoding="utf-8-sig") as stream:
        data = json.load(stream)
    for key in result:
        if not isinstance(data.get(key, []), list):
            raise ValueError(f"Qualification {key} must be an array")
        result[key] = data.get(key, [])
    for item in result["manualDays"]:
        if item.get("source") not in SOURCES or not item.get("date") or not item.get("reason"):
            raise ValueError("A manual day requires source, date and reason")
    for item in result["persistentDomains"]:
        if item.get("source") not in SOURCES or not item.get("domain") or not item.get("effectiveDate") or not item.get("reason"):
            raise ValueError("A persistent domain override requires source, domain, effectiveDate and reason")
    return result


def read_source(path: Path, spec: dict):
    if not path.is_file():
        raise FileNotFoundError(path)
    selected = {}
    identities = set()
    with path.open(newline="", encoding="utf-8-sig") as stream:
        reader = csv.DictReader(stream, delimiter=";")
        needed = set((*IDENTITY, *COMMON, *spec["metrics"]))
        if not reader.fieldnames or needed - set(reader.fieldnames):
            raise ValueError(f"Daily source schema mismatch: {path.name}: {sorted(needed - set(reader.fieldnames or []))}")
        for row in reader:
            identity = tuple(row[field].strip() for field in IDENTITY)
            if not all(identity):
                raise ValueError(f"Empty tenant identity in {path.name}")
            identities.add(identity)
            when = parse_date(row["Date"])
            domain = row["DomainName"].strip()
            if not domain:
                raise ValueError(f"Empty domain in {path.name}")
            key = (when.date(), domain.casefold())
            if key not in selected or when > selected[key][0]:
                selected[key] = (when, row)
            elif when == selected[key][0] and row != selected[key][1]:
                raise ValueError(f"Conflicting readings with identical timestamp in {path.name}")
    if len(identities) != 1:
        raise ValueError(f"Exactly one tenant identity required in {path.name}")
    by_day = defaultdict(dict)
    for (day, normalized_domain), (when, row) in selected.items():
        by_day[day][normalized_domain] = (when, row)
    return next(iter(identities)), by_day


def generate_source(source: str, spec: dict, identity, by_day, threshold: Decimal, persistence: int, config: dict):
    manual = {item["date"]: item["reason"] for item in config["manualDays"] if item["source"] == source}
    overrides = {item["domain"].casefold().strip(): item for item in config["persistentDomains"] if item["source"] == source}
    rows = []
    reference = {}
    reference_day = ""
    streak = defaultdict(int)
    known_domains = set()
    retired = set()
    for day in sorted(by_day):
        readings = by_day[day]
        present = set(readings)
        known_domains.update(present)
        for domain in known_domains:
            if domain in present:
                streak[domain] = 0
                retired.discard(domain)
            else:
                streak[domain] += 1
        persistent_now = set()
        for domain in known_domains - present:
            override = overrides.get(domain)
            forced = override and day.isoformat() >= override["effectiveDate"]
            if domain not in retired and (streak[domain] >= persistence or forced):
                retired.add(domain)
                persistent_now.add(domain)
        absent = set(reference) - present - retired
        prior_total = sum(reference.values(), Decimal(0))
        absent_value = sum((reference[d] for d in absent), Decimal(0))
        share = (absent_value / prior_total * 100) if prior_total else Decimal(0)
        manual_reason = manual.get(day.isoformat(), "")
        partial = bool(manual_reason) or share > threshold
        status = "Partial capture" if partial else "Complete"
        event = "Domain absent (persistent)" if persistent_now else ("Manual qualification" if manual_reason else "")
        reason = manual_reason or ("Persistent domain absence reached threshold" if persistent_now else "")
        ref_date = reference_day
        domain_data = []
        for domain, (when, raw) in sorted(readings.items()):
            values = {metric: numeric(raw[metric]) for metric in spec["metrics"]}
            domain_data.append((domain, when, raw["DomainName"].strip(), values))
        for metric in spec["metrics"]:
            measured = all(item[3][metric] is not None for item in domain_data)
            total = sum((item[3][metric] for item in domain_data), Decimal(0)) if measured else None
            for domain, when, display_domain, values in domain_data:
                rows.append(make_row(identity, day, source, domain_key(identity, domain), display_domain, metric, values[metric], when, status, event, reason, len(present), len(absent), len(retired - present), share, threshold, ref_date))
            rows.append(make_row(identity, day, source, "__all__", "All domains", metric, total, max(item[1] for item in domain_data), status, event, reason, len(present), len(absent), len(retired - present), share, threshold, ref_date))
        if not partial and all(all(item[3][metric] is not None for item in domain_data) for metric in spec["core"]):
            reference = {item[0]: item[3][spec["reference"]] for item in domain_data}
            reference_day = day.isoformat()
    return rows


def make_row(identity, day, source, key, name, metric, value, when, status, event, reason, observed, absent, retired, share, threshold, ref_day):
    return dict(zip(FIELDS, (
        *identity, int(day.strftime("%Y%m%d")), day.isoformat(), source, key, name,
        metric, "GB" if metric.endswith("SizeGB") else "count", "" if value is None else str(value),
        when.isoformat(timespec="seconds"), "Measured" if value is not None else "Not measured",
        status, event, reason, observed, absent, retired, str(share.quantize(Decimal("0.0001"))),
        str(threshold), ref_day,
    )))


def run(data_root: Path, output_root: Path, threshold: Decimal, persistence: int, qualification: Path | None):
    if threshold < 0 or threshold > 100 or persistence < 1:
        raise ValueError("Invalid partial threshold or persistence count")
    if data_root.resolve() == output_root.resolve() or data_root.resolve() in output_root.resolve().parents:
        raise ValueError("Output must be outside the raw data root")
    config = load_qualification(qualification)
    tables = defaultdict(list)
    identities = set()
    for source, spec in SOURCES.items():
        identity, days = read_source(data_root / "DATA-ALL" / spec["file"], spec)
        identities.add(identity)
        tables[spec["output"]].extend(generate_source(source, spec, identity, days, threshold, persistence, config))
    if len(identities) != 1:
        raise ValueError("Daily sources must share exactly one tenant identity")
    output_root.mkdir(parents=True, exist_ok=True)
    for name, rows in tables.items():
        with (output_root / name).open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(rows)
        print(f"{name}: {len(rows)} rows")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--partial-threshold-percent", type=Decimal, default=Decimal("1"))
    parser.add_argument("--persistent-absence-observations", type=int, default=7)
    parser.add_argument("--qualification-path", type=Path)
    args = parser.parse_args()
    run(args.data_root, args.output_root, args.partial_threshold_percent, args.persistent_absence_observations, args.qualification_path)
