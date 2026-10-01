"""Prepare only CMDB report-specific CSVs from a validated current collection.

The 20 unchanged PowerBI contract tables stay at DATA-LAST/PowerBI and are not
copied. This command never changes the collection or an existing report.
"""

import argparse
import csv
import datetime as dt
import json
import shutil
from pathlib import Path

import build_report as base
import report_360
import report_cockpit as cockpit
import report_hardware as hardware


ENRICHED = (
    "DimUser", "FactMailbox", "FactDataQuality", "FactUserDeviceRelationship",
    "FactUserLicense", "DimGroup", "DimDevice", "DimLicenseSku",
)
ROWLESS = ("FactDeviceApplication", "FactUserServicePlan")


def source_hash_path(source, relative):
    return (source.parent if Path(relative).parts[0] in {"DATA-LAST", "LOG-ALL"} else source) / relative


def hardware_hash_key(path, output):
    path = Path(path)
    try:
        return "@report/" + path.relative_to(output).as_posix()
    except ValueError:
        return str(path)


def hardware_hash_path(key, output):
    if key.startswith("@report/"):
        return output / key[len("@report/"):]
    return Path(key)


def report_sidecar_path(output, final_output=None):
    """Record the promoted directory, not its temporary staging name."""
    final = Path(final_output).resolve() if final_output is not None else output
    if final.parent != output.parent:
        raise ValueError("Final report directory must be beside the staging directory")
    return final


def write_rows(path, columns, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def validate_evidence(path, identity):
    if not path.is_file():
        raise ValueError("Required evidence not found: " + str(path))
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames or not any(
            name in reader.fieldnames for name in ("PrimarySmtpAddress", "PrimarySMTPaddress", "WindowsEmailAddress")
        ):
            raise ValueError("Mailbox evidence lacks an SMTP address: " + path.name)
        if not any(name in reader.fieldnames for name in ("RecipientTypeDetails", "RecipientType")):
            raise ValueError("Mailbox evidence lacks recipient type: " + path.name)
        count = 0
        for row in reader:
            if None in row or any(value is None for value in row.values()):
                raise ValueError("Malformed mailbox evidence: " + path.name)
            if any(row.get(key) != value for key, value in identity.items()):
                raise ValueError("Mailbox evidence tenant mismatch: " + path.name)
            count += 1
    return count


def top_application_rows(rows):
    groups = {}
    for row in rows:
        key = (
            (row.get("DisplayName") or "Unknown application").strip(),
            (row.get("Publisher") or "Unknown publisher").strip(),
            (row.get("Platform") or "Unknown platform").strip(),
        )
        item = groups.setdefault(key, {"versions": set(), "occurrences": 0})
        item["versions"].add((row.get("Version") or "Unknown version").strip())
        try:
            item["occurrences"] += int(float(row.get("DeviceCount") or 0))
        except ValueError as exc:
            raise ValueError("Invalid detected-application DeviceCount") from exc
    result = [
        {"ApplicationProduct": f"{name} · {publisher} · {platform}",
         "DisplayName": name, "Publisher": publisher, "Platform": platform,
         "VersionCount": len(value["versions"]), "ReportedDeviceCount": value["occurrences"]}
        for (name, publisher, platform), value in groups.items()
    ]
    result.sort(key=lambda item: (-item["ReportedDeviceCount"], item["ApplicationProduct"].casefold()))
    return result[:5]


def prepare(source, output, ci_hardware, local_mailboxes, remote_mailboxes,
            final_output=None):
    source, output = source.resolve(), output.resolve()
    final_output = report_sidecar_path(output, final_output)
    if output.exists() or source == output or output in source.parents:
        raise ValueError("Output must be new and separate from the collection source")
    if source in output.parents and (source / "PowerBI") not in output.parents:
        raise ValueError("Report data inside DATA-LAST must stay under PowerBI")
    if base.PRODUCT == output or base.PRODUCT in output.parents:
        raise ValueError("Private report data must not be written inside the repository")
    collection_manifest = source / "CMDB" / "CMDB_BuildManifest.csv"
    manifest_hash = base.sha(collection_manifest)
    _, manifest_rows = base.read_csv(collection_manifest)
    if len(manifest_rows) != 1 or int(manifest_rows[0]["PowerBITableCount"]) != 28:
        raise ValueError("Expected one complete 28-table CMDB build manifest")
    identity, data, columns, hashes = base.prepare_data(source, ROWLESS)
    if any(manifest_rows[0].get(key) != value for key, value in identity.items()):
        raise ValueError("CMDB build manifest tenant mismatch")
    report_360.enrich(source, identity, data, columns, hashes, base.read_csv, base.sha, base.PRODUCT)
    local_count = validate_evidence(local_mailboxes, identity)
    remote_count = validate_evidence(remote_mailboxes, identity)
    output.mkdir(parents=True)
    selected = (*ENRICHED, "SourceHealth", "DeviceSource", "LicenseAssignmentPath", "EntityFinding")
    counts = {}
    for name in selected:
        write_rows(output / (name + ".csv"), columns[name], data[name])
        counts[name] = len(data[name])

    copied_hardware = output / "CMDB_CIDeviceHardware.csv"
    copied_manifest = output / "CIRegistry.manifest.json.txt"
    shutil.copyfile(ci_hardware, copied_hardware)
    original_manifest = ci_hardware.parent / "CIRegistry.manifest.json.txt"
    if not original_manifest.is_file():
        original_manifest = ci_hardware.parent / "CIRegistry.manifest.json"
    shutil.copyfile(original_manifest, copied_manifest)
    if base.sha(ci_hardware) != base.sha(copied_hardware) or base.sha(original_manifest) != base.sha(copied_manifest):
        raise ValueError("CI hardware evidence changed while copying")
    _, hardware_rows, hardware_coverage, hardware_hashes = hardware.prepare_data(
        output / "DimDevice.csv", copied_hardware
    )
    hardware.write_report_data(output, hardware_rows, hardware_coverage)
    counts.update(DeviceHardware=len(hardware_rows), HardwareCoverage=1)

    _, local_rows = base.read_csv(local_mailboxes)
    _, remote_rows = base.read_csv(remote_mailboxes)
    for row in local_rows:
        row["PrimarySmtpAddress"] = row["PrimarySMTPaddress"]
    user_country_by_key = {
        row["TenantUserKey"]: row.get("CountryLabel") or "Unknown / unassigned"
        for row in data["DimUser"] if row.get("TenantUserKey")
    }
    user_country_by_address = {
        row["UserPrincipalName"].strip().lower(): row.get("CountryLabel") or "Unknown / unassigned"
        for row in data["DimUser"] if row.get("UserPrincipalName")
    }
    hosting_rows = cockpit.reconcile_mailboxes(
        data["FactMailbox"], remote_rows, local_rows, identity,
        user_country_by_key, user_country_by_address,
    )
    hosting_metadata = {
        "online": sum(row["HostingLocation"] == "Exchange Online" for row in hosting_rows),
        "onPremises": sum(row["HostingLocation"] == "Exchange On-premises" for row in hosting_rows),
        "total": len(hosting_rows),
        "localEvidenceDate": local_mailboxes.stat().st_mtime,
        "remoteEvidenceDate": remote_mailboxes.stat().st_mtime,
    }
    write_rows(output / "FactMailboxHosting.csv", cockpit.MAILBOX_HOSTING_COLUMNS, hosting_rows)
    counts["FactMailboxHosting"] = len(hosting_rows)
    countries = sorted({
        row.get("CountryLabel") or "Unknown / unassigned"
        for row in [*data["DimUser"], *data["DimDevice"], *hosting_rows]
    })
    write_rows(output / "DimCountry.csv", ["CountryLabel"], ({"CountryLabel": c} for c in countries))
    counts["DimCountry"] = len(countries)
    top = top_application_rows(data["DimDetectedApplication"])
    write_rows(output / "TopApplication.csv", cockpit.TOP_APPLICATION_COLUMNS, top)
    counts["TopApplication"] = len(top)
    bridge = cockpit.build_intune_device_bridge(data["DeviceSource"], identity)
    write_rows(output / "DimIntuneManagedDevice.csv", cockpit.INTUNE_DEVICE_BRIDGE_COLUMNS, bridge)
    counts["DimIntuneManagedDevice"] = len(bridge)
    relationships = []
    for name, file_name, evidence, predicate in (
        ("PrimaryUser", "FactUserDeviceRelationship.csv", "Exact curated user-device links", lambda row: True),
        ("HasMailbox", "FactMailbox.csv", "Exact curated user-mailbox links", lambda row: bool((row.get("CmdbUserId") or "").strip())),
        ("AssignedLicense", "FactUserLicense.csv", "Exact curated user-SKU assignments", lambda row: True),
        ("MemberOfGroup", "FactTeamMember.csv", "Exact Microsoft Teams membership enumeration", lambda row: True),
        ("DeviceHasApplication", "FactDeviceApplication.csv", "Exact Intune detected-app to managed-device links", lambda row: True),
        ("DeviceInAutopilot", "FactAutopilotDevice.csv", "Exact Intune managed-device to Autopilot links", lambda row: bool((row.get("ManagedDeviceId") or "").strip())),
    ):
        path = output / file_name
        if not path.is_file():
            path = source / "PowerBI" / file_name
        relationships.append({"RelationshipType": name,
                              "RelationshipCount": cockpit.count_csv_rows(path, predicate),
                              "EvidenceSource": evidence})
    write_rows(output / "FactRelationshipOverview.csv", cockpit.RELATIONSHIP_OVERVIEW_COLUMNS, relationships)
    counts["FactRelationshipOverview"] = len(relationships)

    # Recheck the controlling inputs after output generation; synchronized data
    # that changed during preparation cannot be promoted as one snapshot.
    if base.sha(collection_manifest) != manifest_hash:
        raise ValueError("CMDB build manifest changed during report preparation")
    for relative, digest in hashes.items():
        path = source_hash_path(source, relative)
        if not path.is_file() or base.sha(path) != digest:
            raise ValueError("Source changed during report preparation: " + relative)
    result = {
        "status": "Prepared", "generatedUtc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "sourceBuildDateTime": manifest_rows[0]["BuildDateTime"],
        "sourceRoot": str(source), "reportSidecar": str(final_output), "identity": identity,
        "rowCounts": counts, "sourceHashes": hashes,
        "mailboxEvidence": {
            "localPath": str(local_mailboxes), "localRows": local_count,
            "localSha256": base.sha(local_mailboxes),
            "remotePath": str(remote_mailboxes), "remoteRows": remote_count,
            "remoteSha256": base.sha(remote_mailboxes),
            "reconciled": hosting_metadata,
        },
        "hardwareInputHashes": {hardware_hash_key(path, output): digest
                                for path, digest in hardware_hashes.items()},
        "outputHashes": {name + ".csv": base.sha(output / (name + ".csv"))
                         for name in counts},
    }
    (output / "report-data.manifest.json.txt").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {"status": "Prepared", "sourceBuildDateTime": result["sourceBuildDateTime"],
            "tables": len(counts), "mailboxes": len(hosting_rows), "hardware": len(hardware_rows)}


def validate_current(source, output):
    """Fail closed if any collection, auxiliary evidence, or sidecar byte changed."""
    source, output = source.resolve(), output.resolve()
    manifest_path = output / "report-data.manifest.json.txt"
    if not manifest_path.is_file():
        raise ValueError("Prepared report-data manifest is missing")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("status") != "Prepared" or Path(manifest.get("sourceRoot", "")) != source:
        raise ValueError("Prepared report-data manifest has the wrong source")
    _, current_build = base.read_csv(source / "CMDB" / "CMDB_BuildManifest.csv")
    if len(current_build) != 1 or current_build[0]["BuildDateTime"] != manifest["sourceBuildDateTime"]:
        raise ValueError("Report-derived data does not match the latest CMDB build")
    if any(current_build[0].get(key) != value for key, value in manifest["identity"].items()):
        raise ValueError("Prepared report-data tenant identity changed")
    expected = manifest["outputHashes"]
    actual_names = {path.name for path in output.glob("*.csv")}
    evidence_name = "CMDB_CIDeviceHardware.csv"
    has_portable_evidence = "@report/" + evidence_name in manifest["hardwareInputHashes"]
    allowed_names = set(expected) | ({evidence_name} if has_portable_evidence else set())
    if actual_names != allowed_names:
        raise ValueError("Prepared report-data file set changed")
    for name, digest in expected.items():
        if base.sha(output / name) != digest:
            raise ValueError("Prepared report-data file changed: " + name)
    for relative, digest in manifest["sourceHashes"].items():
        path = source_hash_path(source, relative)
        if not path.is_file() or base.sha(path) != digest:
            raise ValueError("Collection source changed: " + relative)
    mailbox = manifest["mailboxEvidence"]
    for role in ("local", "remote"):
        path = Path(mailbox[role + "Path"])
        if not path.is_file() or base.sha(path) != mailbox[role + "Sha256"]:
            raise ValueError("Exchange " + role + " evidence changed")
    for name, digest in manifest["hardwareInputHashes"].items():
        path = hardware_hash_path(name, output)
        if not path.is_file() or base.sha(path) != digest:
            raise ValueError("Hardware evidence changed: " + name)
    return {"status": "Current", "sourceBuildDateTime": manifest["sourceBuildDateTime"],
            "sourceFiles": len(manifest["sourceHashes"]), "reportTables": len(expected)}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--ci-hardware", type=Path)
    parser.add_argument("--exchange-onprem-local", type=Path)
    parser.add_argument("--exchange-onprem-remote", type=Path)
    parser.add_argument("--final-output", type=Path,
                        help="Promoted report directory recorded in the manifest")
    parser.add_argument("--validate-only", action="store_true", help="Verify the prepared report matches every current source without writing.")
    args = parser.parse_args()
    if args.validate_only:
        result = validate_current(args.data_root, args.output)
    else:
        if not all((args.ci_hardware, args.exchange_onprem_local, args.exchange_onprem_remote)):
            parser.error("preparation requires --ci-hardware and both --exchange-onprem paths")
        result = prepare(args.data_root, args.output, args.ci_hardware,
                         args.exchange_onprem_local, args.exchange_onprem_remote,
                         args.final_output)
    print(json.dumps(result, ensure_ascii=False))
