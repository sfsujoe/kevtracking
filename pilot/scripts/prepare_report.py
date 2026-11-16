"""Read Defender CSV/XLSX exports and XLSX tracking; prepare JSON. Never writes XLSX."""
from __future__ import annotations

import argparse
import collections
import csv
import datetime as dt
import json
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from tracking_state import (ASSET_FIELDS, REGISTER_EXTRA, STATUSES, collect_saved_edits,
                            finalize_snapshot, normalize_manual)

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
TRACK_HEADERS = ["DeviceId", "DeviceName", "RemediationStatus", "NeedsEscalation",
                 "EscalationReason", "ServiceNowTickets", "Circumstances", "Notes",
                 "LastReviewed", "IdentityMatch", "LegacyColor", "AnnotationSource"]
TRACK_HEADERS += REGISTER_EXTRA
TRACK_HEADERS += ASSET_FIELDS
REQUIRED_EXPORT_FIELDS = ("DeviceId", "DeviceName", "KEVCount", "OverdueKEVs", "KEVs")
COUNT_FIELDS = ("KEVCount", "OverdueKEVs", "RansomwareLinkedKEVs",
                "OverdueRansomwareLinkedKEVs", "MissingDueDateKEVs",
                "AlertsLast30Days", "IncidentsLast6Months")
BOOLEAN_FIELDS = ("IsInternetFacing", "CveListTruncated")


def normal(value):
    return str(value or "").strip().lower()


def format_recent_users(value, os_platform=""):
    """Turn source user labels into the report's readable account list."""
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        value = value.strip()
        if not value or value.lower() == "null":
            return ""
        if value.startswith("["):
            try:
                value = json.loads(value)
            except ValueError as error:
                raise ValueError("RecentUsers contains an invalid JSON array") from error
        else:
            value = [value]
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError("RecentUsers must be a JSON array of user labels")
    mac = normal(os_platform).startswith(("mac", "osx"))
    labels, seen = [], set()
    pseudo_accounts = {"system", "local service", "network service", "anonymous logon"}
    for item in value:
        label = re.sub(r"\s+", " ", item).strip()
        if len(label) >= 2 and label[0] == label[-1] == '"':
            label = label[1:-1].strip()
        parts = re.split(r"\s+[\u2014\u2013-]\s+", label, maxsplit=1)
        username, display_name = parts[0].strip(), parts[1].strip() if len(parts) == 2 else ""
        # Accept an already formatted label without adding another pair of parentheses.
        if not display_name:
            match = re.fullmatch(r"(.+?)\s+\((.+)\)", username)
            if match:
                username, display_name = match.groups()
        username = username.rsplit("\\", 1)[-1].split("@", 1)[0].strip()
        key = username.casefold()
        if not username or key in pseudo_accounts or key.startswith(("dwm-", "umfd-")) or username.endswith("$"):
            continue
        formatted = f"{username} ({display_name})" if not mac and re.fullmatch(r"[0-9]{9}", username) and display_name else username
        if formatted not in seen:
            labels.append(formatted)
            seen.add(formatted)
    return ", ".join(labels)


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(x.itertext()) for x in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("s:si", NS)]
        styles = ET.fromstring(z.read("xl/styles.xml"))
        fills = []
        for fill in styles.find("s:fills", NS):
            fg = fill.find("s:patternFill/s:fgColor", NS)
            fills.append(fg.get("rgb", "") if fg is not None else "")
        colors = [fills[int(x.get("fillId", 0))] for x in styles.find("s:cellXfs", NS)]
        rels = {x.get("Id"): x.get("Target") for x in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        sheets = {}
        for sheet in ET.fromstring(z.read("xl/workbook.xml")).find("s:sheets", NS):
            target = rels[sheet.get(REL)]
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            rows = []
            for row in ET.fromstring(z.read(target)).findall("s:sheetData/s:row", NS):
                values, cell_colors, formulas = {}, {}, {}
                for cell in row:
                    address = cell.get("r")
                    if not address:
                        continue
                    col = re.sub(r"\d", "", address)
                    val = cell.find("s:v", NS)
                    value = (val.text or "") if val is not None else ""
                    kind = cell.get("t")
                    if kind == "s":
                        value = strings[int(value)]
                    elif kind == "inlineStr":
                        value = "".join(cell.find("s:is", NS).itertext())
                    elif kind == "b":
                        value = value == "1"
                    elif kind == "d" and value:
                        stamp = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
                        if stamp.tzinfo:
                            stamp = stamp.astimezone(dt.timezone.utc).replace(tzinfo=None)
                        serial = (stamp - dt.datetime(1899, 12, 30)).total_seconds() / 86400
                        value = int(serial) if serial.is_integer() else serial
                    elif kind not in ("str", "e") and value != "":
                        number = float(value)
                        value = int(number) if number.is_integer() else number
                    values[col] = value
                    cell_colors[col] = colors[int(cell.get("s", 0))]
                    formula = cell.find("s:f", NS)
                    if formula is not None:
                        formulas[col] = formula.text or ""
                rows.append({"row": int(row.get("r")), "values": values, "colors": cell_colors, "formulas": formulas})
            sheets[sheet.get("name")] = rows
        return sheets


def table_records(rows, first_header):
    for i, row in enumerate(rows):
        if row["values"].get("A") == first_header:
            headers = row["values"]
            return [{str(header): r["values"].get(col, "") for col, header in headers.items() if header != ""}
                    for r in rows[i + 1:] if any(v != "" for v in r["values"].values())]
    raise ValueError(f"Table with first header {first_header!r} not found")


def read_csv_export(path):
    # utf-8-sig accepts both ordinary UTF-8 and the optional Excel/MDE BOM.
    # Also support BOM-marked UTF-16 without guessing an ANSI code page.
    with Path(path).open("rb") as stream:
        prefix = stream.read(4)
    encoding = "utf-16" if prefix[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
    csv.field_size_limit(16 * 1024 * 1024)  # Long CVE/alert arrays can exceed the default 128 KiB.
    records = []
    try:
        with Path(path).open("r", encoding=encoding, newline="") as stream:
            reader = csv.reader(stream, strict=True)
            headers = next((row for row in reader if any(cell.strip() for cell in row)), [])
            headers = [header.strip() for header in headers]
            if not headers:
                raise ValueError("The Defender CSV has no header row")
            if any(not header for header in headers) or len(headers) != len(set(headers)):
                raise ValueError("The Defender CSV contains blank or duplicate column headers")
            missing = [field for field in REQUIRED_EXPORT_FIELDS if field not in headers]
            if missing:
                raise ValueError("Missing required Defender CSV columns: " + ", ".join(missing))
            for values in reader:
                if not any(value.strip() for value in values):
                    continue
                if len(values) != len(headers):
                    raise ValueError(f"Defender CSV line {reader.line_num} has {len(values)} fields; expected {len(headers)}. Check CSV quoting.")
                row = dict(zip(headers, values))
                # Convert only known measures/booleans. Device/user identifiers
                # stay strings, including their leading zeroes; JSON stays intact.
                for field in COUNT_FIELDS:
                    if field not in row:
                        continue
                    value = row[field].strip()
                    if value == "" and field not in ("KEVCount", "OverdueKEVs"):
                        row[field] = ""
                        continue
                    if not re.fullmatch(r"[0-9]+", value):
                        raise ValueError(f"{field} must be a nonnegative integer (CSV line {reader.line_num})")
                    row[field] = int(value)
                for field in BOOLEAN_FIELDS:
                    if field not in row:
                        continue
                    value = row[field].strip().lower()
                    if value in ("", "null"):
                        row[field] = ""
                    elif value in ("true", "1"):
                        row[field] = True
                    elif value in ("false", "0"):
                        row[field] = False
                    else:
                        raise ValueError(f"{field} must be true, false, 1, 0, or blank (CSV line {reader.line_num})")
                records.append(row)
    except UnicodeError as error:
        raise ValueError("Could not decode the Defender CSV. Use UTF-8 or BOM-marked UTF-16.") from error
    except csv.Error as error:
        raise ValueError(f"Invalid Defender CSV: {error}") from error
    return records


def read_export(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        records = read_csv_export(path)
    elif suffix == ".xlsx":
        for rows in read_xlsx(path).values():
            if rows and rows[0]["values"].get("A") == "DeviceId":
                records = table_records(rows, "DeviceId")
                break
        else:
            raise ValueError("Expected Defender export with DeviceId as the first header")
    else:
        raise ValueError("The Defender export must be a .csv or .xlsx file")
    if not records:
        raise ValueError("The export contains no devices; refusing to generate an apparently healthy empty report")
    ids = set()
    for row in records:
        for key in REQUIRED_EXPORT_FIELDS:
            if key not in row or row[key] is None or (isinstance(row[key], str) and not row[key].strip()):
                raise ValueError(f"Missing required source field: {key}")
        row["DeviceId"] = normal(row["DeviceId"])
        if row["DeviceId"] in ids:
            raise ValueError(f"Duplicate DeviceId in source: {row['DeviceId']}")
        ids.add(row["DeviceId"])
        for key in ("KEVCount", "OverdueKEVs"):
            value = row[key]
            if isinstance(value, bool) or not isinstance(value, (float, int)) or value < 0 or int(value) != value:
                raise ValueError(f"{key} must be a nonnegative integer")
            row[key] = int(value)
        if row["OverdueKEVs"] > row["KEVCount"]:
            raise ValueError("OverdueKEVs exceeds KEVCount")
        try:
            cves = json.loads(row["KEVs"])
        except (ValueError, TypeError) as error:
            raise ValueError(f"Invalid JSON in KEVs for device {row['DeviceId']}") from error
        if not isinstance(cves, list) or not all(isinstance(cve, str) and cve for cve in cves) or len(cves) != len(set(cves)) or len(cves) > row["KEVCount"]:
            raise ValueError("Invalid or duplicated CVEs in KEVs array")
        if row.get("OverdueCves"):
            try:
                overdue = json.loads(row["OverdueCves"])
            except (ValueError, TypeError) as error:
                raise ValueError(f"Invalid JSON in OverdueCves for device {row['DeviceId']}") from error
            if not isinstance(overdue, list) or not all(isinstance(cve, str) and cve for cve in overdue) or len(overdue) != len(set(overdue)) or len(overdue) != row["OverdueKEVs"] or not set(overdue).issubset(cves):
                raise ValueError("OverdueCves does not reconcile to OverdueKEVs and KEVs")
    return records


def read_legacy(path, report_date):
    book = read_xlsx(path)
    annotations, counts = [], {}
    palette = {"FF00B0F0": "Blue", "FF92D050": "Green", "FFFFFF00": "Yellow", "FFFFC000": "Orange"}
    for row in book["Sheet2"]:
        v = row["values"]
        name = v.get("A", "")
        if name in ("", "Row Labels", "(blank)", "Grand Total"):
            continue
        key = normal(name)
        if key in counts:
            raise ValueError(f"Ambiguous normalized legacy name: {name}")
        counts[key] = {"name": name, "count": int(v["B"])}
        color = palette.get(row["colors"].get("A"), "")
        if color or v.get("C") or v.get("D"):
            annotations.append({"date": report_date, "name": name, "color": color,
                                "note": v.get("C", ""), "ticket": v.get("D", ""),
                                "source": f"{Path(path).name}, Sheet2!A{row['row']}:D{row['row']}"})
    return annotations, counts


def legacy_tracking(events):
    merged = {}
    for event in sorted(events, key=lambda x: x["date"]):
        key = normal(event["name"])
        t = merged.setdefault(key, {h: "" for h in TRACK_HEADERS})
        t["DeviceName"] = event["name"]
        t["RemediationStatus"] = t["RemediationStatus"] or "—"
        t["NeedsEscalation"] = t["NeedsEscalation"] or "No"
        if event["color"]:
            t["LegacyColor"] = event["color"]
            t["NeedsEscalation"] = "Yes" if event["color"] == "Blue" else "No"
            status = {"Green": "Resolved", "Yellow": "Ticket open", "Orange": "In progress"}.get(event["color"])
            if status:
                t["RemediationStatus"] = status
        if event["note"]:
            t["Circumstances"] = event["note"]
        if event["ticket"]:
            value = str(event["ticket"]).strip()
            if re.fullmatch(r"INC\d+", value, re.IGNORECASE):
                tickets = t["ServiceNowTickets"].split("; ") if t["ServiceNowTickets"] else []
                if value not in tickets:
                    tickets.append(value)
                t["ServiceNowTickets"] = "; ".join(tickets)
            else:
                t["Notes"] = (t["Notes"] + "; " if t["Notes"] else "") + "Legacy ticket field: " + value
        t["AnnotationSource"] += ("; " if t["AnnotationSource"] else "") + event["source"]
    return list(merged.values())


def merge_tracking(records, tracking):
    by_name = collections.defaultdict(list)
    for row in records:
        by_name[normal(row["DeviceName"])].append(row)
    existing, orphan_names = {}, set()
    for t in tracking:
        if t.get("DeviceId"):
            t["DeviceId"] = normal(t["DeviceId"])
            if t["DeviceId"] in existing:
                raise ValueError("Duplicate DeviceId in Tracking; resolve before refreshing")
            existing[t["DeviceId"]] = t
        else:
            name = normal(t.get("DeviceName"))
            if not name or name in orphan_names:
                raise ValueError("Blank or duplicate name in unmapped Tracking records")
            orphan_names.add(name)
    for t in tracking:
        if not t.get("DeviceId"):
            matches = by_name.get(normal(t["DeviceName"]), [])
            if len(matches) == 1 and matches[0]["DeviceId"] not in existing:
                t["DeviceId"] = matches[0]["DeviceId"]
                t["IdentityMatch"] = "Unique name match; verify"
                existing[t["DeviceId"]] = t
            else:
                t["IdentityMatch"] = "No current match" if not matches else "Ambiguous name; review"
    for row in records:
        if row["DeviceId"] not in existing:
            t = {h: "" for h in TRACK_HEADERS}
            t.update(DeviceId=row["DeviceId"], DeviceName=row["DeviceName"], RemediationStatus="—",
                     NeedsEscalation="No", IdentityMatch="DeviceId from export")
            tracking.append(t)
            existing[row["DeviceId"]] = t
        else:
            existing[row["DeviceId"]]["DeviceName"] = row["DeviceName"]
    for t in tracking:
        t['RemediationStatus'] = normalize_manual('RemediationStatus', t.get('RemediationStatus'))
        if t["RemediationStatus"] not in STATUSES:
            raise ValueError(f"Unknown remediation status: {t['RemediationStatus']}")
        if t["NeedsEscalation"] not in ("Yes", "No"):
            raise ValueError("NeedsEscalation must be Yes or No")
    return sorted(tracking, key=lambda t: (normal(t["DeviceName"]), t["DeviceId"]))


def load_saved_state(book):
    tracking = table_records(book['Tracking'], 'DeviceId')
    records = table_records(book['Defender import'], 'DeviceId')
    if len({normal(r['DeviceId']) for r in records}) != len(records):
        raise ValueError('Duplicate DeviceIds in the saved Defender import')
    guide = {r['Item']: r['Details'] for r in table_records(book['Guide'], 'Item')}
    weekly = table_records(book['Weekly report'], 'Device')
    history = table_records(book['History'], 'Report date') if 'History' in book else []
    header = next(r for r in book['Weekly report'] if r['values'].get('A') == 'Device')['values']
    id_column = next(col for col, label in header.items() if label == 'DeviceId')
    formula_fields = {normal(r['values'].get(id_column)): {header[col] for col in r.get('formulas', {}) if col in header}
                      for r in book['Weekly report'] if r['row'] > 8 and r['values'].get(id_column)}
    tracking, history, changes = collect_saved_edits(tracking, weekly, records, history, guide, formula_fields)
    return tracking, records, guide, history, changes


def prepare(export, report_date, previous=None, legacy_paths=()):
    records = read_export(export)
    all_names = collections.Counter(normal(x["DeviceName"]) for x in records)
    events, baseline, comparison = [], {}, {}
    previous_records, previous_date, history = [], "", []
    previous_weekly_ids = set()
    if previous:
        book = read_xlsx(previous)
        tracking, previous_records, guide, history, _ = load_saved_state(book)
        previous_weekly_ids = {normal(r['DeviceId']) for r in table_records(book['Weekly report'], 'Device')}
        previous_date = str(guide.get("Report date", ""))
        if previous_date and report_date < previous_date:
            raise ValueError("Report date precedes the tracking workbook; refusing to roll tracking backward")
    else:
        for path, date in legacy_paths:
            ann, counts = read_legacy(path, date)
            events.extend(ann)
            if not baseline:
                baseline = counts
                previous_date = date
            else:
                comparison = counts
        tracking = legacy_tracking(events)
    known_ids = {normal(t['DeviceId']) for t in tracking if t.get('DeviceId')}
    tracking = merge_tracking(records, tracking)
    tracked = {t["DeviceId"]: t for t in tracking if t["DeviceId"]}
    current_ids = {r["DeviceId"] for r in records}
    prev = {normal(r["DeviceId"]): r for r in previous_records}
    reviews = []

    def review(kind, name, device_id, detail):
        reviews.append([kind, name, device_id, detail])

    for t in tracking:
        if not t["DeviceId"] or t["DeviceId"] not in current_ids:
            if t["LegacyColor"] or any(t.get(h) for h in ("ServiceNowTickets", "Circumstances", "Notes", *ASSET_FIELDS)) or t["RemediationStatus"] != "—" or t["NeedsEscalation"] == "Yes":
                review("Tracked device not observed", t["DeviceName"], t["DeviceId"], "Tracking retained. Absence from this export does not establish remediation.")
        if t["IdentityMatch"] in ("Unique name match; verify", "Unverified DeviceId"):
            review("Verify migrated identity", t["DeviceName"], t["DeviceId"], "Legacy annotations matched by unique full name. Verify, then set IdentityMatch to Verified DeviceId.")
        if t["NeedsEscalation"] == "Yes" and not t["EscalationReason"]:
            review("Escalation reason needed", t["DeviceName"], t["DeviceId"], "Blue legacy highlight preserved. Add the reason progress is difficult.")
        if t["RemediationStatus"] == "Ticket open" and not t["ServiceNowTickets"]:
            review("Ticket number needed", t["DeviceName"], t["DeviceId"], "Ticket-open status has no ServiceNow ticket reference.")
    output = []
    for row in records:
        device_id, key = row["DeviceId"], normal(row["DeviceName"])
        t = tracked[device_id]
        previous_count = None
        comparison_method = "No comparable previous record"
        if device_id in prev:
            previous_count = prev[device_id]["OverdueKEVs"]
            comparison_method = "DeviceId"
        elif baseline and key in baseline and all_names[key] == 1:
            previous_count = baseline[key]["count"]
            comparison_method = "Legacy name match"
        if all_names[key] > 1:
            review("Shared hostname", row["DeviceName"], device_id, "Multiple current DeviceIds share this name. Annotations are joined by DeviceId only.")
        if t["RemediationStatus"] == "Resolved" and row["OverdueKEVs"] > 0:
            review("Resolution needs checking", row["DeviceName"], device_id, "Recorded resolved; current export still contains overdue KEVs. Check finding dates and ticket scope.")
        if t["RemediationStatus"] in ("Ticket open", "In progress") and row["OverdueKEVs"] == 0:
            review("No overdue findings observed", row["DeviceName"], device_id, "No overdue KEVs in this export. Verify remediation and ticket scope before changing the recorded status.")
        if len(json.loads(row["KEVs"])) < row["KEVCount"]:
            review("CVE list incomplete", row["DeviceName"], device_id, "Count exceeds exported CVE list length. Rerun the revised query before CVE-level comparison.")
        if row.get("MissingDueDateKEVs", 0):
            review("CISA due date missing", row["DeviceName"], device_id, "At least one KEV has no due date and cannot be classified as overdue.")
        if row["OverdueKEVs"] > 0:
            output.append({"device": row, "previous": previous_count, "comparison": comparison_method,
                           "appearance": '' if device_id in previous_weekly_ids else 'Returning' if device_id in known_ids else 'New',
                           "recentUsers": format_recent_users(row.get("RecentUsers", ""), row.get("OSPlatform", ""))})
    if previous:
        for device_id, row in prev.items():
            if device_id not in current_ids:
                review("Previous device not observed", row["DeviceName"], device_id, "Not in the current export. Confirm inventory and query coverage before declaring resolved.")
    else:
        for key, item in baseline.items():
            if key not in all_names:
                review("Previous name not observed", item["name"], "", "Present in the previous legacy report, absent from this export. Identity and query scope need review.")
    output.sort(key=lambda x: (-x["device"]["OverdueKEVs"], -int(bool(x["device"].get("IsInternetFacing"))), -int(x["device"].get("RansomwareLinkedKEVs", 0) or 0), normal(x["device"]["DeviceName"]), x["device"]["DeviceId"]))
    reconciliation = []
    if comparison:
        overdue_names = collections.Counter()
        for r in records:
            if r["OverdueKEVs"]:
                overdue_names[normal(r["DeviceName"])]+=r["OverdueKEVs"]
        for key in sorted(comparison.keys() | overdue_names.keys()):
            old_count = comparison.get(key, {}).get("count", 0)
            new_count = overdue_names.get(key, 0)
            if old_count != new_count:
                name = comparison.get(key, {}).get("name", next((r["DeviceName"] for r in records if normal(r["DeviceName"]) == key), key))
                reconciliation.append([name, old_count, new_count, new_count-old_count, "Absent from new export" if key not in overdue_names else "Shared hostname" if all_names[key] > 1 else "Counts differ"])
        if sum(r[3] for r in reconciliation) != sum(r["OverdueKEVs"] for r in records)-sum(v["count"] for v in comparison.values()):
            raise AssertionError("Reconciliation does not balance")
    payload = {"reportDate": report_date, "previousDate": previous_date,
            "source": Path(export).name, "trackingSource": Path(previous).name if previous else "Legacy workbooks",
            "records": records, "tracking": tracking, "trackingHeaders": TRACK_HEADERS,
            "report": output, "reviews": reviews, "reconciliation": reconciliation,
            "legacyAlertLabels": any("IncidentsLast6Months" in r for r in records),
            "summary": {"sourceDevices": len(records), "reportDevices": len(output),
                        "overdue": sum(r["OverdueKEVs"] for r in records),
                        "allKevs": sum(r["KEVCount"] for r in records),
                        "legacyAnnotations": sum(bool(t.get("LegacyColor") or t.get("AnnotationSource")) for t in tracking),
                        "trackingRows": len(tracking)}}
    return finalize_snapshot(payload, history)


def prepare_existing_report(path):
    """Rebuild a saved report's presentation without changing its snapshot or comparisons."""
    book = read_xlsx(path)
    guide = {r["Item"]: r["Details"] for r in table_records(book["Guide"], "Item")}
    records = table_records(book["Defender import"], "DeviceId")
    tracking, records, guide, history, changes = load_saved_state(book)
    report_rows = table_records(book["Weekly report"], "Device")
    by_id = {normal(row["DeviceId"]): row for row in records}
    if len(by_id) != len(records):
        raise ValueError("Duplicate DeviceIds in the saved Defender import")
    report = []
    appearances = {normal(h['DeviceId']): h.get('Weekly appearance', '') for h in history
                   if h.get('SnapshotId') == guide.get('Snapshot id') and h.get('Event') == 'Generated'}
    saved_rows = {normal(row['DeviceId']): row for row in report_rows}
    order = {normal(row['DeviceId']): index for index, row in enumerate(report_rows)}
    devices = sorted((r for r in records if r['OverdueKEVs'] > 0),
                     key=lambda r: (-r['OverdueKEVs'], order.get(normal(r['DeviceId']), len(order))))
    for device in devices:
        row = saved_rows.get(normal(device['DeviceId']), {})
        previous = row.get("Previous overdue", "")
        report.append({"device": device, "previous": previous if previous != "" else None,
                       "comparison": "Saved comparison",
                       "appearance": appearances.get(normal(device['DeviceId']), ''),
                       "recentUsers": format_recent_users(device.get("RecentUsers", ""), device.get("OSPlatform", ""))})
    reviews = [[r["Review item"], r["Device"], r["DeviceId"], r["Detail"]]
               for r in table_records(book["Review"], "Review item")]
    for device in devices:
        if normal(device['DeviceId']) not in saved_rows:
            reviews.append(['Report row restored', device['DeviceName'], device['DeviceId'],
                            'Device is still present with overdue KEVs. The deleted report row was restored; its previous comparison is unavailable.'])
    payload = {"reportDate": guide["Report date"], "previousDate": guide.get("Previous report date", ""),
            "source": guide["Defender source"], "trackingSource": Path(path).name,
            "records": records, "tracking": tracking, "trackingHeaders": TRACK_HEADERS,
            "report": report, "reviews": reviews, "reconciliation": [],
            "legacyAlertLabels": any("IncidentsLast6Months" in r for r in records),
            "summary": {"sourceDevices": len(records), "reportDevices": len(report),
                        "overdue": sum(r["OverdueKEVs"] for r in devices),
                        "allKevs": sum(r["KEVCount"] for r in records),
                        "legacyAnnotations": sum(bool(t.get("LegacyColor") or t.get("AnnotationSource")) for t in tracking),
                        "trackingRows": len(tracking)}}
    return finalize_snapshot(payload, history)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--export")
    parser.add_argument("--date")
    parser.add_argument("--existing-report")
    parser.add_argument("--previous")
    parser.add_argument("--legacy", action="append", default=[])
    parser.add_argument("--json", required=True)
    args = parser.parse_args()
    if args.existing_report:
        if args.export or args.date or args.previous or args.legacy:
            parser.error("--existing-report cannot be combined with export/date/previous/legacy inputs")
        result = prepare_existing_report(args.existing_report)
    else:
        if not args.export or not args.date:
            parser.error("--export and --date are required for a new snapshot")
        dt.date.fromisoformat(args.date)
        legacy = [(item.split("|", 1)[1], item.split("|", 1)[0]) for item in args.legacy]
        result = prepare(args.export, args.date, args.previous, legacy)
    Path(args.json).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(result["summary"]))


if __name__ == "__main__":
    main()
