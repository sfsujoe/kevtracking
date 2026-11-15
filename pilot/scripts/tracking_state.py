"""DeviceId-based edit reconciliation and periodic device-state history."""
import copy
import datetime as dt
import json
import uuid

MANUAL_COLUMNS = {
    'RemediationStatus': 'Remediation status', 'ServiceNowTickets': 'ServiceNow tickets',
    'NeedsEscalation': 'Escalation', 'EscalationReason': 'Escalation reason',
    'Circumstances': 'Circumstances', 'Notes': 'Notes', 'LastReviewed': 'Last reviewed',
    'IdentityMatch': 'Identity match',
}
STATUSES = ['Unspecified', 'Ticket open', 'In progress', 'Resolved']
IDENTITY_VALUES = ['Verified DeviceId', 'DeviceId from export', 'Unique name match; verify',
                   'No current match', 'Ambiguous name; review', 'Unverified DeviceId']
REGISTER_EXTRA = ['TrackingKey', 'LastObservedReportDate', 'LastOverdueKEVs', 'LastKEVCount',
                  'LastCves', 'LastOverdueCves', 'ObservedInExport']
HISTORY_HEADERS = ['Report date', 'Device', 'Status', 'ServiceNow tickets', 'Overdue KEVs',
                   'All KEVs', 'Observed in export', 'Event', 'Changed fields', 'Notes',
                   'Escalation', 'Escalation reason', 'Circumstances', 'Last reviewed',
                   'DeviceId', 'SnapshotId', 'Captured UTC', 'Identity match', 'KEVs',
                   'OverdueCves', 'TrackingKey']
HISTORY_MANUAL = dict(MANUAL_COLUMNS, RemediationStatus='Status')


def norm(value):
    return str(value or '').strip().lower()


def utc_now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z')


def normalize_manual(field, value):
    if value is None or (isinstance(value, str) and not value.strip()):
        return {'RemediationStatus': 'Unspecified', 'NeedsEscalation': 'No',
                'IdentityMatch': 'Unverified DeviceId'}.get(field, '')
    if field == 'LastReviewed':
        if isinstance(value, (int, float)) and not isinstance(value, bool) and 0 <= value <= 2958465:
            return value
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
            try:
                return (dt.datetime.strptime(str(value).strip(), fmt) - dt.datetime(1899, 12, 30)).days
            except ValueError:
                pass
        raise ValueError('Last reviewed must be an Excel date or yyyy-mm-dd date')
    value = str(value).strip()
    if field == 'RemediationStatus' and value not in STATUSES:
        raise ValueError(f'Unknown remediation status: {value}')
    if field == 'NeedsEscalation' and value not in ('Yes', 'No'):
        raise ValueError('Escalation must be Yes or No')
    if field == 'IdentityMatch' and value not in IDENTITY_VALUES:
        raise ValueError(f'Unknown identity match value: {value}')
    return value


def seed_register(tracking, records, report_date):
    source = {norm(r['DeviceId']): r for r in records}
    result, ids, keys = [], set(), set()
    for original in tracking:
        row = copy.deepcopy(original)
        row['DeviceId'] = norm(row.get('DeviceId'))
        if row['DeviceId']:
            if row['DeviceId'] in ids:
                raise ValueError('Duplicate DeviceId in Tracking; resolve before refreshing')
            ids.add(row['DeviceId'])
        key = row.get('TrackingKey') or row['DeviceId'] or 'legacy:' + norm(row['DeviceName'])
        if key in keys:
            raise ValueError('Duplicate internal tracking key')
        row['TrackingKey'] = key
        keys.add(key)
        for field in MANUAL_COLUMNS:
            row[field] = normalize_manual(field, row.get(field, ''))
        for field in REGISTER_EXTRA:
            row.setdefault(field, '')
        current = source.get(row['DeviceId'])
        row['ObservedInExport'] = current is not None
        if current:
            row.update(LastObservedReportDate=report_date, LastOverdueKEVs=current['OverdueKEVs'],
                       LastKEVCount=current['KEVCount'], LastCves=current['KEVs'],
                       LastOverdueCves=current.get('OverdueCves', ''))
        result.append(row)
    return result


def history_rows(tracking, records, snapshot_id, report_date, captured, event, changes=None):
    source = {norm(r['DeviceId']): r for r in records}
    rows = []
    for t in tracking:
        key = t['TrackingKey']
        if changes is not None and key not in changes:
            continue
        current = source.get(t['DeviceId'])
        row = dict.fromkeys(HISTORY_HEADERS, '')
        row.update({'Report date': report_date, 'Device': t['DeviceName'], 'DeviceId': t['DeviceId'],
                    'TrackingKey': key, 'SnapshotId': snapshot_id, 'Captured UTC': captured,
                    'Event': event, 'Observed in export': current is not None,
                    'Changed fields': ', '.join(MANUAL_COLUMNS[f] for f in changes.get(key, [])) if changes else ''})
        if current:
            row.update({'Overdue KEVs': current['OverdueKEVs'], 'All KEVs': current['KEVCount'],
                        'KEVs': current['KEVs'], 'OverdueCves': current.get('OverdueCves', '')})
        for field, column in HISTORY_MANUAL.items():
            row[column] = t[field]
        rows.append(row)
    return rows


def collect_saved_edits(tracking, weekly, records, history, guide, formula_fields=None, captured=None):
    """Three-way merge: generated baseline, backing register, saved report values."""
    captured = captured or utc_now()
    previous_date = str(guide.get('Report date', ''))
    register = seed_register(tracking, records, previous_date)
    by_id = {r['DeviceId']: r for r in register if r['DeviceId']}
    source = {norm(r['DeviceId']): r for r in records}
    formula_fields = formula_fields or {}
    history = copy.deepcopy(history)
    modern = str(guide.get('Schema version', '')) == '2'
    snapshot_id = str(guide.get('Snapshot id') or uuid.uuid4())
    baselines = {}
    if modern:
        for h in history:
            if h['SnapshotId'] == snapshot_id and h['Event'] == 'Generated':
                if h['TrackingKey'] in baselines:
                    raise ValueError('Duplicate history baseline for the current snapshot')
                baselines[h['TrackingKey']] = {f: normalize_manual(f, h[HISTORY_MANUAL[f]]) for f in MANUAL_COLUMNS}
        if any(r['TrackingKey'] not in baselines for r in register):
            raise ValueError('History baseline is missing for a tracked device; choose an intact saved report')
    else:
        baselines = {r['TrackingKey']: {f: r[f] for f in MANUAL_COLUMNS} for r in register}
        history.extend(history_rows(register, records, snapshot_id, previous_date, captured, 'Migrated baseline'))
    weekly_by_id = {}
    for row in weekly:
        device_id = norm(row.get('DeviceId'))
        if not device_id or device_id not in by_id or device_id not in source:
            raise ValueError('Weekly report contains a missing or unrecognized DeviceId; restore the original device identity')
        if device_id in weekly_by_id:
            raise ValueError('Duplicate DeviceId in Weekly report; resolve conflicting or copied rows before refreshing')
        original = source[device_id]
        for column, field in [('Device', 'DeviceName'), ('Overdue KEVs', 'OverdueKEVs'),
                              ('All KEVs', 'KEVCount'), ('Exposure', 'ExposureLevel'), ('OS', 'OSPlatform')]:
            if column in row and row[column] != original.get(field, ''):
                raise ValueError(f'Imported field changed for {device_id}: {column}. Edit only the manual report columns.')
        weekly_by_id[device_id] = row
    changes = {}
    for t in register:
        baseline = baselines[t['TrackingKey']]
        report_row = weekly_by_id.get(t['DeviceId'])
        changed = []
        for field, column in MANUAL_COLUMNS.items():
            current = t[field]
            candidate = current
            if report_row is not None and column in report_row:
                has_formula = column in formula_fields.get(t['DeviceId'], set())
                if modern and has_formula:
                    raise ValueError(f'Manual field {column} contains a formula; enter a value instead')
                # Old report formulas are views, not user edits. Literal edits
                # that replaced those formulas are migrated deliberately.
                if not has_formula:
                    candidate = normalize_manual(field, report_row[column])
                    if current != baseline[field] and candidate != baseline[field] and current != candidate:
                        raise ValueError(f'Conflicting edits for {t["DeviceName"]}: {column}. Resolve the report/register disagreement.')
                    if candidate == baseline[field]:
                        candidate = current
            if candidate != baseline[field]:
                changed.append(field)
            t[field] = candidate
        if changed:
            changes[t['TrackingKey']] = changed
    history.extend(history_rows(register, records, snapshot_id, previous_date, captured, 'Saved edits', changes))
    return register, history, changes


def finalize_snapshot(payload, history, captured=None):
    captured = captured or utc_now()
    payload['tracking'] = seed_register(payload['tracking'], payload['records'], payload['reportDate'])
    payload['schemaVersion'] = 2
    payload['snapshotId'] = str(uuid.uuid4())
    payload['generatedUtc'] = captured
    payload['historyHeaders'] = HISTORY_HEADERS
    payload['history'] = list(history) + history_rows(payload['tracking'], payload['records'],
        payload['snapshotId'], payload['reportDate'], captured, 'Generated')
    return payload
