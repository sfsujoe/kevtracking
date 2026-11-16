# KEV Tracking

A Windows utility for building weekly device reports from Microsoft Defender for Endpoint advanced hunting exports, with directly editable remediation tracking and retained device history.

See the [project runbook](PROJECT-RUNBOOK.md) for weekly operation, setup, troubleshooting, recovery and maintenance handoff.

## What it does

- Imports Defender CSV or XLSX results and ranks devices by overdue CISA Known Exploited Vulnerabilities (KEVs).
- Lets staff edit status, ServiceNow ticket references, escalation details, notes, review dates, SFSU Tag, AT Tag and Serial No. on the **Weekly report** sheet.
- Carries saved edits forward by **DeviceId**, including intentional clearing of fields.
- Retains absent devices in a hidden **Tracking** register and records periodic snapshots and saved edits in **History**.
- Archives the previous saved report before generating a new one. Existing output files are never overwritten.
- Flags duplicate identities, conflicting edits, missing observations and recorded resolutions that still have overdue findings.
- Formats recent user labels as `912345678 (Example User), localusername`.

The utility processes local exports. It does not connect directly to Defender or ServiceNow. Ticket references are entered manually; history records saved state at refresh rather than every Excel keystroke.

## Requirements

- Windows with Windows PowerShell 5.1 for the desktop form.
- The Codex desktop bundled runtime, including Node.js, Python and `@oai/artifact-tool`, available under `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies`.
- Access to Defender advanced hunting to run the supplied KQL and download the results.

This initial implementation depends on the bundled runtime. It is not yet a standalone installer. Dependencies are not vendored or downloaded automatically.

## Start

1. Run [the KQL query](pilot/queries/DevicesByKEV.kql) in Defender Advanced Hunting with a 30-day hunting range, then download the CSV export. Validate query availability against your tenant's schema and permissions.
2. Double-click `pilot\Refresh-KEVReport.cmd` to open the form. The launcher uses process-scoped `ExecutionPolicy Bypass`.
3. Choose the Defender export, report date, output directory and filename. On the first run, leave **Previous report** blank. For later runs, select the most recent saved working report.
4. Create the report, then edit the fields with amber headers on **Weekly report**.
5. Save and close the workbook before the next refresh, and select that workbook as the previous report.

For command-line use:

```powershell
# First report
.\pilot\Refresh-KEVReport.cmd `
  -ExportPath 'C:\Reports\Defender-export.csv' `
  -ReportDate '2026-11-16' `
  -OutputPath 'C:\Reports\KEV-Report-2026-11-16.xlsx'

# Subsequent report, preserving saved edits and history
.\pilot\Refresh-KEVReport.cmd `
  -ExportPath 'C:\Reports\Defender-export-new.csv' `
  -ReportDate '2026-11-23' `
  -TrackingWorkbook 'C:\Reports\KEV-Report-2026-11-16.xlsx' `
  -OutputPath 'C:\Reports\KEV-Report-2026-11-23.xlsx'
```

`-UpgradeReport -TrackingWorkbook <saved-report.xlsx> -OutputPath <new-report.xlsx>` converts a previously generated report to the current editing model using its saved Defender snapshot.

## Workbook behavior

**Weekly report** shows records with `OverdueKEVs > 0`. Its manual fields are values, and the colors and row review message react immediately to edits. Imported device facts and DeviceId must stay attached to the correct row; sort the complete table.

Remediation status defaults to **—**. Blank status cells and older `Unspecified` values become **—** on refresh. **SFSU Tag**, **AT Tag** and **Serial No.** sit together immediately before DeviceId. Enter them manually as text; blank values remain blank. They carry forward by DeviceId from saved weekly reports and are retained in device history, including while a device is absent. Values are never inferred from external exports.

**Tracking** is a hidden device register. **History** retains generated snapshots and saved edits, including the CVE sets available in the export. **Defender import** retains source data, **Review** lists refresh-time exceptions, and **Guide** explains the workflow.

Pink rows mark devices new or returning to Weekly report compared with the previous saved Weekly report, including returns after zero overdue findings. Yellow means ticket open, orange means in progress, green means recorded resolved, and blue marks a device needing escalation. Status and escalation colors take priority in their cells. A recorded resolution with overdue findings is flagged for review. Disappearance from an export does not establish remediation.

## Copying the tool

Copy `pilot/Refresh-KEVReport.cmd` and the entire `pilot/scripts/` folder together to a writable Windows location. Tests and verification utilities live separately in `pilot/tests/` and are not needed to run the tool. The launcher resolves its files relative to itself, including when started from another working directory. Runtime dependencies listed above must still be available. Archives, diagnostics and the runtime package junction are created beside the CMD, outside `scripts`. Transfer exports and the latest saved working workbook separately; the query can be run from the repository or saved in Defender.

Use one current working report. Separate edited copies are not merged automatically. Conflicting report/register changes and duplicate DeviceIds stop generation instead of silently selecting a value.

## Data and repository contents

Actual device exports, generated workbooks, screenshots, archives, logs, runtime packages and private fixture-dependent checks are excluded by `.gitignore`. The published regression tests use synthetic data. Keep operational files in their generated local folders or another access-controlled location.

## Tests

The Python tests use the standard library:

```powershell
python -B -m unittest discover -s pilot/tests -p 'test_public.py' -v
python -B -m unittest discover -s pilot/tests -p 'test_tracking_state.py' -v
python -B -m unittest discover -s pilot/tests -p 'test_portable.py' -v
node --test pilot/tests/test_export_inspection.mjs
```

On Windows, exercise the form handlers and hidden worker with synthetic fixtures:

```powershell
powershell.exe -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass `
  -File pilot/tests/Test-ReportGui.ps1
```

The portable test requires the bundled runtime. It copies only the CMD and scripts to a separate folder, exercises multiple synthetic refreshes, manual edits, returns and report upgrade, and retains its diagnostics under `pilot/validation`.

Report generation additionally validates totals and comparisons, scans formula errors, and renders previews under `pilot/validation`. The finalization step hides the register by changing only its workbook metadata; all cells and formulas are authored with `@oai/artifact-tool`.
