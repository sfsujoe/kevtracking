# KEV reporting runbook

Updated: 2026-11-16
Repository: [sfsujoe/kevtracking](https://github.com/sfsujoe/kevtracking)
Baseline before this update: `c5c8418` on `main`.

Current implementation: portable CMD/scripts layout, unset status shown as **—**, pink new/returning rows, manual asset identifiers with retained history and export-diagnostic moves across filesystem boundaries.

## Purpose and current state

Create the University's weekly device report from Microsoft Defender for Endpoint (MDE), while retaining remediation work, ServiceNow ticket references and device history. Staff edit the **Weekly report** directly rather than maintaining a separate visible tracking sheet or recoloring rows each week.

The pilot is implemented and published. It accepts native MDE CSV exports or XLSX exports, supplies file and directory pickers, archives the previous saved report and refuses to overwrite output files. It processes local files; Defender querying, export download and ticket updates remain manual. No scheduled job or live ServiceNow integration is configured.

The project folder on the original workstation is `%USERPROFILE%\src\Vulnerability Reporting`. Commands below assume PowerShell is open in the project folder. Actual device data remains local and is excluded from Git.

## Weekly operator procedure

1. Identify the **single current working workbook**. Save all edits in Excel and close it. Use that saved file as **Previous report**; the utility cannot read unsaved changes or merge multiple edited copies.
2. In Defender Advanced Hunting, run [`pilot/queries/DevicesByKEV.kql`](pilot/queries/DevicesByKEV.kql) with the hunting range set to **30 days**. Download the native CSV results into an access-controlled location. Keep its column names, arrays and identifiers intact; no conversion to Excel is required.
3. Double-click `pilot\Refresh-KEVReport.cmd`.
4. Choose the Defender export, the intended report date, the previous saved workbook, an existing output folder and a **new** `.xlsx` filename, such as `KEV-Report-2026-11-16.xlsx`. Check the previous-file choice even if the form suggests one. Reports saved outside `pilot` may need to be located manually.
5. Click **Create report** and wait for the form to report success. Retain the input and any diagnostic files if generation fails.
6. Open the new workbook in Excel. Confirm the report date and source on **Guide**, then check the device count, overdue totals, large changes and entries on **Review**. Compare unexpected changes with the export before distributing the report.
7. Work in **Weekly report**, using the editable fields described below. Save the workbook. This becomes the current working report and the input to the next refresh.

For the first report in a fresh checkout, leave **Previous report** blank. On the original workstation, supply the latest saved pilot report so existing annotations are retained. Blank previous input can trigger a one-time migration when both original November annotated workbooks are present. If a `KEV-Report-*.xlsx` already exists directly in `pilot`, the utility requires a previous report to guard against losing its edits.

The report date is a label for the run. Changing it does **not** retrieve historical vulnerability data or recalculate the export's overdue counts.

## Editing and interpreting the workbook

| Sheet | Operator use |
| --- | --- |
| **Weekly report** | Current devices with `OverdueKEVs > 0`; edit remediation fields here. |
| **History** | Retained snapshots and saved manual changes captured when a workbook is processed. Filter by DeviceId to follow a device over time. |
| **Defender import** | Source snapshot, including KEV-bearing devices that are not yet overdue. Use to investigate counts. |
| **Review** | Exceptions captured at generation. Changes made later appear here on the next refresh. |
| **Guide** | Run metadata, definitions and workflow guidance. |
| **Tracking** (hidden) | Persistent register and comparison baseline, including absent devices. Leave it to the utility. |

Editable fields on **Weekly report** are **Remediation status**, **ServiceNow tickets**, **Escalation**, **Escalation reason**, **Circumstances**, **Notes**, **Last reviewed**, **Identity match**, **SFSU Tag**, **AT Tag** and **Serial No.** Their headers are amber; body cells are amber unless another highlight applies. Identity match should change only after someone verifies a migrated association. Last reviewed is entered manually; use an Excel date or `yyyy-MM-dd`. Multiple ticket references can be separated with semicolons.

**SFSU Tag → AT Tag → Serial No.** appear together immediately before DeviceId. Enter these values manually. They are stored as text to preserve leading zeros and remain blank when no value is recorded. The utility carries them forward only from saved weekly reports by DeviceId, preserves intentional clearing and records the saved values in History, including while devices are absent. It never infers asset identifiers from Defender or any other external export. Older workbooks without these columns start with blank asset fields.

| Input | Meaning / appearance |
| --- | --- |
| `—` | No remediation status selected. Blank status cells and older `Unspecified` values become `—` on refresh. |
| `Ticket open` | Yellow; enter the ServiceNow reference. |
| `In progress` | Orange; remediation is actively progressing. |
| `Resolved` | Green; recorded resolution still needs to agree with fresh Defender evidence. |
| Escalation = `Yes` | Blue device/escalation cells; progress is difficult. Add a reason. This is independent of status, so both colors can appear. |
| New or returning device | Pink row when absent from the previous saved Weekly report, including returns after zero overdue KEVs. Status and escalation colors take priority in their cells. On a first report, all rows are new. |

Change the values to control colors. Manually applied fill colors are not tracking inputs. Clearing a manual field intentionally clears its carried-forward value. Sort the **entire table**, keeping DeviceId and imported facts with their row. Do not replace identifiers, edit imported counts or remove workbook sheets/headers to change report scope; use Excel filters instead.

Colors and the row **Review** formula respond to edits in Excel. Imported facts, comparisons, the hidden register, History and the separate Review queue update when the utility next processes the saved workbook. History is a refresh-time record, not an audit of every Excel edit or an authenticated editor log.

Recent users are formatted as `912345678 (Example User), macusername, localusername` when the source provides the necessary identity information. These are recent logons, not a verified device owner. **Exposure** is the imported Defender `ExposureLevel`, a broader assessment than the device's KEV count; see Guide for the implementation's documented interpretation.

## Identity, counts and history rules

- Normal carry-forward uses **DeviceId**. A new DeviceId with the same hostname does not automatically inherit the previous device's tickets or notes. Some original annotations were migrated by unique name and require identity review.
- Counts represent unique device/CVE pairs, not vulnerable software installations. The Weekly report ranks overdue findings; **All KEVs** includes KEVs that are not yet overdue.
- The source query returns devices with at least one KEV. A device with no remaining KEVs, missing telemetry or changed query visibility may disappear. Disappearance preserves its tracking record and is **not** automatic proof of resolution.
- A device still present in the export with zero overdue KEVs leaves Weekly report but remains in the source and tracking state. Review any open-ticket follow-up.
- A recorded `Resolved` device with overdue findings is flagged. Verify the findings and adjust the recorded status or remediation work as appropriate.
- Duplicate identities, conflicting manual edits and damaged history/baseline data stop generation for investigation. Do not bypass these checks by starting a new report without its history.

## Command-line operation and migration

The CMD launcher uses `ExecutionPolicy Bypass` for its PowerShell process; it does not change machine policy. It also forwards command-line arguments. Example weekly refresh:

```powershell
.\pilot\Refresh-KEVReport.cmd `
  -ExportPath 'C:\Reports\Defender-export-2026-11-16.csv' `
  -ReportDate '2026-11-16' `
  -TrackingWorkbook 'C:\Reports\KEV-Report-2026-11-09.xlsx' `
  -OutputPath 'C:\Reports\KEV-Report-2026-11-16.xlsx'
```

To convert a previously generated pilot report to the current editing model using its **saved source snapshot**, use a new filename. This does not fetch new Defender findings:

```powershell
.\pilot\Refresh-KEVReport.cmd -UpgradeReport `
  -TrackingWorkbook 'C:\Reports\KEV-Report-previous.xlsx' `
  -OutputPath 'C:\Reports\KEV-Report-editable.xlsx'
```

## Setup on another workstation

1. Clone the repository for maintenance, or copy only `pilot\Refresh-KEVReport.cmd` and the entire `pilot\scripts\` folder together to a writable Windows deployment folder. Tests and verification utilities are in `pilot\tests\` and are not required at the deployment location. The CMD resolves all tool paths relative to itself, independent of the current working directory.
2. Ensure Windows PowerShell 5.1 is available for the WinForms interface.
3. Open the project in Codex and provision/check its bundled workspace dependencies. The current launcher expects:
   - `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`
   - `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe`
   - `%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules` containing `@oai/artifact-tool`.
4. Confirm access to run the KQL in Defender. Table availability, permissions and external CISA CSV access must work in the University's tenant.
5. Transfer the latest saved working workbook separately through approved University storage if continuing existing tracking. Git contains no operational workbooks.
6. Run the synthetic checks below and generate a report to a new filename before using the workstation for the weekly process.

The utility creates a `node_modules` junction beside the CMD to the bundled packages and can repoint its own outdated junction. Archives and diagnostics are also created beside the CMD, outside `scripts`. The paths below use `pilot` as that deployment folder. Exports and previous working reports are selected separately; the KQL can be run from the repository or saved in Defender. Dependencies are not installed automatically. This is currently a Codex-runtime-dependent pilot, not a standalone packaged application.

## Files, backups and diagnostics

| Location | Contents / retention use |
| --- | --- |
| Chosen output path | Generated workbook; the saved working copy is the source of manual tracking for the next run. |
| `pilot\archive\` | Timestamped, uniquely named copies of the previous saved workbook, created after source preparation and before workbook generation. An archive can exist even when a later generation step fails. |
| `pilot\validation\run-<id>\prepared.json` | Per-run prepared data for investigating validation/generation failures. |
| `pilot\validation\gui-<id>\` | GUI worker logs: `refresh-output.log` and `refresh-error.log`. |
| `pilot\validation\` | Formula scans, inspection results and rendered previews. Some filenames are reused; retain relevant diagnostics before another run. |

Treat exports, archives, prepared JSON, logs and previews as operational device data. Keep them in access-controlled storage. Git exclusions help prevent accidental publication but are not access controls or backups. There is no automatic retention cleanup. Agree on retention and backup ownership before removing old reports, archives or diagnostics.

## Troubleshooting and recovery

| Symptom | Action |
| --- | --- |
| Output already exists | Choose a new filename. Never use the previous working file as the output. |
| Previous report required | Select the latest saved working workbook, including when it lives outside `pilot`. |
| Missing bundled runtime / missing module | Check the runtime paths above and the package junction. Open the project in Codex to investigate dependencies; reinstalling Python alone will not supply the artifact library. |
| Invalid CSV, missing columns, duplicate IDs or mismatched CVE counts | Re-export the supplied query without editing it in Excel. Check whether results were truncated or a different query was used. Preserve the original export for diagnosis. |
| Invalid status/date or conflicting report/register edits | Check the named device and manual field in a copy of the previous workbook. Use valid dropdown values/dates. Reconcile competing copies manually into one authoritative workbook; do not discard the hidden baseline. |
| Older report date rejected | Use a date at least as recent as the selected prior workbook. For historical investigation, preserve the original files rather than rolling the working history backward. |
| Browse dialog error | Enter the full path directly if needed. Record the message and run the GUI checks below. The current implementation handles blank, missing and whitespace paths. |
| Generation fails or a workbook remains after an error | Keep logs and source files. Treat the output as incomplete until investigated. Retry with a new filename from the last known-good saved previous report. |
| `EXDEV` moving `.xlsx.inspect.ndjson` from Box or another mounted folder | Older builders saved the workbook, then failed moving its diagnostic across filesystem boundaries, before hiding Tracking. The current builder falls back to copy-then-delete for this move. Retry with a new output filename and the same last known-good previous report. |
| Finalization cannot hide Tracking | Treat the run as incomplete; inspect logs and the workbook before use. The generated file may exist even though the overall run failed. |
| Notes appear missing | Verify which previous workbook was selected and whether edits were saved. Check DeviceId, History and archived inputs. An older saved input cannot contain later unsaved edits. |

To recover from an incorrect refresh, close the new workbook and retain it for comparison. Locate the last known-good saved report or its copy in `pilot\archive`, confirm it contains the required edits, then rerun with that file as Previous report and a new output filename. Re-enter any later valid edits deliberately; there is no automatic merge of divergent workbook histories. The archive protects the selected saved input, not work that was never saved or never selected.

## Maintainer checks and implementation map

Run these from the project folder. The tests use synthetic data and do not require private University workbooks:

```powershell
$pythonExe = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $pythonExe -B -m unittest discover -s pilot/tests -p 'test_public.py' -v
& $pythonExe -B -m unittest discover -s pilot/tests -p 'test_tracking_state.py' -v
& $pythonExe -B -m unittest discover -s pilot/tests -p 'test_portable.py' -v
& "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --test pilot/tests/test_export_inspection.mjs
powershell.exe -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass `
  -File pilot/tests/Test-ReportGui.ps1
```

At the published baseline, 9 import/user-format tests, 14 tracking/history tests and the GUI checks passed, including from a clean staged source copy. These checks do not validate live tenant connectivity. After changes to workbook generation, also generate a new report and inspect its totals, formulas, editable fields, history and layout. The builder performs reconciliation checks, a formula-error scan and preview rendering.

The local revision adds five tracking/asset/appearance checks and one end-to-end portable deployment check (29 Python tests total). The latter copies only the CMD and scripts to a separate path containing spaces, an ampersand and an apostrophe, then exercises refreshes from another working directory, saved asset edits, absence/return, intentional clearing, report upgrade and generated archives. It also checks saved Excel text formats, status validation and Tracking visibility. Synthetic workbooks and previews remain under `pilot\validation` for inspection. Old schema 2 workbooks are accepted; new reports use schema 3 to include the asset history baseline.

Six Node checks cover export-diagnostic relocation, including simulated `EXDEV`, an absent sidecar and preservation on failures. The portable deployment check also forces `EXDEV` during its first refresh and verifies that diagnostics reach validation and Tracking is hidden before success.

| File | Responsibility |
| --- | --- |
| `pilot/Refresh-KEVReport.cmd` | Process-scoped policy and Windows PowerShell launcher. |
| `pilot/scripts/Refresh-KEVReport.ps1` | Validation, runtime resolution, archive and generation pipeline. |
| `pilot/scripts/ReportGui.ps1` | Pickers, request validation, background worker and progress/error display. |
| `pilot/scripts/prepare_report.py` | CSV/XLSX reading, input checks, legacy migration and report payload. |
| `pilot/scripts/tracking_state.py` | Manual-edit reconciliation, stable identity, retained register and history. |
| `pilot/scripts/build_report.mjs` | Workbook authoring with `@oai/artifact-tool`, formulas, formatting and previews. |
| `pilot/scripts/export_inspection.mjs` | Relocates the optional export diagnostic, including across filesystem boundaries such as Box mounts. |
| `pilot/scripts/finish_workbook.py` | Hides Tracking through a narrow workbook metadata change. |
| `pilot/queries/DevicesByKEV.kql` | Defender query maintained alongside the utility. |
| `pilot/tests/` | Synthetic regression tests and development verification utilities; excluded from deployment. |

Keep real exports and identifiers out of source and synthetic fixtures. Before publishing, review `git status`, stage intended files explicitly and inspect the staged diff. The local fixture-dependent `test_report.py` and `verify_workbook.mjs` are deliberately ignored and are not portable public tests.

## Open decisions and next improvements

- **Query semantics:** confirm whether University reporting should classify items overdue at `dueDate < export time UTC` (current rule) or only after the full due date has passed. Keep the chosen rule consistent across weeks.
- **Source coverage:** verify tenant table access, data retention and query completeness. The revised query uses a 30-day alert window. Older imports labeled `IncidentsLast6Months` count alert IDs, not verified unique incidents or ServiceNow tickets; six-month coverage was not established.
- **Workflow ownership:** agree on who maintains the authoritative workbook, reviews exceptions, verifies resolutions and retains backups. ServiceNow references currently have no validated device/CVE coverage or live status sync.
- **Deployment:** remove or package the dependency on the Codex bundled runtime before wider staff rollout.
- **Automation:** evaluate scheduled/on-demand Defender export retrieval first, then ServiceNow lookups using a reliable device/ticket mapping. Confirm API permissions, licensing and whether ServiceNow Vulnerability Response is available before selecting connectors. Its use has not been confirmed.
- **Shared state and reporting:** if multiple staff need concurrent edits, consider a central tracking store with an audit trail, then Power Automate orchestration and Power BI reporting. Preserve DeviceId-based identity, explicit exceptions and historical snapshots when changing platforms.

No implementation of these future integrations is implied by the current pilot. Preserve the working local workflow while validating each integration separately.
