import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';
import { relocateExportInspection } from './export_inspection.mjs';

const args = process.argv.slice(2);
const payload = JSON.parse(await fs.readFile(args[0], 'utf8'));
const outputPath = path.resolve(args[1]);
await fs.access(outputPath).then(() => { throw new Error('Output already exists. Choose a new dated filename.'); }, e => { if (e.code !== 'ENOENT') throw e; });
const wb = Workbook.create();
const report = wb.worksheets.add('Weekly report');
const history = wb.worksheets.add('History');
const tracking = wb.worksheets.add('Tracking');
const source = wb.worksheets.add('Defender import');
const review = wb.worksheets.add('Review');
const guide = wb.worksheets.add('Guide');
const HEADER = 8, START = 9;
const colors = {Blue:'#00B0F0',Green:'#92D050',Yellow:'#FFFF00',Orange:'#FFC000'};
const count = payload.report.length;
const last = START + count - 1;
const trackLast = START + payload.tracking.length - 1;
const excelText = value => typeof value === 'string' && /^[=+@]/.test(value) ? "'" + value : value;
const safeMatrix = rows => rows.map(row => row.map(v => excelText(v ?? null)));
const column = n => { let s=''; for (n++; n>0; n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; };
const dateValue=value=>typeof value==='string' && /^\d{4}-\d{2}-\d{2}/.test(value) ? new Date(value.length===10?value+'T00:00:00Z':value) : value;

function section(sheet, title, note, headers, rows, tableName) {
  const width = headers.length;
  const end = START + Math.max(rows.length,1)-1;
  sheet.showGridLines = false;
  sheet.getRange(`A1:${column(width-1)}${end}`).format.font = {name:'Arial',size:10,color:'#172C3C'};
  sheet.getRange(`A1:${column(width-1)}${end}`).format.verticalAlignment='center';
  sheet.getRange('A2').values = [[title]];
  sheet.getRange('A2').format.font = {name:'Arial',size:15,bold:true,color:'#172C3C'};
  sheet.getRange('A3').values = [[note]];
  sheet.getRange('A3').format.font = {name:'Arial',size:10,italic:true,color:'#526779'};
  sheet.getRange(`A${HEADER}:${column(width-1)}${HEADER}`).values = [headers];
  sheet.getRange(`A${HEADER}:${column(width-1)}${HEADER}`).format = {fill:'#243E55',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:32,horizontalAlignment:'center',verticalAlignment:'center',wrapText:true};
  if (rows.length) {
    sheet.getRange(`A${START}:${column(width-1)}${end}`).values = safeMatrix(rows);
    sheet.getRange(`A${START}:${column(width-1)}${end}`).format.rowHeight=23;
    if (sheet.name !== 'Guide') {
      sheet.getRange(`A${START}:${column(width-1)}${end}`).format.horizontalAlignment='center';
    }
    const table = sheet.tables.add(`A${HEADER}:${column(width-1)}${end}`,true,tableName);
    table.style='TableStyleLight9';
  }
  sheet.getRange(`A1:${column(width-1)}${end}`).format.columnWidth=20;
  sheet.freezePanes.freezeRows(HEADER);
  return end;
}

const reportHeaders = ['Device','Overdue KEVs','Previous overdue','Change','All KEVs','Recent users','Remediation status','ServiceNow tickets','Escalation','Escalation reason','Circumstances','Notes','Last reviewed','Review','Exposure','OS','Identity match','SFSU Tag','AT Tag','Serial No.','DeviceId'];
const byId=new Map(payload.tracking.filter(t=>t.DeviceId).map(t=>[t.DeviceId.toLowerCase(),t]));
const reportRows = payload.report.map(({device:d,previous,recentUsers}) => {
  const t=byId.get(d.DeviceId.toLowerCase());
  if (!t) throw new Error('Device is missing from the register: '+d.DeviceId);
  return [d.DeviceName,d.OverdueKEVs,previous,null,d.KEVCount,recentUsers||'',t.RemediationStatus,t.ServiceNowTickets,t.NeedsEscalation,t.EscalationReason,t.Circumstances,t.Notes,t.LastReviewed,null,d.ExposureLevel||'',d.OSPlatform||'',t.IdentityMatch,t.SFSUTag||'',t.ATTag||'',t.SerialNo||'',d.DeviceId];
});
section(report,'Overdue KEVs by device',`${payload.reportDate}. Previous: ${payload.previousDate || 'unavailable'}. Prior legacy counts use unique full-name matches; later runs use DeviceId.`,reportHeaders,reportRows,'WeeklyReport');
report.tabColor='#243E55';
report.getRange('A5:B5').values=[['Devices with overdue KEVs',null]];
report.getRange('D5:E5').values=[['Overdue pairs',null]];
report.getRange('G5:H5').values=[['Current export devices',payload.summary.sourceDevices]];
report.getRange('B5').formulas=[[count ? `=COUNT(B${START}:B${last})` : '=0']];
report.getRange('E5').formulas=[[count ? `=SUM(B${START}:B${last})` : '=0']];
report.getRange('A5:H5').format.font={name:'Arial',size:10,bold:true,color:'#243E55'};
report.getRange('A5').format.wrapText=true;
report.getRange('D5').format.wrapText=true;
report.getRange('G5').format.wrapText=true;
report.getRange('A5:H5').format.rowHeight=34;
report.getRange('A6').values=[['Edit G:M and Q:T (amber headers). Pink rows are new or returning. Save and close before the next refresh.']];
report.getRange('A6').format.font={name:'Arial',size:10,italic:true,color:'#526779'};
for (const range of ['G8:M8','Q8:T8']) report.getRange(range).format.font.color='#FFE3A1';

section(tracking,'Device register','Internal register. Edit the Weekly report; saved edits are reconciled on refresh. Absent devices remain here.',payload.trackingHeaders,payload.tracking.map(t=>payload.trackingHeaders.map(h=>h==='LastObservedReportDate'?dateValue(t[h]):t[h]??'')),'DeviceTracking');
tracking.tabColor='#B78326';
tracking.getRange(`C${START}:C${trackLast}`).dataValidation={rule:{type:'list',values:['—','Ticket open','In progress','Resolved']}};
tracking.getRange(`D${START}:D${trackLast}`).dataValidation={rule:{type:'list',values:['Yes','No']}};
const identityValues=['Verified DeviceId','DeviceId from export','Unique name match; verify','No current match','Ambiguous name; review','Unverified DeviceId'];
tracking.getRange(`J${START}:J${trackLast}`).dataValidation={rule:{type:'list',values:identityValues}};
tracking.getRange(`I${START}:I${trackLast}`).setNumberFormat('mm/dd/yy');
tracking.getRange(`N${START}:N${trackLast}`).setNumberFormat('mm/dd/yy');

if (count) {
  report.getRange(`G${START}:M${last}`).format.fill='#FFF2CC';
  report.getRange(`Q${START}:T${last}`).format.fill='#FFF2CC';
  report.getRange(`R${START}:T${last}`).setNumberFormat('@');
  report.getRange(`G${START}:G${last}`).dataValidation={rule:{type:'list',values:['—','Ticket open','In progress','Resolved']}};
  report.getRange(`I${START}:I${last}`).dataValidation={rule:{type:'list',values:['Yes','No']}};
  report.getRange(`Q${START}:Q${last}`).dataValidation={rule:{type:'list',values:identityValues}};
  report.getRange(`M${START}:M${last}`).setNumberFormat('mm/dd/yy');
  const formulas=payload.report.map((entry,i)=>{
    const r=START+i, shared=payload.records.filter(x=>x.DeviceName.trim().toLowerCase()===entry.device.DeviceName.trim().toLowerCase()).length>1;
    return [`=IF(ISNUMBER(C${r}),B${r}-C${r},"")`,
            `=IF(AND(G${r}="Resolved",B${r}>0),"Resolved; overdue KEVs detected",IF(AND(G${r}="Ticket open",H${r}=""),"Ticket number needed",IF(AND(I${r}="Yes",J${r}=""),"Escalation reason needed",IF(OR(Q${r}="Unique name match; verify",Q${r}="Unverified DeviceId"),"Verify migrated identity",${shared ? '"Shared hostname; use DeviceId"' : `IF(I${r}="Yes","Needs escalation","")` }))))`];
  });
  report.getRange(`D${START}:D${last}`).formulas=formulas.map(x=>[x[0]]);
  report.getRange(`N${START}:N${last}`).formulas=formulas.map(x=>[x[1]]);
  report.getRange(`B${START}:E${last}`).setNumberFormat('#,##0');
  report.getRange(`D${START}:D${last}`).setNumberFormat('+#,##0;-#,##0;0');
  report.getRange(`F${START}:F${last}`).format.wrapText=true;
  report.getRange(`J${START}:L${last}`).format.wrapText=true;
  payload.report.forEach((entry,index)=>{
    const t=byId.get(entry.device.DeviceId.toLowerCase());
    if (['New','Returning'].includes(entry.appearance)) report.getRange(`A${START+index}:U${START+index}`).format.fill='#F4B6D2';
    const lines=Math.max(1,Math.ceil((entry.recentUsers||'').length/65),Math.ceil((t.Notes||'').length/45),Math.ceil((t.EscalationReason||'').length/42));
    if (lines>1) report.getRange(`A${START+index}:U${START+index}`).format.rowHeight=lines*17+8;
  });
}

function statusFormatting(sheet,statusCol,escalationCol,end) {
  for (const [status,color] of [['Ticket open',colors.Yellow],['In progress',colors.Orange],['Resolved',colors.Green]])
    sheet.getRange(`${statusCol}${START}:${statusCol}${end}`).conditionalFormats.addCustom(`$${statusCol}${START}="${status}"`,{fill:color});
  sheet.getRange(`${escalationCol}${START}:${escalationCol}${end}`).conditionalFormats.addCustom(`$${escalationCol}${START}="Yes"`,{fill:colors.Blue,font:{bold:true}});
}
if (count) {
  statusFormatting(report,'G','I',last);
  report.getRange(`A${START}:A${last}`).conditionalFormats.addCustom(`$I${START}="Yes"`,{fill:colors.Blue});
  report.getRange(`N${START}:N${last}`).conditionalFormats.addCustom(`$G${START}="Resolved"`,{fill:'#FFE2DE',font:{color:'#A12718',bold:true}});
}
statusFormatting(tracking,'C','D',trackLast);

const sourceHeaders=Object.keys(payload.records[0]);
const sourceLast=section(source,'Defender export',payload.legacyAlertLabels ? 'Original export preserved. IncidentsLast6Months counts alert IDs; six-month coverage is unverified. See Guide.' : 'Original export values. Alert counts describe retained Defender data, not ServiceNow tickets.',sourceHeaders,payload.records.map(r=>sourceHeaders.map(h=>r[h]??'')),'DefenderSource');
sourceHeaders.forEach((h,i)=>{
  if (['ReportAsOfUtc','ContextWindowStartUtc','DeviceInfoLastSeen'].includes(h)) source.getRange(`${column(i)}${START}:${column(i)}${sourceLast}`).setNumberFormat('mm/dd/yy hh:mm');
});
const queue=[...payload.reviews];
for (const [name,oldCount,newCount,delta,reason] of payload.reconciliation)
  queue.push(['Same-week reconciliation',name,'',`Legacy ${oldCount}; new ${newCount}; change ${delta>0?'+':''}${delta}. ${reason}.`]);
section(review,'Review queue','Snapshot at generation. Saved Weekly report edits are reflected in this queue at the next refresh.',['Review item','Device','DeviceId','Detail'],queue,'ReviewQueue');

const historyLast=section(history,'Device history','Saved state at each refresh. Filter Device or DeviceId. A missing export record is not evidence of remediation.',payload.historyHeaders,payload.history.map(row=>payload.historyHeaders.map(h=>['Report date','Captured UTC'].includes(h)?dateValue(row[h]):row[h]??'')),'DeviceHistory');
history.tabColor='#526779';
history.getRange(`A${START}:A${historyLast}`).setNumberFormat('mm/dd/yy');
history.getRange(`N${START}:N${historyLast}`).setNumberFormat('mm/dd/yy');
history.getRange(`Q${START}:Q${historyLast}`).setNumberFormat('mm/dd/yy hh:mm:ss');
for (const [sheet,headers,names,end] of [
  [tracking,payload.trackingHeaders,['SFSUTag','ATTag','SerialNo'],trackLast],
  [history,payload.historyHeaders,['SFSU Tag','AT Tag','Serial No.'],historyLast],
]) {
  for (const name of names) {
    const col=column(headers.indexOf(name));
    sheet.getRange(`${col}${START}:${col}${end}`).setNumberFormat('@');
    sheet.getRange(`${col}1:${col}${end}`).format.columnWidthPx=name.includes('Serial')?180:125;
  }
}

const guideRows=[
  ['Report date',payload.reportDate],
  ['Previous report date',payload.previousDate || 'Unavailable'],
  ['Defender source',payload.source],
  ['Tracking source',payload.trackingSource],
  ['Schema version',payload.schemaVersion],
  ['Editing model','Weekly report'],
  ['Snapshot id',payload.snapshotId],
  ['Generated UTC',payload.generatedUtc],
  ['Report scope','Current export records with OverdueKEVs > 0. Defender import keeps all KEV-bearing devices, including those not yet overdue. Devices with no KEVs are not in this source.'],
  ['Weekly step 1','Edit the columns with amber headers on Weekly report. Save and close the workbook. Select this saved workbook as the previous report during the next refresh.'],
  ['Weekly step 2','Export DevicesByKEV.kql from Defender as CSV. CSV and XLSX exports are accepted. The revised query has not been executed against the University tenant; validate it in Advanced Hunting first.'],
  ['Weekly step 3','Double-click Refresh-KEVReport.cmd. Choose the Defender CSV/XLSX export, previous XLSX report, report date, output folder and filename, then Create report. Command-line examples are in Refresh-KEVReport.ps1.'],
  ['Weekly step 4','Review the generated report and Review queue. Save a dated copy. No workbook, source file, or ticket is overwritten by the utility.'],
  ['Runtime','The refresh utility uses the bundled Node.js, Python and artifact-tool installed with Codex on this computer. No cloud connection or tenant credentials are used.'],
  ['Manual inputs','Edit status, tickets, escalation flag/reason, circumstances, notes, last-reviewed date, identity match, SFSU Tag, AT Tag and Serial No. on Weekly report. Clearing a field is saved as an intentional change. Last reviewed is a manual date.'],
  ['Asset identifiers','SFSU Tag, AT Tag and Serial No. are manual text fields. Blank means no value recorded. Values carry forward by DeviceId from saved weekly reports and are retained in History, including while devices are absent. They are never inferred from Defender or another external source.'],
  ['Device register','Tracking is a hidden internal register. It retains records for devices absent from the current report. Never delete this sheet. Weekly report edits are collected by DeviceId before a new export is merged.'],
  ['History','History records generated snapshots and saved edits observed during refresh. It captures saved state rather than every keystroke. Migrated baseline records reflect data available in the previous workbook; older unavailable states are not reconstructed.'],
  ['Sorting and imported values','Sort the complete WeeklyReport table normally. Keep each DeviceId attached to its row. Imported device facts are checked against Defender import during refresh; conflicting edits and duplicate IDs stop generation.'],
  ['Copies and conflicts','Use the most recent saved working report. A refresh reads one previous workbook. Changes in separate copies are not merged automatically. Disagreements between edited report values and the backing register are flagged.'],
  ['Archived input','The refresh utility archives the previous saved workbook under archive beside the CMD launcher before generating the next report. Existing outputs are never overwritten.'],
  ['Color meanings','Pink rows: new or returning to Weekly report compared with the previous saved Weekly report, including returns after zero overdue KEVs. Yellow: Ticket open. Orange: In progress. Green: recorded Resolved. Blue device/escalation cells: progress is difficult. Status and escalation colors take priority in their cells.'],
  ['Unset status','— means no remediation status selected. Old Unspecified values and blank status cells become — on refresh. A ticket reference alone does not prove that the ticket is open.'],
  ['Identity migration',`${payload.tracking.filter(t=>t.IdentityMatch==='Unique name match; verify').length} tracking records have candidate unique-name matches. Verify DeviceId, then choose Verified DeviceId in the Weekly report Identity match column.`],
  ['Unmatched identities','Unmatched legacy tracking rows retain a blank DeviceId. A later unique name match is a candidate requiring review. Existing populated DeviceIds are never reassigned by hostname.'],
  ['Resolving an unmatched identity','After verifying the device in Defender, enter its DeviceId and mark Verified DeviceId. If an unannotated row already has that ID, consolidate the tracking fields into that row and remove the duplicate row. Refresh refuses duplicate DeviceIds.'],
  ['Resolved status','A recorded resolution is retained. Any overdue KEVs still detected create a review flag. No resolution date or ServiceNow closure is invented.'],
  ['Missing records','Absence from an export does not establish remediation. Tracking is retained, including for devices missing from more than one report.'],
  ['Counts and comparisons','The initial Previous overdue column uses Nov 2 legacy counts only where the current full name is unique. Blank means unavailable or ambiguous. Later runs compare by DeviceId against the previous Defender import, including legitimate zero counts.'],
  ['Historic comparison scope','Legacy counts and current device/CVE counts have different grouping. Initial deltas require interpretation and are not proof of patching. The old query itself was not provided.'],
  ['Manual ticket field','The legacy value Manual is preserved as a note, not treated as a ServiceNow incident number.'],
  ['Ticket scope','ServiceNowTickets stores references; multiple references can be separated by semicolons. Device/CVE ticket coverage and live ticket state are not established by this pilot.'],
  ['User context','RecentUsers records recent logons, not confirmed device ownership. The revised KQL excludes common Windows pseudo-accounts visible in the supplied export.'],
  ['Recent users display','Windows domain labels with a resolved nine-digit SFSU ID display as ID (Display Name). Macs and local accounts display usernames only. Multiple users are comma-separated. Source JSON punctuation and common system pseudo-accounts are omitted; original RecentUsers remains in Defender import.'],
  ['Exposure',"Imported DeviceInfo.ExposureLevel is Defender's assessment of vulnerability to exploitation across the device, based on its exposure score. Low: 0-29. Medium: 30-69. High: 70-100. This is broader than KEV count and does not establish that the device has been compromised."],
  ['Exposure scoring','The calculation uses vulnerability severity and exploitability. The updated model also includes internet-facing status and asset criticality. Microsoft is rolling out the model across tenants; the raw score is not included in this export.'],
  ['Alert context','The original IncidentsLast6Months and IncidentNames columns actually count AlertId values and alert titles. They do not describe ServiceNow tickets or unique Defender incidents. The revised KQL uses AlertsLast30Days and AlertTitles.'],
  ['Alert retention','An ago(180d) filter cannot extend retained hunting data. Standard Defender hunting event data is available for up to 30 days unless extended retention is configured. A zero means no matching alert observed in the queried data.'],
  ['Due date boundary','The revised KQL preserves dueDate < export time in UTC. Confirm whether University policy instead treats an item as overdue only after the entire CISA due date has passed.'],
  ['Refresh behavior','Manual values are edited directly on Weekly report; colors and the row Review formula update immediately. Tracking, History, imported values, comparisons and the Review queue update when the utility processes the saved workbook.'],
  ['Automatic identity checks','The refresh fails on duplicate source/Tracking DeviceIds, missing required counts, invalid counts, and malformed CVE arrays. New source DeviceIds get new blank tracking entries.'],
  ['Future integrations','After validating the pilot, replace manual export and ticket-state updates with Microsoft Graph Security and ServiceNow reads. Native ServiceNow Vulnerability Response availability is unconfirmed.'],
  ['Reference: alert schema','https://learn.microsoft.com/en-us/defender-xdr/advanced-hunting-alertevidence-table'],
  ['Reference: hunting retention','https://learn.microsoft.com/en-us/defender-xdr/advanced-hunting-overview'],
  ['Reference: small-set dcount','https://learn.microsoft.com/en-us/kusto/query/dcount-aggregation-function'],
  ['Reference: exposure score','https://learn.microsoft.com/en-us/defender-vulnerability-management/tvm-exposure-score'],
];
section(guide,'Using the weekly report',`Source: ${payload.source}. Report date: ${payload.reportDate}.`,['Item','Details'],guideRows,'ReportGuide');
guide.getRange(`A${START}:A${START+guideRows.length-1}`).format.font={name:'Arial',size:10,bold:true,color:'#243E55'};
guide.getRange(`B${START}:B${START+guideRows.length-1}`).format.wrapText=true;
guide.getRange(`A${START}:B${START+guideRows.length-1}`).format.rowHeight=52;
guide.getRange(`B${START}:B${START+guideRows.length-1}`).format.verticalAlignment='top';
guide.getRange(`B${START+guideRows.findIndex(row=>row[0]==='Generated UTC')}`).setNumberFormat('mm/dd/yy hh:mm:ss');

// Widths reserve the first screen for actionable report fields.
for (const [col,pixels] of Object.entries({A:290,B:80,C:90,D:65,E:65,F:460,G:125,H:140,I:95,J:280,K:145,L:300,M:115,N:245,O:85,P:105,Q:210,R:125,S:125,T:180,U:320})) report.getRange(`${col}1:${col}${Math.max(last,9)}`).format.columnWidthPx=pixels;
for (const [col,pixels] of Object.entries({A:320,B:290,C:125,D:135,E:260,F:145,G:140,H:270,I:105,J:200,K:95,L:450})) tracking.getRange(`${col}1:${col}${trackLast}`).format.columnWidthPx=pixels;
source.getRange(`A1:A${sourceLast}`).format.columnWidthPx=320;
source.getRange(`B1:B${sourceLast}`).format.columnWidthPx=290;
for (const h of ['KEVs','OverdueCves','RecentUsers','IncidentNames','AlertTitles']) {
  const index=sourceHeaders.indexOf(h);
  if (index>=0) source.getRange(`${column(index)}1:${column(index)}${sourceLast}`).format.columnWidthPx=460;
}
for (const [col,pixels] of Object.entries({A:210,B:290,C:320,D:780})) review.getRange(`${col}1:${col}${Math.max(START+queue.length-1,9)}`).format.columnWidthPx=pixels;
if (queue.length) {
  review.getRange(`D${START}:D${START+queue.length-1}`).format.wrapText=true;
  review.getRange(`A${START}:D${START+queue.length-1}`).format.rowHeight=40;
}
for (const [col,pixels] of Object.entries({A:110,B:290,C:125,D:145,E:90,F:80,G:100,H:135,I:260,J:300,K:95,L:280,M:150,N:115,O:320,P:300,Q:180,R:210,S:460,T:460,U:320})) history.getRange(`${col}1:${col}${historyLast}`).format.columnWidthPx=pixels;
history.getRange(`I${START}:M${historyLast}`).format.wrapText=true;
history.getRange(`Y1:Y${historyLast}`).format.columnWidthPx=145;
history.getRange(`A${START}:Y${historyLast}`).format.rowHeight=38;
payload.history.forEach((row,index)=>{
  const lines=Math.max(1,Math.ceil((row['Changed fields']||'').length/38),Math.ceil((row.Notes||'').length/45),Math.ceil((row['Escalation reason']||'').length/42));
  if(lines>2) history.getRange(`A${START+index}:Y${START+index}`).format.rowHeight=lines*17+8;
});
report.freezePanes.freezeColumns(1);
history.freezePanes.freezeColumns(2);
guide.getRange(`A1:A${START+guideRows.length}`).format.columnWidthPx=235;
guide.getRange(`B1:B${START+guideRows.length}`).format.columnWidthPx=990;

wb.recalculate();
const expected = {devices:count,overdue:payload.summary.overdue};
const actual = {devices:report.getRange('B5').values[0][0],overdue:report.getRange('E5').values[0][0]};
if (JSON.stringify(actual)!==JSON.stringify(expected)) throw new Error(`Summary mismatch ${JSON.stringify({actual,expected})}`);
if (count) {
  const changes=report.getRange(`D${START}:D${last}`).values;
  payload.report.forEach((entry,index)=>{
    const expectedChange=entry.previous===null || entry.previous===undefined ? '' : entry.device.OverdueKEVs-entry.previous;
    if (changes[index][0]!==expectedChange) throw new Error('Previous-count comparison failed, including zero/blank handling.');
  });
}
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},maxChars:4000,summary:'Formula error scan'});
const inspect=await wb.inspect({kind:'table',range:'Weekly report!A8:J15',include:'values,formulas',tableMaxRows:8,tableMaxCols:10,maxChars:6000});
const outDir=path.dirname(outputPath);
await fs.mkdir(outDir,{recursive:true});
const validationDir=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..','validation');
await fs.mkdir(validationDir,{recursive:true});
await fs.writeFile(path.join(validationDir,'formula-scan.ndjson'),errors.ndjson);
await fs.writeFile(path.join(validationDir,'report-inspect.ndjson'),inspect.ndjson);
// Render a bounded preview so long source/notes columns do not create a huge image.
if (process.env.KEV_SKIP_PREVIEW!=='1') {
  for (const [sheetName,range,file] of [['Weekly report','A1:J17','weekly-report.png'],['Weekly report','N8:U17','weekly-assets.png'],['History','A1:J15','history.png'],['History','V8:Y15','history-assets.png'],['Tracking','A1:J15','tracking.png'],['Defender import','A1:J13','defender-import.png'],['Review','A1:D15','review.png'],['Guide','A1:B18','guide.png']]) {
    const preview=await wb.render({sheetName,range,scale:1,format:'png'});
    await fs.writeFile(path.join(validationDir,file),new Uint8Array(await preview.arrayBuffer()));
  }
}
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
// Keep the author's diagnostic sidecar out of the report folder when the runtime emits one.
await relocateExportInspection(outputPath+'.inspect.ndjson',path.join(validationDir,'latest-export-inspect.ndjson'));
console.log(JSON.stringify({output:outputPath,...actual,trackingRows:payload.tracking.length,reviewRows:queue.length}));
