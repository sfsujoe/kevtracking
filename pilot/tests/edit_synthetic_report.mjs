// Saved manual edits for the portable refresh regression; synthetic devices only.
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const [input,output,mode]=process.argv.slice(2);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const sheet=wb.worksheets.getItem('Weekly report');
const rows=sheet.getUsedRange().values, headers=rows[7];
for (let i=8;i<rows.length;i++) {
  const id=rows[i][headers.indexOf('DeviceId')];
  if (!['steady','zero','absent'].includes(id)) continue;
  const set=(name,value)=>{sheet.getCell(i,headers.indexOf(name)).values=[[value]];};
  if (mode==='clear') {
    if (id==='steady') {
      set('SFSU Tag',''); set('Remediation status','');
    }
  } else {
    set('SFSU Tag','000123'); set('AT Tag','AT-'+id); set('Serial No.','0000'+id);
    if (id==='steady') {
      set('Remediation status','Ticket open'); set('ServiceNow tickets','INC000123');
      set('Escalation','Yes'); set('Escalation reason','Synthetic maintenance window');
    }
  }
}
wb.recalculate();
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
