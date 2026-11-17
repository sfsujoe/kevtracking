"""End-to-end synthetic refresh through a copy containing only CMD and scripts.

Run explicitly: python -B -m unittest discover -s pilot/tests -p test_portable.py -v
Requires the bundled runtime. Keeps diagnostics under pilot/validation.
"""
import csv
import os
from pathlib import Path
import shutil
import subprocess
import sys
import unittest
import uuid
import zipfile
import xml.etree.ElementTree as ET

sys.dont_write_bytecode = True
PILOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PILOT/'scripts'))
import prepare_report as p


class PortableReportTests(unittest.TestCase):
    def test_three_weeks_assets_appearance_upgrade_and_clears(self):
        root=PILOT/'validation'/('portable & apostrophe\'s '+uuid.uuid4().hex)
        root.mkdir(parents=True)
        deploy=root/'production copy'
        deploy.mkdir()
        shutil.copy2(PILOT/'Refresh-KEVReport.cmd',deploy)
        shutil.copytree(PILOT/'scripts',deploy/'scripts')
        self.assertEqual({f.name for f in deploy.iterdir()},{'Refresh-KEVReport.cmd','scripts'})
        runtime=Path(os.environ['USERPROFILE'])/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
        powershell=Path(os.environ['WINDIR'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
        node=runtime/'node/bin/node.exe'
        env=dict(os.environ,KEV_SKIP_PREVIEW='1',PYTHONDONTWRITEBYTECODE='1')

        def run(args):
            result=subprocess.run([str(a) for a in args],cwd=PILOT.parent,env=env,
                                  capture_output=True,text=True,timeout=180)
            self.assertEqual(result.returncode,0,result.stdout+'\n'+result.stderr)

        def export(name,ids,with_counts=True):
            target=root/(name+'.csv')
            rows=[dict(DeviceId=key,DeviceName='host-'+key,KEVCount=1,OverdueKEVs=overdue,
                       KEVs='["CVE-2026-0001"]',OverdueCves='["CVE-2026-0001"]' if overdue else '[]',
                       SFSUTag='DO NOT IMPORT',ATTag='DO NOT IMPORT',SerialNo='DO NOT IMPORT')
                  for key,overdue in ids]
            for row in rows:
                if row['DeviceId'] in ('vuln-only','clean'):
                    row.update(KEVCount=0,KEVs='[]')
                if with_counts:
                    counts=(0,0,0) if row['DeviceId']=='clean' else (12,3,4) if name=='week3' else (8,2,3)
                    row.update(zip(p.VULNERABILITY_COLUMNS,counts))
            with target.open('w',newline='',encoding='utf-8-sig') as stream:
                writer=csv.DictWriter(stream,fieldnames=list(rows[0]));writer.writeheader();writer.writerows(rows)
            return target

        def refresh(source,date,previous='',upgrade=False):
            output=root/('KEV-Report-'+date+('-upgrade' if upgrade else '')+'.xlsx')
            env.update(KEV_TEST_CMD=str(deploy/'Refresh-KEVReport.cmd'),KEV_TEST_EXPORT=str(source),
                       KEV_TEST_DATE=date,KEV_TEST_OUTPUT=str(output),KEV_TEST_PREVIOUS=str(previous))
            command=('& $env:KEV_TEST_CMD -UpgradeReport -TrackingWorkbook $env:KEV_TEST_PREVIOUS -OutputPath $env:KEV_TEST_OUTPUT'
                     if upgrade else '& $env:KEV_TEST_CMD -ExportPath $env:KEV_TEST_EXPORT -ReportDate $env:KEV_TEST_DATE -OutputPath $env:KEV_TEST_OUTPUT'
                     + (' -TrackingWorkbook $env:KEV_TEST_PREVIOUS' if previous else ''))
            run([powershell,'-NoLogo','-NoProfile','-NonInteractive','-Command',command+'; exit $LASTEXITCODE'])
            return output

        def weekly(book):
            return {r['DeviceId']:r for r in p.table_records(book['Weekly report'],'Device')}

        def appearance(book):
            guide={r['Item']:r['Details'] for r in p.table_records(book['Guide'],'Item')}
            return {h['DeviceId']:h['Weekly appearance'] for h in p.table_records(book['History'],'Report date')
                    if h['SnapshotId']==guide['Snapshot id'] and h['Event']=='Generated'}

        # Reproduce Box's EXDEV response through the real builder and launcher.
        preload=root/'simulate-cross-device.mjs'
        preload.write_text("""import fs from 'node:fs/promises';
const rename = fs.rename;
fs.rename = async (source, destination) => {
  if (String(source).endsWith('.inspect.ndjson')) {
    await fs.writeFile(new URL('./cross-device-exercised', import.meta.url), 'EXDEV');
    throw Object.assign(new Error('Simulated cross-device rename'), {code: 'EXDEV'});
  }
  return rename(source, destination);
};
""",encoding='utf-8')
        original_node_options=env.get('NODE_OPTIONS')
        env['NODE_OPTIONS']=((original_node_options or '')+' --import='+preload.as_uri()).strip()
        first=refresh(export('week1',[('steady',1),('zero',1),('absent',1)],with_counts=False),'2026-11-02')
        if original_node_options is None:
            env.pop('NODE_OPTIONS')
        else:
            env['NODE_OPTIONS']=original_node_options
        self.assertTrue((root/'cross-device-exercised').is_file())
        self.assertFalse(Path(str(first)+'.inspect.ndjson').exists())
        self.assertTrue((deploy/'validation/latest-export-inspect.ndjson').is_file())
        with zipfile.ZipFile(first) as archive:
            self.assertTrue(any(s.get('name')=='Tracking' and s.get('state')=='hidden'
                                for s in ET.fromstring(archive.read('xl/workbook.xml')).find('s:sheets',p.NS)))
        book=p.read_xlsx(first)
        first_snapshot=next(r['Details'] for r in p.table_records(book['Guide'],'Item') if r['Item']=='Snapshot id')
        self.assertTrue(all(r['Remediation status']=='—' for r in weekly(book).values()))
        self.assertTrue(all(r[h]=='' for r in weekly(book).values() for h in ['SFSU Tag','AT Tag','Serial No.']))
        self.assertTrue(all(v=='New' for v in appearance(book).values()))
        self.assertTrue(all(row[c]=='' for row in weekly(book).values() for c in p.VULNERABILITY_COLUMNS.values()))
        edited=root/'edited-first.xlsx'
        run([node,PILOT/'tests/edit_synthetic_report.mjs',first,edited,'set'])
        second=refresh(export('week2',[('steady',1),('zero',0)]),'2026-11-09',edited)
        book=p.read_xlsx(second)
        self.assertEqual(set(weekly(book)),{'steady','zero'})
        self.assertEqual([weekly(book)['zero'][c] for c in p.VULNERABILITY_COLUMNS.values()],[8,2,3])
        self.assertEqual(weekly(book)['zero']['Overdue KEVs'],0)
        self.assertEqual(weekly(book)['steady']['SFSU Tag'],'000123')
        self.assertEqual(appearance(book)['steady'],'')
        history=p.table_records(book['History'],'Report date')
        self.assertTrue(any(h['DeviceId']=='absent' and h['SFSU Tag']=='000123' and not h['Observed in export'] for h in history))
        first_history=[h for h in history if h['SnapshotId']==first_snapshot and h['Event']=='Generated']
        self.assertEqual(len(first_history),3)
        self.assertTrue(all(h['Discovered vulnerabilities']=='' for h in first_history))
        env.pop('KEV_SKIP_PREVIEW')
        third=refresh(export('week3',[('steady',1),('zero',1),('absent',1),('fresh',1),('vuln-only',0),('clean',0)]),'2026-11-16',second)
        env['KEV_SKIP_PREVIEW']='1'
        book=p.read_xlsx(third); rows=weekly(book)
        headers=next(r['values'] for r in book['Weekly report'] if r['values'].get('A')=='Device')
        self.assertEqual([headers[c] for c in ('F','G','H')],list(p.VULNERABILITY_COLUMNS.values()))
        self.assertEqual([headers[c] for c in ('U','V','W','X')],['SFSU Tag','AT Tag','Serial No.','DeviceId'])
        self.assertEqual(appearance(book),{'steady':'','zero':'','absent':'Returning','fresh':'New','vuln-only':'New','clean':'New'})
        self.assertEqual([rows['steady'][c] for c in p.VULNERABILITY_COLUMNS.values()],[12,3,4])
        self.assertEqual([rows['clean'][c] for c in p.VULNERABILITY_COLUMNS.values()],[0,0,0])
        self.assertEqual(rows['vuln-only']['All KEVs'],0)
        self.assertEqual(rows['vuln-only']['Discovered vulnerabilities'],12)
        summary=next(r['values'] for r in book['Weekly report'] if r['row']==5)
        self.assertEqual([summary[c] for c in ('B','E','H')],[6,4,4])
        self.assertEqual([rows['zero'][h] for h in ['SFSU Tag','AT Tag','Serial No.']],['000123','AT-zero','0000zero'])
        self.assertEqual(rows['absent']['SFSU Tag'],'000123')
        self.assertEqual(rows['fresh']['SFSU Tag'],'')
        # Check the saved Excel types/formats, not just the rendered preview.
        with zipfile.ZipFile(third) as archive:
            sheet=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            style_xml=ET.fromstring(archive.read('xl/styles.xml'))
            styles=style_xml.find('s:cellXfs',p.NS)
            cell=sheet.find('.//s:c[@r="U9"]',p.NS)
            self.assertIn(cell.get('t'),('s','inlineStr','str'))
            format_id=styles[int(cell.get('s'))].get('numFmtId')
            formats={'49':'@',**{f.get('numFmtId'):f.get('formatCode') for f in style_xml.findall('s:numFmts/s:numFmt',p.NS)}}
            self.assertEqual(formats[format_id],'@')
            validation=sheet.find('s:dataValidations',p.NS)
            self.assertTrue(any('—' in ''.join(v.itertext()) and 'Unspecified' not in ''.join(v.itertext()) for v in validation))
            self.assertIsNotNone(sheet.find('s:conditionalFormatting',p.NS))
            self.assertTrue(any(s.get('name')=='Tracking' and s.get('state')=='hidden'
                                for s in ET.fromstring(archive.read('xl/workbook.xml')).find('s:sheets',p.NS)))
        for row in book['Weekly report']:
            if row['row']>8 and row['values'].get('X') in ('absent','fresh','vuln-only','clean'):
                self.assertTrue(row['colors']['B'].endswith('F4B6D2'))
        upgraded=refresh('', '2026-11-16',third,upgrade=True)
        self.assertEqual(appearance(p.read_xlsx(upgraded)),appearance(book))
        self.assertEqual(weekly(p.read_xlsx(upgraded)),rows)
        cleared=root/'cleared.xlsx'
        run([node,PILOT/'tests/edit_synthetic_report.mjs',third,cleared,'clear'])
        final=refresh(root/'week3.csv','2026-11-23',cleared)
        book=p.read_xlsx(final); row=weekly(book)['steady']
        self.assertEqual(row['Remediation status'],'—')
        self.assertEqual(row['SFSU Tag'],'')
        self.assertEqual(row['AT Tag'],'AT-steady')
        self.assertTrue(any(h['Event']=='Saved edits' and h['DeviceId']=='steady' and h['SFSU Tag']==''
                            for h in p.table_records(book['History'],'Report date')))
        self.assertFalse((deploy/'scripts/__pycache__').exists())
        self.assertTrue((deploy/'archive').is_dir())
        self.assertTrue((deploy/'validation/weekly-assets.png').is_file())
        self.assertEqual(len(list((deploy/'archive').glob('*.xlsx'))),4)
        print('Synthetic deployment and previews: '+str(deploy))


if __name__=='__main__':
    unittest.main(verbosity=2)
