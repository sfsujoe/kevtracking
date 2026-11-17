"""Read-only verification of requested report edits and preserved snapshot data."""
import json
import re
import sys
sys.dont_write_bytecode = True
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import prepare_report as p

original_path, updated_path = sys.argv[1:3]
original, updated = p.read_xlsx(original_path), p.read_xlsx(updated_path)
assert set(updated) == {'Weekly report','History','Tracking','Defender import','Review','Guide'}
for sheet, first in [('Defender import','DeviceId'),('Review','Review item')]:
    assert p.table_records(original[sheet],first) == p.table_records(updated[sheet],first), f'{sheet}: unrelated data changed'
before_tracking=p.table_records(original['Tracking'],'DeviceId')
after_tracking=p.table_records(updated['Tracking'],'DeviceId')
for old,new in zip(before_tracking,after_tracking):
    for field,value in old.items():
        assert new[field]==value,('Tracking',field)
before=p.table_records(original['Weekly report'],'Device')
after=p.table_records(updated['Weekly report'],'Device')
assert len(after)>=len(before)
source={r['DeviceId']:r for r in p.table_records(updated['Defender import'],'DeviceId')}
tracking={r['DeviceId']:r for r in p.table_records(updated['Tracking'],'DeviceId') if r['DeviceId']}
assert {r['DeviceId'] for r in after} == set(source)
after_by_id={r['DeviceId']:r for r in after}
for old in before:
    new=after_by_id[old['DeviceId']]
    for field in ('Device','DeviceId','Overdue KEVs','Previous overdue','Change','All KEVs','Exposure','OS'):
        assert old[field]==new[field],(field,old['DeviceId'])
    d=source[new['DeviceId']]
    assert new['Recent users']==p.format_recent_users(d.get('RecentUsers',''),d.get('OSPlatform',''))
    t=tracking[new['DeviceId']]
    for field,tracking_field in [('Remediation status','RemediationStatus'),('ServiceNow tickets','ServiceNowTickets'),('Escalation','NeedsEscalation'),('Circumstances','Circumstances')]:
        assert new[field]==t[tracking_field],(field,new['DeviceId'])
old_guide={r['Item']:r['Details'] for r in p.table_records(original['Guide'],'Item')}
new_guide={r['Item']:r['Details'] for r in p.table_records(updated['Guide'],'Item')}
for field in ('Report date','Previous report date','Defender source'):
    assert old_guide[field]==new_guide[field]
assert new_guide['Schema version']==4
assert new_guide['Editing model']=='Weekly report'
register,_,_,history,changes=p.load_saved_state(updated)
assert not changes
assert len(history)>=len(register)

with zipfile.ZipFile(updated_path) as z:
    styles=ET.fromstring(z.read('xl/styles.xml')).find('s:cellXfs',p.NS)
    rels={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    checked=0
    for sh in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets',p.NS):
        target=rels[sh.get(p.REL)]
        if sh.get('name')=='Tracking':
            assert sh.get('state')=='hidden'
        target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        xml=ET.fromstring(z.read(target))
        for row in xml.findall('s:sheetData/s:row',p.NS):
            if int(row.get('r'))<=8:
                continue
            for cell in row:
                assert cell.get('t')!='e',f"Formula error at {sh.get('name')}!{cell.get('r')}"
                align=styles[int(cell.get('s',0))].find('s:alignment',p.NS)
                if sh.get('name')!='Guide':
                    assert align is not None and align.get('horizontal')=='center' and align.get('vertical')=='center',(sh.get('name'),cell.get('r'))
                    checked+=1
                elif cell.get('r')=='B9':
                    assert align is None or align.get('horizontal')!='center','Guide alignment changed'
                if sh.get('name')=='Weekly report':
                    col=re.sub(r'\d','',cell.get('r'))
                    if col in ('J','K','L','M','N','O','P','T','U','V','W'):
                        assert cell.find('s:f',p.NS) is None,'Manual cell contains a lookup formula'
        if sh.get('name')=='Weekly report':
            assert xml.find('s:dataValidations',p.NS) is not None
            pane=xml.find('s:sheetViews/s:sheetView/s:pane',p.NS)
            assert pane is not None and float(pane.get('xSplit',0))==1 and float(pane.get('ySplit',0))==8
    assert z.testzip() is None
print(json.dumps({'verified':'Editable report values, hidden register, history, preserved data/comparisons, centered cells, validation and frozen panes',
                  'bodyCellsChecked':checked,'reportDevices':len(after),'overdue':sum(r['Overdue KEVs'] for r in after)}))
