"""Self-contained regression checks with synthetic CSV and device data."""
import copy
import csv
import json
from pathlib import Path
import tempfile
import unittest
import sys
from unittest.mock import patch
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import prepare_report as p


class PublicImportTests(unittest.TestCase):
    def setUp(self):
        validation=Path(__file__).resolve().parents[1]/'validation'
        validation.mkdir(exist_ok=True)
        temp=tempfile.TemporaryDirectory(prefix='public-tests-',dir=validation)
        self.addCleanup(temp.cleanup)
        self.root=Path(temp.name)
        self.row=dict(DeviceId='000abc',DeviceName='example-device',KEVCount=2,OverdueKEVs=1,
                      KEVs='["CVE-2026-0001","CVE-2026-0002"]',OverdueCves='["CVE-2026-0001"]',
                      IsInternetFacing='false',RecentUsers=json.dumps(['912345678 — Renée Example'],ensure_ascii=False),
                      Description='One, "quoted" line\r\nA second line',AssetTag='00007')

    def write(self,rows,encoding='utf-8-sig',headers=None):
        target=self.root/'example export & users.csv'
        headers=headers or list(rows[0])
        with target.open('w',encoding=encoding,newline='') as stream:
            writer=csv.writer(stream);writer.writerow(headers)
            writer.writerows([[r[h.strip()] for h in headers] for r in rows])
        return target

    def test_encodings_quotes_unicode_and_identifiers(self):
        for encoding in ('utf-8','utf-8-sig','utf-16'):
            with self.subTest(encoding=encoding):
                row=p.read_export(self.write([self.row],encoding))[0]
                self.assertEqual(row['DeviceId'],'000abc')
                self.assertEqual(row['AssetTag'],'00007')
                self.assertIs(row['IsInternetFacing'],False)
                self.assertEqual(row['Description'],self.row['Description'])
                self.assertEqual(row['RecentUsers'],self.row['RecentUsers'])
                self.assertEqual(row['OverdueKEVs'],1)

    def test_reordered_headers(self):
        headers=[' '+h+' ' for h in reversed(self.row)]
        self.assertEqual(p.read_export(self.write([self.row],headers=headers))[0]['KEVCount'],2)

    def test_empty_and_malformed_csv_fail(self):
        for data in ('','DeviceId,DeviceId\n','DeviceId,DeviceName,KEVCount,OverdueKEVs,KEVs\n',
                     'DeviceId,DeviceName,KEVCount,OverdueKEVs,KEVs\nid,name,1,1\n',
                     'DeviceId,DeviceName,KEVCount,OverdueKEVs,KEVs\n"unterminated\n'):
            with self.subTest(data=data):
                target=self.root/'bad.csv';target.write_text(data,encoding='utf-8')
                with self.assertRaises(ValueError):p.read_export(target)

    def test_invalid_counts_arrays_and_duplicate_ids_fail(self):
        for key,value in [('KEVCount',''),('OverdueKEVs','-1'),('OverdueKEVs','1.5'),
                          ('IsInternetFacing','maybe'),('KEVs','invalid JSON'),
                          ('OverdueCves','["CVE-2026-0001","CVE-2026-0001"]')]:
            with self.subTest(key=key,value=value):
                row=dict(self.row);row[key]=value
                with self.assertRaises(ValueError):p.read_export(self.write([row]))
        with self.assertRaisesRegex(ValueError,'Duplicate DeviceId'):
            p.read_export(self.write([self.row,self.row]))

    def test_large_quoted_field(self):
        row=dict(self.row);row['AlertTitles']=json.dumps(['x'*150000])
        self.assertEqual(p.read_export(self.write([row]))[0]['AlertTitles'],row['AlertTitles'])

    def test_first_report_needs_no_private_workbook(self):
        result=p.prepare(self.write([self.row]),'2026-11-16')
        self.assertEqual(result['summary']['reportDevices'],1)
        self.assertEqual(result['summary']['overdue'],1)
        self.assertEqual(result['tracking'][0]['RemediationStatus'],'—')
        self.assertEqual(result['history'][0]['Event'],'Generated')

    def test_vulnerability_counts_are_numeric_and_old_exports_remain_unknown(self):
        self.assertNotIn('DiscoveredVulnerabilities', p.read_export(self.write([self.row]))[0])
        row=dict(self.row, DiscoveredVulnerabilities=10, CriticalVulnerabilities=2, HighVulnerabilities=3)
        actual=p.read_export(self.write([row]))[0]
        self.assertEqual([actual[f] for f in p.VULNERABILITY_COLUMNS], [10,2,3])
        blank=dict(self.row, **dict.fromkeys(p.VULNERABILITY_COLUMNS, ''))
        self.assertTrue(all(p.read_export(self.write([blank]))[0][f]=='' for f in p.VULNERABILITY_COLUMNS))

    def test_invalid_or_partial_vulnerability_counts_fail(self):
        row=dict(self.row, DiscoveredVulnerabilities=10, CriticalVulnerabilities=2, HighVulnerabilities=3)
        for field,value in [('DiscoveredVulnerabilities',-1),('CriticalVulnerabilities',1.5),
                            ('HighVulnerabilities','invalid'),('HighVulnerabilities',''),
                            ('DiscoveredVulnerabilities',4)]:
            with self.subTest(field=field,value=value):
                with self.assertRaises(ValueError):
                    p.read_export(self.write([dict(row, **{field:value})]))
        partial=dict(self.row, DiscoveredVulnerabilities=10)
        with self.assertRaisesRegex(ValueError,'all three'):
            p.read_export(self.write([partial]))
        too_few=dict(row, DiscoveredVulnerabilities=1, CriticalVulnerabilities=0, HighVulnerabilities=0)
        with self.assertRaisesRegex(ValueError,'KEVCount exceeds'):
            p.read_export(self.write([too_few]))

    def test_xlsx_vulnerability_counts_use_the_same_validation(self):
        row=dict(self.row, KEVCount=2, OverdueKEVs=1,
                 DiscoveredVulnerabilities=10, CriticalVulnerabilities=True, HighVulnerabilities=3)
        book={'Export':[{'values':{'A':'DeviceId'}}]}
        with patch.object(p,'read_xlsx',return_value=book), patch.object(p,'table_records',return_value=[row]):
            with self.assertRaisesRegex(ValueError,'nonnegative integer'):
                p.read_export('synthetic.xlsx')

    def test_weekly_report_includes_non_kev_and_zero_count_devices(self):
        rows=[dict(self.row, DiscoveredVulnerabilities=10, CriticalVulnerabilities=2, HighVulnerabilities=3),
              dict(self.row, DeviceId='non-kev', KEVCount=0, OverdueKEVs=0, KEVs='[]', OverdueCves='[]',
                   DiscoveredVulnerabilities=5, CriticalVulnerabilities=1, HighVulnerabilities=2),
              dict(self.row, DeviceId='zero', KEVCount=0, OverdueKEVs=0, KEVs='[]', OverdueCves='[]',
                   DiscoveredVulnerabilities=0, CriticalVulnerabilities=0, HighVulnerabilities=0)]
        result=p.prepare(self.write(rows),'2026-11-16')
        self.assertEqual({r['device']['DeviceId'] for r in result['report']}, {'000abc','non-kev','zero'})
        self.assertEqual(result['summary']['reportDevices'],3)
        self.assertEqual(result['summary']['overdue'],1)
        self.assertEqual(result['schemaVersion'],4)


class UserLabelTests(unittest.TestCase):
    def test_domain_and_local_user_formats(self):
        users=['912345678 — Example User','macusername','PC\\localuser — Local Display']
        self.assertEqual(p.format_recent_users(users,'Windows11'),'912345678 (Example User), macusername, localuser')
        self.assertEqual(p.format_recent_users(['macusername — Mac Display'],'macOS'),'macusername')

    def test_unicode_punctuation_and_pseudo_accounts(self):
        self.assertEqual(p.format_recent_users(['912345678 — Renée Example',"923456789 — Anne-Marie O'Example"],'Windows11'),
                         "912345678 (Renée Example), 923456789 (Anne-Marie O'Example)")
        self.assertEqual(p.format_recent_users(['dwm-1','umfd-0','SYSTEM','HOST$','localuser','localuser']),'localuser')

    def test_empty_already_formatted_and_invalid_values(self):
        for value in ('',None,'[]','null'):self.assertEqual(p.format_recent_users(value),'')
        self.assertEqual(p.format_recent_users('912345678 (Example User)'),'912345678 (Example User)')
        for value in ('[broken','[123]',{'unexpected':'object'}):
            with self.assertRaises(ValueError):p.format_recent_users(value)


if __name__=='__main__':
    unittest.main(verbosity=2)
