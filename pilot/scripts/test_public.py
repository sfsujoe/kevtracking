"""Self-contained regression checks with synthetic CSV and device data."""
import copy
import csv
import json
from pathlib import Path
import tempfile
import unittest

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
        self.assertEqual(result['tracking'][0]['RemediationStatus'],'Unspecified')
        self.assertEqual(result['history'][0]['Event'],'Generated')


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
