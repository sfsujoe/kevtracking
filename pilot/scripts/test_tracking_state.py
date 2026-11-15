import copy
import unittest

import prepare_report as p
import tracking_state as s


class DirectEditingTests(unittest.TestCase):
    def setUp(self):
        self.records=[dict(DeviceId='one',DeviceName='host-one',KEVCount=2,OverdueKEVs=1,
                           KEVs='["CVE-1","CVE-2"]',OverdueCves='["CVE-1"]',ExposureLevel='High',OSPlatform='Windows11'),
                      dict(DeviceId='two',DeviceName='host-two',KEVCount=1,OverdueKEVs=1,
                           KEVs='["CVE-3"]',OverdueCves='["CVE-3"]',ExposureLevel='Medium',OSPlatform='macOS')]
        tracking=[]
        for d in self.records:
            t=dict.fromkeys(p.TRACK_HEADERS,'')
            t.update(DeviceId=d['DeviceId'],DeviceName=d['DeviceName'],RemediationStatus='Ticket open',
                     ServiceNowTickets='INC-'+d['DeviceId'],NeedsEscalation='No',Notes='original note',IdentityMatch='Verified DeviceId')
            tracking.append(t)
        self.payload=s.finalize_snapshot(dict(tracking=tracking,records=self.records,reportDate='2026-11-09'),[],captured='2026-11-09T20:00:00Z')
        self.register=self.payload['tracking']
        self.history=self.payload['history']
        self.guide={'Report date':'2026-11-09','Schema version':2,'Snapshot id':self.payload['snapshotId']}
        self.weekly=[]
        for d,t in zip(self.records,self.register):
            row=dict(Device=d['DeviceName'],DeviceId=d['DeviceId'],Exposure=d['ExposureLevel'],OS=d['OSPlatform'],**{'Overdue KEVs':d['OverdueKEVs'],'All KEVs':d['KEVCount']})
            row.update({column:t[field] for field,column in s.MANUAL_COLUMNS.items()})
            self.weekly.append(row)

    def collect(self,weekly=None,register=None,history=None,guide=None,formulas=None):
        return s.collect_saved_edits(register if register is not None else self.register,
            weekly if weekly is not None else self.weekly,self.records,
            history if history is not None else self.history,guide if guide is not None else self.guide,
            formulas,captured='2026-11-16T20:00:00Z')

    def test_sorted_report_edits_and_explicit_clears(self):
        weekly=copy.deepcopy(list(reversed(self.weekly)))
        edited=next(r for r in weekly if r['DeviceId']=='one')
        edited.update({'Remediation status':'In progress','ServiceNow tickets':'','Notes':'',
                       'Escalation':'Yes','Escalation reason':'Waiting for access','Last reviewed':'2026-11-15'})
        register,history,changes=self.collect(weekly)
        row=next(t for t in register if t['DeviceId']=='one')
        self.assertEqual(row['ServiceNowTickets'],'')
        self.assertEqual(row['Notes'],'')
        self.assertEqual(row['RemediationStatus'],'In progress')
        self.assertEqual(row['EscalationReason'],'Waiting for access')
        self.assertIsInstance(row['LastReviewed'],int)
        self.assertIn('Notes',changes['one'])
        self.assertEqual(next(h for h in history if h['Event']=='Generated' and h['DeviceId']=='one')['Notes'],'original note')
        self.assertEqual(next(h for h in history if h['Event']=='Saved edits')['Notes'],'')

    def test_deleted_report_row_is_not_a_deleted_record(self):
        register,_,changes=self.collect(self.weekly[1:])
        self.assertEqual(len(register),2)
        self.assertEqual(next(r for r in register if r['DeviceId']=='one')['Notes'],'original note')
        self.assertEqual(changes,{})

    def test_missing_device_and_reappearance_preserve_state_and_last_counts(self):
        register,history,_=self.collect()
        missing=s.finalize_snapshot(dict(tracking=register,records=self.records[1:],reportDate='2026-11-16'),history)
        absent=next(t for t in missing['tracking'] if t['DeviceId']=='one')
        self.assertFalse(absent['ObservedInExport'])
        self.assertEqual(absent['LastObservedReportDate'],'2026-11-09')
        self.assertEqual(absent['LastOverdueKEVs'],1)
        missing_history=next(h for h in missing['history'] if h['SnapshotId']==missing['snapshotId'] and h['DeviceId']=='one')
        self.assertEqual(missing_history['Overdue KEVs'],'')
        self.assertFalse(missing_history['Observed in export'])
        returned=s.finalize_snapshot(dict(tracking=missing['tracking'],records=self.records,reportDate='2026-11-23'),missing['history'])
        self.assertEqual(next(t for t in returned['tracking'] if t['DeviceId']=='one')['Notes'],'original note')
        self.assertEqual(len(returned['history']),6)

    def test_same_hostname_new_id_does_not_inherit_notes(self):
        new=copy.deepcopy(self.records[0]);new['DeviceId']='new-device'
        register=p.merge_tracking([new],copy.deepcopy(self.register))
        self.assertEqual(next(t for t in register if t['DeviceId']=='new-device')['Notes'],'')
        self.assertEqual(next(t for t in register if t['DeviceId']=='one')['Notes'],'original note')

    def test_conflicting_register_and_report_edits_fail(self):
        backing=copy.deepcopy(self.register);backing[0]['Notes']='backing edit'
        weekly=copy.deepcopy(self.weekly);weekly[0]['Notes']='different report edit'
        with self.assertRaisesRegex(ValueError,'Conflicting edits'):
            self.collect(weekly,backing)
        self.assertEqual(self.register[0]['Notes'],'original note')

    def test_backing_only_edit_is_retained_and_captured(self):
        backing=copy.deepcopy(self.register);backing[0]['Notes']='backing edit'
        result,history,_=self.collect(register=backing)
        self.assertEqual(result[0]['Notes'],'backing edit')
        self.assertEqual(history[-1]['Event'],'Saved edits')

    def test_duplicate_and_changed_device_id_fail(self):
        with self.assertRaisesRegex(ValueError,'Duplicate DeviceId in Weekly'):
            self.collect(self.weekly+[copy.deepcopy(self.weekly[0])])
        weekly=copy.deepcopy(self.weekly);weekly[0]['DeviceId']='unknown'
        with self.assertRaisesRegex(ValueError,'unrecognized DeviceId'):
            self.collect(weekly)

    def test_imported_counts_and_names_cannot_be_overridden(self):
        for field,value in [('Overdue KEVs',0),('All KEVs',999),('Device','another-device'),('Exposure','Low')]:
            with self.subTest(field=field):
                weekly=copy.deepcopy(self.weekly);weekly[0][field]=value
                with self.assertRaisesRegex(ValueError,'Imported field changed'):
                    self.collect(weekly)

    def test_blank_defaults_and_zero_note_are_intentional_edits(self):
        weekly=copy.deepcopy(self.weekly)
        weekly[0].update({'Remediation status':'','Escalation':'','Identity match':'','Notes':0})
        result,_,_=self.collect(weekly)
        self.assertEqual(result[0]['RemediationStatus'],'Unspecified')
        self.assertEqual(result[0]['NeedsEscalation'],'No')
        self.assertEqual(result[0]['IdentityMatch'],'Unverified DeviceId')
        self.assertEqual(result[0]['Notes'],'0')

    def test_invalid_status_date_or_manual_formula_fails(self):
        for field,value in [('Remediation status','Done'),('Last reviewed','yesterday'),('Escalation','Maybe')]:
            with self.subTest(field=field):
                weekly=copy.deepcopy(self.weekly);weekly[0][field]=value
                with self.assertRaises(ValueError):
                    self.collect(weekly)
        with self.assertRaisesRegex(ValueError,'contains a formula'):
            self.collect(formulas={'one':{'Notes'}})

    def test_legacy_formula_caches_are_not_mistaken_for_edits(self):
        guide={'Report date':'2026-11-09'}
        weekly=copy.deepcopy(self.weekly);weekly[0]['ServiceNow tickets']=0
        result,history,_=self.collect(weekly,guide=guide,history=[],formulas={'one':{'ServiceNow tickets'}})
        self.assertEqual(result[0]['ServiceNowTickets'],'INC-one')
        self.assertEqual(history[0]['Event'],'Migrated baseline')

    def test_literal_legacy_report_edits_are_migrated(self):
        guide={'Report date':'2026-11-09'}
        weekly=copy.deepcopy(self.weekly);weekly[0]['Notes']='saved in an old report'
        result,history,_=self.collect(weekly,guide=guide,history=[])
        self.assertEqual(result[0]['Notes'],'saved in an old report')
        self.assertEqual(history[-1]['Event'],'Saved edits')

    def test_missing_or_duplicate_history_baseline_fails(self):
        for history in ([],self.history+[copy.deepcopy(self.history[0])]):
            with self.assertRaises(ValueError):
                self.collect(history=history)

    def test_history_keeps_actual_cve_sets_and_distinct_same_day_runs(self):
        second=s.finalize_snapshot(dict(tracking=copy.deepcopy(self.register),records=self.records,reportDate='2026-11-09'),self.history)
        self.assertNotEqual(second['snapshotId'],self.payload['snapshotId'])
        self.assertEqual(len(second['history']),4)
        self.assertEqual(second['history'][-2]['KEVs'],'["CVE-1","CVE-2"]')


if __name__=='__main__':
    unittest.main(verbosity=2)
