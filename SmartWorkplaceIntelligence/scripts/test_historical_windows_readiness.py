"""Synthetic tests for historical readiness qualification and identity uncertainty."""
import csv
import tempfile
import unittest
from pathlib import Path
import historical_windows_readiness as h


def sample(**changes):
    row = dict.fromkeys(h.RAW_COLUMNS, "")
    row.update({"SourceFile": "A", "pc.Date": "2025-09-16 11:48:48Z: ", "pc.name": "PC01", "pc.GLOBALResult": "CAPABLE", "pc.Model": "Model A"})
    row.update(changes)
    return row

class ReadinessTest(unittest.TestCase):
    def test_qualification_requires_exact_hash_and_never_overwrites(self):
        with tempfile.TemporaryDirectory() as d:
            src, out = Path(d)/"input.csv", Path(d)/"output.csv"
            with src.open("w", newline="", encoding="utf-8") as f:
                w=csv.DictWriter(f,fieldnames=h.RAW_COLUMNS);w.writeheader();w.writerow(sample())
            expected=h.digest(src)
            with self.assertRaises(ValueError):h.qualify(src,out,"test","0"*64)
            self.assertFalse(out.exists())
            self.assertEqual(h.qualify(src,out,"test",expected)["rows"],1)
            before=out.read_bytes()
            with self.assertRaises(ValueError):h.qualify(src,out,"test",expected)
            self.assertEqual(out.read_bytes(),before)
            with out.open(encoding="utf-8",newline="") as f: rows=list(csv.DictReader(f))
            self.assertEqual(rows[0]["TenantKey"],"test")
            self.assertEqual(h.validate_qualified(out,"test")["rows"],1)
            with self.assertRaises(ValueError):h.validate_qualified(out,"other")
            with self.assertRaises(ValueError):h.project(rows,"other",[],[],[])

    def test_only_unassessed_latest_and_intune_precedence(self):
        prefix={"TenantKey":"test","HistoricalSourceSHA256":"A"*64,"HistoricalObservationUtc":"2025-09-16T11:48:48Z"}
        a={**prefix,**sample()}; b={**prefix,**sample(SourceFile="B")}
        current={"Device Name":"pc01","Device Source ID":"id1","Inventory Source":"Intune","Upgrade Eligibility":"Unknown"}
        inv={"TenantKey":"test","Device name":"pc01","Device ID":"id1","Model":"Model A"}
        rows=h.project([a,b],"test",[current],[inv],[])
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]["Duplicate Count"],2)
        self.assertEqual(rows[0]["Source Files"],"A | B")
        self.assertIn("identity unverified",rows[0]["Match Status"])
        self.assertEqual(rows[0]["Observation UTC"],"2025-09-16T11:48:48Z")
        for status in ("Capable", "Not capable", "Already upgraded"):
            self.assertEqual(h.project([a],"test",[{**current,"Upgrade Eligibility":status}],[inv],[]),[])
        for status in ("Unknown", "Not assessed", "Undetermined"):
            self.assertEqual(len(h.project([a],"test",[{**current,"Upgrade Eligibility":status}],[inv],[])),1)
        newer={**prefix,**sample(**{"pc.Date":"2025-10-03 10:00:00Z:","pc.GLOBALResult":"NOT CAPABLE"}),"HistoricalObservationUtc":"2025-10-03T10:00:00Z"}
        self.assertEqual(h.project([a,newer],"test",[current],[inv],[])[0]["Historical Result"],"NOT CAPABLE")

    def test_identity_ambiguity_model_and_tie_are_quarantined(self):
        prefix={"TenantKey":"test","HistoricalSourceSHA256":"A"*64,"HistoricalObservationUtc":"2025-09-16T11:48:48Z"}
        a={**prefix,**sample()}
        current={"Device Name":"PC01","Device Source ID":"a","Inventory Source":"Intune","Upgrade Eligibility":"Unknown"}
        inv={"TenantKey":"test","Device name":"PC01","Device ID":"a","Model":"Model A"}
        bad={**prefix,**sample(**{"pc.GLOBALResult":"NOT CAPABLE"})}
        rows=h.project([a,bad],"test",[current],[inv],[])
        self.assertEqual(rows[0]["Observation Status"],"Conflicting observations")
        self.assertEqual(rows[0]["Historical Result"],"")
        rows=h.project([a],"test",[current],[inv,{**inv,"Device ID":"b"}],[])
        self.assertEqual(rows[0]["Match Status"],"Ambiguous current name")
        self.assertEqual(rows[0]["Historical Result"],"")
        rows=h.project([a],"test",[current],[{**inv,"Model":"Other Model"}],[])
        self.assertEqual(rows[0]["Match Status"],"Model mismatch")
        self.assertEqual(rows[0]["Historical Result"],"")
        rows=h.project([a],"test",[{**current,"Upgrade Eligibility":"Capable"}],[inv],[])
        self.assertEqual(rows, [])
        rows=h.project([a],"test",[{**current,"Upgrade Eligibility":"Not capable"}],[inv],[])
        self.assertEqual(rows, [])
        with self.assertRaises(ValueError):h.project([a],"test",[current],[{**inv,"TenantKey":"foreign"}],[])
        ad={"TenantKey":"test","ObjectGUID":"guid1","Name":"PC01"}
        ad_current={"Device Name":"PC01","Device Source ID":"ad:test:guid1","Inventory Source":"AD only","Upgrade Eligibility":"Not assessed"}
        self.assertEqual(h.project([a],"test",[ad_current],[],[ad])[0]["Historical Result"],"CAPABLE")
        self.assertEqual(h.project([a],"test",[{**ad_current,"Device Source ID":"ad:test:other"}],[],[ad])[0]["Match Status"],"AD identity mismatch")

    def test_full_projection_and_optional_absence(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d); source=root/"source.csv"; qualified=root/"qualified.csv"; output=root/"prepared.csv"
            def write(path, fields, rows):
                with path.open("w",newline="",encoding="utf-8") as f:
                    w=csv.DictWriter(f,fieldnames=fields);w.writeheader();w.writerows(rows)
            write(source,h.RAW_COLUMNS,[sample(**{"pc.GLOBALResult":"NOT CAPABLE"})])
            h.qualify(source,qualified,"test",h.digest(source))
            migration=root/"migration.csv";inventory=root/"intune.csv";ad=root/"ad.csv"
            write(migration,["Device Source ID","Device Name","Inventory Source","Upgrade Eligibility"],[{"Device Source ID":"id1","Device Name":"PC01","Inventory Source":"Intune","Upgrade Eligibility":"Unknown"}])
            write(inventory,["TenantKey","Device ID","Device name","Model"],[{"TenantKey":"test","Device ID":"id1","Device name":"PC01","Model":"Model A"}])
            write(ad,["TenantKey","ObjectGUID","Name"],[])
            result=h.build_projection(qualified,migration,output,"test",inventory,ad)
            self.assertEqual(result["observations"],1)
            with output.open(encoding="utf-8",newline="") as f:rows=list(csv.DictReader(f))
            self.assertEqual(rows[0]["Historical Result"],"NOT CAPABLE")
            self.assertEqual(rows[0]["Candidate Device Source ID"],"id1")
            result=h.build_projection(root/"missing.csv",migration,output,"test",inventory,ad)
            self.assertFalse(result["source_available"])
            with output.open(encoding="utf-8",newline="") as f:self.assertEqual(list(csv.DictReader(f)),[])

    def test_dates_and_source_hash_cannot_be_invented(self):
        with self.assertRaises(ValueError):h.observation_utc("2025-09-16 11:48:48")
        a={"TenantKey":"test","HistoricalSourceSHA256":"A"*64,"HistoricalObservationUtc":"2025-09-16T11:48:48Z",**sample()}
        b={"TenantKey":"test","HistoricalSourceSHA256":"B"*64,"HistoricalObservationUtc":"2025-09-16T11:48:48Z",**sample(SourceFile="B")}
        with self.assertRaises(ValueError):h.project([a,b],"test",[],[],[])

if __name__ == "__main__": unittest.main()
