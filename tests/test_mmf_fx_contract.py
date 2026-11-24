"""Static regression for MMF FX import lookup with workbook formula migration.

The original XLSM remains unchanged in Git. VBAUpdater imports the new
module; Workbook_Open installs formulas in the local user's XLSM.
"""
import json
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
CODE = (ROOT / "source/standard/ModuleMMFFX.bas").read_text(encoding="utf-8")
WB = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
CONFIG = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def read_legacy_formulas():
    target = ROOT / CONFIG["target"]["directory"] / CONFIG["target"]["file"]
    with zipfile.ZipFile(target) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        links_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        links = {x.get("Id"): x.get("Target") for x in links_xml}
        sh = next(s for s in book.findall("x:sheets/x:sheet", NS)
                  if s.get("name") == "MMF取引台帳")
        p = links[sh.get("{" + NS["r"] + "}id")].lstrip("/")
        if not p.startswith("xl/"):
            p = "xl/" + p
        xml = ET.fromstring(z.read(p))
        cells = xml.findall(".//x:sheetData/x:row/x:c", NS)
        return {c.get("r"): c.find("x:f", NS).text for c in cells
                if c.find("x:f", NS) is not None}


class FXContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.legacy = read_legacy_formulas()

    def test_old_formula_can_block_loaded_fx_data(self):
        y12 = self.legacy["Y12"]
        self.assertIn("COUNTIF(為替データ!$A$5:$A$10000,A12)<>1", y12)
        self.assertIn("MATCH($F$3,為替データ!$C$3:$CL$3,0)", y12)
        self.assertIn("FX取得(A12,H12,$F$3)", self.legacy["L12"])

    def test_currency_code_normalized_and_merged_3_column_headers(self):
        self.assertIn('Replace(textLabel, "（", "(")', CODE)
        self.assertIn('Replace(textLabel, "）", ")")', CODE)
        self.assertIn('For col = 3 To lastCol Step 3', CODE)
        self.assertIn('Case "TTS": rateOffset = 0', CODE)
        self.assertIn('Case "TTB": rateOffset = 1', CODE)
        self.assertIn('Case "TTM": rateOffset = 2', CODE)

    def test_duplicate_date_conflicts_not_silently_accepted(self):
        self.assertIn('If Abs(CDbl(rawRate) - foundRate) > 0.00000001 Then GoTo AmbiguousRate', CODE)
        self.assertIn('MMF為替_取得 = CVErr(xlErrNum)', CODE)
        self.assertIn('MMF為替_取得 = foundRate', CODE)
        self.assertIn('If sourceDay = targetDay Then', CODE)
        self.assertIn('If CDbl(rawRate) <= 0 Then GoTo InvalidInput', CODE)

    def test_rate_lookup_bulk_reads_without_changing_other_forex_macros(self):
        self.assertIn('datesData = fx.Range("A5:A" & lastRow).Value2', CODE)
        self.assertIn('ratesData = fx.Range(fx.Cells(5, sourceCol)', CODE)
        old = (ROOT / "source/standard/Module1.bas").read_text(encoding="utf-8")
        self.assertIn('Function FX取得(', old)
        self.assertIn('Set ws = ThisWorkbook.Worksheets("為替データ")', old)

    def test_update_has_backup_and_is_idempotent(self):
        for term in (
            'Public Sub MMF為替_参照式更新()',
            'Public Sub MMF為替_起動時設定()',
            'If changed Then',
            'If writeStarted Then',
            'ledger.Range("L" & FIRST_ROW & ":L" & lastRow).Formula = origL',
            'ledger.Range("Y" & FIRST_ROW & ":Y" & lastRow).Formula = origY',
            'IFERROR(MMF為替_取得(',
            'MMF為替_参照式更新本体 = True',
        ):
            with self.subTest(term=term):
                self.assertIn(term, CODE)

    def test_installer_is_registered(self):
        self.assertEqual(1, WB.count("Call MMF為替_起動時設定"))
        components = {(x["type"], x["name"]) for x in CONFIG["components"]}
        managed = {(x["type"], x["name"]) for x in CONFIG["sync"]["managedComponents"]}
        self.assertIn(("standard", "ModuleMMFFX"), components)
        self.assertIn(("standard", "ModuleMMFFX"), managed)

    def test_diagnostic_macro_exposes_reason(self):
        self.assertIn('Public Sub MMF為替_初回取引診断()', CODE)
        for term in ('"通貨コード："', '"取引日のデータ行数："',
                     '"為替ヘッダー列："', '"取得レート："'):
            self.assertIn(term, CODE)


if __name__ == "__main__":
    unittest.main()
