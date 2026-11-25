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

    def test_fx_udf_arguments_do_not_use_vba_reserved_identifiers(self):
        # Currency is a VBA built-in type/reserved identifier. Using
        # 'ByVal currency As Variant' causes 'Expected: identifier' in VBE.
        import re
        self.assertRegex(CODE, r"ByVal\s+currencyLabel\s+As\s+Variant")
        self.assertIn("IsError(currencyLabel)", CODE)
        self.assertIn("CStr(currencyLabel)", CODE)
        self.assertNotRegex(CODE, r"(?i)\bByVal\s+Currency\s+As\b")
        self.assertNotRegex(CODE, r"(?i)\bByRef\s+Currency\s+As\b")

    def test_excel_value2_date_serials_are_accepted(self):
        # Excel Range.Value2 returns a Double for a date formatted cell;
        # direct IsDate(value2) is not a reliable numeric-date validator.
        import datetime
        excel_serial = (datetime.date(2020, 5, 15) -
                        datetime.date(1899, 12, 30)).days
        self.assertEqual(43966, excel_serial)
        self.assertIn('Private Function MMF為替_日付シリアル(', CODE)
        self.assertIn('ElseIf IsNumeric(cellDate) Then', CODE)
        self.assertIn('numericDay = CDbl(cellDate)', CODE)
        self.assertIn('If ThisWorkbook.Date1904 Then numericDay = numericDay + 1462#', CODE)
        self.assertIn('dayNumber = CLng(Fix(numericDay))', CODE)
        self.assertIn('If Not MMF為替_日付シリアル(tradeDate, targetDay)', CODE)
        self.assertIn('If MMF為替_日付シリアル(cellValue, sourceDay)', CODE)
        self.assertIn('If Not MMF為替_日付シリアル(dayValue, targetDay)', CODE)
        self.assertNotIn('If Not IsDate(tradeDate)', CODE)
        self.assertNotIn('If Not IsDate(dayValue)', CODE)

    def test_fx_date_single_row_array_shape(self):
        self.assertIn('If lastRow = 5 Then', CODE)
        self.assertIn('ReDim datesData(1 To 1, 1 To 1)', CODE)
        self.assertIn('ReDim ratesData(1 To 1, 1 To 1)', CODE)

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
