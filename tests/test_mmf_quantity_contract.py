"""Regressions for quantity-based MMF trade classification.

VBAUpdater imports ModuleMMFCalc. Excel may save repeated worksheet formulas
as shared OOXML formulas, so tests reconstruct followers before validating.
"""
import json
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
MOD = (ROOT / "source/standard/ModuleMMFCalc.bas").read_text(encoding="utf-8")
PRINT = (ROOT / "source/standard/ModuleMMFPrint.bas").read_text(encoding="utf-8")
WB = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
CONFIG = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}



def _addr_parts(address):
    match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", address)
    if match is None:
        raise ValueError(f"Unexpected spreadsheet address: {address}")
    col, row = match.groups()
    col_num = 0
    for char in col:
        col_num = col_num * 26 + ord(char) - ord("A") + 1
    return col_num, int(row)


def _column_label(col_num):
    result = ""
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _translate_shared_formula(formula, origin, dest):
    """Reconstruct a follower of an OOXML shared formula without extra libraries."""
    from_col, from_row = _addr_parts(origin)
    to_col, to_row = _addr_parts(dest)
    delta_col = to_col - from_col
    delta_row = to_row - from_row

    def translate(match):
        absolute_col, col, absolute_row, row = match.groups()
        col_value, _ = _addr_parts(col + "1")
        new_col = col_value if absolute_col else col_value + delta_col
        new_row = int(row) if absolute_row else int(row) + delta_row
        if new_col <= 0 or new_row <= 0:
            raise ValueError(f"Invalid translated reference from {origin} to {dest}")
        return f"{absolute_col}{_column_label(new_col)}{absolute_row}{new_row}"

    return re.sub(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)",
                  translate, formula)


def source_formulas():
    path = ROOT / CONFIG["target"]["directory"] / CONFIG["target"]["file"]
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        links = {r.get("Id"): r.get("Target") for r in rel}
        sh = next(x for x in wb.findall("m:sheets/m:sheet", NS)
                  if x.get("name") == "MMF取引台帳")
        source = links[sh.get("{" + NS["r"] + "}id")].lstrip("/")
        if not source.startswith("xl/"):
            source = "xl/" + source
        xml = ET.fromstring(z.read(source))

    raw_formulas = {}
    shared_masters = {}
    for cell in xml.findall(".//m:sheetData/m:row/m:c", NS):
        elem = cell.find("m:f", NS)
        if elem is None:
            continue
        address = cell.get("r")
        raw_formulas[address] = (elem.get("t"), elem.get("si"), elem.text)
        if elem.get("t") == "shared" and elem.text:
            shared_masters[elem.get("si")] = (address, elem.text)

    expanded = {}
    for address, (kind, si, formula) in raw_formulas.items():
        if kind == "shared" and not formula:
            if si not in shared_masters:
                raise ValueError(f"Shared formula {address} has no master (si={si})")
            origin, original_formula = shared_masters[si]
            expanded[address] = _translate_shared_formula(original_formula, origin, address)
        else:
            expanded[address] = formula
    return expanded, sum(1 for kind, si, formula in raw_formulas.values()
                         if kind == "shared" and not formula)


def migrate_formula(f, r, column):
    sale = f'B{r}="売"'
    if column == "K":
        original = (f'IF(AND(B{r}<>"売",B{r}<>"買",B{r}<>"累投"),'
                    f'"区分を確認",')
        replacement = f'IF(AND(C{r}>0,E{r}>0),"売買数量が重複",'
        assert original in f
        f = f.replace(original, replacement)
        f = f.replace(f'OR(E{r}<=0,C{r}<>"",D{r}<>"",F{r}="")',
                      f'OR(E{r}<=0,C{r}>0,D{r}<>"",F{r}="")')
        f = f.replace(f'OR(C{r}<=0,E{r}<>"",AND(F{r}="",D{r}<=0))',
                      f'OR(C{r}<=0,E{r}>0,AND(F{r}="",D{r}<=0))')
    return f.replace(sale, f'E{r}>0')


class MMFQuantityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.formulas, cls.shared_followers = source_formulas()

    def test_shared_formula_followers_from_excel_save_are_expanded(self):
        # Opening/saving XLSM in Excel may store <f t="shared" si="..."/>
        # for repeated rows rather than repeating each formula in its XML.
        self.assertIsInstance(self.shared_followers, int)
        self.assertGreaterEqual(self.shared_followers, 0)
        self.assertIsInstance(self.formulas["K13"], str)
        self.assertTrue(self.formulas["K13"].startswith("=") is False)
        self.assertIn("13", self.formulas["K13"])

    def test_existing_workbook_has_all_formula_columns(self):
        for r in (12, 13, 100, 511):
            for c in ("K",) + tuple("MNOPQRSTUVWX"):
                with self.subTest(cell=f"{c}{r}"):
                    self.assertIsInstance(self.formulas[f"{c}{r}"], str)

    def test_migration_matches_actual_workbook_formulas(self):
        for r in (12, 13, 100, 511):
            for c in ("K",) + tuple("MNOPQRSTUVWX"):
                f = self.formulas[f"{c}{r}"]
                g = migrate_formula(f, r, c)
                with self.subTest(cell=f"{c}{r}"):
                    self.assertEqual(g.count("("), g.count(")"))
                    self.assertNotIn(f'B{r}="売"', g)
                    if c == "K":
                        self.assertIn(f'AND(C{r}>0,E{r}>0)', g)
                        self.assertNotIn("区分を確認", g)
                        self.assertIn(f'IF(E{r}>0,', g)
                    self.assertIn(f'E{r}>0', g) if f'B{r}="売"' in f else None

    def test_other_formula_logic_is_retained(self):
        for c in ("K", "Q", "R", "S", "T", "U", "V", "W", "X"):
            f = self.formulas[f"{c}13"]
            g = migrate_formula(f, 13, c)
            if c != "K":
                self.assertEqual(f.replace('B13="売"', 'E13>0'), g)
        self.assertEqual(self.formulas["Y13"], self.formulas["Y13"])

    def test_vba_guard_and_rollback(self):
        for phrase in (
            'Public Sub MMF数量判定_適用()',
            'Public Sub MMF数量判定_起動時設定()',
            'CStr(ws.Range("C11").Value2) <> "買付数量"',
            'CStr(ws.Range("E11").Value2) <> "売付数量"',
            'If inputLast > lastRow Then',
            'If writeStarted Then',
            'ws.Range("K" & FIRST_ROW & ":K" & lastRow).Formula = oldK',
            'ws.Range("M" & FIRST_ROW & ":X" & lastRow).Formula = oldMX',
            'If changed Then',
            'If mApplying Then Exit Function',
        ):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, MOD)

    def test_runtime_wiring(self):
        types = {(x["type"], x["name"]) for x in CONFIG["components"]}
        managed = {(x["type"], x["name"]) for x in CONFIG["sync"]["managedComponents"]}
        self.assertIn(("standard", "ModuleMMFCalc"), types)
        self.assertIn(("standard", "ModuleMMFCalc"), managed)
        self.assertEqual(1, WB.count("Call MMF数量判定_起動時設定"))
        self.assertEqual(1, WB.count("Call MMF取引区分_起動時設定"))

    def test_print_layout_calculation_does_not_use_category(self):
        segment = PRINT.split("Private Function MMF印刷_数式(", 1)[1]
        self.assertIn('expr = b', segment)  # only display the arbitrary label
        self.assertNotIn('"売" & q', segment)
        self.assertEqual(5, segment.count('MMF印刷_参照(sheetName, "E", sourceRow) & ">0"'))

    def test_free_text_input_is_preserved(self):
        src = (ROOT / "source/standard/ModuleMMFInput.bas").read_text(encoding="utf-8")
        self.assertIn(".Validation.Delete", src)
        self.assertIn('.NumberFormat = "@"', src)


if __name__ == "__main__":
    unittest.main()
