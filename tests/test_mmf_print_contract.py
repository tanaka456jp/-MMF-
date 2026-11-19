"""Static regression checks for the A4 portrait print VBA / workbook contract.

These tests do NOT replace Excel/VBE compilation or a printer preview on Windows.
"""
import json
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "source/standard/ModuleMMFPrint.bas").read_text(encoding="utf-8")
WORKBOOK = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
MANIFEST = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def get_sheet_cells(book, name):
    wb = ET.fromstring(book.read("xl/workbook.xml"))
    rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    targets = {node.get("Id"): node.get("Target") for node in rels}
    sheets = wb.findall("m:sheets/m:sheet", NS)
    sh = next(s for s in sheets if s.get("name") == name)
    path = targets[sh.get("{" + NS["r"] + "}id")].lstrip("/")
    if not path.startswith("xl/"):
        path = "xl/" + path
    root = ET.fromstring(book.read(path))
    shared = []
    if "xl/sharedStrings.xml" in book.namelist():
        strings = ET.fromstring(book.read("xl/sharedStrings.xml"))
        shared = [
            "".join(t.text or "" for t in element.findall(".//m:t", NS))
            for element in strings.findall("m:si", NS)
        ]
    values = {}
    for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
        inline = cell.find("m:is", NS)
        value = cell.find("m:v", NS)
        if inline is not None:
            text = "".join(t.text or "" for t in inline.findall(".//m:t", NS))
        elif value is not None and cell.get("t") == "s":
            text = shared[int(value.text)]
        elif value is not None:
            text = value.text
        else:
            continue
        values[cell.get("r")] = text
    return values


class PrintContractTests(unittest.TestCase):
    def test_existing_workbook_schema(self):
        target = ROOT / MANIFEST["target"]["directory"] / MANIFEST["target"]["file"]
        with zipfile.ZipFile(target) as book:
            cells = get_sheet_cells(book, "MMF取引台帳")
        expected = {
            "A11": "日付", "B11": "取引区分", "C11": "買付数量",
            "E11": "売付数量", "J11": "備考", "K11": "入力判定",
            "L11": "換算レート", "M11": "買付価額・外貨",
            "N11": "買付価額・円", "O11": "売付価額・外貨",
            "P11": "売付価額・円", "R11": "売却取得費・円",
            "S11": "売却損益・円", "T11": "残高数量",
            "U11": "残高・外貨", "V11": "残高・円",
        }
        for cell, text in expected.items():
            with self.subTest(cell=cell):
                self.assertEqual(text, cells.get(cell))

    def test_vbaupdater_modules_are_managed(self):
        components = {(c["type"], c["name"]) for c in MANIFEST["components"]}
        managed = {(c["type"], c["name"]) for c in MANIFEST["sync"]["managedComponents"]}
        for item in [("standard", "ModuleMMFPrint"), ("document", "ThisWorkbook")]:
            self.assertIn(item, components)
            self.assertIn(item, managed)
        self.assertEqual(MANIFEST["sync"]["deletePolicy"], "report-only")

    def test_a4_portrait_12_columns_repeated_header(self):
        for pattern in (
            'PRINT_COLUMNS As Long = 12',
            '.PrintArea = "$A$1:$L$"',
            '.PrintTitleRows = "$1:$5"',
            '.PaperSize = xlPaperA4',
            '.Orientation = xlPortrait',
            '.FitToPagesWide = 1',
            '.FitToPagesTall = False',
        ):
            with self.subTest(pattern=pattern):
                self.assertIn(pattern, MODULE)
        self.assertNotIn('.PrintArea = "$A$1:$O$"', MODULE)

    def test_no_blank_date_transaction_loss(self):
        self.assertIn('Application.CountA(src.Range("A" & rowNo & ":J" & rowNo))', MODULE)
        self.assertIn('"日付未入力"', MODULE)
        self.assertIn('ReDim data(1 To n, 1 To PRINT_COLUMNS)', MODULE)
        self.assertIn('If n > 0 Then', MODULE)
        self.assertIn('lastSource = STANDARD_LAST_ROW', MODULE)

    def test_bad_row_is_visible_and_highlighted(self):
        self.assertIn('CStr(src.Cells(rowNo, "K").Value2) <> "OK"', MODULE)
        self.assertIn('=LEFT($L6,3)=""要確認""', MODULE)
        self.assertIn('CStr(bad) & " 件"', MODULE)
        self.assertIn('"要確認："', MODULE)

    def test_before_print_event_is_gated_and_cancellable(self):
        self.assertEqual(WORKBOOK.count("Private Sub Workbook_BeforePrint("), 1)
        self.assertIn('If ActiveSheet.Name <> "A4縦印刷" Then Exit Sub', WORKBOOK)
        self.assertIn('If Not MMF印刷_印刷直前更新() Then Cancel = True', WORKBOOK)
        self.assertIn('If mUpdating Then', MODULE)

    def test_core_calculation_is_not_rewritten(self):
        self.assertIn('expr = MMF印刷_参照(sheetName, "T", sourceRow)', MODULE)
        self.assertIn('expr = MMF印刷_参照(sheetName, "U", sourceRow)', MODULE)
        self.assertIn('expr = MMF印刷_参照(sheetName, "V", sourceRow)', MODULE)
        self.assertIn('expr = MMF印刷_参照(sheetName, "L", sourceRow)', MODULE)


if __name__ == "__main__":
    unittest.main()
