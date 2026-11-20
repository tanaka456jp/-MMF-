"""Contract tests for manual transaction-category input in the MMF ledger.

Checks VBAUpdater wiring and safety of code changes; Excel/VBE runtime tests still required.
"""
import json
import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "source/standard/ModuleMMFInput.bas").read_text(encoding="utf-8")
BOOK_CODE = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
MANIFEST = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


class MMFInputContractTests(unittest.TestCase):
    def test_vbaupdater_manifest_covers_new_module(self):
        comps = {(c["type"], c["name"]) for c in MANIFEST["components"]}
        managed = {(c["type"], c["name"]) for c in MANIFEST["sync"]["managedComponents"]}
        self.assertIn(("standard", "ModuleMMFInput"), comps)
        self.assertIn(("standard", "ModuleMMFInput"), managed)
        self.assertIn(("document", "ThisWorkbook"), managed)

    def test_existing_workbook_has_specific_type_dropdown(self):
        path = ROOT / MANIFEST["target"]["directory"] / MANIFEST["target"]["file"]
        with zipfile.ZipFile(path) as book:
            wb = ET.fromstring(book.read("xl/workbook.xml"))
            rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
            targets = {node.get("Id"): node.get("Target") for node in rels}
            sh = next(x for x in wb.findall("m:sheets/m:sheet", NS)
                      if x.get("name") == "MMF取引台帳")
            target = targets[sh.get("{" + NS["r"] + "}id")].lstrip("/")
            target = target if target.startswith("xl/") else "xl/" + target
            doc = ET.fromstring(book.read(target))
            validations = doc.findall("m:dataValidations/m:dataValidation", NS)
            refs = {x.get("sqref"): x.get("type") for x in validations}
        self.assertEqual(refs.get("B12:B511"), "list")
        self.assertEqual(refs.get("H12:H511"), "list")
        self.assertEqual(refs.get("F3"), "list")

    def test_only_b_transactions_have_validation_removed(self):
        self.assertIn('ws.Range("B" & FIRST_TRADE_ROW & ":B" & lastInputRow)', SOURCE)
        self.assertIn(".Validation.Delete", SOURCE)
        self.assertIn('.NumberFormat = "@"', SOURCE)
        self.assertNotIn("ws.Cells.Clear", SOURCE)
        self.assertNotIn("ws.Range(\"A12", SOURCE)
        self.assertNotIn(".Formula =", SOURCE)
        self.assertNotIn(".ClearContents", SOURCE)
        self.assertNotIn('.Range("H', SOURCE)

    def test_calculation_semantics_remain_explicit(self):
        # The user's label is unconstrained at entry, but unsupported labels must
        # still be rejected by the pre-existing K-column validation, not guessed.
        self.assertIn("既存の K 列", SOURCE)
        self.assertIn('CStr(ws.Range("B11").Value2) <> "取引区分"', SOURCE)
        self.assertIn("買付数量(C列)", SOURCE)
        self.assertIn("売付数量(E列)", SOURCE)
        self.assertIn("取引区分の文字列は計算に使用しません", SOURCE)
        self.assertNotIn("Worksheets(MMF_INPUT_SHEET).Calculate", SOURCE)

    def test_startup_and_manual_macros(self):
        self.assertIn("Public Sub MMF取引区分_自由入力設定()", SOURCE)
        self.assertIn("Public Sub MMF取引区分_起動時設定()", SOURCE)
        self.assertEqual(BOOK_CODE.count("Private Sub Workbook_Open()"), 1)
        self.assertIn("Call MMF取引区分_起動時設定", BOOK_CODE)
        self.assertEqual(BOOK_CODE.count("Private Sub Workbook_BeforePrint("), 1)


if __name__ == "__main__":
    unittest.main()
