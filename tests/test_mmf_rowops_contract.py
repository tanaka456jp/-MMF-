"""MMF safe row operations and moving-average formula repair contracts.

These tests validate the VBAUpdater wiring and ensure F7/F8 no longer use
whole-row deletes for MMF ledgers. Native Excel execution is a separate gate.
"""
import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ROWOPS = (ROOT / "source/standard/ModuleMMFRowOps.bas").read_text(encoding="utf-8")
LEGACY = (ROOT / "source/standard/Module1.bas").read_text(encoding="utf-8")
WORKBOOK = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
MANIFEST = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))


class MMFRowOpsContracts(unittest.TestCase):
    def test_vbaupdater_registers_module(self):
        components = {(c["type"], c["name"]) for c in MANIFEST["components"]}
        managed = {(c["type"], c["name"]) for c in MANIFEST["sync"]["managedComponents"]}
        self.assertIn(("standard", "ModuleMMFRowOps"), components)
        self.assertIn(("standard", "ModuleMMFRowOps"), managed)
        self.assertIn(("document", "ThisWorkbook"), managed)

    def test_f7_dispatches_before_legacy_full_row_delete(self):
        section = LEGACY.split("Sub F7_行削除()", 1)[1].split("Sub F8_行追加()", 1)[0]
        self.assertIn("If MMF台帳_対象(ws) Then", section)
        self.assertIn("Call MMF台帳_選択行削除", section)
        self.assertLess(section.index("Call MMF台帳_選択行削除"),
                        section.index("ws.Rows(arr(i)).Delete"))

    def test_f8_dispatches_before_legacy_full_row_copy(self):
        section = LEGACY.split("Sub F8_行追加()", 1)[1].split("Public Sub SetSpaceKey()", 1)[0]
        self.assertIn("If MMF台帳_対象(ws) Then", section)
        self.assertIn("Call MMF台帳_選択行追加", section)
        self.assertLess(section.index("Call MMF台帳_選択行追加"),
                        section.index("ws.Rows(baseRow + 1).Insert"))

    def test_row_ops_only_relocate_input_cells(self):
        self.assertIn('ws.Range("A" & rowNumber & ":J" & (lastRow - 1)).FormulaR1C1', ROWOPS)
        self.assertIn('ws.Range("A" & (rowNumber + 1) & ":J" & lastRow).FormulaR1C1', ROWOPS)
        self.assertIn('ws.Range("A" & (shiftStart + 1) & ":J" & lastRow).FormulaR1C1', ROWOPS)
        self.assertIn('ws.Range("A" & shiftStart & ":J" & (lastRow - 1)).FormulaR1C1', ROWOPS)
        self.assertIn('ws.Range("A" & lastRow & ":J" & lastRow).ClearContents', ROWOPS)
        self.assertIn('ws.Range("A" & shiftStart & ":J" & shiftStart).ClearContents', ROWOPS)
        self.assertNotIn("ws.Rows(rowNumber).Delete", ROWOPS)
        self.assertNotIn("ws.Rows(rowNumber).Insert", ROWOPS)

    def test_header_and_tail_protected(self):
        self.assertIn('If rowNumber < MMF_FIRST Or rowNumber > lastRow Then', ROWOPS)
        self.assertIn('MMF_FIRST As Long = 12', ROWOPS)
        self.assertIn('If selectedRows(1) >= lastRow Then', ROWOPS)
        self.assertIn('If Application.CountA(ws.Range("A" & (lastRow - n + 1)', ROWOPS)
        self.assertIn('末尾の空行が不足しています', ROWOPS)
        self.assertIn('vbYesNo + vbQuestion + vbDefaultButton2', ROWOPS)

    def test_formula_restoration_uses_first_and_subsequent_row_templates(self):
        self.assertIn('firstFormula = master.Range("K12:Y12").FormulaR1C1', ROWOPS)
        self.assertIn('nextFormula = master.Range("K13:Y13").FormulaR1C1', ROWOPS)
        self.assertIn('target.Cells(12, col).FormulaR1C1 =', ROWOPS)
        self.assertIn('target.Range(target.Cells(13, col), target.Cells(lastRow, col)).FormulaR1C1', ROWOPS)
        self.assertIn('For col = 11 To 25', ROWOPS)  # K:Y including FX helper Y
        self.assertIn('If Not ws.Cells(r, col).HasFormula Then Exit Function', ROWOPS)
        self.assertIn('InStr(1, f, "#REF!", vbTextCompare)', ROWOPS)
        self.assertIn('Set master = candidate', ROWOPS)
        self.assertIn('If Not MMF台帳_数式復元(ws, True) Then', ROWOPS)

    def test_native_row_delete_triggers_structural_repair(self):
        self.assertEqual(1, WORKBOOK.count('Private Sub Workbook_SheetChange('))
        self.assertIn('Call MMF台帳_構造変更時修復(Sh, Target)', WORKBOOK)
        self.assertIn('Set hit = Application.Intersect(changed, ws.Range("K12:Y"', ROWOPS)
        self.assertIn('If hit Is Nothing Then Exit Sub', ROWOPS)
        self.assertIn('Call MMF台帳_数式復元(ws, False)', ROWOPS)
        self.assertIn('If mRowHandling Then Exit Sub', ROWOPS)

    def test_can_repair_existing_corrupted_rows_without_input_loss(self):
        sub = ROWOPS.split('Public Sub MMF台帳_計算式修復()', 1)[1].split('End Sub', 1)[0]
        self.assertIn('MMF台帳_数式復元(ws, False)', sub)
        repair = ROWOPS.split('Private Function MMF台帳_数式復元(', 1)[1].split('End Function', 1)[0]
        self.assertNotIn('"A" &', repair)
        self.assertNotIn('"J" &', repair)
        self.assertIn('target.Calculate', repair)

    def test_general_specific_and_legacy_names(self):
        for name in ('"MMF取引台帳"', '"一般口座"', '"特定口座"'):
            self.assertIn(name, ROWOPS)

    def test_model_of_input_only_deletion_preserves_formula_row_count(self):
        # F7 deletes data at selected row then shifts inputs upward.
        # K:Y remain in their original rows and are reconstructed from templates.
        original = ["trade-12", "trade-13", "trade-14", "trade-15", None, None]
        data = list(original)
        for selected in sorted((12, 14), reverse=True):
            index = selected - 12
            data[index:-1] = data[index + 1:]
            data[-1] = None
        self.assertEqual(["trade-13", "trade-15", None, None, None, None], data)
        self.assertEqual(len(data), len(original))

    def test_model_of_input_only_insertion_does_not_duplicate_transactions(self):
        data = ["A", "B", "C", None, None]
        insert_after = 12
        index = insert_after - 12 + 1
        data[index + 1:] = data[index:-1]
        data[index] = None
        self.assertEqual(["A", None, "B", "C", None], data)


if __name__ == "__main__":
    unittest.main()
