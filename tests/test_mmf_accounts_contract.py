"""Static MMF dual-account migration and A4 summary contracts.

The user-provided one-page sample has specific/general quantity changes,
year-end quantities, yen acquisition unit cost, total units and foreign
balance. These checks do not replace Excel COM validation.
"""
import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "source/standard/ModuleMMFAccounts.bas").read_text(encoding="utf-8")
PRINT = (ROOT / "source/standard/ModuleMMFPrint.bas").read_text(encoding="utf-8")
WB = (ROOT / "source/document/ThisWorkbook.cls").read_text(encoding="utf-8")
MANIFEST = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))


class AccountSummaryContractTests(unittest.TestCase):
    def test_vbaupdater_registers_new_component(self):
        components = {(x["type"], x["name"]) for x in MANIFEST["components"]}
        managed = {(x["type"], x["name"]) for x in MANIFEST["sync"]["managedComponents"]}
        self.assertIn(("standard", "ModuleMMFAccounts"), components)
        self.assertIn(("standard", "ModuleMMFAccounts"), managed)
        self.assertEqual(1, WB.count("Call MMF口座_起動時設定"))
        self.assertIn('Call MMF印刷_ボタン設置', WB)

    def test_general_and_specific_sheets_do_not_duplicate_initial_transactions(self):
        for phrase in [
            'GENERAL_LEDGER As String = "一般口座"',
            'SPECIFIC_LEDGER As String = "特定口座"',
            'source.Copy After:=source',
            'source.Copy After:=gen',
            'MMF口座_台帳入力初期化 spec',
            'MMF口座_台帳入力初期化 gen',
            'source.Visible = xlSheetVeryHidden',
            'If answer = vbCancel Then GoTo CleanExit',
            'If (gen Is Nothing) Xor (spec Is Nothing) Then',
            'If Not gen Is Nothing And Not spec Is Nothing Then',
        ]:
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, SOURCE)
        self.assertIn("vbYesNoCancel", SOURCE)
        self.assertIn("既存「MMF取引台帳」の取引をどちら", SOURCE)

    def test_summary_columns_match_example(self):
        for col in (
            '"取引開始日"', '"終了日"', '"備考／取引区分"',
            '"特定口座（口）"', '"一般口座（口）"', '"取得単価（円）"',
            '"合計残高（口）"', '"合計残高（外貨）"',
        ):
            self.assertIn(col, SOURCE)
        self.assertIn('SUMMARY_SHEET As String = "口座別まとめ"', SOURCE)
        self.assertIn('If qty(1) + qty(2) > 0 Then', SOURCE)
        self.assertIn('CStr(prevYear) & "年12月31日現在"', SOURCE)
        self.assertIn('avgPrice = Round(totalCost / total, 9)', SOURCE)
        self.assertIn('WorksheetFunction.Round(total * valuePerUnit, 0)', SOURCE)

    def test_summary_is_only_from_two_account_sheets(self):
        self.assertIn('Set gen = MMF口座_シート取得(GENERAL_LEDGER)', SOURCE)
        self.assertIn('Set spec = MMF口座_シート取得(SPECIFIC_LEDGER)', SOURCE)
        self.assertNotIn('Set sourceWs = ThisWorkbook.Worksheets(LEGACY_LEDGER)', SOURCE)
        self.assertIn('If CStr(v) <> "OK" Then GoTo InvalidTransaction', SOURCE)
        self.assertIn('If CDbl(v) < 61# Then GoTo InvalidTransaction', SOURCE)
        self.assertIn('sourceWs.Cells(rowNo, "T").Value2', SOURCE)
        self.assertIn('sourceWs.Cells(rowNo, "V").Value2', SOURCE)
        self.assertIn('MMF口座_イベント整列 d, acct, sourceRows, 1, n', SOURCE)
        self.assertIn('If name1 = "" Or name1 <> name2', SOURCE)

    def test_summary_prints_with_selected_paper_and_headers(self):
        self.assertIn(".PaperSize = xlPaperA4", SOURCE)
        self.assertIn(".Orientation = xlLandscape", SOURCE)
        self.assertIn('.PrintArea = "$A$1:$H$" & lastOut', SOURCE)
        self.assertIn('.PrintTitleRows = "$1:$5"', SOURCE)
        self.assertIn('Application.Dialogs(xlDialogPrint).Show', SOURCE)
        self.assertIn('Public Sub MMF口座_まとめ表印刷()', SOURCE)
        self.assertIn('If Not MMF口座_まとめ更新(False) Then Cancel = True', WB)
        self.assertIn('MMF口座_ボタン out, out.Range("J9:K10")', SOURCE)

    def test_a4_detail_print_button_routes_by_ledger(self):
        self.assertIn('If returnWs.Name = "一般口座" Or returnWs.Name = "特定口座"', PRINT)
        self.assertIn('printWs.Range("O2").Value2 = returnWs.Name', PRINT)
        self.assertIn('For Each nameValue In Array("一般口座", "特定口座")', PRINT)
        self.assertIn('If ledger.Visible = xlSheetVisible Then', PRINT)

    def test_safe_external_import_and_no_overwrite(self):
        for term in (
            "Public Sub MMF口座_別ブックから取込()",
            "If Application.CountA(targetWs.Range(",
            "に既存取引があります。上書きはしません。",
            "期首残高があります。上書きせず中止しました。",
            "Application.AutomationSecurity = 3",
            "Set sourceBook = Application.Workbooks.Open(",
            "ReadOnly:=True",
            "If dataWritten Then",
            "If appStateSaved Then",
            "sourceBook.Close SaveChanges:=False",
        ):
            with self.subTest(term=term):
                self.assertIn(term, SOURCE)

    def test_summary_sample_2023_to_2025_reference_totals(self):
        # Reference example on page 1: separate account movements may be
        # positive and negative, while total holdings add both accounts.
        # These are arithmetic checks for the sample, not a claim that
        # transfers are already tax-neutral ledger transactions.
        specified = 18_599_890
        general = 0
        specified -= 18_599_890
        general += 18_599_890
        self.assertEqual(specified + general, 18_599_890)
        general += 257_812  # 2019
        general += 483_335  # 2020
        general += 1_942_685  # 2021
        general += 602_252  # 2022
        general += 2_198_228  # 2023, before the account policy change
        self.assertEqual(specified + general, 24_084_202)
        specified += 106_435 + 381
        general += 232_814
        self.assertEqual(specified + general, 24_423_832)
        specified += 316_994 + 1_669
        general += 202_392
        self.assertEqual(specified + general, 24_944_887)
        general -= 9_280_109
        self.assertEqual(specified + general, 15_664_778)
        specified += 1_588_780 + 37_059
        general += 426_166
        self.assertEqual(specified + general, 17_716_783)
        specified += 947_448 + 78_067
        general += 464_484
        self.assertEqual(specified, 3_076_833)
        self.assertEqual(general, 16_129_949)
        self.assertEqual(specified + general, 19_206_782)
        self.assertEqual(round((specified + general) * 0.01), 192_068)

    def test_repeated_click_does_not_duplicate_buttons(self):
        self.assertIn("Set s = ws.Shapes(shapeName)", SOURCE)
        self.assertIn("If s Is Nothing Then", SOURCE)
        self.assertIn('MMF口座_ボタン out, out.Range("J6:K7")', SOURCE)
        self.assertIn('MMF口座_ボタン out, out.Range("J12:K13")', SOURCE)


if __name__ == "__main__":
    unittest.main()
