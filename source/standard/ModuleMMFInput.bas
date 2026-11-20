Attribute VB_Name = "ModuleMMFInput"
Option Explicit

' 取引区分はリスト選択に制限せず、利用者が自由に入力できる。
' 既存の K 列「入力判定」および M:X 列の取得費計算には変更を加えない。
Private Const MMF_INPUT_SHEET As String = "MMF取引台帳"
Private Const FIRST_TRADE_ROW As Long = 12
Private Const BASE_LAST_TRADE_ROW As Long = 511
Private Const EXTRA_INPUT_ROWS As Long = 500

' VBAUpdater反映直後、ブックを開き直さずに実行する場合の手動マクロ。
Public Sub MMF取引区分_自由入力設定()
    If MMF取引区分_設定本体() Then
        MsgBox "取引区分を自由入力に変更しました。" & vbCrLf & _
               "B列の選択リスト制限は解除済みです。" & vbCrLf & _
               "取得費計算は買付数量(C列)・売付数量(E列)から判定します。" & vbCrLf & _
               "取引区分の文字列は計算に使用しません。", _
               vbInformation, "外貨建MMF"
    End If
End Sub

' ThisWorkbook.Workbook_Open から呼び出す。起動時は成功ダイアログを出さない。
Public Sub MMF取引区分_起動時設定()
    Call MMF取引区分_設定本体
End Sub

Private Function MMF取引区分_設定本体() As Boolean
    Dim ws As Worksheet
    Dim lastInputRow As Long
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim errMsg As String
    Dim eventsCaptured As Boolean

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    eventsCaptured = True

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(MMF_INPUT_SHEET)

    If CStr(ws.Range("B11").Value2) <> "取引区分" Then
        Err.Raise vbObjectError + 521, , _
                  "「" & MMF_INPUT_SHEET & "」の B11 が取引区分ではありません。"
    End If
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 522, , _
                  "台帳が保護されているため設定できません。保護解除後に再実行してください。"
    End If

    ' 既存500件に加え、既存入力の最終行＋余裕行を対象とする。
    lastInputRow = Application.Max(BASE_LAST_TRADE_ROW, _
                   ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
                   ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
                   ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)
    lastInputRow = Application.Min(ws.Rows.Count, lastInputRow + EXTRA_INPUT_ROWS)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    With ws.Range("B" & FIRST_TRADE_ROW & ":B" & lastInputRow)
        ' 取引区分の入力規則だけを削除（為替区分 H列、通貨 F3 は維持）。
        .Validation.Delete
        ' 「001」「買付」等の文字入力をExcelが数値や日付へ自動変換しない。
        .NumberFormat = "@"
    End With

    MMF取引区分_設定本体 = True

CleanExit:
    On Error Resume Next
    If eventsCaptured Then
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
        Application.DisplayAlerts = oldAlerts
    End If
    On Error GoTo 0
    Exit Function

Failed:
    errMsg = Err.Description
    MsgBox "取引区分の自由入力設定に失敗しました。" & vbCrLf & errMsg, _
           vbExclamation, "外貨建MMF"
    Resume CleanExit
End Function
