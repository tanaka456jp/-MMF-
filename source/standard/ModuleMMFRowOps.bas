Attribute VB_Name = "ModuleMMFRowOps"
Option Explicit

' MMF台帳の行削除・行挿入は行全体を操作しない。
' A:Jの入力欄だけを移動させ、K:Yの行連動計算式を保持する。
' Excelの通常の行削除で計算式が壊れた場合は原本を基に再構築する。
Private Const MMF_FIRST As Long = 12
Private Const MMF_BASE_LAST As Long = 511
Private Const MMF_MAX_LAST As Long = 5011
Private Const MMF_PASS As String = "1234"
Private mRowHandling As Boolean

Public Function MMF台帳_対象(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Not (ws.Parent Is ThisWorkbook) Then Exit Function
    Select Case ws.Name
        Case "MMF取引台帳", "一般口座", "特定口座"
            MMF台帳_対象 = True
    End Select
End Function

' F7専用: 選択行の入力データを削除し、その下の入力だけを上へ詰める。
Public Sub MMF台帳_選択行削除()
    Call MMF台帳_選択行編集(False)
End Sub

' F8専用: 選択行の直後に空の入力行を作る（コピーによる取引重複を防ぐ）。
Public Sub MMF台帳_選択行追加()
    Call MMF台帳_選択行編集(True)
End Sub

Private Sub MMF台帳_選択行編集(ByVal addMode As Boolean)
    Dim ws As Worksheet, sel As Range
    Dim rowFlags() As Boolean, selectedRows() As Long
    Dim area As Range, n As Long, i As Long, rowNumber As Long
    Dim lastRow As Long, maxData As Long, newBottom As Long
    Dim shiftStart As Long
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalc As XlCalculation
    Dim wasProtected As Boolean, stateTaken As Boolean, success As Boolean
    Dim actionName As String, errDetail As String
    Dim inputsBefore As Variant, dataChanged As Boolean

    If mRowHandling Then Exit Sub
    On Error GoTo Failed
    Set ws = ActiveSheet
    If Not MMF台帳_対象(ws) Then
        Err.Raise vbObjectError + 911, , "MMF取引台帳・一般口座・特定口座だけが対象です。"
    End If
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection

    maxData = MMF台帳_最終入力行(ws)
    lastRow = Application.Max(MMF_BASE_LAST, maxData)
    If lastRow > MMF_MAX_LAST Then
        Err.Raise vbObjectError + 912, , "台帳が上限を超えています。"
    End If
    ReDim rowFlags(MMF_FIRST To lastRow)

    For Each area In sel.Areas
        If area.Rows.CountLarge > MMF_MAX_LAST Then
            Err.Raise vbObjectError + 913, , "選択行数が多すぎます。"
        End If
        For i = 1 To area.Rows.Count
            rowNumber = area.Row + i - 1
            If rowNumber < MMF_FIRST Or rowNumber > lastRow Then
                Err.Raise vbObjectError + 914, , _
                    "見出し行・台帳範囲外の行は操作できません（12行目以降）。"
            End If
            If Not rowFlags(rowNumber) Then
                rowFlags(rowNumber) = True
                n = n + 1
            End If
        Next i
    Next area
    If n = 0 Then Exit Sub

    ReDim selectedRows(1 To n)
    n = 0
    For i = lastRow To MMF_FIRST Step -1
        If rowFlags(i) Then
            n = n + 1
            selectedRows(n) = i
        End If
    Next i
    If addMode Then
        actionName = "追加"
    Else
        actionName = "削除"
    End If

    If addMode Then
        If selectedRows(1) >= lastRow Then
            Err.Raise vbObjectError + 915, , _
                "最終行の後には空行を追加できません。"
        End If
        newBottom = lastRow
        ' 末尾に既存のデータがある場合は絶対に上書きしない。
        If Application.CountA(ws.Range("A" & (lastRow - n + 1) & ":J" & lastRow)) > 0 Then
            Err.Raise vbObjectError + 916, , _
                "末尾の空行が不足しています。入力を失う可能性があるため追加を中止しました。"
        End If
    End If

    If MsgBox(CStr(n) & "行の取引データを" & actionName & "します。" & vbCrLf & _
              "A:Jの入力欄だけを移動し、K:Yの取得費計算式は削除しません。" & vbCrLf & _
              "このマクロ操作はExcelの元に戻す機能を使用できません。" & vbCrLf & _
              "続行しますか？", vbYesNo + vbQuestion + vbDefaultButton2, _
              "外貨建MMF 安全な行" & actionName) <> vbYes Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    stateTaken = True
    mRowHandling = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect Password:=MMF_PASS
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 917, , "台帳の保護を解除できませんでした。"
    End If

    ' 参照式の破損がある場合はデータ移動前に先に修復する。
    If Not MMF台帳_数式復元(ws, True) Then
        Err.Raise vbObjectError + 919, , _
                  "正常な計算式のひな形が見つかりません。行操作を中止しました。"
    End If
    ' 操作途中で失敗した場合に元入力を戻す（Excelの通常のUndoは使用できない）。
    inputsBefore = ws.Range("A12:J" & lastRow).FormulaR1C1
    dataChanged = True

    For i = 1 To n
        rowNumber = selectedRows(i)
        If addMode Then
            shiftStart = rowNumber + 1
            If shiftStart < lastRow Then
                ws.Range("A" & (shiftStart + 1) & ":J" & lastRow).FormulaR1C1 = _
                    ws.Range("A" & shiftStart & ":J" & (lastRow - 1)).FormulaR1C1
            End If
            ws.Range("A" & shiftStart & ":J" & shiftStart).ClearContents
        Else
            If rowNumber < lastRow Then
                ws.Range("A" & rowNumber & ":J" & (lastRow - 1)).FormulaR1C1 = _
                    ws.Range("A" & (rowNumber + 1) & ":J" & lastRow).FormulaR1C1
            End If
            ws.Range("A" & lastRow & ":J" & lastRow).ClearContents
        End If
    Next i

    If Not MMF台帳_数式復元(ws, True) Then
        Err.Raise vbObjectError + 918, , _
            "移動平均計算式を復元できません。直ちに元ブックのバックアップから確認してください。"
    End If
    ws.Calculate
    success = True

CleanExit:
    On Error Resume Next
    If wasProtected Then
        ws.Protect Password:=MMF_PASS, UserInterfaceOnly:=True, _
                   AllowInsertingRows:=False, AllowDeletingRows:=False
        ws.EnableSelection = xlNoRestrictions
    End If
    If stateTaken Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
        Application.EnableEvents = oldEvents
    End If
    mRowHandling = False
    On Error GoTo 0
    If success Then
        ' まとめ表の更新は任意。異常な別口座の取引があっても行編集を巻き戻さない。
        Call MMF口座_まとめ更新(True)
        ' まとめ表の再集計後も操作元の台帳に戻す。
        ws.Activate
        If addMode Then
            ws.Cells(selectedRows(n) + 1, 1).Select
        Else
            ws.Cells(selectedRows(n), 1).Select
        End If
        MsgBox "安全な行" & actionName & "を実行しました。" & vbCrLf & _
               "入力データと移動平均取得費を再計算しました。", vbInformation, "外貨建MMF"
    End If
    Exit Sub

Failed:
    errDetail = Err.Description
    If dataChanged Then
        On Error Resume Next
        Application.EnableEvents = False
        ws.Range("A12:J" & lastRow).FormulaR1C1 = inputsBefore
        Call MMF台帳_数式復元(ws, True)
        On Error GoTo 0
    End If
    MsgBox "MMFの行" & actionName & "を実行できませんでした。" & vbCrLf & _
           errDetail, vbExclamation, "外貨建MMF"
    Resume CleanExit
End Sub

' ブック起動時に口座別台帳の式を一括点検し、#REF!や行参照ずれを修復する。
Public Sub MMF台帳_起動時点検()
    Dim names As Variant, item As Variant, ws As Worksheet
    Dim fixedCount As Long

    names = Array("一般口座", "特定口座")
    For Each item In names
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(item))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If MMF台帳_式ずれ検出(ws) Then
                If MMF台帳_数式復元(ws, False) Then fixedCount = fixedCount + 1
            End If
        End If
    Next item

    If fixedCount > 0 Then
        Call MMF口座_まとめ更新(True)
        MsgBox CStr(fixedCount) & "口座の計算式を再構築しました。" & vbCrLf & _
               "行削除後の参照ずれを修復しています。", _
               vbInformation, "外貨建MMF"
    End If
End Sub

Private Function MMF台帳_式ずれ検出(ByVal ws As Worksheet) As Boolean
    Dim source As Worksheet
    Dim originFirst As Variant, originOther As Variant
    Dim current As Variant
    Dim r As Long, c As Long, lastRow As Long
    Dim expectFormula As String
    Dim cellValue As Variant

    On Error GoTo NeedsAttention
    Set source = ThisWorkbook.Worksheets("MMF取引台帳")
    If Not MMF台帳_ひな形有効(source) Then GoTo NeedsAttention

    originFirst = source.Range("K12:Y12").FormulaR1C1
    originOther = source.Range("K13:Y13").FormulaR1C1
    lastRow = MMF台帳_最終入力行(ws)
    If lastRow > MMF_MAX_LAST Then GoTo NeedsAttention

    current = ws.Range("K12:Y" & lastRow).FormulaR1C1
    For r = 1 To lastRow - MMF_FIRST + 1
        For c = 1 To 15
            cellValue = current(r, c)
            If IsError(cellValue) Then GoTo NeedsAttention
            If r = 1 Then
                expectFormula = CStr(originFirst(1, c))
            Else
                expectFormula = CStr(originOther(1, c))
            End If
            If StrComp(CStr(cellValue), expectFormula, vbBinaryCompare) <> 0 Then
                GoTo NeedsAttention
            End If
        Next c
    Next r
    Exit Function

NeedsAttention:
    MMF台帳_式ずれ検出 = True
End Function

' Excel標準の行削除・行挿入で K:Yの計算式が壊れた場合に手動修復する。
' 数式だけを書き換え、A:Jの入力値は絶対に変更しない。
Public Sub MMF台帳_計算式修復()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not MMF台帳_対象(ws) Then
        MsgBox "MMF取引台帳・一般口座・特定口座を開いてください。", vbExclamation
        Exit Sub
    End If
    If MMF台帳_数式復元(ws, False) Then
        MsgBox "K:Yの計算式を修復しました。" & vbCrLf & _
               "A:Jの入力データは変更していません。", vbInformation, "外貨建MMF"
    End If
End Sub

' Workbook_SheetChange で使用。通常の入力変更では重い処理を実行しない。
Public Sub MMF台帳_構造変更時修復(ByVal sh As Object, ByVal changed As Range)
    Dim ws As Worksheet, hit As Range
    If mRowHandling Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    Set ws = sh
    If Not MMF台帳_対象(ws) Then Exit Sub
    If changed.Row > MMF_MAX_LAST Then Exit Sub
    ' 列全体を削除／一部セルをシフトした場合も数式欠損を検出する。
    Set hit = Application.Intersect(changed, ws.Range("K12:Y" & MMF_MAX_LAST))
    If hit Is Nothing Then Exit Sub
    ' 標準の「シート行全体の削除」で式列まで移動した場合の救済。
    Call MMF台帳_数式復元(ws, False)
End Sub

Private Function MMF台帳_最終入力行(ByVal ws As Worksheet) As Long
    Dim col As Long, rr As Long, lastRow As Long
    lastRow = MMF_BASE_LAST
    For col = 1 To 10
        rr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If rr > lastRow Then lastRow = rr
    Next col
    MMF台帳_最終入力行 = lastRow
End Function

' 元台帳（移行前の原本）からR1C1形式の計算式を復元する。
' 先頭行と2行目以降の式は異なるため必ず別のひな形を使用する。
Private Function MMF台帳_数式復元(ByVal target As Worksheet, _
                                 ByVal silent As Boolean) As Boolean
    Dim master As Worksheet, candidate As Worksheet
    Dim names As Variant, item As Variant
    Dim firstFormula As Variant, nextFormula As Variant
    Dim col As Long, countCols As Long, lastRow As Long
    Dim wasEvents As Boolean, wasScreen As Boolean
    Dim wasProtected As Boolean, captured As Boolean
    Dim valid As Boolean, msg As String

    If mRowHandling And Not silent Then Exit Function
    On Error GoTo Failed
    names = Array("MMF取引台帳", "一般口座", "特定口座")
    For Each item In names
        Set candidate = Nothing
        On Error Resume Next
        Set candidate = ThisWorkbook.Worksheets(CStr(item))
        On Error GoTo Failed
        If Not candidate Is Nothing Then
            If Not candidate Is target Then
                valid = MMF台帳_ひな形有効(candidate)
                If valid Then
                    Set master = candidate
                    Exit For
                End If
            End If
        End If
    Next item
    If master Is Nothing Then
        If MMF台帳_ひな形有効(target) Then Set master = target
    End If
    If master Is Nothing Then
        Err.Raise vbObjectError + 921, , _
                  "修復元となる正常な計算式が見つかりません。"
    End If

    firstFormula = master.Range("K12:Y12").FormulaR1C1
    nextFormula = master.Range("K13:Y13").FormulaR1C1
    lastRow = MMF台帳_最終入力行(target)
    If lastRow > MMF_MAX_LAST Then
        Err.Raise vbObjectError + 922, , "取引行数が上限を超えています。"
    End If

    wasEvents = Application.EnableEvents
    wasScreen = Application.ScreenUpdating
    captured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    wasProtected = target.ProtectContents
    If wasProtected Then target.Unprotect Password:=MMF_PASS
    If target.ProtectContents Then
        Err.Raise vbObjectError + 923, , "シート保護を解除できません。"
    End If

    For col = 11 To 25
        countCols = col - 10
        target.Cells(12, col).FormulaR1C1 = CStr(firstFormula(1, countCols))
        target.Range(target.Cells(13, col), target.Cells(lastRow, col)).FormulaR1C1 = _
            CStr(nextFormula(1, countCols))
    Next col

    target.Calculate
    MMF台帳_数式復元 = True

CleanExit:
    On Error Resume Next
    If wasProtected Then
        target.Protect Password:=MMF_PASS, UserInterfaceOnly:=True, _
                   AllowInsertingRows:=False, AllowDeletingRows:=False
    End If
    If captured Then
        Application.ScreenUpdating = wasScreen
        Application.EnableEvents = wasEvents
    End If
    On Error GoTo 0
    Exit Function

Failed:
    msg = Err.Description
    If Not silent Then
        MsgBox "計算式を修復できませんでした。" & vbCrLf & msg, _
               vbExclamation, "外貨建MMF"
    End If
    Resume CleanExit
End Function

Private Function MMF台帳_ひな形有効(ByVal ws As Worksheet) As Boolean
    Dim col As Long, r As Long, f As String
    On Error GoTo Invalid
    If CStr(ws.Range("K11").Value2) <> "入力判定" Then Exit Function
    For r = 12 To 13
        For col = 11 To 25
            If Not ws.Cells(r, col).HasFormula Then Exit Function
            f = CStr(ws.Cells(r, col).FormulaR1C1)
            If InStr(1, f, "#REF!", vbTextCompare) > 0 Then Exit Function
        Next col
    Next r
    MMF台帳_ひな形有効 = True
Invalid:
End Function


' 操作バー: 既存の見出し(10～11行)・入力(12行以降)は変更しない。
' 9行目の空欄を使い、従来のI3:L4のA4印刷ボタンは維持する。
Public Sub MMF操作_ボタン設置()
    Dim ws As Worksheet
    Dim problem As String, detail As String
    For Each ws In ThisWorkbook.Worksheets
        If MMF台帳_対象(ws) And ws.Visible = xlSheetVisible Then
            detail = MMF操作_台帳バー(ws)
            If Len(detail) > 0 Then problem = problem & vbCrLf & ws.Name & "：" & detail
        End If
    Next ws
    If Len(problem) > 0 Then
        MsgBox "一部の操作ボタンを配置できませんでした。" & problem, _
               vbExclamation, "外貨建MMF"
    End If
End Sub

Private Function MMF操作_台帳バー(ByVal ws As Worksheet) As String
    Dim wasProtected As Boolean, didUnprotect As Boolean
    Dim drawProtected As Boolean, contentsProtected As Boolean, scenarioProtected As Boolean
    Dim allowFmtCells As Boolean, allowFmtCols As Boolean, allowFmtRows As Boolean
    Dim allowInsCols As Boolean, allowInsRows As Boolean, allowLinks As Boolean
    Dim allowDelCols As Boolean, allowDelRows As Boolean
    Dim allowSort As Boolean, allowFilter As Boolean, allowPivot As Boolean
    Dim selectionMode As XlEnableSelection, messageText As String
    On Error GoTo Failed

    ' 9行目に独自の情報がある場合は上書きや重ね表示をしない。
    If Application.CountA(ws.Range("A9:L9")) > 0 Then
        MMF操作_台帳バー = "9行目に既存データがあるため配置を中止しました。"
        Exit Function
    End If
    If CStr(ws.Range("C11").Value2) <> "買付数量" Or _
       CStr(ws.Range("K11").Value2) <> "入力判定" Then
        MMF操作_台帳バー = "台帳の見出しが想定と異なります。"
        Exit Function
    End If

    drawProtected = ws.ProtectDrawingObjects
    contentsProtected = ws.ProtectContents
    scenarioProtected = ws.ProtectScenarios
    wasProtected = drawProtected Or contentsProtected Or scenarioProtected
    If wasProtected Then
        selectionMode = ws.EnableSelection
        With ws.Protection
            allowFmtCells = .AllowFormattingCells
            allowFmtCols = .AllowFormattingColumns
            allowFmtRows = .AllowFormattingRows
            allowInsCols = .AllowInsertingColumns
            allowInsRows = .AllowInsertingRows
            allowLinks = .AllowInsertingHyperlinks
            allowDelCols = .AllowDeletingColumns
            allowDelRows = .AllowDeletingRows
            allowSort = .AllowSorting
            allowFilter = .AllowFiltering
            allowPivot = .AllowUsingPivotTables
        End With
        ws.Unprotect Password:=MMF_PASS
        didUnprotect = True
    End If

    If ws.Rows(9).RowHeight < 30 Then ws.Rows(9).RowHeight = 30
    MMF操作_図形 ws, ws.Range("A9:B9"), "MMF_UI_Delete", _
        "行削除 [F7]", "MMF操作_行削除", RGB(155, 64, 51)
    MMF操作_図形 ws, ws.Range("C9:D9"), "MMF_UI_Add", _
        "行追加 [F8]", "MMF操作_行追加", RGB(31, 90, 130)
    MMF操作_図形 ws, ws.Range("E9:F9"), "MMF_UI_Repair", _
        "計算式修復", "MMF台帳_計算式修復", RGB(31, 90, 130)
    MMF操作_図形 ws, ws.Range("G9:H9"), "MMF_UI_Summary", _
        "まとめ表更新", "MMF口座_まとめ表更新", RGB(24, 113, 103)
    If ws.Name = "MMF取引台帳" Then
        MMF操作_図形 ws, ws.Range("I9:J9"), "MMF_UI_Import", _
            "2口座作成", "MMF操作_口座作成", RGB(31, 90, 130)
    Else
        MMF操作_図形 ws, ws.Range("I9:J9"), "MMF_UI_Import", _
            "別ブック取込", "MMF口座_別ブックから取込", RGB(31, 90, 130)
    End If
    MMF操作_図形 ws, ws.Range("K9:L9"), "MMF_UI_Help", _
        "操作ガイド", "MMF操作_ガイド", RGB(75, 88, 102)

CleanExit:
    If didUnprotect Then
        On Error Resume Next
        Err.Clear
        ws.Protect Password:=MMF_PASS, DrawingObjects:=drawProtected, _
            Contents:=contentsProtected, Scenarios:=scenarioProtected, _
            UserInterfaceOnly:=True, AllowFormattingCells:=allowFmtCells, _
            AllowFormattingColumns:=allowFmtCols, AllowFormattingRows:=allowFmtRows, _
            AllowInsertingColumns:=allowInsCols, AllowInsertingRows:=allowInsRows, _
            AllowInsertingHyperlinks:=allowLinks, AllowDeletingColumns:=allowDelCols, _
            AllowDeletingRows:=allowDelRows, AllowSorting:=allowSort, _
            AllowFiltering:=allowFilter, AllowUsingPivotTables:=allowPivot
        If Err.Number <> 0 Then
            messageText = messageText & "／シート保護の復元失敗：" & Err.Description
        Else
            ws.EnableSelection = selectionMode
        End If
        On Error GoTo 0
    End If
    MMF操作_台帳バー = messageText
    Exit Function
Failed:
    messageText = Err.Description
    Resume CleanExit
End Function

' 取引行の選択が失われた場合は誤って別の行を操作しない。
Public Sub MMF操作_行削除()
    If Not MMF操作_選択確認() Then Exit Sub
    Call MMF台帳_選択行削除
End Sub

Public Sub MMF操作_行追加()
    If Not MMF操作_選択確認() Then Exit Sub
    Call MMF台帳_選択行追加
End Sub

Private Function MMF操作_選択確認() As Boolean
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    If Not MMF台帳_対象(ActiveSheet) Then Exit Function
    If TypeName(Selection) = "Range" Then
        If Selection.Row >= MMF_FIRST Then
            MMF操作_選択確認 = True
            Exit Function
        End If
    End If
    MsgBox "先に12行目以降の取引セルを選択してください。" & vbCrLf & _
           "行削除はF7、行追加はF8でも実行できます。", vbInformation, "外貨建MMF"
End Function

Public Sub MMF操作_口座作成()
    Call MMF口座_2口座シート作成
    Call MMF印刷_ボタン設置
    Call MMF操作_ボタン設置
End Sub

Public Sub MMF操作_ガイド()
    MsgBox "【操作例】" & vbCrLf & _
           "行削除：対象行のセルを選択 → 行削除 [F7] → 確認" & vbCrLf & _
           "行追加：挿入したい位置の直前の行を選択 → 行追加 [F8]" & vbCrLf & _
           "複数行：Ctrlキーで離れた行も選択できます。" & vbCrLf & _
           "計算式修復：対象口座を開いてボタンを押す（入力値は保持）。" & vbCrLf & _
           "まとめ表更新：両口座を集計し、まとめ表を表示。" & vbCrLf & _
           "別ブック取込：空の口座へ既存ファイルを取り込む。" & vbCrLf & _
           "A4印刷：現在の口座の取引明細を印刷。" & vbCrLf & _
           "まとめ表ではA4まとめ印刷を使用してください。" & vbCrLf & _
           "※標準の行削除ではなくF7または行削除ボタンを推奨。", _
           vbInformation, "外貨建MMF 操作ガイド"
End Sub

' 印刷シートの設定欄に既存マクロを直接呼び出すボタンを配置。
' A:Lの印刷範囲に重ならない。印刷再作成時にも呼び出される。
Public Sub MMF操作_印刷補助設置(ByVal ws As Worksheet)
    MMF操作_図形 ws, ws.Range("N11:O12"), "MMF_UI_PrintRefresh", _
        "印刷明細更新", "MMF印刷_1段表示へ更新", RGB(31, 90, 130)
    MMF操作_図形 ws, ws.Range("N13:O14"), "MMF_UI_PrintPreview", _
        "指定ページプレビュー", "MMF印刷_指定頁プレビュー", RGB(31, 90, 130)
End Sub

Private Sub MMF操作_図形(ByVal ws As Worksheet, ByVal place As Range, _
                         ByVal shapeId As String, ByVal buttonText As String, _
                         ByVal macroName As String, ByVal fillColor As Long)
    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes(shapeId)
    On Error GoTo 0
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            place.Left + 2, place.Top + 2, place.Width - 4, place.Height - 4)
        btn.Name = shapeId
    End If
    With btn
        .Left = place.Left + 2
        .Top = place.Top + 2
        .Width = place.Width - 4
        .Height = place.Height - 4
        .Placement = xlMove
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
        .AlternativeText = buttonText
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = buttonText
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Name = "Yu Gothic"
            .Size = 11
            .Bold = True
            .Color = vbWhite
        End With
    End With
End Sub
