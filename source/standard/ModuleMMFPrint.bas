Attribute VB_Name = "ModuleMMFPrint"
Option Explicit

' A4縦 / 1取引1行の印刷専用ビュー。取得費等の計算は一切変更しない。
' 列構成は対象のMMF取引台帳（A:J入力、K入力判定、L:V計算）に対応。
Private Const PRINT_SHEET As String = "A4縦印刷"
Private Const DEFAULT_SOURCE As String = "MMF取引台帳"
Private Const FIRST_SOURCE_ROW As Long = 12
Private Const STANDARD_LAST_ROW As Long = 511
Private Const FIRST_PRINT_ROW As Long = 6
Private Const PRINT_COLUMNS As Long = 12
Private Const MAX_PRINT_ROWS As Long = 5000
Private Const PRINT_PASSWORD As String = "1234"
Private mUpdating As Boolean

' VBAUpdater反映後は、このマクロを一度実行する。
Public Sub MMF印刷_1段表示へ更新()
    If MMF印刷_更新本体(False) Then
        MsgBox "1取引1行・A4縦の印刷明細を更新しました。" & vbCrLf & _
               "以後はこのシートの印刷・プレビュー直前にも自動更新されます。", vbInformation
    End If
End Sub

' このマクロから印刷プレビューすると最新の明細が表示される。
Public Sub MMF印刷_1段表示プレビュー()
    If MMF印刷_更新本体(False) Then
        ThisWorkbook.Worksheets(PRINT_SHEET).PrintPreview
    End If
End Sub

' ThisWorkbook.Workbook_BeforePrint から使用する。
' 戻り値Falseの場合は印刷そのものを中止する。
Public Function MMF印刷_印刷直前更新() As Boolean
    If mUpdating Then
        MMF印刷_印刷直前更新 = True
        Exit Function
    End If
    MMF印刷_印刷直前更新 = MMF印刷_更新本体(True)
End Function

Private Function MMF印刷_更新本体(ByVal quiet As Boolean) As Boolean
    Dim wb As Workbook
    Dim src As Worksheet, out As Worksheet
    Dim sourceName As String
    Dim data() As Variant
    Dim heads As Variant, widths As Variant
    Dim rowNo As Long, lastSource As Long, lastOutput As Long
    Dim col As Long, idx As Long, n As Long, bad As Long
    Dim oldEvents As Boolean, oldScreen As Boolean
    Dim stateCaptured As Boolean, wasProtected As Boolean
    Dim warnRule As FormatCondition, stripeRule As FormatCondition
    Dim failureText As String
    Dim success As Boolean

    If mUpdating Then Exit Function
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    stateCaptured = True
    On Error GoTo Failed

    mUpdating = True
    Set wb = ThisWorkbook
    On Error Resume Next
    Set out = wb.Worksheets(PRINT_SHEET)
    On Error GoTo Failed

    sourceName = DEFAULT_SOURCE
    If Not out Is Nothing Then
        If Len(Trim$(CStr(out.Range("O2").Value2))) > 0 Then
            sourceName = Trim$(CStr(out.Range("O2").Value2))
        ElseIf Len(Trim$(CStr(out.Range("R2").Value2))) > 0 Then
            sourceName = Trim$(CStr(out.Range("R2").Value2))
        ElseIf Len(Trim$(CStr(out.Range("J2").Value2))) > 0 Then
            sourceName = Trim$(CStr(out.Range("J2").Value2))
        End If
    End If
    If sourceName = PRINT_SHEET Then
        Err.Raise vbObjectError + 420, , "印刷用シート自身は対象に指定できません。"
    End If

    On Error Resume Next
    Set src = wb.Worksheets(sourceName)
    On Error GoTo Failed
    If src Is Nothing Then
        Err.Raise vbObjectError + 421, , "対象台帳「" & sourceName & "」が見つかりません。"
    End If
    If CStr(src.Range("A11").Value2) <> "日付" _
       Or CStr(src.Range("B11").Value2) <> "取引区分" _
       Or CStr(src.Range("K11").Value2) <> "入力判定" _
       Or CStr(src.Range("T11").Value2) <> "残高数量" _
       Or CStr(src.Range("V11").Value2) <> "残高・円" Then
        Err.Raise vbObjectError + 422, , _
                  "台帳の項目構成が想定と異なります。対象シート名と11行目の項目をご確認ください。"
    End If

    ' 日付欠落でも他の入力が存在する取引は除外しない。
    ' 511行を超えて追加した明細も範囲判定に含める。
    lastSource = STANDARD_LAST_ROW
    For col = 1 To 10
        rowNo = src.Cells(src.Rows.Count, col).End(xlUp).Row
        If rowNo > lastSource Then lastSource = rowNo
    Next col
    If lastSource - FIRST_SOURCE_ROW + 1 > MAX_PRINT_ROWS Then
        Err.Raise vbObjectError + 423, , _
                  "取引行が上限を超えています。空白以外の入力位置を確認してください。"
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    src.Calculate

    For rowNo = FIRST_SOURCE_ROW To lastSource
        If Application.CountA(src.Range("A" & rowNo & ":J" & rowNo)) > 0 Then
            n = n + 1
            If CStr(src.Cells(rowNo, "K").Value2) <> "OK" Then bad = bad + 1
        End If
    Next rowNo

    If n > 0 Then
        ReDim data(1 To n, 1 To PRINT_COLUMNS)
        idx = 0
        For rowNo = FIRST_SOURCE_ROW To lastSource
            If Application.CountA(src.Range("A" & rowNo & ":J" & rowNo)) > 0 Then
                idx = idx + 1
                For col = 1 To PRINT_COLUMNS
                    data(idx, col) = MMF印刷_数式(sourceName, rowNo, col)
                Next col
            End If
        Next rowNo
    End If

    If out Is Nothing Then
        Set out = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        out.Name = PRINT_SHEET
    End If

    wasProtected = out.ProtectContents
    If wasProtected Then out.Unprotect Password:=PRINT_PASSWORD

    ' 旧2段レイアウトの結合・書式・計算式をリセットする。
    out.UsedRange.UnMerge
    out.UsedRange.Clear

    With out
        .Range("A1:L" & CStr(Application.Max(6, n + FIRST_PRINT_ROW - 1))).Font.Name = "Yu Gothic"
        .Range("A1:L" & CStr(Application.Max(6, n + FIRST_PRINT_ROW - 1))).Font.Size = 10
        .Range("A1:L" & CStr(Application.Max(6, n + FIRST_PRINT_ROW - 1))).Font.Color = RGB(25, 48, 66)
        .Range("A1:L" & CStr(Application.Max(6, n + FIRST_PRINT_ROW - 1))).VerticalAlignment = xlCenter
        .Range("A1:L" & CStr(Application.Max(6, n + FIRST_PRINT_ROW - 1))).ShrinkToFit = True

        ' 合計を約114文字幅に抑え、A4縦での縮小率を改善する。
        widths = Array(10, 5, 11, 11, 11, 11, 11, 11, 12, 11, 9, 14)
        For col = 1 To PRINT_COLUMNS
            .Columns(col).ColumnWidth = widths(col - 1)
        Next col
        .Columns("M").ColumnWidth = 2
        .Columns("N").ColumnWidth = 20
        .Columns("O").ColumnWidth = 28

        .Range("A1:L1").Merge
        .Range("A1").Value = "外貨建MMF  取得費計算明細｜A4縦・1取引1行"
        .Range("A1:L1").Interior.Color = RGB(23, 52, 72)
        .Range("A1:L1").Font.Color = vbWhite
        .Range("A1:L1").Font.Bold = True
        .Range("A1:L1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlLeft
        .Rows(1).RowHeight = 28

        .Range("A2:F2").Merge
        .Range("G2:J2").Merge
        .Range("K2:L2").Merge
        .Range("A2").Formula = "=""銘柄・口座：""&" & MMF印刷_参照(sourceName, "B", 3)
        .Range("G2").Formula = "=""通貨：""&" & MMF印刷_参照(sourceName, "F", 3)
        .Range("K2").Formula = "=""出力：""&TEXT(TODAY(),""yy/m/d"")"
        .Range("A2:L2").Interior.Color = RGB(219, 234, 243)
        .Range("A2:L2").Font.Color = RGB(23, 52, 72)
        .Range("A2:L2").Font.Bold = True
        .Range("A2:L2").HorizontalAlignment = xlLeft
        .Rows(2).RowHeight = 22

        .Range("A3:D3").Merge
        .Range("E3:H3").Merge
        .Range("I3:L3").Merge
        .Range("A3").Formula = "=""期首数量：""&TEXT(" & MMF印刷_参照(sourceName, "B", 4) & ",""#,##0.########"")"
        .Range("E3").Formula = "=""期首外貨：""&TEXT(" & MMF印刷_参照(sourceName, "D", 4) & ",""#,##0.########"")"
        .Range("I3").Formula = "=""期首円貨：""&TEXT(" & MMF印刷_参照(sourceName, "F", 4) & ",""#,##0"")"
        .Range("A3:L3").Interior.Color = RGB(238, 247, 251)
        .Range("A3:L3").HorizontalAlignment = xlLeft
        .Rows(3).RowHeight = 21

        .Range("A4:L4").Merge
        .Range("A4").Value = "売買額は外貨・円貨を併記。売却取得費・売却損益は円貨表示。残高は時価ではなく移動平均取得原価です。"
        .Range("A4:L4").Interior.Color = RGB(245, 248, 250)
        .Range("A4:L4").HorizontalAlignment = xlLeft
        .Range("A4:L4").Font.Size = 8.5
        .Rows(4).RowHeight = 19

        heads = Array("取引日", "区分", "売買数量", "売買額" & vbLf & "外貨", _
                      "売買額" & vbLf & "円貨", "売取得費" & vbLf & "円貨", _
                      "売損益" & vbLf & "円貨", "残数量", "残高" & vbLf & "外貨", _
                      "残高" & vbLf & "円貨", "換算" & vbLf & "レート", "備考・確認")
        For col = 1 To PRINT_COLUMNS
            .Cells(5, col).Value = heads(col - 1)
        Next col
        With .Range("A5:L5")
            .Interior.Color = RGB(31, 70, 95)
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .ShrinkToFit = True
        End With
        .Rows(5).RowHeight = 31

        .Range("N1:O1").Merge
        .Range("N1").Value = "印刷設定（用紙には出しません）"
        .Range("N1:O1").Interior.Color = RGB(31, 70, 95)
        .Range("N1:O1").Font.Color = vbWhite
        .Range("N2").Value = "参照台帳シート名"
        .Range("O2").Value = sourceName
        .Range("N3").Value = "印刷取引件数"
        .Range("O3").Value = n
        .Range("N4").Value = "入力要確認件数"
        .Range("O4").Value = bad
        .Range("N5").Value = "更新方法"
        .Range("O5").Value = "Alt+F8：MMF印刷_1段表示へ更新"
        .Range("O2").Interior.Color = RGB(255, 245, 211)
        .Range("O2").Font.Color = RGB(30, 88, 160)
        .Range("N2:N5").Interior.Color = RGB(235, 244, 249)
        .Range("N2:O5").HorizontalAlignment = xlLeft
        .Range("O5").ShrinkToFit = True

        lastOutput = FIRST_PRINT_ROW + n - 1
        If n > 0 Then
            ' n×12の正確なサイズで一括書込み（旧版は500×15のまま代入）。
            .Range("A" & FIRST_PRINT_ROW).Resize(n, PRINT_COLUMNS).Formula = data
            With .Range("A6:L" & lastOutput)
                .RowHeight = 21
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(222, 231, 238)
                .Borders.Weight = xlHairline
            End With
            With .Range("A6:L" & lastOutput).FormatConditions
                .Delete
                Set stripeRule = .Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0")
                stripeRule.Interior.Color = RGB(244, 249, 252)
                Set warnRule = .Add(Type:=xlExpression, Formula1:="=LEFT($L6,3)=""要確認""")
                warnRule.Interior.Color = RGB(255, 233, 226)
                warnRule.Font.Color = RGB(156, 37, 32)
                warnRule.SetFirstPriority
                warnRule.StopIfTrue = True
            End With
            .Range("A6:A" & lastOutput).NumberFormat = "yyyy/m/d"
            .Range("C6:D" & lastOutput).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("E6:G" & lastOutput).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("H6:I" & lastOutput).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("J6:J" & lastOutput).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("K6:K" & lastOutput).NumberFormat = "0.####;[Red](0.####);-"
            .Range("B6:B" & lastOutput).HorizontalAlignment = xlCenter
            .Range("L6:L" & lastOutput).HorizontalAlignment = xlLeft
            .Range("L6:L" & lastOutput).Font.Size = 8.5
        Else
            lastOutput = 6
            .Range("A6:L6").Merge
            .Range("A6").Value = "印刷対象の取引はありません。元台帳のA～J列を確認してください。"
            .Range("A6").HorizontalAlignment = xlLeft
            .Rows(6).RowHeight = 24
        End If

        .Activate
        If Not ActiveWindow Is Nothing Then ActiveWindow.DisplayGridlines = False
        With .PageSetup
            .PrintArea = "$A$1:$L$" & CStr(lastOutput)
            .PrintTitleRows = "$1:$5"
            .PaperSize = xlPaperA4
            .Orientation = xlPortrait
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftMargin = Application.CentimetersToPoints(0.35)
            .RightMargin = Application.CentimetersToPoints(0.35)
            .TopMargin = Application.CentimetersToPoints(0.55)
            .BottomMargin = Application.CentimetersToPoints(0.55)
            .HeaderMargin = Application.CentimetersToPoints(0.2)
            .FooterMargin = Application.CentimetersToPoints(0.25)
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintGridlines = False
            .CenterFooter = "外貨建MMF 取得費計算　&P / &N"
            .RightFooter = "&D"
        End With
    End With
    out.Calculate
    success = True
    MMF印刷_更新本体 = True

CleanUp:
    On Error Resume Next
    If wasProtected And Not out Is Nothing Then
        out.Range("O2").Locked = False
        out.Protect Password:=PRINT_PASSWORD, UserInterfaceOnly:=True
    End If
    If stateCaptured Then
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    mUpdating = False
    On Error GoTo 0

    If Not success Then
        MsgBox "印刷明細の更新に失敗しました。印刷を中止してください。" & vbCrLf & _
               failureText, vbExclamation
    ElseIf bad > 0 And Not quiet Then
        MsgBox "更新しました。ただし入力要確認の取引が " & bad & " 件あります。" & vbCrLf & _
               "赤色行と元台帳の入力判定を確認してください。", vbExclamation
    End If
    Exit Function

Failed:
    failureText = Err.Description
    MMF印刷_更新本体 = False
    Resume CleanUp
End Function

' 数式は元台帳の算定結果を参照するだけ。入力のある日付欠落行も「日付未入力」で出力する。
Private Function MMF印刷_参照(ByVal sheetName As String, ByVal col As String, ByVal rowNo As Long) As String
    MMF印刷_参照 = "'" & Replace(sheetName, "'", "''") & "'!" & col & CStr(rowNo)
End Function

Private Function MMF印刷_表示式(ByVal expression As String) As String
    Dim q As String
    q = Chr$(34)
    MMF印刷_表示式 = "=IFERROR(IF(" & expression & "=" & q & q & _
                     "," & q & q & "," & expression & ")," & q & q & ")"
End Function

Private Function MMF印刷_数式(ByVal sheetName As String, _
                                ByVal sourceRow As Long, ByVal colNo As Long) As String
    Dim q As String, b As String, k As String, expr As String
    q = Chr$(34)
    b = MMF印刷_参照(sheetName, "B", sourceRow)
    k = MMF印刷_参照(sheetName, "K", sourceRow)

    Select Case colNo
        Case 1
            expr = MMF印刷_参照(sheetName, "A", sourceRow)
            MMF印刷_数式 = "=IFERROR(IF(" & expr & "=" & q & q & _
                            "," & q & "日付未入力" & q & "," & expr & ")," & q & "日付要確認" & q & ")"
            Exit Function
        Case 2
            expr = b
        Case 3
            expr = "IF(" & b & "=" & q & "売" & q & "," & _
                   MMF印刷_参照(sheetName, "E", sourceRow) & "," & _
                   MMF印刷_参照(sheetName, "C", sourceRow) & ")"
        Case 4
            expr = "IF(" & b & "=" & q & "売" & q & "," & _
                   MMF印刷_参照(sheetName, "O", sourceRow) & "," & _
                   MMF印刷_参照(sheetName, "M", sourceRow) & ")"
        Case 5
            expr = "IF(" & b & "=" & q & "売" & q & "," & _
                   MMF印刷_参照(sheetName, "P", sourceRow) & "," & _
                   MMF印刷_参照(sheetName, "N", sourceRow) & ")"
        Case 6
            expr = "IF(" & b & "=" & q & "売" & q & "," & _
                   MMF印刷_参照(sheetName, "R", sourceRow) & "," & q & q & ")"
        Case 7
            expr = "IF(" & b & "=" & q & "売" & q & "," & _
                   MMF印刷_参照(sheetName, "S", sourceRow) & "," & q & q & ")"
        Case 8
            expr = MMF印刷_参照(sheetName, "T", sourceRow)
        Case 9
            expr = MMF印刷_参照(sheetName, "U", sourceRow)
        Case 10
            expr = MMF印刷_参照(sheetName, "V", sourceRow)
        Case 11
            expr = MMF印刷_参照(sheetName, "L", sourceRow)
        Case 12
            expr = "IF(" & k & "=" & q & "OK" & q & "," & _
                   MMF印刷_参照(sheetName, "J", sourceRow) & "," & _
                   q & "要確認：" & q & "&" & k & "&" & q & " " & q & _
                   "&" & MMF印刷_参照(sheetName, "J", sourceRow) & ")"
        Case Else
            Err.Raise vbObjectError + 424, , "未定義の印刷列：" & CStr(colNo)
    End Select
    MMF印刷_数式 = MMF印刷_表示式(expr)
End Function
