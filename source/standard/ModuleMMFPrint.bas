Attribute VB_Name = "ModuleMMFPrint"
Option Explicit

' 外貨建MMF: 1取引=1行 / A4縦 印刷用ビュー
' This module changes only the dedicated print worksheet.
' 取得費、円換算、移動平均、売却損益は「MMF取引台帳」を正本とする。
Private Const MMF_PRINT_SHEET As String = "A4縦印刷"
Private Const MMF_DEFAULT_SOURCE As String = "MMF取引台帳"
Private Const MMF_FIRST_SOURCE_ROW As Long = 12
Private Const MMF_LAST_SOURCE_ROW As Long = 511
Private Const MMF_FIRST_PRINT_ROW As Long = 7
Private Const MMF_PRINT_PASSWORD As String = "1234"
Private MMF_PrintLastSuccess As Boolean

' VBAUpdater適用後、Alt+F8から実行。
' 対象変更時は印刷用シート R2 に既存の台帳シート名を入力して再実行。
Public Sub MMF印刷_1段表示へ更新()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim sourceName As String
    Dim i As Long, j As Long, n As Long
    Dim rr As Long, lastOut As Long
    Dim data() As Variant
    Dim headers As Variant, widths As Variant
    Dim found As Range, fmt As FormatCondition
    Dim wasProtected As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean
    Dim complete As Boolean, failureMessage As String

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    MMF_PrintLastSuccess = False
    On Error GoTo Failure
    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsOut = wb.Worksheets(MMF_PRINT_SHEET)
    On Error GoTo Failure

    sourceName = MMF_DEFAULT_SOURCE
    If Not wsOut Is Nothing Then
        If Len(Trim$(CStr(wsOut.Range("R2").Value2))) > 0 Then
            sourceName = Trim$(CStr(wsOut.Range("R2").Value2))
        ElseIf Len(Trim$(CStr(wsOut.Range("J2").Value2))) > 0 Then
            sourceName = Trim$(CStr(wsOut.Range("J2").Value2))
        End If
    End If
    If sourceName = MMF_PRINT_SHEET Then
        Err.Raise vbObjectError + 400, , "印刷用シート自身を参照することはできません。"
    End If

    On Error Resume Next
    Set wsSrc = wb.Worksheets(sourceName)
    On Error GoTo Failure
    If wsSrc Is Nothing Then
        Err.Raise vbObjectError + 401, , "対象台帳「" & sourceName & "」が存在しません。"
    End If
    If CStr(wsSrc.Range("A11").Value2) <> "日付" _
       Or CStr(wsSrc.Range("B11").Value2) <> "取引区分" _
       Or CStr(wsSrc.Range("T11").Value2) <> "残高数量" Then
        Err.Raise vbObjectError + 402, , _
            "対象シートの列構成が想定と異なります。MMF取引台帳を指定してください。"
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 判定・計算列(K:Y)は触れず、入力列(A:J)に実データのある行だけ表示する。
    ReDim data(1 To 500, 1 To 15)
    For i = MMF_FIRST_SOURCE_ROW To MMF_LAST_SOURCE_ROW
        If Application.WorksheetFunction.CountA(wsSrc.Range("A" & i & ":J" & i)) > 0 Then
            n = n + 1
            rr = MMF_FIRST_PRINT_ROW + n - 1
            For j = 1 To 15
                data(n, j) = MMF_印刷数式(sourceName, i, rr, j)
            Next j
        End If
    Next i

    If wsOut Is Nothing Then
        Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsOut.Name = MMF_PRINT_SHEET
    End If
    wasProtected = wsOut.ProtectContents
    If wasProtected Then wsOut.Unprotect Password:=MMF_PRINT_PASSWORD
    wsOut.Range("A1:R6").UnMerge
    wsOut.Cells.Clear

    With wsOut
        .Cells.Font.Name = "Yu Gothic"
        .Cells.Font.Size = 9
        .Cells.Font.Color = RGB(30, 49, 65)
        .Cells.VerticalAlignment = xlCenter
        .Cells.HorizontalAlignment = xlRight
        .Cells.ShrinkToFit = True

        widths = Array(9, 4.4, 10.5, 8, 10.7, 10.7, 10, 10, 10, 10.3, 10.7, 10.7, 8.4, 7.4, 15)
        For j = 1 To 15
            .Columns(j).ColumnWidth = widths(j - 1)
        Next j
        .Columns("P").ColumnWidth = 2
        .Columns("Q").ColumnWidth = 18
        .Columns("R").ColumnWidth = 24

        .Range("A1:O1").Merge
        .Range("A1").Value = "外貨建MMF　取得費計算明細　｜　A4縦・1取引1行"
        .Range("A1:O1").Interior.Color = RGB(25, 54, 76)
        .Range("A1:O1").Font.Color = vbWhite
        .Range("A1:O1").Font.Bold = True
        .Range("A1:O1").Font.Size = 15
        .Range("A1").HorizontalAlignment = xlLeft
        .Rows(1).RowHeight = 31

        .Range("A2:F2").Merge
        .Range("G2:J2").Merge
        .Range("K2:O2").Merge
        .Range("A2").Formula = "=""銘柄・口座：""&" & MMF_参照(sourceName, "B", 3)
        .Range("G2").Formula = "=""通貨：""&" & MMF_参照(sourceName, "F", 3)
        .Range("K2").Formula = "=""出力日：""&TEXT(TODAY(),""yyyy/m/d"")"
        .Range("A2:O2").Interior.Color = RGB(221, 235, 243)
        .Range("A2:O2").Font.Color = RGB(25, 54, 76)
        .Range("A2:O2").Font.Bold = True
        .Range("A2:O2").HorizontalAlignment = xlLeft
        .Rows(2).RowHeight = 22

        .Range("A3:F3").Merge
        .Range("G3:J3").Merge
        .Range("K3:O3").Merge
        .Range("A3").Formula = "=""期首数量：""&TEXT(" & MMF_参照(sourceName, "B", 4) & ",""#,##0.########"")"
        .Range("G3").Formula = "=""期首外貨簿価：""&TEXT(" & MMF_参照(sourceName, "D", 4) & ",""#,##0.########"")"
        .Range("K3").Formula = "=""期首円貨簿価：""&TEXT(" & MMF_参照(sourceName, "F", 4) & ",""#,##0"")"
        .Range("A3:O3").Interior.Color = RGB(235, 245, 249)
        .Range("A3:O3").HorizontalAlignment = xlLeft
        .Rows(3).RowHeight = 21

        .Range("A4:O4").Merge
        .Range("A4").Value = "買・累投は買付数量と買付価額、売は売付数量と売付価額を表示。売取得費と売損益は売付時のみ表示。"
        .Range("A5:O5").Merge
        .Range("A5").Value = "残高は移動平均法による取得原価（時価ではありません）。「要確認」の取引は元台帳の入力判定をご確認ください。"
        .Range("A4:O5").HorizontalAlignment = xlLeft
        .Range("A4:O5").Interior.Color = RGB(239, 248, 251)
        .Range("A4:O5").Font.Size = 8.5
        .Rows("4:5").RowHeight = 19

        headers = Array("取引日", "区分", "売買数量", "買付単価", _
                        "売買価額" & vbLf & "外貨", "売買価額" & vbLf & "円貨", _
                        "売取得費" & vbLf & "外貨", "売取得費" & vbLf & "円貨", _
                        "売損益" & vbLf & "円貨", "残数量", "残高" & vbLf & "外貨", _
                        "残高" & vbLf & "円貨", "換算レート", "為替区分", "備考／確認")
        For j = 1 To 15
            .Cells(6, j).Value = headers(j - 1)
        Next j
        With .Range("A6:O6")
            .Interior.Color = RGB(32, 71, 97)
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Size = 8.5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .ShrinkToFit = True
        End With
        .Rows(6).RowHeight = 34

        .Range("Q1:R1").Merge
        .Range("Q1").Value = "印刷対象の設定（印刷対象外）"
        .Range("Q1:R1").Interior.Color = RGB(32, 71, 97)
        .Range("Q1:R1").Font.Color = vbWhite
        .Range("Q1:R1").Font.Bold = True
        .Range("Q2").Value = "対象シート名"
        .Range("R2").Value = sourceName
        .Range("R2").Font.Color = RGB(24, 80, 167)
        .Range("R2").Interior.Color = RGB(255, 242, 199)
        .Range("Q3").Value = "明細件数"
        .Range("R3").Value = n
        .Range("Q4").Value = "変更・追加後"
        .Range("R4").Value = "Alt+F8 → MMF印刷_1段表示へ更新"
        .Range("R4").WrapText = True
        .Rows(4).RowHeight = 29
        .Range("Q2:Q4").Interior.Color = RGB(231, 241, 247)
        .Range("Q2:Q4").HorizontalAlignment = xlLeft
        .Range("R2:R4").HorizontalAlignment = xlLeft

        lastOut = MMF_FIRST_PRINT_ROW + n - 1
        If n > 0 Then
            .Range("A" & MMF_FIRST_PRINT_ROW).Resize(n, 15).Formula = data
            With .Range("A7:O" & lastOut)
                .RowHeight = 20
                .Font.Size = 9
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(217, 228, 235)
                .Borders.Weight = xlHairline
            End With
            With .Range("A7:O" & lastOut).FormatConditions
                .Delete
                Set fmt = .Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0")
                fmt.Interior.Color = RGB(241, 248, 252)
            End With
            .Range("A7:A" & lastOut).NumberFormat = "yy/m/d"
            .Range("C7:C" & lastOut).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("D7:E" & lastOut).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("F7:F" & lastOut).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("G7:G" & lastOut).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("H7:I" & lastOut).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("J7:K" & lastOut).NumberFormat = "#,##0.########;[Red](#,##0.########);-"
            .Range("L7:L" & lastOut).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("M7:M" & lastOut).NumberFormat = "0.####;[Red](0.####);-"
            .Range("B7:B" & lastOut).HorizontalAlignment = xlCenter
            .Range("N7:N" & lastOut).HorizontalAlignment = xlCenter
            .Range("O7:O" & lastOut).HorizontalAlignment = xlLeft
        Else
            lastOut = 6
        End If

        .Activate
        ActiveWindow.DisplayGridlines = False
        .PageSetup.PrintArea = "$A$1:$O$" & lastOut
        .PageSetup.PrintTitleRows = "$1:$6"
        With .PageSetup
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
            .CenterFooter = "外貨建MMF 取得費計算　　&P / &N"
            .RightFooter = "&D"
        End With
        .Range("A7").Select
    End With
    complete = True
    MMF_PrintLastSuccess = True

CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsOut.Range("R2").Locked = False
        wsOut.Protect Password:=MMF_PRINT_PASSWORD, UserInterfaceOnly:=True
    End If
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If complete Then
        MsgBox "A4縦・1取引1行の印刷明細を更新しました。" & vbCrLf & _
               "取引件数：" & CStr(n) & " 件。印刷プレビューで文字の大きさをご確認ください。", vbInformation
    Else
        MsgBox "印刷明細を更新できませんでした。" & vbCrLf & failureMessage, vbExclamation
    End If
    Exit Sub

Failure:
    failureMessage = Err.Description
    Resume CleanExit
End Sub

' 入力変更・対象シート変更を反映してからプレビューを開く。
Public Sub MMF印刷_1段表示プレビュー()
    MMF印刷_1段表示へ更新
    If MMF_PrintLastSuccess Then
        ThisWorkbook.Worksheets(MMF_PRINT_SHEET).PrintPreview
    End If
End Sub

Private Function MMF_参照(ByVal sheetName As String, ByVal col As String, ByVal rowNum As Long) As String
    MMF_参照 = "'" & Replace(sheetName, "'", "''") & "'!" & col & CStr(rowNum)
End Function

Private Function MMF_空欄対応(ByVal expression As String, ByVal printRow As Long) As String
    Dim q As String
    q = Chr$(34)
    MMF_空欄対応 = "=IFERROR(IF($A" & CStr(printRow) & "=" & q & q & "," & q & q & _
                    ",IF(" & expression & "=" & q & q & "," & q & q & "," & expression & "))," & q & q & ")"
End Function

Private Function MMF_印刷数式(ByVal sourceName As String, ByVal srcRow As Long, _
                               ByVal printRow As Long, ByVal colNo As Long) As String
    Dim a As String, b As String, e As String, k As String
    Dim expr As String, q As String
    q = Chr$(34)
    a = MMF_参照(sourceName, "A", srcRow)
    b = MMF_参照(sourceName, "B", srcRow)
    e = MMF_参照(sourceName, "E", srcRow)
    k = MMF_参照(sourceName, "K", srcRow)

    Select Case colNo
        Case 1
            ' A列自体は空欄判定用のため、自己参照を避ける。
            MMF_印刷数式 = "=IFERROR(IF(" & a & "=" & q & q & "," & q & q & "," & a & ")," & q & q & ")"
            Exit Function
        Case 2
            expr = b
        Case 3
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & e & "," & MMF_参照(sourceName, "C", srcRow) & ")"
        Case 4
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & q & q & "," & MMF_参照(sourceName, "D", srcRow) & ")"
        Case 5
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & MMF_参照(sourceName, "O", srcRow) & "," & MMF_参照(sourceName, "M", srcRow) & ")"
        Case 6
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & MMF_参照(sourceName, "P", srcRow) & "," & MMF_参照(sourceName, "N", srcRow) & ")"
        Case 7
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & MMF_参照(sourceName, "Q", srcRow) & "," & q & q & ")"
        Case 8
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & MMF_参照(sourceName, "R", srcRow) & "," & q & q & ")"
        Case 9
            expr = "IF($B" & printRow & "=" & q & "売" & q & "," & MMF_参照(sourceName, "S", srcRow) & "," & q & q & ")"
        Case 10
            expr = MMF_参照(sourceName, "T", srcRow)
        Case 11
            expr = MMF_参照(sourceName, "U", srcRow)
        Case 12
            expr = MMF_参照(sourceName, "V", srcRow)
        Case 13
            expr = MMF_参照(sourceName, "L", srcRow)
        Case 14
            expr = MMF_参照(sourceName, "H", srcRow)
        Case 15
            expr = "IF(" & k & "=" & q & "OK" & q & "," & _
                    MMF_参照(sourceName, "J", srcRow) & "," & _
                    q & "要確認：" & q & "&" & k & "&" & q & " " & q & _
                    "&" & MMF_参照(sourceName, "J", srcRow) & ")"
        Case Else
            Err.Raise vbObjectError + 403, , "未定義の印刷列：" & CStr(colNo)
    End Select
    MMF_印刷数式 = MMF_空欄対応(expr, printRow)
End Function
