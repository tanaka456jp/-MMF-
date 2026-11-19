Attribute VB_Name = "Module1"
Option Explicit
Public Const PWD As String = "1234"
Public Const START_ROW As Long = 10
Public Const END_ROW As Long = 1000

Function FX取得(日付 As Date, 為替タイプ As String, 通貨 As String) As Variant

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, colBase As Long, col As Long
    Dim sumVal As Double, cnt As Long
    
    Set ws = ThisWorkbook.Worksheets("為替データ")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    '--- 通貨列取得
    colBase = 0
    For i = 3 To lastCol
        If ws.Cells(3, i).Value = 通貨 Then
            colBase = i
            Exit For
        End If
    Next i
    
    If colBase = 0 Then
        FX取得 = CVErr(xlErrNA)
        Exit Function
    End If
    
    '--- 為替タイプで列調整
    col = colBase
    If InStr(為替タイプ, "TTB") > 0 Then col = colBase + 1
    If InStr(為替タイプ, "TTM") > 0 Then col = colBase + 2
    
    '========================
    ' 単日取得
    '========================
    If InStr(為替タイプ, "月") = 0 And InStr(為替タイプ, "年") = 0 Then
        
        For i = lastRow To 4 Step -1
            If IsDate(ws.Cells(i, 1).Value) Then
                If ws.Cells(i, 1).Value <= 日付 Then
                    FX取得 = ws.Cells(i, col).Value
                    Exit Function
                End If
            End If
        Next i
        
    End If
    
    '========================
    ' 月平均
    '========================
    If InStr(為替タイプ, "月") > 0 Then
        
        For i = 4 To lastRow
            If IsDate(ws.Cells(i, 1).Value) Then
                If Year(ws.Cells(i, 1).Value) = Year(日付) _
                And Month(ws.Cells(i, 1).Value) = Month(日付) Then
                    
                    If IsNumeric(ws.Cells(i, col).Value) Then
                        sumVal = sumVal + ws.Cells(i, col).Value
                        cnt = cnt + 1
                    End If
                    
                End If
            End If
        Next i
        
        If cnt > 0 Then
            FX取得 = sumVal / cnt
        Else
            FX取得 = ""
        End If
        
        Exit Function
    End If
    
    '========================
    ' 年平均
    '========================
    If InStr(為替タイプ, "年") > 0 Then
        
        For i = 4 To lastRow
            If IsDate(ws.Cells(i, 1).Value) Then
                If Year(ws.Cells(i, 1).Value) = Year(日付) Then
                    
                    If IsNumeric(ws.Cells(i, col).Value) Then
                        sumVal = sumVal + ws.Cells(i, col).Value
                        cnt = cnt + 1
                    End If
                    
                End If
            End If
        Next i
        
        If cnt > 0 Then
            FX取得 = sumVal / cnt
        Else
            FX取得 = ""
        End If
        
        Exit Function
    End If

End Function
Sub 現金出納帳_行追加_高速()

    Dim ws As Worksheet
    Dim addVal As Variant
    Dim addCount As Long
    Dim lastRow As Long
    Dim newRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    '入力
    addVal = Application.InputBox("追加する行数", Type:=1)
    If addVal = False Or addVal <= 0 Then Exit Sub
    addCount = CLng(addVal)
    
    '高速化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '保護解除（最初だけ）
    ws.Unprotect "1234"
    
    For i = 1 To addCount
        
        lastRow = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        newRow = lastRow + 1
        
        '行追加
        ws.Rows(newRow).Insert
        
        '数式・書式コピー（確実）
        ws.Range("A" & lastRow & ":L" & lastRow).Copy
        ws.Range("A" & newRow).PasteSpecial xlPasteAll
        
        Application.CutCopyMode = False
        
        '入力列クリア
        ws.Range("A" & newRow).ClearContents
        ws.Range("B" & newRow).ClearContents
        ws.Range("D" & newRow).ClearContents
        ws.Range("I" & newRow & ":L" & newRow).ClearContents
        
    Next i
    
    '共通処理
    Call 印刷設定_外枠(ws)
    
    '再保護（最後だけ）
    ws.Protect "1234", UserInterfaceOnly:=True, _
        AllowInsertingRows:=True, AllowDeletingRows:=True
    
    '戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
Sub 現金出納帳_行削除_高速()

    Dim ws As Worksheet
    Dim delVal As Variant
    Dim delCount As Long
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    '入力
    delVal = Application.InputBox("削除する行数", Type:=1)
    If delVal = False Or delVal <= 0 Then Exit Sub
    delCount = CLng(delVal)
    
    '高速化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '保護解除
    ws.Unprotect "1234"
    
    For i = 1 To delCount
        
        lastRow = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        
        If lastRow <= 10 Then Exit For
        
        ws.Rows(lastRow).Delete
        
    Next i
    
    '共通処理
    Call 印刷設定_外枠(ws)
    
    '再保護
    ws.Protect "1234", UserInterfaceOnly:=True, _
        AllowInsertingRows:=True, AllowDeletingRows:=True
    
    '戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
Sub 印刷設定_外枠(ws As Worksheet)

    Dim lastRow As Long
    Dim rng As Range
    Dim breakRow As Long
    
    lastRow = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    If lastRow < 10 Then Exit Sub
    
    Set rng = ws.Range("A10:L" & lastRow)
    
    '========================
    ' ■ 罫線リセット
    '========================
    rng.Borders.LineStyle = xlNone
    
    '========================
    ' ■ 内枠（極細）
    '========================
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    '========================
    ' ■ 外枠（細線）
    '========================
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    '========================
    ' ■ 印刷範囲
    '========================
    ws.PageSetup.PrintArea = "$A$1:$L$" & lastRow
    
    '========================
    ' ■ 印刷設定
    '========================
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    '========================
    ' ■ 改ページ（66行単位）
    '========================
    ws.ResetAllPageBreaks
    
    breakRow = 10 + 66
    
    Do While breakRow <= lastRow
        ws.HPageBreaks.Add Before:=ws.Rows(breakRow)
        breakRow = breakRow + 66
    Loop

End Sub
Sub チェックボックス連動_H4K4()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Const PW As String = "1234" 'パスワード必要なら設定

    Application.ScreenUpdating = False

    '① 保護解除
    ws.Unprotect Password:="1234"

    With ws.Range("H4:K4")

        If ws.Range("A2").Value = True Then
            '? ON → 黒文字
            .Font.Color = RGB(0, 0, 0)
        Else
            '? OFF → 背景と同色
            .Font.Color = .Interior.Color
        End If

    End With

    '② 再保護
    ws.Protect Password:="1234", UserInterfaceOnly:=True

    Application.ScreenUpdating = True

End Sub
Sub 入力内容削除_確認付き()

    Dim ret As VbMsgBoxResult
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim isProtected As Boolean

    Set ws = ActiveSheet

    ret = MsgBox("入力内容を全部削除します。よろしいですか？", _
                 vbYesNo + vbQuestion, "確認")

    If ret = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    With ws

        isProtected = .ProtectContents

        If isProtected Then
            .Unprotect Password:="1234"
        End If

        '▼入力値クリア
        .Range("B4:E4," & _
               "A10:A1048576,B10:B1048576,D10:D1048576," & _
               "I10:I1048576,J10:J1048576,K10:K1048576,L10:L1048576").ClearContents

        '■■ここ追加（コメント全削除）
        Dim cmt As Object
        For Each cmt In .CommentsThreaded
            cmt.Delete
        Next cmt

        Dim cmtOld As Comment
        For Each cmtOld In .Comments
            cmtOld.Delete
        Next cmtOld

        '▼最終行取得
        lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        If lastRow < 10 Then lastRow = 10

        '① H4:K4 の数式復元
        .Range("H4").Formula = "=LOOKUP(2,1/(F:F<>""""),F:F)"
        .Range("I4").Formula = "=IFERROR(ROUND(FX取得(B6,I5,$D$4),2),"""")"
        .Range("J4").Formula = "=K4-LOOKUP(2,1/(G:G<>""""),G:G)"
        .Range("K4").Formula = "=ROUND(H4*I4,0)"

        '② C列
        .Range("C10").Formula = "=IFERROR(出納帳!$B10*出納帳!$H10,"""")"
        .Range("C10:C" & lastRow).FillDown

        '③ E～H列
        .Range("E10").Formula = "=IFERROR(出納帳!$D10*出納帳!$H10,"""")"
        .Range("F10").Formula = "=$B$4+SUM($B$10:B10)-SUM($D$10:D10)"
        .Range("G10").Formula = "=$C$4+SUM($C$10:C10)-SUM($E$10:E10)"
        .Range("H10").Formula = "=IFERROR(ROUND(FX取得(A10,I10,$D$4),2),"""")"

        .Range("E10:H" & lastRow).FillDown

        If isProtected Then
            .Protect Password:="1234", UserInterfaceOnly:=True
        End If

    End With

    Application.ScreenUpdating = True

    MsgBox "入力内容を削除し、数式を復元しました", vbInformation

End Sub

Sub CheckBox_Control()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cb As CheckBox
    Set cb = ws.CheckBoxes(Application.Caller)

    Dim isChecked As Boolean
    isChecked = (cb.Value = 1)

    Application.ScreenUpdating = False

    ws.Unprotect Password:="1234"

    '■■状態保持
    ws.Range("AA1").Value = isChecked
    ws.Range("AA1").EntireColumn.Hidden = True

    '■■対象範囲
    Dim rng As Range
    Set rng = Union(ws.Range("C10:C1000"), _
                    ws.Range("E10:E1000"), _
                    ws.Range("H10:H1000"), _
                    ws.Range("H4,I4,J4,K4"))

    If isChecked Then
        '■■対象だけ編集可
        rng.Locked = False

        ws.Shapes("txtOverwrite").Visible = msoTrue

        Call RestoreItalic

    Else
        '■■対象だけ編集不可
        rng.Locked = True

        ws.Shapes("txtOverwrite").Visible = msoFalse

        rng.Font.Italic = False
    End If

    ws.Protect Password:="1234", _
        UserInterfaceOnly:=True, _
        AllowDeletingRows:=isChecked, _
        AllowInsertingRows:=isChecked

    Application.ScreenUpdating = True

End Sub


'========================================
' 斜体復元（コメント基準）
'========================================
Public Sub RestoreItalic()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng As Range
    Set rng = Union(ws.Range("C10:C1000"), _
                    ws.Range("E10:E1000"), _
                    ws.Range("H10:H1000"), _
                    ws.Range("H4,I4,J4,K4"))

    '■■全解除
    rng.Font.Italic = False

    If ws.Range("AA1").Value <> True Then Exit Sub

    Dim c As Range

    For Each c In rng
        If Not c.CommentThreaded Is Nothing Then
            If Not c.HasFormula Then
                c.Font.Italic = True
            End If
        End If
    Next c

End Sub



Sub 印刷プレビュー_斜体制御()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng As Range
    Set rng = Union(ws.Range("C10:C1000"), _
                    ws.Range("E10:E1000"), _
                    ws.Range("H10:H1000"), _
                    ws.Range("H4:I4:J4:K4"))

    Dim c As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Application.ScreenUpdating = False

    '■■状態記録（安全化）
    For Each c In rng.Cells
        On Error Resume Next
        If c.Font.Italic = True Then
            dict(c.Address) = True
        End If
        On Error GoTo 0
    Next c

    '■■解除
    rng.Font.Italic = False

    Application.ScreenUpdating = True

    '■■プレビュー
    ws.PrintPreview

    Application.ScreenUpdating = False

    '■■復元
    For Each c In rng.Cells
        If dict.exists(c.Address) Then
            c.Font.Italic = True
        Else
            c.Font.Italic = False
        End If
    Next c

SafeExit:
    Application.ScreenUpdating = True

End Sub
Sub F7_行削除()

    Dim ws As Worksheet
    Dim dict As Object
    Dim a As Range, r As Range
    Dim arr() As Long
    Dim key As Variant
    Dim i As Long, j As Long, temp As Long

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    If TypeName(Selection) <> "Range" Then Exit Sub

    '■■行取得（範囲を直書き：10～1000）
    For Each a In Selection.Areas
        For Each r In a.Rows

            If r.Row >= 10 And r.Row <= 1000 Then
                If Not dict.exists(r.Row) Then
                    dict.Add r.Row, r.Row
                End If
            End If

        Next r
    Next a

    If dict.Count = 0 Then
        MsgBox "削除対象がありません（10～1000行のみ対象）", vbExclamation
        Exit Sub
    End If

    If MsgBox("選択した行を削除しますか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ws.Unprotect Password:="1234"

    ReDim arr(0 To dict.Count - 1)

    i = 0
    For Each key In dict.Keys
        arr(i) = key
        i = i + 1
    Next key

    '降順ソート
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    '削除
    For i = LBound(arr) To UBound(arr)
        ws.Rows(arr(i)).Delete
    Next i

    ws.Protect Password:="1234", _
        UserInterfaceOnly:=True, _
        AllowDeletingRows:=True, _
        AllowInsertingRows:=True

    ws.EnableSelection = xlNoRestrictions

    Application.ScreenUpdating = True

End Sub
Sub F8_行追加()

    Dim ws As Worksheet
    Dim dict As Object
    Dim a As Range, r As Range
    Dim arr() As Long
    Dim key As Variant
    Dim i As Long, j As Long, temp As Long
    Dim baseRow As Long

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    If TypeName(Selection) <> "Range" Then Exit Sub

    '■■行取得（10～1000）
    For Each a In Selection.Areas
        For Each r In a.Rows

            If r.Row >= 10 And r.Row <= 1000 Then
                If Not dict.exists(r.Row) Then
                    dict.Add r.Row, r.Row
                End If
            End If

        Next r
    Next a

    If dict.Count = 0 Then
        MsgBox "追加対象がありません（10～1000行のみ対象）", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ws.Unprotect Password:="1234"

    ReDim arr(0 To dict.Count - 1)

    i = 0
    For Each key In dict.Keys
        arr(i) = key
        i = i + 1
    Next key

    '降順ソート
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    '追加＋コピー
    For i = LBound(arr) To UBound(arr)

        baseRow = arr(i)

        ws.Rows(baseRow + 1).Insert

        ws.Rows(baseRow).Copy
        ws.Rows(baseRow + 1).PasteSpecial xlPasteFormulas
        ws.Rows(baseRow).Copy Destination:=ws.Rows(baseRow + 1)

    Next i

    Application.CutCopyMode = False

    ws.Protect Password:="1234", _
        UserInterfaceOnly:=True, _
        AllowDeletingRows:=True, _
        AllowInsertingRows:=True

    ws.EnableSelection = xlNoRestrictions

    Application.ScreenUpdating = True

End Sub


'========================================
' スペースキー割当
'========================================
Public Sub SetSpaceKey()
    Application.OnKey " ", "ShowKamokuUI"
End Sub

'========================================
' 解除
'========================================
Public Sub ResetSpaceKey()
    Application.OnKey " "
End Sub

'========================================
' UI起動判定
'========================================
Public Sub ShowKamokuUI()

    '■■ 対象シート限定（必要に応じて変更）
    If ActiveSheet.Name <> "出納帳" Then Exit Sub

    '■■ J10以降のみ
    If Intersect(Selection, ActiveSheet.Range("J10:J" & Rows.Count)) Is Nothing Then Exit Sub

    '■■ 複数セル防止
    If Selection.Cells.Count > 1 Then Exit Sub

    '■■ UI起動
    frmKamoku.Show

End Sub
