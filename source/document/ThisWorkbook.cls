VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub DownloadMurcFX_MultiYear_Final()

    Dim wsActive As Worksheet   ' 元シート保持
    Set wsActive = ActiveSheet
    
    Dim yearInput As String
    Dim years() As String
    Dim y As Variant
    
    Dim url As String
    Dim ws As Worksheet
    Dim wsM As Worksheet
    Dim tmpWB As Workbook
    Dim tmpFile As String
    
    Dim pasteRow As Long
    Dim lastRowTmp As Long
    Dim lastColTmp As Long
    Dim lastCol As Long
    Dim lastRow As Long
    
    Dim i As Long, j As Long
    Dim startRow As Long
    Dim isFirst As Boolean
    
    Dim yy As Long
    Dim targetY As Long
    
    ' ===== 高速化 =====
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ===== シート設定 =====
    On Error Resume Next
    Set ws = Worksheets("為替データ")
    Set wsM = Worksheets("マスタ")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "「為替データ」シートが存在しません。", vbCritical
        GoTo ExitHandler
    End If
    
    If wsM Is Nothing Then
        MsgBox "「マスタ」シートが存在しません。", vbCritical
        GoTo ExitHandler
    End If
    
    ' ===== 年入力 =====
    yearInput = InputBox("取得する西暦（例：2025 または 2024,2025）")
    
    If Trim(yearInput) = "" Then
        MsgBox "年が入力されていません。", vbExclamation
        GoTo ExitHandler
    End If
    
    years = Split(yearInput, ",")
    
    ' ===== 初期化 =====
    ws.Cells.Clear
    wsM.Range("B2:B1000").ClearContents
    
    pasteRow = 3
    isFirst = True
    
    ' ===== 年ループ =====
    For Each y In years
    
        y = Trim(y)
        If y = "" Then GoTo NextYear
        
        targetY = CLng(y)
        
        ' ===== 前年＋当年 =====
        For yy = targetY - 1 To targetY
        
            url = "https://www.murc-kawasesouba.jp/fx/xls/murc_" & yy & ".xls"
            tmpFile = Environ("TEMP") & "\murc_" & yy & ".xls"
            
            On Error Resume Next
            Kill tmpFile
            On Error GoTo 0
            
            ' ===== ダウンロード =====
            With CreateObject("MSXML2.ServerXMLHTTP.6.0")
                .Open "GET", url, False
                .send
                
                If .Status <> 200 Then
                    MsgBox yy & "年の取得失敗", vbExclamation
                    GoTo NextYY
                End If
                
                Dim stream As Object
                Set stream = CreateObject("ADODB.Stream")
                stream.Type = 1
                stream.Open
                stream.Write .responseBody
                stream.SaveToFile tmpFile, 2
                stream.Close
            End With
            
            ' ===== 開く =====
            Set tmpWB = Workbooks.Open(tmpFile)
            
            With tmpWB.Sheets(1)
            
                lastRowTmp = .Cells(.Rows.Count, "A").End(xlUp).Row
                lastColTmp = .Cells(2, .Columns.Count).End(xlToLeft).Column
                
                ' ===== 開始行判定 =====
                startRow = 0
                For i = 1 To 10
                    If IsDate(.Cells(i, 1).Value) Then
                        startRow = i
                        Exit For
                    End If
                Next i
                
                If startRow = 0 Then
                    tmpWB.Close False
                    GoTo NextYY
                End If
                
                ' ===== コピー =====
                If isFirst Then
                    .Range(.Cells(1, 1), .Cells(lastRowTmp, lastColTmp)).Copy ws.Cells(pasteRow, 1)
                    isFirst = False
                Else
                    .Range(.Cells(startRow, 1), .Cells(lastRowTmp, lastColTmp)).Copy ws.Cells(pasteRow, 1)
                End If
                
                pasteRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
                
            End With
            
            tmpWB.Close False

NextYY:
        Next yy
        
NextYear:
    Next y
    
    ' ===== 仕上げ =====
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    ws.Range("A4:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).NumberFormat = "yyyy/m/d"
    
    ws.Columns("A").ColumnWidth = 12
    For j = 3 To lastCol
        ws.Columns(j).ColumnWidth = 6.2
    Next j
    
    ' ===== C列空白削除 =====
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For i = lastRow To 4 Step -1
        If Trim(ws.Cells(i, "C").Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' ===== 通貨マスタ =====
    Dim masterRow As Long
    masterRow = 2
    
    For j = 3 To lastCol
        If ws.Cells(3, j).Value <> "" Then
            wsM.Cells(masterRow, 2).Value = ws.Cells(3, j).Value
            masterRow = masterRow + 1
        End If
    Next j
    
    wsM.Range("B2:B" & masterRow - 1).RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' ===== 表示調整 =====
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True

ExitHandler:

    ' ===== 元シートへ戻る =====
    wsActive.Activate
    
    ' ===== 再計算 =====
    Application.CalculateFull
    
    ' ===== 設定戻し =====
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "前年＋当年の取得完了", vbInformation

End Sub

Private Sub Workbook_Open()

    '■■ キー割当
    Application.OnKey "{F7}", "F7_行削除"
    Application.OnKey "{F8}", "F8_行追加"
    
    '■■ スペースキー
    Call SetSpaceKey

    '■■ MMF取引台帳の取引区分を自由入力に設定（元の計算式は変更しない）
    Call MMF取引区分_起動時設定
    ' 入力内容ではなく数量に基づく取得費判定へ切替
    Call MMF数量判定_起動時設定

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    '■■ キー解除
    Application.OnKey "{F7}"
    Application.OnKey "{F8}"
    
    '■■ スペースキー解除
    Call ResetSpaceKey

End Sub


' A4印刷専用シートだけ印刷直前に再集計する。
' 通常の出納帳・台帳や別ブックの印刷には介入しない。
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    On Error GoTo PrintFailed

    If Not (ActiveWorkbook Is Me) Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Name <> "A4縦印刷" Then Exit Sub

    If Not MMF印刷_印刷直前更新() Then Cancel = True
    Exit Sub

PrintFailed:
    Cancel = True
    MsgBox "A4縦印刷シートの更新中に問題が発生したため、印刷を中止しました。" & _
           vbCrLf & Err.Description, vbExclamation
End Sub
