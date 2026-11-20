Attribute VB_Name = "ModuleMMFCalc"
Option Explicit

' 外貨建MMFの取引区分（B列）は表示・備考用途のみ。
' 買付は C列>0、売付は E列>0 で判定する。
' この移行では元の計算・丸め・為替換算ロジックは維持し、
' K列の入力判定と、M:X列の売買判定だけを変更する。
Private Const LEDGER_SHEET As String = "MMF取引台帳"
Private Const FIRST_ROW As Long = 12
Private Const MIN_LAST_ROW As Long = 511
Private Const MAX_ROWS As Long = 5000
Private mApplying As Boolean

' VBAUpdater反映後にブックを閉じずに適用したい場合に使用する。
Public Sub MMF数量判定_適用()
    If MMF数量判定_変更本体() Then
        MsgBox "取引区分と計算区分を分離しました。" & vbCrLf & _
               "買付数量(C列)>0は買付、売付数量(E列)>0は売付として計算します。" & vbCrLf & _
               "取引区分(B列)の文字は計算に影響しません。" & vbCrLf & _
               "両方の数量に正数が入力された場合は入力エラーとして扱います。", _
               vbInformation, "外貨建MMF"
    End If
End Sub

' ThisWorkbook.Workbook_Openから呼び出す。更新済みなら再変更しない。
Public Sub MMF数量判定_起動時設定()
    Call MMF数量判定_変更本体
End Sub

Private Function MMF数量判定_変更本体() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, inputLast As Long
    Dim r As Long, p As Long, c As Long
    Dim oldK As Variant, newK As Variant
    Dim oldMX As Variant, newMX As Variant
    Dim f As String, oldSale As String
    Dim oldType As String, newType As String
    Dim oldSellCondition As String, newSellCondition As String
    Dim oldBuyCondition As String, newBuyCondition As String
    Dim changed As Boolean, writeStarted As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean
    Dim errorText As String
    Dim quote As String

    If mApplying Then Exit Function
    mApplying = True
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(LEDGER_SHEET)
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 601, , "元台帳が保護されています。保護解除後に実行してください。"
    End If
    If CStr(ws.Range("B11").Value2) <> "取引区分" _
       Or CStr(ws.Range("C11").Value2) <> "買付数量" _
       Or CStr(ws.Range("E11").Value2) <> "売付数量" _
       Or CStr(ws.Range("K11").Value2) <> "入力判定" Then
        Err.Raise vbObjectError + 602, , "MMF取引台帳の列構成が変更されています。更新を中止しました。"
    End If

    lastRow = Application.Max(MIN_LAST_ROW, _
               ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
    inputLast = Application.Max(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
                ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
                ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    If inputLast > lastRow Then
        Err.Raise vbObjectError + 603, , _
                  "入力済みの行に計算式が不足しています。計算式を複写してから再実行してください。"
    End If
    If lastRow - FIRST_ROW + 1 > MAX_ROWS Then
        Err.Raise vbObjectError + 604, , "対象行が上限を超えています。"
    End If

    quote = Chr$(34)
    oldK = ws.Range("K" & FIRST_ROW & ":K" & lastRow).Formula
    oldMX = ws.Range("M" & FIRST_ROW & ":X" & lastRow).Formula
    newK = oldK
    newMX = oldMX

    ' すべての数式を先に検証・変換してから、一括書込みする。
    ' 不明な旧式数式は推測して書き換えず停止する。
    For r = FIRST_ROW To lastRow
        p = r - FIRST_ROW + 1
        oldSale = "B" & r & "=" & quote & "売" & quote
        oldType = "IF(AND(B" & r & "<>" & quote & "売" & quote & _
                  ",B" & r & "<>" & quote & "買" & quote & _
                  ",B" & r & "<>" & quote & "累投" & quote & _
                  ")," & quote & "区分を確認" & quote & ","
        newType = "IF(AND(C" & r & ">0,E" & r & ">0)," & _
                  quote & "売買数量が重複" & quote & ","
        oldSellCondition = "OR(E" & r & "<=0,C" & r & "<>" & quote & quote & _
                           ",D" & r & "<>" & quote & quote & _
                           ",F" & r & "=" & quote & quote & ")"
        newSellCondition = "OR(E" & r & "<=0,C" & r & ">0," & _
                           "D" & r & "<>" & quote & quote & _
                           ",F" & r & "=" & quote & quote & ")"
        oldBuyCondition = "OR(C" & r & "<=0,E" & r & "<>" & quote & quote & _
                          ",AND(F" & r & "=" & quote & quote & ",D" & r & "<=0))"
        newBuyCondition = "OR(C" & r & "<=0,E" & r & ">0," & _
                          "AND(F" & r & "=" & quote & quote & ",D" & r & "<=0))"

        f = CStr(oldK(p, 1))
        If Left$(f, 1) <> "=" Then
            Err.Raise vbObjectError + 605, , "K" & r & "に入力判定の計算式がありません。"
        End If

        If InStr(1, f, oldType, vbBinaryCompare) > 0 Then
            If InStr(1, f, oldSale, vbBinaryCompare) = 0 _
               Or InStr(1, f, oldSellCondition, vbBinaryCompare) = 0 _
               Or InStr(1, f, oldBuyCondition, vbBinaryCompare) = 0 Then
                Err.Raise vbObjectError + 606, , "K" & r & "の旧計算式が想定外です。"
            End If
            f = Replace(f, oldType, newType, , , vbBinaryCompare)
            f = Replace(f, oldSellCondition, newSellCondition, , , vbBinaryCompare)
            f = Replace(f, oldBuyCondition, newBuyCondition, , , vbBinaryCompare)
            f = Replace(f, oldSale, "E" & r & ">0", , , vbBinaryCompare)
            newK(p, 1) = f
            changed = True
        ElseIf InStr(1, f, newType, vbBinaryCompare) = 0 _
               Or InStr(1, f, oldSale, vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 607, , _
                      "K" & r & "の計算式を安全に判定できません。変更を中止しました。"
        End If

        For c = 1 To 12
            f = CStr(oldMX(p, c))
            If Left$(f, 1) <> "=" Then
                Err.Raise vbObjectError + 608, , _
                          ws.Cells(r, c + 12).Address(False, False) & _
                          "に計算式がありません。"
            End If
            If InStr(1, f, oldSale, vbBinaryCompare) > 0 Then
                newMX(p, c) = Replace(f, oldSale, "E" & r & ">0", , , vbBinaryCompare)
                changed = True
            End If
        Next c
    Next r

    If changed Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        writeStarted = True
        ws.Range("K" & FIRST_ROW & ":K" & lastRow).Formula = newK
        ws.Range("M" & FIRST_ROW & ":X" & lastRow).Formula = newMX
        ws.Calculate
    End If

    MMF数量判定_変更本体 = True

CleanUp:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    mApplying = False
    On Error GoTo 0
    Exit Function

Failed:
    errorText = Err.Description
    If writeStarted Then
        On Error Resume Next
        Application.EnableEvents = False
        ws.Range("K" & FIRST_ROW & ":K" & lastRow).Formula = oldK
        ws.Range("M" & FIRST_ROW & ":X" & lastRow).Formula = oldMX
        On Error GoTo 0
    End If
    MsgBox "数量判定への移行を中止しました。変更内容を確認してください。" & vbCrLf & _
           errorText, vbExclamation, "外貨建MMF"
    Resume CleanUp
End Function
