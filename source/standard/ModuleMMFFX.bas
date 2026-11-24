Attribute VB_Name = "ModuleMMFFX"
Option Explicit

' MMF専用の為替参照。元の出納帳で使用する FX取得 は変更しない。
' 米ドル(USD) / 米ドル（USD）のような括弧表記差を通貨コードで吸収する。
' 同日の重複は同一レートのみ許容し、異なるレートならエラー扱い。
Private Const MMF_LEDGER As String = "MMF取引台帳"
Private Const MMF_FX_DATA As String = "為替データ"
Private Const FIRST_ROW As Long = 12
Private Const BASE_LAST_ROW As Long = 511
Private mFXInstalling As Boolean

Public Function MMF為替_取得(ByVal tradeDate As Variant, ByVal rateType As Variant, _
                              ByVal currencyLabel As Variant) As Variant
    Dim fx As Worksheet
    Dim targetCode As String, headCode As String
    Dim sourceCol As Long, lastCol As Long, lastRow As Long
    Dim col As Long, rowNo As Long, rateOffset As Long
    Dim targetDay As Long, sourceDay As Long
    Dim cellValue As Variant, rawRate As Variant
    Dim found As Boolean, foundRate As Double

    On Error GoTo InvalidInput

    If IsError(tradeDate) Or IsError(rateType) Or IsError(currencyLabel) Then GoTo InvalidInput
    If Not IsDate(tradeDate) Then GoTo InvalidInput
    targetDay = CLng(Int(CDbl(CDate(tradeDate))))
    targetCode = MMF為替_通貨コード(CStr(currencyLabel))
    If targetCode = "" Then GoTo MissingValue

    Select Case UCase$(Trim$(CStr(rateType)))
        Case "TTS": rateOffset = 0
        Case "TTB": rateOffset = 1
        Case "TTM": rateOffset = 2
        Case Else: GoTo MissingValue
    End Select

    Set fx = ThisWorkbook.Worksheets(MMF_FX_DATA)
    lastCol = fx.Cells(3, fx.Columns.Count).End(xlToLeft).Column
    For col = 3 To lastCol Step 3
        headCode = MMF為替_通貨コード(CStr(fx.Cells(3, col).Value2))
        If headCode = targetCode Then
            sourceCol = col + rateOffset
            Exit For
        End If
    Next col
    If sourceCol = 0 Then GoTo MissingValue

    lastRow = fx.Cells(fx.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then GoTo MissingValue
    ' Excelとの往復アクセスを避け、通貨列と日付列を一括読み込みする。
    ' 大量の明細行での数式再計算を高速化する。
    Dim datesData As Variant, ratesData As Variant
    datesData = fx.Range("A5:A" & lastRow).Value2
    ratesData = fx.Range(fx.Cells(5, sourceCol), fx.Cells(lastRow, sourceCol)).Value2

    For rowNo = 1 To UBound(datesData, 1)
        cellValue = datesData(rowNo, 1)
        If Not IsError(cellValue) Then
            If IsDate(cellValue) Then
                sourceDay = CLng(Int(CDbl(CDate(cellValue))))
                If sourceDay = targetDay Then
                    rawRate = ratesData(rowNo, 1)
                    If IsError(rawRate) Then GoTo InvalidInput
                    If Not IsNumeric(rawRate) Then GoTo InvalidInput
                    If CDbl(rawRate) <= 0 Then GoTo InvalidInput
                    If found Then
                        ' 複数回の取込で同じ日付・同じ値の場合は正しいものと認める。
                        ' 同日で異なる為替レートがあれば自動選択しない。
                        If Abs(CDbl(rawRate) - foundRate) > 0.00000001 Then GoTo AmbiguousRate
                    Else
                        found = True
                        foundRate = CDbl(rawRate)
                    End If
                End If
            End If
        End If
    Next rowNo

    If found Then
        MMF為替_取得 = foundRate
        Exit Function
    End If

MissingValue:
    MMF為替_取得 = CVErr(xlErrNA)
    Exit Function

AmbiguousRate:
    MMF為替_取得 = CVErr(xlErrNum)
    Exit Function

InvalidInput:
    MMF為替_取得 = CVErr(xlErrValue)
End Function

Private Function MMF為替_通貨コード(ByVal rawLabel As String) As String
    Dim textLabel As String
    Dim beginAt As Long, endAt As Long
    Dim code As String

    textLabel = UCase$(Trim$(rawLabel))
    textLabel = Replace(textLabel, "（", "(")
    textLabel = Replace(textLabel, "）", ")")
    beginAt = InStr(1, textLabel, "(", vbBinaryCompare)

    If beginAt > 0 Then
        endAt = InStr(beginAt + 1, textLabel, ")", vbBinaryCompare)
        If endAt > beginAt Then
            code = Trim$(Mid$(textLabel, beginAt + 1, endAt - beginAt - 1))
        End If
    Else
        code = textLabel
    End If

    If Len(code) <> 3 Then Exit Function
    If Not code Like "[A-Z][A-Z][A-Z]" Then Exit Function
    MMF為替_通貨コード = code
End Function

' 更新プログラムからの起動と Alt+F8 からの手動適用に共通。
Public Sub MMF為替_参照式更新()
    Call MMF為替_参照式更新本体
End Sub

Public Sub MMF為替_起動時設定()
    Call MMF為替_参照式更新本体
End Sub

Private Function MMF為替_参照式更新本体() As Boolean
    Dim ledger As Worksheet
    Dim lastRow As Long, dataRow As Long
    Dim rowNo As Long, pos As Long
    Dim origL As Variant, origY As Variant
    Dim destL() As Variant, destY() As Variant
    Dim formulaY As String, formulaL As String
    Dim q As String, changed As Boolean, writeStarted As Boolean
    Dim wasEvents As Boolean, wasScreen As Boolean
    Dim errorMessage As String

    If mFXInstalling Then Exit Function
    mFXInstalling = True
    wasEvents = Application.EnableEvents
    wasScreen = Application.ScreenUpdating
    On Error GoTo Failure

    Set ledger = ThisWorkbook.Worksheets(MMF_LEDGER)
    If ledger.ProtectContents Then
        Err.Raise vbObjectError + 701, , "MMF取引台帳が保護されています。保護解除後に再実行してください。"
    End If
    If CStr(ledger.Range("L11").Value2) <> "換算レート" Or _
       CStr(ledger.Range("Y11").Value2) <> "" Then
        ' Y列は既存の自動参照補助列。変更されている場合は推測で編集しない。
        Err.Raise vbObjectError + 702, , "換算レートの列構成が想定と異なるため更新を中止しました。"
    End If

    lastRow = Application.Max(BASE_LAST_ROW, ledger.Cells(ledger.Rows.Count, "K").End(xlUp).Row)
    dataRow = Application.Max(ledger.Cells(ledger.Rows.Count, "A").End(xlUp).Row, _
                              ledger.Cells(ledger.Rows.Count, "J").End(xlUp).Row)
    If dataRow > lastRow Then
        Err.Raise vbObjectError + 703, , "計算式のない入力行があります。元台帳の計算式を確認してください。"
    End If
    If lastRow - FIRST_ROW + 1 > 5000 Then
        Err.Raise vbObjectError + 704, , "対象の計算行数が上限を超えました。"
    End If

    origL = ledger.Range("L" & FIRST_ROW & ":L" & lastRow).Formula
    origY = ledger.Range("Y" & FIRST_ROW & ":Y" & lastRow).Formula
    ReDim destL(1 To lastRow - FIRST_ROW + 1, 1 To 1)
    ReDim destY(1 To lastRow - FIRST_ROW + 1, 1 To 1)

    q = Chr$(34)
    For rowNo = FIRST_ROW To lastRow
        pos = rowNo - FIRST_ROW + 1
        If Left$(CStr(origL(pos, 1)), 1) <> "=" Or _
           Left$(CStr(origY(pos, 1)), 1) <> "=" Then
            Err.Raise vbObjectError + 705, , "L" & rowNo & "またはY" & rowNo & "に計算式がありません。"
        End If

        If InStr(1, CStr(origY(pos, 1)), "MMF為替_取得(", vbTextCompare) = 0 Then
            If InStr(1, CStr(origY(pos, 1)), "COUNTIF(為替データ!", vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 706, , "Y" & rowNo & "に未対応の計算式があるため更新を中止しました。"
            End If
        End If

        formulaY = "=IF(COUNTA(A" & rowNo & ":J" & rowNo & ")=0," & q & q & _
                   ",IF(OR(A" & rowNo & "=" & q & q & ",H" & rowNo & "=" & q & q & _
                   ",$F$3=" & q & q & ")," & q & q & _
                   ",IFERROR(MMF為替_取得(A" & rowNo & ",H" & rowNo & ",$F$3)," & q & q & ")))"
        formulaL = "=IF(COUNTA(A" & rowNo & ":J" & rowNo & ")=0," & q & q & _
                   ",IF(I" & rowNo & "<>" & q & q & ",I" & rowNo & ",Y" & rowNo & "))"

        destY(pos, 1) = formulaY
        destL(pos, 1) = formulaL
        If CStr(origY(pos, 1)) <> formulaY Or CStr(origL(pos, 1)) <> formulaL Then
            changed = True
        End If
    Next rowNo

    If changed Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        writeStarted = True
        ledger.Range("Y" & FIRST_ROW & ":Y" & lastRow).Formula = destY
        ledger.Range("L" & FIRST_ROW & ":L" & lastRow).Formula = destL
        ledger.Calculate
    End If
    MMF為替_参照式更新本体 = True

CleanUp:
    On Error Resume Next
    Application.EnableEvents = wasEvents
    Application.ScreenUpdating = wasScreen
    mFXInstalling = False
    On Error GoTo 0
    Exit Function

Failure:
    errorMessage = Err.Description
    If writeStarted Then
        On Error Resume Next
        Application.EnableEvents = False
        ledger.Range("Y" & FIRST_ROW & ":Y" & lastRow).Formula = origY
        ledger.Range("L" & FIRST_ROW & ":L" & lastRow).Formula = origL
        On Error GoTo 0
    End If
    MsgBox "MMFの為替参照式を更新できませんでした。" & vbCrLf & _
           errorMessage, vbExclamation, "外貨建MMF"
    Resume CleanUp
End Function

' 1行目のFXが取得できない原因を、日付・通貨・重複・指定レートに分けて示す。
Public Sub MMF為替_初回取引診断()
    Dim ledger As Worksheet, fx As Worksheet
    Dim dayValue As Variant, typeValue As String, ccy As String
    Dim n As Long, col As Long, dateMatches As Long, headerMatch As Long
    Dim lastCol As Long, lastRow As Long, targetDay As Long, sourceDay As Long
    Dim cellValue As Variant, rateValue As Variant, result As Variant
    Dim msg As String

    On Error GoTo DiagnosticFailed
    Set ledger = ThisWorkbook.Worksheets(MMF_LEDGER)
    Set fx = ThisWorkbook.Worksheets(MMF_FX_DATA)
    dayValue = ledger.Range("A12").Value2
    typeValue = CStr(ledger.Range("H12").Value2)
    ccy = MMF為替_通貨コード(CStr(ledger.Range("F3").Value2))
    If Not IsDate(dayValue) Then
        MsgBox "A12が正しい日付として認識されていません。", vbExclamation
        Exit Sub
    End If
    targetDay = CLng(Int(CDbl(CDate(dayValue))))
    lastCol = fx.Cells(3, fx.Columns.Count).End(xlToLeft).Column
    lastRow = fx.Cells(fx.Rows.Count, 1).End(xlUp).Row
    For col = 3 To lastCol Step 3
        If MMF為替_通貨コード(CStr(fx.Cells(3, col).Value2)) = ccy Then
            headerMatch = col
            Exit For
        End If
    Next col
    For n = 5 To lastRow
        cellValue = fx.Cells(n, 1).Value2
        If Not IsError(cellValue) Then
            If IsDate(cellValue) Then
                sourceDay = CLng(Int(CDbl(CDate(cellValue))))
                If targetDay = sourceDay Then dateMatches = dateMatches + 1
            End If
        End If
    Next n
    result = MMF為替_取得(dayValue, typeValue, ledger.Range("F3").Value2)
    msg = "取引日：" & Format$(CDate(dayValue), "yyyy/m/d") & vbCrLf & _
          "為替区分：" & typeValue & vbCrLf & "通貨コード：" & ccy & vbCrLf & _
          "為替ヘッダー列：" & CStr(headerMatch) & vbCrLf & _
          "取引日のデータ行数：" & CStr(dateMatches) & vbCrLf
    If IsError(result) Then
        msg = msg & "取得結果：取得不可（同日レートの相違・入力不備等を確認）"
    Else
        msg = msg & "取得レート：" & CStr(result)
    End If
    MsgBox msg, vbInformation, "MMF 為替データ診断"
    Exit Sub

DiagnosticFailed:
    MsgBox "為替データを診断できませんでした：" & Err.Description, vbExclamation, "外貨建MMF"
End Sub
