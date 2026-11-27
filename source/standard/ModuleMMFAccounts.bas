Attribute VB_Name = "ModuleMMFAccounts"
Option Explicit

' 外貨建MMFを一般口座・特定口座で別管理し、合計残高の一覧表を作る。
' 元の「MMF取引台帳」は移行後も非常時の原本として保持（非表示）。
' 各口座の取得費計算式・移動平均法・円換算は既存の台帳をそのまま利用。
Private Const LEGACY_LEDGER As String = "MMF取引台帳"
Private Const GENERAL_LEDGER As String = "一般口座"
Private Const SPECIFIC_LEDGER As String = "特定口座"
Private Const SUMMARY_SHEET As String = "口座別まとめ"
Private Const FIRST_TRADE As Long = 12
Private Const MIN_LAST_TRADE As Long = 511
Private Const MAX_TRADE_ROWS As Long = 5000
Private Const FIRST_SUMMARY_ROW As Long = 6
Private mSummaryRunning As Boolean
Private mSetupRunning As Boolean

' 起動時に実行。データの所属が未確定なら初回だけ確認する。
Public Sub MMF口座_起動時設定()
    If Not MMF口座_初期設定(False) Then Exit Sub
    ' 必要な条件が整えばブック起動時にまとめ表も更新する。
    Call MMF口座_まとめ更新(True)
End Sub

' Alt+F8から任意実行（アップデーター適用直後はこちらでも可）。
Public Sub MMF口座_2口座シート作成()
    If MMF口座_初期設定(True) Then
        MsgBox "「一般口座」「特定口座」を利用できます。" & vbCrLf & _
               "まとめ表は「MMF口座_まとめ表更新」で作成してください。", _
               vbInformation, "外貨建MMF"
    End If
End Sub

Private Function MMF口座_シート取得(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set MMF口座_シート取得 = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function MMF口座_初期設定(ByVal showErrors As Boolean) As Boolean
    Dim source As Worksheet, gen As Worksheet, spec As Worksheet
    Dim carryToGeneral As Boolean, haveTrades As Boolean
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean
    Dim createdGeneral As Boolean, createdSpecific As Boolean
    Dim failure As String

    If mSetupRunning Then Exit Function
    mSetupRunning = True
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failure

    Set gen = MMF口座_シート取得(GENERAL_LEDGER)
    Set spec = MMF口座_シート取得(SPECIFIC_LEDGER)
    If Not gen Is Nothing And Not spec Is Nothing Then
        MMF口座_初期設定 = True
        GoTo CleanExit
    End If
    If (gen Is Nothing) Xor (spec Is Nothing) Then
        Err.Raise vbObjectError + 811, , _
                  "口座シートが片方だけ存在します。既存のシートを保護するため自動追加を中止しました。"
    End If

    Set source = MMF口座_シート取得(LEGACY_LEDGER)
    If source Is Nothing Then
        Err.Raise vbObjectError + 812, , _
                  "元のMMF取引台帳が見つかりません。"
    End If
    If ThisWorkbook.ProtectStructure Then
        Err.Raise vbObjectError + 813, , _
                  "ブックの構成が保護されています。保護を解除してください。"
    End If
    If CStr(source.Range("C11").Value2) <> "買付数量" Or _
       CStr(source.Range("E11").Value2) <> "売付数量" Then
        Err.Raise vbObjectError + 814, , _
                  "元台帳の列構成が異なります。自動移行は中止しました。"
    End If

    haveTrades = (Application.CountA(source.Range("A12:J511")) > 0)
    carryToGeneral = True
    If haveTrades Then
        answer = MsgBox("初回のみ、既存「MMF取引台帳」の取引をどちらへ引き継ぎますか？" & _
                        vbCrLf & vbCrLf & _
                        "［はい］＝一般口座" & vbCrLf & _
                        "［いいえ］＝特定口座" & vbCrLf & _
                        "［キャンセル］＝今回は変更しない" & vbCrLf & vbCrLf & _
                        "元の台帳はバックアップとして保持し、取引を重複集計しません。", _
                        vbYesNoCancel + vbQuestion + vbDefaultButton3, _
                        "外貨建MMF 口座別シートの初期移行")
        If answer = vbCancel Then GoTo CleanExit
        carryToGeneral = (answer = vbYes)
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    source.Visible = xlSheetVisible
    source.Copy After:=source
    Set gen = ActiveSheet
    gen.Name = GENERAL_LEDGER
    createdGeneral = True

    source.Copy After:=gen
    Set spec = ActiveSheet
    spec.Name = SPECIFIC_LEDGER
    createdSpecific = True

    ' 元の入力データが所属しない口座のみクリア。数式・書式は残す。
    If haveTrades Then
        If carryToGeneral Then
            MMF口座_台帳入力初期化 spec
        Else
            MMF口座_台帳入力初期化 gen
        End If
    Else
        MMF口座_台帳入力初期化 gen
        MMF口座_台帳入力初期化 spec
    End If

    gen.Range("B12:B511").Validation.Delete
    spec.Range("B12:B511").Validation.Delete
    gen.Range("B12:B511").NumberFormat = "@"
    spec.Range("B12:B511").NumberFormat = "@"
    gen.Tab.Color = RGB(28, 109, 146)
    spec.Tab.Color = RGB(89, 115, 158)

    ' 非表示の原本はVBAUpdater互換のため名前を保持する。
    gen.Activate
    source.Visible = xlSheetVeryHidden
    MMF口座_初期設定 = True

CleanExit:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    mSetupRunning = False
    On Error GoTo 0
    Exit Function

Failure:
    failure = Err.Description
    On Error Resume Next
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If createdSpecific Then spec.Delete
    If createdGeneral Then gen.Delete
    Application.DisplayAlerts = oldAlerts
    If Not source Is Nothing Then source.Visible = xlSheetVisible
    On Error GoTo 0
    If showErrors Then
        MsgBox "口座シートの作成を中止しました。" & vbCrLf & failure, _
               vbExclamation, "外貨建MMF"
    Else
        MsgBox "口座別管理の初期設定を完了できませんでした。" & vbCrLf & failure, _
               vbExclamation, "外貨建MMF"
    End If
    Resume CleanExit
End Function

Private Sub MMF口座_台帳入力初期化(ByVal ws As Worksheet)
    ws.Range("A12:J511").ClearContents
    ws.Range("B4").Value2 = 0
    ws.Range("D4").Value2 = 0
    ws.Range("F4").Value2 = 0
End Sub

' 既存の別ファイルで保有している口座取引を、空の口座シートへ取り込む。
' 上書き・結合は行わない。移行元のファイルは常に読み取り専用。
Public Sub MMF口座_別ブックから取込()
    Dim chooser As Object
    Dim sourceBook As Workbook
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim sourceName As String, targetName As String, selectedFile As String
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean, oldScreen As Boolean, oldSecurity As Long
    Dim sourceLast As Long, lastTarget As Long
    Dim originalInput As Variant, originalStart As Variant
    Dim dataWritten As Boolean, errorText As String, appStateSaved As Boolean

    On Error GoTo Failed
    If MMF口座_シート取得(GENERAL_LEDGER) Is Nothing Or _
       MMF口座_シート取得(SPECIFIC_LEDGER) Is Nothing Then
        Err.Raise vbObjectError + 841, , _
                  "一般口座・特定口座のシートを先に作成してください。"
    End If

    answer = MsgBox("別ファイルの取引をどちらへ取り込みますか？" & vbCrLf & _
                    "［はい］＝一般口座、［いいえ］＝特定口座" & vbCrLf & _
                    "［キャンセル］＝中止", vbYesNoCancel + vbQuestion + _
                    vbDefaultButton3, "MMF 既存取引の取り込み")
    If answer = vbCancel Then Exit Sub
    If answer = vbYes Then targetName = GENERAL_LEDGER Else targetName = SPECIFIC_LEDGER
    Set targetWs = ThisWorkbook.Worksheets(targetName)

    ' 既存入力への無断上書きは行わず、空の台帳に限って取り込む。
    lastTarget = MMF口座_最終入力行(targetWs)
    If Application.CountA(targetWs.Range("A12:J" & lastTarget)) > 0 Then
        Err.Raise vbObjectError + 842, , _
                  targetName & "に既存取引があります。上書きはしません。"
    End If
    If targetWs.ProtectContents Then
        Err.Raise vbObjectError + 843, , _
                  targetName & "のシート保護を解除してください。"
    End If
    If Val(CStr(targetWs.Range("B4").Value2)) <> 0 Or _
       Val(CStr(targetWs.Range("D4").Value2)) <> 0 Or _
       Val(CStr(targetWs.Range("F4").Value2)) <> 0 Then
        Err.Raise vbObjectError + 849, , _
                  targetName & "に期首残高があります。上書きせず中止しました。"
    End If

    Set chooser = Application.FileDialog(3)
    chooser.AllowMultiSelect = False
    chooser.Title = "取り込むMMF取得費計算ファイルを選択してください"
    chooser.Filters.Clear
    chooser.Filters.Add "Excelブック", "*.xlsm;*.xlsx"
    If chooser.Show <> -1 Then Exit Sub
    selectedFile = CStr(chooser.SelectedItems(1))
    If StrComp(selectedFile, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 844, , "現在のブックは移行元として選択できません。"
    End If

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldSecurity = Application.AutomationSecurity
    appStateSaved = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AutomationSecurity = 3   ' msoAutomationSecurityForceDisable

    Set sourceBook = Application.Workbooks.Open( _
        Filename:=selectedFile, UpdateLinks:=0, ReadOnly:=True)
    Set sourceWs = Nothing
    On Error Resume Next
    Set sourceWs = sourceBook.Worksheets(targetName)
    If sourceWs Is Nothing Then Set sourceWs = sourceBook.Worksheets(LEGACY_LEDGER)
    On Error GoTo Failed
    If sourceWs Is Nothing Then
        Err.Raise vbObjectError + 845, , _
                  "取り込み元に「" & targetName & "」または「MMF取引台帳」がありません。"
    End If
    If CStr(sourceWs.Range("C11").Value2) <> "買付数量" Or _
       CStr(sourceWs.Range("E11").Value2) <> "売付数量" Then
        Err.Raise vbObjectError + 846, , "取り込み元の台帳列構成が異なります。"
    End If
    If Trim$(CStr(sourceWs.Range("B3").Value2)) <> _
       Trim$(CStr(targetWs.Range("B3").Value2)) Or _
       Trim$(CStr(sourceWs.Range("F3").Value2)) <> _
       Trim$(CStr(targetWs.Range("F3").Value2)) Then
        Err.Raise vbObjectError + 847, , _
                  "銘柄または通貨が一致しません。別銘柄の取引は取り込みません。"
    End If
    sourceLast = MMF口座_最終入力行(sourceWs)
    If sourceLast > MIN_LAST_TRADE Then
        Err.Raise vbObjectError + 848, , _
                  "500行を超える入力があります。元台帳の分割を検討してください。"
    End If

    originalInput = targetWs.Range("A12:J511").Value2
    originalStart = Array(targetWs.Range("B4").Value2, _
                          targetWs.Range("D4").Value2, _
                          targetWs.Range("F4").Value2)
    dataWritten = True
    targetWs.Range("A12:J511").Value2 = sourceWs.Range("A12:J511").Value2
    targetWs.Range("B4").Value2 = sourceWs.Range("B4").Value2
    targetWs.Range("D4").Value2 = sourceWs.Range("D4").Value2
    targetWs.Range("F4").Value2 = sourceWs.Range("F4").Value2
    targetWs.Calculate
    dataWritten = False

    MsgBox targetName & "に取引データを取り込みました。" & vbCrLf & _
           "移行元ファイルには書き込んでいません。" & vbCrLf & _
           "K列の入力判定を確認後、まとめ表を更新してください。", _
           vbInformation, "外貨建MMF"

CleanUp:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    If appStateSaved Then
        Application.AutomationSecurity = oldSecurity
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    Exit Sub

Failed:
    errorText = Err.Description
    If dataWritten Then
        On Error Resume Next
        targetWs.Range("A12:J511").Value2 = originalInput
        targetWs.Range("B4").Value2 = originalStart(0)
        targetWs.Range("D4").Value2 = originalStart(1)
        targetWs.Range("F4").Value2 = originalStart(2)
        On Error GoTo 0
    End If
    MsgBox "既存の別ブックから取り込めませんでした。" & vbCrLf & _
           errorText, vbExclamation, "外貨建MMF"
    Resume CleanUp
End Sub

Public Sub MMF口座_まとめ表更新()
    If MMF口座_まとめ更新(False) Then
        ThisWorkbook.Worksheets(SUMMARY_SHEET).Activate
        MsgBox "一般口座・特定口座の取引からまとめ表を更新しました。", vbInformation
    End If
End Sub

Public Sub MMF口座_まとめ表印刷()
    Dim ws As Worksheet, returnSheet As Worksheet
    On Error GoTo PrintFailed
    If ActiveWorkbook Is ThisWorkbook Then
        If TypeName(ActiveSheet) = "Worksheet" Then Set returnSheet = ActiveSheet
    End If
    If Not MMF口座_まとめ更新(False) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    ws.Select Replace:=True
    Application.Dialogs(xlDialogPrint).Show
CleanExit:
    On Error Resume Next
    If Not returnSheet Is Nothing Then returnSheet.Activate
    On Error GoTo 0
    Exit Sub
PrintFailed:
    MsgBox "まとめ表の印刷ダイアログを開けませんでした。" & vbCrLf & _
           Err.Description, vbExclamation, "外貨建MMF"
    Resume CleanExit
End Sub

' Ctrl+Pで直接印刷する場合にも最新データへ更新する。
Public Function MMF口座_まとめ更新(Optional ByVal quiet As Boolean = True) As Boolean
    Dim gen As Worksheet, spec As Worksheet, out As Worksheet
    Dim d() As Double, acct() As Long, sourceRows() As Long
    Dim allRows As Long, n As Long, idx As Long, lastRow As Long
    Dim accountNo As Long, rowNo As Long, prevYear As Long, yearNo As Long
    Dim rowIndex As Long, col As Long
    Dim sourceWs As Worksheet, eventDate As Double, lastAccountDate As Double
    Dim qty(1 To 2) As Double, cost(1 To 2) As Double
    Dim delta As Double, amount As Double, unitValue As Double
    Dim opening As Variant, v As Variant, contents As Variant
    Dim records As Collection, output() As Variant, rowValues As Variant
    Dim lastOut As Long, invalidRow As String, errText As String
    Dim name1 As String, name2 As String, ccy1 As String, ccy2 As String
    Dim oldEvents As Boolean, oldScreen As Boolean
    Dim valuePerUnit As Double

    If mSummaryRunning Then
        MMF口座_まとめ更新 = True
        Exit Function
    End If
    mSummaryRunning = True
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failure

    Set gen = MMF口座_シート取得(GENERAL_LEDGER)
    Set spec = MMF口座_シート取得(SPECIFIC_LEDGER)
    If gen Is Nothing Or spec Is Nothing Then
        Err.Raise vbObjectError + 821, , _
                  "口座シートがありません。「MMF口座_2口座シート作成」を先に実行してください。"
    End If

    name1 = Trim$(CStr(gen.Range("B3").Value2))
    name2 = Trim$(CStr(spec.Range("B3").Value2))
    ccy1 = Trim$(CStr(gen.Range("F3").Value2))
    ccy2 = Trim$(CStr(spec.Range("F3").Value2))
    If name1 = "" Or name1 <> name2 Or ccy1 = "" Or ccy1 <> ccy2 Then
        Err.Raise vbObjectError + 822, , _
                  "2口座のB3（銘柄）・F3（通貨）を同じ値に設定してください。"
    End If

    Set out = MMF口座_シート取得(SUMMARY_SHEET)
    valuePerUnit = 0.01
    If Not out Is Nothing Then
        v = out.Range("K2").Value2
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then valuePerUnit = CDbl(v)
        End If
    End If

    ReDim d(1 To 2 * MAX_TRADE_ROWS)
    ReDim acct(1 To 2 * MAX_TRADE_ROWS)
    ReDim sourceRows(1 To 2 * MAX_TRADE_ROWS)
    Set records = New Collection
    For accountNo = 1 To 2
        If accountNo = 1 Then
            Set sourceWs = gen
        Else
            Set sourceWs = spec
        End If

        If Not IsNumeric(sourceWs.Range("B4").Value2) Or _
           Not IsNumeric(sourceWs.Range("F4").Value2) Then
            Err.Raise vbObjectError + 823, , _
                      sourceWs.Name & "の期首数量または円貨簿価が数値ではありません。"
        End If
        qty(accountNo) = CDbl(sourceWs.Range("B4").Value2)
        cost(accountNo) = CDbl(sourceWs.Range("F4").Value2)
        If qty(accountNo) < 0 Or cost(accountNo) < 0 Then
            Err.Raise vbObjectError + 824, , _
                      sourceWs.Name & "の期首残高が負の数です。"
        End If

        sourceWs.Calculate
        lastAccountDate = 0#
        lastRow = MMF口座_最終入力行(sourceWs)
        If lastRow > FIRST_TRADE + MAX_TRADE_ROWS - 1 Then
            Err.Raise vbObjectError + 825, , _
                      sourceWs.Name & "の取引数が上限を超えました。"
        End If
        For rowNo = FIRST_TRADE To lastRow
            If Application.CountA(sourceWs.Range("A" & rowNo & ":J" & rowNo)) > 0 Then
                v = sourceWs.Cells(rowNo, "K").Value2
                If IsError(v) Then GoTo InvalidTransaction
                If CStr(v) <> "OK" Then GoTo InvalidTransaction
                v = sourceWs.Cells(rowNo, "A").Value2
                If Not IsNumeric(v) Then GoTo InvalidTransaction
                If CDbl(v) < 61# Then GoTo InvalidTransaction
                If CDbl(v) < lastAccountDate Then
                    Err.Raise vbObjectError + 829, , _
                              sourceWs.Name & "!" & CStr(rowNo) & _
                              "の日付が前行よりも前です。日付順に入力してください。"
                End If
                lastAccountDate = CDbl(v)
                If Not IsNumeric(sourceWs.Cells(rowNo, "T").Value2) Or _
                   Not IsNumeric(sourceWs.Cells(rowNo, "V").Value2) Then
                    GoTo InvalidTransaction
                End If
                n = n + 1
                d(n) = CDbl(Fix(CDbl(v)))
                acct(n) = accountNo
                sourceRows(n) = rowNo
            End If
        Next rowNo
    Next accountNo

    If n > 1 Then MMF口座_イベント整列 d, acct, sourceRows, 1, n

    If qty(1) + qty(2) > 0 Then
        If n > 0 Then
            eventDate = d(1) - 1#
        Else
            eventDate = CDbl(Date)
        End If
        MMF口座_行追加 records, eventDate, 0#, "期首残高（口座別簿価）", _
                           qty(2), qty(1), cost(1) + cost(2), valuePerUnit, True
    End If

    prevYear = 0
    For idx = 1 To n
        eventDate = d(idx)
        yearNo = Year(CDate(eventDate))
        If prevYear > 0 And yearNo > prevYear Then
            Do While prevYear < yearNo
                MMF口座_行追加 records, CDbl(DateSerial(prevYear, 12, 31)), 0#, _
                    CStr(prevYear) & "年12月31日現在", _
                    qty(2), qty(1), cost(1) + cost(2), valuePerUnit, True
                prevYear = prevYear + 1
            Loop
        End If
        prevYear = yearNo

        accountNo = acct(idx)
        rowNo = sourceRows(idx)
        If accountNo = 1 Then Set sourceWs = gen Else Set sourceWs = spec
        delta = 0#
        v = sourceWs.Cells(rowNo, "C").Value2
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then delta = CDbl(v)
        End If
        v = sourceWs.Cells(rowNo, "E").Value2
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then delta = -CDbl(v)
        End If

        qty(accountNo) = CDbl(sourceWs.Cells(rowNo, "T").Value2)
        cost(accountNo) = CDbl(sourceWs.Cells(rowNo, "V").Value2)
        If qty(accountNo) < -0.00000001 Then
            Err.Raise vbObjectError + 826, , sourceWs.Name & "!" & rowNo & "の残高数量が負です。"
        End If

        contents = sourceWs.Cells(rowNo, "B").Value2
        If IsError(contents) Then contents = ""
        invalidRow = Trim$(CStr(contents))
        If invalidRow = "" Then
            If delta >= 0 Then invalidRow = "買付" Else invalidRow = "売付"
        End If
        contents = sourceWs.Cells(rowNo, "J").Value2
        If Not IsError(contents) Then
            If Len(Trim$(CStr(contents))) > 0 Then _
                invalidRow = invalidRow & "／" & Trim$(CStr(contents))
        End If

        If accountNo = 1 Then
            MMF口座_行追加 records, eventDate, eventDate, invalidRow, _
                              0#, delta, cost(1) + cost(2), valuePerUnit, False, qty(2), qty(1)
        Else
            MMF口座_行追加 records, eventDate, eventDate, invalidRow, _
                              delta, 0#, cost(1) + cost(2), valuePerUnit, False, qty(2), qty(1)
        End If
    Next idx

    If prevYear > 0 Then
        MMF口座_行追加 records, CDbl(DateSerial(prevYear, 12, 31)), 0#, _
                    CStr(prevYear) & "年12月31日現在", _
                    qty(2), qty(1), cost(1) + cost(2), valuePerUnit, True
    End If
    ' 数量・円貨簿価は口座ごとに計算。合計取得単価は表示用加重平均。
    ' 口座をまたぐ移管を売買入力だけで代用しないこと（別途調整が必要）。

    If out Is Nothing Then
        Set out = ThisWorkbook.Worksheets.Add(After:=spec)
        out.Name = SUMMARY_SHEET
    End If
    If out.ProtectContents Then
        Err.Raise vbObjectError + 827, , _
                  "まとめ表のシート保護を解除してください。"
    End If
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    out.UsedRange.UnMerge
    out.Cells.Clear
    With out
        .Cells.Font.Name = "Yu Gothic"
        .Cells.Font.Size = 9
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 13
        .Columns("C").ColumnWidth = 38
        .Columns("D").ColumnWidth = 17
        .Columns("E").ColumnWidth = 17
        .Columns("F").ColumnWidth = 17
        .Columns("G").ColumnWidth = 18
        .Columns("H").ColumnWidth = 18
        .Columns("I").ColumnWidth = 2
        .Columns("J").ColumnWidth = 22
        .Columns("K").ColumnWidth = 25

        .Range("A1:H1").Merge
        .Range("A1").Value = name1 & "　口座別保有・取得費まとめ表"
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Font.Size = 14
        .Range("A1:H1").Font.Color = vbWhite
        .Range("A1:H1").Interior.Color = RGB(24, 56, 78)
        .Range("A1").HorizontalAlignment = xlLeft
        .Rows(1).RowHeight = 31

        .Range("A2:H2").Merge
        .Range("A2").Value = "通貨：" & ccy1 & _
            "　　特定口座・一般口座の増減と年末現在残高を表示"
        .Range("A2:H2").Interior.Color = RGB(229, 239, 245)
        .Range("A2").HorizontalAlignment = xlLeft
        .Rows(2).RowHeight = 23

        .Range("A3:H3").Merge
        .Range("A3").Value = "年末現在行は口座別残高、取引行は口座別増減です。取得単価は2口座合計の円貨簿価÷合計数量（参考値）。"
        .Range("A3").HorizontalAlignment = xlLeft
        .Range("A3:H3").Font.Size = 8
        .Rows(3).RowHeight = 19
        .Range("A4:H4").Merge
        .Range("A4").Value = "注意：口座移管は単なる売買として入力すると取得費が変動します。移管時は別途整合確認が必要です。"
        .Range("A4").HorizontalAlignment = xlLeft
        .Range("A4:H4").Font.Color = RGB(153, 65, 15)
        .Range("A4:H4").Font.Size = 8
        .Rows(4).RowHeight = 19

        .Range("A5:H5").Value = Array("取引開始日", "終了日", "備考／取引区分", _
                    "特定口座（口）", "一般口座（口）", "取得単価（円）", _
                    "合計残高（口）", "合計残高（外貨）")
        .Range("A5:H5").Interior.Color = RGB(31, 70, 95)
        .Range("A5:H5").Font.Color = vbWhite
        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").HorizontalAlignment = xlCenter
        .Rows(5).RowHeight = 27

        .Range("J1:K1").Merge
        .Range("J1").Value = "まとめ表の設定（印刷対象外）"
        .Range("J1:K1").Interior.Color = RGB(31, 70, 95)
        .Range("J1:K1").Font.Color = vbWhite
        .Range("J2").Value = "1口当たり外貨額"
        .Range("K2").Value2 = valuePerUnit
        .Range("K2").NumberFormat = "0.########"
        .Range("K2").Interior.Color = RGB(255, 244, 202)
        .Range("J3").Value = "集計元"
        .Range("K3").Value = "特定口座 ＋ 一般口座"
        .Range("J4").Value = "集計取引件数"
        .Range("K4").Value = n

        If records.Count > 0 Then
            ReDim output(1 To records.Count, 1 To 8)
            For rowIndex = 1 To records.Count
                rowValues = records.Item(rowIndex)
                For col = 1 To 8
                    output(rowIndex, col) = rowValues(col)
                Next col
            Next rowIndex
            .Range("A" & FIRST_SUMMARY_ROW).Resize(records.Count, 8).Value = output
            lastOut = FIRST_SUMMARY_ROW + records.Count - 1
            With .Range("A6:H" & lastOut)
                .RowHeight = 20
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 209, 217)
                .Borders.Weight = xlHairline
            End With
            .Range("A6:B" & lastOut).NumberFormat = "yyyy/m/d"
            .Range("D6:E" & lastOut).NumberFormat = "#,##0;[Red](#,##0);-"
            .Range("F6:F" & lastOut).NumberFormat = "0.########"
            .Range("G6:G" & lastOut).NumberFormat = "#,##0"
            .Range("H6:H" & lastOut).NumberFormat = "#,##0"
            .Range("C6:C" & lastOut).HorizontalAlignment = xlLeft
            .Range("D6:H" & lastOut).HorizontalAlignment = xlRight

            For rowIndex = 1 To records.Count
                rowValues = records.Item(rowIndex)
                If CBool(rowValues(0)) Then
                    With .Range("A" & FIRST_SUMMARY_ROW + rowIndex - 1 & _
                                ":H" & FIRST_SUMMARY_ROW + rowIndex - 1)
                        .Interior.Color = RGB(224, 235, 244)
                        .Font.Bold = True
                    End With
                ElseIf rowIndex Mod 2 = 0 Then
                    .Range("A" & FIRST_SUMMARY_ROW + rowIndex - 1 & _
                            ":H" & FIRST_SUMMARY_ROW + rowIndex - 1).Interior.Color = RGB(246, 249, 252)
                End If
            Next rowIndex
        Else
            lastOut = FIRST_SUMMARY_ROW
            .Range("A6:H6").Merge
            .Range("A6").Value = "取引は登録されていません。"
            .Range("A6").HorizontalAlignment = xlLeft
        End If
        .Tab.Color = RGB(29, 130, 113)
        .Activate
        .PageSetup.PrintArea = "$A$1:$H$" & lastOut
        .PageSetup.PrintTitleRows = "$1:$5"
        With .PageSetup
            .PaperSize = xlPaperA4
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftMargin = Application.CentimetersToPoints(0.5)
            .RightMargin = Application.CentimetersToPoints(0.5)
            .TopMargin = Application.CentimetersToPoints(0.6)
            .BottomMargin = Application.CentimetersToPoints(0.6)
            .HeaderMargin = Application.CentimetersToPoints(0.2)
            .FooterMargin = Application.CentimetersToPoints(0.3)
            .CenterHorizontally = True
            .CenterFooter = "外貨建MMF 口座別まとめ　&P / &N"
            .RightFooter = "&D"
            .PrintGridlines = False
        End With
    End With

    MMF口座_ボタン out, out.Range("J6:K7"), "MMF_SummaryRefresh", _
                       "まとめ表更新", "MMF口座_まとめ表更新"
    MMF口座_ボタン out, out.Range("J9:K10"), "MMF_SummaryPrint", _
                       "A4まとめ印刷", "MMF口座_まとめ表印刷"
    MMF口座_ボタン out, out.Range("J12:K13"), "MMF_SummaryImport", _
                       "別ブック取込", "MMF口座_別ブックから取込"

    MMF口座_まとめ更新 = True

Finish:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    mSummaryRunning = False
    On Error GoTo 0
    Exit Function

InvalidTransaction:
    Err.Raise vbObjectError + 828, , _
              sourceWs.Name & "!" & CStr(rowNo) & _
              "の入力判定を確認してください。K列がOKでなければ集計しません。"

Failure:
    errText = Err.Description
    If Not quiet Then
        MsgBox "口座別まとめ表を更新できませんでした。" & vbCrLf & errText, _
               vbExclamation, "外貨建MMF"
    End If
    MMF口座_まとめ更新 = False
    Resume Finish
End Function

Private Function MMF口座_最終入力行(ByVal ws As Worksheet) As Long
    Dim col As Long, rr As Long, last As Long
    last = MIN_LAST_TRADE
    For col = 1 To 10
        rr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If rr > last Then last = rr
    Next col
    MMF口座_最終入力行 = last
End Function

' dateEnd=0: 年末現在行、dateEnd=dateStart: 取引行。
' 資産数量（第7列）は直近の口座別残高の合計で表示する。
Private Sub MMF口座_行追加(ByVal records As Collection, ByVal dateStart As Double, _
                          ByVal dateEnd As Double, ByVal notes As String, _
                          ByVal specific As Double, ByVal general As Double, _
                          ByVal totalCost As Double, ByVal valuePerUnit As Double, _
                          ByVal snapshot As Boolean, _
                          Optional ByVal specificBalance As Double = -1#, _
                          Optional ByVal generalBalance As Double = -1#)
    Dim total As Double, avgPrice As Variant
    If snapshot Then
        total = specific + general
    Else
        total = specificBalance + generalBalance
    End If
    avgPrice = ""
    If total > 0 Then avgPrice = Round(totalCost / total, 9)
    If snapshot Then
        records.Add Array(True, CDate(dateStart), Empty, notes, specific, general, _
                          avgPrice, total, Application.WorksheetFunction.Round(total * valuePerUnit, 0))
    Else
        records.Add Array(False, CDate(dateStart), CDate(dateEnd), notes, specific, general, _
                          avgPrice, total, Application.WorksheetFunction.Round(total * valuePerUnit, 0))
    End If
End Sub

Private Function MMF口座_比較(ByVal da As Double, ByVal aa As Long, ByVal ra As Long, _
                             ByVal db As Double, ByVal ab As Long, ByVal rb As Long) As Long
    If da < db Then
        MMF口座_比較 = -1
    ElseIf da > db Then
        MMF口座_比較 = 1
    ElseIf aa < ab Then
        MMF口座_比較 = -1
    ElseIf aa > ab Then
        MMF口座_比較 = 1
    ElseIf ra < rb Then
        MMF口座_比較 = -1
    ElseIf ra > rb Then
        MMF口座_比較 = 1
    End If
End Function

Private Sub MMF口座_イベント整列(ByRef d() As Double, ByRef acct() As Long, _
                               ByRef rr() As Long, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, k As Long
    Dim pd As Double, pa As Long, pr As Long
    Dim td As Double, ta As Long, tr As Long
    i = low
    j = high
    k = (low + high) \ 2
    pd = d(k): pa = acct(k): pr = rr(k)
    Do While i <= j
        Do While MMF口座_比較(d(i), acct(i), rr(i), pd, pa, pr) < 0
            i = i + 1
        Loop
        Do While MMF口座_比較(d(j), acct(j), rr(j), pd, pa, pr) > 0
            j = j - 1
        Loop
        If i <= j Then
            td = d(i): d(i) = d(j): d(j) = td
            ta = acct(i): acct(i) = acct(j): acct(j) = ta
            tr = rr(i): rr(i) = rr(j): rr(j) = tr
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then MMF口座_イベント整列 d, acct, rr, low, j
    If i < high Then MMF口座_イベント整列 d, acct, rr, i, high
End Sub

Private Sub MMF口座_ボタン(ByVal ws As Worksheet, ByVal place As Range, _
                        ByVal shapeName As String, ByVal label As String, _
                        ByVal action As String)
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(shapeName)
    On Error GoTo 0
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            place.Left + 2, place.Top + 2, place.Width - 4, place.Height - 4)
        s.Name = shapeName
    End If
    With s
        .Left = place.Left + 2
        .Top = place.Top + 2
        .Width = place.Width - 4
        .Height = place.Height - 4
        .Placement = xlMove
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
        .Fill.ForeColor.RGB = RGB(31, 90, 130)
        .Line.ForeColor.RGB = RGB(20, 63, 94)
        .TextFrame.Characters.Text = label
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
