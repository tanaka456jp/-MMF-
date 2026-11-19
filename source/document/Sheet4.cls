VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ExitHandler

    Dim wsM As Worksheet
    Dim lastRow As Long
    Dim code As String
    Dim foundCell As Range

    Dim rng As Range
    Set rng = Union(Me.Range("C10:C1000"), _
                    Me.Range("E10:E1000"), _
                    Me.Range("H10:H1000"), _
                    Me.Range("H4,I4,J4,K4"))

    Dim hit As Range
    Set hit = Intersect(Target, rng)

    '■■ J列用（科目自動表示）--------------------
    If Not Intersect(Target, Me.Range("J10:J" & Me.Rows.Count)) Is Nothing Then

        If Target.Cells.Count = 1 Then

            Application.EnableEvents = False

            code = Trim(Target.Value)

            If code = "" Then
                Target.Offset(0, 1).ClearContents
            Else
                Set wsM = ThisWorkbook.Sheets("マスタ")
                lastRow = wsM.Cells(wsM.Rows.Count, "C").End(xlUp).Row

                Set foundCell = wsM.Range("C2:C" & lastRow).Find( _
                                    What:=code, _
                                    LookAt:=xlWhole)

                If Not foundCell Is Nothing Then
                    Target.Offset(0, 1).Value = foundCell.Offset(0, 1).Value
                Else
                    Target.Offset(0, 1).Value = "該当なし"
                End If
            End If

            Application.EnableEvents = True

        End If

    End If

    '■■ チェックOFFなら履歴処理しない
    If Me.Range("AA1").Value <> True Then Exit Sub

    '■■ 範囲外は無視
    If hit Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Me.Unprotect "1234"

    Dim c As Range
    Dim msg As String
    Dim oldVal As Variant

    For Each c In hit

        '変更前値取得
        If Not OldValues Is Nothing Then
            If OldValues.exists(c.Address) Then
                oldVal = OldValues(c.Address)
            Else
                oldVal = ""
            End If
        Else
            oldVal = ""
        End If

        If c.HasFormula Then
            c.Font.Italic = False
            GoTo NextCell
        End If

        If c.Value <> "" Then

            msg = Application.UserName & " " & Format(Now, "yyyy/mm/dd HH:mm") & _
                  " / 変更前：" & oldVal

            On Error Resume Next
            If Not c.CommentThreaded Is Nothing Then
                c.CommentThreaded.Text Text:=msg
            Else
                c.AddCommentThreaded msg
            End If
            On Error GoTo 0

            c.Font.Italic = True

        Else
            c.Font.Italic = False

            On Error Resume Next
            If Not c.CommentThreaded Is Nothing Then
                c.CommentThreaded.Delete
            End If
            On Error GoTo 0
        End If

NextCell:
    Next c

ExitHandler:

    On Error Resume Next
    Me.Protect "1234", _
        UserInterfaceOnly:=True, _
        AllowDeletingRows:=Me.Range("AA1").Value, _
        AllowInsertingRows:=Me.Range("AA1").Value

    Application.EnableEvents = True

End Sub

