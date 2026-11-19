VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKamoku 
   Caption         =   "勘定科目一覧"
   ClientHeight    =   7368
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5892
   OleObjectBlob   =   "frmKamoku.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKamoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================================
' 初期表示
'========================================
Private Sub UserForm_Initialize()

    '位置調整（エラー回避付き）
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ActiveWindow.Top + ActiveCell.Top + ActiveCell.Height
    Me.Left = Application.Left + ActiveWindow.Left + ActiveCell.Left
    On Error GoTo 0

    LoadKamokuList ""

End Sub

'========================================
' マスタ読込
'========================================
Private Sub LoadKamokuList(ByVal keyword As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    '■■ シート存在チェック
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("マスタ")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "マスタシートが見つかりません", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    With lstKamoku
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "80;200"
    End With

    For i = 2 To lastRow

        If keyword = "" Or InStr(1, ws.Cells(i, "D").Value, keyword, vbTextCompare) > 0 Then

            lstKamoku.AddItem ws.Cells(i, "C").Value
            lstKamoku.List(lstKamoku.ListCount - 1, 1) = ws.Cells(i, "D").Value

        End If

    Next i

End Sub

'========================================
' 検索
'========================================
Private Sub txtSearch_Change()
    LoadKamokuList txtSearch.Text
End Sub

'========================================
' ダブルクリック
'========================================
Private Sub lstKamoku_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SetValue
End Sub

'========================================
' Enterキー
'========================================
Private Sub lstKamoku_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SetValue
    End If

End Sub

'========================================
' 値セット
'========================================
Private Sub SetValue()

    If lstKamoku.ListIndex = -1 Then Exit Sub

    '■■ ActiveCell消失対策
    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.EnableEvents = False

    With Selection.Cells(1, 1)
        .Value = lstKamoku.List(lstKamoku.ListIndex, 0)
        .Offset(0, 1).Value = lstKamoku.List(lstKamoku.ListIndex, 1)
    End With

    Application.EnableEvents = True

    Unload Me

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub
