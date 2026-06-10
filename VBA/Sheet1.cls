VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub
    If Target.Row < 11 Then Exit Sub

    Application.EnableEvents = False

    ' 관리번호
    If Me.Cells(Target.Row, "A").Value = "" Then
        Me.Range("A1").Value = "-"
    Else
        Me.Range("A1").Value = Me.Cells(Target.Row, "A").Value
    End If

    ' 선택 행
    Me.Range("A8").Value = Target.Row

    Application.EnableEvents = True

End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    If Target.Address <> "$H$5" Then Exit Sub
    
    Cancel = True
    
    Dim wb As Workbook

    ' DB 경로 없으면 종료
    If gDB_Path = "" Then Exit Sub

    ' 전역 DB 객체 확인
    If gDB_WB Is Nothing Then
        Set wb = OpenDB
    Else
        Set wb = gDB_WB
    End If

    If wb Is Nothing Then Exit Sub

    ' 숨김 해제 후 활성화
    On Error Resume Next
    wb.Windows(1).Visible = True
    wb.Activate
    On Error GoTo 0

End Sub

