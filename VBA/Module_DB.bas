Attribute VB_Name = "Module_DB"
Option Explicit

'========================================================
' [OpenDB]
'
' 기능
' - gDB_Path 경로 기준으로 QA_DB 연결
' - 이미 열려있으면 재사용
' - 없으면 새로 열기
' - DB 창은 성적서 프로그램에서만 숨김
'
' 안정화
' - Workbooks.Open 오류 방지
' - Windows 객체 존재 확인
'========================================================
Public Function OpenDB() As Workbook

    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' 성적서 시트 참조
    Set ws = ThisWorkbook.Sheets("성적서")

    ' DB 경로 없으면 종료
    If gDB_Path = "" Then Exit Function


    '----------------------------------------------------
    ' 1) 이미 열려있는지 확인
    '----------------------------------------------------
    For Each wb In Application.Workbooks
        
        If wb.FullName = gDB_Path Then
            
            ' 전역 변수 설정
            Set gDB_WB = wb
            
            ' DB 창 숨김 (안전 처리)
            If gDB_WB.Windows.Count > 0 Then
                gDB_WB.Windows(1).Visible = False
            End If
            
            ' 반환값
            Set OpenDB = gDB_WB
            
            ' 마지막 성공 경로 기록
            ws.Range("H8").Value = gDB_Path
            
            Exit Function
            
        End If
        
    Next wb


    '----------------------------------------------------
    ' 2) DB 새로 열기
    '----------------------------------------------------
    On Error Resume Next
    Set gDB_WB = Workbooks.Open(gDB_Path)
    On Error GoTo 0

    ' 열기 실패 시 종료
    If gDB_WB Is Nothing Then Exit Function


    '----------------------------------------------------
    ' 3) DB 창 숨김
    '----------------------------------------------------
    If gDB_WB.Windows.Count > 0 Then
        gDB_WB.Windows(1).Visible = False
    End If


    ' 반환값 설정
    Set OpenDB = gDB_WB


    '----------------------------------------------------
    ' 4) 마지막 성공 경로 저장
    '----------------------------------------------------
    ws.Range("H8").Value = gDB_Path


End Function
