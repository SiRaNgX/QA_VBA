VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'========================================================
' [프로그램(성적서 업무용) 백업용 전역 변수]
'========================================================
Dim OpenBackupPath As String   ' "열 때" 생성한 백업 파일 전체 경로
Dim IsModified As Boolean      ' 수정 발생 여부(수정 없으면 백업 삭제용)


'========================================================
' [Workbook_Open]
' - 파일 열릴 때 실행
' - Excel 환경 복구(개선 5)
' - DB 연결 시도(I5 우선, 실패 시 H8 사용)
' - DB 연결 실패 시 사용자 알림(개선 1)
' - 프로그램 백업 생성 + 10개 유지(열 때 백업 / 수정 없으면 삭제)
'========================================================
Private Sub Workbook_Open()

    '----------------------------------------------------
    ' (개선 5) 엑셀 환경 복구
    ' - 이벤트가 꺼져 있거나 계산이 수동으로 남아 있으면
    '   프로그램이 "멈춘 것처럼" 보일 수 있음
    '----------------------------------------------------
    
    ' Excel 창 최대화
    Application.WindowState = xlMaximized
    
    Call RestoreExcelEnvironment

    ' 링크 업데이트 관련(기존 코드 유지)
    On Error Resume Next
    Application.AskToUpdateLinks = False
    ThisWorkbook.UpdateLinks = xlUpdateLinksAlways
    On Error GoTo 0

    '----------------------------------------------------
    ' 성적서 시트 설정 및 보호(기존 코드 유지)
    '----------------------------------------------------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("성적서")

    ws.EnableSelection = xlNoRestrictions
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True

    '----------------------------------------------------
    ' DB 경로 읽기(기존 로직)
    ' - I5: 기본 DB 경로(사용자가 관리/세팅)
    ' - H8: 마지막으로 성공했던 백업 경로(자동 기록)
    '----------------------------------------------------
    Dim path_I5 As String
    Dim path_H8 As String

    path_I5 = ws.Range("I5").Value
    path_H8 = ws.Range("H8").Value

    '----------------------------------------------------
    ' 1) I5 경로 우선 DB 연결
    '----------------------------------------------------
    gDB_Path = path_I5
    Set gDB_WB = OpenDB

    '----------------------------------------------------
    ' 2) 실패 시 H8 백업 경로로 재시도
    '----------------------------------------------------
    If gDB_WB Is Nothing Then

        If path_H8 <> "" Then

            gDB_Path = path_H8
            Set gDB_WB = OpenDB

            ' H8 경로로 성공했다면 I5도 갱신(다음부터 바로 성공하도록)
            If Not gDB_WB Is Nothing Then
                ws.Range("I5").Value = path_H8
            End If

        End If

    End If

    '----------------------------------------------------
    ' (개선 1) DB 연결 실패 시 사용자 알림
    ' - DB가 없으면 이후 매크로에서 오류가 날 수 있으니
    '   시작 시점에 확실히 알려주기
    '----------------------------------------------------
    If gDB_WB Is Nothing Then

        MsgBox "QA_DB 파일 연결에 실패했습니다." & vbCrLf & _
               "1) 성적서 시트 I5의 DB 경로 확인" & vbCrLf & _
               "2) DB 파일 존재/이름/폴더 확인" & vbCrLf & _
               "3) 접근 권한(공용드라이브 등) 확인", vbCritical

    End If

    '----------------------------------------------------
    ' [프로그램(성적서 업무용) 백업 생성]
    ' - 열 때 현재 상태를 백업(수정 전 상태 확보)
    ' - 백업은 10개 유지
    ' - 수정이 없으면 닫을 때 방금 만든 백업 삭제
    '----------------------------------------------------
    Call MakeProgramOpenBackup_Keep10

End Sub


'========================================================
' [Workbook_Activate]
' - 성적서 업무용.xlsb가 활성화될 때
' - DB가 열려있는지 재확인 후, 없으면 재연결
'========================================================
Private Sub Workbook_Activate()

    Dim wb As Workbook
    Dim found As Boolean

    found = False

    ' DB 경로 없으면 종료
    If gDB_Path = "" Then Exit Sub

    ' 현재 열린 파일 중 DB가 있는지 검사
    For Each wb In Application.Workbooks

        If wb.FullName = gDB_Path Then
            Set gDB_WB = wb
            found = True
            Exit For
        End If

    Next wb

    ' 없으면 다시 연결 시도
    If found = False Then
        Set gDB_WB = Nothing
        Call OpenDB
    End If

End Sub


'========================================================
' [Workbook_SheetChange]
' - 어떤 시트든 셀 값 변경이 발생하면 수정 플래그 ON
' - 이후 닫을 때: 수정이 있으면 열 때 만든 백업 유지
'========================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    IsModified = True

End Sub


'========================================================
' [Workbook_BeforeClose]
' - 닫기 직전
' - 수정이 없으면 "열 때 만든 백업"은 불필요하므로 삭제
'========================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If IsModified = False Then

        If OpenBackupPath <> "" Then

            If Dir(OpenBackupPath) <> "" Then
                Kill OpenBackupPath
            End If

        End If

    End If

End Sub


'========================================================
' [내부 유틸] 열 때 백업 생성 + 10개 유지
' - Workbook_Open에서 호출
'========================================================
Private Sub MakeProgramOpenBackup_Keep10()

    Dim backupFolder As String
    Dim backupPath As String

    Dim fileList() As String
    Dim fileDates() As Date

    Dim fileCount As Long
    Dim f As String

    Dim i As Long, j As Long
    Dim tempName As String
    Dim tempDate As Date

    ' 백업 폴더(기존 경로 사용)
    backupFolder = "D:\WORK\QA_DB\Backup\"

    ' 폴더 없으면 생성
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    ' 백업 파일 생성(초 단위로 중복 방지)
    backupPath = backupFolder & "성적서 업무용_" & _
                 Format(Now, "yyyymmdd_hhmmss") & ".xlsb"

    ' 현재 파일 복사본 백업
    ThisWorkbook.SaveCopyAs backupPath

    ' 닫을 때 삭제/유지 판단을 위해 경로 저장
    OpenBackupPath = backupPath

    ' "열자마자 백업 생성"이므로, 아직 수정은 없다고 초기화
    IsModified = False

    '--------------------------------------------
    ' 백업 파일 목록 수집 (성적서 업무용_*.xlsb)
    '--------------------------------------------
    f = Dir(backupFolder & "성적서 업무용_*.xlsb")
    fileCount = 0

    Do While f <> ""

        fileCount = fileCount + 1

        ReDim Preserve fileList(1 To fileCount)
        ReDim Preserve fileDates(1 To fileCount)

        fileList(fileCount) = f
        fileDates(fileCount) = FileDateTime(backupFolder & f)

        f = Dir

    Loop

    '--------------------------------------------
    ' 10개 초과 시 오래된 것 삭제
    '--------------------------------------------
    If fileCount > 10 Then

        ' 날짜 기준 정렬(오래된 → 최신)
        For i = 1 To fileCount - 1
            For j = i + 1 To fileCount

                If fileDates(i) > fileDates(j) Then

                    tempDate = fileDates(i)
                    fileDates(i) = fileDates(j)
                    fileDates(j) = tempDate

                    tempName = fileList(i)
                    fileList(i) = fileList(j)
                    fileList(j) = tempName

                End If

            Next j
        Next i

        ' 오래된 것부터 삭제하여 10개 유지
        For i = 1 To fileCount - 10
            Kill backupFolder & fileList(i)
        Next i

    End If

End Sub

