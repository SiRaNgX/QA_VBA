Attribute VB_Name = "Module_resetEvt"
Option Explicit

'========================================================
' [Excel 환경 복구]
' - 이벤트/화면갱신/알림/계산모드가 매크로 실행 중 꺼진 상태로
'   남아버리는 사고를 복구하기 위한 루틴
' - "이벤트 활성화" 버튼에서 호출
' - Workbook_Open에서 자동 호출(권장)
'========================================================
Public Sub RestoreExcelEnvironment()

    Application.EnableEvents = True                  ' 이벤트 다시 켜기
    Application.ScreenUpdating = True                ' 화면 갱신 켜기
    Application.DisplayAlerts = True                 ' 경고/알림 표시 켜기
    Application.Calculation = xlCalculationAutomatic ' 계산 모드 자동

End Sub
