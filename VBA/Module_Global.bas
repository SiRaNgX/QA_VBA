Attribute VB_Name = "Module_Global"
Option Explicit

'========================================================
' [전역 변수]
' - 성적서 업무용.xlsb 어디서든 접근 가능한 DB 정보
'========================================================

Public gDB_Path As String      ' QA_DB.xlsb의 전체 경로(FullName)
Public gDB_WB As Workbook      ' 열린 QA_DB Workbook 객체 (핸들)
