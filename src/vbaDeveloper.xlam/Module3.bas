Attribute VB_Name = "Module3"
'디버깅을 위한 코드
'주요 상태 체크 값을 위한 루틴들 이곳에 작성

Option Explicit

Sub Reset_Used_Range()
    Dim A As Integer
    A = ActiveSheet.UsedRange.Rows.Count
End Sub

Sub check_settlement()
'결산서 초기화가 필요한지 체크
'계정과목/예산 설정을 매뉴얼대로 하지 않을 때 결산서에 나오는 Div/0! 메시지 등이 나오는 상황을 원천 예방

End Sub

'회계원장 체크
Sub check_ledger()

End Sub

'회계원장 문제 수정
Sub repair_ledger()

End Sub

'관항목/예산서 체크
'함수 이름 변경 필요성 확인
Sub check_accounts()

End Sub
