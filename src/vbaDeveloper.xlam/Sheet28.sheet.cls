'Dim started As Integer

Private Sub Worksheet_Activate()
    'If Not started > 0 Then
    '    started = 0
    'End If
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    'Application.EnableEvents = False
    ' Note: this is a sheet-level setting.
    'ActiveSheet.DisplayPageBreaks = False

    'If started = 0 Then
    '    UserForm_설정.Hide
    '
    '    ' #2 데이터 복사하란 안내 메시지
    '    Range("A3:P1000").ClearContents
    '    MsgBox "가져올 데이터를 이 시트 (세번째 줄)에 복사해주세요"
    '    Range("A3").Select
    'End If
    
    ' #3 시트에 데이터가 생기면, 일자, 관항목, 적요, 수입/지출, 은/현 구분 라벨 첫 줄에 넣으라는 메시지
    ' #4 첫 행에 모든 라벨 있는지 확인
    ' #5 각 열 순회하며 복사 & 회계원장에 붙이기
    ' #6 추가된 부분 순회하며 관항목 검증 : 없으면 붉은색 표시, 있으면 기존에 등록된 코드인지 확인
    ' #6-1. 신규 관항목이 들어 있으면 관항목 생성
    ' #7 회계원장 정렬

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' #3 시트에 데이터가 생기면, 일자, 관항목, 적요, 수입/지출, 은/현 구분 라벨 첫 줄에 넣으라는 메시지
    'With Worksheets("가져오기")
    '    If .Range("A3").Value <> "" And started = 0 Then
    '        MsgBox "두번째 행에 라벨을 추가해주세요(일자, 관, 항, 목, 적요, 수입, 지출, 은/현)"
    '        .Range("A2").Select
    '        started = 1
    '    End If
    'End With
End Sub

Private Sub Worksheet_Deactivate()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    'Application.EnableEvents = False
    ' Note: this is a sheet-level setting.
    'ActiveSheet.DisplayPageBreaks = True
End Sub