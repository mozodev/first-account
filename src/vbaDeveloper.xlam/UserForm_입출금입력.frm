VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_입출금입력 
   Caption         =   "회계장부 입력 "
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550.001
   OleObjectBlob   =   "UserForm_입출금입력.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_입출금입력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private save_error As Integer
Const 데이터시트 As String = "회계원장"
Const 기준셀 As String = "일자필드레이블"
Const 헤더줄수 As Integer = 5
Const 전체행수 As Integer = 4000
    
Const 열offset_날짜 As Integer = 0
Const 열offset_관항목 As Integer = 1
Const 열offset_code As Integer = 2
Const 열offset_관 As Integer = 3
Const 열offset_항 As Integer = 4
Const 열offset_목 As Integer = 5
Const 열offset_세목 As Integer = 6 '열 추가 : 2015.3.8
Const 열offset_적요 As Integer = 7
Const 열offset_수입 As Integer = 8
Const 열offset_지출 As Integer = 9
Const 열offset_은현 As Integer = 10
Const 열offset_VAT As Integer = 11
Const 열offset_대차 As Integer = 12
Const 열offset_프로젝트 As Integer = 13
Const 열offset_부서 As Integer = 14
Const 열offset_현금잔액 As Integer = 15
Const 열offset_통장잔액 As Integer = 16
Const 열offset_총잔액 As Integer = 17

Private Sub CheckBox_부가세포함_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "지출 내역일 경우 부가세가 포함된 가격을 적으셨으면 체크해주세요"
End Sub

Private Sub ComboBox_guan_change()
    Dim 관 As String
    관 = ComboBox_guan.Value
    
    If 관 = "수입" Or 관 = "지출" Then
        Call 항_초기화(관)
        ComboBox_hang.Enabled = True
        ComboBox_mok.Enabled = True
        ComboBox_세목.Enabled = True
        ComboBox_hang.SetFocus
    Else '예산외수입/지출
        ComboBox_hang.Enabled = False
        ComboBox_mok.Enabled = False
        ComboBox_세목.Enabled = False
        TextBox_summary.SetFocus
    End If
    
End Sub

Private Sub ComboBox_hang_Change()
    If ComboBox_hang.Enabled Then
        Call 목_초기화(ComboBox_guan.Value, ComboBox_hang.Value)
        ComboBox_mok.SetFocus
    End If
End Sub

Private Sub ComboBox_mok_Change()
    If ComboBox_mok.Enabled Then
        Call 세목_초기화(ComboBox_guan.Value, ComboBox_hang.Value, ComboBox_mok.Value)
        ComboBox_세목.SetFocus
    End If
End Sub

Private Sub ComboBox_세목_Change()
    If ComboBox_세목.Enabled Then
        MultiPage1.Value = 0
        TextBox_summary.SetFocus
    End If
End Sub

Private Sub ComboBox_부서_Change()
    'ComboBox_guan.SetFocus
End Sub

Private Sub ComboBox_프로젝트_Change()
    ComboBox_부서.SetFocus
End Sub

Private Sub CommandButton_close_Click()
    Unload Me
    If Parent <> "회계원장" Then
        홈
    End If
End Sub

Private Sub CommandButton1_Click()
    If IsEmpty(ComboBox_hang.Value) Or ComboBox_hang.Value = "" Then
        MsgBox "항을 선택해주세요"
        MultiPage1.Value = 0
        ComboBox_hang.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ComboBox_mok.Value) Or ComboBox_mok.Value = "" Then
        MsgBox "목을 선택해주세요"
        MultiPage1.Value = 0
        ComboBox_mok.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ComboBox_세목.Value) Or ComboBox_세목.Value = "" Then
        MsgBox "세목을 선택해주세요"
        MultiPage1.Value = 0
        ComboBox_세목.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(TextBox_summary.Value) Or TextBox_summary.Value = "" Then
        MsgBox "적요를 입력해주세요"
        MultiPage1.Value = 0
        TextBox_summary.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(TextBox_amount.Value) Or TextBox_amount.Value = "" Then
        MsgBox "금액을 입력해주세요"
        MultiPage1.Value = 0
        TextBox_amount.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_amount.Value) Or TextBox_amount.Value = 0 Then
        MsgBox "금액을 0보다 큰 숫자로 입력해주세요"
        MultiPage1.Value = 0
        TextBox_amount.Value = ""
        TextBox_amount.SetFocus
        Exit Sub
    End If
    
    Worksheets("회계원장").Unprotect PWD
    
    Call 저장
    If save_error = 0 Then
        MsgBox "입력되었습니다"
        Call 초기화
    End If
    
    If (Worksheets("설정").Range("시트잠금설정").Offset(, 1).Value = True) Then
        Worksheets("회계원장").Protect PWD
    End If
End Sub

Private Sub CommandButton2_Click()
    '상세입력부분에 있는 "회계원장입력" 버튼
    Call CommandButton1_Click
End Sub

Sub 저장()
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 코드 As String
    Dim 프로젝트 As String, 부서 As String
    Dim 적요 As String
    Dim 금액 As Long
    
    If ComboBox_guan.Value = "" Then
        MsgBox "관항목을 설정해주세요"
        save_error = 1
        Exit Sub
    End If
    
    If TextBox_summary.Value = "" Then
        MsgBox "적요를 입력해주세요"
        save_error = 2
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_amount.Value) Or Not TextBox_amount.Value > 0 Then
        MsgBox "금액을 입력해주세요"
        save_error = 3
        Exit Sub
    End If
    
    프로젝트 = ComboBox_프로젝트.Value
    부서 = ComboBox_부서.Value
    관 = ComboBox_guan.Value
    항 = ComboBox_hang.Value
    목 = ComboBox_mok.Value
    세목 = ComboBox_세목.Value
    코드 = get_code(관, 항, 목, 세목)
    적요 = TextBox_summary.Value
    금액 = TextBox_amount.Value
    
    Dim 일자 As String
    Dim 입출금유형 As Integer
    
    save_error = 0
    
    Dim r저장위치 As Range
    Set r저장위치 = Worksheets(데이터시트).Range("일자필드레이블").Offset(TextBox_행번호.Value - 5)
    Dim 첫행 As Integer
    Dim 끝행 As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim st As Range
    Dim 관항목 As String
    Dim newData() As Variant
    ReDim newData(1 To 14)
        
    관항목 = 코드 & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
    newData(열offset_관항목) = 관항목
    newData(열offset_code) = 코드
    newData(열offset_관) = 관
    newData(열offset_항) = 항
    newData(열offset_목) = 목
    newData(열offset_세목) = 세목

    If CheckBox_현금여부.Value Then
        입출금유형 = 1
    Else
        입출금유형 = 0
    End If
    newData(열offset_은현) = 입출금유형
    newData(열offset_프로젝트) = 프로젝트
    newData(열offset_부서) = 부서
    newData(열offset_적요) = 적요
    If 관 = "수입" Or 관 = "예산외수입" Then
        newData(열offset_수입) = 금액
    Else
        newData(열offset_지출) = 금액
    End If
    
    Worksheets(데이터시트).Unprotect
    
    일자 = TextBox_date.Value
    With r저장위치
        .Value = 일자
        .Offset(, 열offset_code).NumberFormat = "@"
        Worksheets(데이터시트).Range(.Offset(0, 1), .Offset(0, 13)).Value = newData
    End With
    
    If Worksheets("설정").Range("a2").Offset(, 1).Value = True Then
        Worksheets(데이터시트).Protect
    End If
        
End Sub

'자주 쓰는 입출력 폼에서 이 함수를 사용한다
Sub 회계원장입력(ByVal 관 As String, ByVal 항 As String, ByVal 목 As String, ByVal 세목 As String, ByVal 적요 As String, ByVal 금액 As Long)
    Dim ws As Worksheet
    Set ws = Worksheets("회계원장")
    Dim 기준점 As Range
    Set 기준점 = ws.Range("일자필드레이블").End(xlDown).Offset(1, 0)
    Dim 관항목, 코드 As String
    
    코드 = get_code(관, 항, 목, 세목)
    If 코드 = "" Then
        MsgBox "등록되지 않은 관항목입니다. 설정하신 관항목이 맞는지 확인해주세요"
        Exit Sub
    End If
    관항목 = 코드 & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
    
    ws.Unprotect PWD

    With 기준점
        .Value = Date
        .Offset(, 열offset_관항목).Value = 관항목
        .Offset(, 열offset_code).NumberFormat = "@"
        .Offset(, 열offset_code).Value = 코드
        .Offset(, 열offset_관).Value = 관
        .Offset(, 열offset_항).Value = 항
        .Offset(, 열offset_목).Value = 목
        .Offset(, 열offset_세목).Value = 세목
        .Offset(, 열offset_적요).Value = 적요
        If 관 = "수입" Or 관 = "예산외수입" Then
            .Offset(, 열offset_수입).Value = 금액
        Else
            .Offset(, 열offset_지출).Value = 금액
        End If
        .Offset(, 열offset_은현).Value = 0
    End With
    
    If (Worksheets("설정").Range("시트잠금설정").Offset(, 1).Value = True) Then
        ws.Protect PWD
    End If
    MsgBox "회계원장에 입력되었습니다"
End Sub

Sub 초기화()
    Dim 컨트롤 As Control
    For Each 컨트롤 In Me.Controls
        If TypeOf 컨트롤 Is MSForms.TextBox Then 컨트롤.Value = ""
        If TypeOf 컨트롤 Is MSForms.combobox Then 컨트롤.Value = ""
    Next
    
    With Worksheets(데이터시트).Range("일자필드레이블").End(xlDown)
        TextBox_date.Value = .Value
        TextBox_행번호.Value = .Offset(1, 0).Row
    End With
    
    CheckBox_현금여부.Value = False '거의 대부분 거래는 현금거래가 아니므로 은행으로 초기화
    TextBox_date.SetFocus
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "검색할 날짜를 입력하고 '검색'버튼을 눌러주세요"
End Sub

Private Sub TextBox_income_Change()
    TextBox_income.Value = format(TextBox_income.Value, "#,#")
End Sub

Private Sub CommandButton3_Click()
    Call CommandButton_close_Click
End Sub

Private Sub TextBox_amount_Change()
    TextBox_amount.Value = format(TextBox_amount.Value, "#,#")
End Sub

Private Sub TextBox_outgoings_Change()
    TextBox_outgoings.Value = format(TextBox_outgoings.Value, "#,#")
End Sub

Private Sub TextBox_search_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "검색할 날짜를 이곳에 입력합니다"
End Sub

Private Sub UserForm_Initialize()

    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    이전항목 = ""
    관수 = ws.Range("b2").CurrentRegion.Rows.Count
    목수 = ws.Range("d4").CurrentRegion.Rows.Count

    Call 프로젝트_초기화
    Call 부서_초기화
    
    For Each 관항목 In ws.Range("b2", "b" & 관수)
        
        If 관항목.Value <> "" Then
            If 관항목.Value <> 이전항목 Then
                ComboBox_guan.AddItem 관항목.Value
                이전항목 = 관항목.Value
            End If
        End If
        
    Next 관항목
    
    With Worksheets(데이터시트).Range("일자필드레이블")
        If .Offset(1, 0).Value <> "" Then
            TextBox_행번호.Value = .End(xlDown).Offset(1, 0).Row
        Else
            TextBox_행번호.Value = .Offset(1, 0).Row
        End If
    End With
    
    Call 항_초기화("지출")
    TextBox_date.Value = Date

    MultiPage1.Value = 0
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Sub 항_초기화(관 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")

    항수 = ws.Range("c4").CurrentRegion.Rows.Count
    이전항목 = ""
    ComboBox_hang.Clear
    
    For Each 관항목 In ws.Range("c4", "c" & 항수)
        If 관항목.Value <> "" Then
            If 관항목.Offset(, -1).Value = 관 And 관항목.Value <> 이전항목 Then
                ComboBox_hang.AddItem 관항목.Value
                이전항목 = 관항목.Value
            End If
        End If
    Next 관항목
    
End Sub

Sub 목_초기화(관 As String, 항 As String)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets("예산서")

        Dim 목수 As Integer
        Dim 이전항목 As String
        목수 = ws.Range("d4").CurrentRegion.Rows.Count
        이전항목 = ""
        ComboBox_mok.Clear
        
        For Each 관항목 In ws.Range("d4", "d" & 목수)
            With 관항목
                If .Value <> "" Then
                    If .Offset(, -2).Value = 관 And .Offset(, -1).Value = 항 And .Value <> 이전항목 Then
                        ComboBox_mok.AddItem .Value
                        이전항목 = .Value
                    End If
                End If
            End With
        Next 관항목
    
End Sub

Sub 세목_초기화(관 As String, 항 As String, 목 As String)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets("예산서")

        Dim 세목수 As Integer
        Dim 이전항목 As String
        세목수 = ws.Range("e4").CurrentRegion.Rows.Count
        이전항목 = ""
        ComboBox_세목.Clear
        
        For Each 관항목 In ws.Range("e4", "e" & 세목수)
            With 관항목
                If .Value <> "" Then
                    If .Offset(, -3).Value = 관 And .Offset(, -2).Value = 항 And .Offset(, -1).Value = 목 And .Value <> 이전항목 Then
                        ComboBox_세목.AddItem .Value
                        이전항목 = .Value
                    End If
                End If
            End With
        Next 관항목
    
End Sub

Sub 프로젝트_초기화()
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim 프로젝트 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("설정")
    Dim 프로젝트수 As Integer
    
    With ws.Range("프로젝트설정레이블")
        If .Offset(1, 0).Value <> "" Then
            프로젝트수 = .CurrentRegion.Rows.Count - 1
        Else
            프로젝트수 = 0
        End If
        
        If 프로젝트수 > 0 Then
            ComboBox_프로젝트.Enabled = True
            Set 시작행 = .Offset(1)
            If 시작행.Offset(1, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
            Else
                Set 종료행 = 시작행
            End If
            
            For Each 프로젝트 In ws.Range(시작행, 종료행)
                If 프로젝트.Value <> "" And 프로젝트.Value <> "프로젝트명" Then
                    ComboBox_프로젝트.AddItem 프로젝트.Value
                End If
            Next 프로젝트
        Else
            ComboBox_프로젝트.Enabled = False
        End If
    End With
End Sub

Sub 부서_초기화()
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim 부서 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("설정")
    Dim 부서수 As Integer
    
    With ws.Range("부서설정레이블")
        If .Offset(1, 0).Value <> "" Then
            부서수 = .CurrentRegion.Rows.Count - 1
        Else
            부서수 = 0
        End If
        
        If 부서수 > 0 Then
            ComboBox_부서.Enabled = True
            Set 시작행 = .Offset(1)
            If 시작행.Offset(1, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
            Else
                Set 종료행 = 시작행
            End If
            
            For Each 부서 In ws.Range(시작행, 종료행)
                If 부서.Value <> "" And 부서.Value <> "부서명" Then
                    ComboBox_부서.AddItem 부서.Value
                End If
            Next 부서
        Else
            ComboBox_부서.Enabled = False
        End If
    End With
End Sub
