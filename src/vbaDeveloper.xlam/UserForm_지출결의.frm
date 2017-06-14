VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_지출결의 
   Caption         =   "지출결의"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505.001
   OleObjectBlob   =   "UserForm_지출결의.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_지출결의"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public save_error As Integer

Private Sub ComboBox_guan_AfterUpdate()
    Call 항_초기화(ComboBox_guan.Value)
End Sub

Private Sub ComboBox_hang_AfterUpdate()
    Call 목_초기화(ComboBox_hang.Value)
End Sub

Private Sub ComboBox_hang_Change()
    ComboBox_mok.Clear
    ComboBox_semok.Clear
End Sub

Private Sub ComboBox_mok_AfterUpdate()
    Call 세목_초기화(ComboBox_hang.Value, ComboBox_mok.Value)
End Sub

Private Sub ComboBox_mok_Change()
    ComboBox_semok.Clear
End Sub

Private Sub ComboBox_semok_Change()
    TextBox_코드.Value = get_code("지출", ComboBox_hang.Value, ComboBox_mok.Value, ComboBox_semok.Value)
    If TextBox_코드.Value <> "" Then
        
    End If
End Sub

Private Sub CommandButton1_Click()
'입력 버튼 클릭시
    Call 저장
    If save_error = 0 Then
        MsgBox "입력되었습니다"
        Call 초기화
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    If Parent = "지출결의대장" Then
        Worksheets("지출결의대장").Activate
    Else
        홈
    End If
End Sub

Sub 저장()
    save_error = 0
    Dim 코드 As String, 관 As String, 항 As String, 목 As String, 세목 As String, 지출명 As String
    Dim 수량 As Integer
    Dim 단가 As Long, 금액 As Long
    
    코드 = TextBox_코드.Value
    관 = "지출"
    항 = ComboBox_hang.Value
    목 = ComboBox_mok.Value
    지출명 = TextBox_지출명.Value
    
    If IsEmpty(TextBox_금액.Value) Or Not IsNumeric(TextBox_금액.Value) Then
        MsgBox "수량과 단가를 입력해 금액을 채워주세요"
        save_error = 3
        Exit Sub
    End If
    금액 = CLng(TextBox_금액.Value)
    
    If IsEmpty(TextBox_수량.Value) Or Not IsNumeric(TextBox_수량.Value) Then
        MsgBox "수량을 숫자로 입력해주세요"
        save_error = 3
        Exit Sub
    End If
    수량 = CInt(TextBox_수량.Value)
    
    If IsEmpty(TextBox_단가.Value) Or Not IsNumeric(TextBox_단가.Value) Then
        MsgBox "단가를 숫자로 입력해주세요"
        save_error = 3
        Exit Sub
    End If
    단가 = CLng(TextBox_단가.Value)
    
    세목 = ComboBox_semok.Value
    
    If 코드 = "" Then
        MsgBox "관항목을 설정해주세요"
        save_error = 1
        Exit Sub
    End If
    
    If 관 = "" Then
        MsgBox "관항목을 설정해주세요"
        save_error = 1
        Exit Sub
    End If
    
    If 지출명 = "" Then
        MsgBox "지출명을 입력해주세요"
        save_error = 2
        Exit Sub
    End If
    
    코드 = 코드 & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
    Dim r저장위치 As Range
    With Worksheets("지출결의대장").Range("결의날짜레이블")
        If .Offset(1, 0).Value = "" Then
            Set r저장위치 = .Offset(1, 0)
        Else
            Set r저장위치 = .End(xlDown).Offset(1)  '빈 파일에서 처리하게
        End If
        
    End With
    
    Worksheets("지출결의대장").Unprotect
    With r저장위치
        .Value = TextBox_날짜.Value
        .Offset(, 1).Value = 코드
        .Offset(, 2).Value = 지출명
        .Offset(, 3).Value = TextBox_규격.Value
        .Offset(, 4).Value = 수량
        .Offset(, 5).Value = 단가
        .Offset(, 6) = "=RC[-2] * RC[-1]" '금액
        .Offset(, 7).Value = TextBox_비고.Value
        .Offset(, 8).Value = TextBox_하단비고.Value
    End With
    
    Call 지출결의대장정렬
End Sub

Sub 초기화()
    Dim 컨트롤 As Control
    For Each 컨트롤 In UserForm_지출결의.Controls
        If TypeOf 컨트롤 Is MSForms.TextBox Then 컨트롤.Value = ""
        If TypeOf 컨트롤 Is MSForms.combobox Then 컨트롤.Value = ""
    Next
End Sub

Sub 항_초기화(관 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")

    Dim 항수 As Integer
    Dim 이전항목 As String
    항수 = ws.Range("c4").CurrentRegion.Rows.Count
    이전항목 = ""
    
    For Each 관항목 In ws.Range("c4", "c" & 항수)
        If 관항목.Value <> "" Then
            If 관항목.Offset(, -1).Value = 관 And 관항목.Value <> 이전항목 Then
                ComboBox_hang.AddItem 관항목.Value
                이전항목 = 관항목.Value
            End If
        End If
    Next 관항목
    
End Sub

Sub 목_초기화(항 As String)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets("예산서")

        Dim 목수 As Integer
        Dim 이전항목 As String
        목수 = ws.Range("d4").CurrentRegion.Rows.Count
        이전항목 = ""
        
        For Each 관항목 In ws.Range("d4", "d" & 목수)
            If 관항목.Value <> "" Then
                If 관항목.Offset(, -2).Value = "지출" And 관항목.Offset(, -1).Value = 항 And 관항목.Value <> 이전항목 Then
                    ComboBox_mok.AddItem 관항목.Value
                    이전항목 = 관항목.Value
                End If
            End If
        Next 관항목
    
End Sub

Sub 세목_초기화(항 As String, 목 As String)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets("예산서")

        Dim 세목수 As Integer
        Dim 이전항목 As String
        세목수 = ws.Range("e4").CurrentRegion.Rows.Count
        이전항목 = ""

        For Each 관항목 In ws.Range("e4", "e" & 세목수)
            If 관항목.Value <> "" Then
                If 관항목.Offset(, -3).Value = "지출" And 관항목.Offset(, -2).Value = 항 And 관항목.Offset(, -1).Value = 목 And 관항목.Value <> 이전항목 Then
                    ComboBox_semok.AddItem 관항목.Value
                    이전항목 = 관항목.Value
                End If
            End If
        Next 관항목
    
End Sub

Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "왼쪽의 '관/항/목'을 선택하시면 자동으로 입력됩니다."
End Sub

Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "지출결의서 제일 아래쪽에 표시되는 내용입니다"
End Sub

Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "개수, 상자, 회 등 단위를 적어주세요"
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "수량과 단가를 입력하시면 자동으로 계산됩니다"
End Sub

Private Sub TextBox_금액_Change()
    TextBox_금액.Value = format(TextBox_금액.Value, "#,#")
End Sub

Private Sub TextBox_금액_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "수량과 단가를 입력하시면 자동으로 계산됩니다"
End Sub

Private Sub TextBox_단가_AfterUpdate()
    Dim 수량 As Integer
    Dim 단가 As Long
    If Not IsNull(TextBox_수량.Value) And IsNumeric(TextBox_수량.Value) Then
        수량 = TextBox_수량.Value
    Else
        수량 = 1
    End If
    
    If Not IsNull(TextBox_단가.Value) And IsNumeric(TextBox_단가.Value) Then
        단가 = TextBox_단가.Value
    Else
        단가 = 0
    End If
    
    If IsEmpty(단가) Or Not IsNumeric(단가) Then
        MsgBox "숫자를 입력하셔야 합니다"
        save_error = 3
        Exit Sub
    End If

    If (Not IsEmpty(수량) And 수량 > 0) Then
        If (단가 > 0) Then
            TextBox_금액.Value = 단가 * 수량
        End If
    End If
End Sub

Private Sub TextBox_단가_Change()
    TextBox_단가.Value = format(TextBox_단가.Value, "#,#")
End Sub

Private Sub TextBox_수량_AfterUpdate()
    Dim 수량 As Integer
    Dim 단가 As Integer
    
    If Not IsNull(TextBox_수량.Value) And IsNumeric(TextBox_수량.Value) Then
        수량 = TextBox_수량.Value
    Else
        수량 = 1
    End If
    
    If Not IsNull(TextBox_단가.Value) And IsNumeric(TextBox_단가.Value) Then
        단가 = TextBox_단가.Value
    Else
        단가 = 0
    End If
    
    If IsEmpty(수량) Or Not IsNumeric(수량) Then
        MsgBox "숫자를 입력하셔야 합니다"
        save_error = 3
        Exit Sub
    End If
    
    If (Not IsEmpty(단가) And 단가 > 0) Then
        If (수량 > 0) Then
            TextBox_금액.Value = 단가 * 수량
        End If
    End If
End Sub

Private Sub TextBox_코드_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "왼쪽의 '관/항/목'을 선택하시면 자동으로 입력됩니다"
End Sub

Private Sub TextBox_하단비고_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "지출결의서 가장 아래쪽에 표시되는 내용입니다"
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")

    TextBox_수량.Value = 1
    TextBox_단가.Value = 0
    Call 항_초기화("지출")
    Call 프로젝트_초기화
    Call 부서_초기화
    
    TextBox_날짜.Value = Date
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
