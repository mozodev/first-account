VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_입출금내역 
   Caption         =   "입출금 내역 입력"
   ClientHeight    =   8505.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670.001
   OleObjectBlob   =   "UserForm_입출금내역.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_입출금내역"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim save_error As Integer
Dim 대차 As String
Const 데이터시트 As String = "회계원장"
Const 기준셀 As String = "일자필드레이블"
Const 헤더줄수 As Integer = 5
Const 전체행수 As Integer = 20000
    
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
        TextBox_summary.SetFocus
    End If
End Sub

Private Sub ComboBox_부서_Change()
    ComboBox_guan.SetFocus
End Sub

Private Sub ComboBox_프로젝트_Change()
    ComboBox_부서.SetFocus
End Sub

Private Sub CommandButton_next_Click()
    Call move_next
End Sub

Sub move_next()
    Dim 행번호 As Integer
    On Error Resume Next
    If TextBox_행번호.Value Then
        행번호 = TextBox_행번호.Value + 1
        With Worksheets(데이터시트).Range("a" & 행번호)
            If .Offset(0, 0).Value <> "" Then
                Call load_입출금레코드(행번호)
            Else
                MsgBox "끝 입력값입니다(다음 입력값이 없습니다)"
            End If
        End With
    Else
        Exit Sub
    End If

End Sub

Private Sub CommandButton_prev_Click()
    Call move_prev
End Sub

Sub move_prev()
    Dim 행번호 As Integer
    If TextBox_행번호.Value Then
        행번호 = TextBox_행번호.Value - 1
    Else
        행번호 = 5
    End If
    
    With Worksheets(데이터시트).Range("a" & 행번호)
        If .Offset(0, 0).Value <> "일자" Then
            Call load_입출금레코드(행번호)
            If 행번호 < 8 Then  '전기이월과 통장입금은 삭제 되지 않도록 보호
                CommandButton_삭제.Enabled = False
            End If
        Else
            MsgBox "첫 입력값입니다(이전 입력값이 없습니다)"
        End If
    End With
End Sub

Private Sub CommandButton_검색_Click()
    Dim 전체 As Range
    Dim 찾은날짜 As Range
    Dim 레코드 As Range
    Dim cell As Range
    Dim x As Integer, y As Integer
    Dim 키워드 As String
    Dim 첫위치 As String
    Dim vlist() As Variant
    
    키워드 = TextBox_search.Value
    y = 0
    
    If (Len(키워드) > 0) Then
        Set 전체 = Worksheets(데이터시트).Range("일자필드레이블").CurrentRegion.columns(1)
        Set 찾은날짜 = 전체.Find(What:=키워드, LookAt:=xlPart)
        
        If Not 찾은날짜 Is Nothing Then
            첫위치 = 찾은날짜.Address
            
            Do
                ReDim Preserve vlist(8, x)
                Set 레코드 = 찾은날짜.Resize(1, 10)
                vlist(0, x) = 레코드.Row
                vlist(1, x) = 레코드.Cells(, 열offset_날짜 + 1)
                vlist(2, x) = 레코드.Cells(, 열offset_관 + 1)
                vlist(3, x) = 레코드.Cells(, 열offset_항 + 1)
                vlist(4, x) = 레코드.Cells(, 열offset_목 + 1)
                vlist(5, x) = 레코드.Cells(, 열offset_적요 + 1)
                vlist(6, x) = 레코드.Cells(, 열offset_수입 + 1)
                vlist(7, x) = 레코드.Cells(, 열offset_지출 + 1)
                
                x = x + 1
                y = 0
                
                Set 찾은날짜 = 전체.FindNext(찾은날짜)
            Loop While Not 찾은날짜 Is Nothing And 찾은날짜.Address <> 첫위치
            
            ListBox1.Column = vlist
        Else
            MsgBox "검색결과가 존재하지 않습니다"
            ListBox1.Clear
        End If
    End If
End Sub

Private Sub CommandButton_검색_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "검색한 날짜를 입력한 후 '검색' 버튼을 눌러주세요"
End Sub

Private Sub CommandButton_삭제_Click()
    Dim 행번호 As Integer
    Dim r시작위치 As Range
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)

    행번호 = 0
    If TextBox_행번호.Value Then
        행번호 = TextBox_행번호.Value
    End If
    
    If Not 행번호 > 0 Then
        MsgBox "삭제할 수 없습니다"
        Exit Sub
    Else
        If MsgBox("삭제하겠습니까? (" & TextBox_date.Value & "/" & TextBox_summary.Value & ")", vbYesNo, "삭제확인") = vbYes Then
            Set r시작위치 = ws.Range(기준셀).Offset(행번호 - 헤더줄수)
            With r시작위치
                ws.Unprotect PWD
                ws.Range(.Offset(0, 0), .Offset(0, 열offset_부서)).Delete Shift:=xlUp
            End With
            
            Set r시작위치 = ws.Range(기준셀).Offset(행번호 - 헤더줄수)
            With r시작위치
                Range(.Offset(-1, 열offset_현금잔액), .Offset(전체행수, 열offset_총잔액)).Select
                Selection.FillDown
            End With
            MsgBox "삭제되었습니다"
            r시작위치.Select
        End If
    End If
    
    Set r시작위치 = ws.Range(기준셀).Offset(행번호 - 헤더줄수) '2016.10.4 추가(런타임 오류)
    If r시작위치.Value <> "" And r시작위치.Value <> "일자" Then
        Call load_입출금레코드(r시작위치.Row)
    Else
        Call 초기화
    End If
    
End Sub

Private Sub CommandButton_신규_Click()
    Call 초기화
End Sub

Private Sub CommandButton_신규_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "아래에 불러온 자료가 있으면 비우고 새롭게 입력합니다"
End Sub

Private Sub CommandButton_편집_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox1
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 0)
            Call load_입출금레코드(행번호)
        End If
    End With
End Sub

Sub load_입출금레코드(행번호 As Integer)
    Dim 부가세포함여부 As String
    Dim 지출유형구분 As Integer
    
    If IsEmpty(Worksheets(데이터시트).Range("a" & 행번호)) Then
        행번호 = Worksheets(데이터시트).Range("일자필드레이블").End(xlDown).Row
    End If
    
    TextBox_행번호.Value = 행번호
    
    With Worksheets(데이터시트).Range("a" & 행번호)
        If .Offset(0, 열offset_날짜).Value <> "" And .Offset(0, 열offset_날짜).Value <> "일자" Then
            TextBox_date.Value = .Offset(0, 열offset_날짜).Value
            ComboBox_프로젝트.Value = .Offset(0, 열offset_프로젝트).Value
            ComboBox_부서.Value = .Offset(0, 열offset_부서).Value
            
            ComboBox_guan.Value = .Offset(0, 열offset_관).Value
            Call 항_초기화(ComboBox_guan.Value)
            ComboBox_hang.Value = .Offset(0, 열offset_항).Value
            Call 목_초기화(ComboBox_guan.Value, ComboBox_hang.Value)
            ComboBox_mok.Value = .Offset(0, 열offset_목).Value
            Call 세목_초기화(ComboBox_guan.Value, ComboBox_hang.Value, ComboBox_mok.Value)
            ComboBox_세목.Value = .Offset(0, 열offset_세목).Value
            
            TextBox_summary.Value = .Offset(0, 열offset_적요).Value
            If ComboBox_guan.Value = "수입" Or ComboBox_guan.Value = "예산외수입" Then
                TextBox_amount.Value = .Offset(0, 열offset_수입).Value
            Else
                TextBox_amount.Value = .Offset(0, 열offset_지출).Value
            End If

                    
            지출유형구분 = .Offset(0, 열offset_은현).Value
            If 지출유형구분 = "0" Then
                ComboBox_out_type.Value = "은행"
            ElseIf 지출유형구분 = "1" Then
                ComboBox_out_type.Value = "현금"
            Else
                ComboBox_out_type.Value = "카드"
            End If
                
        End If
    End With
    
    If 행번호 > 7 Then
        CommandButton_삭제.Enabled = True
    End If
End Sub

Private Sub CommandButton_편집_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "왼쪽 목록상자(검색결과)에서 클릭한 후, 이 버튼을 누르면 내용을 고칠 수 있습니다"
End Sub

Private Sub CommandButton1_Click()
    Worksheets("회계원장").Unprotect PWD
    
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
    Unload UserForm_입출금내역
    If Parent = "회계원장" Then
        
    Else
        홈
    End If
End Sub

Sub 저장()
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 프로젝트 As String, 부서 As String, 코드 As String, 적요 As String
    Dim 금액 As Long
    Dim 일자 As String
    Dim 입출금유형 As Integer
    
    save_error = 0
    
    If ComboBox_guan.Value = "" Then
        MsgBox "관항목을 설정해주세요"
        save_error = 1
        Exit Sub
    End If
    관 = ComboBox_guan.Value
    
    If TextBox_summary.Value = "" Then
        MsgBox "적요를 입력해주세요"
        save_error = 2
        Exit Sub
    End If
    적요 = TextBox_summary.Value
    
    If Not IsNumeric(TextBox_amount.Value) Or Not TextBox_amount.Value > 0 Then
        MsgBox "금액을 입력해주세요"
        save_error = 3
        Exit Sub
    End If
    
    금액 = CLng(TextBox_amount.Value)
    프로젝트 = ComboBox_프로젝트.Value
    부서 = ComboBox_부서.Value
    항 = ComboBox_hang.Value
    목 = ComboBox_mok.Value
    세목 = ComboBox_세목.Value
    
    코드 = get_code(관, 항, 목, 세목)
    
    Dim r저장위치 As Range
    Set r저장위치 = Worksheets(데이터시트).Range("일자필드레이블").Offset(TextBox_행번호.Value - 5)

    Dim 첫행 As Integer
    Dim 끝행 As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim st As Range
        
    관항목 = 코드 & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
    Select Case ComboBox_out_type.Value
        Case "은행"
            입출금유형 = 0
        Case "현금"
            입출금유형 = 1
        Case "카드"
            입출금유형 = 2
        Case Else
            입출금유형 = 0
    End Select
    
    Worksheets(데이터시트).Unprotect
    
    일자 = TextBox_date.Value
    With r저장위치
        .Value = 일자
        .Offset(, 열offset_관항목).Value = 관항목

        .Offset(, 열offset_code).NumberFormatLocal = "G/표준"
        .Offset(, 열offset_code).FormulaR1C1 = "=left(RC[-1], 8)"
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

        .Offset(, 열offset_은현).Value = 입출금유형

        .Offset(, 열offset_프로젝트).Value = 프로젝트
        .Offset(, 열offset_부서).Value = 부서
        
    End With
    
    If Worksheets("설정").Range("a2").Offset(, 1).Value = True Then
        Worksheets(데이터시트).Protect
    End If
        
End Sub

Sub 초기화()
    Dim 컨트롤 As Control
    For Each 컨트롤 In UserForm_입출금내역.Controls
        If TypeOf 컨트롤 Is MSForms.TextBox Then 컨트롤.Value = ""
        If TypeOf 컨트롤 Is MSForms.combobox Then 컨트롤.Value = ""
    Next
    
    With Worksheets(데이터시트).Range("일자필드레이블").End(xlDown)
        TextBox_date.Value = .Value
        TextBox_행번호.Value = .Offset(1, 0).Row
    End With
    
    ComboBox_out_type.Value = "은행"
    TextBox_date.SetFocus
End Sub

Function get_상황도움말(상황코드 As String)
    Dim ws As Worksheet
    Set ws = Worksheets("상황도움말")
    Dim r시작위치 As Range
    Set r시작위치 = ws.Range("상황코드레이블")
    Dim 끝행 As Integer
    끝행 = r시작위치.End(xlDown).Row
    For i = 1 To 끝행
        With r시작위치.Offset(i, 0)
            If .Value = 상황코드 Then
                get_상황도움말 = .Offset(0, 4).Value
                Exit For
            End If
        End With
    Next i
    
    If get_상황도움말 = "" Then
        get_상황도움말 = "준비된 도움말이 없습니다."
    End If
    
End Function

Private Sub Image1_Click()
    MsgBox get_상황도움말("일상_입출금_검색")
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "검색할 날짜를 입력하고 '검색'버튼을 눌러주세요"
End Sub

Private Sub TextBox_income_Change()
    TextBox_income.Value = format(TextBox_income.Value, "#,#")
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
    Dim 이전항목 As String, 관수 As Integer, 목수 As Integer
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
        
    With ListBox1
        .columnCount = 8
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;1.5cm;3cm;1cm;1cm"
    End With
    
    With Worksheets(데이터시트).Range("일자필드레이블")
        If .Offset(1, 0).Value <> "" Then
            TextBox_행번호.Value = .End(xlDown).Offset(1, 0).Row
        Else
            TextBox_행번호.Value = .Offset(1, 0).Row
        End If
    End With
    
    TextBox_date.Value = Date
    ComboBox_out_type.List = Array("은행", "현금", "카드")
    ComboBox_out_type.Value = "은행"
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Sub 항_초기화(관 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    
    Dim 항수 As Integer, 이전항목 As String

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

