VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_계정과목 
   Caption         =   "계정과목 설정"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11430
   OleObjectBlob   =   "UserForm_계정과목.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_계정과목"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const 헤더줄수 As Integer = 3
Const 데이터시트 As String = "예산서"
Const 관위치 As Integer = 1 'offset 용
Const 항위치 As Integer = 2 'offset 용
Const 목위치 As Integer = 3 'offset 용
Const 세목위치 As Integer = 4 'offset 용
Const 과목설명위치 As Integer = 6
Const 첫행 As Integer = 6

Private Sub ComboBox_guan_change()
    Call 항_초기화(ComboBox_guan.Value)
    CommandButton_항_신규.Enabled = True

    TextBox_항.Value = ""
    TextBox_항.Enabled = False
    TextBox_목.Value = ""
    TextBox_목.Enabled = False
    TextBox_세목.Value = ""
    TextBox_세목.Enabled = False
   
    ListBox_목.Clear
    ListBox_세목.Clear
End Sub

Private Sub ComboBox_hang_AfterUpdate()
    TextBox_guan.Enabled = False
    Call 목_초기화(ComboBox_guan.Value, ComboBox_hang.Value)
    textbox_hang.Value = ComboBox_hang.Value
End Sub

Private Sub ComboBox_guan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "수입/지출을 선택해주세요. 그에 따른 '항'들이 표시됩니다"
End Sub

Private Sub CommandButton_목_신규_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    TextBox_목_행번호.Value = ""
    Dim 마지막행 As Integer

    With ws.Range("a4")
        .Select
        마지막행 = .CurrentRegion.Rows.Count
    End With
    TextBox_목.Enabled = True
    TextBox_목.Value = ""
    TextBox_목_행번호.Value = 마지막행 + 1
    CommandButton_목삭제.Enabled = False
    CommandButton_목_저장.Enabled = True
    마지막행 = 0

    TextBox_목.SetFocus
End Sub

Private Sub CommandButton_세목_신규_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    TextBox_세목_행번호.Value = ""
    Dim 마지막행 As Integer

    With ws.Range("a4")
        .Select
        마지막행 = .CurrentRegion.Rows.Count

    End With
    TextBox_세목.Enabled = True
    TextBox_세목.Value = ""
    TextBox_세목_행번호.Value = 마지막행 + 1
    CommandButton_세목삭제.Enabled = False
    CommandButton_세목_저장.Enabled = True
    마지막행 = 0

    TextBox_세목.SetFocus
End Sub

Private Sub CommandButton_목삭제_Click()
    Dim 행번호 As Integer
    Dim r시작위치 As Range
    Dim 항 As String
    Dim 관 As String
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    Dim 끝행 As Integer
    끝행 = ws.Range("관필드").End(xlDown).Row
    
    행번호 = TextBox_목_행번호.Value
    
    If Not 행번호 > 0 Then
        MsgBox "삭제할 수 없습니다"
    Else
        Set r시작위치 = Worksheets(데이터시트).Range("관필드").Offset(행번호 - 1)
        With r시작위치
            관 = .Value
            항 = .Offset(0, 1).Value
            If 관 <> "" Then
                Range(.Offset(0, -1), .Offset(0, 6)).Delete Shift:=xlUp
                Call 줄긋기(행번호)
            End If
        End With
        
        TextBox_목.Value = ""
        TextBox_목_행번호.Value = 끝행 + 1
        
        Call 목_초기화(관, 항)
        TextBox_목.Enabled = False
    End If
End Sub

Private Sub CommandButton_세목삭제_Click()
    Dim 행번호 As Integer
    Dim r시작위치 As Range
    Dim 항 As String
    Dim 관 As String
    Dim 목 As String
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    Dim 끝행 As Integer
    끝행 = ws.Range("관필드").End(xlDown).Row
    
    행번호 = TextBox_세목_행번호.Value
    
    If Not 행번호 > 0 Then
        MsgBox "삭제할 수 없습니다"
    Else
        Set r시작위치 = ws.Range("관필드").Offset(행번호 - 1)
        With r시작위치
            관 = .Value
            항 = .Offset(0, 1).Value
            목 = .Offset(0, 2).Value
            If 관 <> "" Then
                Range(.Offset(0, -1), .Offset(0, 6)).Delete Shift:=xlUp
                Call 줄긋기(행번호)
            End If
        End With
        
        TextBox_세목.Value = ""
        TextBox_세목_행번호.Value = 끝행 + 1
        Call 세목_초기화(관, 항, 목)
        TextBox_세목.Enabled = False
        'CommandButton_세목_신규.Enabled = False '2016.11.9
    End If
End Sub

Private Sub CommandButton_항_신규_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "편집을 취소하고 새로운 '항'을 추가하려면 클릭하세요. 입력란을 비웁니다"
End Sub

Private Sub CommandButton_항_저장_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "위의 입력란에서 '항' 이름을 변경했거나 새로 작성했으면 이 버튼을 눌러주세요"
End Sub

Private Sub CommandButton_항_신규_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    Dim 마지막행 As Integer

    With ws.Range("a4")
        .Select
        마지막행 = .CurrentRegion.Rows.Count
    End With
    TextBox_항.Enabled = True
    CommandButton_항_저장.Enabled = True
    TextBox_항.Value = ""
    TextBox_항_행번호.Value = 마지막행 + 1
    TextBox_항.SetFocus
End Sub

Private Sub CommandButton_목_저장_Click()
    Dim ws As Worksheet
    Dim 관 As String, 항 As String, 목 As String, 이전목 As String
    
    Dim 행번호 As Integer, 끝행 As Integer, 목수 As Integer
    Dim i선택 As Integer
    
    Set ws = Worksheets(데이터시트)
    관 = ComboBox_guan.Value
    With ListBox_항
        i선택 = .ListIndex
        If i선택 > -1 Then
            항 = .List(i선택, 1)
        End If
    End With
    
    With ws.Range("목필드")
        If .Offset(3, 0).Value = "" Then
            끝행 = 첫행
            목수 = 0
        Else
            끝행 = .Offset(3, 0).End(xlDown).Row
            목수 = 끝행 - 헤더줄수
        End If
    End With
    
    행번호 = TextBox_목_행번호.Value
    이전목 = ws.Range("A" & 행번호).Offset(, 목위치).Value
    목 = TextBox_목.Value
    
    If 목 = "" Then
        MsgBox "'목'을 입력해주세요"
        Exit Sub
    End If
    
    '목 저장
     If 목수 > 0 Then
        목수 = 0
        이전항목 = ""
        For Each 관항목 In ws.Range("A" & 행번호, "A" & 끝행)
            If 관항목.Value <> "" Then
                If 관항목.Offset(, 관위치).Value = 관 And 관항목.Offset(, 항위치).Value = 항 And 관항목.Offset(, 목위치).Value = 이전목 Then
                    목수 = 목수 + 1
                    관항목.Offset(, 목위치).Value = 목
                End If
            End If
        Next 관항목
        
    End If
    
    If 목수 = 0 Then '추가
        With ws.Range("A" & 행번호)
            .Offset(0, 관위치).Value = ComboBox_guan.Value
            .Offset(0, 항위치).Value = 항
            .Offset(0, 목위치).Value = 목 '우선 항과 같은 이름의 목 생성
            .Offset(0, 세목위치).Value = 목 '같은 이름의 세목도 생성. (세목을 옵션으로 하면 구현이 복잡하므로 목 처럼 모두 쓰는 것으로)
            Call 줄긋기(행번호)
        End With
    End If

    MsgBox "입력됐습니다"
    
    '관,항,목 정렬
    Call 계정과목_시트정렬
    ws.Range("관항목코드레이블").Select
    code_changed = True
    
    '목 초기화
    Call 목_초기화(관, 항)
    TextBox_목_행번호.Value = ""
    TextBox_목.Value = ""
    TextBox_목.Enabled = False
End Sub

Private Sub CommandButton_세목_저장_Click()
    Dim ws As Worksheet
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 행번호 As Integer, 끝행 As Integer, 목수 As Integer, i선택 As Integer
    
    Set ws = Worksheets(데이터시트)
    관 = ComboBox_guan.Value
    With ListBox_항
        i선택 = .ListIndex
        If i선택 > -1 Then
            항 = .List(i선택, 1)
        End If
    End With
    
    With ListBox_목
        i선택 = .ListIndex
        If i선택 > -1 Then
            목 = .List(i선택, 1)
        End If
    End With
    
    With ws.Range("세목필드")
        If .Offset(3, 0).Value = "" Then
            끝행 = 첫행
            목수 = 0
        Else
            끝행 = .End(xlDown).Row
            목수 = 끝행 - 헤더줄수
        End If
    End With
    
    If TextBox_세목_행번호.Value = "" Then
        MsgBox "추가/변경할 항목의 위치를 다시 지정해주세요"
        Exit Sub
    End If
    행번호 = TextBox_세목_행번호.Value
    
    세목 = TextBox_세목.Value
    If 세목 = "" Then
        MsgBox "'세목'을 입력해주세요"
        Exit Sub
    End If
    
    '목 저장
    With ws.Range("A" & 행번호)
        If 목수 >= 1 And 행번호 <= 끝행 Then ' 2개 이상 입력되어 있고, 기존의 값을 수정하는 경우
            .Offset(0, 세목위치).Value = 세목
        Else
            .Offset(0, 1).Value = ComboBox_guan.Value
            .Offset(0, 항위치).Value = 항
            .Offset(0, 목위치).Value = 목
            .Offset(0, 세목위치).Value = 세목
            Call 줄긋기(행번호)
        End If
    End With
    MsgBox "입력됐습니다"
    
    '관,항,목 정렬
    Call 계정과목_시트정렬
    ws.Range("관항목코드레이블").Select
    code_changed = True
    
    '목 초기화
    Call 세목_초기화(관, 항, 목)
    TextBox_세목_행번호.Value = ""
    TextBox_세목.Value = ""
    TextBox_세목.Enabled = False
End Sub

Private Sub CommandButton_항_저장_Click()
    Dim ws As Worksheet
    Dim 관 As String
    Dim 행번호 As Integer
    Dim 끝행 As Integer
    Dim 항수 As Integer
    Dim 항 As String
    Dim 이전항 As String
    Dim 관항목 As Range
        
    관 = ComboBox_guan.Value
    Set ws = Worksheets(데이터시트)
    
    With ws.Range("항필드")
        If .Offset(3, 0).Value = "" Then
            끝행 = 첫행
            항수 = 0
        Else
            끝행 = .Offset(3, 0).End(xlDown).Row
            항수 = 끝행 - 헤더줄수
        End If
    End With
    
    행번호 = TextBox_항_행번호.Value
    이전항 = ws.Range("A" & 행번호).Offset(, 항위치).Value
    항 = TextBox_항.Value
    
    If 항수 > 0 Then
        항수 = 0
        이전항목 = ""
        For Each 관항목 In ws.Range("A" & 행번호, "A" & 끝행)
            If 관항목.Value <> "" Then
                If 관항목.Offset(, 관위치).Value = 관 And 관항목.Offset(, 항위치).Value = 이전항 Then
                    항수 = 항수 + 1
                    관항목.Offset(, 항위치).Value = 항
                End If
            End If
        Next 관항목
        
    End If
    
    If 항수 = 0 Then '추가
        With ws.Range("A" & 행번호)
            .Offset(0, 관위치).Value = ComboBox_guan.Value
            .Offset(0, 항위치).Value = 항
            .Offset(0, 목위치).Value = 항 '우선 항과 같은 이름의 목 생성
            .Offset(0, 세목위치).Value = 항 '같은 이름의 세목도 생성. (세목을 옵션으로 하면 구현이 복잡하므로 목 처럼 모두 쓰는 것으로)
            Call 줄긋기(행번호)
        End With
    End If
    
    '관,항,목 정렬
    Call 계정과목_시트정렬
    ws.Range("관항목코드레이블").Select
    code_changed = True
    
    '항 초기화
    Call 항_초기화(관)

    TextBox_항.Value = ""
    TextBox_항.Enabled = False
End Sub

Sub 줄긋기(행번호 As Integer)
    Dim r기준위치 As Range
    Dim 범위 As Range
    Set r기준위치 = Range("A" & 행번호)
    Set 범위 = Range(r기준위치, r기준위치.Offset(0, 6))
    
    With 범위
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Private Sub CommandButton_항삭제_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "'항'아래 '목'이 있을 경우 삭제되지 않습니다. '목'을 모두 삭제하면 '항'도 삭제됩니다"
End Sub

Private Sub CommandButton1_Click()
    Call 저장
    Call 초기화
    Call 예산서코드입력
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm_계정과목

    If code_changed Then
        Call 예산서코드입력
        Call 결산서초기화
        code_changed = False
    End If
    홈
End Sub

Sub 초기화()
    Dim ws As Worksheet
    Dim 끝행 As Integer
    Dim 컨트롤 As Control
    
    For Each 컨트롤 In UserForm2.Controls
        If TypeOf 컨트롤 Is MSForms.TextBox Then 컨트롤.Value = ""
    Next
    
    Set ws = Worksheets(데이터시트)
    끝행 = ws.Range("관항목코드레이블").End(xlDown).Row + 1
    
    TextBox_항_행번호.Value = 끝행
    TextBox_목_행번호.Value = 끝행
    끝행 = 0
End Sub

Sub 저장()
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    
    Dim r저장위치 As Range
    Set r저장위치 = ws.Range("관항목코드레이블").End(xlDown).Offset(1)
    
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 예산액 As String
    관 = ComboBox_guan.Value
    항 = TextBox_항.Value
    목 = TextBox_목.Value
    세목 = TextBox_세목.Value
    예산액 = TextBox_budget.Value
    
    Dim code As String
    code = get_code(관, 항, 목, 세목)
    Dim c As Range
    If code <> "" Then
        With ws.Range("관항목코드레이블").CurrentRegion.columns(1)
            Set c = .Find(code)
            If Not c Is Nothing Then
                If c.Offset(0, 5).Value <> 예산액 Then
                    c.Offset(0, 5).Value = 예산액
                End If
            End If
        End With
    Else
        With r저장위치
            
            .Offset(0, 1).Value = 관
            .Offset(0, 2).Value = 항
            .Offset(0, 3).Value = 목
            .Offset(0, 4).Value = 세목
            .Offset(0, 5).Value = 예산액

            Call 계정과목_시트정렬

            code_changed = True
        End With
    End If
    
End Sub

Sub 항_초기화(관 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim 시작행 As Range, 종료행 As Range
    Dim vlist() As Variant
    Dim 항수 As Integer, x As Integer
    x = 0

    Dim 이전항목 As String
    이전항목 = ""
    
    ListBox_항.Clear
    
    With ws.Range("항필드")
        Set 시작행 = .Offset(3)
        If 시작행.Value = "" Then
            Set 종료행 = 시작행
            항수 = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
                항수 = 종료행.Row - 헤더줄수
            End If
        End If
            
        If 항수 > 0 Then
            For Each 관항목 In ws.Range(시작행, 종료행)
                If 관항목.Value <> "" Then
                    If 관항목.Offset(, -1).Value = 관 And 관항목.Value <> 이전항목 Then
                        ReDim Preserve vlist(2, x)
                        vlist(0, x) = 관항목.Row
                        vlist(1, x) = 관항목.Value
                        이전항목 = 관항목.Value
                        x = x + 1
                    End If
                End If
            Next 관항목
            ListBox_항.Column = vlist
        End If
    End With
    
    If x = 0 Then
        CommandButton_항삭제.Enabled = True
    End If
    
End Sub

Sub 목_초기화(관 As String, 항 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim vlist() As Variant
    Dim 목수 As Integer
    Dim 목 As String
    Dim x As Integer
    x = 0
    Dim 이전항목 As String
    이전항목 = ""
        
    ListBox_목.Clear
    
    With ws.Range("목필드")
        Set 시작행 = .Offset(3)
        If 시작행.Value = "" Then
            Set 종료행 = 시작행
            목수 = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
                목수 = 종료행.Row - 헤더줄수
            End If
        End If
            
        If 목수 > 0 Then
            For Each 관항목 In ws.Range(시작행, 종료행)
                With 관항목
                    목 = .Value
                    If 목 <> "" Then
                        If .Offset(, -2).Value = 관 And .Offset(, -1).Value = 항 And .Value <> 이전항목 Then
                            ReDim Preserve vlist(3, x)
                            vlist(0, x) = .Row
                            vlist(1, x) = 목
                            vlist(2, x) = .Offset(0, 3).Value
                            이전항목 = 목
                            x = x + 1
                        End If
                    End If
                End With
            Next 관항목
            
            If x > 0 Then
                ListBox_목.Column = vlist
            Else
                Call 항_초기화(ComboBox_guan.Value)
            End If
        End If
    End With
End Sub

Sub 세목_초기화(관 As String, 항 As String, 목 As String)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim vlist() As Variant
    Dim 세목수 As Integer
    Dim 세목 As String
    Dim x As Integer
    x = 0
    Dim 이전항목 As String
    이전항목 = ""
        
    ListBox_세목.Clear

    With ws.Range("세목필드")
        Set 시작행 = .Offset(3)
        If 시작행.Value = "" Then
            Set 종료행 = 시작행
            세목수 = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
                세목수 = 종료행.Row - 헤더줄수
            Else
                Set 종료행 = 시작행
                세목수 = 1
            End If
        End If
            
        If 세목수 > 0 Then

            For Each 관항목 In ws.Range(시작행, 종료행)
                With 관항목
                    세목 = .Value
                    If 세목 <> "" Then
                        If .Offset(, -3).Value = 관 And .Offset(, -2).Value = 항 And .Offset(, -1).Value = 목 And .Value <> 이전항목 Then
                            ReDim Preserve vlist(3, x)
                            vlist(0, x) = .Row
                            vlist(1, x) = 세목
                            vlist(2, x) = .Offset(0, 2).Value '세목/예산액/설명 순서
                            이전항목 = 세목
                            x = x + 1
                        End If
                    End If
                End With
            Next 관항목
            
            If x > 0 Then
                ListBox_세목.Column = vlist
            Else
                Call 목_초기화(ComboBox_guan.Value, TextBox_항.Value)
                CommandButton_세목_신규.Enabled = False ' 2016.11.9
            End If
            
        End If
    End With
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "'항'을 선택해주세요. 그에 따른 '목'이 오른쪽에 표시됩니다"
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "수입/지출을 선택해주세요. 그에 따른 '항'들이 표시됩니다"
End Sub

Private Sub ListBox_목_Click()
    Dim i선택 As Integer
    Dim 목 As String
    
    With ListBox_목
        i선택 = .ListIndex
        If i선택 > -1 Then
            목 = .List(i선택, 1)

            TextBox_목_행번호.Value = .List(i선택, 0)
            TextBox_목.Value = 목

            CommandButton_목삭제.Enabled = True
            Call 세목_초기화(ComboBox_guan.Value, TextBox_항.Value, 목)
        End If
    End With
    
    TextBox_목.Enabled = True

    CommandButton_목_저장.Enabled = True
    CommandButton_항_저장.Enabled = False
    CommandButton_세목_저장.Enabled = False
    CommandButton_세목_신규.Enabled = True
    CommandButton_세목삭제.Enabled = False
End Sub

Private Sub ListBox_목_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_목_저장.Enabled = True
    CommandButton_항_저장.Enabled = False
    CommandButton_세목_저장.Enabled = False
    CommandButton_세목삭제.Enabled = False
End Sub

Private Sub ListBox_세목_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_세목_저장.Enabled = True
    CommandButton_항_저장.Enabled = False
    CommandButton_목_저장.Enabled = False
    CommandButton_세목삭제.Enabled = True
End Sub

Private Sub ListBox_항_Click()
    Dim i선택 As Integer
    Dim 항 As String
    
    With ListBox_항
        i선택 = .ListIndex
        If i선택 > -1 Then
            항 = .List(i선택, 1)
            Call 목_초기화(ComboBox_guan.Value, 항)
            ListBox_세목.Clear
            TextBox_항_행번호.Value = .List(i선택, 0)
            TextBox_항.Value = 항
        End If
    End With
    TextBox_항.Enabled = True

    CommandButton_항_저장.Enabled = True
    CommandButton_목_저장.Enabled = False
    CommandButton_세목_저장.Enabled = False
    
    CommandButton_목_신규.Enabled = True
    CommandButton_세목삭제.Enabled = False
End Sub

Private Sub ListBox_세목_Click()
    Dim i선택 As Integer
    Dim 세목 As String
    Dim 과목설명 As String
    
    With ListBox_세목
        i선택 = .ListIndex
        If i선택 > -1 Then
            세목 = .List(i선택, 1)
            과목설명 = .List(i선택, 2)
            TextBox_세목_행번호.Value = .List(i선택, 0)
            TextBox_세목.Value = 세목

            CommandButton_세목삭제.Enabled = True
        End If
    End With
    
    TextBox_세목.Enabled = True

    CommandButton_목_저장.Enabled = False
    CommandButton_항_저장.Enabled = False
    CommandButton_세목삭제.Enabled = True
End Sub

Private Sub ListBox_항_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_항_저장.Enabled = True
    CommandButton_목_저장.Enabled = False
    CommandButton_세목_저장.Enabled = False
    CommandButton_세목삭제.Enabled = False
End Sub

Private Sub TextBox_목_Change()
    If TextBox_목.Value <> "" Then
        CommandButton_목_저장.Enabled = True
    End If
End Sub

Private Sub TextBox_세목_Change()
    If TextBox_목.Value <> "" Then
        CommandButton_세목_저장.Enabled = True
    End If
End Sub

Private Sub TextBox_항_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "'항'의 이름을 바꾸거나 추가하려면 이곳에 입력후 '저장'버튼을 눌러주세요"
End Sub

Private Sub UserForm_Initialize()
    code_changed = False
    
    Call 관_초기화
    
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    Dim 끝행 As Integer
    끝행 = ws.Range("관필드").End(xlDown).Row
    TextBox_항_행번호.Value = 끝행 + 1
    TextBox_목_행번호.Value = 끝행 + 1
    
    With ListBox_항
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_목
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_세목
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    CommandButton_항_신규.Enabled = False
    CommandButton_항_저장.Enabled = False
    CommandButton_항삭제.Enabled = False
    CommandButton_목_신규.Enabled = False
    CommandButton_목_저장.Enabled = False
    CommandButton_목삭제.Enabled = False
    CommandButton_세목_신규.Enabled = False
    CommandButton_세목_저장.Enabled = False
    CommandButton_세목삭제.Enabled = False

End Sub

Sub 관_초기화()
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    이전항목 = ""
    관수 = ws.Range("b2").CurrentRegion.Rows.Count
    목수 = ws.Range("d4").CurrentRegion.Rows.Count
    Dim 관 As String
    
    ComboBox_guan.Clear
    
    For Each 관항목 In ws.Range("b2", "b" & 관수)
        관 = 관항목.Value
        If 관 <> "" Then
            If 관 <> 이전항목 Then
                If 관 <> "예산외수입" And 관 <> "예산외지출" Then
                    ComboBox_guan.AddItem 관
                    이전항목 = 관
                End If
            End If
        End If
    
    Next 관항목
    
End Sub
