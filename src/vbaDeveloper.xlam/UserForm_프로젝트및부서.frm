VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_프로젝트및부서 
   Caption         =   "프로젝트/통장 설정"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   OleObjectBlob   =   "UserForm_프로젝트및부서.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_프로젝트및부서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim save_error As Integer
Dim 프로젝트입력행번호 As Integer
Dim 부서입력행번호 As Integer
Const 데이터시트 As String = "설정"
Const 프로젝트기준셀 As String = "프로젝트설정레이블"
Const 부서기준셀 As String = "부서설정레이블"
Const 헤더줄수 As Integer = 5

Private Sub CommandButton_닫기1_Click()
    Unload Me
End Sub

Private Sub CommandButton_닫기2_Click()
    Unload Me
End Sub

Sub 프로젝트초기화()
    Dim ws As Worksheet
    Dim 프로젝트 As Range
    Set ws = Worksheets(데이터시트)
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim vlist() As Variant
    Dim 프로젝트수 As Integer
    Dim x As Integer
    x = 0
    
    With ws.Range(프로젝트기준셀)
        If .Offset(1, 0).Value <> "" Then
            프로젝트수 = .CurrentRegion.Rows.Count - 1
        Else
            프로젝트수 = 0
        End If
        TextBox_프로젝트행번호.Value = .Offset(프로젝트수 + 1).Row '새 프로젝트가 들어갈 행번호
        
        If 프로젝트수 > 0 Then
            Set 시작행 = .Offset(1)
            If 시작행.Offset(1, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
            Else
                Set 종료행 = 시작행
            End If
            
            For Each 프로젝트 In ws.Range(시작행, 종료행)
                If 프로젝트.Value <> "" And 프로젝트.Value <> "프로젝트명" Then
                    ReDim Preserve vlist(2, x)
                    vlist(0, x) = 프로젝트.Row
                    vlist(1, x) = 프로젝트.Value
                End If
                x = x + 1
            Next 프로젝트
            ListBox_프로젝트.Column = vlist
        End If
    End With
    
    TextBox_프로젝트명.Value = ""
End Sub

Sub 부서초기화()
    Dim ws As Worksheet
    Dim 부서 As Range
    Set ws = Worksheets(데이터시트)
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim vlist() As Variant
    Dim x As Integer
    x = 0
    
    With ws.Range(부서기준셀)
        If .Offset(1, 0).Value <> "" Then
            부서수 = .CurrentRegion.Rows.Count - 1
        Else
            부서수 = 0
        End If
        TextBox_부서행번호.Value = .Offset(부서수 + 1).Row '새 부서가 들어갈 행번호
        
        If 부서수 > 0 Then
            Set 시작행 = .Offset(1)
            If 시작행.Offset(1, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
            Else
                Set 종료행 = 시작행
            End If
            
            For Each 부서 In ws.Range(시작행, 종료행)
                If 부서.Value <> "" And 부서.Value <> "부서명" Then
                    ReDim Preserve vlist(2, x)
                    vlist(0, x) = 부서.Row
                    vlist(1, x) = 부서.Value
                End If
                x = x + 1
            Next 부서
            ListBox_부서.Column = vlist
        End If
    End With
    
    TextBox_부서명.Value = ""
End Sub

Private Sub CommandButton_부서삭제_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox_부서
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 0)
            Worksheets(데이터시트).Range(부서기준셀).Offset(행번호 - 헤더줄수).Delete Shift:=xlUp
            Call 부서초기화
        End If
    End With
End Sub

Private Sub CommandButton_부서저장_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    If TextBox_부서명.Value <> "" Then
        ws.Range(부서기준셀).Offset(TextBox_부서행번호.Value - 헤더줄수).Value = TextBox_부서명.Value
        Call 부서초기화
    Else
        MsgBox "부서명을 입력해주세요"
    End If
    TextBox_부서명.SetFocus
End Sub

Private Sub CommandButton_부서편집_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox_부서
        i선택 = .ListIndex
        If i선택 > -1 Then
            TextBox_부서행번호.Value = .List(i선택, 0)
            TextBox_부서명.Value = .List(i선택, 1)
        End If
    End With
    TextBox_부서명.SetFocus
End Sub

Private Sub CommandButton_신규부서_Click()
    TextBox_부서행번호.Value = 새입력행번호("부서")
    TextBox_부서명.Value = ""
    TextBox_부서명.SetFocus
End Sub

Private Sub CommandButton_신규프로젝트_Click()
    TextBox_프로젝트행번호.Value = 새입력행번호("프로젝트")
    TextBox_프로젝트명.Value = ""
    TextBox_프로젝트명.SetFocus
End Sub

Private Sub CommandButton_신규프로젝트_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "불러온 값을 비우고 새로운 프로젝트를 만듭니다"
End Sub

Private Sub CommandButton_프로젝트삭제_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox_프로젝트
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 0)
            Worksheets(데이터시트).Range(프로젝트기준셀).Offset(행번호 - 헤더줄수).Delete Shift:=xlUp
            Call 프로젝트초기화
        End If
    End With
End Sub

Private Sub CommandButton_프로젝트저장_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    If TextBox_프로젝트명.Value <> "" Then
        ws.Range(프로젝트기준셀).Offset(TextBox_프로젝트행번호.Value - 헤더줄수).Value = TextBox_프로젝트명.Value
        Call 프로젝트초기화
    Else
        MsgBox "프로젝트명을 입력해주세요"
    End If
    TextBox_프로젝트명.SetFocus
End Sub

Private Sub CommandButton_프로젝트저장_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "새로운 이름으로 프로젝트를 추가 혹은 변경합니다"
End Sub

Private Sub CommandButton_프로젝트편집_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox_프로젝트
        i선택 = .ListIndex
        If i선택 > -1 Then
            TextBox_프로젝트행번호.Value = .List(i선택, 0)
            TextBox_프로젝트명.Value = .List(i선택, 1)
        End If
    End With
    
    TextBox_프로젝트명.SetFocus
End Sub

Function 새입력행번호(구분 As String)
    Dim 기준셀 As String
    If 구분 = "프로젝트" Then
        기준셀 = 프로젝트기준셀
    Else
        기준셀 = 부서기준셀
    End If
    
    With Worksheets(데이터시트).Range(기준셀)
        If .Offset(1, 0).Value <> "" Then
            새입력행번호 = .End(xlDown).Offset(1, 0).Row
        Else
            새입력행번호 = .Offset(1, 0).Row
        End If
    End With
End Function

Private Sub CommandButton_프로젝트편집_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "선택한 프로젝트 이름을 불러옵니다"
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "프로젝트 이름을 쓰고 '저장'을 누르면 반영됩니다"
End Sub

Private Sub TextBox_프로젝트명_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "이곳에 프로젝트의 새로운 이름을 적어주세요"
End Sub

Private Sub UserForm_Initialize()
    With ListBox_프로젝트
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_부서
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With Worksheets(데이터시트).Range(프로젝트기준셀)
        If .Offset(1, 0).Value <> "" Then
            프로젝트입력행번호 = .End(xlDown).Offset(1, 0).Row
        Else
            프로젝트입력행번호 = .Offset(1, 0).Row
        End If
    End With
    
    Call 프로젝트초기화
    
    With Worksheets(데이터시트).Range(부서기준셀)
        If .Offset(1, 0).Value <> "" Then
            부서입력행번호 = .End(xlDown).Offset(1, 0).Row
        Else
            부서입력행번호 = .Offset(1, 0).Row
        End If
    End With
    
    Call 부서초기화
    
    MultiPage2.Value = 0  '첫페이지(프로젝트설정)가 항상 먼저 뜨도록
End Sub

