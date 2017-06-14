VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_자주쓰는입출금 
   Caption         =   "간편입력 (자주쓰는입출금내역)"
   ClientHeight    =   6700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030.001
   OleObjectBlob   =   "UserForm_자주쓰는입출금.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_자주쓰는입출금"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox_관_Change()
    Call 항_초기화(ComboBox_관.Value)
    ComboBox_항.SetFocus
End Sub

Private Sub ComboBox_항_Change()
    Call 목_초기화(ComboBox_관.Value, ComboBox_항.Value)
    ComboBox_목.SetFocus
End Sub

Private Sub ComboBox_목_Change()
    Call 세목_초기화(ComboBox_관.Value, ComboBox_항.Value, ComboBox_목.Value)
    ComboBox_세목.SetFocus
End Sub

Private Sub ComboBox_세목_Change()
    TextBox_적요.SetFocus
End Sub

Private Sub CommandButton_close_Click()
    Unload Me
    If Parent <> "회계원장" Then
        홈
    End If
End Sub

Private Sub CommandButton_reset_Click()
    ComboBox_관.Value = ""
    ComboBox_항.Value = ""
    ComboBox_목.Value = ""
    ComboBox_세목.Value = ""
    TextBox_적요.Value = ""
    TextBox_금액.Value = ""
    TextBox_행번호.Value = ""
End Sub

Private Sub CommandButton_삭제_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    Dim 기준점 As Range
    Set 기준점 = Worksheets("설정").Range("템플릿설정레이블")
    
    With ListBox_입출금템플릿
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 6)

            If MsgBox("삭제하겠습니까?", vbYesNo) Then
                Set 기준점 = 기준점.Offset(행번호, 0)
                Range(기준점, 기준점.End(xlToRight)).Delete Shift:=xlUp
            End If

        End If
    End With
    
    Call load_입출력템플릿
End Sub

Private Sub CommandButton_수정_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox_입출금템플릿
        i선택 = .ListIndex
        If i선택 > -1 Then
            TextBox_행번호 = .List(i선택, 6)
            ComboBox_관.Value = .List(i선택, 0)
            ComboBox_항.Value = .List(i선택, 1)
            ComboBox_목.Value = .List(i선택, 2)
            ComboBox_세목.Value = .List(i선택, 3)
            TextBox_적요.Value = .List(i선택, 4)
            TextBox_금액.Value = .List(i선택, 5)
        End If
    End With
End Sub

Private Sub CommandButton_입력_Click()
    Dim i선택 As Integer
    Dim 행번호 As Integer

    Dim 관 As String, 항 As String, 목 As String, 세목 As String, 적요 As String
    Dim 금액 As Long
    
    With ListBox_입출금템플릿
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 6)
            
            관 = .List(i선택, 0)
            항 = .List(i선택, 1)
            목 = .List(i선택, 2)
            세목 = .List(i선택, 3)
            적요 = .List(i선택, 4)
            금액 = .List(i선택, 5)

            Call UserForm_입출금입력.회계원장입력(관, 항, 목, 세목, 적요, 금액)

        End If
    End With
End Sub

Private Sub CommandButton_추가_Click()
    If ComboBox_관.Value = "" Then
        MsgBox "관을 선택해주세요"
        ComboBox_관.SetFocus
        Exit Sub
    End If
    
    If ComboBox_항.Value = "" Then
        MsgBox "항을 선택해주세요"
        ComboBox_항.SetFocus
        Exit Sub
    End If
    
    If ComboBox_목.Value = "" Then
        MsgBox "목을 선택해주세요"
        ComboBox_목.SetFocus
        Exit Sub
    End If
    
    If ComboBox_세목.Value = "" Then
        MsgBox "세목을 선택해주세요"
        ComboBox_세목.SetFocus
        Exit Sub
    End If
    
    If TextBox_적요.Value = "" Then
        MsgBox "적요를 입력해주세요"
        TextBox_적요.SetFocus
        Exit Sub
    End If
    
    If TextBox_금액.Value = "" Then
        MsgBox "금액을 입력해주세요"
        TextBox_금액.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_금액.Value) Then
        MsgBox "금액은 숫자로 입력해주세요"
        TextBox_금액.SetFocus
        Exit Sub
    End If
    
    Call 저장
    Call CommandButton_reset_Click
    ComboBox_관.SetFocus
    
    Call load_입출력템플릿
End Sub

Sub 저장()
    Dim 행번호 As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("설정")
    Dim 기준점 As Range
    Set 기준점 = ws.Range("템플릿설정레이블")
    
    Dim 관, 항, 목, 세목, 적요 As String
    Dim 금액 As Long
    
    If TextBox_행번호.Value <> "" Then
        행번호 = CInt(TextBox_행번호.Value)
    Else
        If 기준점.Offset(1, 0).Value = "" Then
            행번호 = 기준점.Offset(1, 0).Row - 기준점.Row
        Else
            행번호 = 기준점.End(xlDown).Offset(1, 0).Row - 기준점.Row
        End If
    End If
    
    If 행번호 > 0 Then
        Set 기준점 = 기준점.Offset(행번호, 0)
        관 = ComboBox_관.Value
        항 = ComboBox_항.Value
        목 = ComboBox_목.Value
        세목 = ComboBox_세목.Value
        적요 = TextBox_적요.Value
        금액 = TextBox_금액.Value
        With 기준점
            .Value = 관
            .Offset(, 1).Value = 항
            .Offset(, 2).Value = 목
            .Offset(, 3).Value = 세목
            .Offset(, 4).Value = 적요
            .Offset(, 5).Value = 금액
        End With
        
        MsgBox "저장했습니다"
    End If
End Sub

Private Sub ListBox_입출금템플릿_Click()
    CommandButton_입력.Enabled = True
    CommandButton_수정.Enabled = True
    CommandButton_삭제.Enabled = True
End Sub

Private Sub TextBox_금액_Change()
    TextBox_금액.Value = format(TextBox_금액.Value, "#,#")
End Sub

Private Sub UserForm_Initialize()
    With ListBox_입출금템플릿
        .columnCount = 6
        .ColumnWidths = "1cm;2.2cm;2.4cm;2.5cm;2.7cm;1.5cm"
    End With
    Call load_입출력템플릿
    
    ComboBox_관.AddItem "수입"
    ComboBox_관.AddItem "지출"
    ComboBox_관.ListIndex = 1
    
    '항 초기화
    Call 항_초기화("지출")
End Sub

Sub 항_초기화(관 As String)
    Call UserForm_예산.항_초기화(관, UserForm_자주쓰는입출금)
End Sub

Sub 목_초기화(관 As String, 항 As String)
    Call UserForm_예산.목_초기화(관, 항, UserForm_자주쓰는입출금)
End Sub

Sub 세목_초기화(관 As String, 항 As String, 목 As String)
    Call UserForm_예산.세목_초기화(관, 항, 목, UserForm_자주쓰는입출금)
End Sub

Sub load_입출력템플릿()
    Dim ws As Worksheet
    Set ws = Worksheets("설정")
    Dim 기준점 As Range
    Set 기준점 = ws.Range("템플릿설정레이블")
    Dim vlist() As Variant
    Dim x As Integer
    
    If 기준점.Offset(1, 0).Value <> "" Then
        '관 항 목 세목 적요 금액
        
        Do
            ReDim Preserve vlist(7, x)
            Set 기준점 = 기준점.Offset(1, 0)
            vlist(0, x) = 기준점.Value
            vlist(1, x) = 기준점.Offset(0, 1).Value
            vlist(2, x) = 기준점.Offset(0, 2).Value
            vlist(3, x) = 기준점.Offset(0, 3).Value
            vlist(4, x) = 기준점.Offset(0, 4).Value
            vlist(5, x) = 기준점.Offset(0, 5).Value
            vlist(6, x) = 기준점.Row - ws.Range("템플릿설정레이블").Row
            x = x + 1
        Loop While Not IsEmpty(기준점.Offset(1, 0).Value)
        
        ListBox_입출금템플릿.Column = vlist
    Else
        ListBox_입출금템플릿.Clear
    End If
End Sub
