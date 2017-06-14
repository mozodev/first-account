VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_예산 
   Caption         =   "예산설정"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "UserForm_예산.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_예산"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const 데이터시트 As String = "예산서"
Const 헤더줄수 As Integer = 1
Dim error_num As Integer
Dim data_changed As Integer

Private Sub ComboBox_관_Change()
    Call 항_초기화(ComboBox_관.Value)
    ComboBox_항.SetFocus
End Sub

Private Sub ComboBox_세목_Change()
    Dim 관항목 As Range
    Dim 이전항목 As String, 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 목수 As Integer
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)

    Dim 세목수 As Integer
    세목수 = ws.Range("e4").CurrentRegion.Rows.Count
    이전항목 = ""
    관 = ComboBox_관.Value
    항 = ComboBox_항.Value
    목 = ComboBox_목.Value
    세목 = ComboBox_세목.Value
        
    For Each 관항목 In ws.Range("e4", "e" & 세목수)
        With 관항목
            If .Value <> "" Then
                If .Offset(, -3).Value = 관 And .Offset(, -2).Value = 항 And .Offset(, -1).Value = 목 And .Value = 세목 Then
                    TextBox_행번호.Value = .Row
                    
                    Exit For
                End If
            End If
        End With
    Next 관항목
    
    TextBox_예산액.SetFocus
End Sub

Private Sub ComboBox_항_Change()
    Call 목_초기화(ComboBox_관.Value, ComboBox_항.Value)
    ComboBox_목.SetFocus
End Sub

Private Sub ComboBox_목_Change()
    Call 세목_초기화(ComboBox_관.Value, ComboBox_항.Value, ComboBox_목.Value)
    ComboBox_세목.SetFocus
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim 행번호 As Integer
        
    If ComboBox_관.Value = "" Then
        MsgBox "'관'을 선택해주십시오"
        ComboBox_관.SetFocus
        Exit Sub
    End If
    
    If ComboBox_항.Value = "" Then
        MsgBox "'항'을 선택해주십시오"
        ComboBox_항.SetFocus
        Exit Sub
    End If
    
    If ComboBox_목.Value = "" Then
        MsgBox "'목'을 선택해주십시오"
        ComboBox_목.SetFocus
        Exit Sub
    End If
    
    If ComboBox_세목.Value = "" Then
        MsgBox "'세목'을 선택해주십시오"
        ComboBox_세목.SetFocus
        Exit Sub
    End If
    
    If TextBox_예산액.Value = "" Or Not IsNumeric(TextBox_예산액.Value) Then
        MsgBox "예산액을 숫자로 입력해주십시오"
        TextBox_예산액.SetFocus
        Exit Sub
    End If
    
    Set ws = Worksheets(데이터시트)
    행번호 = TextBox_행번호.Value
    
    If 행번호 > 0 Then
        With ws.Range("예산액필드")
            .Offset(행번호 - 헤더줄수).Value = TextBox_예산액.Value
        End With
        '연속 입력 위해 필드 초기화
        MsgBox "입력됐습니다"
        error_num = 0
        data_changed = 1
        Call 초기화
    Else
        error_num = 1
        MsgBox "저장에 실패했습니다"
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me

    If error_num = 0 And data_changed = 1 Then
        Call 결산서초기화
    End If
    홈
End Sub

Private Sub TextBox_예산액_Change()
    TextBox_예산액.Value = format(TextBox_예산액.Value, "#,#")
End Sub

Private Sub UserForm_Initialize()
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)
    Dim 이전항목 As String
    이전항목 = ""
    Dim 관수 As Integer
    Dim 목수 As Integer
    Dim 관 As String
    
    관수 = ws.Range("b2").CurrentRegion.Rows.Count
    목수 = ws.Range("d4").CurrentRegion.Rows.Count

    For Each 관항목 In ws.Range("b2", "b" & 관수)
        관 = 관항목.Value
        
        If 관항목.Value <> "" Then
            If 관 <> "예산외수입" And 관 <> "예산외지출" Then
                If 관 <> 이전항목 Then
                    ComboBox_관.AddItem 관
                    이전항목 = 관
                End If
            End If
        End If
        
    Next 관항목
    
    error_num = 0
    data_changed = 0

End Sub

Sub 항_초기화(관 As String, Optional ByRef frm As UserForm)
    Dim 관항목 As Range
    Dim ws As Worksheet
    Set ws = Worksheets(데이터시트)

    Dim 항수 As Integer
    항수 = ws.Range("c4").CurrentRegion.Rows.Count
    Dim 이전항목 As String
    이전항목 = ""
    Dim 폼 As UserForm
        
    If Not frm Is Nothing Then
        Set 폼 = frm
    Else
        Set 폼 = UserForm_예산
    End If
        
    폼.ComboBox_항.Clear
    
    For Each 관항목 In ws.Range("c4", "c" & 항수)
        If 관항목.Value <> "" Then
            If 관항목.Offset(, -1).Value = 관 And 관항목.Value <> 이전항목 Then
                폼.ComboBox_항.AddItem 관항목.Value
                이전항목 = 관항목.Value
            End If
        End If
    Next 관항목

End Sub

Sub 목_초기화(관 As String, 항 As String, Optional ByRef frm As UserForm)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets(데이터시트)

        Dim 목수 As Integer
        Dim 이전항목 As String
        목수 = ws.Range("d4").CurrentRegion.Rows.Count
        이전항목 = ""
        Dim 폼 As UserForm
        
        If Not frm Is Nothing Then
            Set 폼 = frm
        Else
            Set 폼 = UserForm_예산
        End If
        
        폼.ComboBox_목.Clear
        
        For Each 관항목 In ws.Range("d4", "d" & 목수)
            With 관항목
                If .Value <> "" Then
                    If .Offset(, -2).Value = 관 And .Offset(, -1).Value = 항 And .Value <> 이전항목 Then
                        폼.ComboBox_목.AddItem .Value
                        
                        이전항목 = .Value
                    End If
                End If
            End With
        Next 관항목
    
End Sub

Sub 세목_초기화(관 As String, 항 As String, 목 As String, Optional ByRef frm As UserForm)
        Dim 관항목 As Range
        Dim ws As Worksheet
        Set ws = Worksheets(데이터시트)

        Dim 세목수 As Integer
        Dim 이전항목 As String
        세목수 = ws.Range("d4").CurrentRegion.Rows.Count
        이전항목 = ""
        Dim 세목 As String
        Dim 폼 As UserForm
        
        If Not frm Is Nothing Then
            Set 폼 = frm
        Else
            Set 폼 = UserForm_예산
        End If
        
        폼.ComboBox_세목.Clear
        
        For Each 관항목 In ws.Range("e4", "e" & 세목수)
            With 관항목
                세목 = .Value
            
                If 세목 <> "" Then
                    If .Offset(, -3).Value = 관 And .Offset(, -2).Value = 항 And .Offset(, -1).Value = 목 And 세목 <> 이전항목 Then
                        폼.ComboBox_세목.AddItem 세목
                        
                        이전항목 = 세목
                    End If
                End If
            End With
        Next 관항목
    
End Sub

Sub 초기화()
    Dim 컨트롤 As Control
    For Each 컨트롤 In UserForm_입출금내역.Controls
        If TypeOf 컨트롤 Is MSForms.TextBox Then 컨트롤.Value = ""
        If TypeOf 컨트롤 Is MSForms.combobox Then 컨트롤.Value = ""
    Next
End Sub
