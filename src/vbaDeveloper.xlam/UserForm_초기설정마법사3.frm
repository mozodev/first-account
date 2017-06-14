VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_초기설정마법사3 
   Caption         =   "초기설정마법사_3단계"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   OleObjectBlob   =   "UserForm_초기설정마법사3.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_초기설정마법사3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_계정과목1_Click()
    Dim 계정과목수 As Integer
    Call UserForm_계정과목보기.계정과목로드("공통", "계정과목샘플")
    UserForm_계정과목보기.Show
End Sub

Private Sub CommandButton_계정과목2_Click()
    Call UserForm_계정과목보기.계정과목로드("위탁", "계정과목샘플")
    UserForm_계정과목보기.Show
End Sub

Private Sub CommandButton_계정과목3_Click()
    Call UserForm_계정과목보기.계정과목로드("수익", "계정과목샘플")
    UserForm_계정과목보기.Show
End Sub

Private Sub CommandButton1_Click()
    Dim 유형 As String

    If OptionButton_계정과목1.Value = True Then
        유형 = "공통"
    ElseIf OptionButton_계정과목2.Value = True Then
        유형 = "위탁"
    ElseIf OptionButton_계정과목3.Value = True Then
        유형 = "수익"
    Else
        유형 = ""
    End If
    
    If 유형 <> "" Then
        Call 계정과목가져오기("계정과목샘플", 유형)
        MsgBox "선택하신 유형의 계정과목을 기본으로 적용했습니다"
    End If
    
    Unload Me
    If CheckBox_계정과목self.Value = True Then
        MsgBox "관항목을 추가로 입력할 수 있는 시트로 이동합니다"
        Dim ws As Worksheet
        Set ws = Worksheets("가져오기2")
        ws.Visible = xlSheetVisible
        ws.Activate
    Else
        MsgBox "마법사를 마칩니다. 처음엑셀회계를 마음껏 활용해주세요"
        홈
    End If
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    홈
End Sub

Private Sub CommandButton3_Click()
    Unload Me
    Dim 기관명 As String
    Dim 회계시작일 As String
    Dim 담당자직함 As String
    Dim 결재1직함 As String
    Dim 결재2직함 As String
    
    With Worksheets("설정")
        기관명 = .Range("기관명설정").Offset(0, 1).Value
        회계시작일 = .Range("회계시작일설정").Offset(0, 1).Value
        담당자직함 = .Range("담당자직함설정").Offset(0, 1).Value
        결재1직함 = .Range("결재1설정").Offset(0, 1).Value
        결재2직함 = .Range("결재2설정").Offset(0, 1).Value
    End With
    
    With UserForm_초기설정마법사2
        .TextBox_기관명 = 기관명
        .TextBox_회계시작일 = 회계시작일
        .TextBox_담당자직함 = 담당자직함
        .TextBox_결재1직함 = 결재1직함
        .TextBox_결재2직함 = 결재2직함
    End With
    UserForm_초기설정마법사2.Show
    
End Sub

