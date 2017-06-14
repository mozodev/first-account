VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_초기설정마법사1 
   Caption         =   "초기설정마법사_1단계"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7215
   OleObjectBlob   =   "UserForm_초기설정마법사1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_초기설정마법사1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton_cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_first_Click()
    Unload Me
    UserForm_초기설정마법사2.Show
End Sub

Private Sub CommandButton_flush_Click()
    If MsgBox("주의:설정과 자료가 모두 사라집니다. 계속하시겠습니까?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    '모든 설정을 지우고, 처음상태로 돌림
    'UserForm_설정
    UserForm_설정.회계자료초기화 ("전체")
    UserForm_설정.회계설정초기화 ("전체")
    UserForm_설정.기본설정초기화
    MsgBox "모두 초기화됐습니다"
    Unload Me
    홈
End Sub

Private Sub CommandButton_newyear_Click()
    If MsgBox("주의: 회계원장에 입력한 자료가 사라집니다. 계속하시겠습니까?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    UserForm_설정.회계자료초기화 ("전체")
    '회계시작일 초기화하기
    Unload Me
    홈
End Sub
