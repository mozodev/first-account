VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_front_config 
   Caption         =   "설정"
   ClientHeight    =   8710.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570.001
   OleObjectBlob   =   "UserForm_front_config.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_front_config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_setup_budget_Click()
    Call 메뉴_예산설정
End Sub

Private Sub CommandButton_setup_currentsubject_Click()
    Call 메뉴_예산서보기
End Sub

Private Sub CommandButton_setup_favorite_Click()
    Call 메뉴_자주쓰는입출금
End Sub

Private Sub CommandButton_setup_import_Click()
    Call 메뉴_원장데이터가져오기
    Unload Me
End Sub

Private Sub CommandButton_setup_subject_Click()
    Call 메뉴_계정과목입력
End Sub

Private Sub CommandButton_setup_tag_Click()
    Call 메뉴_프로젝트및부서설정
End Sub

Private Sub CommandButton_환경설정_Click()
    Call 메뉴_환경설정
End Sub

Private Sub CommandButton_setup_close_Click()
    Unload Me
End Sub
