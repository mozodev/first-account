VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_시작화면 
   Caption         =   "처음엑셀회계 v1.5"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000.001
   OleObjectBlob   =   "UserForm_시작화면.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_시작화면"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    UserForm_초기설정마법사1.Show
End Sub

