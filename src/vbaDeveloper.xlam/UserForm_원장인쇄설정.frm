VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_원장인쇄설정 
   Caption         =   "회계원장인쇄"
   ClientHeight    =   4065
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4365
   OleObjectBlob   =   "UserForm_원장인쇄설정.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_원장인쇄설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_close_Click()
    Unload Me
End Sub

Private Sub CommandButton_print_Click()
    '원장인쇄
    Dim 시작일 As String
    Dim 종료일 As String
    시작일 = TextBox_startdate.Value
    종료일 = TextBox_enddate.Value
    Unload Me
    Call 원장인쇄(시작일, 종료일)
End Sub

