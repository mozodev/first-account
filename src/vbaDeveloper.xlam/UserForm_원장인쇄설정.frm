VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�����μ⼳�� 
   Caption         =   "ȸ������μ�"
   ClientHeight    =   4065
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4365
   OleObjectBlob   =   "UserForm_�����μ⼳��.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�����μ⼳��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_close_Click()
    Unload Me
End Sub

Private Sub CommandButton_print_Click()
    '�����μ�
    Dim ������ As String
    Dim ������ As String
    ������ = TextBox_startdate.Value
    ������ = TextBox_enddate.Value
    Unload Me
    Call �����μ�(������, ������)
End Sub

