VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_front_config 
   Caption         =   "����"
   ClientHeight    =   8710.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570.001
   OleObjectBlob   =   "UserForm_front_config.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_front_config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_setup_budget_Click()
    Call �޴�_���꼳��
End Sub

Private Sub CommandButton_setup_currentsubject_Click()
    Call �޴�_���꼭����
End Sub

Private Sub CommandButton_setup_favorite_Click()
    Call �޴�_���־��������
End Sub

Private Sub CommandButton_setup_import_Click()
    Call �޴�_���嵥���Ͱ�������
    Unload Me
End Sub

Private Sub CommandButton_setup_subject_Click()
    Call �޴�_���������Է�
End Sub

Private Sub CommandButton_setup_tag_Click()
    Call �޴�_������Ʈ�׺μ�����
End Sub

Private Sub CommandButton_ȯ�漳��_Click()
    Call �޴�_ȯ�漳��
End Sub

Private Sub CommandButton_setup_close_Click()
    Unload Me
End Sub
