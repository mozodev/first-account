VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�ʱ⼳��������1 
   Caption         =   "�ʱ⼳��������_1�ܰ�"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7215
   OleObjectBlob   =   "UserForm_�ʱ⼳��������1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�ʱ⼳��������1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton_cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_first_Click()
    Unload Me
    UserForm_�ʱ⼳��������2.Show
End Sub

Private Sub CommandButton_flush_Click()
    If MsgBox("����:������ �ڷᰡ ��� ������ϴ�. ����Ͻðڽ��ϱ�?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    '��� ������ �����, ó�����·� ����
    'UserForm_����
    UserForm_����.ȸ���ڷ��ʱ�ȭ ("��ü")
    UserForm_����.ȸ�輳���ʱ�ȭ ("��ü")
    UserForm_����.�⺻�����ʱ�ȭ
    MsgBox "��� �ʱ�ȭ�ƽ��ϴ�"
    Unload Me
    Ȩ
End Sub

Private Sub CommandButton_newyear_Click()
    If MsgBox("����: ȸ����忡 �Է��� �ڷᰡ ������ϴ�. ����Ͻðڽ��ϱ�?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    UserForm_����.ȸ���ڷ��ʱ�ȭ ("��ü")
    'ȸ������� �ʱ�ȭ�ϱ�
    Unload Me
    Ȩ
End Sub
