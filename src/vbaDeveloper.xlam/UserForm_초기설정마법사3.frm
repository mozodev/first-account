VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�ʱ⼳��������3 
   Caption         =   "�ʱ⼳��������_3�ܰ�"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   OleObjectBlob   =   "UserForm_�ʱ⼳��������3.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�ʱ⼳��������3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_��������1_Click()
    Dim ��������� As Integer
    Call UserForm_�������񺸱�.��������ε�("����", "�����������")
    UserForm_�������񺸱�.Show
End Sub

Private Sub CommandButton_��������2_Click()
    Call UserForm_�������񺸱�.��������ε�("��Ź", "�����������")
    UserForm_�������񺸱�.Show
End Sub

Private Sub CommandButton_��������3_Click()
    Call UserForm_�������񺸱�.��������ε�("����", "�����������")
    UserForm_�������񺸱�.Show
End Sub

Private Sub CommandButton1_Click()
    Dim ���� As String

    If OptionButton_��������1.Value = True Then
        ���� = "����"
    ElseIf OptionButton_��������2.Value = True Then
        ���� = "��Ź"
    ElseIf OptionButton_��������3.Value = True Then
        ���� = "����"
    Else
        ���� = ""
    End If
    
    If ���� <> "" Then
        Call ��������������("�����������", ����)
        MsgBox "�����Ͻ� ������ ���������� �⺻���� �����߽��ϴ�"
    End If
    
    Unload Me
    If CheckBox_��������self.Value = True Then
        MsgBox "���׸��� �߰��� �Է��� �� �ִ� ��Ʈ�� �̵��մϴ�"
        Dim ws As Worksheet
        Set ws = Worksheets("��������2")
        ws.Visible = xlSheetVisible
        ws.Activate
    Else
        MsgBox "�����縦 ��Ĩ�ϴ�. ó������ȸ�踦 ������ Ȱ�����ּ���"
        Ȩ
    End If
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    Ȩ
End Sub

Private Sub CommandButton3_Click()
    Unload Me
    Dim ����� As String
    Dim ȸ������� As String
    Dim ��������� As String
    Dim ����1���� As String
    Dim ����2���� As String
    
    With Worksheets("����")
        ����� = .Range("�������").Offset(0, 1).Value
        ȸ������� = .Range("ȸ������ϼ���").Offset(0, 1).Value
        ��������� = .Range("��������Լ���").Offset(0, 1).Value
        ����1���� = .Range("����1����").Offset(0, 1).Value
        ����2���� = .Range("����2����").Offset(0, 1).Value
    End With
    
    With UserForm_�ʱ⼳��������2
        .TextBox_����� = �����
        .TextBox_ȸ������� = ȸ�������
        .TextBox_��������� = ���������
        .TextBox_����1���� = ����1����
        .TextBox_����2���� = ����2����
    End With
    UserForm_�ʱ⼳��������2.Show
    
End Sub

