VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�ʱ⼳��������2 
   Caption         =   "�ʱ⼳��������_2�ܰ�"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865.001
   OleObjectBlob   =   "UserForm_�ʱ⼳��������2.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�ʱ⼳��������2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    '������ ����
    Dim �����, ȸ�������, ���������, ����1����, ����2���� As String
    Dim ��Ʈ��� As String
    
    ����� = TextBox_�����.Value
    ȸ������� = TextBox_ȸ�������.Value
    ��������� = TextBox_���������.Value
    ����1���� = TextBox_����1����.Value
    ����2���� = TextBox_����2����.Value

    ��Ʈ��� = True
    
    With Worksheets("����")
        If ��������� <> "" Then
            .Range("E2").Value = ���������
        End If
        If ����1���� <> "" Then
            .Range("F2").Value = ����1����
        End If
        If ����2���� <> "" Then
            .Range("G2").Value = ����2����
        End If
    End With
    
    With Worksheets("����")
        .Range("��Ʈ��ݼ���").Offset(0, 1).Value = ��Ʈ���
        If ����� <> "" And ����� <> .Range("�������").Offset(0, 1).Value Then
            .Range("�������").Offset(0, 1).Value = �����
        End If
        
        If ��������� <> "" Then
            .Range("��������Լ���").Offset(0, 1).Value = ���������
        End If
        
        If ȸ������� <> "" And ȸ������� <> .Range("ȸ������ϼ���").Offset(0, 1).Value Then
            .Range("ȸ������ϼ���").Offset(0, 1).Value = ȸ�������
        End If
        
        If ����1���� <> "" Then
            .Range("����1����").Offset(0, 1).Value = ����1����
        End If
        
        If ����2���� <> "" Then
            .Range("����2����").Offset(0, 1).Value = ����2����
        End If

    End With
    
    With Worksheets("ȸ�����")
        .Unprotect PWD
        If ����� <> "" And ����� <> .Range("�����").Value Then
            .Range("�����").Value = �����
        End If

        If Worksheets("����").Range("��Ʈ��ݼ���").Offset(, 1).Value = True Then
            .Protect PWD
        End If
    End With
    
    Unload Me
    UserForm_�ʱ⼳��������3.Show
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    UserForm_�ʱ⼳��������1.Show
End Sub

