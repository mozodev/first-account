VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_���־�������� 
   Caption         =   "�����Է� (���־�������ݳ���)"
   ClientHeight    =   6700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030.001
   OleObjectBlob   =   "UserForm_���־��������.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_���־��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox_��_Change()
    Call ��_�ʱ�ȭ(ComboBox_��.Value)
    ComboBox_��.SetFocus
End Sub

Private Sub ComboBox_��_Change()
    Call ��_�ʱ�ȭ(ComboBox_��.Value, ComboBox_��.Value)
    ComboBox_��.SetFocus
End Sub

Private Sub ComboBox_��_Change()
    Call ����_�ʱ�ȭ(ComboBox_��.Value, ComboBox_��.Value, ComboBox_��.Value)
    ComboBox_����.SetFocus
End Sub

Private Sub ComboBox_����_Change()
    TextBox_����.SetFocus
End Sub

Private Sub CommandButton_close_Click()
    Unload Me
    If Parent <> "ȸ�����" Then
        Ȩ
    End If
End Sub

Private Sub CommandButton_reset_Click()
    ComboBox_��.Value = ""
    ComboBox_��.Value = ""
    ComboBox_��.Value = ""
    ComboBox_����.Value = ""
    TextBox_����.Value = ""
    TextBox_�ݾ�.Value = ""
    TextBox_���ȣ.Value = ""
End Sub

Private Sub CommandButton_����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    Dim ������ As Range
    Set ������ = Worksheets("����").Range("���ø��������̺�")
    
    With ListBox_��������ø�
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 6)

            If MsgBox("�����ϰڽ��ϱ�?", vbYesNo) Then
                Set ������ = ������.Offset(���ȣ, 0)
                Range(������, ������.End(xlToRight)).Delete Shift:=xlUp
            End If

        End If
    End With
    
    Call load_��������ø�
End Sub

Private Sub CommandButton_����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox_��������ø�
        i���� = .ListIndex
        If i���� > -1 Then
            TextBox_���ȣ = .List(i����, 6)
            ComboBox_��.Value = .List(i����, 0)
            ComboBox_��.Value = .List(i����, 1)
            ComboBox_��.Value = .List(i����, 2)
            ComboBox_����.Value = .List(i����, 3)
            TextBox_����.Value = .List(i����, 4)
            TextBox_�ݾ�.Value = .List(i����, 5)
        End If
    End With
End Sub

Private Sub CommandButton_�Է�_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer

    Dim �� As String, �� As String, �� As String, ���� As String, ���� As String
    Dim �ݾ� As Long
    
    With ListBox_��������ø�
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 6)
            
            �� = .List(i����, 0)
            �� = .List(i����, 1)
            �� = .List(i����, 2)
            ���� = .List(i����, 3)
            ���� = .List(i����, 4)
            �ݾ� = .List(i����, 5)

            Call UserForm_������Է�.ȸ������Է�(��, ��, ��, ����, ����, �ݾ�)

        End If
    End With
End Sub

Private Sub CommandButton_�߰�_Click()
    If ComboBox_��.Value = "" Then
        MsgBox "���� �������ּ���"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_��.Value = "" Then
        MsgBox "���� �������ּ���"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_��.Value = "" Then
        MsgBox "���� �������ּ���"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_����.Value = "" Then
        MsgBox "������ �������ּ���"
        ComboBox_����.SetFocus
        Exit Sub
    End If
    
    If TextBox_����.Value = "" Then
        MsgBox "���並 �Է����ּ���"
        TextBox_����.SetFocus
        Exit Sub
    End If
    
    If TextBox_�ݾ�.Value = "" Then
        MsgBox "�ݾ��� �Է����ּ���"
        TextBox_�ݾ�.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_�ݾ�.Value) Then
        MsgBox "�ݾ��� ���ڷ� �Է����ּ���"
        TextBox_�ݾ�.SetFocus
        Exit Sub
    End If
    
    Call ����
    Call CommandButton_reset_Click
    ComboBox_��.SetFocus
    
    Call load_��������ø�
End Sub

Sub ����()
    Dim ���ȣ As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("����")
    Dim ������ As Range
    Set ������ = ws.Range("���ø��������̺�")
    
    Dim ��, ��, ��, ����, ���� As String
    Dim �ݾ� As Long
    
    If TextBox_���ȣ.Value <> "" Then
        ���ȣ = CInt(TextBox_���ȣ.Value)
    Else
        If ������.Offset(1, 0).Value = "" Then
            ���ȣ = ������.Offset(1, 0).Row - ������.Row
        Else
            ���ȣ = ������.End(xlDown).Offset(1, 0).Row - ������.Row
        End If
    End If
    
    If ���ȣ > 0 Then
        Set ������ = ������.Offset(���ȣ, 0)
        �� = ComboBox_��.Value
        �� = ComboBox_��.Value
        �� = ComboBox_��.Value
        ���� = ComboBox_����.Value
        ���� = TextBox_����.Value
        �ݾ� = TextBox_�ݾ�.Value
        With ������
            .Value = ��
            .Offset(, 1).Value = ��
            .Offset(, 2).Value = ��
            .Offset(, 3).Value = ����
            .Offset(, 4).Value = ����
            .Offset(, 5).Value = �ݾ�
        End With
        
        MsgBox "�����߽��ϴ�"
    End If
End Sub

Private Sub ListBox_��������ø�_Click()
    CommandButton_�Է�.Enabled = True
    CommandButton_����.Enabled = True
    CommandButton_����.Enabled = True
End Sub

Private Sub TextBox_�ݾ�_Change()
    TextBox_�ݾ�.Value = format(TextBox_�ݾ�.Value, "#,#")
End Sub

Private Sub UserForm_Initialize()
    With ListBox_��������ø�
        .columnCount = 6
        .ColumnWidths = "1cm;2.2cm;2.4cm;2.5cm;2.7cm;1.5cm"
    End With
    Call load_��������ø�
    
    ComboBox_��.AddItem "����"
    ComboBox_��.AddItem "����"
    ComboBox_��.ListIndex = 1
    
    '�� �ʱ�ȭ
    Call ��_�ʱ�ȭ("����")
End Sub

Sub ��_�ʱ�ȭ(�� As String)
    Call UserForm_����.��_�ʱ�ȭ(��, UserForm_���־��������)
End Sub

Sub ��_�ʱ�ȭ(�� As String, �� As String)
    Call UserForm_����.��_�ʱ�ȭ(��, ��, UserForm_���־��������)
End Sub

Sub ����_�ʱ�ȭ(�� As String, �� As String, �� As String)
    Call UserForm_����.����_�ʱ�ȭ(��, ��, ��, UserForm_���־��������)
End Sub

Sub load_��������ø�()
    Dim ws As Worksheet
    Set ws = Worksheets("����")
    Dim ������ As Range
    Set ������ = ws.Range("���ø��������̺�")
    Dim vlist() As Variant
    Dim x As Integer
    
    If ������.Offset(1, 0).Value <> "" Then
        '�� �� �� ���� ���� �ݾ�
        
        Do
            ReDim Preserve vlist(7, x)
            Set ������ = ������.Offset(1, 0)
            vlist(0, x) = ������.Value
            vlist(1, x) = ������.Offset(0, 1).Value
            vlist(2, x) = ������.Offset(0, 2).Value
            vlist(3, x) = ������.Offset(0, 3).Value
            vlist(4, x) = ������.Offset(0, 4).Value
            vlist(5, x) = ������.Offset(0, 5).Value
            vlist(6, x) = ������.Row - ws.Range("���ø��������̺�").Row
            x = x + 1
        Loop While Not IsEmpty(������.Offset(1, 0).Value)
        
        ListBox_��������ø�.Column = vlist
    Else
        ListBox_��������ø�.Clear
    End If
End Sub
