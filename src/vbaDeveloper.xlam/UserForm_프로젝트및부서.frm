VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_������Ʈ�׺μ� 
   Caption         =   "������Ʈ/���� ����"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   OleObjectBlob   =   "UserForm_������Ʈ�׺μ�.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_������Ʈ�׺μ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim save_error As Integer
Dim ������Ʈ�Է����ȣ As Integer
Dim �μ��Է����ȣ As Integer
Const �����ͽ�Ʈ As String = "����"
Const ������Ʈ���ؼ� As String = "������Ʈ�������̺�"
Const �μ����ؼ� As String = "�μ��������̺�"
Const ����ټ� As Integer = 5

Private Sub CommandButton_�ݱ�1_Click()
    Unload Me
End Sub

Private Sub CommandButton_�ݱ�2_Click()
    Unload Me
End Sub

Sub ������Ʈ�ʱ�ȭ()
    Dim ws As Worksheet
    Dim ������Ʈ As Range
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim ������ As Range
    Dim ������ As Range
    Dim vlist() As Variant
    Dim ������Ʈ�� As Integer
    Dim x As Integer
    x = 0
    
    With ws.Range(������Ʈ���ؼ�)
        If .Offset(1, 0).Value <> "" Then
            ������Ʈ�� = .CurrentRegion.Rows.Count - 1
        Else
            ������Ʈ�� = 0
        End If
        TextBox_������Ʈ���ȣ.Value = .Offset(������Ʈ�� + 1).Row '�� ������Ʈ�� �� ���ȣ
        
        If ������Ʈ�� > 0 Then
            Set ������ = .Offset(1)
            If ������.Offset(1, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
            Else
                Set ������ = ������
            End If
            
            For Each ������Ʈ In ws.Range(������, ������)
                If ������Ʈ.Value <> "" And ������Ʈ.Value <> "������Ʈ��" Then
                    ReDim Preserve vlist(2, x)
                    vlist(0, x) = ������Ʈ.Row
                    vlist(1, x) = ������Ʈ.Value
                End If
                x = x + 1
            Next ������Ʈ
            ListBox_������Ʈ.Column = vlist
        End If
    End With
    
    TextBox_������Ʈ��.Value = ""
End Sub

Sub �μ��ʱ�ȭ()
    Dim ws As Worksheet
    Dim �μ� As Range
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim ������ As Range
    Dim ������ As Range
    Dim vlist() As Variant
    Dim x As Integer
    x = 0
    
    With ws.Range(�μ����ؼ�)
        If .Offset(1, 0).Value <> "" Then
            �μ��� = .CurrentRegion.Rows.Count - 1
        Else
            �μ��� = 0
        End If
        TextBox_�μ����ȣ.Value = .Offset(�μ��� + 1).Row '�� �μ��� �� ���ȣ
        
        If �μ��� > 0 Then
            Set ������ = .Offset(1)
            If ������.Offset(1, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
            Else
                Set ������ = ������
            End If
            
            For Each �μ� In ws.Range(������, ������)
                If �μ�.Value <> "" And �μ�.Value <> "�μ���" Then
                    ReDim Preserve vlist(2, x)
                    vlist(0, x) = �μ�.Row
                    vlist(1, x) = �μ�.Value
                End If
                x = x + 1
            Next �μ�
            ListBox_�μ�.Column = vlist
        End If
    End With
    
    TextBox_�μ���.Value = ""
End Sub

Private Sub CommandButton_�μ�����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox_�μ�
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 0)
            Worksheets(�����ͽ�Ʈ).Range(�μ����ؼ�).Offset(���ȣ - ����ټ�).Delete Shift:=xlUp
            Call �μ��ʱ�ȭ
        End If
    End With
End Sub

Private Sub CommandButton_�μ�����_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    If TextBox_�μ���.Value <> "" Then
        ws.Range(�μ����ؼ�).Offset(TextBox_�μ����ȣ.Value - ����ټ�).Value = TextBox_�μ���.Value
        Call �μ��ʱ�ȭ
    Else
        MsgBox "�μ����� �Է����ּ���"
    End If
    TextBox_�μ���.SetFocus
End Sub

Private Sub CommandButton_�μ�����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox_�μ�
        i���� = .ListIndex
        If i���� > -1 Then
            TextBox_�μ����ȣ.Value = .List(i����, 0)
            TextBox_�μ���.Value = .List(i����, 1)
        End If
    End With
    TextBox_�μ���.SetFocus
End Sub

Private Sub CommandButton_�űԺμ�_Click()
    TextBox_�μ����ȣ.Value = ���Է����ȣ("�μ�")
    TextBox_�μ���.Value = ""
    TextBox_�μ���.SetFocus
End Sub

Private Sub CommandButton_�ű�������Ʈ_Click()
    TextBox_������Ʈ���ȣ.Value = ���Է����ȣ("������Ʈ")
    TextBox_������Ʈ��.Value = ""
    TextBox_������Ʈ��.SetFocus
End Sub

Private Sub CommandButton_�ű�������Ʈ_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�ҷ��� ���� ���� ���ο� ������Ʈ�� ����ϴ�"
End Sub

Private Sub CommandButton_������Ʈ����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox_������Ʈ
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 0)
            Worksheets(�����ͽ�Ʈ).Range(������Ʈ���ؼ�).Offset(���ȣ - ����ټ�).Delete Shift:=xlUp
            Call ������Ʈ�ʱ�ȭ
        End If
    End With
End Sub

Private Sub CommandButton_������Ʈ����_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    If TextBox_������Ʈ��.Value <> "" Then
        ws.Range(������Ʈ���ؼ�).Offset(TextBox_������Ʈ���ȣ.Value - ����ټ�).Value = TextBox_������Ʈ��.Value
        Call ������Ʈ�ʱ�ȭ
    Else
        MsgBox "������Ʈ���� �Է����ּ���"
    End If
    TextBox_������Ʈ��.SetFocus
End Sub

Private Sub CommandButton_������Ʈ����_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "���ο� �̸����� ������Ʈ�� �߰� Ȥ�� �����մϴ�"
End Sub

Private Sub CommandButton_������Ʈ����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox_������Ʈ
        i���� = .ListIndex
        If i���� > -1 Then
            TextBox_������Ʈ���ȣ.Value = .List(i����, 0)
            TextBox_������Ʈ��.Value = .List(i����, 1)
        End If
    End With
    
    TextBox_������Ʈ��.SetFocus
End Sub

Function ���Է����ȣ(���� As String)
    Dim ���ؼ� As String
    If ���� = "������Ʈ" Then
        ���ؼ� = ������Ʈ���ؼ�
    Else
        ���ؼ� = �μ����ؼ�
    End If
    
    With Worksheets(�����ͽ�Ʈ).Range(���ؼ�)
        If .Offset(1, 0).Value <> "" Then
            ���Է����ȣ = .End(xlDown).Offset(1, 0).Row
        Else
            ���Է����ȣ = .Offset(1, 0).Row
        End If
    End With
End Function

Private Sub CommandButton_������Ʈ����_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ ������Ʈ �̸��� �ҷ��ɴϴ�"
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������Ʈ �̸��� ���� '����'�� ������ �ݿ��˴ϴ�"
End Sub

Private Sub TextBox_������Ʈ��_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�̰��� ������Ʈ�� ���ο� �̸��� �����ּ���"
End Sub

Private Sub UserForm_Initialize()
    With ListBox_������Ʈ
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_�μ�
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With Worksheets(�����ͽ�Ʈ).Range(������Ʈ���ؼ�)
        If .Offset(1, 0).Value <> "" Then
            ������Ʈ�Է����ȣ = .End(xlDown).Offset(1, 0).Row
        Else
            ������Ʈ�Է����ȣ = .Offset(1, 0).Row
        End If
    End With
    
    Call ������Ʈ�ʱ�ȭ
    
    With Worksheets(�����ͽ�Ʈ).Range(�μ����ؼ�)
        If .Offset(1, 0).Value <> "" Then
            �μ��Է����ȣ = .End(xlDown).Offset(1, 0).Row
        Else
            �μ��Է����ȣ = .Offset(1, 0).Row
        End If
    End With
    
    Call �μ��ʱ�ȭ
    
    MultiPage2.Value = 0  'ù������(������Ʈ����)�� �׻� ���� �ߵ���
End Sub

