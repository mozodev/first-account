VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_������� 
   Caption         =   "�������"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505.001
   OleObjectBlob   =   "UserForm_�������.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public save_error As Integer

Private Sub ComboBox_guan_AfterUpdate()
    Call ��_�ʱ�ȭ(ComboBox_guan.Value)
End Sub

Private Sub ComboBox_hang_AfterUpdate()
    Call ��_�ʱ�ȭ(ComboBox_hang.Value)
End Sub

Private Sub ComboBox_hang_Change()
    ComboBox_mok.Clear
    ComboBox_semok.Clear
End Sub

Private Sub ComboBox_mok_AfterUpdate()
    Call ����_�ʱ�ȭ(ComboBox_hang.Value, ComboBox_mok.Value)
End Sub

Private Sub ComboBox_mok_Change()
    ComboBox_semok.Clear
End Sub

Private Sub ComboBox_semok_Change()
    TextBox_�ڵ�.Value = get_code("����", ComboBox_hang.Value, ComboBox_mok.Value, ComboBox_semok.Value)
    If TextBox_�ڵ�.Value <> "" Then
        
    End If
End Sub

Private Sub CommandButton1_Click()
'�Է� ��ư Ŭ����
    Call ����
    If save_error = 0 Then
        MsgBox "�ԷµǾ����ϴ�"
        Call �ʱ�ȭ
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    If Parent = "������Ǵ���" Then
        Worksheets("������Ǵ���").Activate
    Else
        Ȩ
    End If
End Sub

Sub ����()
    save_error = 0
    Dim �ڵ� As String, �� As String, �� As String, �� As String, ���� As String, ����� As String
    Dim ���� As Integer
    Dim �ܰ� As Long, �ݾ� As Long
    
    �ڵ� = TextBox_�ڵ�.Value
    �� = "����"
    �� = ComboBox_hang.Value
    �� = ComboBox_mok.Value
    ����� = TextBox_�����.Value
    
    If IsEmpty(TextBox_�ݾ�.Value) Or Not IsNumeric(TextBox_�ݾ�.Value) Then
        MsgBox "������ �ܰ��� �Է��� �ݾ��� ä���ּ���"
        save_error = 3
        Exit Sub
    End If
    �ݾ� = CLng(TextBox_�ݾ�.Value)
    
    If IsEmpty(TextBox_����.Value) Or Not IsNumeric(TextBox_����.Value) Then
        MsgBox "������ ���ڷ� �Է����ּ���"
        save_error = 3
        Exit Sub
    End If
    ���� = CInt(TextBox_����.Value)
    
    If IsEmpty(TextBox_�ܰ�.Value) Or Not IsNumeric(TextBox_�ܰ�.Value) Then
        MsgBox "�ܰ��� ���ڷ� �Է����ּ���"
        save_error = 3
        Exit Sub
    End If
    �ܰ� = CLng(TextBox_�ܰ�.Value)
    
    ���� = ComboBox_semok.Value
    
    If �ڵ� = "" Then
        MsgBox "���׸��� �������ּ���"
        save_error = 1
        Exit Sub
    End If
    
    If �� = "" Then
        MsgBox "���׸��� �������ּ���"
        save_error = 1
        Exit Sub
    End If
    
    If ����� = "" Then
        MsgBox "������� �Է����ּ���"
        save_error = 2
        Exit Sub
    End If
    
    �ڵ� = �ڵ� & "/" & �� & "/" & �� & "/" & �� & "/" & ����
    Dim r������ġ As Range
    With Worksheets("������Ǵ���").Range("���ǳ�¥���̺�")
        If .Offset(1, 0).Value = "" Then
            Set r������ġ = .Offset(1, 0)
        Else
            Set r������ġ = .End(xlDown).Offset(1)  '�� ���Ͽ��� ó���ϰ�
        End If
        
    End With
    
    Worksheets("������Ǵ���").Unprotect
    With r������ġ
        .Value = TextBox_��¥.Value
        .Offset(, 1).Value = �ڵ�
        .Offset(, 2).Value = �����
        .Offset(, 3).Value = TextBox_�԰�.Value
        .Offset(, 4).Value = ����
        .Offset(, 5).Value = �ܰ�
        .Offset(, 6) = "=RC[-2] * RC[-1]" '�ݾ�
        .Offset(, 7).Value = TextBox_���.Value
        .Offset(, 8).Value = TextBox_�ϴܺ��.Value
    End With
    
    Call ������Ǵ�������
End Sub

Sub �ʱ�ȭ()
    Dim ��Ʈ�� As Control
    For Each ��Ʈ�� In UserForm_�������.Controls
        If TypeOf ��Ʈ�� Is MSForms.TextBox Then ��Ʈ��.Value = ""
        If TypeOf ��Ʈ�� Is MSForms.combobox Then ��Ʈ��.Value = ""
    Next
End Sub

Sub ��_�ʱ�ȭ(�� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")

    Dim �׼� As Integer
    Dim �����׸� As String
    �׼� = ws.Range("c4").CurrentRegion.Rows.Count
    �����׸� = ""
    
    For Each ���׸� In ws.Range("c4", "c" & �׼�)
        If ���׸�.Value <> "" Then
            If ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                ComboBox_hang.AddItem ���׸�.Value
                �����׸� = ���׸�.Value
            End If
        End If
    Next ���׸�
    
End Sub

Sub ��_�ʱ�ȭ(�� As String)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets("���꼭")

        Dim ��� As Integer
        Dim �����׸� As String
        ��� = ws.Range("d4").CurrentRegion.Rows.Count
        �����׸� = ""
        
        For Each ���׸� In ws.Range("d4", "d" & ���)
            If ���׸�.Value <> "" Then
                If ���׸�.Offset(, -2).Value = "����" And ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                    ComboBox_mok.AddItem ���׸�.Value
                    �����׸� = ���׸�.Value
                End If
            End If
        Next ���׸�
    
End Sub

Sub ����_�ʱ�ȭ(�� As String, �� As String)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets("���꼭")

        Dim ����� As Integer
        Dim �����׸� As String
        ����� = ws.Range("e4").CurrentRegion.Rows.Count
        �����׸� = ""

        For Each ���׸� In ws.Range("e4", "e" & �����)
            If ���׸�.Value <> "" Then
                If ���׸�.Offset(, -3).Value = "����" And ���׸�.Offset(, -2).Value = �� And ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                    ComboBox_semok.AddItem ���׸�.Value
                    �����׸� = ���׸�.Value
                End If
            End If
        Next ���׸�
    
End Sub

Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ '��/��/��'�� �����Ͻø� �ڵ����� �Էµ˴ϴ�."
End Sub

Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������Ǽ� ���� �Ʒ��ʿ� ǥ�õǴ� �����Դϴ�"
End Sub

Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "����, ����, ȸ �� ������ �����ּ���"
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ �ܰ��� �Է��Ͻø� �ڵ����� ���˴ϴ�"
End Sub

Private Sub TextBox_�ݾ�_Change()
    TextBox_�ݾ�.Value = format(TextBox_�ݾ�.Value, "#,#")
End Sub

Private Sub TextBox_�ݾ�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ �ܰ��� �Է��Ͻø� �ڵ����� ���˴ϴ�"
End Sub

Private Sub TextBox_�ܰ�_AfterUpdate()
    Dim ���� As Integer
    Dim �ܰ� As Long
    If Not IsNull(TextBox_����.Value) And IsNumeric(TextBox_����.Value) Then
        ���� = TextBox_����.Value
    Else
        ���� = 1
    End If
    
    If Not IsNull(TextBox_�ܰ�.Value) And IsNumeric(TextBox_�ܰ�.Value) Then
        �ܰ� = TextBox_�ܰ�.Value
    Else
        �ܰ� = 0
    End If
    
    If IsEmpty(�ܰ�) Or Not IsNumeric(�ܰ�) Then
        MsgBox "���ڸ� �Է��ϼž� �մϴ�"
        save_error = 3
        Exit Sub
    End If

    If (Not IsEmpty(����) And ���� > 0) Then
        If (�ܰ� > 0) Then
            TextBox_�ݾ�.Value = �ܰ� * ����
        End If
    End If
End Sub

Private Sub TextBox_�ܰ�_Change()
    TextBox_�ܰ�.Value = format(TextBox_�ܰ�.Value, "#,#")
End Sub

Private Sub TextBox_����_AfterUpdate()
    Dim ���� As Integer
    Dim �ܰ� As Integer
    
    If Not IsNull(TextBox_����.Value) And IsNumeric(TextBox_����.Value) Then
        ���� = TextBox_����.Value
    Else
        ���� = 1
    End If
    
    If Not IsNull(TextBox_�ܰ�.Value) And IsNumeric(TextBox_�ܰ�.Value) Then
        �ܰ� = TextBox_�ܰ�.Value
    Else
        �ܰ� = 0
    End If
    
    If IsEmpty(����) Or Not IsNumeric(����) Then
        MsgBox "���ڸ� �Է��ϼž� �մϴ�"
        save_error = 3
        Exit Sub
    End If
    
    If (Not IsEmpty(�ܰ�) And �ܰ� > 0) Then
        If (���� > 0) Then
            TextBox_�ݾ�.Value = �ܰ� * ����
        End If
    End If
End Sub

Private Sub TextBox_�ڵ�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ '��/��/��'�� �����Ͻø� �ڵ����� �Էµ˴ϴ�"
End Sub

Private Sub TextBox_�ϴܺ��_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������Ǽ� ���� �Ʒ��ʿ� ǥ�õǴ� �����Դϴ�"
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")

    TextBox_����.Value = 1
    TextBox_�ܰ�.Value = 0
    Call ��_�ʱ�ȭ("����")
    Call ������Ʈ_�ʱ�ȭ
    Call �μ�_�ʱ�ȭ
    
    TextBox_��¥.Value = Date
End Sub

Sub ������Ʈ_�ʱ�ȭ()
    Dim ������ As Range
    Dim ������ As Range
    Dim ������Ʈ As Range
    Dim ws As Worksheet
    Set ws = Worksheets("����")
    Dim ������Ʈ�� As Integer
    
    With ws.Range("������Ʈ�������̺�")
        If .Offset(1, 0).Value <> "" Then
            ������Ʈ�� = .CurrentRegion.Rows.Count - 1
        Else
            ������Ʈ�� = 0
        End If
        
        If ������Ʈ�� > 0 Then
            ComboBox_������Ʈ.Enabled = True
            Set ������ = .Offset(1)
            If ������.Offset(1, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
            Else
                Set ������ = ������
            End If
            
            For Each ������Ʈ In ws.Range(������, ������)
                If ������Ʈ.Value <> "" And ������Ʈ.Value <> "������Ʈ��" Then
                    ComboBox_������Ʈ.AddItem ������Ʈ.Value
                End If
            Next ������Ʈ
        Else
            ComboBox_������Ʈ.Enabled = False
        End If
    End With
End Sub

Sub �μ�_�ʱ�ȭ()
    Dim ������ As Range
    Dim ������ As Range
    Dim �μ� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("����")
    Dim �μ��� As Integer
    
    With ws.Range("�μ��������̺�")
        If .Offset(1, 0).Value <> "" Then
            �μ��� = .CurrentRegion.Rows.Count - 1
        Else
            �μ��� = 0
        End If
        
        If �μ��� > 0 Then
            ComboBox_�μ�.Enabled = True
            Set ������ = .Offset(1)
            If ������.Offset(1, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
            Else
                Set ������ = ������
            End If
            
            For Each �μ� In ws.Range(������, ������)
                If �μ�.Value <> "" And �μ�.Value <> "�μ���" Then
                    ComboBox_�μ�.AddItem �μ�.Value
                End If
            Next �μ�
        Else
            ComboBox_�μ�.Enabled = False
        End If
    End With
End Sub
