VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_������Է� 
   Caption         =   "ȸ����� �Է� "
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550.001
   OleObjectBlob   =   "UserForm_������Է�.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_������Է�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private save_error As Integer
Const �����ͽ�Ʈ As String = "ȸ�����"
Const ���ؼ� As String = "�����ʵ巹�̺�"
Const ����ټ� As Integer = 5
Const ��ü��� As Integer = 4000
    
Const ��offset_��¥ As Integer = 0
Const ��offset_���׸� As Integer = 1
Const ��offset_code As Integer = 2
Const ��offset_�� As Integer = 3
Const ��offset_�� As Integer = 4
Const ��offset_�� As Integer = 5
Const ��offset_���� As Integer = 6 '�� �߰� : 2015.3.8
Const ��offset_���� As Integer = 7
Const ��offset_���� As Integer = 8
Const ��offset_���� As Integer = 9
Const ��offset_���� As Integer = 10
Const ��offset_VAT As Integer = 11
Const ��offset_���� As Integer = 12
Const ��offset_������Ʈ As Integer = 13
Const ��offset_�μ� As Integer = 14
Const ��offset_�����ܾ� As Integer = 15
Const ��offset_�����ܾ� As Integer = 16
Const ��offset_���ܾ� As Integer = 17

Private Sub CheckBox_�ΰ�������_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "���� ������ ��� �ΰ����� ���Ե� ������ ���������� üũ���ּ���"
End Sub

Private Sub ComboBox_guan_change()
    Dim �� As String
    �� = ComboBox_guan.Value
    
    If �� = "����" Or �� = "����" Then
        Call ��_�ʱ�ȭ(��)
        ComboBox_hang.Enabled = True
        ComboBox_mok.Enabled = True
        ComboBox_����.Enabled = True
        ComboBox_hang.SetFocus
    Else '����ܼ���/����
        ComboBox_hang.Enabled = False
        ComboBox_mok.Enabled = False
        ComboBox_����.Enabled = False
        TextBox_summary.SetFocus
    End If
    
End Sub

Private Sub ComboBox_hang_Change()
    If ComboBox_hang.Enabled Then
        Call ��_�ʱ�ȭ(ComboBox_guan.Value, ComboBox_hang.Value)
        ComboBox_mok.SetFocus
    End If
End Sub

Private Sub ComboBox_mok_Change()
    If ComboBox_mok.Enabled Then
        Call ����_�ʱ�ȭ(ComboBox_guan.Value, ComboBox_hang.Value, ComboBox_mok.Value)
        ComboBox_����.SetFocus
    End If
End Sub

Private Sub ComboBox_����_Change()
    If ComboBox_����.Enabled Then
        MultiPage1.Value = 0
        TextBox_summary.SetFocus
    End If
End Sub

Private Sub ComboBox_�μ�_Change()
    'ComboBox_guan.SetFocus
End Sub

Private Sub ComboBox_������Ʈ_Change()
    ComboBox_�μ�.SetFocus
End Sub

Private Sub CommandButton_close_Click()
    Unload Me
    If Parent <> "ȸ�����" Then
        Ȩ
    End If
End Sub

Private Sub CommandButton1_Click()
    If IsEmpty(ComboBox_hang.Value) Or ComboBox_hang.Value = "" Then
        MsgBox "���� �������ּ���"
        MultiPage1.Value = 0
        ComboBox_hang.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ComboBox_mok.Value) Or ComboBox_mok.Value = "" Then
        MsgBox "���� �������ּ���"
        MultiPage1.Value = 0
        ComboBox_mok.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(ComboBox_����.Value) Or ComboBox_����.Value = "" Then
        MsgBox "������ �������ּ���"
        MultiPage1.Value = 0
        ComboBox_����.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(TextBox_summary.Value) Or TextBox_summary.Value = "" Then
        MsgBox "���並 �Է����ּ���"
        MultiPage1.Value = 0
        TextBox_summary.SetFocus
        Exit Sub
    End If
    
    If IsEmpty(TextBox_amount.Value) Or TextBox_amount.Value = "" Then
        MsgBox "�ݾ��� �Է����ּ���"
        MultiPage1.Value = 0
        TextBox_amount.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_amount.Value) Or TextBox_amount.Value = 0 Then
        MsgBox "�ݾ��� 0���� ū ���ڷ� �Է����ּ���"
        MultiPage1.Value = 0
        TextBox_amount.Value = ""
        TextBox_amount.SetFocus
        Exit Sub
    End If
    
    Worksheets("ȸ�����").Unprotect PWD
    
    Call ����
    If save_error = 0 Then
        MsgBox "�ԷµǾ����ϴ�"
        Call �ʱ�ȭ
    End If
    
    If (Worksheets("����").Range("��Ʈ��ݼ���").Offset(, 1).Value = True) Then
        Worksheets("ȸ�����").Protect PWD
    End If
End Sub

Private Sub CommandButton2_Click()
    '���Էºκп� �ִ� "ȸ������Է�" ��ư
    Call CommandButton1_Click
End Sub

Sub ����()
    Dim �� As String, �� As String, �� As String, ���� As String
    Dim �ڵ� As String
    Dim ������Ʈ As String, �μ� As String
    Dim ���� As String
    Dim �ݾ� As Long
    
    If ComboBox_guan.Value = "" Then
        MsgBox "���׸��� �������ּ���"
        save_error = 1
        Exit Sub
    End If
    
    If TextBox_summary.Value = "" Then
        MsgBox "���並 �Է����ּ���"
        save_error = 2
        Exit Sub
    End If
    
    If Not IsNumeric(TextBox_amount.Value) Or Not TextBox_amount.Value > 0 Then
        MsgBox "�ݾ��� �Է����ּ���"
        save_error = 3
        Exit Sub
    End If
    
    ������Ʈ = ComboBox_������Ʈ.Value
    �μ� = ComboBox_�μ�.Value
    �� = ComboBox_guan.Value
    �� = ComboBox_hang.Value
    �� = ComboBox_mok.Value
    ���� = ComboBox_����.Value
    �ڵ� = get_code(��, ��, ��, ����)
    ���� = TextBox_summary.Value
    �ݾ� = TextBox_amount.Value
    
    Dim ���� As String
    Dim ��������� As Integer
    
    save_error = 0
    
    Dim r������ġ As Range
    Set r������ġ = Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").Offset(TextBox_���ȣ.Value - 5)
    Dim ù�� As Integer
    Dim ���� As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim st As Range
    Dim ���׸� As String
    Dim newData() As Variant
    ReDim newData(1 To 14)
        
    ���׸� = �ڵ� & "/" & �� & "/" & �� & "/" & �� & "/" & ����
    newData(��offset_���׸�) = ���׸�
    newData(��offset_code) = �ڵ�
    newData(��offset_��) = ��
    newData(��offset_��) = ��
    newData(��offset_��) = ��
    newData(��offset_����) = ����

    If CheckBox_���ݿ���.Value Then
        ��������� = 1
    Else
        ��������� = 0
    End If
    newData(��offset_����) = ���������
    newData(��offset_������Ʈ) = ������Ʈ
    newData(��offset_�μ�) = �μ�
    newData(��offset_����) = ����
    If �� = "����" Or �� = "����ܼ���" Then
        newData(��offset_����) = �ݾ�
    Else
        newData(��offset_����) = �ݾ�
    End If
    
    Worksheets(�����ͽ�Ʈ).Unprotect
    
    ���� = TextBox_date.Value
    With r������ġ
        .Value = ����
        .Offset(, ��offset_code).NumberFormat = "@"
        Worksheets(�����ͽ�Ʈ).Range(.Offset(0, 1), .Offset(0, 13)).Value = newData
    End With
    
    If Worksheets("����").Range("a2").Offset(, 1).Value = True Then
        Worksheets(�����ͽ�Ʈ).Protect
    End If
        
End Sub

'���� ���� ����� ������ �� �Լ��� ����Ѵ�
Sub ȸ������Է�(ByVal �� As String, ByVal �� As String, ByVal �� As String, ByVal ���� As String, ByVal ���� As String, ByVal �ݾ� As Long)
    Dim ws As Worksheet
    Set ws = Worksheets("ȸ�����")
    Dim ������ As Range
    Set ������ = ws.Range("�����ʵ巹�̺�").End(xlDown).Offset(1, 0)
    Dim ���׸�, �ڵ� As String
    
    �ڵ� = get_code(��, ��, ��, ����)
    If �ڵ� = "" Then
        MsgBox "��ϵ��� ���� ���׸��Դϴ�. �����Ͻ� ���׸��� �´��� Ȯ�����ּ���"
        Exit Sub
    End If
    ���׸� = �ڵ� & "/" & �� & "/" & �� & "/" & �� & "/" & ����
    
    ws.Unprotect PWD

    With ������
        .Value = Date
        .Offset(, ��offset_���׸�).Value = ���׸�
        .Offset(, ��offset_code).NumberFormat = "@"
        .Offset(, ��offset_code).Value = �ڵ�
        .Offset(, ��offset_��).Value = ��
        .Offset(, ��offset_��).Value = ��
        .Offset(, ��offset_��).Value = ��
        .Offset(, ��offset_����).Value = ����
        .Offset(, ��offset_����).Value = ����
        If �� = "����" Or �� = "����ܼ���" Then
            .Offset(, ��offset_����).Value = �ݾ�
        Else
            .Offset(, ��offset_����).Value = �ݾ�
        End If
        .Offset(, ��offset_����).Value = 0
    End With
    
    If (Worksheets("����").Range("��Ʈ��ݼ���").Offset(, 1).Value = True) Then
        ws.Protect PWD
    End If
    MsgBox "ȸ����忡 �ԷµǾ����ϴ�"
End Sub

Sub �ʱ�ȭ()
    Dim ��Ʈ�� As Control
    For Each ��Ʈ�� In Me.Controls
        If TypeOf ��Ʈ�� Is MSForms.TextBox Then ��Ʈ��.Value = ""
        If TypeOf ��Ʈ�� Is MSForms.combobox Then ��Ʈ��.Value = ""
    Next
    
    With Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").End(xlDown)
        TextBox_date.Value = .Value
        TextBox_���ȣ.Value = .Offset(1, 0).Row
    End With
    
    CheckBox_���ݿ���.Value = False '���� ��κ� �ŷ��� ���ݰŷ��� �ƴϹǷ� �������� �ʱ�ȭ
    TextBox_date.SetFocus
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�˻��� ��¥�� �Է��ϰ� '�˻�'��ư�� �����ּ���"
End Sub

Private Sub TextBox_income_Change()
    TextBox_income.Value = format(TextBox_income.Value, "#,#")
End Sub

Private Sub CommandButton3_Click()
    Call CommandButton_close_Click
End Sub

Private Sub TextBox_amount_Change()
    TextBox_amount.Value = format(TextBox_amount.Value, "#,#")
End Sub

Private Sub TextBox_outgoings_Change()
    TextBox_outgoings.Value = format(TextBox_outgoings.Value, "#,#")
End Sub

Private Sub TextBox_search_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�˻��� ��¥�� �̰��� �Է��մϴ�"
End Sub

Private Sub UserForm_Initialize()

    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    �����׸� = ""
    ���� = ws.Range("b2").CurrentRegion.Rows.Count
    ��� = ws.Range("d4").CurrentRegion.Rows.Count

    Call ������Ʈ_�ʱ�ȭ
    Call �μ�_�ʱ�ȭ
    
    For Each ���׸� In ws.Range("b2", "b" & ����)
        
        If ���׸�.Value <> "" Then
            If ���׸�.Value <> �����׸� Then
                ComboBox_guan.AddItem ���׸�.Value
                �����׸� = ���׸�.Value
            End If
        End If
        
    Next ���׸�
    
    With Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�")
        If .Offset(1, 0).Value <> "" Then
            TextBox_���ȣ.Value = .End(xlDown).Offset(1, 0).Row
        Else
            TextBox_���ȣ.Value = .Offset(1, 0).Row
        End If
    End With
    
    Call ��_�ʱ�ȭ("����")
    TextBox_date.Value = Date

    MultiPage1.Value = 0
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Sub ��_�ʱ�ȭ(�� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")

    �׼� = ws.Range("c4").CurrentRegion.Rows.Count
    �����׸� = ""
    ComboBox_hang.Clear
    
    For Each ���׸� In ws.Range("c4", "c" & �׼�)
        If ���׸�.Value <> "" Then
            If ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                ComboBox_hang.AddItem ���׸�.Value
                �����׸� = ���׸�.Value
            End If
        End If
    Next ���׸�
    
End Sub

Sub ��_�ʱ�ȭ(�� As String, �� As String)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets("���꼭")

        Dim ��� As Integer
        Dim �����׸� As String
        ��� = ws.Range("d4").CurrentRegion.Rows.Count
        �����׸� = ""
        ComboBox_mok.Clear
        
        For Each ���׸� In ws.Range("d4", "d" & ���)
            With ���׸�
                If .Value <> "" Then
                    If .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value <> �����׸� Then
                        ComboBox_mok.AddItem .Value
                        �����׸� = .Value
                    End If
                End If
            End With
        Next ���׸�
    
End Sub

Sub ����_�ʱ�ȭ(�� As String, �� As String, �� As String)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets("���꼭")

        Dim ����� As Integer
        Dim �����׸� As String
        ����� = ws.Range("e4").CurrentRegion.Rows.Count
        �����׸� = ""
        ComboBox_����.Clear
        
        For Each ���׸� In ws.Range("e4", "e" & �����)
            With ���׸�
                If .Value <> "" Then
                    If .Offset(, -3).Value = �� And .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value <> �����׸� Then
                        ComboBox_����.AddItem .Value
                        �����׸� = .Value
                    End If
                End If
            End With
        Next ���׸�
    
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
