VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_����ݳ��� 
   Caption         =   "����� ���� �Է�"
   ClientHeight    =   8505.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670.001
   OleObjectBlob   =   "UserForm_����ݳ���.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_����ݳ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim save_error As Integer
Dim ���� As String
Const �����ͽ�Ʈ As String = "ȸ�����"
Const ���ؼ� As String = "�����ʵ巹�̺�"
Const ����ټ� As Integer = 5
Const ��ü��� As Integer = 20000
    
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
        TextBox_summary.SetFocus
    End If
End Sub

Private Sub ComboBox_�μ�_Change()
    ComboBox_guan.SetFocus
End Sub

Private Sub ComboBox_������Ʈ_Change()
    ComboBox_�μ�.SetFocus
End Sub

Private Sub CommandButton_next_Click()
    Call move_next
End Sub

Sub move_next()
    Dim ���ȣ As Integer
    On Error Resume Next
    If TextBox_���ȣ.Value Then
        ���ȣ = TextBox_���ȣ.Value + 1
        With Worksheets(�����ͽ�Ʈ).Range("a" & ���ȣ)
            If .Offset(0, 0).Value <> "" Then
                Call load_����ݷ��ڵ�(���ȣ)
            Else
                MsgBox "�� �Է°��Դϴ�(���� �Է°��� �����ϴ�)"
            End If
        End With
    Else
        Exit Sub
    End If

End Sub

Private Sub CommandButton_prev_Click()
    Call move_prev
End Sub

Sub move_prev()
    Dim ���ȣ As Integer
    If TextBox_���ȣ.Value Then
        ���ȣ = TextBox_���ȣ.Value - 1
    Else
        ���ȣ = 5
    End If
    
    With Worksheets(�����ͽ�Ʈ).Range("a" & ���ȣ)
        If .Offset(0, 0).Value <> "����" Then
            Call load_����ݷ��ڵ�(���ȣ)
            If ���ȣ < 8 Then  '�����̿��� �����Ա��� ���� ���� �ʵ��� ��ȣ
                CommandButton_����.Enabled = False
            End If
        Else
            MsgBox "ù �Է°��Դϴ�(���� �Է°��� �����ϴ�)"
        End If
    End With
End Sub

Private Sub CommandButton_�˻�_Click()
    Dim ��ü As Range
    Dim ã����¥ As Range
    Dim ���ڵ� As Range
    Dim cell As Range
    Dim x As Integer, y As Integer
    Dim Ű���� As String
    Dim ù��ġ As String
    Dim vlist() As Variant
    
    Ű���� = TextBox_search.Value
    y = 0
    
    If (Len(Ű����) > 0) Then
        Set ��ü = Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").CurrentRegion.columns(1)
        Set ã����¥ = ��ü.Find(What:=Ű����, LookAt:=xlPart)
        
        If Not ã����¥ Is Nothing Then
            ù��ġ = ã����¥.Address
            
            Do
                ReDim Preserve vlist(8, x)
                Set ���ڵ� = ã����¥.Resize(1, 10)
                vlist(0, x) = ���ڵ�.Row
                vlist(1, x) = ���ڵ�.Cells(, ��offset_��¥ + 1)
                vlist(2, x) = ���ڵ�.Cells(, ��offset_�� + 1)
                vlist(3, x) = ���ڵ�.Cells(, ��offset_�� + 1)
                vlist(4, x) = ���ڵ�.Cells(, ��offset_�� + 1)
                vlist(5, x) = ���ڵ�.Cells(, ��offset_���� + 1)
                vlist(6, x) = ���ڵ�.Cells(, ��offset_���� + 1)
                vlist(7, x) = ���ڵ�.Cells(, ��offset_���� + 1)
                
                x = x + 1
                y = 0
                
                Set ã����¥ = ��ü.FindNext(ã����¥)
            Loop While Not ã����¥ Is Nothing And ã����¥.Address <> ù��ġ
            
            ListBox1.Column = vlist
        Else
            MsgBox "�˻������ �������� �ʽ��ϴ�"
            ListBox1.Clear
        End If
    End If
End Sub

Private Sub CommandButton_�˻�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�˻��� ��¥�� �Է��� �� '�˻�' ��ư�� �����ּ���"
End Sub

Private Sub CommandButton_����_Click()
    Dim ���ȣ As Integer
    Dim r������ġ As Range
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)

    ���ȣ = 0
    If TextBox_���ȣ.Value Then
        ���ȣ = TextBox_���ȣ.Value
    End If
    
    If Not ���ȣ > 0 Then
        MsgBox "������ �� �����ϴ�"
        Exit Sub
    Else
        If MsgBox("�����ϰڽ��ϱ�? (" & TextBox_date.Value & "/" & TextBox_summary.Value & ")", vbYesNo, "����Ȯ��") = vbYes Then
            Set r������ġ = ws.Range(���ؼ�).Offset(���ȣ - ����ټ�)
            With r������ġ
                ws.Unprotect PWD
                ws.Range(.Offset(0, 0), .Offset(0, ��offset_�μ�)).Delete Shift:=xlUp
            End With
            
            Set r������ġ = ws.Range(���ؼ�).Offset(���ȣ - ����ټ�)
            With r������ġ
                Range(.Offset(-1, ��offset_�����ܾ�), .Offset(��ü���, ��offset_���ܾ�)).Select
                Selection.FillDown
            End With
            MsgBox "�����Ǿ����ϴ�"
            r������ġ.Select
        End If
    End If
    
    Set r������ġ = ws.Range(���ؼ�).Offset(���ȣ - ����ټ�) '2016.10.4 �߰�(��Ÿ�� ����)
    If r������ġ.Value <> "" And r������ġ.Value <> "����" Then
        Call load_����ݷ��ڵ�(r������ġ.Row)
    Else
        Call �ʱ�ȭ
    End If
    
End Sub

Private Sub CommandButton_�ű�_Click()
    Call �ʱ�ȭ
End Sub

Private Sub CommandButton_�ű�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�Ʒ��� �ҷ��� �ڷᰡ ������ ���� ���Ӱ� �Է��մϴ�"
End Sub

Private Sub CommandButton_����_Click()
    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox1
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 0)
            Call load_����ݷ��ڵ�(���ȣ)
        End If
    End With
End Sub

Sub load_����ݷ��ڵ�(���ȣ As Integer)
    Dim �ΰ������Կ��� As String
    Dim ������������ As Integer
    
    If IsEmpty(Worksheets(�����ͽ�Ʈ).Range("a" & ���ȣ)) Then
        ���ȣ = Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").End(xlDown).Row
    End If
    
    TextBox_���ȣ.Value = ���ȣ
    
    With Worksheets(�����ͽ�Ʈ).Range("a" & ���ȣ)
        If .Offset(0, ��offset_��¥).Value <> "" And .Offset(0, ��offset_��¥).Value <> "����" Then
            TextBox_date.Value = .Offset(0, ��offset_��¥).Value
            ComboBox_������Ʈ.Value = .Offset(0, ��offset_������Ʈ).Value
            ComboBox_�μ�.Value = .Offset(0, ��offset_�μ�).Value
            
            ComboBox_guan.Value = .Offset(0, ��offset_��).Value
            Call ��_�ʱ�ȭ(ComboBox_guan.Value)
            ComboBox_hang.Value = .Offset(0, ��offset_��).Value
            Call ��_�ʱ�ȭ(ComboBox_guan.Value, ComboBox_hang.Value)
            ComboBox_mok.Value = .Offset(0, ��offset_��).Value
            Call ����_�ʱ�ȭ(ComboBox_guan.Value, ComboBox_hang.Value, ComboBox_mok.Value)
            ComboBox_����.Value = .Offset(0, ��offset_����).Value
            
            TextBox_summary.Value = .Offset(0, ��offset_����).Value
            If ComboBox_guan.Value = "����" Or ComboBox_guan.Value = "����ܼ���" Then
                TextBox_amount.Value = .Offset(0, ��offset_����).Value
            Else
                TextBox_amount.Value = .Offset(0, ��offset_����).Value
            End If

                    
            ������������ = .Offset(0, ��offset_����).Value
            If ������������ = "0" Then
                ComboBox_out_type.Value = "����"
            ElseIf ������������ = "1" Then
                ComboBox_out_type.Value = "����"
            Else
                ComboBox_out_type.Value = "ī��"
            End If
                
        End If
    End With
    
    If ���ȣ > 7 Then
        CommandButton_����.Enabled = True
    End If
End Sub

Private Sub CommandButton_����_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "���� ��ϻ���(�˻����)���� Ŭ���� ��, �� ��ư�� ������ ������ ��ĥ �� �ֽ��ϴ�"
End Sub

Private Sub CommandButton1_Click()
    Worksheets("ȸ�����").Unprotect PWD
    
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
    Unload UserForm_����ݳ���
    If Parent = "ȸ�����" Then
        
    Else
        Ȩ
    End If
End Sub

Sub ����()
    Dim �� As String, �� As String, �� As String, ���� As String
    Dim ������Ʈ As String, �μ� As String, �ڵ� As String, ���� As String
    Dim �ݾ� As Long
    Dim ���� As String
    Dim ��������� As Integer
    
    save_error = 0
    
    If ComboBox_guan.Value = "" Then
        MsgBox "���׸��� �������ּ���"
        save_error = 1
        Exit Sub
    End If
    �� = ComboBox_guan.Value
    
    If TextBox_summary.Value = "" Then
        MsgBox "���並 �Է����ּ���"
        save_error = 2
        Exit Sub
    End If
    ���� = TextBox_summary.Value
    
    If Not IsNumeric(TextBox_amount.Value) Or Not TextBox_amount.Value > 0 Then
        MsgBox "�ݾ��� �Է����ּ���"
        save_error = 3
        Exit Sub
    End If
    
    �ݾ� = CLng(TextBox_amount.Value)
    ������Ʈ = ComboBox_������Ʈ.Value
    �μ� = ComboBox_�μ�.Value
    �� = ComboBox_hang.Value
    �� = ComboBox_mok.Value
    ���� = ComboBox_����.Value
    
    �ڵ� = get_code(��, ��, ��, ����)
    
    Dim r������ġ As Range
    Set r������ġ = Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").Offset(TextBox_���ȣ.Value - 5)

    Dim ù�� As Integer
    Dim ���� As Integer
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim st As Range
        
    ���׸� = �ڵ� & "/" & �� & "/" & �� & "/" & �� & "/" & ����
    Select Case ComboBox_out_type.Value
        Case "����"
            ��������� = 0
        Case "����"
            ��������� = 1
        Case "ī��"
            ��������� = 2
        Case Else
            ��������� = 0
    End Select
    
    Worksheets(�����ͽ�Ʈ).Unprotect
    
    ���� = TextBox_date.Value
    With r������ġ
        .Value = ����
        .Offset(, ��offset_���׸�).Value = ���׸�

        .Offset(, ��offset_code).NumberFormatLocal = "G/ǥ��"
        .Offset(, ��offset_code).FormulaR1C1 = "=left(RC[-1], 8)"
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

        .Offset(, ��offset_����).Value = ���������

        .Offset(, ��offset_������Ʈ).Value = ������Ʈ
        .Offset(, ��offset_�μ�).Value = �μ�
        
    End With
    
    If Worksheets("����").Range("a2").Offset(, 1).Value = True Then
        Worksheets(�����ͽ�Ʈ).Protect
    End If
        
End Sub

Sub �ʱ�ȭ()
    Dim ��Ʈ�� As Control
    For Each ��Ʈ�� In UserForm_����ݳ���.Controls
        If TypeOf ��Ʈ�� Is MSForms.TextBox Then ��Ʈ��.Value = ""
        If TypeOf ��Ʈ�� Is MSForms.combobox Then ��Ʈ��.Value = ""
    Next
    
    With Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�").End(xlDown)
        TextBox_date.Value = .Value
        TextBox_���ȣ.Value = .Offset(1, 0).Row
    End With
    
    ComboBox_out_type.Value = "����"
    TextBox_date.SetFocus
End Sub

Function get_��Ȳ����(��Ȳ�ڵ� As String)
    Dim ws As Worksheet
    Set ws = Worksheets("��Ȳ����")
    Dim r������ġ As Range
    Set r������ġ = ws.Range("��Ȳ�ڵ巹�̺�")
    Dim ���� As Integer
    ���� = r������ġ.End(xlDown).Row
    For i = 1 To ����
        With r������ġ.Offset(i, 0)
            If .Value = ��Ȳ�ڵ� Then
                get_��Ȳ���� = .Offset(0, 4).Value
                Exit For
            End If
        End With
    Next i
    
    If get_��Ȳ���� = "" Then
        get_��Ȳ���� = "�غ�� ������ �����ϴ�."
    End If
    
End Function

Private Sub Image1_Click()
    MsgBox get_��Ȳ����("�ϻ�_�����_�˻�")
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�˻��� ��¥�� �Է��ϰ� '�˻�'��ư�� �����ּ���"
End Sub

Private Sub TextBox_income_Change()
    TextBox_income.Value = format(TextBox_income.Value, "#,#")
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
    Dim �����׸� As String, ���� As Integer, ��� As Integer
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
        
    With ListBox1
        .columnCount = 8
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;1.5cm;3cm;1cm;1cm"
    End With
    
    With Worksheets(�����ͽ�Ʈ).Range("�����ʵ巹�̺�")
        If .Offset(1, 0).Value <> "" Then
            TextBox_���ȣ.Value = .End(xlDown).Offset(1, 0).Row
        Else
            TextBox_���ȣ.Value = .Offset(1, 0).Row
        End If
    End With
    
    TextBox_date.Value = Date
    ComboBox_out_type.List = Array("����", "����", "ī��")
    ComboBox_out_type.Value = "����"
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Sub ��_�ʱ�ȭ(�� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    
    Dim �׼� As Integer, �����׸� As String

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

