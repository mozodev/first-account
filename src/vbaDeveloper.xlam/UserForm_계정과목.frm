VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�������� 
   Caption         =   "�������� ����"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11430
   OleObjectBlob   =   "UserForm_��������.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ����ټ� As Integer = 3
Const �����ͽ�Ʈ As String = "���꼭"
Const ����ġ As Integer = 1 'offset ��
Const ����ġ As Integer = 2 'offset ��
Const ����ġ As Integer = 3 'offset ��
Const ������ġ As Integer = 4 'offset ��
Const ���񼳸���ġ As Integer = 6
Const ù�� As Integer = 6

Private Sub ComboBox_guan_change()
    Call ��_�ʱ�ȭ(ComboBox_guan.Value)
    CommandButton_��_�ű�.Enabled = True

    TextBox_��.Value = ""
    TextBox_��.Enabled = False
    TextBox_��.Value = ""
    TextBox_��.Enabled = False
    TextBox_����.Value = ""
    TextBox_����.Enabled = False
   
    ListBox_��.Clear
    ListBox_����.Clear
End Sub

Private Sub ComboBox_hang_AfterUpdate()
    TextBox_guan.Enabled = False
    Call ��_�ʱ�ȭ(ComboBox_guan.Value, ComboBox_hang.Value)
    textbox_hang.Value = ComboBox_hang.Value
End Sub

Private Sub ComboBox_guan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "����/������ �������ּ���. �׿� ���� '��'���� ǥ�õ˴ϴ�"
End Sub

Private Sub CommandButton_��_�ű�_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    TextBox_��_���ȣ.Value = ""
    Dim �������� As Integer

    With ws.Range("a4")
        .Select
        �������� = .CurrentRegion.Rows.Count
    End With
    TextBox_��.Enabled = True
    TextBox_��.Value = ""
    TextBox_��_���ȣ.Value = �������� + 1
    CommandButton_�����.Enabled = False
    CommandButton_��_����.Enabled = True
    �������� = 0

    TextBox_��.SetFocus
End Sub

Private Sub CommandButton_����_�ű�_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    TextBox_����_���ȣ.Value = ""
    Dim �������� As Integer

    With ws.Range("a4")
        .Select
        �������� = .CurrentRegion.Rows.Count

    End With
    TextBox_����.Enabled = True
    TextBox_����.Value = ""
    TextBox_����_���ȣ.Value = �������� + 1
    CommandButton_�������.Enabled = False
    CommandButton_����_����.Enabled = True
    �������� = 0

    TextBox_����.SetFocus
End Sub

Private Sub CommandButton_�����_Click()
    Dim ���ȣ As Integer
    Dim r������ġ As Range
    Dim �� As String
    Dim �� As String
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim ���� As Integer
    ���� = ws.Range("���ʵ�").End(xlDown).Row
    
    ���ȣ = TextBox_��_���ȣ.Value
    
    If Not ���ȣ > 0 Then
        MsgBox "������ �� �����ϴ�"
    Else
        Set r������ġ = Worksheets(�����ͽ�Ʈ).Range("���ʵ�").Offset(���ȣ - 1)
        With r������ġ
            �� = .Value
            �� = .Offset(0, 1).Value
            If �� <> "" Then
                Range(.Offset(0, -1), .Offset(0, 6)).Delete Shift:=xlUp
                Call �ٱ߱�(���ȣ)
            End If
        End With
        
        TextBox_��.Value = ""
        TextBox_��_���ȣ.Value = ���� + 1
        
        Call ��_�ʱ�ȭ(��, ��)
        TextBox_��.Enabled = False
    End If
End Sub

Private Sub CommandButton_�������_Click()
    Dim ���ȣ As Integer
    Dim r������ġ As Range
    Dim �� As String
    Dim �� As String
    Dim �� As String
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim ���� As Integer
    ���� = ws.Range("���ʵ�").End(xlDown).Row
    
    ���ȣ = TextBox_����_���ȣ.Value
    
    If Not ���ȣ > 0 Then
        MsgBox "������ �� �����ϴ�"
    Else
        Set r������ġ = ws.Range("���ʵ�").Offset(���ȣ - 1)
        With r������ġ
            �� = .Value
            �� = .Offset(0, 1).Value
            �� = .Offset(0, 2).Value
            If �� <> "" Then
                Range(.Offset(0, -1), .Offset(0, 6)).Delete Shift:=xlUp
                Call �ٱ߱�(���ȣ)
            End If
        End With
        
        TextBox_����.Value = ""
        TextBox_����_���ȣ.Value = ���� + 1
        Call ����_�ʱ�ȭ(��, ��, ��)
        TextBox_����.Enabled = False
        'CommandButton_����_�ű�.Enabled = False '2016.11.9
    End If
End Sub

Private Sub CommandButton_��_�ű�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ ����ϰ� ���ο� '��'�� �߰��Ϸ��� Ŭ���ϼ���. �Է¶��� ���ϴ�"
End Sub

Private Sub CommandButton_��_����_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "���� �Է¶����� '��' �̸��� �����߰ų� ���� �ۼ������� �� ��ư�� �����ּ���"
End Sub

Private Sub CommandButton_��_�ű�_Click()
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim �������� As Integer

    With ws.Range("a4")
        .Select
        �������� = .CurrentRegion.Rows.Count
    End With
    TextBox_��.Enabled = True
    CommandButton_��_����.Enabled = True
    TextBox_��.Value = ""
    TextBox_��_���ȣ.Value = �������� + 1
    TextBox_��.SetFocus
End Sub

Private Sub CommandButton_��_����_Click()
    Dim ws As Worksheet
    Dim �� As String, �� As String, �� As String, ������ As String
    
    Dim ���ȣ As Integer, ���� As Integer, ��� As Integer
    Dim i���� As Integer
    
    Set ws = Worksheets(�����ͽ�Ʈ)
    �� = ComboBox_guan.Value
    With ListBox_��
        i���� = .ListIndex
        If i���� > -1 Then
            �� = .List(i����, 1)
        End If
    End With
    
    With ws.Range("���ʵ�")
        If .Offset(3, 0).Value = "" Then
            ���� = ù��
            ��� = 0
        Else
            ���� = .Offset(3, 0).End(xlDown).Row
            ��� = ���� - ����ټ�
        End If
    End With
    
    ���ȣ = TextBox_��_���ȣ.Value
    ������ = ws.Range("A" & ���ȣ).Offset(, ����ġ).Value
    �� = TextBox_��.Value
    
    If �� = "" Then
        MsgBox "'��'�� �Է����ּ���"
        Exit Sub
    End If
    
    '�� ����
     If ��� > 0 Then
        ��� = 0
        �����׸� = ""
        For Each ���׸� In ws.Range("A" & ���ȣ, "A" & ����)
            If ���׸�.Value <> "" Then
                If ���׸�.Offset(, ����ġ).Value = �� And ���׸�.Offset(, ����ġ).Value = �� And ���׸�.Offset(, ����ġ).Value = ������ Then
                    ��� = ��� + 1
                    ���׸�.Offset(, ����ġ).Value = ��
                End If
            End If
        Next ���׸�
        
    End If
    
    If ��� = 0 Then '�߰�
        With ws.Range("A" & ���ȣ)
            .Offset(0, ����ġ).Value = ComboBox_guan.Value
            .Offset(0, ����ġ).Value = ��
            .Offset(0, ����ġ).Value = �� '�켱 �װ� ���� �̸��� �� ����
            .Offset(0, ������ġ).Value = �� '���� �̸��� ���� ����. (������ �ɼ����� �ϸ� ������ �����ϹǷ� �� ó�� ��� ���� ������)
            Call �ٱ߱�(���ȣ)
        End With
    End If

    MsgBox "�Էµƽ��ϴ�"
    
    '��,��,�� ����
    Call ��������_��Ʈ����
    ws.Range("���׸��ڵ巹�̺�").Select
    code_changed = True
    
    '�� �ʱ�ȭ
    Call ��_�ʱ�ȭ(��, ��)
    TextBox_��_���ȣ.Value = ""
    TextBox_��.Value = ""
    TextBox_��.Enabled = False
End Sub

Private Sub CommandButton_����_����_Click()
    Dim ws As Worksheet
    Dim �� As String, �� As String, �� As String, ���� As String
    Dim ���ȣ As Integer, ���� As Integer, ��� As Integer, i���� As Integer
    
    Set ws = Worksheets(�����ͽ�Ʈ)
    �� = ComboBox_guan.Value
    With ListBox_��
        i���� = .ListIndex
        If i���� > -1 Then
            �� = .List(i����, 1)
        End If
    End With
    
    With ListBox_��
        i���� = .ListIndex
        If i���� > -1 Then
            �� = .List(i����, 1)
        End If
    End With
    
    With ws.Range("�����ʵ�")
        If .Offset(3, 0).Value = "" Then
            ���� = ù��
            ��� = 0
        Else
            ���� = .End(xlDown).Row
            ��� = ���� - ����ټ�
        End If
    End With
    
    If TextBox_����_���ȣ.Value = "" Then
        MsgBox "�߰�/������ �׸��� ��ġ�� �ٽ� �������ּ���"
        Exit Sub
    End If
    ���ȣ = TextBox_����_���ȣ.Value
    
    ���� = TextBox_����.Value
    If ���� = "" Then
        MsgBox "'����'�� �Է����ּ���"
        Exit Sub
    End If
    
    '�� ����
    With ws.Range("A" & ���ȣ)
        If ��� >= 1 And ���ȣ <= ���� Then ' 2�� �̻� �ԷµǾ� �ְ�, ������ ���� �����ϴ� ���
            .Offset(0, ������ġ).Value = ����
        Else
            .Offset(0, 1).Value = ComboBox_guan.Value
            .Offset(0, ����ġ).Value = ��
            .Offset(0, ����ġ).Value = ��
            .Offset(0, ������ġ).Value = ����
            Call �ٱ߱�(���ȣ)
        End If
    End With
    MsgBox "�Էµƽ��ϴ�"
    
    '��,��,�� ����
    Call ��������_��Ʈ����
    ws.Range("���׸��ڵ巹�̺�").Select
    code_changed = True
    
    '�� �ʱ�ȭ
    Call ����_�ʱ�ȭ(��, ��, ��)
    TextBox_����_���ȣ.Value = ""
    TextBox_����.Value = ""
    TextBox_����.Enabled = False
End Sub

Private Sub CommandButton_��_����_Click()
    Dim ws As Worksheet
    Dim �� As String
    Dim ���ȣ As Integer
    Dim ���� As Integer
    Dim �׼� As Integer
    Dim �� As String
    Dim ������ As String
    Dim ���׸� As Range
        
    �� = ComboBox_guan.Value
    Set ws = Worksheets(�����ͽ�Ʈ)
    
    With ws.Range("���ʵ�")
        If .Offset(3, 0).Value = "" Then
            ���� = ù��
            �׼� = 0
        Else
            ���� = .Offset(3, 0).End(xlDown).Row
            �׼� = ���� - ����ټ�
        End If
    End With
    
    ���ȣ = TextBox_��_���ȣ.Value
    ������ = ws.Range("A" & ���ȣ).Offset(, ����ġ).Value
    �� = TextBox_��.Value
    
    If �׼� > 0 Then
        �׼� = 0
        �����׸� = ""
        For Each ���׸� In ws.Range("A" & ���ȣ, "A" & ����)
            If ���׸�.Value <> "" Then
                If ���׸�.Offset(, ����ġ).Value = �� And ���׸�.Offset(, ����ġ).Value = ������ Then
                    �׼� = �׼� + 1
                    ���׸�.Offset(, ����ġ).Value = ��
                End If
            End If
        Next ���׸�
        
    End If
    
    If �׼� = 0 Then '�߰�
        With ws.Range("A" & ���ȣ)
            .Offset(0, ����ġ).Value = ComboBox_guan.Value
            .Offset(0, ����ġ).Value = ��
            .Offset(0, ����ġ).Value = �� '�켱 �װ� ���� �̸��� �� ����
            .Offset(0, ������ġ).Value = �� '���� �̸��� ���� ����. (������ �ɼ����� �ϸ� ������ �����ϹǷ� �� ó�� ��� ���� ������)
            Call �ٱ߱�(���ȣ)
        End With
    End If
    
    '��,��,�� ����
    Call ��������_��Ʈ����
    ws.Range("���׸��ڵ巹�̺�").Select
    code_changed = True
    
    '�� �ʱ�ȭ
    Call ��_�ʱ�ȭ(��)

    TextBox_��.Value = ""
    TextBox_��.Enabled = False
End Sub

Sub �ٱ߱�(���ȣ As Integer)
    Dim r������ġ As Range
    Dim ���� As Range
    Set r������ġ = Range("A" & ���ȣ)
    Set ���� = Range(r������ġ, r������ġ.Offset(0, 6))
    
    With ����
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Private Sub CommandButton_�׻���_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "'��'�Ʒ� '��'�� ���� ��� �������� �ʽ��ϴ�. '��'�� ��� �����ϸ� '��'�� �����˴ϴ�"
End Sub

Private Sub CommandButton1_Click()
    Call ����
    Call �ʱ�ȭ
    Call ���꼭�ڵ��Է�
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm_��������

    If code_changed Then
        Call ���꼭�ڵ��Է�
        Call ��꼭�ʱ�ȭ
        code_changed = False
    End If
    Ȩ
End Sub

Sub �ʱ�ȭ()
    Dim ws As Worksheet
    Dim ���� As Integer
    Dim ��Ʈ�� As Control
    
    For Each ��Ʈ�� In UserForm2.Controls
        If TypeOf ��Ʈ�� Is MSForms.TextBox Then ��Ʈ��.Value = ""
    Next
    
    Set ws = Worksheets(�����ͽ�Ʈ)
    ���� = ws.Range("���׸��ڵ巹�̺�").End(xlDown).Row + 1
    
    TextBox_��_���ȣ.Value = ����
    TextBox_��_���ȣ.Value = ����
    ���� = 0
End Sub

Sub ����()
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    
    Dim r������ġ As Range
    Set r������ġ = ws.Range("���׸��ڵ巹�̺�").End(xlDown).Offset(1)
    
    Dim �� As String, �� As String, �� As String, ���� As String
    Dim ����� As String
    �� = ComboBox_guan.Value
    �� = TextBox_��.Value
    �� = TextBox_��.Value
    ���� = TextBox_����.Value
    ����� = TextBox_budget.Value
    
    Dim code As String
    code = get_code(��, ��, ��, ����)
    Dim c As Range
    If code <> "" Then
        With ws.Range("���׸��ڵ巹�̺�").CurrentRegion.columns(1)
            Set c = .Find(code)
            If Not c Is Nothing Then
                If c.Offset(0, 5).Value <> ����� Then
                    c.Offset(0, 5).Value = �����
                End If
            End If
        End With
    Else
        With r������ġ
            
            .Offset(0, 1).Value = ��
            .Offset(0, 2).Value = ��
            .Offset(0, 3).Value = ��
            .Offset(0, 4).Value = ����
            .Offset(0, 5).Value = �����

            Call ��������_��Ʈ����

            code_changed = True
        End With
    End If
    
End Sub

Sub ��_�ʱ�ȭ(�� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim ������ As Range, ������ As Range
    Dim vlist() As Variant
    Dim �׼� As Integer, x As Integer
    x = 0

    Dim �����׸� As String
    �����׸� = ""
    
    ListBox_��.Clear
    
    With ws.Range("���ʵ�")
        Set ������ = .Offset(3)
        If ������.Value = "" Then
            Set ������ = ������
            �׼� = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
                �׼� = ������.Row - ����ټ�
            End If
        End If
            
        If �׼� > 0 Then
            For Each ���׸� In ws.Range(������, ������)
                If ���׸�.Value <> "" Then
                    If ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                        ReDim Preserve vlist(2, x)
                        vlist(0, x) = ���׸�.Row
                        vlist(1, x) = ���׸�.Value
                        �����׸� = ���׸�.Value
                        x = x + 1
                    End If
                End If
            Next ���׸�
            ListBox_��.Column = vlist
        End If
    End With
    
    If x = 0 Then
        CommandButton_�׻���.Enabled = True
    End If
    
End Sub

Sub ��_�ʱ�ȭ(�� As String, �� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim ������ As Range
    Dim ������ As Range
    Dim vlist() As Variant
    Dim ��� As Integer
    Dim �� As String
    Dim x As Integer
    x = 0
    Dim �����׸� As String
    �����׸� = ""
        
    ListBox_��.Clear
    
    With ws.Range("���ʵ�")
        Set ������ = .Offset(3)
        If ������.Value = "" Then
            Set ������ = ������
            ��� = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
                ��� = ������.Row - ����ټ�
            End If
        End If
            
        If ��� > 0 Then
            For Each ���׸� In ws.Range(������, ������)
                With ���׸�
                    �� = .Value
                    If �� <> "" Then
                        If .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value <> �����׸� Then
                            ReDim Preserve vlist(3, x)
                            vlist(0, x) = .Row
                            vlist(1, x) = ��
                            vlist(2, x) = .Offset(0, 3).Value
                            �����׸� = ��
                            x = x + 1
                        End If
                    End If
                End With
            Next ���׸�
            
            If x > 0 Then
                ListBox_��.Column = vlist
            Else
                Call ��_�ʱ�ȭ(ComboBox_guan.Value)
            End If
        End If
    End With
End Sub

Sub ����_�ʱ�ȭ(�� As String, �� As String, �� As String)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim ������ As Range
    Dim ������ As Range
    Dim vlist() As Variant
    Dim ����� As Integer
    Dim ���� As String
    Dim x As Integer
    x = 0
    Dim �����׸� As String
    �����׸� = ""
        
    ListBox_����.Clear

    With ws.Range("�����ʵ�")
        Set ������ = .Offset(3)
        If ������.Value = "" Then
            Set ������ = ������
            ����� = 0
        Else
            If .Offset(4, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
                ����� = ������.Row - ����ټ�
            Else
                Set ������ = ������
                ����� = 1
            End If
        End If
            
        If ����� > 0 Then

            For Each ���׸� In ws.Range(������, ������)
                With ���׸�
                    ���� = .Value
                    If ���� <> "" Then
                        If .Offset(, -3).Value = �� And .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value <> �����׸� Then
                            ReDim Preserve vlist(3, x)
                            vlist(0, x) = .Row
                            vlist(1, x) = ����
                            vlist(2, x) = .Offset(0, 2).Value '����/�����/���� ����
                            �����׸� = ����
                            x = x + 1
                        End If
                    End If
                End With
            Next ���׸�
            
            If x > 0 Then
                ListBox_����.Column = vlist
            Else
                Call ��_�ʱ�ȭ(ComboBox_guan.Value, TextBox_��.Value)
                CommandButton_����_�ű�.Enabled = False ' 2016.11.9
            End If
            
        End If
    End With
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "'��'�� �������ּ���. �׿� ���� '��'�� �����ʿ� ǥ�õ˴ϴ�"
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "����/������ �������ּ���. �׿� ���� '��'���� ǥ�õ˴ϴ�"
End Sub

Private Sub ListBox_��_Click()
    Dim i���� As Integer
    Dim �� As String
    
    With ListBox_��
        i���� = .ListIndex
        If i���� > -1 Then
            �� = .List(i����, 1)

            TextBox_��_���ȣ.Value = .List(i����, 0)
            TextBox_��.Value = ��

            CommandButton_�����.Enabled = True
            Call ����_�ʱ�ȭ(ComboBox_guan.Value, TextBox_��.Value, ��)
        End If
    End With
    
    TextBox_��.Enabled = True

    CommandButton_��_����.Enabled = True
    CommandButton_��_����.Enabled = False
    CommandButton_����_����.Enabled = False
    CommandButton_����_�ű�.Enabled = True
    CommandButton_�������.Enabled = False
End Sub

Private Sub ListBox_��_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_��_����.Enabled = True
    CommandButton_��_����.Enabled = False
    CommandButton_����_����.Enabled = False
    CommandButton_�������.Enabled = False
End Sub

Private Sub ListBox_����_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_����_����.Enabled = True
    CommandButton_��_����.Enabled = False
    CommandButton_��_����.Enabled = False
    CommandButton_�������.Enabled = True
End Sub

Private Sub ListBox_��_Click()
    Dim i���� As Integer
    Dim �� As String
    
    With ListBox_��
        i���� = .ListIndex
        If i���� > -1 Then
            �� = .List(i����, 1)
            Call ��_�ʱ�ȭ(ComboBox_guan.Value, ��)
            ListBox_����.Clear
            TextBox_��_���ȣ.Value = .List(i����, 0)
            TextBox_��.Value = ��
        End If
    End With
    TextBox_��.Enabled = True

    CommandButton_��_����.Enabled = True
    CommandButton_��_����.Enabled = False
    CommandButton_����_����.Enabled = False
    
    CommandButton_��_�ű�.Enabled = True
    CommandButton_�������.Enabled = False
End Sub

Private Sub ListBox_����_Click()
    Dim i���� As Integer
    Dim ���� As String
    Dim ���񼳸� As String
    
    With ListBox_����
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 1)
            ���񼳸� = .List(i����, 2)
            TextBox_����_���ȣ.Value = .List(i����, 0)
            TextBox_����.Value = ����

            CommandButton_�������.Enabled = True
        End If
    End With
    
    TextBox_����.Enabled = True

    CommandButton_��_����.Enabled = False
    CommandButton_��_����.Enabled = False
    CommandButton_�������.Enabled = True
End Sub

Private Sub ListBox_��_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CommandButton_��_����.Enabled = True
    CommandButton_��_����.Enabled = False
    CommandButton_����_����.Enabled = False
    CommandButton_�������.Enabled = False
End Sub

Private Sub TextBox_��_Change()
    If TextBox_��.Value <> "" Then
        CommandButton_��_����.Enabled = True
    End If
End Sub

Private Sub TextBox_����_Change()
    If TextBox_��.Value <> "" Then
        CommandButton_����_����.Enabled = True
    End If
End Sub

Private Sub TextBox_��_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "'��'�� �̸��� �ٲٰų� �߰��Ϸ��� �̰��� �Է��� '����'��ư�� �����ּ���"
End Sub

Private Sub UserForm_Initialize()
    code_changed = False
    
    Call ��_�ʱ�ȭ
    
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim ���� As Integer
    ���� = ws.Range("���ʵ�").End(xlDown).Row
    TextBox_��_���ȣ.Value = ���� + 1
    TextBox_��_���ȣ.Value = ���� + 1
    
    With ListBox_��
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_��
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    With ListBox_����
        .columnCount = 2
        .ColumnWidths = "0cm;5cm"
    End With
    
    CommandButton_��_�ű�.Enabled = False
    CommandButton_��_����.Enabled = False
    CommandButton_�׻���.Enabled = False
    CommandButton_��_�ű�.Enabled = False
    CommandButton_��_����.Enabled = False
    CommandButton_�����.Enabled = False
    CommandButton_����_�ű�.Enabled = False
    CommandButton_����_����.Enabled = False
    CommandButton_�������.Enabled = False

End Sub

Sub ��_�ʱ�ȭ()
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    �����׸� = ""
    ���� = ws.Range("b2").CurrentRegion.Rows.Count
    ��� = ws.Range("d4").CurrentRegion.Rows.Count
    Dim �� As String
    
    ComboBox_guan.Clear
    
    For Each ���׸� In ws.Range("b2", "b" & ����)
        �� = ���׸�.Value
        If �� <> "" Then
            If �� <> �����׸� Then
                If �� <> "����ܼ���" And �� <> "���������" Then
                    ComboBox_guan.AddItem ��
                    �����׸� = ��
                End If
            End If
        End If
    
    Next ���׸�
    
End Sub
