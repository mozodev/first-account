VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ǰ�Ǽ����Ǽ����� 
   Caption         =   "ǰ�Ǽ�/���Ǽ� ����"
   ClientHeight    =   4980
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5280
   OleObjectBlob   =   "UserForm_ǰ�Ǽ����Ǽ�����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_ǰ�Ǽ����Ǽ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ��offset_��¥ As Integer = 0
Const ��offset_�ڵ� As Integer = 1
Const ��offset_����� As Integer = 2
Const ��offset_�԰� As Integer = 3
Const ��offset_���� As Integer = 4
Const ��offset_�ܰ� As Integer = 5
Const ��offset_�ݾ� As Integer = 6
Const ��offset_��� As Integer = 7
Const ��offset_�ϴܺ�� As Integer = 8
    
Sub ��¥����()
    
    Dim ws As Worksheet
    Set ws = Worksheets("ȸ�����")
    ws.Activate
    
    ������ = IIf(TextBox_������.Value <> "", TextBox_������.Value, "2014-01-01")
    ������ = IIf(TextBox_������.Value <> "", TextBox_������.Value, Date)
    
    Dim r���� As Range

    ������_�� = format(������, "yyyy")
    ������_�� = format(������, "m")
    ������_�� = format(������, "d")
    Set r���� = ws.Range("�����ʵ巹�̺�").CurrentRegion.columns(1).Find(������_�� & "/" & ������_�� & "/" & ������_��)
    
    If r���� Is Nothing Then
        MsgBox "������ ��¥(������) �ڷḦ ã�� ���߽��ϴ�. �ֱٿ� �Է��� ������ ���ϴ�"
        ws.Range("�����ʵ巹�̺�").End(xlDown).Select
    Else

        With Worksheets("����")
            .Activate
            .Range("�۾������ϼ���").Offset(0, 1).Value = ������
            .Range("�۾������ϼ���").Offset(0, 1).Value = ������
        End With
        ws.Activate
        r����.Select
    End If

End Sub

Sub ��¥����2()
    Worksheets("ǰ�Ǽ�����").Activate
    
    Dim ws As Worksheet
    Set ws = Worksheets("ǰ�Ǽ�����")

    ������ = IIf(TextBox_������.Value <> "", TextBox_������.Value, Date)
    
    Dim r���� As Range

    ������_�� = format(������, "yyyy")
    ������_�� = format(������, "m")
    ������_�� = format(������, "d")
    Set r���� = ws.Range("ǰ�ǳ�¥���̺�").CurrentRegion.columns(1).Find(������_�� & "/" & ������_�� & "/" & ������_��)
    
    If r���� Is Nothing Then
        MsgBox "������ ��¥(������) �ڷḦ ã�� ���߽��ϴ�. �ֱ� ��¥�� ǰ�Ǽ��� �����մϴ�"
        ws.Range("ǰ�ǳ�¥���̺�").End(xlDown).Select
    Else
        r����.Select
    End If

End Sub

Sub ��¥����3()
    Dim ws As Worksheet
    Set ws = Worksheets("������Ǵ���")
    ws.Activate
    Dim ������ As String
    Dim ������_�� As String
    Dim ������_�� As String
    Dim ������_�� As String
    
    ������ = IIf(TextBox_������.Value <> "", TextBox_������.Value, Date)
    
    Dim r���� As Range

    ������_�� = format(������, "yyyy")
    ������_�� = format(������, "m")
    ������_�� = format(������, "d")
    Set r���� = ws.Range("���ǳ�¥���̺�").CurrentRegion.columns(1).Find(������_�� & "/" & ������_�� & "/" & ������_��)
    
    If r���� Is Nothing Then
        MsgBox "������ ��¥(������) �ڷḦ ã�� ���߽��ϴ�. �ֱ� ��¥�� ǰ�Ǽ��� �����մϴ�"
        ws.Range("���ǳ�¥���̺�").End(xlDown).Select
    Else
        r����.Select
    End If

End Sub

'Private Sub btn_�Աݿ���_Click()
'    Call ��¥����
'    Call �Աݿ����ۼ�
'    Unload Me
'End Sub

'Private Sub btn_��ݿ���_Click()
'    Call ��¥����
'    Call ��ݿ����ۼ�
'    Unload Me
'End Sub

'Private Sub btn_ǰ�Ǽ�_Click()
'
'    Dim i���� As Integer
'    Dim ���ȣ As Integer
'
'    With ListBox2
'        i���� = .ListIndex
'        If i���� > -1 Then
'            ���ȣ = .List(i����, 0)
'            Worksheets("ǰ�Ǽ�����").Range("A" & ���ȣ).Select
'            Call ǰ�Ǽ��ۼ�(False)
'        Else
'            Call ��¥����2
'            Call ǰ�Ǽ��ۼ�(True)
'        End If
'    End With
'
'    Unload Me
'End Sub

Private Sub btn_ǰ�Ǽ�_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ ��¥�� ǰ�Ǽ��� �����մϴ�"
End Sub

Private Sub CommandButton_close2_Click()
    Unload Me
    If Parent = "ȸ�����" Then
        Worksheets("ȸ�����").Activate
    ElseIf Parent = "ǰ�Ǽ�����" Then
        Worksheets("ǰ�Ǽ�����").Activate
    ElseIf Parent = "������Ǵ���" Then
        Worksheets("������Ǵ���").Activate
    Else
        Ȩ
    End If
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

    Ű���� = TextBox_������.Value
    If Not IsNumeric(Ű����) Then '��¥ ������ �Է��� ���
        Ű���� = format(Ű����, "m") & "/" & format(Ű����, "d") & "/" & format(Ű����, "yyyy")
    End If
    
    y = 0
    
    If (Len(Ű����) > 0) Then
        Set ��ü = Worksheets("������Ǵ���").Range("���ǳ�¥���̺�").CurrentRegion.columns(1)
        Set ã����¥ = ��ü.Find(What:=Ű����, LookAt:=xlPart)
        
        If Not ã����¥ Is Nothing Then
            ù��ġ = ã����¥.Address
            
            Do
                ReDim Preserve vlist(10, x)
                Set ���ڵ� = ã����¥.Resize(1, 10)
                vlist(0, x) = ���ڵ�.Row
                vlist(1, x) = ���ڵ�.Cells(, ��offset_��¥ + 1)
                vlist(2, x) = ���ڵ�.Cells(, ��offset_�ڵ� + 1)
                vlist(3, x) = ���ڵ�.Cells(, ��offset_����� + 1)
                vlist(4, x) = ���ڵ�.Cells(, ��offset_�԰� + 1)
                vlist(5, x) = ���ڵ�.Cells(, ��offset_���� + 1)
                vlist(6, x) = ���ڵ�.Cells(, ��offset_�ܰ� + 1)
                vlist(7, x) = ���ڵ�.Cells(, ��offset_�ݾ� + 1)
                vlist(8, x) = ���ڵ�.Cells(, ��offset_��� + 1)
                vlist(9, x) = ���ڵ�.Cells(, ��offset_�ϴܺ�� + 1)
                
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

Private Sub CommandButton2_Click()
    Unload Me
    If Parent = "ȸ�����" Then
        Worksheets("ȸ�����").Activate
    ElseIf Parent = "ǰ�Ǽ�����" Then
        Worksheets("ǰ�Ǽ�����").Activate
    ElseIf Parent = "������Ǵ���" Then
        Worksheets("������Ǵ���").Activate
    Else
        Ȩ
    End If
End Sub

Private Sub CommandButton4_Click()
    Worksheets("��꼭").Activate
    Unload Me
End Sub

Private Sub CommandButton7_Click()
    Worksheets("�׸�а���").Activate
    Unload Me
End Sub

Private Sub OptionButton_thismonth_Click()
    Dim dtLastDayofMonth As Date
    dtLastDayofMonth = DateAdd("d", -1, DateSerial(Year(Now), Month(Now) + 1, 1))
    TextBox_������.Value = DateSerial(Year(Now), Month(Now), 1)
    TextBox_������.Value = dtLastDayofMonth
    
End Sub

Private Sub OptionButton_thisyear_Click()
    TextBox_������.Value = DateSerial(Year(Now), 1, 1)
    TextBox_������.Value = DateAdd("d", -1, DateSerial(Year(Now) + 1, 1, 1))
End Sub

Private Sub CommandButton3_Click()

    Dim i���� As Integer
    Dim ���ȣ As Integer
    
    With ListBox1
        i���� = .ListIndex
        If i���� > -1 Then
            ���ȣ = .List(i����, 0)
            Worksheets("������Ǵ���").Range("A" & ���ȣ).Select
            Call ������Ǽ��ۼ�(False)
        Else
            Call ��¥����3
            Call ������Ǽ��ۼ�(True)
        End If
    End With
    
    Unload Me
End Sub

Private Sub CommandButton3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������ ��¥�� ������Ǽ��� ����մϴ�"
End Sub

Private Sub CommandButton5_Click()
    Dim ��ü As Range
    Dim ã����¥ As Range
    Dim ���ڵ� As Range
    Dim cell As Range
    Dim x As Integer, y As Integer
    Dim Ű���� As String
    Dim ù��ġ As String
    Dim vlist() As Variant

    Ű���� = TextBox_ǰ�ǳ�¥.Value
    If Not IsNumeric(Ű����) Then '��¥ ������ �Է��� ���
        Ű���� = format(Ű����, "m") & "/" & format(Ű����, "d") & "/" & format(Ű����, "yyyy")
    End If
    
    y = 0
    
    If (Len(Ű����) > 0) Then
        Set ��ü = Worksheets("ǰ�Ǽ�����").Range("ǰ�ǳ�¥���̺�").CurrentRegion.columns(1)
        Set ã����¥ = ��ü.Find(What:=Ű����, LookAt:=xlPart)
        
        If Not ã����¥ Is Nothing Then
            ù��ġ = ã����¥.Address
            
            Do
                ReDim Preserve vlist(10, x)
                Set ���ڵ� = ã����¥.Resize(1, 10)
                vlist(0, x) = ���ڵ�.Row
                vlist(1, x) = ���ڵ�.Cells(, ��offset_��¥ + 1)
                vlist(2, x) = ���ڵ�.Cells(, ��offset_�ڵ� + 1)
                vlist(3, x) = ���ڵ�.Cells(, ��offset_����� + 1)
                vlist(4, x) = ���ڵ�.Cells(, ��offset_�԰� + 1)
                vlist(5, x) = ���ڵ�.Cells(, ��offset_���� + 1)
                vlist(6, x) = ���ڵ�.Cells(, ��offset_�ܰ� + 1)
                vlist(7, x) = ���ڵ�.Cells(, ��offset_�ݾ� + 1)
                vlist(8, x) = ���ڵ�.Cells(, ��offset_��� + 1)
                vlist(9, x) = ���ڵ�.Cells(, ��offset_�ϴܺ�� + 1)
                
                x = x + 1
                y = 0
                
                Set ã����¥ = ��ü.FindNext(ã����¥)
            Loop While Not ã����¥ Is Nothing And ã����¥.Address <> ù��ġ
            
            ListBox2.Column = vlist
        Else
            MsgBox "�˻������ �������� �ʽ��ϴ�"
            ListBox2.Clear
        End If
    End If
End Sub

Private Sub SpinButton_������_SpinDown()
    TextBox_������.Value = DateAdd("d", -1, TextBox_������.Value)
End Sub

Private Sub SpinButton_������_SpinUp()
    TextBox_������.Value = DateAdd("d", 1, TextBox_������.Value)
End Sub

Private Sub UserForm_Initialize()

    If Parent = "ǰ�Ǽ�����" Or Parent = "ǰ�Ǽ�����_from_Ȩ" Then
        With Worksheets("ǰ�Ǽ�����")
            .Activate
            TextBox_������ = .Range("ǰ�ǳ�¥���̺�").Offset(1, 0).Value
            TextBox_ǰ�ǳ�¥ = .Range("ǰ�ǳ�¥���̺�").End(xlDown).Value
        End With
    
    ElseIf Parent = "������Ǵ���" Or Parent = "������Ǵ���_from_Ȩ" Then
        With Worksheets("������Ǵ���")
            .Activate
            TextBox_������ = .Range("���ǳ�¥���̺�").Offset(1, 0).Value
            TextBox_������ = .Range("���ǳ�¥���̺�").End(xlDown).Value
        End With
        
    Else

        With Worksheets("����")
            .Activate
            ������ = .Range("�۾������ϼ���").Offset(0, 1).Value
            If ������ = "" Then
                ������ = .Range("ȸ������ϼ���").Offset(0, 1).Value
            End If
            TextBox_������.Value = ������
            ������ = .Range("�۾������ϼ���").Offset(0, 1).Value
            If ������ = "" Then
                ������ = Date
            End If
            TextBox_������.Value = ������
            TextBox_ǰ�ǳ�¥.Value = ������
        End With
    End If
    
    With ListBox1
        .columnCount = 9
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;0cm;0cm;0cm;1cm;3cm"
    End With
    With ListBox2
        .columnCount = 9
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;0cm;0cm;0cm;1cm;3cm"
    End With
End Sub

