VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_���� 
   Caption         =   "ȯ�漳��"
   ClientHeight    =   6100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   OleObjectBlob   =   "UserForm_����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const �ִ���� As Integer = 10000
Const �ִ����2 As Integer = 5000
'Const PWD = "1234"

'����ڰ� ���� ��Ʈ�� �����⺸�� �غ�� �� ���� �̿��ϵ��� ����
Sub ��Ʈ��ȣ(��Ʈ�̸� As String)
    With Worksheets(��Ʈ�̸�)
        .Protect 'PWD '��Ʈ��ȣ
    End With
   
    MsgBox "���ÿ����� ��� �۾��� �Ϸ�Ǿ����ϴ�.", vbExclamation, ""
End Sub

Private Sub CheckBox_��Ʈ���_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "���� ��뿡 �ͼ��ϰų� ����� �����ϰ��� ��쿡�� �����ϼ���"
End Sub

'���� ������Ʈ ���� ��� ���� ��.
'VBA �ڵ常 ������Ʈ �ǰ� ��Ʈ���� ������ ���� ��� ���� ���ϰ� ���� ��ġ�� ������Ʈ ������ �ΰ� �� ��ư Ŭ���ϰ�
Private Sub CommandButton_update_Click()
    Dim Filt$, title$, fileName$, Message As VbMsgBoxResult
    Filt = "VB Files (*.bas; *.frm; *.cls)(*.bas; *.frm; *.cls)," & _
        "*.bas;*.frm;*.cls"
    Dim vbp As Object 'VBIDE.VBProject
    Dim vbc As Object 'VBIDE.VBComponent
    
    Set vbp = ActiveWorkbook.VBProject
    
    fileName = Application.GetOpenFilename(FileFilter:=Filt, _
        FilterIndex:=5, title:=title)
        
    If fileName <> vbNullString Then
        Dim n As Integer
        n = Len(fileName)
        Dim m As String
        m = Left(fileName, n - 4)
        Dim m_array() As String
        m_array() = Split(m, "\")
        m = m_array(UBound(m_array))
            
        On Error Resume Next
        Set vbc = vbp.VBComponents(m)
        If Err = 0 Then
            vbp.VBComponents.Remove vbc
        Else
            Err.Clear
        End If
        vbp.VBComponents.Import fileName
        MsgBox "������Ʈ �߽��ϴ�"
        On Error GoTo 0
    End If

End Sub

Function module_name(fileName As String)
    Dim n As Integer
    n = Len(fileName)
    Dim m As String
    m = Left(fileName, n - 4)
    Dim m_array() As String
    m_array() = Split(m, "\")
    module_name = m_array(UBound(m_array))
End Function

Private Sub CommandButton_updateAll_Click()

    Dim vbp As Object 'VBIDE.VBProject
    Dim vbc As Object 'VBIDE.VBComponent
    
    Set vbp = ActiveWorkbook.VBProject
    
    Dim sThisFilePath          As String
    Dim sFile                  As String
    
    sThisFilePath = ThisWorkbook.Path
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    
    sFile = Dir(sThisFilePath & "*.bas")
    Dim m As String
    
    Do While sFile <> vbNullString
        'MsgBox "The next file is " & sFile
        On Error Resume Next
        m = module_name(sFile)
        Set vbc = vbp.VBComponents(m)
        If Err = 0 Then
            vbp.VBComponents.Remove vbc
        Else
            Err.Clear
        End If
        vbp.VBComponents.Import sFile
        
        On Error GoTo 0
        sFile = Dir
    Loop
    MsgBox "������Ʈ �߽��ϴ�"
End Sub

Private Sub CommandButton_��������������_Click()
    ' #1 �� ��Ʈ ��
    Dim ws As Worksheet
    Set ws = Worksheets("��������2")
    ws.Visible = xlSheetVisible
    ws.Activate
    Unload Me

    ' #2 ������ �����϶� �ȳ� �޽���
   
    ' #3 ��Ʈ�� �����Ͱ� �����, ����, ���׸�, ����, ����/����, ��/�� ���� �� ù �ٿ� ������� �޽���
    ' #4 ù �࿡ ��� �� �ִ��� Ȯ��
    ' #5 �� �� ��ȸ�ϸ� ���� & ȸ����忡 ���̱�
    ' #6 �߰��� �κ� ��ȸ�ϸ� ���׸� ���� : ������ ������ ǥ��, ������ ������ ��ϵ� �ڵ����� Ȯ��
    ' #6-1. �ű� ���׸��� ��� ������ ���׸� ����
    ' #7 ȸ����� ����
End Sub

Private Sub CommandButton_���������ʱ�ȭ_Click()
    Call ȸ�輳���ʱ�ȭ("��������")
End Sub

Private Sub CommandButton_�μ��ʱ�ȭ_Click()
    Call ȸ�輳���ʱ�ȭ("�μ�")
End Sub

Private Sub CommandButton_����_Click()
    Dim orgName As String
    Dim accountStartDate As String
    Dim startDate As String
    Dim endDate As String
    Dim ��������� As String
    Dim ����1���� As String
    Dim ����2���� As String
    Dim ����3���� As String
    Dim ��Ʈ��� As String
    
    orgName = TextBox_�����.Value
    accountStartDate = TextBox_ȸ�������.Value
    ��������� = TextBox_���������.Value
    ����1���� = TextBox_����1����.Value
    ����2���� = TextBox_����2����.Value
    ����3���� = TextBox_����3����.Value

    ��Ʈ��� = CheckBox_��Ʈ���.Value
    
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
        If ����3���� <> "" Then
            .Range("H2").Value = ����3����
            '.Pictures("Picture 2").Formula = "$L$26:$O$27"
        Else
            '.Pictures("Picture 2").Formula = "$L$26:$N$27"
        End If
    End With
    
    With Worksheets("����")
        .Range("��Ʈ��ݼ���").Offset(0, 1).Value = ��Ʈ���
        If orgName <> "" And orgName <> .Range("�������").Offset(0, 1).Value Then
            .Range("�������").Offset(0, 1).Value = orgName
        End If
        
        If ��������� <> "" Then
            .Range("��������Լ���").Offset(0, 1).Value = ���������
        End If
        
        If accountStartDate <> "" And accountStartDate <> .Range("ȸ������ϼ���").Offset(0, 1).Value Then
            .Range("ȸ������ϼ���").Offset(0, 1).Value = accountStartDate
        End If
        
        .Range("�۾������ϼ���").Offset(0, 1).Value = startDate
        .Range("�۾������ϼ���").Offset(0, 1).Value = endDate
        
        If ����1���� <> "" Then
            .Range("����1����").Offset(0, 1).Value = ����1����
        End If
        
        If ����2���� <> "" Then
            .Range("����2����").Offset(0, 1).Value = ����2����
        End If
        
        'If ����3���� <> "" Then
            .Range("����3����").Offset(0, 1).Value = ����3����
        'End If

    End With
    
    With Worksheets("ȸ�����")
        .Unprotect PWD
        If orgName <> "" And orgName <> .Range("�����").Value Then
            .Range("�����").Value = orgName
        End If

        If Worksheets("����").Range("��Ʈ��ݼ���").Offset(, 1).Value = True Then
            .Protect PWD
        End If
    End With
    
    MsgBox "�����Ǿ����ϴ�"
    Unload Me
    Ȩ
End Sub

Private Sub CommandButton_��������ʱ�ȭ_Click()
    Call ȸ���ڷ��ʱ�ȭ("�������")
End Sub

Private Sub CommandButton_���_Click()
    Unload Me
    Ȩ
End Sub

Private Sub CommandButton_������Ʈ�ʱ�ȭ_Click()
    Call ȸ�輳���ʱ�ȭ("������Ʈ")
End Sub

Private Sub CommandButton_ȸ�輳����ü�ʱ�ȭ_Click()
    Call ȸ�輳���ʱ�ȭ("��ü")
End Sub

Private Sub CommandButton_ȸ����尡������_Click()
    ' #1 �� ��Ʈ ��
    Dim ws As Worksheet
    Set ws = Worksheets("��������")
    ws.Visible = xlSheetVisible
    ws.Activate
    Unload Me

    ' #2 ������ �����϶� �ȳ� �޽���
   
    ' #3 ��Ʈ�� �����Ͱ� �����, ����, ���׸�, ����, ����/����, ��/�� ���� �� ù �ٿ� ������� �޽���
    ' #4 ù �࿡ ��� �� �ִ��� Ȯ��
    ' #5 �� �� ��ȸ�ϸ� ���� & ȸ����忡 ���̱�
    ' #6 �߰��� �κ� ��ȸ�ϸ� ���׸� ���� : ������ ������ ǥ��, ������ ������ ��ϵ� �ڵ����� Ȯ��
    ' #6-1. �ű� ���׸��� ��� ������ ���׸� ����
    ' #7 ȸ����� ����
End Sub

Private Sub CommandButton_ȸ������ʱ�ȭ_Click()
    Call ȸ���ڷ��ʱ�ȭ("ȸ�����")
End Sub

Sub ȸ���ڷ��ʱ�ȭ(���� As String)
    Dim �ڵ� As String
    
    If ���� = "ȸ�����" Or ���� = "��ü" Then
        Call ��Ʈ�������("ȸ�����")
        Dim �̿��� As String
        �̿��� = InputBox("�̿����� ������ �ݾ��� �����ּ���. ������ 0 Ȥ�� �׳� Ȯ���� �����ּ���")
        With Worksheets("ȸ�����")
            .Range("A6:O7").ClearContents
            .Range("A8:O" & �ִ����).ClearContents
            .Range("a6").Value = Worksheets("����").Range("ȸ������ϼ���").Offset(0, 1).Value
            .Range("a7").Value = .Range("a6").Value
            .Range("c6").NumberFormat = "@"
            .Range("c7").NumberFormat = "@"
            .Range("c6").Value = get_code("����", "�̿���", "�̿���", "�����̿�")
            .Range("b6").Value = .Range("c6").Value & "/����/�̿���/�̿���/�����̿�"
            .Range("c7").Value = "00010101"
            .Range("b7").Value = "00010101/����ܼ���/"
            .Range("d6").Value = "����"
            .Range("e6").Value = "�̿���"
            .Range("f6").Value = "�̿���"
            .Range("g6").Value = "�����̿�"
            .Range("d7").Value = "����ܼ���"
            .Range("H6").Value = "�����̿�"
            .Range("i6").Value = �̿���
            .Range("i7").Value = �̿���
            .Range("H7").Value = "�����Ա�"
            .Range("K6:O7").ClearContents
            .Range("K6").Value = 1
            .Range("K7").Value = 0
        End With
        Call ��Ʈ���("ȸ�����")
        MsgBox "ȸ������� �ʱ�ȭ�Ǿ����ϴ�"
    End If
    
    If ���� = "�������" Or ���� = "��ü" Then
        With Worksheets("������Ǵ���")
            .Range("A4:I" & �ִ����2).ClearContents
        End With
        MsgBox "������Ǵ����� �ʱ�ȭ�Ǿ����ϴ�"
    End If

End Sub

Sub ȸ�輳���ʱ�ȭ(���� As String)
    If ���� = "������Ʈ" Or ���� = "��ü" Then
        With Worksheets("����").Range("������Ʈ�������̺�")
            Range(.Offset(1, 0), .End(xlDown)).ClearContents
        End With
        MsgBox "������Ʈ ������ �ʱ�ȭ�Ǿ����ϴ�"
    End If
    
    If ���� = "�μ�" Or ���� = "��ü" Then
        With Worksheets("����").Range("�μ��������̺�")
            Range(.Offset(1, 0), .End(xlDown)).ClearContents
        End With
        MsgBox "�μ� ������ �ʱ�ȭ�Ǿ����ϴ�"
    End If
    
    If ���� = "��������" Or ���� = "��ü" Then
        With Worksheets("���꼭")
            .Range("A4:G1000").ClearContents
            With .Range("b4")
                .Value = "����"
                .Offset(, 1).Value = "�̿���"
                .Offset(, 2).Value = "�̿���"
                .Offset(, 3).Value = "�����̿�"
            End With
            
            With .Range("b5")
                .Value = "����"
                .Offset(, 1).Value = "�����"
                .Offset(, 2).Value = "�����"
                .Offset(, 3).Value = "�����"
            End With
            
        End With
        Call ���꼭�ڵ��Է�
        MsgBox "�������� ������ �ʱ�ȭ�Ǿ����ϴ�"
    End If
        
    If ���� = "����" Or ���� = "��ü" Then
        With Worksheets("���꼭")
            .Range("F4", .Range("F4").End(xlDown)).ClearContents
        End With
        MsgBox "���� ������ �ʱ�ȭ�Ǿ����ϴ�"
    End If
    
End Sub

Sub �⺻�����ʱ�ȭ()
    With Worksheets("����")
        .Range("E2").Value = "" '����1
        .Range("F2").Value = "" '����2
        .Range("G2").Value = "" '����3
        .Range("H2").Value = "" '����4
        
        .Range("��Ʈ��ݼ���").Offset(0, 1).Value = ""
        .Range("�������").Offset(0, 1).Value = ""
        .Range("��������Լ���").Offset(0, 1).Value = ""
        .Range("ȸ������ϼ���").Offset(0, 1).Value = ""
    End With
End Sub

Private Sub CommandButton_ȸ���ڷ���ü�ʱ�ȭ_Click()
    '��ü �ʱ�ȭ�� �� ��쿡�� ȸ������� ��¥�� ���� ù�ط� ����
    Call init_firstday
    
    Call ȸ���ڷ��ʱ�ȭ("��ü")
End Sub


Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "��ü/�������� �Է����ּ���. ���� ���� �� ��� ��½� ǥ�õ˴ϴ�"
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If MultiPage1.Value = 0 Then
        Label_�޽���.caption = "�ʼ������� �����ؾ� �� ������Դϴ�"
    ElseIf MultiPage1.Value = 1 Then
        Label_�޽���.caption = "�ΰ������� �����ϴ� ������Դϴ�"
    Else
        Label_�޽���.caption = "�Է��� ȸ���ڷḦ ��� ���� ó������ �ǵ����ϴ�"
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim accountStartDate As String
    'Dim mytoday As String
    'Dim days As Integer
    
    With Worksheets("����")
        TextBox_�����.Value = .Range("�������").Offset(0, 1).Value
        If Not .Range("ȸ������ϼ���").Offset(0, 1).Value = "" Then
            'mytoday = Date
            accountStartDate = .Range("ȸ������ϼ���").Offset(0, 1).Value
            'days = DateDiff("d", mytoday, ȸ�������)
            'If days > 365 Then
            '    ȸ������� = Year(Now) & "-01-01"
            'End If
            TextBox_ȸ�������.Value = accountStartDate
        Else
            'accountStartDate = Year(Now) & "-01-01"
            Call init_firstday
        End If
        
        TextBox_���������.Value = .Range("��������Լ���").Offset(0, 1).Value
        TextBox_����1����.Value = .Range("����1����").Offset(0, 1).Value
        TextBox_����2����.Value = .Range("����2����").Offset(0, 1).Value
        TextBox_����3����.Value = .Range("����3����").Offset(0, 1).Value

        CheckBox_��Ʈ���.Value = Worksheets("����").Range("��Ʈ��ݼ���").Offset(0, 1).Value
    End With
    MultiPage1.Value = 0  'ù������(�⺻����)�� �׻� ���� �ߵ���
End Sub
