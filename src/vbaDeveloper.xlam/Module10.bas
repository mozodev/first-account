Attribute VB_Name = "Module10"
'module10: �޴�/�������̽�
Option Explicit
Public Parent As String

Sub Ȩ()
    Worksheets("ù������").Activate
End Sub

Sub �޴�_����()
    Worksheets("����").Activate
End Sub

Sub �޴�_��Ȳ����()
    '���� ���� ����
    UserForm_����.Show
End Sub

Sub �޴�_����()
    UserForm_front_config.Show
End Sub

Sub �޴�_ȯ�漳��()
    UserForm_����.Show
End Sub

Sub �޴�_��������()
    UserForm_�����Ͱ���.MultiPage1.Value = 0
    UserForm_�����Ͱ���.Show
End Sub

Sub �޴�_��������()
    UserForm_�����Ͱ���.MultiPage1.Value = 1
    UserForm_�����Ͱ���.Show
End Sub

Sub �޴�_���⳻���Է�()
    If ActiveSheet.name = "ȸ�����" Then
        Parent = "ȸ�����"
    Else
        Worksheets("ȸ�����").Activate
    End If
    
    UserForm_������Է�.Show
    Parent = ""
End Sub

Sub �޴�_���־��������()
    If ActiveSheet.name = "ȸ�����" Then
        Parent = "ȸ�����"
    Else
        Worksheets("ȸ�����").Activate
    End If
    UserForm_���־��������.Show
End Sub

Sub ����޴�_���⳻������()
    Dim �� As Integer
    If ActiveSheet.name = "ȸ�����" Then
        Parent = "ȸ�����"
        �� = ActiveCell.Row
        If �� < 6 Then
            �� = 6
        End If
        
        UserForm_����ݳ���.load_����ݷ��ڵ� (��)
        UserForm_����ݳ���.Show
    End If
End Sub

Sub ����޴�_������Ǽ�����()
    Dim �� As Integer, ���� As Integer
    Dim dataCount As Integer
    Dim ws As Worksheet, ws_ledger As Worksheet
    Set ws = Worksheets("������Ǵ���")
    Set ws_ledger = Worksheets("ȸ�����")
    Dim error_code As Integer, answer As Integer
    
    If ActiveSheet.name = "ȸ�����" Then
        Parent = "ȸ�����"
        error_code = 0
        
        �� = ActiveCell.Row
        If �� < 6 Then
            error_code = 1
        Else
            With ws_ledger.Range("A5")
                dataCount = .End(xlDown).Row - 5
            End With
            If �� > dataCount + 5 Then
                error_code = 1
            End If
        End If
        
        If error_code > 0 Then
            MsgBox "������Ǽ��� ������ ������ ��� �ִ� ���� �������ּ���"
            Exit Sub
        End If
        
        With ws_ledger.Range("A" & ��)
            answer = MsgBox("������Ǵ��忡 �Է��ϰ� ������Ǽ� ����� �����մϴ� : " & .Value & " / " & .Offset(, 7).Value & " ( " & .Offset(, 9).Value & " ) ", vbYesNo + vbQuestion, "������Ǽ� ���� Ȯ��")
            If answer <> vbYes Then
                Exit Sub
            End If
        End With
            
        ws.Activate
        
        '������Ǵ��忡 copy
        Call ȸ�����_��������Է�(��)
        
        '������Ǵ����� �ٹ�ȣ ��������
        If ws.Range("���ǳ�¥���̺�").Offset(1, 0).Value Then
            ���� = ws.Range("���ǳ�¥���̺�").End(xlDown).Row

            '������Ǽ� ����
            ws.Range("A" & ����).Select

            Call ������Ǽ��ۼ�(False)
        End If
        
    End If
End Sub

Sub �޴�_������Ʈ�׺μ�����()
    UserForm_������Ʈ�׺μ�.Show
End Sub

Sub �޴�_���������Է�()
    With Worksheets("���꼭")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = True
        .columns(7).Hidden = False
        .Activate
    End With
    
    UserForm_��������.Show
End Sub

Sub �޴�_�������񺸱�()
    Worksheets("���꼭").Activate
End Sub

Sub �޴�_���꼳��()
    With Worksheets("���꼭")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = False
        .columns(7).Hidden = True
        .Activate
    End With
    UserForm_����.Show
End Sub

Sub �޴�_���꼭����()
    With Worksheets("���꼭")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = False
        .columns(7).Hidden = True
        .Activate
    End With
End Sub

Sub �޴�_�������Է�()
    Worksheets("2014������").Activate
    'UserForm3.Show '�̱����̹Ƿ� �� ��Ȱ��ȭ
End Sub

Sub �޴�_������Ǽ��ۼ�()
    If ActiveSheet.name = "������Ǵ���" Then
        Parent = "������Ǵ���"
    Else
        Worksheets("������Ǵ���").Activate
    End If
    On Error GoTo error
    UserForm_�������.Show
    Parent = ""
error:
    If Err.Number <> 0 Then

        MsgBox "������ȣ : " & Err.Number & vbCr & _
        "�������� : " & Err.Description, vbCritical, "����"

    End If
End Sub

Sub �޴�_ȸ�������ȸ()
    Worksheets("ȸ�����").Activate
End Sub

Sub ����޴�_ȸ������μ�()
    Worksheets("ȸ�����").Activate
    UserForm_�����μ⼳��.Show
End Sub

Sub �޴�_��꼭��ȸ()
    If ActiveSheet.name = "ȸ�����" Then
        Parent = "ȸ�����"
    Else
        Worksheets("ȸ�����").Activate
    End If

    UserForm5.Show
    Parent = ""
End Sub

Sub �޴�_������Ǵ��庸��()
    Worksheets("������Ǵ���").Activate
End Sub

Sub �޴�_������Ǽ�����()
    
    If ActiveSheet.name = "������Ǵ���" Then
        Parent = "������Ǵ���"
    Else
        Parent = "������Ǵ���_from_Ȩ"
    End If
    Worksheets("������Ǵ���").Activate

    UserForm_ǰ�Ǽ����Ǽ�����.MultiPage1.Value = 0
    UserForm_ǰ�Ǽ����Ǽ�����.Show
    Parent = ""
End Sub

Sub �޴�_�ʱ⼳��������()
    UserForm_�ʱ⼳��������1.Show
End Sub

Sub �޴�_���嵥���Ͱ�������()
    Dim ws As Worksheet
    Set ws = Worksheets("��������")
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub
