VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_���� 
   Caption         =   "���꼳��"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "UserForm_����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const �����ͽ�Ʈ As String = "���꼭"
Const ����ټ� As Integer = 1
Dim error_num As Integer
Dim data_changed As Integer

Private Sub ComboBox_��_Change()
    Call ��_�ʱ�ȭ(ComboBox_��.Value)
    ComboBox_��.SetFocus
End Sub

Private Sub ComboBox_����_Change()
    Dim ���׸� As Range
    Dim �����׸� As String, �� As String, �� As String, �� As String, ���� As String
    Dim ��� As Integer
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)

    Dim ����� As Integer
    ����� = ws.Range("e4").CurrentRegion.Rows.Count
    �����׸� = ""
    �� = ComboBox_��.Value
    �� = ComboBox_��.Value
    �� = ComboBox_��.Value
    ���� = ComboBox_����.Value
        
    For Each ���׸� In ws.Range("e4", "e" & �����)
        With ���׸�
            If .Value <> "" Then
                If .Offset(, -3).Value = �� And .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value = ���� Then
                    TextBox_���ȣ.Value = .Row
                    
                    Exit For
                End If
            End If
        End With
    Next ���׸�
    
    TextBox_�����.SetFocus
End Sub

Private Sub ComboBox_��_Change()
    Call ��_�ʱ�ȭ(ComboBox_��.Value, ComboBox_��.Value)
    ComboBox_��.SetFocus
End Sub

Private Sub ComboBox_��_Change()
    Call ����_�ʱ�ȭ(ComboBox_��.Value, ComboBox_��.Value, ComboBox_��.Value)
    ComboBox_����.SetFocus
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim ���ȣ As Integer
        
    If ComboBox_��.Value = "" Then
        MsgBox "'��'�� �������ֽʽÿ�"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_��.Value = "" Then
        MsgBox "'��'�� �������ֽʽÿ�"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_��.Value = "" Then
        MsgBox "'��'�� �������ֽʽÿ�"
        ComboBox_��.SetFocus
        Exit Sub
    End If
    
    If ComboBox_����.Value = "" Then
        MsgBox "'����'�� �������ֽʽÿ�"
        ComboBox_����.SetFocus
        Exit Sub
    End If
    
    If TextBox_�����.Value = "" Or Not IsNumeric(TextBox_�����.Value) Then
        MsgBox "������� ���ڷ� �Է����ֽʽÿ�"
        TextBox_�����.SetFocus
        Exit Sub
    End If
    
    Set ws = Worksheets(�����ͽ�Ʈ)
    ���ȣ = TextBox_���ȣ.Value
    
    If ���ȣ > 0 Then
        With ws.Range("������ʵ�")
            .Offset(���ȣ - ����ټ�).Value = TextBox_�����.Value
        End With
        '���� �Է� ���� �ʵ� �ʱ�ȭ
        MsgBox "�Էµƽ��ϴ�"
        error_num = 0
        data_changed = 1
        Call �ʱ�ȭ
    Else
        error_num = 1
        MsgBox "���忡 �����߽��ϴ�"
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me

    If error_num = 0 And data_changed = 1 Then
        Call ��꼭�ʱ�ȭ
    End If
    Ȩ
End Sub

Private Sub TextBox_�����_Change()
    TextBox_�����.Value = format(TextBox_�����.Value, "#,#")
End Sub

Private Sub UserForm_Initialize()
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)
    Dim �����׸� As String
    �����׸� = ""
    Dim ���� As Integer
    Dim ��� As Integer
    Dim �� As String
    
    ���� = ws.Range("b2").CurrentRegion.Rows.Count
    ��� = ws.Range("d4").CurrentRegion.Rows.Count

    For Each ���׸� In ws.Range("b2", "b" & ����)
        �� = ���׸�.Value
        
        If ���׸�.Value <> "" Then
            If �� <> "����ܼ���" And �� <> "���������" Then
                If �� <> �����׸� Then
                    ComboBox_��.AddItem ��
                    �����׸� = ��
                End If
            End If
        End If
        
    Next ���׸�
    
    error_num = 0
    data_changed = 0

End Sub

Sub ��_�ʱ�ȭ(�� As String, Optional ByRef frm As UserForm)
    Dim ���׸� As Range
    Dim ws As Worksheet
    Set ws = Worksheets(�����ͽ�Ʈ)

    Dim �׼� As Integer
    �׼� = ws.Range("c4").CurrentRegion.Rows.Count
    Dim �����׸� As String
    �����׸� = ""
    Dim �� As UserForm
        
    If Not frm Is Nothing Then
        Set �� = frm
    Else
        Set �� = UserForm_����
    End If
        
    ��.ComboBox_��.Clear
    
    For Each ���׸� In ws.Range("c4", "c" & �׼�)
        If ���׸�.Value <> "" Then
            If ���׸�.Offset(, -1).Value = �� And ���׸�.Value <> �����׸� Then
                ��.ComboBox_��.AddItem ���׸�.Value
                �����׸� = ���׸�.Value
            End If
        End If
    Next ���׸�

End Sub

Sub ��_�ʱ�ȭ(�� As String, �� As String, Optional ByRef frm As UserForm)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets(�����ͽ�Ʈ)

        Dim ��� As Integer
        Dim �����׸� As String
        ��� = ws.Range("d4").CurrentRegion.Rows.Count
        �����׸� = ""
        Dim �� As UserForm
        
        If Not frm Is Nothing Then
            Set �� = frm
        Else
            Set �� = UserForm_����
        End If
        
        ��.ComboBox_��.Clear
        
        For Each ���׸� In ws.Range("d4", "d" & ���)
            With ���׸�
                If .Value <> "" Then
                    If .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And .Value <> �����׸� Then
                        ��.ComboBox_��.AddItem .Value
                        
                        �����׸� = .Value
                    End If
                End If
            End With
        Next ���׸�
    
End Sub

Sub ����_�ʱ�ȭ(�� As String, �� As String, �� As String, Optional ByRef frm As UserForm)
        Dim ���׸� As Range
        Dim ws As Worksheet
        Set ws = Worksheets(�����ͽ�Ʈ)

        Dim ����� As Integer
        Dim �����׸� As String
        ����� = ws.Range("d4").CurrentRegion.Rows.Count
        �����׸� = ""
        Dim ���� As String
        Dim �� As UserForm
        
        If Not frm Is Nothing Then
            Set �� = frm
        Else
            Set �� = UserForm_����
        End If
        
        ��.ComboBox_����.Clear
        
        For Each ���׸� In ws.Range("e4", "e" & �����)
            With ���׸�
                ���� = .Value
            
                If ���� <> "" Then
                    If .Offset(, -3).Value = �� And .Offset(, -2).Value = �� And .Offset(, -1).Value = �� And ���� <> �����׸� Then
                        ��.ComboBox_����.AddItem ����
                        
                        �����׸� = ����
                    End If
                End If
            End With
        Next ���׸�
    
End Sub

Sub �ʱ�ȭ()
    Dim ��Ʈ�� As Control
    For Each ��Ʈ�� In UserForm_����ݳ���.Controls
        If TypeOf ��Ʈ�� Is MSForms.TextBox Then ��Ʈ��.Value = ""
        If TypeOf ��Ʈ�� Is MSForms.combobox Then ��Ʈ��.Value = ""
    Next
End Sub
