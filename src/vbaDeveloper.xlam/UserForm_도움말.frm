VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_���� 
   Caption         =   "����"
   ClientHeight    =   6015
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   9765.001
   OleObjectBlob   =   "UserForm_����.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ���򸻰˻����() As String
Dim �˻�����ε��� As Integer
Dim �˻������ As Integer
Dim ws As Worksheet

Private Sub CommandButton_�˻�_Click()
    Set ws = Worksheets("��ɵ���")
    Dim c As Range
    Dim �˻��� As String
    Dim i As Integer
    Dim ã���ڵ� As String
    Dim ���򸻼� As Integer
    ���򸻼� = ws.Range("����ڵ巹�̺�").End(xlDown).Row

    ReDim ���򸻰˻����(2, ���򸻼�)
        
    i = 1
    �˻��� = TextBox_���򸻰˻�.Value
    �˻�����ε��� = 1
    
    With ws.Cells
        Set c = .Find(What:=�˻���)
        If Not c Is Nothing Then
            Dim firstaddress As String
            firstaddress = c.Address
            Do
                ã���ڵ� = c.End(xlToLeft).Value
                If i = 1 Then
                    ���򸻰˻����(1, i) = ã���ڵ�
                    ���򸻰˻����(2, i) = c.Row
                    i = i + 1
                ElseIf i > 1 And ���򸻰˻����(2, i - 1) <> c.Row Then
                    ���򸻰˻����(1, i) = ã���ڵ�
                    ���򸻰˻����(2, i) = c.Row
                    i = i + 1
                End If
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstaddress
        End If
    End With
    
    �˻������ = i - 1

    Label_�˻������.caption = "�˻���� : " & �˻������ & "��"
    If �˻������ > 0 Then
        
        '�˻��� ��� �� ù��° ��� ǥ��
        Dim �� As Integer
        Dim j As Integer
        �� = ���򸻰˻����(2, 1)

        With ws.Range("A" & ��)
            ��з� = .Offset(0, 1).Value
            �з� = .Offset(0, 3).Value

            Select Case ��з�
                Case "�ϻ�ȸ��"
                    MultiPage1.Value = 0
                    For j = 1 To ListBox_�ϻ�ȸ��.ListCount

                        If ListBox_�ϻ�ȸ��.List(j - 1, 1) = .Value Then
                            ListBox_�ϻ�ȸ��.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "�������"
                    MultiPage1.Value = 1
                    For j = 1 To ListBox_�������.ListCount

                        If ListBox_�������.List(j - 1, 1) = .Value Then
                            ListBox_�������.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "����ǰ��"
                    MultiPage1.Value = 2
                    For j = 1 To ListBox_����ǰ��.ListCount

                        If ListBox_����ǰ��.List(j - 1, 1) = .Value Then
                            ListBox_����ǰ��.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "����"
                    MultiPage1.Value = 3
                    For j = 1 To ListBox_����.ListCount

                        If ListBox_����.List(j - 1, 1) = .Value Then
                            ListBox_����.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "����"
                    MultiPage1.Value = 4
                    For j = 1 To ListBox_����.ListCount

                        If ListBox_����.List(j - 1, 1) = .Value Then
                            ListBox_����.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "���"
                    MultiPage1.Value = 5
                    For j = 1 To ListBox_���.ListCount

                        If ListBox_���.List(j - 1, 1) = .Value Then
                            ListBox_���.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "�ڻ�ä��"
                    MultiPage1.Value = 6
                    For j = 1 To ListBox_�ڻ�ä��.ListCount

                        If ListBox_�ڻ�ä��.List(j - 1, 1) = .Value Then
                            ListBox_�ڻ�ä��.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
            End Select
        End With

        CommandButton_����ã��.Visible = True
    Else
        CommandButton_����ã��.Visible = False
        TextBox_���򸻰˻�.SetFocus
    End If
End Sub

Private Sub CommandButton_����ã��_Click()
    �˻�����ε��� = �˻�����ε��� + 1

    If �˻�����ε��� <= �˻������ Then
        Call �˻����ǥ��(�˻�����ε���)
    Else
        MsgBox "�� �̻��� �˻������ �����ϴ�"
        TextBox_���򸻰˻�.SetFocus
    End If
End Sub

Sub �˻����ǥ��(�ε��� As Integer)
    Dim �� As Integer
    Dim j As Integer
    �� = ���򸻰˻����(2, �ε���)
    Set ws = Worksheets("��ɵ���")

    With ws.Range("A" & ��)
        ��з� = .Offset(0, 1).Value
        �з� = .Offset(0, 3).Value

        Select Case ��з�
            Case "�ϻ�ȸ��"
                MultiPage1.Value = 0
                For j = 0 To ListBox_�ϻ�ȸ��.ListCount - 1

                    If ListBox_�ϻ�ȸ��.List(j, 1) = .Value Then
                        ListBox_�ϻ�ȸ��.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "�������"
                MultiPage1.Value = 1
                For j = 0 To ListBox_�������.ListCount - 1

                    If ListBox_�������.List(j, 1) = .Value Then
                        ListBox_�������.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "����ǰ��"
                MultiPage1.Value = 2
                For j = 0 To ListBox_����ǰ��.ListCount - 1

                    If ListBox_����ǰ��.List(j, 1) = .Value Then
                        ListBox_����ǰ��.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "����"
                MultiPage1.Value = 3
                For j = 0 To ListBox_����.ListCount - 1

                    If ListBox_����.List(j, 1) = .Value Then
                        ListBox_����.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "����"
                MultiPage1.Value = 4
                For j = 0 To ListBox_����.ListCount - 1

                    If ListBox_����.List(j, 1) = .Value Then
                        ListBox_����.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "���"
                MultiPage1.Value = 5
                For j = 0 To ListBox_���.ListCount - 1

                    If ListBox_���.List(j, 1) = .Value Then
                        ListBox_���.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "�ڻ�ä��"
                MultiPage1.Value = 6
                For j = 0 To ListBox_�ڻ�ä��.ListCount - 1

                    If ListBox_�ڻ�ä��.List(j, 1) = .Value Then
                        ListBox_�ڻ�ä��.Selected(j) = True
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub ListBox_���_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_���
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_��굵��.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_����_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_����
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_��������.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_����_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_����
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_���굵��.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_�ϻ�ȸ��_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_�ϻ�ȸ��
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_�ϻ�ȸ�赵��.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_�ڻ�ä��_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_�ڻ�ä��
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_�ڻ�ä������.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_�������_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_�������
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_������ǵ���.caption = ����
        End If
    End With
End Sub

Private Sub ListBox_����ǰ��_Click()
    Dim i���� As Integer
    Dim ���� As String
    
    With ListBox_����ǰ��
        i���� = .ListIndex
        If i���� > -1 Then
            ���� = .List(i����, 3)

            Label_����ǰ�ǵ���.caption = ����
        End If
    End With
End Sub

Private Sub MultiPage1_click(ByVal Index As Long)
    Select Case MultiPage1.SelectedItem.name
        Case "page_�ϻ�ȸ�����":  '�ϻ�ȸ�����
            Call listbox_�ʱ�ȭ("�ϻ�ȸ��")
        Case "page_�������":  '�������
            Call listbox_�ʱ�ȭ("�������")
        Case "page_����ǰ��":  '����ǰ��
            Call listbox_�ʱ�ȭ("����ǰ��")
        Case "page_����":  '����
            Call listbox_�ʱ�ȭ("����")
        Case "page_����":  '����
            Call listbox_�ʱ�ȭ("����")
        Case "page_���":  '���
            Call listbox_�ʱ�ȭ("���")
        Case Else '�ڻ�/ä������
            Call listbox_�ʱ�ȭ("�ڻ�ä��")
    End Select
End Sub

Private Sub UserForm_Initialize()
    
    With ListBox_�ϻ�ȸ��
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_�������
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_����ǰ��
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_����
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_����
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_���
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_�ڻ�ä��
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    
    Call listbox_�ʱ�ȭ("�ϻ�ȸ��")
    Call listbox_�ʱ�ȭ("�������")
    Call listbox_�ʱ�ȭ("����ǰ��")
    Call listbox_�ʱ�ȭ("����")
    Call listbox_�ʱ�ȭ("����")
    Call listbox_�ʱ�ȭ("���")
    Call listbox_�ʱ�ȭ("�ڻ�ä��")
    
    MultiPage1.Value = 0  'ù������(�ϻ�ȸ�����)�� �׻� ���� �ߵ���
End Sub

Sub listbox_�ʱ�ȭ(�����׸� As String)
    
    Dim ��Ȳ As Range
    Dim ws As Worksheet
    Set ws = Worksheets("��ɵ���")
    Dim vlist() As Variant
    Dim ���򸻼� As Integer
    Dim x As Integer
    x = 0
    Const ����ټ� As Integer = 1
        
    Select Case �����׸�
        Case "�ϻ�ȸ��"
            ListBox_�ϻ�ȸ��.Clear
        Case "�������"
            ListBox_�������.Clear
        Case "����ǰ��"
            ListBox_����ǰ��.Clear
        Case "����"
            ListBox_����.Clear
        Case "����"
            ListBox_����.Clear
        Case "���"
            ListBox_���.Clear
        Case "�ڻ�ä��"
            ListBox_�ڻ�ä��.Clear
    End Select
    
    With ws.Range("����ڵ巹�̺�")
        Set ��Ȳ = .Offset(1)
        With ��Ȳ
            If .Value = "" Then
                ���򸻼� = 0
            Else
                If .Offset(1, 0).Value <> "" Then
                    ���򸻼� = .End(xlDown).Row - ����ټ�
                End If
            End If
        End With
            
        If ���򸻼� > 0 Then
            
            For i = 0 To ���򸻼� - 1
                With ��Ȳ.Offset(i, 0)
            
                    If .Value <> "" Then
                        If .Offset(0, 1).Value = �����׸� Then
                            ReDim Preserve vlist(3, x)
                            vlist(1, x) = .Value
                            vlist(2, x) = .Offset(0, 3).Value
                            vlist(3, x) = .Offset(0, 4).Value
                            x = x + 1
                        End If
                    End If
                End With
            Next i
            
            If x > 0 Then
                Select Case �����׸�
                    Case "�ϻ�ȸ��"
                        ListBox_�ϻ�ȸ��.Column = vlist
                    Case "�������"
                        ListBox_�������.Column = vlist
                    Case "����ǰ��"
                        ListBox_����ǰ��.Column = vlist
                    Case "����"
                        ListBox_����.Column = vlist
                    Case "����"
                        ListBox_����.Column = vlist
                    Case "���"
                        ListBox_���.Column = vlist
                    Case "�ڻ�ä��"
                        ListBox_�ڻ�ä��.Column = vlist
                End Select
                
            End If
        End If
    End With
End Sub
