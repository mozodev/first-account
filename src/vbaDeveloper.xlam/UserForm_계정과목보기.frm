VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_�������񺸱� 
   Caption         =   "�������񺸱�"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595.001
   OleObjectBlob   =   "UserForm_�������񺸱�.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm_�������񺸱�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_close_Click()
    Unload Me
End Sub

Sub ��������ε�(�з� As String, �����ͼҽ� As String)
    Dim ��ü As Range, ã���� As Range, ���ڵ� As Range
    Dim x As Integer, y As Integer

    Dim ���ؿ� As Range
    Dim ws_source As Worksheet
    Set ws_source = Worksheets(�����ͼҽ�)

    Set ���ؿ� = ws_source.Range("���úз�����")
    Dim ���� As Range, �׿� As Range, �� As Range, ���� As Range
    
    Set ���� = ws_source.Range("���ð�����")
    Set �׿� = ws_source.Range("�����׿���")
    Set �� = ws_source.Range("���ø񿭶�")
    Set ���� = ws_source.Range("���ü��񿭶�")
    
    Dim vlist() As Variant
    
    x = 0

    ' Step 1 : ����Ǵ� ���� �׸��� ���� ������
    Set ��ü = ws_source.Range("���ð�����").CurrentRegion.columns(���ؿ�.Column)
    Dim ��� As Integer
       
    ��� = ��ü.Rows.Count
    Dim �з��� As String
    Dim i As Integer
     
    For i = 1 To ���
        �з��� = ws_source.Range("A" & i).Offset(, ���ؿ�.Column - 1).Value
        If �з��� = "����" Or �з��� = �з� Then
            ReDim Preserve vlist(5, x)
            Set ���ڵ� = ws_source.Range("A" & i).Resize(1, ���ؿ�.Column) '�� ���ڸ� �˾Ƴ��� �ٲ���
            vlist(0, x) = ���ڵ�.Cells(, ����.Column)
            vlist(1, x) = ���ڵ�.Cells(, �׿�.Column)
            vlist(2, x) = ���ڵ�.Cells(, ��.Column)
            vlist(3, x) = ���ڵ�.Cells(, ����.Column)
            vlist(4, x) = ���ڵ�.Cells(, ���ؿ�.Column)
    
            x = x + 1
            ListBox_�������񺸱�.Column = vlist
        End If
    Next i
    
    If x = 0 Then
        MsgBox "�˻������ �������� �ʽ��ϴ�"
        ListBox_�������񺸱�.Clear
    End If

End Sub

Sub ��������ε�2(���� As String, �����ͼҽ� As String)
    ' ���� ��������ε� �Լ��� ����� ��
    Dim ��ü As Range, ã���� As Range, ���ڵ� As Range, ���ؿ� As Range
    Dim x As Integer, y As Integer
    Dim ù��ġ As String, �ʵ�� As String
    
    Dim ws_source As Worksheet
    Set ws_source = Worksheets(�����ͼҽ�)
    
    Dim ���� As Range
    Set ���� = ws_source.Range("A2").CurrentRegion.Rows(2)
    Set ���ؿ� = ws_source.Range("�з�����")
    
    Dim ���� As Range, �׿� As Range, �� As Range, ���� As Range
    Set ���� = ����.Find(What:="��", LookAt:=xlPart)
    Set �׿� = ����.Find(What:="��", LookAt:=xlPart)
    Set �� = ����.Find(What:="��", LookAt:=xlPart)
    Set ���� = ����.Find(What:="����", LookAt:=xlPart)
    
    Dim vlist() As Variant
    
    y = 0
    
    Set ��ü = ws_source.Range("���ð�����").CurrentRegion.columns(���ؿ�.Column)
    Set ã���� = ��ü.Find(What:=1, LookAt:=xlPart)
    
    If Not ã���� Is Nothing Then
        ù��ġ = ã����.Address
        
        Do
            ReDim Preserve vlist(5, x)
            Set ���ڵ� = ã����.End(xlToLeft).Resize(1, ����.columns.Count)
            vlist(0, x) = ���ڵ�.Cells(, ����.Column)
            vlist(1, x) = ���ڵ�.Cells(, �׿�.Column)
            vlist(2, x) = ���ڵ�.Cells(, ��.Column)
            vlist(3, x) = ���ڵ�.Cells(, ����.Column)
            vlist(4, x) = ���ڵ�.Cells(, ���ؿ�.Column)
    
            x = x + 1
    '        Y = 0
    '
            Set ã���� = ��ü.FindNext(ã����)
        Loop While Not ã���� Is Nothing And ã����.Address <> ù��ġ
    
        ListBox_�������񺸱�.Column = vlist
    Else
        MsgBox "�˻������ �������� �ʽ��ϴ�"
        ListBox_�������񺸱�.Clear
    End If

End Sub

Private Sub UserForm_Initialize()
    With ListBox_�������񺸱�
        .columnCount = 4
        .ColumnWidths = "1cm;2.7cm;3cm;3.5cm"
    End With
End Sub
