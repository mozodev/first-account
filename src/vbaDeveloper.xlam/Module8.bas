Attribute VB_Name = "Module8"
'module8 : ǰ�Ǽ�, ������Ǽ�
Option Explicit

Sub ������Ǵ�������()
    Dim ws_target As Worksheet
    Set ws_target = Worksheets("������Ǵ���")
    Dim ���� As Integer
    ���� = ws_target.Range("���ǳ�¥���̺�").End(xlDown).Row
    MsgBox ����
    
    ws_target.Unprotect

    ws_target.Range("A4:I" & ����).Sort Key1:=ws_target.Range("A3"), Order1:=xlAscending, Key2:=ws_target.Range("B3") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom

End Sub

Sub ������Ǽ��μ�_Click()
    Dim ws As Worksheet
    Set ws = Worksheets("������Ǽ�")
    ws.PageSetup.PrintArea = "$b$3:$l$24"
    ActiveWindow.SelectedSheets.PrintPreview

End Sub

Sub �������_ȸ������Է�()
    Dim ws As Worksheet
    Set ws = Worksheets("������Ǽ�")

    Dim �� As String
    Dim �� As String
    Dim ���� As String
    Dim ȸ���� As String
    Dim �ڵ� As String
    Dim �ݾ� As Long
    Dim �ݾ�2 As Long
    Dim ���� As String
    On Error Resume Next
    With ws
        .Range("t2").Select

        �� = .Range("������").Value
        �� = .Range("�����").Value
        ���� = .Range("���⼼��").Value
        
        ȸ���� = .Range("a1")
        �ڵ� = .Range("b1")
        �ݾ� = .Range("i18")
        �ݾ�2 = .Range("i10")
        
        If �ݾ�2 > 0 Then
            ���� = .Range("b9") & " ��"
        Else
            ���� = .Range("b9")
        End If
    
        .Range("A1").Select
    End With
    
    'ȸ����忡 �Է�
    Dim ws_target As Worksheet
    Dim r������ġ As Range
    Set ws_target = Worksheets("ȸ�����")
    Set r������ġ = ws_target.Range("A5").End(xlDown)
    
    With r������ġ
        .Offset(1, 0).FormulaR1C1 = ȸ����
        .Offset(1, 1).FormulaR1C1 = �ڵ�
        .Offset(1, 3).FormulaR1C1 = "����"
        .Offset(1, 4).FormulaR1C1 = ��
        .Offset(1, 5).FormulaR1C1 = ��
        .Offset(1, 6).FormulaR1C1 = ����
        .Offset(1, 7).FormulaR1C1 = ����
        .Offset(1, 9).FormulaR1C1 = �ݾ�
        .Offset(1, 10).FormulaR1C1 = 0
    End With
    
    If Err.Number = 0 Then
        MsgBox "�Էµƽ��ϴ�"
    End If
        
End Sub

Sub ȸ�����_��������Է�(���ȣ As Integer)
    Dim ws As Worksheet
    Set ws = Worksheets("ȸ�����")

    Dim ȸ���� As String
    Dim �ڵ� As String
    Dim �ݾ� As Long
    Dim ���� As String
    
    On Error Resume Next
    With ws.Range("A" & ���ȣ)
        ȸ���� = .Value
        �ڵ� = .Offset(, 1).Value
        ���� = .Offset(, 7).Value
        �ݾ� = .Offset(, 9).Value

    End With
    
    If �ݾ� > 0 Then
        '������Ǵ��忡 �Է�
        Dim ws_target As Worksheet
        Dim r������ġ As Range
        Set ws_target = Worksheets("������Ǵ���")
        If ws_target.Range("���ǳ�¥���̺�").Offset(1, 0).Value Then
            Set r������ġ = ws_target.Range("���ǳ�¥���̺�").End(xlDown).Offset(1, 0)
        Else
            Set r������ġ = ws_target.Range("���ǳ�¥���̺�").Offset(1, 0)
        End If
        
        With r������ġ
            .FormulaR1C1 = ȸ����
            .Offset(, 1).FormulaR1C1 = �ڵ�
            .Offset(, 2).FormulaR1C1 = ���� '�����
            .Offset(, 4).Value = 1
            .Offset(, 5).Value = �ݾ� '�ܰ�
            .Offset(, 6).Value = �ݾ�
        End With
        
    Else
        MsgBox "������Ǽ��� ������ �� �����ϴ�"
    End If
        
End Sub

Sub ������Ǽ��ۼ�(��¥��ü As Boolean)

    Dim �����(20)
    Dim �԰�(20)
    Dim ����(20)
    Dim �ܰ�(20)
    Dim ���(20)
    Dim �ϴܺ��(20)
    Dim ws As Worksheet
    Set ws = Worksheets("������Ǽ�")
    Dim ws_source As Worksheet
    Set ws_source = Worksheets("������Ǵ���")
    
    Dim ������ As Integer
    Dim ��¥ As String, �ڵ� As String, ���� As String, ������ As String, �׸��ڵ� As String
    Dim i As Integer, i2 As Integer, i3 As Integer
    
    ������ = ActiveCell.Row
    With ws_source
        .Range("a" & ������).Select
        ��¥ = .Range("a" & ������).Value
        �ڵ� = .Range("b" & ������).Value
    
        ���� = ��¥
            
        i2 = 1
            
        For i = 1 To 300
        
            ������ = ActiveCell.Row
            ������ = .Range("a" & ������)
            �׸��ڵ� = .Range("b" & ������)
            
            If ��¥ <> ������ Then 'Or �ڵ� <> �׸��ڵ� Then
                Exit For
            Else
            
                �����(i) = .Range("c" & ������)
                �԰�(i) = .Range("d" & ������)
                ����(i) = .Range("e" & ������)
                �ܰ�(i) = .Range("f" & ������)
                ���(i) = .Range("h" & ������)
                �ϴܺ��(i) = .Range("i" & ������)
                
                i2 = i2 + 1
                
            End If
            
            ActiveCell.Offset(1, 0).Range("A1").Select
            If Not ��¥��ü Then
                Exit For
            End If
        
        Next i
    End With
    
    With ws
        .Activate
        
        .Range("b9:g17").Select
        Selection.ClearContents
        
        .Range("b1").Value = �ڵ�
        .Range("b5").Value = ����
        
        For i3 = 1 To i2
        
            .Range("b" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = �����(i3)
            
            .Range("d" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = �԰�(i3)
            
            .Range("f" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = ����(i3)
            
            .Range("g" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = �ܰ�(i3)
        
        Next i3
        
        .Range("k9").Value = ���(1)
        .Range("c19").Value = �ϴܺ��(1)
    End With
    Erase �����, �԰�, ����, �ܰ�, ���, �ϴܺ��

    ws.Activate
    ws.Visible = xlSheetVisible
    
End Sub
