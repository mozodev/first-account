Attribute VB_Name = "Module6"
'module6 : ���꼭�ڵ��Է�, �������Է�
Option Explicit

Public code_changed As Boolean

Sub ��������_��Ʈ����()

    If Not code_changed Then
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")
    Dim ������ As Integer
    Dim ������ As Integer
    Dim r������ġ As Range
    Set r������ġ = ws.Range("���ʵ�").Offset(3, 0)
    ������ = r������ġ.Row
    ������ = r������ġ.End(xlDown).Row
    ws.Range("A" & ������, "G" & ������).Select
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B4", "B" & ������), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("C4", "C" & ������), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("D4", "D" & ������), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A" & ������, "G" & ������)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.CutCopyMode = False
End Sub

Sub ���꼭�ڵ��Է�()
Attribute ���꼭�ڵ��Է�.VB_Description = ".��(��) 2006-11-15�� ����� ��ũ��"
Attribute ���꼭�ڵ��Է�.VB_ProcData.VB_Invoke_Func = "u\n14"

    Dim ��(100), ��(100), ��(100), ����(100) As String
    Dim c1, c2, c3, c4 As Integer
    Dim c1_1, c2_2, c3_3, c4_4 As String

    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("���꼭")
    ws.Activate
    
With ws
  
    c1 = 0
    c2 = 1
    c3 = 1
    c4 = 1
    
    ��(4) = .Range("b4").Value
    ��(4) = .Range("c4").Value
    ��(4) = .Range("d4").Value
    ����(4) = .Range("d4").Value

    For i = 4 To 100
    
        ��(i) = .Range("b" & i).Value
        ��(i) = .Range("c" & i).Value
        ��(i) = .Range("d" & i).Value
        ����(i) = .Range("e" & i).Value
        
        If ��(i) = "" Then
            Exit For
        End If
        
        If ��(i - 1) <> ��(i) Then
            c1 = c1 + 1
            c2 = 1
            c3 = 1
            c4 = 0
        Else
            If ��(i - 1) <> ��(i) Then
                c2 = c2 + 1
                c3 = 1
                c4 = 0
            Else
                If ��(i - 1) <> ��(i) Then
                    c3 = c3 + 1
                    c4 = 0
                End If
                
            End If
        
        End If
        
        c4 = c4 + 1
 
        If c1 < 10 Then
            c1_1 = "'0" & c1
        Else
            c1_1 = "'" & c1
        End If
        
        If c2 < 10 Then
            c2_2 = "0" & c2
        Else
            c2_2 = c2
        End If
        
        If c3 < 10 Then
            c3_3 = "0" & c3
        Else
            c3_3 = c3
        End If
        
        If c4 < 10 Then
            c4_4 = "0" & c4
        Else
            c4_4 = c4
        End If
         
        .Cells(i, 1) = c1_1 & c2_2 & c3_3 & c4_4
    
    Next i
    
    Dim ii As Integer
    
    For ii = 2 To i - 1
        .Range("o" & ii) = "=IF(RC[-14]>"""",RC[-14]&""/""&RC[-13]&""/""&RC[-12]&""/""&RC[-11]&""/""&RC[-10],"""")"
    Next ii

End With
Erase ��
Erase ��
Erase ��
Erase ����
    ActiveWorkbook.Names.Add name:="�׸���", RefersToR1C1:="=���꼭!R2C15:R" & i - 1 & "C15"

    ActiveWorkbook.Names.Add name:="���׸�", RefersToR1C1:="=���꼭!R2C1:R" & i - 1 & "C5"
    
End Sub
