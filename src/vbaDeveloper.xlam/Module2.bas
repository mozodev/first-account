Attribute VB_Name = "Module2"
'module2 : ��꼭�ʱ�ȭ, ��������1, ��������2
Option Explicit

Const ���̸�_code As String = "A"
Const ���̸�_�� As String = "B"
Const ���̸�_�� As String = "C"
Const ���̸�_�� As String = "D"
Const ���̸�_���� As String = "e"
Const ���̸�_����� As String = "f"
Const ���̸�_�Ⱓ���� As String = "g"
Const ���̸�_�Ⱓ���� As String = "h"
Const ���̸�_�������� As String = "i"
Const ���̸�_�������� As String = "j"
Const ���̸�_�����ܾ� As String = "k"
Const ���̸�_���� As String = "l"

Sub ��꼭�ʱ�ȭ2()
    '��ü ������ �ٽ� �������� �ʰ� �Ⱓ����/����, ��������/���� ���� ����.
    '��꼭 ������Ҷ� �δ� ���̱� ���� ��� �ʱ�ȭ
    Dim ���꼭 As Worksheet
    Dim ��꼭 As Worksheet
    Set ���꼭 = Worksheets("���꼭")
    Set ��꼭 = Worksheets("��꼭")
    ���꼭.Activate
    
    Dim ���� As Integer
    ���� = ���꼭.Range("a4").End(xlDown).Row
    
    ��꼭.Activate
    
    With Range(���̸�_�Ⱓ���� & "6:" & ���̸�_�������� & ���� + 2)
        .ClearContents
    End With
End Sub

Sub ��꼭�ʱ�ȭ()
Attribute ��꼭�ʱ�ȭ.VB_ProcData.VB_Invoke_Func = "i\n14"

    Dim ���꼭 As Worksheet
    Dim ��꼭 As Worksheet
    Set ���꼭 = Worksheets("���꼭")
    Set ��꼭 = Worksheets("��꼭")
    ���꼭.Activate
    
    Dim ���� As Integer
    ���� = ���꼭.Range("a4").End(xlDown).Row
    
    ��꼭.Activate
   
    With Range("A6:" & ���̸�_���� & "200")
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .ClearContents
    End With
        
    Range("A6:A180").NumberFormatLocal = "G/ǥ��"

    Dim i As Integer
    With ��꼭
        For i = 6 To ���� + 2
            Application.ScreenUpdating = False
            With .Range("a" & i)
                .FormulaR1C1 = "=���꼭!R[-2]C" '�ڵ�
                .Offset(0, 1).FormulaR1C1 = "=���꼭!R[-2]C"  '��
                .Offset(0, 2).FormulaR1C1 = "=���꼭!R[-2]C"  '��
                .Offset(0, 3).FormulaR1C1 = "=���꼭!R[-2]C"  '��
                .Offset(0, 4).FormulaR1C1 = "=���꼭!R[-2]C"  '����
                .Offset(0, 5).FormulaR1C1 = "=���꼭!R[-2]C"  '�����
                .Offset(0, 10).FormulaR1C1 = "=RC[-5]-RC[-2]-RC[-1]" '�����ܾ�
                .Offset(0, 11).FormulaR1C1 = "=If(RC[-6] > 0, (RC[-6]-RC[-1])/RC[-6], 0)"
            End With
        Next i
        
        Dim i3 As Integer

        i3 = .Range("a" & ���� + 2).Row
        
        .Range("b" & i3 + 1).FormulaR1C1 = "�հ�"
        
        .Range(���̸�_�Ⱓ���� & i3 + 1).FormulaR1C1 = "=sum(r[-" & ���� - 3 & "]C:R[-1]C)"
        .Range(���̸�_�Ⱓ���� & i3 + 1).FormulaR1C1 = "=sum(r[-" & ���� - 3 & "]C:R[-1]C)"
        .Range(���̸�_�������� & i3 + 1).FormulaR1C1 = "=sum(r[-" & ���� - 3 & "]C:R[-1]C)"
        .Range(���̸�_�������� & i3 + 1).FormulaR1C1 = "=sum(r[-" & ���� - 3 & "]C:R[-1]C)"
        
        .Range("b" & i3 + 2).FormulaR1C1 = "���ܾ�"
        .Range("h" & i3 + 2).FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
        .Range("j" & i3 + 2).FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
        
        With .Range("A5:" & ���̸�_���� & ���� + 4)
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
    End With

End Sub

Sub ��꼭��������1()
 
Dim ���� As Integer
Dim ws As Worksheet
Dim �ӽ������� As Integer
Dim �ӽù��ڿ� As String
Dim ������ġ As Range
Dim �� As String, �� As String, �� As String
Dim ������ As Integer, ���׸�� As Integer

Dim ��꼭 As Worksheet
Set ��꼭 = Worksheets("��꼭")
    ��꼭.Activate
    ��꼭.Select
    ��꼭.Copy After:=Sheets(9)
    Set ��꼭 = ActiveSheet

With ��꼭
    .Activate
    ���� = .Range("A5").End(xlDown).Row
    
    ���׸�� = ���� - 5
    If ���׸�� > 30 Then
        ReDim �����(���׸��)
        ReDim ������(���׸��)
        ReDim �׻���(���׸��)
        ReDim ������(���׸��)
        ReDim ������(���׸��)
        ReDim ������(���׸��)
    End If

    .Range(���̸�_�� & "5:" & ���̸�_�� & ����).Select

    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With

    .Range("j1").FormulaR1C1 = " �� ��Ʈ�� �����ϼŵ� �˴ϴ�."
    .Range("a6").Select

    �� = Left(ActiveCell.Value, 6)
    �����(1) = 6
    ������(1) = .Range(���̸�_�� & "6")
    
    Dim i As Integer
    Dim i2 As Integer
    i2 = 2
    
    For i = 1 To 300
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 6) = "" Then
    
          �����(i2) = ActiveCell.Row
          ������(i2) = Range(���̸�_�� & �����(i2))
    
          i2 = i2 + 1
    
            Exit For
        End If
    
        If �� <> Left(ActiveCell.Value, 6) Then
          �����(i2) = ActiveCell.Row
          ������(i2) = Range(���̸�_�� & �����(i2))
    
          i2 = i2 + 1
    
        End If
    
        �� = Left(ActiveCell.Value, 6)
    
    Next i
    
    Dim i3 As Integer
    Dim i4 As Integer
    i4 = 0
    
    For i3 = 1 To i2 - 2
    
        Range("A" & �����(i3) + i4).Select
        Selection.EntireRow.Insert
    
        ������ = ActiveCell.Row
        .Range(���̸�_�� & ������).Select
        ActiveCell.FormulaR1C1 = ������(i3)
        Selection.Font.Bold = True
    
        .Range(���̸�_�� & ������).Select
    
        With Selection
            .HorizontalAlignment = xlLeft
            .Font.ThemeColor = xlThemeColorLight1
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        
        With Selection.Borders(xlEdgeLeft)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
        .Range(���̸�_����� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�����ܾ� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & �����(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_���� & ������).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
    
        .Range(���̸�_�� & ������ & ":" & ���̸�_���� & ������).Select
        With Selection.Interior
            .ColorIndex = 35
        End With
    
        i4 = i4 + 1
    
    Next i3
    
    Erase �����
    Erase ������
    
    '�� ����
    
    .Range("a6").Select
    
    �� = Left(ActiveCell.Value, 4)
      
    �׻���(1) = 6
    ������(1) = Range("b6")
    i2 = 2
    
    For i = 1 To 150
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 4) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
        If Left(ActiveCell.Value, 4) = Empty And �� = Empty Then
             Exit For
        End If
    
    
        If �� <> Left(ActiveCell.Value, 4) Then
    
            If �� <> Left(ActiveCell.Value, 4) Then
    
                �׻���(i2) = ActiveCell.Row
                ������(i2) = Range(���̸�_�� & �׻���(i2))
                i2 = i2 + 1
            End If
    
        End If
    
        �� = Left(ActiveCell.Value, 4)
    
    Next i
    �׻���(i2) = �׻���(i2) - 2
    
    i4 = -1
    
    
    For i3 = 2 To i2 - 2
        .Range("A" & �׻���(i3) + i4).Select
        Selection.EntireRow.Insert
    
        ������ = ActiveCell.Row
        .Range(���̸�_�� & ������).Select
        ActiveCell.FormulaR1C1 = ������(i3)
        Selection.Font.Bold = True
    
        .Range(���̸�_�� & ������ & ":" & ���̸�_�� & ������).Select
    
        With Selection
            .HorizontalAlignment = xlLeft
    
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
    
        End With
    
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .Weight = xlThin
        End With
    
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
        .Range(���̸�_����� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�����ܾ� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & �׻���(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_���� & ������).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
        
        .Range(���̸�_�� & ������ & ":" & ���̸�_���� & ������).Select
        With Selection.Interior
            .ColorIndex = 34
        End With
    
        i4 = i4 + 1
    
    Next i3
    
    Erase �׻���
    Erase ������

    '�� ����
    
    .Range("a6").Select
    
    �� = Left(ActiveCell.Value, 2)
    
    ������(1) = 6
    ������(1) = Range("b6")
    i2 = 2
    
    For i = 1 To 150
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 2) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
        
        If Left(ActiveCell.Value, 2) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
         
        If Left(ActiveCell.Value, 2) = Empty And �� = Empty Then
             Exit For
        End If
    
    
        If �� <> Left(ActiveCell.Value, 2) Then
    
            If �� <> Left(ActiveCell.Value, 2) Then
    
                ������(i2) = ActiveCell.Row
                ������(i2) = Range(���̸�_�� & ������(i2))
                i2 = i2 + 1
            End If
    
        End If
    
        �� = Left(ActiveCell.Value, 2)
    
    Next i
    ������(i2) = ������(i2) - 2
    
    i4 = -2
    
    
    For i3 = 2 To i2 - 2
        .Range("A" & ������(i3) + i4).Select
        Selection.EntireRow.Insert
    
        ������ = ActiveCell.Row
        .Range(���̸�_�� & ������).Select
        ActiveCell.FormulaR1C1 = ������(i3)
        Selection.Font.Bold = True
    
        .Range(���̸�_�� & ������ & ":" & ���̸�_�� & ������).Select
    
        With Selection
            .HorizontalAlignment = xlLeft
    
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
    
        End With
    
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .Weight = xlThin
        End With
    
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    
        .Range(���̸�_����� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�Ⱓ���� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�������� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_�����ܾ� & ������).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & ������(i3 + 1) - ������ + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(���̸�_���� & ������).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
        
        .Range(���̸�_�� & ������ & ":" & ���̸�_���� & ������).Select
        With Selection.Interior
            .ColorIndex = 37
        End With
     
        i4 = i4 + 1
    
    Next i3
    Erase ������
    Erase ������
   
    .Range("b5:" & ���̸�_���� & "5").Select
    With Selection
        .Font.Bold = True
        .Interior.ColorIndex = 33
        .Font.ThemeColor = xlThemeColorLight1
    End With


' �Ѱ� ����

    .Range(���̸�_����� & "6").End(xlDown).Select

    ������ = ActiveCell.Row + 1
    Dim ó���� As Integer
    ó���� = ������ - 5

    .Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "��    ��"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":d" & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    .Range(���̸�_�Ⱓ���� & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(���̸�_�Ⱓ���� & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(���̸�_�������� & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(���̸�_�������� & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":" & ���̸�_���� & ������).Select
    With Selection.Interior
        .ColorIndex = 33
    End With


'�ܾ� ����
    ������ = ������ + 1

    .Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    Selection.Font.Bold = True

    .Range(���̸�_�� & ������ & ":" & ���̸�_�� & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeTop)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeBottom)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeRight)

        .Weight = xlThin

    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    .Range(���̸�_�Ⱓ���� & ������).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    .Range(���̸�_�������� & ������).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    .Range(���̸�_�� & ������ & ":" & ���̸�_���� & ������).Select
    With Selection.Interior
        .ColorIndex = 33
    End With

    columns("b:c").ColumnWidth = 4
    columns("e:k").EntireColumn.AutoFit

    .Shapes("Picture 6").Select
    Selection.ShapeRange.IncrementLeft -20

    .Shapes("Button 2").Select
    Selection.Cut

    .Shapes("Button 1").Select
    Selection.Cut

    .PageSetup.PrintArea = "$b$2:$k$" & ������

End With

'�޸� �����
Application.CutCopyMode = False

End Sub


Sub ��꼭����2()

    Dim �� As String
    Dim i As Integer, i2 As Integer, i3 As Integer, i4 As Integer
    
    Dim ��꼭 As Worksheet
    Dim ws As Worksheet
    Set ��꼭 = Worksheets("��꼭")
    
    ��꼭.Select
    ��꼭.Copy After:=Sheets(9)
    
    Range("b5:k5").Select
    With Selection.Interior
        .ColorIndex = 33
    End With
    
    Set ws = ActiveSheet
ws.Range("i1").Select
ActiveCell.FormulaR1C1 = " �� ��Ʈ�� �����ϼŵ� �˴ϴ�."

Dim �׻���(30)

ws.Range("a6").Select

�� = Left(ActiveCell.Value, 4)

i2 = 1

For i = 1 To 300
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    If Left(ActiveCell.Value, 4) = "" Then
      �׻���(i2) = ActiveCell.Row
      i2 = i2 + 1
         Exit For
    End If

    If �� <> Left(ActiveCell.Value, 4) Then
      �׻���(i2) = ActiveCell.Row
      i2 = i2 + 1
    End If
   
    �� = Left(ActiveCell.Value, 4)

Next i

Dim ������ As Integer

i4 = 0
�׻���(0) = 6

For i3 = 1 To i2 - 1
    Range("A" & �׻���(i3) + i4).Select
    Selection.EntireRow.Insert

    ������ = ActiveCell.Row
    ws.Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    Selection.Font.Bold = True
    
    Range(���̸�_�� & ������ & ":d" & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    ws.Range("e" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,R[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("f" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,R[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("j" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - �׻���(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True
    
    ws.Range("k" & ������).Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
    Selection.Font.Bold = True
    
    Range(���̸�_�� & ������ & ":k" & ������).Select
    With Selection.Interior
        .ColorIndex = 35
    End With
    
    If �׻���(i3 - 1) + i4 + 1 <> ������ Then
        Range(���̸�_�� & �׻���(i3 - 1) + i4 + 1 & ":c" & ������ - 1).Select
        Selection.Font.ColorIndex = 2
        
        Range(���̸�_�� & �׻���(i3 - 1) + i4 & ":c" & ������ - 1).Select

        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            
            .Weight = xlThin
     
        End With
        With Selection.Borders(xlEdgeTop)
            
            .Weight = xlThin
     
        End With
        With Selection.Borders(xlEdgeBottom)
            
            .Weight = xlThin
     
        End With
        With Selection.Borders(xlEdgeRight)
            
            .Weight = xlThin
     
        End With
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    End If

    i4 = i4 + 1

Next i3


'�� ����

Dim ������(30)

ws.Range("a6").Select
Dim �� As String
�� = Left(ActiveCell.Value, 2)

i2 = 1

For i = 1 To 150
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    If Left(ActiveCell.Value, 2) = Empty And �� = Empty Then
         Exit For
    End If

    If �� <> Left(ActiveCell.Value, 2) Then
      ActiveCell.Offset(1, 0).Range("A1").Select
      
        If �� <> Left(ActiveCell.Value, 2) Then
      
            ������(i2) = ActiveCell.Row
            i2 = i2 + 1
        End If
      
    End If
    
    �� = Left(ActiveCell.Value, 2)

Next i

i4 = 0
������(0) = 5

For i3 = 1 To i2 - 1
    Range("A" & ������(i3) + i4).Select
    Selection.EntireRow.Insert

    ������ = ActiveCell.Row
    ws.Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":d" & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    ws.Range("e" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("f" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("j" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ������ - ������(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("k" & ������).Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":k" & ������).Select
    With Selection.Interior
        .ColorIndex = 34
    End With
    

    If ������(i3 - 1) + i4 + 1 <> ������ Then
        Range(���̸�_�� & ������(i3 - 1) + i4 + 1 & ":b" & ������ - 1).Select
        Selection.Font.ColorIndex = 2

        Range(���̸�_�� & ������(i3 - 1) + i4 & ":b" & ������ - 1).Select

        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

     End If

    i4 = i4 + 1

Next i3

Dim ó���� As Integer
' �Ѱ� ����
    Range("A" & ������ + 1).Select

    ó���� = ������ - ������(0) + 1

    ������ = ActiveCell.Row

    ws.Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "��    ��"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":d" & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    ws.Range("f" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & ������).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & ó���� & "]C:R[-1]C)"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":k" & ������).Select
    With Selection.Interior
        .ColorIndex = 33
    End With

'�ܾ� ����
    Range("A" & ������ + 1).Select

    ������ = ActiveCell.Row

    ws.Range(���̸�_�� & ������).Select
    ActiveCell.FormulaR1C1 = "�� �� ��"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":d" & ������).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeTop)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeBottom)
'
        .Weight = xlThin

    End With
    With Selection.Borders(xlEdgeRight)
 
        .Weight = xlThin

    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    ws.Range("g" & ������).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    ws.Range("i" & ������).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    Range(���̸�_�� & ������ & ":k" & ������).Select
    With Selection.Interior
        .ColorIndex = 33
    End With
    
    ws.Shapes("Button 2").Select
    Selection.Cut
    ws.Shapes("Button 3").Select
    Selection.Cut
    ws.Shapes("Button 1").Select
    Selection.Cut

    ws.PageSetup.PrintArea = "$b$2:$k$" & ������


'�޸� �����
Application.CutCopyMode = False

End Sub
