Attribute VB_Name = "Module12"
'module12 : ��꼭
Option Explicit

Const ��꼭_��_�������� As String = "i"
Const ��꼭_��_�������� As String = "j"
Const ��꼭_��_�Ⱓ���� As String = "g"
Const ��꼭_��_�Ⱓ���� As String = "h"
Const �׸�а���_��_��¥ As String = "A"
Const �׸�а���_��_�� As String = "C"
Const �׸�а���_��_�� As String = "D"
Const �׸�а���_��_�� As String = "E"
Const �׸�а���_��_���� As String = "f"
Const �׸�а���_��_���� As String = "g"
Const �׸�а���_��_���� As String = "h"
Const �׸�а���_��_���� As String = "i"
Const �ִ���� As Integer = 20000
Const ȸ�����_������ As Integer = 5

Public project As String
Public rebuild_report As Boolean
Public report_1p As Boolean

Sub �׸����ۼ�(ByVal �׸�а������ As Boolean, Optional ByVal project As String)

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "��꼭�� �ۼ��ϰ� �ֽ��ϴ�. ��� ��ٷ��ּ���.."

    Dim ws_ledger As Worksheet
    Dim �׸�а��� As Worksheet
    Dim ws_help As Worksheet, ws_budget As Worksheet, ws_settle As Worksheet
    
    Set ws_ledger = Worksheets("ȸ�����")
    Set �׸�а��� = Worksheets("�׸�а���")
    Set ws_help = Worksheets("����")
    Set ws_budget = Worksheets("���꼭")
    Set ws_settle = Worksheets("��꼭")
    
    ws_ledger.Activate

    Dim wkSht As Worksheet
 
    'On Error Resume Next                                        '���� �߻��ص� ��� �ڵ� ����
    Set wkSht = ThisWorkbook.Worksheets("����")  '��Ʈ�� ��ü������ ����
    
    If Err.Number = 0 Then                                        '������ ���ٸ�
                                                   '�� ��Ʈ�� ������
    Else                               '(���� ��Ʈ�� ��ü������ �����Ƿ�)�׿� �޸� ������ �߻��ϸ�
        MsgBox "������ �ջ�Ǿ����ϴ�. http://firstaccounting.org ���� �ٽ� �ٿ� �޾Ƽ� ����ϼ���.": Exit Sub                                            '�� ��Ʈ�� �������� ����
    End If

    If ws_help.Range("aa1") <> 300 Then MsgBox "������ �ջ�Ǿ����ϴ�. http://kimjy.net ���� �ٽ� �ٿ� �޾Ƽ� ����ϼ���.": Exit Sub
    
    Set ws_help = Nothing
    Set wkSht = Nothing
    
    Dim ������ As Integer, ���� As Integer, �Է��� As Integer
    Dim startDate As Date, endDate As Date, ȸ������� As Date, ���� As Date
    Dim i As Integer, i2 As Integer, i3 As Integer
    
    startDate = format(get_config("������"), "Short Date")
    endDate = format(get_config("������"), "Short Date")
    ȸ������� = format(get_config("ȸ�������"), "Short Date")
    ���� = format(Date, "Short Date")

    �Է��� = 6
    ������ = 6
    ���� = ws_ledger.Range("�����ʵ巹�̺�").End(xlDown).Row

' ���μ���
' #1. ��ü �Ⱓ ��� �ǹ����̺� ���� (�Ʒ����� ��꼭�� �ݿ�)
' #2. ������ �Ⱓ �׸�а��� �ۼ� ����
' #3. ���� ������ ����� �� ������ ī�� ��, ���� �Ⱓ ��� �ǹ����̺� ���� (�Ʒ����� ��꼭�� �ݿ�)
' #4. �׸�а��� �ۼ� �Ϸ�
' #5. ��꼭 ���� ������Ʈ : ��ü �Ⱓ, ���� �Ⱓ

' #0. ������, �����Ͽ� ���� ������ ���� �ľ�
'�����ٰ� ���� �˾Ƴ��� �κ�
    Dim dataCount As Integer
    dataCount = ���� - 5
    Dim startRow As Integer, endRow As Integer
    Dim c As Range
    Set c = ws_ledger.Range("A5")
    Dim dateArray() As Variant
    ReDim dateArray(1 To dataCount)
    
    For i = 1 To dataCount
        dateArray(i) = c.Offset(i, 0).Value
    Next i
    
    If date_compare(startDate, endDate) < 0 Then
        MsgBox "�������� �����Ϻ��� �����ϴ�. 1�� ��ü ��길 ǥ�õ˴ϴ�"
        Exit Sub
    End If
    If date_compare(dateArray(dataCount), startDate) > 0 Then
        MsgBox "�������� �߸� �����Ǿ����ϴ�. 1�� ��ü ��길 ǥ�õ˴ϴ�."
        Exit Sub
    End If
    If date_compare(dateArray(1), endDate) < 0 Then
        MsgBox "�������� �߸� �����Ǿ����ϴ�. 1�� ��ü ��길 ǥ�õ˴ϴ�."
        Exit Sub
    End If
    
    For i = 1 To dataCount
        If dateArray(i) = startDate Then '�� �κ��� ���� �۵�
            startRow = i + 5
            Exit For
        Else
            If date_compare(dateArray(i), startDate) < 0 Then '���� �����Ϻ��� ���� ��¥�̸� �� ��ȸ
                startRow = i + 5
                Exit For
            End If
        End If
    Next i
    
    For i = dataCount To 1 Step -1
        If date_compare(dateArray(i), endDate) = 0 Then
            endRow = i + 5
            Exit For
        Else
            If date_compare(dateArray(i), endDate) > 0 Then '���� ������ �����Ϻ��� ���� ��¥�� ��Ÿ���ٸ� �̰� ����������
                endRow = i + 5
                Exit For
            End If
        End If
    Next i

    Erase dateArray
    '������ ���� ������ ��꼭 ���� ���ϰ�
    If Not startRow < endRow Then
        MsgBox "�ش�Ⱓ�� �Էµ� �ڷᰡ ���ų� �����ϰ� �������� �߸� �����Ǿ����ϴ�. �Ⱓ�� �ٽ� ������ �õ��غ�����"
        Exit Sub
    End If
    
    �׸�а���.Activate
    'ws_help.Activate
    ws_budget.Activate
    ws_settle.Activate
        
' #1. ��ü �Ⱓ ��� �ǹ����̺� ���� (�Ʒ����� ��꼭�� �ݿ�)

    If rebuild_report Then
        'Call ��꼭�ʱ�ȭ
        Call ����ǹ����̺����(ȸ�������, ����)
    End If

' #2. ������ �Ⱓ �׸�а��� �ۼ� ����
    If �׸�а������ Then

        With �׸�а���.Range("A6:" & �׸�а���_��_���� & �ִ����) ' �� �ִ���� ���� ������ �ڵ� ������ ����
            .ClearContents
            .Font.Bold = False
        
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            
        End With
   
    
    '�׸�а������� ȸ������� �ش糯¥ ������ ����
        Call ���嵥����copy("�׸�а���", startRow, endRow)
    
    '�޸� �����
        Application.CutCopyMode = False
 End If
 
' #3. ī�ǵ� �����͸� ��������, ���� �Ⱓ�� ��� �ǹ����̺� ���� (�Ʒ����� ��꼭�� �ݿ�)
    If rebuild_report Then
        Call ����ǹ����̺����(startDate, endDate)
    End If

' #4. ��꼭 ���� ������Ʈ : ��ü �Ⱓ, ���� �Ⱓ (���� 2�� �Ϸ�: 2016.9.20)


' #4-1. ��꼭 �ۼ�
    If rebuild_report Then
        Dim codeCount As Integer
        codeCount = ws_settle.Range("A5").End(xlDown).Row - 5
        
        '�����ڵ带 �迭�� �ϰ� ��ȯ 2016. 9.20
        Dim sCode() As Variant
        ReDim sCode(1 To codeCount)
        Dim settleResult() As Variant
        ReDim settleResult(1 To codeCount, 1 To 4)
        
        Dim j As Integer
        For j = 1 To codeCount
            sCode(j) = ws_settle.Range("A" & j + 5).Value
        Next j
        
        For j = 1 To codeCount
            settleResult(j, 1) = �ڵ���(sCode(j), "����", "�κ�")
            settleResult(j, 2) = �ڵ���(sCode(j), "����", "�κ�")
            settleResult(j, 3) = �ڵ���(sCode(j), "����", "��ü")
            settleResult(j, 4) = �ڵ���(sCode(j), "����", "��ü")
        Next j
    
        With ws_settle
            .Range(��꼭_��_�Ⱓ���� & 6, ��꼭_��_�������� & codeCount + 5).Value = settleResult
            .Range("a3").Value = "(" & startDate & " ~ " & endDate & ")"
        End With
    End If
    ws_settle.Visible = xlSheetVisible
    
    If �׸�а������ Then
    ' #4-2. �׸�а��� �ۼ�
        �׸�а���.Activate
        dataCount = 0
        Application.StatusBar = "������������ �ۼ����Դϴ�. ��ø� ��ٷ��ּ���."
        
        With �׸�а���.Range("a6:" & �׸�а���_��_���� & (���� - ������ + 6))
            .Sort Key1:=Range("B6"), Order1:=xlAscending, Key2:=Range("A6") _
            , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, _
            Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers, DataOption2 _
            :=xlSortNormal
        End With
        
        With �׸�а���
            '.Range("b6").Select  '�׸�а��� ù��° ������
    
            Dim �����ټ� As Integer
            Dim ������ As Integer
            dataCount = .Range("A5").End(xlDown).Row - 5
            �����ټ� = 0
            '����ܼ���, ��������� ���� ����
            For i3 = 6 To (dataCount - 1 + 6)
                If Left(.Range("b" & i3).Value, 2) = "00" Then
                
                    ������ = i3 '.Row
                    .Range("A" & ������ & ":" & �׸�а���_��_���� & ������).Delete Shift:=xlUp
                    '.Range("b" & ������).Select
                    �����ټ� = �����ټ� + 1
                End If
            Next i3
            dataCount = dataCount - �����ټ�
            '����ܼ���, ��������� ���� �Ϸ�
            
            Dim �ڵ� As String, �����ڵ� As String
            '�ڵ� = .Range("b6").Value
            'Dim �Է��ټ� As Integer
           ' �Է��ټ� = 1
            
            '�� �κ��� ���� �����ؾ� ��
            
            Dim tmpDataArray() As Variant
            ReDim tmpDataArray(1 To dataCount, 1 To 9)
            Set c = .Range("A5")
            tmpDataArray = .Range(c.Offset(1, 0), c.Offset(dataCount, 8)).Value
            
            Dim �Է��ټ� As Integer, �Է��� As Integer, �����ټ� As Integer
            �Է��ټ� = 1
            �Է��� = 0
            �����ټ� = 0
            �����ڵ� = ""
            Set c = .Range("b5")
            
            For i3 = 1 To dataCount
                �ڵ� = tmpDataArray(i3, 2)
                
                If �����ڵ� <> "" Then 'ù ���� ������ �ʴ´�
                    If �ڵ� <> �����ڵ� Then
                        �Է��� = i3 + 5 + �����ټ�
                        
                        .Rows(�Է���).Insert
                        �����ټ� = �����ټ� + 1
                        
                        .Range("a" & �Է���).Value = "�� ��"
                        .Range("b" & �Է���).Value = Range("b" & �Է��� - 1).Value
                
                        With .Range(�׸�а���_��_���� & �Է���)
                            .FormulaR1C1 = "=SUbtotal(9,R[-" & �Է��ټ� & "]C:R[-1]C)"
                            '.Font.Bold = True
                        End With
                
                        With .Range(�׸�а���_��_���� & �Է���)
                            .FormulaR1C1 = "=SUbtotal(9,R[-" & �Է��ټ� & "]C:R[-1]C)"
                            '.Font.Bold = True
                        End With
                
                        With .Range("a" & �Է��� & ":" & �׸�а���_��_���� & �Է���)
                            .Borders(xlEdgeBottom).Weight = xlMedium
                            .Font.Bold = True
                        End With
                        
                        �Է��ټ� = 1
                    Else
                        �Է��ټ� = �Է��ټ� + 1
                    End If
                    
                End If
                
                �����ڵ� = �ڵ�
            Next i3
            
            Erase tmpDataArray
            
            '������ �ڵ� �հ�
            �Է��� = dataCount + 5 + �����ټ� + 1
            .Rows(�Է���).Insert
            
            .Range("a" & �Է���).Value = "�� ��"
            .Range("b" & �Է���).Value = Range("b" & �Է��� - 1).Value
    
            With .Range(�׸�а���_��_���� & �Է���)
                .FormulaR1C1 = "=SUbtotal(9,R[-" & �Է��ټ� & "]C:R[-1]C)"
            End With
    
            With .Range(�׸�а���_��_���� & �Է���)
                .FormulaR1C1 = "=SUbtotal(9,R[-" & �Է��ټ� & "]C:R[-1]C)"
            End With
    
            With .Range("a" & �Է��� & ":" & �׸�а���_��_���� & �Է���)
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Font.Bold = True
            End With
                        
            '�Ѿ� ����ϰ� ������
            Dim finalRow As Integer
            finalRow = �Է��� + 1
            
            .Range("a" & finalRow).Value = "�� ��"
        
            With .Range(�׸�а���_��_���� & finalRow)
                .FormulaR1C1 = "=subtotal(9,R[-" & finalRow - 5 & "]C:R[-1]C)"
            End With
        
            With .Range(�׸�а���_��_���� & finalRow)
                .FormulaR1C1 = "=subtotal(9,R[-" & finalRow - 5 & "]C:R[-1]C)"
            End With
        
            With Range("a" & finalRow & ":" & �׸�а���_��_���� & finalRow)
                .Font.Bold = True
                With .Borders(xlEdgeBottom)
                    .Weight = xlMedium
                End With
            End With
        
            .Range("a" & finalRow + 1).Value = "�� ��"
        
            With .Range(�׸�а���_��_���� & finalRow + 1)
                .FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
                '.Font.Bold = True
            End With
        
            With .Range("a" & finalRow + 1 & ":" & �׸�а���_��_���� & finalRow + 1)
                .Font.Bold = True
                With .Borders(xlEdgeBottom)
                    .Weight = xlMedium
                End With
            End With
        
            .Range("a2").Value = "(" & startDate & " ~ " & endDate & ")"
            .PageSetup.PrintArea = "$a$1:$" & �׸�а���_��_���� & "$" & finalRow + 1
    
        End With
    End If

    �׸�а���.Visible = xlSheetVisible
End Sub

Sub ���1p()
'�� ������ �ջ��Ͽ� ������ ��꼭 ����
    Const ��꼭������ = 5
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    Dim ws_source, ws_target As Worksheet
    Set ws_source = Worksheets("��꼭")
    Set ws_target = Worksheets("��꼭1p")
    
    ws_target.Range("A7", "A1000").EntireRow.Delete
    
    Dim ����, �׶�, �Ⱓ���Զ�, �Ⱓ����� As Range
    With ws_source
        Set ���� = .Range("B5")
        Set �׶� = .Range("C5")
        Set �Ⱓ���Զ� = .Range("G5")
        Set �Ⱓ����� = .Range("H5")
    End With
    
    Dim ����list() As Variant, ����list() As Variant
    Dim ������ As Range
    Dim ��� As Integer, ���� As Integer, i As Integer, x As Integer, y As Integer
    x = 0
    y = 0
    Dim ���� As Boolean
    Dim ���Աݾ� As Long, ����ݾ� As Long, ����� As Long
    Dim ������ As String
    
    'Set ������ = �׶�
    ��� = �׶�.End(xlDown).Row - ��꼭������
    
    For i = 1 To ���
        Set ������ = �׶�.Offset(i, 0)
        ����� = ������.Offset(, 3).Value
        ���Աݾ� = ������.Offset(, 4).Value
        ����ݾ� = ������.Offset(, 5).Value
        ���� = True
        
        If ������.Offset(, -1).Value = "����" Then
            If x > 0 Then
                If ������.Value = ������ Then
                    ���� = False
                End If
            End If
            
            If ���� Then
                ReDim Preserve ����list(3, x)
                ����list(0, x) = ������.Value
                ����list(1, x) = �����
                ����list(2, x) = ���Աݾ�
                x = x + 1
                ������ = ������.Value
            Else
                ����list(1, x - 1) = ����list(1, x - 1) + �����
                ����list(2, x - 1) = ����list(2, x - 1) + ���Աݾ�
            End If
        Else
            If y > 0 Then
                If ������.Value = ������ Then
                    ���� = False
                End If
            End If
            
            If ���� Then
                ReDim Preserve ����list(3, y)
                ����list(0, y) = ������.Value
                ����list(1, y) = �����
                ����list(2, y) = ����ݾ�
                y = y + 1
                ������ = ������.Value
            Else
                ����list(2, y - 1) = ����list(2, y - 1) + ����ݾ�
                ����list(1, y - 1) = ����list(1, y - 1) + �����
            End If
        End If
        
    Next i
    
    Set ������ = Nothing
    
    '����� ������ ��꼭1p ä���
    With ws_target
        '����
        Dim �����׼�, �����׼�, �׼� As Integer
        �����׼� = UBound(����list, 2) + 1
        For i = 0 To �����׼� - 1
            With .Range("��꼭1p������").Offset(i + 1, 0)
                .Value = ����list(0, i)

                .Offset(0, 1).Value = ����list(1, i)
                .Offset(0, 2).Value = ����list(2, i)
            End With

        Next i
        Erase ����list
        
        �����׼� = UBound(����list, 2) + 1
        For i = 0 To �����׼� - 1
            With .Range("��꼭1p������").Offset(i + 1, 0)
                .Value = ����list(0, i)

                .Offset(0, 1).Value = ����list(1, i)
                .Offset(0, 2).Value = ����list(2, i)
            End With

        Next i
        Erase ����list
        
        If �����׼� > �����׼� Then
            �׼� = �����׼�
        Else
            �׼� = �����׼�
        End If
        
        ���� = �׼� + ��꼭������ + 1
        
        With .Range("��꼭1p������").Offset(�׼� + 1, 0)
            .Value = "�հ�"
            .Font.Bold = True
            .Offset(0, 1).Formula = "=sum(B7:B" & ���� & ")"
            .Offset(0, 2).Formula = "=sum(C7:C" & ���� & ")"
        End With
        
        With .Range("��꼭1p������").Offset(�׼� + 1, 0)
            .Value = "�հ�"
            .Font.Bold = True
            .Offset(0, 1).Formula = "=sum(E7:E" & ���� & ")"
            .Offset(0, 2).Formula = "=sum(F7:F" & ���� & ")"
        End With
        
    End With
    
    With ws_target.Range("��꼭1p����").CurrentRegion
        .Borders.LineStyle = 1
    End With
    
    ���� = ���� + 1
    With ws_target.Rows("7:" & ����)
        .RowHeight = 30
        .Font.size = 14
    End With
    
    With ws_target.Range("A" & ���� & ":F" & ����)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    With ws_target.Range("D5:D" & ����)
        .Borders(xlEdgeLeft).LineStyle = xlDouble
    End With
    
    Dim startDate, endDate As Date
    startDate = format(get_config("������"), "Short Date")
    endDate = format(get_config("������"), "Short Date")
    ws_target.Range("a3").Value = get_config("�����") & " (" & startDate & " ~ " & endDate & ")"
    
    ws_target.Visible = xlSheetVisible

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub
