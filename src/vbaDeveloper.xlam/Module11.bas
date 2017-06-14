Attribute VB_Name = "Module11"
'module11 : �߰� ��� ���� �Լ��� - get_code, ��굥���Ͱ���(�׽�Ʈ��), �ڵ���,
'  get_config, set_config, ����ǹ����̺����, date_compare, ���嵥����copy, ���嵥��������, init_firstday
Option Explicit

Const ȸ�����_��_���� As String = "a"
Const ȸ�����_��_code As String = "c"
Const ȸ�����_��_�� As String = "d"
Const ȸ�����_��_�� As String = "e"
Const ȸ�����_��_�� As String = "f"
Const ȸ�����_��_���� As String = "g"
Const ȸ�����_��_���� As String = "h"
Const ȸ�����_��_���� As String = "i"
Const ȸ�����_��_���� As String = "j"
Const ȸ�����_��_���� As String = "k"
Const ȸ�����_��_������Ʈ As String = "n"
Const ȸ�����_��_�����ܾ� As String = "P"
Const ȸ�����_��_���ܾ� As String = "R"
Const �ִ���� As Integer = 30000

'get_code �Լ��� ������ �����ϱ� ���� �� ��⿡�� ����ϴ� �������� �迭 ���
Public accountCodes() As Variant
Public codeArrayInitialized As Boolean

Sub home()
' home ��ũ��
' �ٷ� ���� Ű: Ctrl+Shift+H
'
    Sheets("ù������").Activate
End Sub

Function get_code(ByVal �� As String, ByVal �� As String, ByVal �� As String, ByVal ���� As String)
    Dim code As String
    code = ""
    
    'ó�� �����ϴ� �Ÿ� ���������� accountCodes() �� �ʱ�ȭ�Ѵ�.
    If Not codeArrayInitialized Then
        Call init_accountCodes
    End If
    
    Dim i As Integer
    For i = 1 To UBound(accountCodes)
        If accountCodes(i, 2) = �� And accountCodes(i, 3) = �� And accountCodes(i, 4) = �� And accountCodes(i, 5) = ���� Then
            code = accountCodes(i, 1)
        End If
    Next i

    get_code = code
End Function

Sub init_accountCodes()
    Dim ws As Worksheet
    Dim c As Range ', oRng As Range
    Set ws = Worksheets("���꼭")
    Set c = Worksheets("���꼭").Range("���׸��ڵ巹�̺�")
    
    Dim codeCount As Integer
    codeCount = c.End(xlDown).Row - 1
    
    ReDim accountCodes(1 To codeCount, 1 To 5)
    accountCodes = ws.Range(c.Offset(1, 0), c.Offset(codeCount, 4)).Value
    codeArrayInitialized = True
End Sub

Function new_code(ByVal �� As String, ByVal �� As String, ByVal �� As String, ByVal ���� As String, Optional ByVal refresh As Boolean)
    Dim ws As Worksheet
    Set ws = Worksheets("���꼭")

    Dim accountCode As String
    
    If �� = "" Or �� = "" Or �� = "" Or ���� = "" Then
        new_code = ""
    Else
        Dim r������ġ As Range
        Set r������ġ = ws.Range("���׸��ڵ巹�̺�").End(xlDown).Offset(1)
        With r������ġ
            .Offset(0, 1).Value = ��
            .Offset(0, 2).Value = ��
            .Offset(0, 3).Value = ��
            .Offset(0, 4).Value = ����
                
            ws.Range("A4:G" & r������ġ.Row).Sort Key1:=ws.Range("B7"), Order1:=xlAscending, Key2:=ws.Range("C7") _
                , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
                :=xlSortNormal
            
            Call ���꼭�ڵ��Է�
            accountCode = .Value
        End With
        
        new_code = accountCode
    End If
End Function

Sub ��굥���Ͱ���()
    Dim ws_data As Worksheet
    Dim ws_source As Worksheet
    Set ws_data = Worksheets("data")
    Set ws_source = Worksheets("ȸ�����")
    Dim endLine As Integer
    Dim endDataLine As Integer
    
    endLine = ws_source.Range("�����ʵ巹�̺�").End(xlDown).Row
    endDataLine = 0
    Dim i, j As Integer
    
    For i = 6 To endLine
        If (ws_source.Range("b" & j).Value = "") Then
            endDataLine = i - 1
            Exit For
        End If
    Next i

    'Dim ������, ������ As Integer
    
    '������ = 6
    '������ = endDataLine
     
    On Error Resume Next
     
    Dim startDate, endDate, accountStartDate, today As Date
        
    startDate = ws_data.Range("j1").Value
    endDate = ws_data.Range("j2").Value
    accountStartDate = format(get_config("accountStartDate"), "m/d/yyyy")
    today = format(Date, "m/d/yyyy")
    
    
    ' #1. ��ü �Ⱓ �ǹ� ���̺� ����
    Call ����ǹ����̺����(accountStartDate, today)
    
    ' #2. �Ⱓ�� ����
    Call ����ǹ����̺����(startDate, endDate)
    
    ' #3. ���� �� ������. ���� ��¥ �� ǥ��
    With ws_data.Range("i1")
        .Value = "������"
        .Offset(1, 0).Value = "������"
        .Offset(2, 0).Value = "ȸ�������"
        .Offset(3, 0).Value = "����"
        .Offset(0, 1).Value = startDate
        .Offset(1, 1).Value = endDate
        .Offset(2, 1).Value = accountStartDate
        .Offset(3, 1).Value = today

    End With

End Sub

Function �ڵ���(ByVal accountCode As String, ByVal �������� As String, ByVal �κ���ü As String)
    Dim ws_data As Worksheet
    Set ws_data = Worksheets("data")
    Dim c As Range
    Dim ���̵� As Integer
    
    If �������� = "����" Then
        ���̵� = 1
    Else
        ���̵� = 2
    End If

    If �κ���ü = "�κ�" Then
        With ws_data.Range("e1").CurrentRegion.columns(1)
            For Each c In .Cells
                If c.Value = accountCode Then
                    �ڵ��� = c.Offset(0, ���̵�).Value
                    Exit For
                End If
            Next c
        End With
        If Not �ڵ��� > 0 Then
            �ڵ��� = 0
        End If
    Else
        With ws_data.Range("a1").CurrentRegion.columns(1)
            For Each c In .Cells
                If c.Value = accountCode Then
                    �ڵ��� = c.Offset(0, ���̵�).Value
                    Exit For
                End If
            Next c
        End With
        If Not �ڵ��� > 0 Then
            �ڵ��� = 0
        End If
    End If

End Function

Public Function get_config(ByVal item As String)
    With Worksheets("����")
        If item = "������" Then
            get_config = .Range("�۾������ϼ���").Offset(0, 1).Value
        ElseIf item = "������" Then
            get_config = .Range("�۾������ϼ���").Offset(0, 1).Value
        ElseIf item = "ȸ�������" Then
            get_config = .Range("ȸ������ϼ���").Offset(0, 1).Value
        ElseIf item = "�����" Then
            get_config = .Range("�������").Offset(0, 1).Value
        Else
            get_config = ""
        End If
    End With
End Function

Public Sub set_config(ByVal item As String, ByVal newValue As String)
    With Worksheets("����")
        If item = "������" Then
            .Range("�۾������ϼ���").Offset(0, 1).Value = newValue
        ElseIf item = "������" Then
            .Range("�۾������ϼ���").Offset(0, 1).Value = newValue
        ElseIf item = "ȸ�������" Then
            .Range("ȸ������ϼ���").Offset(1, 0).Value = newValue
        ElseIf item = "�����" Then
            .Range("�������").Offset(0, 1).Value = newValue
        Else

        End If
    End With
End Sub

Sub ����ǹ����̺����(ByVal startDate As Date, ByVal endDate As Date)
    Application.ScreenUpdating = False

    Dim ws_source As Worksheet, ws_source_copy As Worksheet, ws_data As Worksheet
    Dim rng_source As Range
    Dim pvt As PivotTable
    Dim pvtName As String, ������ġ As String
    Dim startRow As Integer, endRow As Integer
    Dim endLine As Integer, endDataLine As Integer
    Dim i As Integer
    Dim dataCount As Integer
    
    Dim accountStartDate As Date, today As Date
    accountStartDate = get_config("ȸ�������")
    today = Date
            
    Set ws_source = Worksheets("ȸ�����")
    Set ws_source_copy = Worksheets("�׸�а���")
    Set ws_data = Worksheets("data")
    
    endLine = ws_source.Range("�����ʵ巹�̺�").End(xlDown).Row
    endDataLine = 0
    i = 0

    For i = 6 To endLine
        If (ws_source.Range("a" & i + 1).Value = "") Then
            endDataLine = i
            Exit For
        End If
    Next i
    
    startRow = 6
    endRow = endDataLine
    dataCount = endDataLine - 5
    
    Dim ��¥�� As Integer
    Dim ��üor�κ� As String
    Dim c As Range
    Set c = ws_source.Range("A5")
    
    ws_source.Unprotect PWD
    'On Error GoTo ErrorHandler
    
    '������� �������� ��Ȯ�� �����ϴ� ���� ����. �Ҿ����� �� ������ �����.
    '�ӵ� ������ ��Ȯ�� �Ǵ��� ���� ���� �ϳ��� Ȯ���ϴ� ���� �ƴ϶� �迭�� �̿�
    Dim dateArray() As Variant
    ReDim dateArray(1 To dataCount)
    
    'dateArray = ws_source.Range("A6:A" & endDataLine).Value '�� ����� �ȵȴ�.
    'dateArray = ws_source.Range(c, c.Offset(dataCount - 1, 0)).Value '�̰͵� get_code�� �޸� �ȵ�
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
        If dateArray(i) = endDate Then
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
    
    If startDate = accountStartDate And endDate = today Then
        ws_data.Range("L2:t" & �ִ����).ClearContents
        Call ���嵥����copy("data", 6, endDataLine)
        Set rng_source = ws_data.Range("L1").Resize(endDataLine - 5 + 1, 9) '�� �󺧵� ���ԵǹǷ� + 1
        
        pvtName = "���Ⱓ���"
        ������ġ = "data!R1C1"
    Else  'Ư�� �Ⱓ�� ������ ��� �ش� �Ⱓ�� ���� ��� ����
        Dim ���ڵ�� As Integer
        ���ڵ�� = endRow - startRow + 1
        
        ws_data.Range("L2:t" & �ִ����).ClearContents
        Call ���嵥����copy("data", startRow, endRow)
        Set rng_source = ws_data.Range("L1").Resize(���ڵ�� + 1, 9) '�� �󺧵� ���ԵǹǷ� + 1
        pvtName = "Ư���Ⱓ���"
        ������ġ = "data!R1C5"
    End If
    
    Dim prevStartDate As String
    Dim prevEndDAte As String
    prevStartDate = ws_data.Range("��������������").Value
    prevEndDAte = ws_data.Range("��������������").Value
        
    If startDate <> prevStartDate Or endDate <> prevEndDAte Then  '���� ��ȸ�Ҷ��� �ٸ� �Ⱓ�� �����ߴٸ� ���̺� �����
        
        '���� �ǹ����̺� ����
        For Each pvt In ws_data.PivotTables
            If Not pvt Is Nothing Then
                If pvt.name = pvtName Then
                    pvt.TableRange2.Clear
                    Exit For
                End If
            End If
        Next
        
        '�ǹ����̺� ����
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            rng_source).CreatePivotTable _
            TableDestination:=������ġ, TableName:=pvtName
        
        With ws_data.PivotTables(pvtName)
            .PivotFields("code").Orientation = xlRowField
            .PivotFields("code").position = 1
            .AddDataField .PivotFields("����"), "�հ�: ����", xlSum 'xlCount
            .AddDataField .PivotFields("����"), "�հ�: ����", xlSum
        End With
        
    Else  '���� ��ȸ�� ���� ���� �Ⱓ�� �����ߴٸ� ���� refresh
        ws_data.PivotTables(pvtName).PivotCache.refresh
        
    End If
    
    ' ������ ���� �����ϰ� ������ ���� (����������, ���������Ϸ� Ȱ��)
    ws_data.Range("��������������").Value = startDate
    ws_data.Range("��������������").Value = endDate
    
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MsgBox "������ ������  " & vbCrLf & Err.Number & vbCrLf & Err.Description & Space(10), vbCritical, "���� ���� Ȯ��"
'    End If
    
End Sub

Function date_compare(ByVal startDate As Date, ByVal endDate As Date)
    'datediff �Լ��� �� ��¥���� �� ��¥�� ����. �� ����� ����̸� �� ��¥�� ������, �����̸� �� ��¥�� ������
    date_compare = DateDiff("d", startDate, endDate)
End Function

Sub init_firstday()
    'ȸ������� �ʱ�ȭ
    Dim accountStartDate As String
    accountStartDate = Year(Now) & "-01-01"
    Worksheets("����").Range("ȸ������ϼ���").Offset(0, 1).Value = accountStartDate
    UserForm_����.TextBox_ȸ�������.Value = accountStartDate
End Sub

Sub ���嵥����copy(ByVal targetSheet As String, ByVal startRow As Integer, ByVal endRow As Integer, Optional ByVal tag As String)
    Application.ScreenUpdating = False
    
    Dim ws_source As Worksheet
    Dim ws_data As Worksheet
    Set ws_source = Worksheets("ȸ�����")
    Set ws_data = Worksheets(targetSheet)
    
    Dim copy_point As Range
    
    If targetSheet = "data" Then
        Set copy_point = ws_data.Range("L2")
    ElseIf targetSheet = "�׸�а���" Then
        Set copy_point = ws_data.Range("a6")
    End If
    'MsgBox "test" & targetSheet & " startRow: " & startRow & " endRow: " & endRow
    If project <> "" Then
        tag = project
    End If
    
    If tag <> "" Then
        Dim ��ü As Range
        Set ��ü = ws_source.Range("A" & startRow & ":N" & endRow)
        Dim i As Integer
        Dim Ÿ���� As Integer
        Ÿ���� = copy_point.Row
        Dim ���� As String, code As String, �� As String, �� As String, �� As String, ���� As String, ���� As String
        Dim ���� As Long, ���� As Long
        
        For i = startRow To endRow
            If ws_source.Range(ȸ�����_��_������Ʈ & i).Value = tag Then
                With ws_source
                    ���� = .Range(ȸ�����_��_���� & i).Value
                    code = .Range(ȸ�����_��_code & i).Value
                    �� = .Range(ȸ�����_��_�� & i).Value
                    �� = .Range(ȸ�����_��_�� & i).Value
                    �� = .Range(ȸ�����_��_�� & i).Value
                    ���� = .Range(ȸ�����_��_���� & i).Value
                    ���� = .Range(ȸ�����_��_���� & i).Value
                    ���� = .Range(ȸ�����_��_���� & i).Value
                    ���� = .Range(ȸ�����_��_���� & i).Value
                End With
                    
                With copy_point
                    .Value = ����
                    .Offset(0, 1).Value = code
                    .Offset(0, 2).Value = ��
                    .Offset(0, 3).Value = ��
                    .Offset(0, 4).Value = ��
                    .Offset(0, 5).Value = ����
                    .Offset(0, 6).Value = ����
                    .Offset(0, 7).Value = ����
                    .Offset(0, 8).Value = ����
                End With
                Set copy_point = copy_point.Offset(1, 0)
            End If
        Next i
    
    Else
        '�ӵ� ������ ���� copy&paste ��Ŀ��� range �� ���� ������ ���� ������� �ٲ� 2016.9.18
        'ws_source.Range("A" & startRow & ":a" & endRow).Copy
        'copy_point.PasteSpecial xlPasteValues
        If targetSheet = "data" Then
            ws_data.Range("L1").Value = "��¥"
            ws_data.Range("M1").Value = "code"
            ws_data.Range("N1").Value = "��"
            ws_data.Range("O1").Value = "��"
            ws_data.Range("P1").Value = "��"
            ws_data.Range("Q1").Value = "����"
            ws_data.Range("R1").Value = "����"
            ws_data.Range("S1").Value = "����"
            ws_data.Range("T1").Value = "����"
            'data ��Ʈ �ʵ� �̸� �κ� �ʱ�ȭ �ڵ� �߰� - 2016.11.9
            
            ws_data.Range("L2:L" & endRow - startRow + 2).Value = ws_source.Range("A" & startRow & ":a" & endRow).Value
            ws_data.Range("M2:T" & endRow - startRow + 2).Value = ws_source.Range("C" & startRow & ":" & ȸ�����_��_���� & endRow).Value
        ElseIf targetSheet = "�׸�а���" Then
            ws_data.Range("A6:A" & endRow - startRow + 6).Value = ws_source.Range("A" & startRow & ":a" & endRow).Value
            ws_data.Range("B6:I" & endRow - startRow + 6).Value = ws_source.Range("C" & startRow & ":" & ȸ�����_��_���� & endRow).Value
        End If
        
        'ws_source.Range("c" & startRow & ":" & ȸ�����_��_���� & endRow).Copy
        'copy_point.Offset(0, 1).PasteSpecial xlPasteValues

        Application.CutCopyMode = False
    End If
End Sub

Sub ��������������(���������Ʈ As String, �з� As String)
    Dim ws_source As Worksheet
    Dim ws_target As Worksheet
    Set ws_source = Worksheets(���������Ʈ)
    Set ws_target = Worksheets("���꼭")
    '���꼭 �ʱ�ȭ
    Call UserForm_����.ȸ�輳���ʱ�ȭ("��������")
    
    '���������� �������� �����ͼ� ���꼭 ��Ʈ�� �ִ� �κ�
    Dim ��ü As Range
    Dim x As Integer, y As Integer
    Dim ���ؿ� As Range, ���� As Range, �׿� As Range, �� As Range, ���� As Range
    
    Set ���� = ws_source.Range("���ð�����")
    Set �׿� = ws_source.Range("�����׿���")
    Set �� = ws_source.Range("���ø񿭶�")
    Set ���� = ws_source.Range("���ü��񿭶�")
    Set ���ؿ� = ws_source.Range("���úз�����")
    x = 0

    Set ��ü = ws_source.Range("���ð�����").CurrentRegion.columns(���ؿ�.Column)
    Dim rowCount As Integer
    rowCount = ��ü.Rows.Count
    Dim Ÿ���� As Integer
    Ÿ���� = ws_target.Range("A1").End(xlDown).Offset(1, 0).Row
    Dim �з��� As String
    Dim i As Integer
    Dim �� As String, �� As String, �� As String, ���� As String
    
    For i = 1 To rowCount
        �з��� = ws_source.Range("A" & i).Offset(, ���ؿ�.Column - 1).Value
        If �з��� = "����" Or �з��� = �з� Then
            With ws_source.Range("A" & i)
                �� = .Offset(, ����.Column - 1).Value
                �� = .Offset(, �׿�.Column - 1).Value
                �� = .Offset(, ��.Column - 1).Value
                ���� = .Offset(, ����.Column - 1).Value
            End With
            
            With ws_target.Range("A" & Ÿ����)
                .Offset(, 1).Value = ��
                .Offset(, 2).Value = ��
                .Offset(, 3).Value = ��
                .Offset(, 4).Value = ����
            End With
            Ÿ���� = Ÿ���� + 1
        End If
    Next i
    
    '������
    code_changed = True
    Call ��������_��Ʈ����
    
    Call ���꼭�ڵ��Է�
    Call ��꼭�ʱ�ȭ
End Sub

Sub ������������()
    Dim ws_source As Worksheet
    Dim ws_target As Worksheet
    Set ws_source = Worksheets("��������2")
    Set ws_target = Worksheets("���꼭")
    Dim start As Range, �� As Range
    Dim columnCount As Integer, rowCount As Integer
    Dim �� As String, �� As String, �� As String, ���� As String, accountCode As String
    Dim x(4, 2) As Variant
    Dim ���ÿ� As Integer
    Dim i As Integer
    i = 0

    On Error GoTo error

    With ws_source
        Set start = .Range("a2")
        columnCount = start.CurrentRegion.columns.Count
        rowCount = start.CurrentRegion.Rows.Count - 2
        
        If rowCount = 0 Then
            MsgBox "����° �ٿ� �����͸� �������ּ���"
            ws_source.Range("A3").Select
            Exit Sub
        End If
        
        Set �� = start.Offset(0, columnCount - 1)
        Dim label As Range
        For Each label In Range(start, ��)
            If label.Value <> "" Then
                x(i, 0) = label.Value
                x(i, 1) = label.Column
                i = i + 1
                If label.Value = "����" Then
                    ���ÿ� = label.Column
                End If
            End If
        Next label
    End With

    If x(3, 0) = "" Then
        MsgBox "���� �����մϴ�. ��, ��, ��, ������� 4���� ���� �ʿ��մϴ�"
        start.Select
        Exit Sub
    Else
        Dim �󺧰� As String
        Dim �� As Integer
        Dim Ÿ���� As Integer
        Ÿ���� = ws_target.Range("���׸��ڵ巹�̺�").End(xlDown).Row + 1
        Dim Ÿ�ٿ� As Integer
        Dim ù��, ���� As Range
        Dim Ÿ�ٽ���, Ÿ�ٳ� As Range
        Dim j As Integer
        Dim codeCount As Integer
        codeCount = 0
        
        Application.DisplayStatusBar = True
        
        ' #5 �� �� ��ȸ�ϸ� ���� & ȸ����忡 ���̱�
        For i = 0 To UBound(x)
            �󺧰� = x(i, 0)
            �� = x(i, 1)

            If Not �� > 0 Then
                Exit For
            End If
            
            With start.Offset(0, �� - 1)
                Set ù�� = .Offset(1, 0)
                Set ���� = .Offset(rowCount, 0)

                With ws_target.Range("���׸��ڵ巹�̺�")
                    For j = 1 To 5
                        If �󺧰� = .Offset(0, j).Value Then
  
                            With ws_target.Range("A" & Ÿ����)
                                Set Ÿ�ٽ��� = .Offset(0, j)
                                Set Ÿ�ٳ� = .Offset(rowCount - 1, j)
                            End With
                            
                            If Not IsEmpty(ws_source.Range(ù��, ����).Value) Then
                                ws_target.Range(Ÿ�ٽ���, Ÿ�ٳ�).Value = ws_source.Range(ù��, ����).Value
                                If codeCount = 0 Then
                                    codeCount = rowCount - Application.WorksheetFunction.CountIf(Range(ù��, ����), "")
                                End If
                            End If
                            
                            Exit For
                        End If
                    Next j
                End With
            End With
        Next i
        
        If codeCount > 0 Then
            MsgBox codeCount & "���� ���׸��� ���꼭�� ����Ǿ����ϴ�."
            ws_target.Activate
            ' # �����ϰ�, �ڵ� �ο��� ��, ��꼭 �ʱ�ȭ
            Application.StatusBar = "������ �������� ���� �ڵ带 �ο��ϰ� �ֽ��ϴ�"
            Call ��������_��Ʈ����
            Call ���꼭�ڵ��Է�
            Application.StatusBar = "Ȯ���� �������� ���� ��꼭�� �ʱ�ȭ�ϰ� �ֽ��ϴ�"
            Call ��꼭�ʱ�ȭ
        Else
            MsgBox "�Էµ��� �ʾҽ��ϴ�."
        End If
    End If
    
    Erase x
    Application.DisplayStatusBar = False
    
error:
    If Err.Number <> 0 Then
        MsgBox "������ȣ : " & Err.Number & vbCr & _
        "�������� : " & Err.Description, vbCritical, "����"
    End If
    
End Sub

Sub ���嵥��������()
    ' #4 ù �࿡ ��� �� �ִ��� Ȯ��
    Dim ws_source, ws_target As Worksheet
    Set ws_source = Worksheets("��������")
    Set ws_target = Worksheets("ȸ�����")
    Dim start, �� As Range
    Dim columnCount As Integer, rowCount As Integer
    Dim �� As String, �� As String, �� As String, ���� As String
    Dim ���� As Integer '�� �����Ҷ�, ���� ����ܼ��� Ȥ�� ������������� Ȯ���ϱ� ���� Ư���� �ʿ�
    Dim accountCode As String
    Dim x(8, 2) As Variant
    Dim i, j As Integer
    i = 0
    
    ws_target.Unprotect PWD
    
    With ws_source
        Set start = .Range("a2") 'a1���� ���� ������ ���� �ִ�
        columnCount = start.CurrentRegion.columns.Count
        rowCount = start.CurrentRegion.Rows.Count - 2 'ù��(���� ����) ����
        
        If rowCount = 0 Then
            MsgBox "����° �ٿ� �����͸� �������ּ���"
            ws_source.Range("A3").Select
            Exit Sub
        End If
        
        Set �� = start.Offset(0, columnCount - 1)
        Dim �� As Range
        For Each �� In Range(start, ��)
            If i > 8 Then
                Exit For
            End If
            If ��.Value <> "" Then
                x(i, 0) = ��.Value
                x(i, 1) = ��.Column
                i = i + 1
                If ��.Value = "��" Then
                    ���� = ��.Column
                End If
            End If
        Next ��
    End With
    
    If x(8, 0) = "" Then
        MsgBox "���� �����մϴ�. ����, ��, ��, ��, ����, ����, ����, ����, ��/������ 9���� ���� �ʿ��մϴ�"
        start.Select
        Exit Sub
    Else
        Dim �󺧰� As String
        Dim �� As Integer
        Dim Ÿ����, Ÿ�ٿ� As Integer
        Ÿ���� = ws_target.Range("�����ʵ巹�̺�").End(xlDown).Row + 1
        Dim ù��, ���� As Range
        Dim Ÿ�ٽ���, Ÿ�ٳ� As Range
        Dim ������ As Range
        
        ' #4-1. ������ �����ϱ�
        For i = 0 To UBound(x)
            �󺧰� = x(i, 0)
            �� = x(i, 1)
            ws_source.Range("f1").Value = ��
            If Not �� > 0 Then '��� ���� �� ���Ұų� �ʱ�ȭ�� �ȵ� ��� ����
                Exit For
            End If
            
            If Not IsError(Application.Match(�󺧰�, Array("��", "��", "��", "����"), False)) Then
                With start.Offset(0, �� - 1)
                    For j = 1 To rowCount
                        If IsEmpty(.Offset(j, 0).Value) Then
                            Select Case start.Offset(j, ���� - 1).Value
                                Case "����ܼ���", "���������"
                                    'MsgBox start.Offset(j, ���� - 1).Value
                                    'Exit Sub
                                    'pass
                                Case Else
                                    MsgBox "������ ���� �־� ������������ ������ �� �����ϴ�" & rowCount
                                    .Offset(j, 0).Select
                                    Exit Sub
                            End Select
                        End If
                    Next j
                End With
            End If
        Next i
        
        MsgBox "������ ������ ���ƽ��ϴ�. Ȯ���� ������ ȸ����忡 �����մϴ�"
        
        Application.DisplayStatusBar = True
        Application.StatusBar = "�����͸� ȸ����忡 �����ϰ� �ֽ��ϴ�"
        
        ws_target.Unprotect PWD
        
        ' #5 �� �� ��ȸ�ϸ� ���� & ȸ����忡 ���̱�
        For i = 0 To UBound(x)
            �󺧰� = x(i, 0)
            �� = x(i, 1)
            ws_source.Range("f1").Value = ��
            If Not �� > 0 Then '��� ���� �� ���Ұų� �ʱ�ȭ�� �ȵ� ��� ����
                Exit For
            End If
            
            With start.Offset(0, �� - 1)
                Set ù�� = .Offset(1, 0)
                Set ���� = .Offset(rowCount, 0)
                Set ������ = ws_target.Range("�����ʵ巹�̺�")
                
                For j = 0 To 17 'columnCount?
                    If �󺧰� = ������.Offset(0, j).Value Then
                        With ws_target.Range("A" & Ÿ����)  'ȸ������� ���� ������ �ٿ� �߰�
                            Set Ÿ�ٽ��� = .Offset(0, j)
                            Set Ÿ�ٳ� = .Offset(rowCount - 1, j)
                        End With
                        ws_target.Range(Ÿ�ٽ���, Ÿ�ٳ�).Value = ws_source.Range(ù��, ����).Value
                        Exit For
                    End If
                Next j

            End With
        Next i
        Application.DisplayStatusBar = False

        ws_target.Activate
    
    End If
    
    Erase x
    
    Application.StatusBar = "���ο� ���׸��� �����ϰ� �ֽ��ϴ�"
    
    '#6 �߰��� �κ� ��ȸ�ϸ� ���׸� ���� : ������ ��ϵ� �ڵ����� Ȯ��
    Dim r������ġ As Range
    Dim ws_���꼭 As Worksheet
    Set ws_���꼭 = Worksheets("���꼭")
    Dim k As Integer
    
    '#6-1. �� ���׸��� �ִ��� Ȯ��
    '#6-2. �� ���׸� ����
    'ȸ����� ��Ʈ�� ������ ������ ���׸���� ���꼭 ��Ʈ�� ���� -> �ߺ��� ���� -> ���� -> �ڵ� ��ο�
    Dim ���꼭_������ As Integer
    ���꼭_������ = ws_���꼭.Range("���ʵ�").End(xlDown).Row + 1
    
    Dim ����_������ġ As Range
    Set ����_������ġ = ws_���꼭.Range("B" & ���꼭_������ & ":E" & ���꼭_������ + rowCount - 1)
    ����_������ġ.Value = ws_target.Range("D" & Ÿ���� & ":G" & Ÿ���� + rowCount - 1).Value
    Set ����_������ġ = ws_���꼭.Range("B2:E" & ���꼭_������ + rowCount)
    ����_������ġ.RemoveDuplicates columns:=Array(1, 2, 3, 4), Header:=xlNo
    
    ws_���꼭.Range("A4:G" & Cells(Rows.Count, "A").End(xlDown).Row).Sort Key1:=ws_���꼭.Range("B7"), Order1:=xlAscending, Key2:=ws_���꼭.Range("C7") _
                , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
                :=xlSortNormal
                
    Set ����_������ġ = Nothing
    Application.CutCopyMode = False
            
    Call ���꼭�ڵ��Է�
    
    '#6-3. ȸ����忡 ���׸� �Է�
    '�� �κ��� ����. �迭�� ���� ������ ��ȸ�ϰ� ����
    Application.StatusBar = "ȸ����忡 �� ���׸��� �����ϰ� �ֽ��ϴ�"
    ws_target.Unprotect PWD
    codeArrayInitialized = False 'get_code �Լ��� ���ο� �������� �ڵ带 ���������� �ٽ� �ʱ�ȭ�ؾ� �Ѵٰ� �˸�
    
    Dim subject() As Variant
    ReDim subject(1 To rowCount, 1 To 4)
    Dim newCode() As Variant
    ReDim newCode(1 To rowCount, 1 To 2)
    
    Dim rNewTarget As Range
    Set rNewTarget = ws_target.Range("D" & Ÿ����)
    
    subject = ws_target.Range(rNewTarget, rNewTarget.Offset(rowCount, 3)).Value
    Dim point As Integer
    point = rowCount / 10
    
    Application.StatusBar = "ȸ����忡 �� ���׸��� �����ϰ� �ֽ��ϴ�"
    With ws_target
        For k = 1 To rowCount
            'If k Mod point = 0 Then
            '    'Application.StatusBar = "ȸ����忡 �� ���׸��� �����ϰ� �ֽ��ϴ� (" & k & " / " & rowCount & ")"
            '    Application.StatusBar = "ȸ����忡 �� ���׸��� �����ϰ� �ֽ��ϴ� (" & k * 100 / rowCount & "%)"
            'End If
            �� = subject(k, 1)
            �� = subject(k, 2)
            �� = subject(k, 3)
            ���� = subject(k, 4)
            
            accountCode = get_code(��, ��, ��, ����)
            
            newCode(k, 1) = accountCode & "/" & �� & "/" & �� & "/" & �� & "/" & ����
            newCode(k, 2) = accountCode
            
            'With .Range("A" & Ÿ���� + k - 1)
            '    If accountCode <> "" Then
            '        .Offset(0, 1).Value = accountCode & "/" & �� & "/" & �� & "/" & �� & "/" & ����
            '        '.Offset(0, 2).FillDown
            '        '.Offset(0, 2).NumberFormat = "@" '�� �κ��� ���� �ſ���! �Ʒ����� �ϰ��ǰ� �����ϰ� ���� 2016.9.20
            '        .Offset(0, 2).Value = accountCode
            '    End If
            'End With
        Next k
        
        .Range("B" & Ÿ����, "C" & Ÿ���� + rowCount - 1).Value = newCode
        .Range("C" & Ÿ���� & ":C" & Ÿ���� + rowCount - 1).NumberFormat = "@"
    End With
    
    ' #7 ȸ����� ����
    Application.StatusBar = "ȸ����� ������ �������ϰ� �ֽ��ϴ�"
    Dim endLine As Integer
    endLine = ws_target.Range("�����ʵ巹�̺�").End(xlDown).Row

    ws_target.Range("A8:O" & endLine).Sort Key1:=ws_target.Range("A7"), Order1:=xlAscending, Key2:=ws_target.Range("B7") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

    ' #8 ����
    ' ��/���� ��� ������ 0���� ä��� (���� �ŷ� �켱)
    With ws_target
        For k = 8 To endLine
            If IsEmpty(.Range(ȸ�����_��_���� & k).Value) Then
                .Range(ȸ�����_��_���� & k).Value = 0
            End If
        Next k
    End With
    
    If Worksheets("����").Range("a2").Offset(, 1).Value = True Then
        ws_target.Protect PWD
    End If
    ws_target.Activate
    MsgBox rowCount & "���� �����Ͱ� ���յǾ����ϴ�."
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    
error:
    If Err.Number <> 0 Then

        MsgBox "������ȣ : " & Err.Number & vbCr & _
        "�������� : " & Err.Description, vbCritical, "����"

    End If
End Sub

Sub ���嵥���Ͱ�������_�غ�()
    Dim ws As Worksheet
    Set ws = Worksheets("��������")
    Dim ������ As Range
    Set ������ = ws.Range("A2")
    
    Dim ��¥1 As String
    Dim ��¥2 As Date
    
    '�츮���� ���������� Ȯ��
    'A3���� No. �ŷ��Ͻ� ���� ���系�� ã���űݾ�(Ȥ�� ����(��)), �ñ�űݾ�(Ȥ�� ), �ŷ����ܾ�, �����
    Dim ���� As Integer
    ���� = 0
    If ������.Value = "No." Then
        ���� = 2
    ElseIf ������.Offset(1, 0).Value = "No." Then
        ���� = 3
    End If
    
    If Not ���� > 0 Then
        MsgBox "�츮���� �ڷḦ A3 ��ġ�� �ٽ� �������ּ���"
        Exit Sub
    End If
    
    '2��° �� (A2~) �����ϰ� ����ø�
    If ���� = 3 Then
        ������.EntireRow.Delete Shift:=xlUp
        Set ������ = ws.Range("a2")
    End If
    
    If ������.Offset(1, 0).Value = "" Then
        MsgBox "�����Ͱ� �����ϴ�"
        Exit Sub
    End If
    
    'No. �����ϰ� shift left, ���� ���� ��ȯ
    Dim startRow As Integer, endRow As Integer
    startRow = 2
    endRow = ������.End(xlDown).Row
    
    ws.Range("a" & startRow, "a" & endRow).Select
    Selection.Delete Shift:=xlToLeft
    
    '�Ͻ� �����ʿ� 4�� �߰� & ��, ��, ��, ���� �� ǥ��
    With ws.Range("B2:F" & endRow)
        .Insert Shift:=xlToRight
        .NumberFormatLocal = "@"
    End With
    
    Set ������ = ws.Range("A2")
    
    With ������
        .Value = "����"
        .Offset(, 1).Value = "��"
        .Offset(, 2).Value = "��"
        .Offset(, 3).Value = "��"
        .Offset(, 4).Value = "����"
        .Offset(, 5).Value = "��/��"
    End With
    
    '���� �����ʿ� "��/��" �� �߰�. �� �Ʒ� ��� 0���� ä��
    ws.Range("F3").Value = 0
    ws.Range("F3:F" & endRow).FillDown
    columns("F:F").EntireColumn.AutoFit
    
    '��¥ ���� ���� (����)
    Dim i As Integer
    Dim rŸ����ġ As Range
    For i = startRow + 1 To endRow
        Set rŸ����ġ = ws.Range("A" & i)
        With rŸ����ġ
            .Value = Replace(Left(.Value, 10), ".", "-")
        End With
    Next i
    columns("A:A").EntireColumn.AutoFit
    
    Dim ���Կ�, ���⿭ As Integer
    
    '���� -> ����, ���系�� -> ����� �󺧺���
    'ã���űݾ�(����(��)) -> ���� / �ñ�űݾ� -> ����
    ws.Range(������, ������.End(xlToRight)).Select
    
    Selection.Replace What:="����", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="���系��", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="����(��)", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="ã���űݾ�", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="�Ա�(��)", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="�ñ�űݾ�", Replacement:="����", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Find(What:="����", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
    ���Կ� = ActiveCell.Column
    
    '�� ������ ���Աݾ��̸� '����', ����ݾ��̸� '����'�� ǥ��
    For i = startRow + 1 To endRow
        With ws.Range("B" & i)
            If .Offset(, ���Կ� - 2).Value > 0 Then
                .Value = "����"
            Else
                .Value = "����"
            End If
        End With
    Next i

End Sub
