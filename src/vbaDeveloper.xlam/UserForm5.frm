VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "��꼭����"
   ClientHeight    =   8380.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895.001
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_�Ⱓ���_Click()
    Dim startDate As Date, endDate As Date
    Dim project As String
    
    If TextBox_������.Value = "" Then
        MsgBox "�������� �Է����ּ���"
        Exit Sub
    End If
    
    If TextBox_������.Value = "" Then
        MsgBox "�������� �Է����ּ���"
        Exit Sub
    End If
    
    If Not IsDate(TextBox_������.Value) Then
        MsgBox "�������� �߸� �ԷµǾ����ϴ�. (��: 28�ϸ� �ִ� �޿� 29���� �Է�)"
        Exit Sub
    End If
    If Not IsDate(TextBox_������.Value) Then
        MsgBox "�������� �߸� �ԷµǾ����ϴ�. (��: 30�ϸ� �ִ� �޿� 31���� �Է�)"
        Exit Sub
    End If
    
    startDate = IIf(TextBox_������.Value <> "", TextBox_������.Value, get_config("ȸ�������"))
    endDate = IIf(TextBox_������.Value <> "", TextBox_������.Value, Date)
    
    If date_compare(startDate, endDate) < 0 Then
        MsgBox "�������� �����Ϻ��� �ռ��ϴ�. �������� �ٽ� �������ּ���"
        Exit Sub
    End If
    
    With Worksheets("����")
        .Activate
        .Range("�۾������ϼ���").Offset(0, 1).Value = startDate
        .Range("�۾������ϼ���").Offset(0, 1).Value = endDate
    End With

    'project ������ module12�� ���ǵ� ��������
    If ComboBox_������Ʈ.Value <> "" Then
        project = ComboBox_������Ʈ.Value
    Else
        project = ""
    End If
    
    Application.DisplayStatusBar = True
    
    rebuild_report = CheckBox_default.Value
    If rebuild_report Then
        Application.StatusBar = "��꼭�� �ʱ�ȭ�ϰ� �ֽ��ϴ�."
        Call ��꼭�ʱ�ȭ2
        Application.StatusBar = "��꼭�� �ʱ�ȭ�Ǿ����ϴ�."
    End If
    
    Call �׸����ۼ�(CheckBox_�׸�а���.Value, project) ' module 12

    If CheckBox_1page.Value = True Then
        report_1p = True
        
        Call ���1p
    End If
    
    Application.StatusBar = "��꼭 ������ �Ϸ�Ǿ����ϴ�"
    Unload Me
End Sub

Private Sub btn_�Ⱓ���_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�����ʿ� ������ �Ⱓ�� �ش��ϴ� ��꼭�� �����մϴ�"
End Sub

Private Sub btn_�ϰ�ǥ_Click()
    Call ��¥����("�ϰ�ǥ")
    Call �ϰ�ǥ�ۼ�
    Unload Me
End Sub

Sub ��¥����(paper As String)
    Dim ws As Worksheet
    Set ws = Worksheets("ȸ�����")
    ws.Activate
    Dim curRange As Range
    Dim i As Integer
    Dim ���� As Integer
    ���� = ws.Range("A6").End(xlDown).Row
    Dim ��¥���� As Integer
    ��¥���� = 0
    Dim startDate As Date, endDate As Date
    startDate = IIf(TextBox_������.Value <> "", TextBox_������.Value, get_config("ȸ�������"))
    endDate = IIf(TextBox_������.Value <> "", TextBox_������.Value, Date)
    
    Dim r���� As Range
    '��¥�� ��Ʈ�� ��� ǥ�õǴ�, find�� "m/d/yyyy" ���·� �˻��ؾ� ã������.
    'find �μ��� format( ,"m/d/yyyy") �� ����
    ������_�� = format(endDate, "yyyy")
    ������_�� = format(endDate, "m")
    ������_�� = format(endDate, "d")
    Set r���� = ws.Range("�����ʵ巹�̺�").CurrentRegion.columns(1).Find(������_�� & "/" & ������_�� & "/" & ������_��)
    
    If r���� Is Nothing Then
        MsgBox "������ ��¥(������) �ڷḦ ã�� ���߽��ϴ�. �ֱٿ� �Է��� ������ ���ϴ�"
        Set curRange = ws.Range("�����ʵ巹�̺�").End(xlDown)
        i = ����
        
        If paper = "�Աݿ���" Then

            Do While i > 6
                If ws.Range("A" & i).Offset(, 3).Value = "����" Then
                    ws.Range("A" & i).Select
                    Exit Do
                End If
                i = i - 1

            Loop
            
        ElseIf paper = "��ݿ���" Then

            Do While i > 6
                If ws.Range("A" & i).Offset(, 3).Value = "����" Then
                    ws.Range("A" & i).Select
                    Exit Do
                End If
                i = i - 1
            Loop

        Else '�ϰ�ǥ
            curRange.Select
        End If
        
    Else
        With Worksheets("����")
            .Activate
            .Range("�۾������ϼ���").Offset(0, 1).Value = startDate
            .Range("�۾������ϼ���").Offset(0, 1).Value = endDate
        End With
        ws.Activate
        
        If paper = "�Աݿ���" Then
            If r����.Offset(, 3).Value <> "����" Then
                For i = r����.Row To ����
                    If ws.Range("A" & i).Value <> CDate(������) Then
                        Exit For
                    End If
                    If ws.Range("A" & i).Offset(, 3).Value = "����" Then
                        ws.Range("A" & i).Select
                        ��¥���� = 1
                        Exit For
                    End If
                Next i
                
                If Not ��¥���� > 0 Then
                    MsgBox "�ش� ��¥�� ���� ����� �����ϴ�. �ٸ� ��¥�� �������ּ���"
                    ws.Range("�����ʵ巹�̺�").Select
                End If
            Else
                r����.Select
            End If
            
        ElseIf paper = "��ݿ���" Then
            If r����.Offset(, 3).Value <> "����" Then
                
                For i = r����.Row To ����
                    If ws.Range("A" & i).Value <> CDate(������) Then
                        Exit For
                    End If
                    If ws.Range("A" & i).Offset(, 3).Value = "����" Then
                        ws.Range("A" & i).Select
                        ��¥���� = 1
                        Exit For
                    End If
                Next i
                
                If Not ��¥���� > 0 Then
                    MsgBox "�ش� ��¥�� ���� ����� �����ϴ�. �ٸ� ��¥�� �������ּ���"
                    ws.Range("�����ʵ巹�̺�").Select
                End If
            Else
                r����.Select
            End If
        Else '�ϰ�ǥ
            r����.Select 'curRange.Select
        End If
        
    End If

End Sub

Private Sub btn_�ϰ�ǥ_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�����Ϸ� ������ ��¥�� �ϰ�ǥ�� �����մϴ�"
End Sub

Private Sub CommandButton_endday_today_Click()
    TextBox_������.Value = DateSerial(Year(Now), Month(Now), Day(Now))
End Sub

Private Sub CommandButton_startday_today_Click()
    TextBox_������.Value = DateSerial(Year(Now), Month(Now), Day(Now))
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm5
    If Parent = "ȸ�����" Then
        Worksheets("ȸ�����").Activate
    ElseIf Parent = "ǰ�Ǽ�����" Then
        Worksheets("ǰ�Ǽ�����").Activate
    Else
        Ȩ
    End If
End Sub

Private Sub CommandButton3_Click()
    Worksheets("�׸�а���").Activate
    Worksheets("��꼭").Activate
    
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Worksheets("��꼭").Activate
    Unload Me
End Sub

Private Sub CommandButton7_Click()
    Worksheets("�׸�а���").Activate
    Unload Me
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "������Ʈ�� ��꼭�� �����մϴ�(2014�� 10�� ���� ������)"
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "'��'�� ���� ��¥�� ���� ���� ù������ �������� ���õ˴ϴ�"
End Sub

Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_�޽���.caption = "�Ⱓ����� ������ Ȥ�� �ϰ�ǥ�� ��/��ݿ����� ���س�¥�� �˴ϴ�."
End Sub

Private Sub OptionButton_thismonth_Click()
    Dim dtLastDayofMonth As Date
    dtLastDayofMonth = DateAdd("d", -1, DateSerial(Year(Now), Month(Now) + 1, 1))
    TextBox_������.Value = DateSerial(Year(Now), Month(Now), 1)
    TextBox_������.Value = dtLastDayofMonth
End Sub

Private Sub OptionButton_thisquarter_Click()
    Dim dtLastDayofMonth As Date
    Dim dThisMonth As Integer
    Dim quarter As Integer
    dThisMonth = Month(Now)
    Dim b As Integer
    Dim dStartMonth As Integer
    
    quarter = dThisMonth / 3
    b = dThisMonth Mod 3
    If b > 0 Then
        quarter = quarter + 1
    End If
        
    dStartMonth = 3 * (quarter - 1) + 1
    
    dtFirstDayofQuarter = DateSerial(Year(Now), dStartMonth, 1)
    dtLastDayofQuarter = DateAdd("d", -1, DateSerial(Year(Now), dStartMonth + 3, 1))
    TextBox_������.Value = dtFirstDayofQuarter
    TextBox_������.Value = dtLastDayofQuarter
End Sub

Private Sub OptionButton_thisyear_Click()
    TextBox_������.Value = DateSerial(Year(Now), 1, 1)
    TextBox_������.Value = DateAdd("d", -1, DateSerial(Year(Now) + 1, 1, 1))
End Sub

Private Sub SpinButton_������_SpinDown()
    TextBox_������.Value = DateAdd("d", -1, TextBox_������.Value)
End Sub

Private Sub SpinButton_������_SpinUp()
    TextBox_������.Value = DateAdd("d", 1, TextBox_������.Value)
End Sub

Private Sub SpinButton_������_SpinDown()
    TextBox_������.Value = DateAdd("d", -1, TextBox_������.Value)
End Sub

Private Sub SpinButton_������_SpinUp()
    TextBox_������.Value = DateAdd("d", 1, TextBox_������.Value)
End Sub

Private Sub UserForm_Initialize()

    With Worksheets("����")
        .Activate
        ������ = .Range("�۾������ϼ���").Offset(0, 1).Value
        If ������ = "" Then
            ������ = .Range("ȸ������ϼ���").Offset(0, 1).Value
        End If
        TextBox_������.Value = ������
        ������ = .Range("�۾������ϼ���").Offset(0, 1).Value
        If ������ = "" Then
            ������ = Date
        End If
        TextBox_������.Value = ������
    End With
    CheckBox_default.Value = True
    
    Call ������Ʈ_�ʱ�ȭ
    
End Sub

Sub ������Ʈ_�ʱ�ȭ()
    Dim ������ As Range
    Dim ������ As Range
    Dim ������Ʈ As Range
    Dim ws As Worksheet
    Set ws = Worksheets("����")
    Dim ������Ʈ�� As Integer
    
    With ws.Range("������Ʈ�������̺�")
        If .Offset(1, 0).Value <> "" Then
            ������Ʈ�� = .CurrentRegion.Rows.Count - 1
        Else
            ������Ʈ�� = 0
        End If
        
        If ������Ʈ�� > 0 Then
            ComboBox_������Ʈ.Enabled = True
            Set ������ = .Offset(1)
            If ������.Offset(1, 0).Value <> "" Then
                Set ������ = ������.End(xlDown)
            Else
                Set ������ = ������
            End If
            
            For Each ������Ʈ In ws.Range(������, ������)
                If ������Ʈ.Value <> "" And ������Ʈ.Value <> "������Ʈ��" Then
                    ComboBox_������Ʈ.AddItem ������Ʈ.Value
                End If
            Next ������Ʈ
        Else
            ComboBox_������Ʈ.Enabled = False
        End If
    End With
End Sub

