VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "결산서선택"
   ClientHeight    =   8380.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895.001
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_기간결산_Click()
    Dim startDate As Date, endDate As Date
    Dim project As String
    
    If TextBox_시작일.Value = "" Then
        MsgBox "시작일을 입력해주세요"
        Exit Sub
    End If
    
    If TextBox_종료일.Value = "" Then
        MsgBox "종료일을 입력해주세요"
        Exit Sub
    End If
    
    If Not IsDate(TextBox_시작일.Value) Then
        MsgBox "시작일이 잘못 입력되었습니다. (예: 28일만 있는 달에 29일을 입력)"
        Exit Sub
    End If
    If Not IsDate(TextBox_종료일.Value) Then
        MsgBox "종료일이 잘못 입력되었습니다. (예: 30일만 있는 달에 31일을 입력)"
        Exit Sub
    End If
    
    startDate = IIf(TextBox_시작일.Value <> "", TextBox_시작일.Value, get_config("회계시작일"))
    endDate = IIf(TextBox_종료일.Value <> "", TextBox_종료일.Value, Date)
    
    If date_compare(startDate, endDate) < 0 Then
        MsgBox "종료일이 시작일보다 앞섭니다. 종료일을 다시 설정해주세요"
        Exit Sub
    End If
    
    With Worksheets("설정")
        .Activate
        .Range("작업시작일설정").Offset(0, 1).Value = startDate
        .Range("작업종료일설정").Offset(0, 1).Value = endDate
    End With

    'project 변수는 module12에 정의된 전역변수
    If ComboBox_프로젝트.Value <> "" Then
        project = ComboBox_프로젝트.Value
    Else
        project = ""
    End If
    
    Application.DisplayStatusBar = True
    
    rebuild_report = CheckBox_default.Value
    If rebuild_report Then
        Application.StatusBar = "결산서를 초기화하고 있습니다."
        Call 결산서초기화2
        Application.StatusBar = "결산서가 초기화되었습니다."
    End If
    
    Call 항목결산작성(CheckBox_항목분계장.Value, project) ' module 12

    If CheckBox_1page.Value = True Then
        report_1p = True
        
        Call 결산1p
    End If
    
    Application.StatusBar = "결산서 생성이 완료되었습니다"
    Unload Me
End Sub

Private Sub btn_기간결산_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "오른쪽에 설정한 기간에 해당하는 결산서를 생성합니다"
End Sub

Private Sub btn_일계표_Click()
    Call 날짜선택("일계표")
    Call 일계표작성
    Unload Me
End Sub

Sub 날짜선택(paper As String)
    Dim ws As Worksheet
    Set ws = Worksheets("회계원장")
    ws.Activate
    Dim curRange As Range
    Dim i As Integer
    Dim 끝행 As Integer
    끝행 = ws.Range("A6").End(xlDown).Row
    Dim 날짜존재 As Integer
    날짜존재 = 0
    Dim startDate As Date, endDate As Date
    startDate = IIf(TextBox_시작일.Value <> "", TextBox_시작일.Value, get_config("회계시작일"))
    endDate = IIf(TextBox_종료일.Value <> "", TextBox_종료일.Value, Date)
    
    Dim r일자 As Range
    '날짜가 시트에 어떻게 표시되던, find는 "m/d/yyyy" 형태로 검색해야 찾아진다.
    'find 인수로 format( ,"m/d/yyyy") 안 먹힘
    종료일_연 = format(endDate, "yyyy")
    종료일_월 = format(endDate, "m")
    종료일_일 = format(endDate, "d")
    Set r일자 = ws.Range("일자필드레이블").CurrentRegion.columns(1).Find(종료일_월 & "/" & 종료일_일 & "/" & 종료일_연)
    
    If r일자 Is Nothing Then
        MsgBox "지정한 날짜(종료일) 자료를 찾지 못했습니다. 최근에 입력한 내용을 엽니다"
        Set curRange = ws.Range("일자필드레이블").End(xlDown)
        i = 끝행
        
        If paper = "입금원장" Then

            Do While i > 6
                If ws.Range("A" & i).Offset(, 3).Value = "수입" Then
                    ws.Range("A" & i).Select
                    Exit Do
                End If
                i = i - 1

            Loop
            
        ElseIf paper = "출금원장" Then

            Do While i > 6
                If ws.Range("A" & i).Offset(, 3).Value = "지출" Then
                    ws.Range("A" & i).Select
                    Exit Do
                End If
                i = i - 1
            Loop

        Else '일계표
            curRange.Select
        End If
        
    Else
        With Worksheets("설정")
            .Activate
            .Range("작업시작일설정").Offset(0, 1).Value = startDate
            .Range("작업종료일설정").Offset(0, 1).Value = endDate
        End With
        ws.Activate
        
        If paper = "입금원장" Then
            If r일자.Offset(, 3).Value <> "수입" Then
                For i = r일자.Row To 끝행
                    If ws.Range("A" & i).Value <> CDate(종료일) Then
                        Exit For
                    End If
                    If ws.Range("A" & i).Offset(, 3).Value = "수입" Then
                        ws.Range("A" & i).Select
                        날짜존재 = 1
                        Exit For
                    End If
                Next i
                
                If Not 날짜존재 > 0 Then
                    MsgBox "해당 날짜의 수입 기록이 없습니다. 다른 날짜를 지정해주세요"
                    ws.Range("일자필드레이블").Select
                End If
            Else
                r일자.Select
            End If
            
        ElseIf paper = "출금원장" Then
            If r일자.Offset(, 3).Value <> "지출" Then
                
                For i = r일자.Row To 끝행
                    If ws.Range("A" & i).Value <> CDate(종료일) Then
                        Exit For
                    End If
                    If ws.Range("A" & i).Offset(, 3).Value = "지출" Then
                        ws.Range("A" & i).Select
                        날짜존재 = 1
                        Exit For
                    End If
                Next i
                
                If Not 날짜존재 > 0 Then
                    MsgBox "해당 날짜의 지출 기록이 없습니다. 다른 날짜를 지정해주세요"
                    ws.Range("일자필드레이블").Select
                End If
            Else
                r일자.Select
            End If
        Else '일계표
            r일자.Select 'curRange.Select
        End If
        
    End If

End Sub

Private Sub btn_일계표_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "종료일로 설정한 날짜의 일계표를 생성합니다"
End Sub

Private Sub CommandButton_endday_today_Click()
    TextBox_종료일.Value = DateSerial(Year(Now), Month(Now), Day(Now))
End Sub

Private Sub CommandButton_startday_today_Click()
    TextBox_시작일.Value = DateSerial(Year(Now), Month(Now), Day(Now))
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm5
    If Parent = "회계원장" Then
        Worksheets("회계원장").Activate
    ElseIf Parent = "품의서대장" Then
        Worksheets("품의서대장").Activate
    Else
        홈
    End If
End Sub

Private Sub CommandButton3_Click()
    Worksheets("항목분계장").Activate
    Worksheets("결산서").Activate
    
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Worksheets("결산서").Activate
    Unload Me
End Sub

Private Sub CommandButton7_Click()
    Worksheets("항목분계장").Activate
    Unload Me
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "프로젝트별 결산서를 생성합니다(2014년 10월 현재 제작중)"
End Sub

Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "'월'은 오늘 날짜가 속한 월의 첫날부터 끝날까지 선택됩니다"
End Sub

Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "기간결산의 종료일 혹은 일계표와 입/출금원장의 기준날짜가 됩니다."
End Sub

Private Sub OptionButton_thismonth_Click()
    Dim dtLastDayofMonth As Date
    dtLastDayofMonth = DateAdd("d", -1, DateSerial(Year(Now), Month(Now) + 1, 1))
    TextBox_시작일.Value = DateSerial(Year(Now), Month(Now), 1)
    TextBox_종료일.Value = dtLastDayofMonth
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
    TextBox_시작일.Value = dtFirstDayofQuarter
    TextBox_종료일.Value = dtLastDayofQuarter
End Sub

Private Sub OptionButton_thisyear_Click()
    TextBox_시작일.Value = DateSerial(Year(Now), 1, 1)
    TextBox_종료일.Value = DateAdd("d", -1, DateSerial(Year(Now) + 1, 1, 1))
End Sub

Private Sub SpinButton_시작일_SpinDown()
    TextBox_시작일.Value = DateAdd("d", -1, TextBox_시작일.Value)
End Sub

Private Sub SpinButton_시작일_SpinUp()
    TextBox_시작일.Value = DateAdd("d", 1, TextBox_시작일.Value)
End Sub

Private Sub SpinButton_종료일_SpinDown()
    TextBox_종료일.Value = DateAdd("d", -1, TextBox_종료일.Value)
End Sub

Private Sub SpinButton_종료일_SpinUp()
    TextBox_종료일.Value = DateAdd("d", 1, TextBox_종료일.Value)
End Sub

Private Sub UserForm_Initialize()

    With Worksheets("설정")
        .Activate
        시작일 = .Range("작업시작일설정").Offset(0, 1).Value
        If 시작일 = "" Then
            시작일 = .Range("회계시작일설정").Offset(0, 1).Value
        End If
        TextBox_시작일.Value = 시작일
        종료일 = .Range("작업종료일설정").Offset(0, 1).Value
        If 종료일 = "" Then
            종료일 = Date
        End If
        TextBox_종료일.Value = 종료일
    End With
    CheckBox_default.Value = True
    
    Call 프로젝트_초기화
    
End Sub

Sub 프로젝트_초기화()
    Dim 시작행 As Range
    Dim 종료행 As Range
    Dim 프로젝트 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("설정")
    Dim 프로젝트수 As Integer
    
    With ws.Range("프로젝트설정레이블")
        If .Offset(1, 0).Value <> "" Then
            프로젝트수 = .CurrentRegion.Rows.Count - 1
        Else
            프로젝트수 = 0
        End If
        
        If 프로젝트수 > 0 Then
            ComboBox_프로젝트.Enabled = True
            Set 시작행 = .Offset(1)
            If 시작행.Offset(1, 0).Value <> "" Then
                Set 종료행 = 시작행.End(xlDown)
            Else
                Set 종료행 = 시작행
            End If
            
            For Each 프로젝트 In ws.Range(시작행, 종료행)
                If 프로젝트.Value <> "" And 프로젝트.Value <> "프로젝트명" Then
                    ComboBox_프로젝트.AddItem 프로젝트.Value
                End If
            Next 프로젝트
        Else
            ComboBox_프로젝트.Enabled = False
        End If
    End With
End Sub

