VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_품의서결의서생성 
   Caption         =   "품의서/결의서 생성"
   ClientHeight    =   4980
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5280
   OleObjectBlob   =   "UserForm_품의서결의서생성.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_품의서결의서생성"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const 열offset_날짜 As Integer = 0
Const 열offset_코드 As Integer = 1
Const 열offset_지출명 As Integer = 2
Const 열offset_규격 As Integer = 3
Const 열offset_수량 As Integer = 4
Const 열offset_단가 As Integer = 5
Const 열offset_금액 As Integer = 6
Const 열offset_비고 As Integer = 7
Const 열offset_하단비고 As Integer = 8
    
Sub 날짜선택()
    
    Dim ws As Worksheet
    Set ws = Worksheets("회계원장")
    ws.Activate
    
    시작일 = IIf(TextBox_시작일.Value <> "", TextBox_시작일.Value, "2014-01-01")
    종료일 = IIf(TextBox_종료일.Value <> "", TextBox_종료일.Value, Date)
    
    Dim r일자 As Range

    종료일_연 = format(종료일, "yyyy")
    종료일_월 = format(종료일, "m")
    종료일_일 = format(종료일, "d")
    Set r일자 = ws.Range("일자필드레이블").CurrentRegion.columns(1).Find(종료일_월 & "/" & 종료일_일 & "/" & 종료일_연)
    
    If r일자 Is Nothing Then
        MsgBox "지정한 날짜(종료일) 자료를 찾지 못했습니다. 최근에 입력한 내용을 엽니다"
        ws.Range("일자필드레이블").End(xlDown).Select
    Else

        With Worksheets("설정")
            .Activate
            .Range("작업시작일설정").Offset(0, 1).Value = 시작일
            .Range("작업종료일설정").Offset(0, 1).Value = 종료일
        End With
        ws.Activate
        r일자.Select
    End If

End Sub

Sub 날짜선택2()
    Worksheets("품의서대장").Activate
    
    Dim ws As Worksheet
    Set ws = Worksheets("품의서대장")

    종료일 = IIf(TextBox_종료일.Value <> "", TextBox_종료일.Value, Date)
    
    Dim r일자 As Range

    종료일_연 = format(종료일, "yyyy")
    종료일_월 = format(종료일, "m")
    종료일_일 = format(종료일, "d")
    Set r일자 = ws.Range("품의날짜레이블").CurrentRegion.columns(1).Find(종료일_월 & "/" & 종료일_일 & "/" & 종료일_연)
    
    If r일자 Is Nothing Then
        MsgBox "지정한 날짜(종료일) 자료를 찾지 못했습니다. 최근 날짜의 품의서를 생성합니다"
        ws.Range("품의날짜레이블").End(xlDown).Select
    Else
        r일자.Select
    End If

End Sub

Sub 날짜선택3()
    Dim ws As Worksheet
    Set ws = Worksheets("지출결의대장")
    ws.Activate
    Dim 종료일 As String
    Dim 종료일_연 As String
    Dim 종료일_월 As String
    Dim 종료일_일 As String
    
    종료일 = IIf(TextBox_종료일.Value <> "", TextBox_종료일.Value, Date)
    
    Dim r일자 As Range

    종료일_연 = format(종료일, "yyyy")
    종료일_월 = format(종료일, "m")
    종료일_일 = format(종료일, "d")
    Set r일자 = ws.Range("결의날짜레이블").CurrentRegion.columns(1).Find(종료일_월 & "/" & 종료일_일 & "/" & 종료일_연)
    
    If r일자 Is Nothing Then
        MsgBox "지정한 날짜(종료일) 자료를 찾지 못했습니다. 최근 날짜의 품의서를 생성합니다"
        ws.Range("결의날짜레이블").End(xlDown).Select
    Else
        r일자.Select
    End If

End Sub

'Private Sub btn_입금원장_Click()
'    Call 날짜선택
'    Call 입금원장작성
'    Unload Me
'End Sub

'Private Sub btn_출금원장_Click()
'    Call 날짜선택
'    Call 출금원장작성
'    Unload Me
'End Sub

'Private Sub btn_품의서_Click()
'
'    Dim i선택 As Integer
'    Dim 행번호 As Integer
'
'    With ListBox2
'        i선택 = .ListIndex
'        If i선택 > -1 Then
'            행번호 = .List(i선택, 0)
'            Worksheets("품의서대장").Range("A" & 행번호).Select
'            Call 품의서작성(False)
'        Else
'            Call 날짜선택2
'            Call 품의서작성(True)
'        End If
'    End With
'
'    Unload Me
'End Sub

Private Sub btn_품의서_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "선택한 날짜의 품의서를 생성합니다"
End Sub

Private Sub CommandButton_close2_Click()
    Unload Me
    If Parent = "회계원장" Then
        Worksheets("회계원장").Activate
    ElseIf Parent = "품의서대장" Then
        Worksheets("품의서대장").Activate
    ElseIf Parent = "지출결의대장" Then
        Worksheets("지출결의대장").Activate
    Else
        홈
    End If
End Sub

Private Sub CommandButton_검색_Click()
    Dim 전체 As Range
    Dim 찾은날짜 As Range
    Dim 레코드 As Range
    Dim cell As Range
    Dim x As Integer, y As Integer
    Dim 키워드 As String
    Dim 첫위치 As String
    Dim vlist() As Variant

    키워드 = TextBox_종료일.Value
    If Not IsNumeric(키워드) Then '날짜 전제를 입력한 경우
        키워드 = format(키워드, "m") & "/" & format(키워드, "d") & "/" & format(키워드, "yyyy")
    End If
    
    y = 0
    
    If (Len(키워드) > 0) Then
        Set 전체 = Worksheets("지출결의대장").Range("결의날짜레이블").CurrentRegion.columns(1)
        Set 찾은날짜 = 전체.Find(What:=키워드, LookAt:=xlPart)
        
        If Not 찾은날짜 Is Nothing Then
            첫위치 = 찾은날짜.Address
            
            Do
                ReDim Preserve vlist(10, x)
                Set 레코드 = 찾은날짜.Resize(1, 10)
                vlist(0, x) = 레코드.Row
                vlist(1, x) = 레코드.Cells(, 열offset_날짜 + 1)
                vlist(2, x) = 레코드.Cells(, 열offset_코드 + 1)
                vlist(3, x) = 레코드.Cells(, 열offset_지출명 + 1)
                vlist(4, x) = 레코드.Cells(, 열offset_규격 + 1)
                vlist(5, x) = 레코드.Cells(, 열offset_수량 + 1)
                vlist(6, x) = 레코드.Cells(, 열offset_단가 + 1)
                vlist(7, x) = 레코드.Cells(, 열offset_금액 + 1)
                vlist(8, x) = 레코드.Cells(, 열offset_비고 + 1)
                vlist(9, x) = 레코드.Cells(, 열offset_하단비고 + 1)
                
                x = x + 1
                y = 0
                
                Set 찾은날짜 = 전체.FindNext(찾은날짜)
            Loop While Not 찾은날짜 Is Nothing And 찾은날짜.Address <> 첫위치
            
            ListBox1.Column = vlist
        Else
            MsgBox "검색결과가 존재하지 않습니다"
            ListBox1.Clear
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    If Parent = "회계원장" Then
        Worksheets("회계원장").Activate
    ElseIf Parent = "품의서대장" Then
        Worksheets("품의서대장").Activate
    ElseIf Parent = "지출결의대장" Then
        Worksheets("지출결의대장").Activate
    Else
        홈
    End If
End Sub

Private Sub CommandButton4_Click()
    Worksheets("결산서").Activate
    Unload Me
End Sub

Private Sub CommandButton7_Click()
    Worksheets("항목분계장").Activate
    Unload Me
End Sub

Private Sub OptionButton_thismonth_Click()
    Dim dtLastDayofMonth As Date
    dtLastDayofMonth = DateAdd("d", -1, DateSerial(Year(Now), Month(Now) + 1, 1))
    TextBox_시작일.Value = DateSerial(Year(Now), Month(Now), 1)
    TextBox_종료일.Value = dtLastDayofMonth
    
End Sub

Private Sub OptionButton_thisyear_Click()
    TextBox_시작일.Value = DateSerial(Year(Now), 1, 1)
    TextBox_종료일.Value = DateAdd("d", -1, DateSerial(Year(Now) + 1, 1, 1))
End Sub

Private Sub CommandButton3_Click()

    Dim i선택 As Integer
    Dim 행번호 As Integer
    
    With ListBox1
        i선택 = .ListIndex
        If i선택 > -1 Then
            행번호 = .List(i선택, 0)
            Worksheets("지출결의대장").Range("A" & 행번호).Select
            Call 지출결의서작성(False)
        Else
            Call 날짜선택3
            Call 지출결의서작성(True)
        End If
    End With
    
    Unload Me
End Sub

Private Sub CommandButton3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "선택한 날짜의 지출결의서를 출력합니다"
End Sub

Private Sub CommandButton5_Click()
    Dim 전체 As Range
    Dim 찾은날짜 As Range
    Dim 레코드 As Range
    Dim cell As Range
    Dim x As Integer, y As Integer
    Dim 키워드 As String
    Dim 첫위치 As String
    Dim vlist() As Variant

    키워드 = TextBox_품의날짜.Value
    If Not IsNumeric(키워드) Then '날짜 전제를 입력한 경우
        키워드 = format(키워드, "m") & "/" & format(키워드, "d") & "/" & format(키워드, "yyyy")
    End If
    
    y = 0
    
    If (Len(키워드) > 0) Then
        Set 전체 = Worksheets("품의서대장").Range("품의날짜레이블").CurrentRegion.columns(1)
        Set 찾은날짜 = 전체.Find(What:=키워드, LookAt:=xlPart)
        
        If Not 찾은날짜 Is Nothing Then
            첫위치 = 찾은날짜.Address
            
            Do
                ReDim Preserve vlist(10, x)
                Set 레코드 = 찾은날짜.Resize(1, 10)
                vlist(0, x) = 레코드.Row
                vlist(1, x) = 레코드.Cells(, 열offset_날짜 + 1)
                vlist(2, x) = 레코드.Cells(, 열offset_코드 + 1)
                vlist(3, x) = 레코드.Cells(, 열offset_지출명 + 1)
                vlist(4, x) = 레코드.Cells(, 열offset_규격 + 1)
                vlist(5, x) = 레코드.Cells(, 열offset_수량 + 1)
                vlist(6, x) = 레코드.Cells(, 열offset_단가 + 1)
                vlist(7, x) = 레코드.Cells(, 열offset_금액 + 1)
                vlist(8, x) = 레코드.Cells(, 열offset_비고 + 1)
                vlist(9, x) = 레코드.Cells(, 열offset_하단비고 + 1)
                
                x = x + 1
                y = 0
                
                Set 찾은날짜 = 전체.FindNext(찾은날짜)
            Loop While Not 찾은날짜 Is Nothing And 찾은날짜.Address <> 첫위치
            
            ListBox2.Column = vlist
        Else
            MsgBox "검색결과가 존재하지 않습니다"
            ListBox2.Clear
        End If
    End If
End Sub

Private Sub SpinButton_종료일_SpinDown()
    TextBox_종료일.Value = DateAdd("d", -1, TextBox_종료일.Value)
End Sub

Private Sub SpinButton_종료일_SpinUp()
    TextBox_종료일.Value = DateAdd("d", 1, TextBox_종료일.Value)
End Sub

Private Sub UserForm_Initialize()

    If Parent = "품의서대장" Or Parent = "품의서대장_from_홈" Then
        With Worksheets("품의서대장")
            .Activate
            TextBox_시작일 = .Range("품의날짜레이블").Offset(1, 0).Value
            TextBox_품의날짜 = .Range("품의날짜레이블").End(xlDown).Value
        End With
    
    ElseIf Parent = "지출결의대장" Or Parent = "지출결의대장_from_홈" Then
        With Worksheets("지출결의대장")
            .Activate
            TextBox_시작일 = .Range("결의날짜레이블").Offset(1, 0).Value
            TextBox_종료일 = .Range("결의날짜레이블").End(xlDown).Value
        End With
        
    Else

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
            TextBox_품의날짜.Value = 종료일
        End With
    End If
    
    With ListBox1
        .columnCount = 9
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;0cm;0cm;0cm;1cm;3cm"
    End With
    With ListBox2
        .columnCount = 9
        .ColumnWidths = "0cm;2cm;1.5cm;1.5cm;0cm;0cm;0cm;1cm;3cm"
    End With
End Sub

