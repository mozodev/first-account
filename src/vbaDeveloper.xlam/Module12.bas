Attribute VB_Name = "Module12"
'module12 : 결산서
Option Explicit

Const 결산서_열_누적수입 As String = "i"
Const 결산서_열_누적지출 As String = "j"
Const 결산서_열_기간수입 As String = "g"
Const 결산서_열_기간지출 As String = "h"
Const 항목분계장_열_날짜 As String = "A"
Const 항목분계장_열_관 As String = "C"
Const 항목분계장_열_항 As String = "D"
Const 항목분계장_열_목 As String = "E"
Const 항목분계장_열_세목 As String = "f"
Const 항목분계장_열_적요 As String = "g"
Const 항목분계장_열_수입 As String = "h"
Const 항목분계장_열_지출 As String = "i"
Const 최대행수 As Integer = 20000
Const 회계원장_헤더행수 As Integer = 5

Public project As String
Public rebuild_report As Boolean
Public report_1p As Boolean

Sub 항목결산작성(ByVal 항목분계장생성 As Boolean, Optional ByVal project As String)

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "결산서를 작성하고 있습니다. 잠시 기다려주세요.."

    Dim ws_ledger As Worksheet
    Dim 항목분계장 As Worksheet
    Dim ws_help As Worksheet, ws_budget As Worksheet, ws_settle As Worksheet
    
    Set ws_ledger = Worksheets("회계원장")
    Set 항목분계장 = Worksheets("항목분계장")
    Set ws_help = Worksheets("도움말")
    Set ws_budget = Worksheets("예산서")
    Set ws_settle = Worksheets("결산서")
    
    ws_ledger.Activate

    Dim wkSht As Worksheet
 
    'On Error Resume Next                                        '에러 발생해도 계속 코드 진행
    Set wkSht = ThisWorkbook.Worksheets("도움말")  '시트를 개체변수에 넣음
    
    If Err.Number = 0 Then                                        '에러가 없다면
                                                   '그 시트는 존재함
    Else                               '(없는 시트를 개체변수에 넣으므로)그와 달리 에러가 발생하면
        MsgBox "파일이 손상되었습니다. http://firstaccounting.org 에서 다시 다운 받아서 사용하세요.": Exit Sub                                            '그 시트는 존재하지 않음
    End If

    If ws_help.Range("aa1") <> 300 Then MsgBox "파일이 손상되었습니다. http://kimjy.net 에서 다시 다운 받아서 사용하세요.": Exit Sub
    
    Set ws_help = Nothing
    Set wkSht = Nothing
    
    Dim 시작줄 As Integer, 끝줄 As Integer, 입력줄 As Integer
    Dim startDate As Date, endDate As Date, 회계시작일 As Date, 오늘 As Date
    Dim i As Integer, i2 As Integer, i3 As Integer
    
    startDate = format(get_config("시작일"), "Short Date")
    endDate = format(get_config("종료일"), "Short Date")
    회계시작일 = format(get_config("회계시작일"), "Short Date")
    오늘 = format(Date, "Short Date")

    입력줄 = 6
    시작줄 = 6
    끝줄 = ws_ledger.Range("일자필드레이블").End(xlDown).Row

' 프로세스
' #1. 전체 기간 결산 피벗테이블 생성 (아래에서 결산서에 반영)
' #2. 설정된 기간 항목분계장 작성 시작
' #3. 기존 데이터 지우고 새 데이터 카피 후, 설정 기간 결산 피벗테이블 생성 (아래에서 결산서에 반영)
' #4. 항목분계장 작성 완료
' #5. 결산서 내용 업데이트 : 전체 기간, 설정 기간

' #0. 시작일, 종료일에 따라 데이터 범위 파악
'시작줄과 끝줄 알아내는 부분
    Dim dataCount As Integer
    dataCount = 끝줄 - 5
    Dim startRow As Integer, endRow As Integer
    Dim c As Range
    Set c = ws_ledger.Range("A5")
    Dim dateArray() As Variant
    ReDim dateArray(1 To dataCount)
    
    For i = 1 To dataCount
        dateArray(i) = c.Offset(i, 0).Value
    Next i
    
    If date_compare(startDate, endDate) < 0 Then
        MsgBox "종료일이 시작일보다 빠릅니다. 1년 전체 결산만 표시됩니다"
        Exit Sub
    End If
    If date_compare(dateArray(dataCount), startDate) > 0 Then
        MsgBox "시작일이 잘못 지정되었습니다. 1년 전체 결산만 표시됩니다."
        Exit Sub
    End If
    If date_compare(dateArray(1), endDate) < 0 Then
        MsgBox "종료일이 잘못 지정되었습니다. 1년 전체 결산만 표시됩니다."
        Exit Sub
    End If
    
    For i = 1 To dataCount
        If dateArray(i) = startDate Then '이 부분은 정상 작동
            startRow = i + 5
            Exit For
        Else
            If date_compare(dateArray(i), startDate) < 0 Then '아직 시작일보다 빠른 날짜이면 더 순회
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
            If date_compare(dateArray(i), endDate) > 0 Then '같진 않지만 종료일보다 빠른 날짜가 나타났다면 이걸 종료행으로
                endRow = i + 5
                Exit For
            End If
        End If
    Next i

    Erase dateArray
    '데이터 값이 없으면 결산서 생성 안하게
    If Not startRow < endRow Then
        MsgBox "해당기간에 입력된 자료가 없거나 시작일과 종료일이 잘못 지정되었습니다. 기간을 다시 설정해 시도해보세요"
        Exit Sub
    End If
    
    항목분계장.Activate
    'ws_help.Activate
    ws_budget.Activate
    ws_settle.Activate
        
' #1. 전체 기간 결산 피벗테이블 생성 (아래에서 결산서에 반영)

    If rebuild_report Then
        'Call 결산서초기화
        Call 결산피벗테이블생성(회계시작일, 오늘)
    End If

' #2. 설정된 기간 항목분계장 작성 시작
    If 항목분계장생성 Then

        With 항목분계장.Range("A6:" & 항목분계장_열_지출 & 최대행수) ' 이 최대행수 값이 아직은 자동 계산되지 않음
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
   
    
    '항목분계장으로 회계원장의 해당날짜 데이터 복사
        Call 원장데이터copy("항목분계장", startRow, endRow)
    
    '메모리 지우기
        Application.CutCopyMode = False
 End If
 
' #3. 카피된 데이터를 바탕으로, 설정 기간의 결산 피벗테이블 생성 (아래에서 결산서에 반영)
    If rebuild_report Then
        Call 결산피벗테이블생성(startDate, endDate)
    End If

' #4. 결산서 내용 업데이트 : 전체 기간, 설정 기간 (수정 2차 완료: 2016.9.20)


' #4-1. 결산서 작성
    If rebuild_report Then
        Dim codeCount As Integer
        codeCount = ws_settle.Range("A5").End(xlDown).Row - 5
        
        '계정코드를 배열로 하게 변환 2016. 9.20
        Dim sCode() As Variant
        ReDim sCode(1 To codeCount)
        Dim settleResult() As Variant
        ReDim settleResult(1 To codeCount, 1 To 4)
        
        Dim j As Integer
        For j = 1 To codeCount
            sCode(j) = ws_settle.Range("A" & j + 5).Value
        Next j
        
        For j = 1 To codeCount
            settleResult(j, 1) = 코드결산(sCode(j), "수입", "부분")
            settleResult(j, 2) = 코드결산(sCode(j), "지출", "부분")
            settleResult(j, 3) = 코드결산(sCode(j), "수입", "전체")
            settleResult(j, 4) = 코드결산(sCode(j), "지출", "전체")
        Next j
    
        With ws_settle
            .Range(결산서_열_기간수입 & 6, 결산서_열_누적지출 & codeCount + 5).Value = settleResult
            .Range("a3").Value = "(" & startDate & " ~ " & endDate & ")"
        End With
    End If
    ws_settle.Visible = xlSheetVisible
    
    If 항목분계장생성 Then
    ' #4-2. 항목분계장 작성
        항목분계장.Activate
        dataCount = 0
        Application.StatusBar = "계정별원장을 작성중입니다. 잠시만 기다려주세요."
        
        With 항목분계장.Range("a6:" & 항목분계장_열_지출 & (끝줄 - 시작줄 + 6))
            .Sort Key1:=Range("B6"), Order1:=xlAscending, Key2:=Range("A6") _
            , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, _
            Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers, DataOption2 _
            :=xlSortNormal
        End With
        
        With 항목분계장
            '.Range("b6").Select  '항목분계장 첫번째 데이터
    
            Dim 삭제줄수 As Integer
            Dim 현재줄 As Integer
            dataCount = .Range("A5").End(xlDown).Row - 5
            삭제줄수 = 0
            '예산외수입, 예산외지출 제거 시작
            For i3 = 6 To (dataCount - 1 + 6)
                If Left(.Range("b" & i3).Value, 2) = "00" Then
                
                    현재줄 = i3 '.Row
                    .Range("A" & 현재줄 & ":" & 항목분계장_열_지출 & 현재줄).Delete Shift:=xlUp
                    '.Range("b" & 현재줄).Select
                    삭제줄수 = 삭제줄수 + 1
                End If
            Next i3
            dataCount = dataCount - 삭제줄수
            '예산외수입, 예산외지출 제거 완료
            
            Dim 코드 As String, 이전코드 As String
            '코드 = .Range("b6").Value
            'Dim 입력줄수 As Integer
           ' 입력줄수 = 1
            
            '이 부분을 성능 개선해야 함
            
            Dim tmpDataArray() As Variant
            ReDim tmpDataArray(1 To dataCount, 1 To 9)
            Set c = .Range("A5")
            tmpDataArray = .Range(c.Offset(1, 0), c.Offset(dataCount, 8)).Value
            
            Dim 입력줄수 As Integer, 입력행 As Integer, 삽입줄수 As Integer
            입력줄수 = 1
            입력행 = 0
            삽입줄수 = 0
            이전코드 = ""
            Set c = .Range("b5")
            
            For i3 = 1 To dataCount
                코드 = tmpDataArray(i3, 2)
                
                If 이전코드 <> "" Then '첫 줄은 비교하지 않는다
                    If 코드 <> 이전코드 Then
                        입력행 = i3 + 5 + 삽입줄수
                        
                        .Rows(입력행).Insert
                        삽입줄수 = 삽입줄수 + 1
                        
                        .Range("a" & 입력행).Value = "합 계"
                        .Range("b" & 입력행).Value = Range("b" & 입력행 - 1).Value
                
                        With .Range(항목분계장_열_수입 & 입력행)
                            .FormulaR1C1 = "=SUbtotal(9,R[-" & 입력줄수 & "]C:R[-1]C)"
                            '.Font.Bold = True
                        End With
                
                        With .Range(항목분계장_열_지출 & 입력행)
                            .FormulaR1C1 = "=SUbtotal(9,R[-" & 입력줄수 & "]C:R[-1]C)"
                            '.Font.Bold = True
                        End With
                
                        With .Range("a" & 입력행 & ":" & 항목분계장_열_지출 & 입력행)
                            .Borders(xlEdgeBottom).Weight = xlMedium
                            .Font.Bold = True
                        End With
                        
                        입력줄수 = 1
                    Else
                        입력줄수 = 입력줄수 + 1
                    End If
                    
                End If
                
                이전코드 = 코드
            Next i3
            
            Erase tmpDataArray
            
            '마지막 코드 합계
            입력행 = dataCount + 5 + 삽입줄수 + 1
            .Rows(입력행).Insert
            
            .Range("a" & 입력행).Value = "합 계"
            .Range("b" & 입력행).Value = Range("b" & 입력행 - 1).Value
    
            With .Range(항목분계장_열_수입 & 입력행)
                .FormulaR1C1 = "=SUbtotal(9,R[-" & 입력줄수 & "]C:R[-1]C)"
            End With
    
            With .Range(항목분계장_열_지출 & 입력행)
                .FormulaR1C1 = "=SUbtotal(9,R[-" & 입력줄수 & "]C:R[-1]C)"
            End With
    
            With .Range("a" & 입력행 & ":" & 항목분계장_열_지출 & 입력행)
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Font.Bold = True
            End With
                        
            '총액 계산하고 마무리
            Dim finalRow As Integer
            finalRow = 입력행 + 1
            
            .Range("a" & finalRow).Value = "총 계"
        
            With .Range(항목분계장_열_수입 & finalRow)
                .FormulaR1C1 = "=subtotal(9,R[-" & finalRow - 5 & "]C:R[-1]C)"
            End With
        
            With .Range(항목분계장_열_지출 & finalRow)
                .FormulaR1C1 = "=subtotal(9,R[-" & finalRow - 5 & "]C:R[-1]C)"
            End With
        
            With Range("a" & finalRow & ":" & 항목분계장_열_지출 & finalRow)
                .Font.Bold = True
                With .Borders(xlEdgeBottom)
                    .Weight = xlMedium
                End With
            End With
        
            .Range("a" & finalRow + 1).Value = "잔 액"
        
            With .Range(항목분계장_열_지출 & finalRow + 1)
                .FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
                '.Font.Bold = True
            End With
        
            With .Range("a" & finalRow + 1 & ":" & 항목분계장_열_지출 & finalRow + 1)
                .Font.Bold = True
                With .Borders(xlEdgeBottom)
                    .Weight = xlMedium
                End With
            End With
        
            .Range("a2").Value = "(" & startDate & " ~ " & endDate & ")"
            .PageSetup.PrintArea = "$a$1:$" & 항목분계장_열_지출 & "$" & finalRow + 1
    
        End With
    End If

    항목분계장.Visible = xlSheetVisible
End Sub

Sub 결산1p()
'항 단위로 합산하여 간단한 결산서 생성
    Const 결산서헤더행수 = 5
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    Dim ws_source, ws_target As Worksheet
    Set ws_source = Worksheets("결산서")
    Set ws_target = Worksheets("결산서1p")
    
    ws_target.Range("A7", "A1000").EntireRow.Delete
    
    Dim 관라벨, 항라벨, 기간수입라벨, 기간지출라벨 As Range
    With ws_source
        Set 관라벨 = .Range("B5")
        Set 항라벨 = .Range("C5")
        Set 기간수입라벨 = .Range("G5")
        Set 기간지출라벨 = .Range("H5")
    End With
    
    Dim 수입list() As Variant, 지출list() As Variant
    Dim 기준점 As Range
    Dim 행수 As Integer, 끝행 As Integer, i As Integer, x As Integer, y As Integer
    x = 0
    y = 0
    Dim 새항 As Boolean
    Dim 수입금액 As Long, 지출금액 As Long, 예산액 As Long
    Dim 이전항 As String
    
    'Set 기준점 = 항라벨
    행수 = 항라벨.End(xlDown).Row - 결산서헤더행수
    
    For i = 1 To 행수
        Set 기준점 = 항라벨.Offset(i, 0)
        예산액 = 기준점.Offset(, 3).Value
        수입금액 = 기준점.Offset(, 4).Value
        지출금액 = 기준점.Offset(, 5).Value
        새항 = True
        
        If 기준점.Offset(, -1).Value = "수입" Then
            If x > 0 Then
                If 기준점.Value = 이전항 Then
                    새항 = False
                End If
            End If
            
            If 새항 Then
                ReDim Preserve 수입list(3, x)
                수입list(0, x) = 기준점.Value
                수입list(1, x) = 예산액
                수입list(2, x) = 수입금액
                x = x + 1
                이전항 = 기준점.Value
            Else
                수입list(1, x - 1) = 수입list(1, x - 1) + 예산액
                수입list(2, x - 1) = 수입list(2, x - 1) + 수입금액
            End If
        Else
            If y > 0 Then
                If 기준점.Value = 이전항 Then
                    새항 = False
                End If
            End If
            
            If 새항 Then
                ReDim Preserve 지출list(3, y)
                지출list(0, y) = 기준점.Value
                지출list(1, y) = 예산액
                지출list(2, y) = 지출금액
                y = y + 1
                이전항 = 기준점.Value
            Else
                지출list(2, y - 1) = 지출list(2, y - 1) + 지출금액
                지출list(1, y - 1) = 지출list(1, y - 1) + 예산액
            End If
        End If
        
    Next i
    
    Set 기준점 = Nothing
    
    '계산한 값으로 결산서1p 채우기
    With ws_target
        '수입
        Dim 수입항수, 지출항수, 항수 As Integer
        수입항수 = UBound(수입list, 2) + 1
        For i = 0 To 수입항수 - 1
            With .Range("결산서1p수입항").Offset(i + 1, 0)
                .Value = 수입list(0, i)

                .Offset(0, 1).Value = 수입list(1, i)
                .Offset(0, 2).Value = 수입list(2, i)
            End With

        Next i
        Erase 수입list
        
        지출항수 = UBound(지출list, 2) + 1
        For i = 0 To 지출항수 - 1
            With .Range("결산서1p지출항").Offset(i + 1, 0)
                .Value = 지출list(0, i)

                .Offset(0, 1).Value = 지출list(1, i)
                .Offset(0, 2).Value = 지출list(2, i)
            End With

        Next i
        Erase 지출list
        
        If 수입항수 > 지출항수 Then
            항수 = 수입항수
        Else
            항수 = 지출항수
        End If
        
        끝행 = 항수 + 결산서헤더행수 + 1
        
        With .Range("결산서1p수입항").Offset(항수 + 1, 0)
            .Value = "합계"
            .Font.Bold = True
            .Offset(0, 1).Formula = "=sum(B7:B" & 끝행 & ")"
            .Offset(0, 2).Formula = "=sum(C7:C" & 끝행 & ")"
        End With
        
        With .Range("결산서1p지출항").Offset(항수 + 1, 0)
            .Value = "합계"
            .Font.Bold = True
            .Offset(0, 1).Formula = "=sum(E7:E" & 끝행 & ")"
            .Offset(0, 2).Formula = "=sum(F7:F" & 끝행 & ")"
        End With
        
    End With
    
    With ws_target.Range("결산서1p수입").CurrentRegion
        .Borders.LineStyle = 1
    End With
    
    끝행 = 끝행 + 1
    With ws_target.Rows("7:" & 끝행)
        .RowHeight = 30
        .Font.size = 14
    End With
    
    With ws_target.Range("A" & 끝행 & ":F" & 끝행)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    With ws_target.Range("D5:D" & 끝행)
        .Borders(xlEdgeLeft).LineStyle = xlDouble
    End With
    
    Dim startDate, endDate As Date
    startDate = format(get_config("시작일"), "Short Date")
    endDate = format(get_config("종료일"), "Short Date")
    ws_target.Range("a3").Value = get_config("기관명") & " (" & startDate & " ~ " & endDate & ")"
    
    ws_target.Visible = xlSheetVisible

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub
