Attribute VB_Name = "Module11"
'module11 : 추가 기능 관련 함수들 - get_code, 결산데이터갱신(테스트용), 코드결산,
'  get_config, set_config, 결산피벗테이블생성, date_compare, 원장데이터copy, 원장데이터통합, init_firstday
Option Explicit

Const 회계원장_열_일자 As String = "a"
Const 회계원장_열_code As String = "c"
Const 회계원장_열_관 As String = "d"
Const 회계원장_열_항 As String = "e"
Const 회계원장_열_목 As String = "f"
Const 회계원장_열_세목 As String = "g"
Const 회계원장_열_적요 As String = "h"
Const 회계원장_열_수입 As String = "i"
Const 회계원장_열_지출 As String = "j"
Const 회계원장_열_은현 As String = "k"
Const 회계원장_열_프로젝트 As String = "n"
Const 회계원장_열_현금잔액 As String = "P"
Const 회계원장_열_총잔액 As String = "R"
Const 최대행수 As Integer = 30000

'get_code 함수의 성능을 개선하기 위해 이 모듈에서 사용하는 전역변수 배열 사용
Public accountCodes() As Variant
Public codeArrayInitialized As Boolean

Sub home()
' home 매크로
' 바로 가기 키: Ctrl+Shift+H
'
    Sheets("첫페이지").Activate
End Sub

Function get_code(ByVal 관 As String, ByVal 항 As String, ByVal 목 As String, ByVal 세목 As String)
    Dim code As String
    code = ""
    
    '처음 실행하는 거면 전역변수인 accountCodes() 를 초기화한다.
    If Not codeArrayInitialized Then
        Call init_accountCodes
    End If
    
    Dim i As Integer
    For i = 1 To UBound(accountCodes)
        If accountCodes(i, 2) = 관 And accountCodes(i, 3) = 항 And accountCodes(i, 4) = 목 And accountCodes(i, 5) = 세목 Then
            code = accountCodes(i, 1)
        End If
    Next i

    get_code = code
End Function

Sub init_accountCodes()
    Dim ws As Worksheet
    Dim c As Range ', oRng As Range
    Set ws = Worksheets("예산서")
    Set c = Worksheets("예산서").Range("관항목코드레이블")
    
    Dim codeCount As Integer
    codeCount = c.End(xlDown).Row - 1
    
    ReDim accountCodes(1 To codeCount, 1 To 5)
    accountCodes = ws.Range(c.Offset(1, 0), c.Offset(codeCount, 4)).Value
    codeArrayInitialized = True
End Sub

Function new_code(ByVal 관 As String, ByVal 항 As String, ByVal 목 As String, ByVal 세목 As String, Optional ByVal refresh As Boolean)
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")

    Dim accountCode As String
    
    If 관 = "" Or 항 = "" Or 목 = "" Or 세목 = "" Then
        new_code = ""
    Else
        Dim r저장위치 As Range
        Set r저장위치 = ws.Range("관항목코드레이블").End(xlDown).Offset(1)
        With r저장위치
            .Offset(0, 1).Value = 관
            .Offset(0, 2).Value = 항
            .Offset(0, 3).Value = 목
            .Offset(0, 4).Value = 세목
                
            ws.Range("A4:G" & r저장위치.Row).Sort Key1:=ws.Range("B7"), Order1:=xlAscending, Key2:=ws.Range("C7") _
                , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
                :=xlSortNormal
            
            Call 예산서코드입력
            accountCode = .Value
        End With
        
        new_code = accountCode
    End If
End Function

Sub 결산데이터갱신()
    Dim ws_data As Worksheet
    Dim ws_source As Worksheet
    Set ws_data = Worksheets("data")
    Set ws_source = Worksheets("회계원장")
    Dim endLine As Integer
    Dim endDataLine As Integer
    
    endLine = ws_source.Range("일자필드레이블").End(xlDown).Row
    endDataLine = 0
    Dim i, j As Integer
    
    For i = 6 To endLine
        If (ws_source.Range("b" & j).Value = "") Then
            endDataLine = i - 1
            Exit For
        End If
    Next i

    'Dim 시작행, 종료행 As Integer
    
    '시작행 = 6
    '종료행 = endDataLine
     
    On Error Resume Next
     
    Dim startDate, endDate, accountStartDate, today As Date
        
    startDate = ws_data.Range("j1").Value
    endDate = ws_data.Range("j2").Value
    accountStartDate = format(get_config("accountStartDate"), "m/d/yyyy")
    today = format(Date, "m/d/yyyy")
    
    
    ' #1. 전체 기간 피벗 테이블 생성
    Call 결산피벗테이블생성(accountStartDate, today)
    
    ' #2. 기간별 생성
    Call 결산피벗테이블생성(startDate, endDate)
    
    ' #3. 생성 후 마무리. 참고 날짜 등 표시
    With ws_data.Range("i1")
        .Value = "시작일"
        .Offset(1, 0).Value = "종료일"
        .Offset(2, 0).Value = "회계시작일"
        .Offset(3, 0).Value = "오늘"
        .Offset(0, 1).Value = startDate
        .Offset(1, 1).Value = endDate
        .Offset(2, 1).Value = accountStartDate
        .Offset(3, 1).Value = today

    End With

End Sub

Function 코드결산(ByVal accountCode As String, ByVal 수입지출 As String, ByVal 부분전체 As String)
    Dim ws_data As Worksheet
    Set ws_data = Worksheets("data")
    Dim c As Range
    Dim 열이동 As Integer
    
    If 수입지출 = "수입" Then
        열이동 = 1
    Else
        열이동 = 2
    End If

    If 부분전체 = "부분" Then
        With ws_data.Range("e1").CurrentRegion.columns(1)
            For Each c In .Cells
                If c.Value = accountCode Then
                    코드결산 = c.Offset(0, 열이동).Value
                    Exit For
                End If
            Next c
        End With
        If Not 코드결산 > 0 Then
            코드결산 = 0
        End If
    Else
        With ws_data.Range("a1").CurrentRegion.columns(1)
            For Each c In .Cells
                If c.Value = accountCode Then
                    코드결산 = c.Offset(0, 열이동).Value
                    Exit For
                End If
            Next c
        End With
        If Not 코드결산 > 0 Then
            코드결산 = 0
        End If
    End If

End Function

Public Function get_config(ByVal item As String)
    With Worksheets("설정")
        If item = "시작일" Then
            get_config = .Range("작업시작일설정").Offset(0, 1).Value
        ElseIf item = "종료일" Then
            get_config = .Range("작업종료일설정").Offset(0, 1).Value
        ElseIf item = "회계시작일" Then
            get_config = .Range("회계시작일설정").Offset(0, 1).Value
        ElseIf item = "기관명" Then
            get_config = .Range("기관명설정").Offset(0, 1).Value
        Else
            get_config = ""
        End If
    End With
End Function

Public Sub set_config(ByVal item As String, ByVal newValue As String)
    With Worksheets("설정")
        If item = "시작일" Then
            .Range("작업시작일설정").Offset(0, 1).Value = newValue
        ElseIf item = "종료일" Then
            .Range("작업종료일설정").Offset(0, 1).Value = newValue
        ElseIf item = "회계시작일" Then
            .Range("회계시작일설정").Offset(1, 0).Value = newValue
        ElseIf item = "기관명" Then
            .Range("기관명설정").Offset(0, 1).Value = newValue
        Else

        End If
    End With
End Sub

Sub 결산피벗테이블생성(ByVal startDate As Date, ByVal endDate As Date)
    Application.ScreenUpdating = False

    Dim ws_source As Worksheet, ws_source_copy As Worksheet, ws_data As Worksheet
    Dim rng_source As Range
    Dim pvt As PivotTable
    Dim pvtName As String, 생성위치 As String
    Dim startRow As Integer, endRow As Integer
    Dim endLine As Integer, endDataLine As Integer
    Dim i As Integer
    Dim dataCount As Integer
    
    Dim accountStartDate As Date, today As Date
    accountStartDate = get_config("회계시작일")
    today = Date
            
    Set ws_source = Worksheets("회계원장")
    Set ws_source_copy = Worksheets("항목분계장")
    Set ws_data = Worksheets("data")
    
    endLine = ws_source.Range("일자필드레이블").End(xlDown).Row
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
    
    Dim 날짜비교 As Integer
    Dim 전체or부분 As String
    Dim c As Range
    Set c = ws_source.Range("A5")
    
    ws_source.Unprotect PWD
    'On Error GoTo ErrorHandler
    
    '시작행과 종료행을 정확히 지정하는 것이 관건. 불안정한 주 원인이 여기다.
    '속도 개선과 정확한 판단을 위해 셀을 하나씩 확인하는 것이 아니라 배열을 이용
    Dim dateArray() As Variant
    ReDim dateArray(1 To dataCount)
    
    'dateArray = ws_source.Range("A6:A" & endDataLine).Value '이 방식이 안된다.
    'dateArray = ws_source.Range(c, c.Offset(dataCount - 1, 0)).Value '이것도 get_code와 달리 안됨
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
        If dateArray(i) = endDate Then
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
    
    If startDate = accountStartDate And endDate = today Then
        ws_data.Range("L2:t" & 최대행수).ClearContents
        Call 원장데이터copy("data", 6, endDataLine)
        Set rng_source = ws_data.Range("L1").Resize(endDataLine - 5 + 1, 9) '행 라벨도 포함되므로 + 1
        
        pvtName = "전기간결산"
        생성위치 = "data!R1C1"
    Else  '특정 기간을 지정한 경우 해당 기간에 대한 통계 만듬
        Dim 레코드수 As Integer
        레코드수 = endRow - startRow + 1
        
        ws_data.Range("L2:t" & 최대행수).ClearContents
        Call 원장데이터copy("data", startRow, endRow)
        Set rng_source = ws_data.Range("L1").Resize(레코드수 + 1, 9) '행 라벨도 포함되므로 + 1
        pvtName = "특정기간결산"
        생성위치 = "data!R1C5"
    End If
    
    Dim prevStartDate As String
    Dim prevEndDAte As String
    prevStartDate = ws_data.Range("이전시작일저장").Value
    prevEndDAte = ws_data.Range("이전종료일저장").Value
        
    If startDate <> prevStartDate Or endDate <> prevEndDAte Then  '직전 조회할때와 다른 기간을 설정했다면 테이블 재생성
        
        '기존 피벗테이블 지움
        For Each pvt In ws_data.PivotTables
            If Not pvt Is Nothing Then
                If pvt.name = pvtName Then
                    pvt.TableRange2.Clear
                    Exit For
                End If
            End If
        Next
        
        '피벗테이블 생성
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            rng_source).CreatePivotTable _
            TableDestination:=생성위치, TableName:=pvtName
        
        With ws_data.PivotTables(pvtName)
            .PivotFields("code").Orientation = xlRowField
            .PivotFields("code").position = 1
            .AddDataField .PivotFields("수입"), "합계: 수입", xlSum 'xlCount
            .AddDataField .PivotFields("지출"), "합계: 지출", xlSum
        End With
        
    Else  '직전 조회할 때와 같은 기간을 설정했다면 값만 refresh
        ws_data.PivotTables(pvtName).PivotCache.refresh
        
    End If
    
    ' 다음을 위해 시작일과 종료일 저장 (이전시작일, 이전종료일로 활용)
    ws_data.Range("이전시작일저장").Value = startDate
    ws_data.Range("이전종료일저장").Value = endDate
    
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MsgBox "오류의 원인은  " & vbCrLf & Err.Number & vbCrLf & Err.Description & Space(10), vbCritical, "오류 종류 확인"
'    End If
    
End Sub

Function date_compare(ByVal startDate As Date, ByVal endDate As Date)
    'datediff 함수는 뒷 날짜에서 앞 날짜를 뺀다. 즉 결과가 양수이면 앞 날짜가 빠르고, 음수이면 앞 날짜가 느리다
    date_compare = DateDiff("d", startDate, endDate)
End Function

Sub init_firstday()
    '회계시작일 초기화
    Dim accountStartDate As String
    accountStartDate = Year(Now) & "-01-01"
    Worksheets("설정").Range("회계시작일설정").Offset(0, 1).Value = accountStartDate
    UserForm_설정.TextBox_회계시작일.Value = accountStartDate
End Sub

Sub 원장데이터copy(ByVal targetSheet As String, ByVal startRow As Integer, ByVal endRow As Integer, Optional ByVal tag As String)
    Application.ScreenUpdating = False
    
    Dim ws_source As Worksheet
    Dim ws_data As Worksheet
    Set ws_source = Worksheets("회계원장")
    Set ws_data = Worksheets(targetSheet)
    
    Dim copy_point As Range
    
    If targetSheet = "data" Then
        Set copy_point = ws_data.Range("L2")
    ElseIf targetSheet = "항목분계장" Then
        Set copy_point = ws_data.Range("a6")
    End If
    'MsgBox "test" & targetSheet & " startRow: " & startRow & " endRow: " & endRow
    If project <> "" Then
        tag = project
    End If
    
    If tag <> "" Then
        Dim 전체 As Range
        Set 전체 = ws_source.Range("A" & startRow & ":N" & endRow)
        Dim i As Integer
        Dim 타겟행 As Integer
        타겟행 = copy_point.Row
        Dim 일자 As String, code As String, 관 As String, 항 As String, 목 As String, 세목 As String, 적요 As String
        Dim 수입 As Long, 지출 As Long
        
        For i = startRow To endRow
            If ws_source.Range(회계원장_열_프로젝트 & i).Value = tag Then
                With ws_source
                    일자 = .Range(회계원장_열_일자 & i).Value
                    code = .Range(회계원장_열_code & i).Value
                    관 = .Range(회계원장_열_관 & i).Value
                    항 = .Range(회계원장_열_항 & i).Value
                    목 = .Range(회계원장_열_목 & i).Value
                    세목 = .Range(회계원장_열_세목 & i).Value
                    적요 = .Range(회계원장_열_적요 & i).Value
                    수입 = .Range(회계원장_열_수입 & i).Value
                    지출 = .Range(회계원장_열_지출 & i).Value
                End With
                    
                With copy_point
                    .Value = 일자
                    .Offset(0, 1).Value = code
                    .Offset(0, 2).Value = 관
                    .Offset(0, 3).Value = 항
                    .Offset(0, 4).Value = 목
                    .Offset(0, 5).Value = 세목
                    .Offset(0, 6).Value = 적요
                    .Offset(0, 7).Value = 수입
                    .Offset(0, 8).Value = 지출
                End With
                Set copy_point = copy_point.Offset(1, 0)
            End If
        Next i
    
    Else
        '속도 개선을 위해 copy&paste 방식에서 range 간 직접 데이터 복사 방식으로 바꿈 2016.9.18
        'ws_source.Range("A" & startRow & ":a" & endRow).Copy
        'copy_point.PasteSpecial xlPasteValues
        If targetSheet = "data" Then
            ws_data.Range("L1").Value = "날짜"
            ws_data.Range("M1").Value = "code"
            ws_data.Range("N1").Value = "관"
            ws_data.Range("O1").Value = "항"
            ws_data.Range("P1").Value = "목"
            ws_data.Range("Q1").Value = "세목"
            ws_data.Range("R1").Value = "적요"
            ws_data.Range("S1").Value = "수입"
            ws_data.Range("T1").Value = "지출"
            'data 시트 필드 이름 부분 초기화 코드 추가 - 2016.11.9
            
            ws_data.Range("L2:L" & endRow - startRow + 2).Value = ws_source.Range("A" & startRow & ":a" & endRow).Value
            ws_data.Range("M2:T" & endRow - startRow + 2).Value = ws_source.Range("C" & startRow & ":" & 회계원장_열_지출 & endRow).Value
        ElseIf targetSheet = "항목분계장" Then
            ws_data.Range("A6:A" & endRow - startRow + 6).Value = ws_source.Range("A" & startRow & ":a" & endRow).Value
            ws_data.Range("B6:I" & endRow - startRow + 6).Value = ws_source.Range("C" & startRow & ":" & 회계원장_열_지출 & endRow).Value
        End If
        
        'ws_source.Range("c" & startRow & ":" & 회계원장_열_지출 & endRow).Copy
        'copy_point.Offset(0, 1).PasteSpecial xlPasteValues

        Application.CutCopyMode = False
    End If
End Sub

Sub 계정과목가져오기(계정과목시트 As String, 분류 As String)
    Dim ws_source As Worksheet
    Dim ws_target As Worksheet
    Set ws_source = Worksheets(계정과목시트)
    Set ws_target = Worksheets("예산서")
    '예산서 초기화
    Call UserForm_설정.회계설정초기화("계정과목")
    
    '계정과목을 유형별로 가져와서 예산서 시트에 넣는 부분
    Dim 전체 As Range
    Dim x As Integer, y As Integer
    Dim 기준열 As Range, 관열 As Range, 항열 As Range, 목열 As Range, 세목열 As Range
    
    Set 관열 = ws_source.Range("샘플관열라벨")
    Set 항열 = ws_source.Range("샘플항열라벨")
    Set 목열 = ws_source.Range("샘플목열라벨")
    Set 세목열 = ws_source.Range("샘플세목열라벨")
    Set 기준열 = ws_source.Range("샘플분류열라벨")
    x = 0

    Set 전체 = ws_source.Range("샘플관열라벨").CurrentRegion.columns(기준열.Column)
    Dim rowCount As Integer
    rowCount = 전체.Rows.Count
    Dim 타겟행 As Integer
    타겟행 = ws_target.Range("A1").End(xlDown).Offset(1, 0).Row
    Dim 분류값 As String
    Dim i As Integer
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    
    For i = 1 To rowCount
        분류값 = ws_source.Range("A" & i).Offset(, 기준열.Column - 1).Value
        If 분류값 = "공통" Or 분류값 = 분류 Then
            With ws_source.Range("A" & i)
                관 = .Offset(, 관열.Column - 1).Value
                항 = .Offset(, 항열.Column - 1).Value
                목 = .Offset(, 목열.Column - 1).Value
                세목 = .Offset(, 세목열.Column - 1).Value
            End With
            
            With ws_target.Range("A" & 타겟행)
                .Offset(, 1).Value = 관
                .Offset(, 2).Value = 항
                .Offset(, 3).Value = 목
                .Offset(, 4).Value = 세목
            End With
            타겟행 = 타겟행 + 1
        End If
    Next i
    
    '마무리
    code_changed = True
    Call 계정과목_시트정렬
    
    Call 예산서코드입력
    Call 결산서초기화
End Sub

Sub 계정과목통합()
    Dim ws_source As Worksheet
    Dim ws_target As Worksheet
    Set ws_source = Worksheets("가져오기2")
    Set ws_target = Worksheets("예산서")
    Dim start As Range, 끝 As Range
    Dim columnCount As Integer, rowCount As Integer
    Dim 관 As String, 항 As String, 목 As String, 세목 As String, accountCode As String
    Dim x(4, 2) As Variant
    Dim 선택열 As Integer
    Dim i As Integer
    i = 0

    On Error GoTo error

    With ws_source
        Set start = .Range("a2")
        columnCount = start.CurrentRegion.columns.Count
        rowCount = start.CurrentRegion.Rows.Count - 2
        
        If rowCount = 0 Then
            MsgBox "세번째 줄에 데이터를 복사해주세요"
            ws_source.Range("A3").Select
            Exit Sub
        End If
        
        Set 끝 = start.Offset(0, columnCount - 1)
        Dim label As Range
        For Each label In Range(start, 끝)
            If label.Value <> "" Then
                x(i, 0) = label.Value
                x(i, 1) = label.Column
                i = i + 1
                If label.Value = "선택" Then
                    선택열 = label.Column
                End If
            End If
        Next label
    End With

    If x(3, 0) = "" Then
        MsgBox "라벨이 부족합니다. 관, 항, 목, 세목까지 4개의 라벨이 필요합니다"
        start.Select
        Exit Sub
    Else
        Dim 라벨값 As String
        Dim 열 As Integer
        Dim 타겟행 As Integer
        타겟행 = ws_target.Range("관항목코드레이블").End(xlDown).Row + 1
        Dim 타겟열 As Integer
        Dim 첫행, 끝행 As Range
        Dim 타겟시작, 타겟끝 As Range
        Dim j As Integer
        Dim codeCount As Integer
        codeCount = 0
        
        Application.DisplayStatusBar = True
        
        ' #5 각 열 순회하며 복사 & 회계원장에 붙이기
        For i = 0 To UBound(x)
            라벨값 = x(i, 0)
            열 = x(i, 1)

            If Not 열 > 0 Then
                Exit For
            End If
            
            With start.Offset(0, 열 - 1)
                Set 첫행 = .Offset(1, 0)
                Set 끝행 = .Offset(rowCount, 0)

                With ws_target.Range("관항목코드레이블")
                    For j = 1 To 5
                        If 라벨값 = .Offset(0, j).Value Then
  
                            With ws_target.Range("A" & 타겟행)
                                Set 타겟시작 = .Offset(0, j)
                                Set 타겟끝 = .Offset(rowCount - 1, j)
                            End With
                            
                            If Not IsEmpty(ws_source.Range(첫행, 끝행).Value) Then
                                ws_target.Range(타겟시작, 타겟끝).Value = ws_source.Range(첫행, 끝행).Value
                                If codeCount = 0 Then
                                    codeCount = rowCount - Application.WorksheetFunction.CountIf(Range(첫행, 끝행), "")
                                End If
                            End If
                            
                            Exit For
                        End If
                    Next j
                End With
            End With
        Next i
        
        If codeCount > 0 Then
            MsgBox codeCount & "개의 관항목이 예산서로 복사되었습니다."
            ws_target.Activate
            ' # 정렬하고, 코드 부여한 후, 결산서 초기화
            Application.StatusBar = "가져온 계정과목에 고유 코드를 부여하고 있습니다"
            Call 계정과목_시트정렬
            Call 예산서코드입력
            Application.StatusBar = "확정된 계정과목에 따라 결산서를 초기화하고 있습니다"
            Call 결산서초기화
        Else
            MsgBox "입력되지 않았습니다."
        End If
    End If
    
    Erase x
    Application.DisplayStatusBar = False
    
error:
    If Err.Number <> 0 Then
        MsgBox "오류번호 : " & Err.Number & vbCr & _
        "오류내용 : " & Err.Description, vbCritical, "오류"
    End If
    
End Sub

Sub 원장데이터통합()
    ' #4 첫 행에 모든 라벨 있는지 확인
    Dim ws_source, ws_target As Worksheet
    Set ws_source = Worksheets("가져오기")
    Set ws_target = Worksheets("회계원장")
    Dim start, 끝 As Range
    Dim columnCount As Integer, rowCount As Integer
    Dim 관 As String, 항 As String, 목 As String, 세목 As String
    Dim 관열 As Integer '값 검증할때, 관이 예산외수입 혹은 예산외지출인지 확인하기 위해 특별히 필요
    Dim accountCode As String
    Dim x(8, 2) As Variant
    Dim i, j As Integer
    i = 0
    
    ws_target.Unprotect PWD
    
    With ws_source
        Set start = .Range("a2") 'a1에는 사용법 설명이 적혀 있다
        columnCount = start.CurrentRegion.columns.Count
        rowCount = start.CurrentRegion.Rows.Count - 2 '첫행(사용법 설명) 제외
        
        If rowCount = 0 Then
            MsgBox "세번째 줄에 데이터를 복사해주세요"
            ws_source.Range("A3").Select
            Exit Sub
        End If
        
        Set 끝 = start.Offset(0, columnCount - 1)
        Dim 라벨 As Range
        For Each 라벨 In Range(start, 끝)
            If i > 8 Then
                Exit For
            End If
            If 라벨.Value <> "" Then
                x(i, 0) = 라벨.Value
                x(i, 1) = 라벨.Column
                i = i + 1
                If 라벨.Value = "관" Then
                    관열 = 라벨.Column
                End If
            End If
        Next 라벨
    End With
    
    If x(8, 0) = "" Then
        MsgBox "라벨이 부족합니다. 일자, 관, 항, 목, 세목, 수입, 지출, 적요, 은/현까지 9개의 라벨이 필요합니다"
        start.Select
        Exit Sub
    Else
        Dim 라벨값 As String
        Dim 열 As Integer
        Dim 타겟행, 타겟열 As Integer
        타겟행 = ws_target.Range("일자필드레이블").End(xlDown).Row + 1
        Dim 첫행, 끝행 As Range
        Dim 타겟시작, 타겟끝 As Range
        Dim 기준점 As Range
        
        ' #4-1. 데이터 검증하기
        For i = 0 To UBound(x)
            라벨값 = x(i, 0)
            열 = x(i, 1)
            ws_source.Range("f1").Value = 열
            If Not 열 > 0 Then '모든 열을 다 돌았거나 초기화가 안된 경우 종료
                Exit For
            End If
            
            If Not IsError(Application.Match(라벨값, Array("관", "항", "목", "세목"), False)) Then
                With start.Offset(0, 열 - 1)
                    For j = 1 To rowCount
                        If IsEmpty(.Offset(j, 0).Value) Then
                            Select Case start.Offset(j, 관열 - 1).Value
                                Case "예산외수입", "예산외지출"
                                    'MsgBox start.Offset(j, 관열 - 1).Value
                                    'Exit Sub
                                    'pass
                                Case Else
                                    MsgBox "누락된 값이 있어 데이터통합을 진행할 수 없습니다" & rowCount
                                    .Offset(j, 0).Select
                                    Exit Sub
                            End Select
                        End If
                    Next j
                End With
            End If
        Next i
        
        MsgBox "데이터 검증을 마쳤습니다. 확인을 누르면 회계원장에 통합합니다"
        
        Application.DisplayStatusBar = True
        Application.StatusBar = "데이터를 회계원장에 복사하고 있습니다"
        
        ws_target.Unprotect PWD
        
        ' #5 각 열 순회하며 복사 & 회계원장에 붙이기
        For i = 0 To UBound(x)
            라벨값 = x(i, 0)
            열 = x(i, 1)
            ws_source.Range("f1").Value = 열
            If Not 열 > 0 Then '모든 열을 다 돌았거나 초기화가 안된 경우 종료
                Exit For
            End If
            
            With start.Offset(0, 열 - 1)
                Set 첫행 = .Offset(1, 0)
                Set 끝행 = .Offset(rowCount, 0)
                Set 기준점 = ws_target.Range("일자필드레이블")
                
                For j = 0 To 17 'columnCount?
                    If 라벨값 = 기준점.Offset(0, j).Value Then
                        With ws_target.Range("A" & 타겟행)  '회계원장의 가장 마지막 줄에 추가
                            Set 타겟시작 = .Offset(0, j)
                            Set 타겟끝 = .Offset(rowCount - 1, j)
                        End With
                        ws_target.Range(타겟시작, 타겟끝).Value = ws_source.Range(첫행, 끝행).Value
                        Exit For
                    End If
                Next j

            End With
        Next i
        Application.DisplayStatusBar = False

        ws_target.Activate
    
    End If
    
    Erase x
    
    Application.StatusBar = "새로운 관항목을 생성하고 있습니다"
    
    '#6 추가된 부분 순회하며 관항목 검증 : 기존에 등록된 코드인지 확인
    Dim r저장위치 As Range
    Dim ws_예산서 As Worksheet
    Set ws_예산서 = Worksheets("예산서")
    Dim k As Integer
    
    '#6-1. 새 관항목이 있는지 확인
    '#6-2. 새 관항목 생성
    '회계원장 시트로 직전에 가져온 관항목들을 예산서 시트에 복사 -> 중복을 제거 -> 정렬 -> 코드 재부여
    Dim 예산서_시작행 As Integer
    예산서_시작행 = ws_예산서.Range("관필드").End(xlDown).Row + 1
    
    Dim 예산_저장위치 As Range
    Set 예산_저장위치 = ws_예산서.Range("B" & 예산서_시작행 & ":E" & 예산서_시작행 + rowCount - 1)
    예산_저장위치.Value = ws_target.Range("D" & 타겟행 & ":G" & 타겟행 + rowCount - 1).Value
    Set 예산_저장위치 = ws_예산서.Range("B2:E" & 예산서_시작행 + rowCount)
    예산_저장위치.RemoveDuplicates columns:=Array(1, 2, 3, 4), Header:=xlNo
    
    ws_예산서.Range("A4:G" & Cells(Rows.Count, "A").End(xlDown).Row).Sort Key1:=ws_예산서.Range("B7"), Order1:=xlAscending, Key2:=ws_예산서.Range("C7") _
                , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
                :=xlSortNormal
                
    Set 예산_저장위치 = Nothing
    Application.CutCopyMode = False
            
    Call 예산서코드입력
    
    '#6-3. 회계원장에 관항목 입력
    '이 부분이 느림. 배열로 값을 가져와 순회하게 변경
    Application.StatusBar = "회계원장에 새 관항목을 적용하고 있습니다"
    ws_target.Unprotect PWD
    codeArrayInitialized = False 'get_code 함수가 새로운 계정과목 코드를 가져오도록 다시 초기화해야 한다고 알림
    
    Dim subject() As Variant
    ReDim subject(1 To rowCount, 1 To 4)
    Dim newCode() As Variant
    ReDim newCode(1 To rowCount, 1 To 2)
    
    Dim rNewTarget As Range
    Set rNewTarget = ws_target.Range("D" & 타겟행)
    
    subject = ws_target.Range(rNewTarget, rNewTarget.Offset(rowCount, 3)).Value
    Dim point As Integer
    point = rowCount / 10
    
    Application.StatusBar = "회계원장에 새 관항목을 적용하고 있습니다"
    With ws_target
        For k = 1 To rowCount
            'If k Mod point = 0 Then
            '    'Application.StatusBar = "회계원장에 새 관항목을 적용하고 있습니다 (" & k & " / " & rowCount & ")"
            '    Application.StatusBar = "회계원장에 새 관항목을 적용하고 있습니다 (" & k * 100 / rowCount & "%)"
            'End If
            관 = subject(k, 1)
            항 = subject(k, 2)
            목 = subject(k, 3)
            세목 = subject(k, 4)
            
            accountCode = get_code(관, 항, 목, 세목)
            
            newCode(k, 1) = accountCode & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
            newCode(k, 2) = accountCode
            
            'With .Range("A" & 타겟행 + k - 1)
            '    If accountCode <> "" Then
            '        .Offset(0, 1).Value = accountCode & "/" & 관 & "/" & 항 & "/" & 목 & "/" & 세목
            '        '.Offset(0, 2).FillDown
            '        '.Offset(0, 2).NumberFormat = "@" '이 부분이 느린 거였다! 아래에서 일괄되게 수행하게 변경 2016.9.20
            '        .Offset(0, 2).Value = accountCode
            '    End If
            'End With
        Next k
        
        .Range("B" & 타겟행, "C" & 타겟행 + rowCount - 1).Value = newCode
        .Range("C" & 타겟행 & ":C" & 타겟행 + rowCount - 1).NumberFormat = "@"
    End With
    
    ' #7 회계원장 정렬
    Application.StatusBar = "회계원장 기입을 마무리하고 있습니다"
    Dim endLine As Integer
    endLine = ws_target.Range("일자필드레이블").End(xlDown).Row

    ws_target.Range("A8:O" & endLine).Sort Key1:=ws_target.Range("A7"), Order1:=xlAscending, Key2:=ws_target.Range("B7") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal

    ' #8 검증
    ' 은/현이 비어 있으면 0으로 채우기 (통장 거래 우선)
    With ws_target
        For k = 8 To endLine
            If IsEmpty(.Range(회계원장_열_은현 & k).Value) Then
                .Range(회계원장_열_은현 & k).Value = 0
            End If
        Next k
    End With
    
    If Worksheets("설정").Range("a2").Offset(, 1).Value = True Then
        ws_target.Protect PWD
    End If
    ws_target.Activate
    MsgBox rowCount & "건의 데이터가 통합되었습니다."
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    
error:
    If Err.Number <> 0 Then

        MsgBox "오류번호 : " & Err.Number & vbCr & _
        "오류내용 : " & Err.Description, vbCritical, "오류"

    End If
End Sub

Sub 원장데이터가져오기_준비()
    Dim ws As Worksheet
    Set ws = Worksheets("가져오기")
    Dim 기준점 As Range
    Set 기준점 = ws.Range("A2")
    
    Dim 날짜1 As String
    Dim 날짜2 As Date
    
    '우리은행 데이터인지 확인
    'A3부터 No. 거래일시 적요 기재내용 찾으신금액(혹은 지급(원)), 맡기신금액(혹은 ), 거래후잔액, 취급점
    Dim 라벨행 As Integer
    라벨행 = 0
    If 기준점.Value = "No." Then
        라벨행 = 2
    ElseIf 기준점.Offset(1, 0).Value = "No." Then
        라벨행 = 3
    End If
    
    If Not 라벨행 > 0 Then
        MsgBox "우리은행 자료를 A3 위치에 다시 복사해주세요"
        Exit Sub
    End If
    
    '2번째 행 (A2~) 삭제하고 끌어올림
    If 라벨행 = 3 Then
        기준점.EntireRow.Delete Shift:=xlUp
        Set 기준점 = ws.Range("a2")
    End If
    
    If 기준점.Offset(1, 0).Value = "" Then
        MsgBox "데이터가 없습니다"
        Exit Sub
    End If
    
    'No. 삭제하고 shift left, 일자 형식 변환
    Dim startRow As Integer, endRow As Integer
    startRow = 2
    endRow = 기준점.End(xlDown).Row
    
    ws.Range("a" & startRow, "a" & endRow).Select
    Selection.Delete Shift:=xlToLeft
    
    '일시 오른쪽에 4열 추가 & 관, 항, 목, 세목 라벨 표시
    With ws.Range("B2:F" & endRow)
        .Insert Shift:=xlToRight
        .NumberFormatLocal = "@"
    End With
    
    Set 기준점 = ws.Range("A2")
    
    With 기준점
        .Value = "일자"
        .Offset(, 1).Value = "관"
        .Offset(, 2).Value = "항"
        .Offset(, 3).Value = "목"
        .Offset(, 4).Value = "세목"
        .Offset(, 5).Value = "은/현"
    End With
    
    '세목 오른쪽에 "은/현" 라벨 추가. 그 아래 모두 0으로 채움
    ws.Range("F3").Value = 0
    ws.Range("F3:F" & endRow).FillDown
    columns("F:F").EntireColumn.AutoFit
    
    '날짜 형식 변경 (일자)
    Dim i As Integer
    Dim r타겟위치 As Range
    For i = startRow + 1 To endRow
        Set r타겟위치 = ws.Range("A" & i)
        With r타겟위치
            .Value = Replace(Left(.Value, 10), ".", "-")
        End With
    Next i
    columns("A:A").EntireColumn.AutoFit
    
    Dim 수입열, 지출열 As Integer
    
    '적요 -> 구분, 기재내용 -> 적요로 라벨변경
    '찾으신금액(지급(원)) -> 지출 / 맡기신금액 -> 수입
    ws.Range(기준점, 기준점.End(xlToRight)).Select
    
    Selection.Replace What:="적요", Replacement:="구분", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="기재내용", Replacement:="적요", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="지급(원)", Replacement:="지출", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="찾으신금액", Replacement:="지출", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="입금(원)", Replacement:="수입", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Replace What:="맡기신금액", Replacement:="수입", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Selection.Find(What:="수입", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
    수입열 = ActiveCell.Column
    
    '관 열에서 수입금액이면 '수입', 지출금액이면 '지출'로 표시
    For i = startRow + 1 To endRow
        With ws.Range("B" & i)
            If .Offset(, 수입열 - 2).Value > 0 Then
                .Value = "수입"
            Else
                .Value = "지출"
            End If
        End With
    Next i

End Sub
