Attribute VB_Name = "Module8"
'module8 : 품의서, 지출결의서
Option Explicit

Sub 지출결의대장정렬()
    Dim ws_target As Worksheet
    Set ws_target = Worksheets("지출결의대장")
    Dim 끝줄 As Integer
    끝줄 = ws_target.Range("결의날짜레이블").End(xlDown).Row
    MsgBox 끝줄
    
    ws_target.Unprotect

    ws_target.Range("A4:I" & 끝줄).Sort Key1:=ws_target.Range("A3"), Order1:=xlAscending, Key2:=ws_target.Range("B3") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom

End Sub

Sub 지출결의서인쇄_Click()
    Dim ws As Worksheet
    Set ws = Worksheets("지출결의서")
    ws.PageSetup.PrintArea = "$b$3:$l$24"
    ActiveWindow.SelectedSheets.PrintPreview

End Sub

Sub 지출결의_회계원장입력()
    Dim ws As Worksheet
    Set ws = Worksheets("지출결의서")

    Dim 항 As String
    Dim 목 As String
    Dim 세목 As String
    Dim 회계일 As String
    Dim 코드 As String
    Dim 금액 As Long
    Dim 금액2 As Long
    Dim 적요 As String
    On Error Resume Next
    With ws
        .Range("t2").Select

        항 = .Range("지출항").Value
        목 = .Range("지출목").Value
        세목 = .Range("지출세목").Value
        
        회계일 = .Range("a1")
        코드 = .Range("b1")
        금액 = .Range("i18")
        금액2 = .Range("i10")
        
        If 금액2 > 0 Then
            적요 = .Range("b9") & " 등"
        Else
            적요 = .Range("b9")
        End If
    
        .Range("A1").Select
    End With
    
    '회계원장에 입력
    Dim ws_target As Worksheet
    Dim r저장위치 As Range
    Set ws_target = Worksheets("회계원장")
    Set r저장위치 = ws_target.Range("A5").End(xlDown)
    
    With r저장위치
        .Offset(1, 0).FormulaR1C1 = 회계일
        .Offset(1, 1).FormulaR1C1 = 코드
        .Offset(1, 3).FormulaR1C1 = "지출"
        .Offset(1, 4).FormulaR1C1 = 항
        .Offset(1, 5).FormulaR1C1 = 목
        .Offset(1, 6).FormulaR1C1 = 세목
        .Offset(1, 7).FormulaR1C1 = 적요
        .Offset(1, 9).FormulaR1C1 = 금액
        .Offset(1, 10).FormulaR1C1 = 0
    End With
    
    If Err.Number = 0 Then
        MsgBox "입력됐습니다"
    End If
        
End Sub

Sub 회계원장_지출결의입력(행번호 As Integer)
    Dim ws As Worksheet
    Set ws = Worksheets("회계원장")

    Dim 회계일 As String
    Dim 코드 As String
    Dim 금액 As Long
    Dim 적요 As String
    
    On Error Resume Next
    With ws.Range("A" & 행번호)
        회계일 = .Value
        코드 = .Offset(, 1).Value
        적요 = .Offset(, 7).Value
        금액 = .Offset(, 9).Value

    End With
    
    If 금액 > 0 Then
        '지출결의대장에 입력
        Dim ws_target As Worksheet
        Dim r저장위치 As Range
        Set ws_target = Worksheets("지출결의대장")
        If ws_target.Range("결의날짜레이블").Offset(1, 0).Value Then
            Set r저장위치 = ws_target.Range("결의날짜레이블").End(xlDown).Offset(1, 0)
        Else
            Set r저장위치 = ws_target.Range("결의날짜레이블").Offset(1, 0)
        End If
        
        With r저장위치
            .FormulaR1C1 = 회계일
            .Offset(, 1).FormulaR1C1 = 코드
            .Offset(, 2).FormulaR1C1 = 적요 '지출명
            .Offset(, 4).Value = 1
            .Offset(, 5).Value = 금액 '단가
            .Offset(, 6).Value = 금액
        End With
        
    Else
        MsgBox "지출결의서를 생성할 수 없습니다"
    End If
        
End Sub

Sub 지출결의서작성(날짜전체 As Boolean)

    Dim 지출명(20)
    Dim 규격(20)
    Dim 수량(20)
    Dim 단가(20)
    Dim 비고(20)
    Dim 하단비고(20)
    Dim ws As Worksheet
    Set ws = Worksheets("지출결의서")
    Dim ws_source As Worksheet
    Set ws_source = Worksheets("지출결의대장")
    
    Dim 현재줄 As Integer
    Dim 날짜 As String, 코드 As String, 오늘 As String, 결의일 As String, 항목코드 As String
    Dim i As Integer, i2 As Integer, i3 As Integer
    
    현재줄 = ActiveCell.Row
    With ws_source
        .Range("a" & 현재줄).Select
        날짜 = .Range("a" & 현재줄).Value
        코드 = .Range("b" & 현재줄).Value
    
        오늘 = 날짜
            
        i2 = 1
            
        For i = 1 To 300
        
            현재줄 = ActiveCell.Row
            결의일 = .Range("a" & 현재줄)
            항목코드 = .Range("b" & 현재줄)
            
            If 날짜 <> 결의일 Then 'Or 코드 <> 항목코드 Then
                Exit For
            Else
            
                지출명(i) = .Range("c" & 현재줄)
                규격(i) = .Range("d" & 현재줄)
                수량(i) = .Range("e" & 현재줄)
                단가(i) = .Range("f" & 현재줄)
                비고(i) = .Range("h" & 현재줄)
                하단비고(i) = .Range("i" & 현재줄)
                
                i2 = i2 + 1
                
            End If
            
            ActiveCell.Offset(1, 0).Range("A1").Select
            If Not 날짜전체 Then
                Exit For
            End If
        
        Next i
    End With
    
    With ws
        .Activate
        
        .Range("b9:g17").Select
        Selection.ClearContents
        
        .Range("b1").Value = 코드
        .Range("b5").Value = 오늘
        
        For i3 = 1 To i2
        
            .Range("b" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = 지출명(i3)
            
            .Range("d" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = 규격(i3)
            
            .Range("f" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = 수량(i3)
            
            .Range("g" & i3 + 8).Select
            ActiveCell.FormulaR1C1 = 단가(i3)
        
        Next i3
        
        .Range("k9").Value = 비고(1)
        .Range("c19").Value = 하단비고(1)
    End With
    Erase 지출명, 규격, 수량, 단가, 비고, 하단비고

    ws.Activate
    ws.Visible = xlSheetVisible
    
End Sub
