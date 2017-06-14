Attribute VB_Name = "Module2"
'module2 : 결산서초기화, 서식정리1, 서식정리2
Option Explicit

Const 열이름_code As String = "A"
Const 열이름_관 As String = "B"
Const 열이름_항 As String = "C"
Const 열이름_목 As String = "D"
Const 열이름_세목 As String = "e"
Const 열이름_예산액 As String = "f"
Const 열이름_기간수입 As String = "g"
Const 열이름_기간지출 As String = "h"
Const 열이름_누적수입 As String = "i"
Const 열이름_누적지출 As String = "j"
Const 열이름_예산잔액 As String = "k"
Const 열이름_비율 As String = "l"

Sub 결산서초기화2()
    '전체 포맷을 다시 만들지는 않고 기간수입/지출, 누적수입/지출 값만 비운다.
    '결산서 재생성할때 부담 줄이기 위한 약식 초기화
    Dim 예산서 As Worksheet
    Dim 결산서 As Worksheet
    Set 예산서 = Worksheets("예산서")
    Set 결산서 = Worksheets("결산서")
    예산서.Activate
    
    Dim 끝줄 As Integer
    끝줄 = 예산서.Range("a4").End(xlDown).Row
    
    결산서.Activate
    
    With Range(열이름_기간수입 & "6:" & 열이름_누적지출 & 끝줄 + 2)
        .ClearContents
    End With
End Sub

Sub 결산서초기화()
Attribute 결산서초기화.VB_ProcData.VB_Invoke_Func = "i\n14"

    Dim 예산서 As Worksheet
    Dim 결산서 As Worksheet
    Set 예산서 = Worksheets("예산서")
    Set 결산서 = Worksheets("결산서")
    예산서.Activate
    
    Dim 끝줄 As Integer
    끝줄 = 예산서.Range("a4").End(xlDown).Row
    
    결산서.Activate
   
    With Range("A6:" & 열이름_비율 & "200")
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
        
    Range("A6:A180").NumberFormatLocal = "G/표준"

    Dim i As Integer
    With 결산서
        For i = 6 To 끝줄 + 2
            Application.ScreenUpdating = False
            With .Range("a" & i)
                .FormulaR1C1 = "=예산서!R[-2]C" '코드
                .Offset(0, 1).FormulaR1C1 = "=예산서!R[-2]C"  '관
                .Offset(0, 2).FormulaR1C1 = "=예산서!R[-2]C"  '항
                .Offset(0, 3).FormulaR1C1 = "=예산서!R[-2]C"  '목
                .Offset(0, 4).FormulaR1C1 = "=예산서!R[-2]C"  '세목
                .Offset(0, 5).FormulaR1C1 = "=예산서!R[-2]C"  '예산액
                .Offset(0, 10).FormulaR1C1 = "=RC[-5]-RC[-2]-RC[-1]" '예산잔액
                .Offset(0, 11).FormulaR1C1 = "=If(RC[-6] > 0, (RC[-6]-RC[-1])/RC[-6], 0)"
            End With
        Next i
        
        Dim i3 As Integer

        i3 = .Range("a" & 끝줄 + 2).Row
        
        .Range("b" & i3 + 1).FormulaR1C1 = "합계"
        
        .Range(열이름_기간수입 & i3 + 1).FormulaR1C1 = "=sum(r[-" & 끝줄 - 3 & "]C:R[-1]C)"
        .Range(열이름_기간지출 & i3 + 1).FormulaR1C1 = "=sum(r[-" & 끝줄 - 3 & "]C:R[-1]C)"
        .Range(열이름_누적수입 & i3 + 1).FormulaR1C1 = "=sum(r[-" & 끝줄 - 3 & "]C:R[-1]C)"
        .Range(열이름_누적지출 & i3 + 1).FormulaR1C1 = "=sum(r[-" & 끝줄 - 3 & "]C:R[-1]C)"
        
        .Range("b" & i3 + 2).FormulaR1C1 = "실잔액"
        .Range("h" & i3 + 2).FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
        .Range("j" & i3 + 2).FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
        
        With .Range("A5:" & 열이름_비율 & 끝줄 + 4)
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

Sub 결산서서식정리1()
 
Dim 끝줄 As Integer
Dim ws As Worksheet
Dim 임시정수값 As Integer
Dim 임시문자열 As String
Dim 기준위치 As Range
Dim 관 As String, 목 As String, 항 As String
Dim 현재줄 As Integer, 관항목수 As Integer

Dim 결산서 As Worksheet
Set 결산서 = Worksheets("결산서")
    결산서.Activate
    결산서.Select
    결산서.Copy After:=Sheets(9)
    Set 결산서 = ActiveSheet

With 결산서
    .Activate
    끝줄 = .Range("A5").End(xlDown).Row
    
    관항목수 = 끝줄 - 5
    If 관항목수 > 30 Then
        ReDim 목삽입(관항목수)
        ReDim 목제목(관항목수)
        ReDim 항삽입(관항목수)
        ReDim 항제목(관항목수)
        ReDim 관삽입(관항목수)
        ReDim 관제목(관항목수)
    End If

    .Range(열이름_관 & "5:" & 열이름_목 & 끝줄).Select

    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With

    .Range("j1").FormulaR1C1 = " 이 시트는 삭제하셔도 됩니다."
    .Range("a6").Select

    목 = Left(ActiveCell.Value, 6)
    목삽입(1) = 6
    목제목(1) = .Range(열이름_목 & "6")
    
    Dim i As Integer
    Dim i2 As Integer
    i2 = 2
    
    For i = 1 To 300
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 6) = "" Then
    
          목삽입(i2) = ActiveCell.Row
          목제목(i2) = Range(열이름_목 & 목삽입(i2))
    
          i2 = i2 + 1
    
            Exit For
        End If
    
        If 목 <> Left(ActiveCell.Value, 6) Then
          목삽입(i2) = ActiveCell.Row
          목제목(i2) = Range(열이름_목 & 목삽입(i2))
    
          i2 = i2 + 1
    
        End If
    
        목 = Left(ActiveCell.Value, 6)
    
    Next i
    
    Dim i3 As Integer
    Dim i4 As Integer
    i4 = 0
    
    For i3 = 1 To i2 - 2
    
        Range("A" & 목삽입(i3) + i4).Select
        Selection.EntireRow.Insert
    
        현재줄 = ActiveCell.Row
        .Range(열이름_목 & 현재줄).Select
        ActiveCell.FormulaR1C1 = 목제목(i3)
        Selection.Font.Bold = True
    
        .Range(열이름_목 & 현재줄).Select
    
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
    
        .Range(열이름_예산액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_예산잔액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,R[+" & 목삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_비율 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
    
        .Range(열이름_목 & 현재줄 & ":" & 열이름_비율 & 현재줄).Select
        With Selection.Interior
            .ColorIndex = 35
        End With
    
        i4 = i4 + 1
    
    Next i3
    
    Erase 목삽입
    Erase 목제목
    
    '항 정리
    
    .Range("a6").Select
    
    항 = Left(ActiveCell.Value, 4)
      
    항삽입(1) = 6
    항제목(1) = Range("b6")
    i2 = 2
    
    For i = 1 To 150
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 4) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
        If Left(ActiveCell.Value, 4) = Empty And 항 = Empty Then
             Exit For
        End If
    
    
        If 항 <> Left(ActiveCell.Value, 4) Then
    
            If 항 <> Left(ActiveCell.Value, 4) Then
    
                항삽입(i2) = ActiveCell.Row
                항제목(i2) = Range(열이름_항 & 항삽입(i2))
                i2 = i2 + 1
            End If
    
        End If
    
        항 = Left(ActiveCell.Value, 4)
    
    Next i
    항삽입(i2) = 항삽입(i2) - 2
    
    i4 = -1
    
    
    For i3 = 2 To i2 - 2
        .Range("A" & 항삽입(i3) + i4).Select
        Selection.EntireRow.Insert
    
        현재줄 = ActiveCell.Row
        .Range(열이름_항 & 현재줄).Select
        ActiveCell.FormulaR1C1 = 항제목(i3)
        Selection.Font.Bold = True
    
        .Range(열이름_항 & 현재줄 & ":" & 열이름_목 & 현재줄).Select
    
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
    
        .Range(열이름_예산액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_예산잔액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 항삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_비율 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
        
        .Range(열이름_항 & 현재줄 & ":" & 열이름_비율 & 현재줄).Select
        With Selection.Interior
            .ColorIndex = 34
        End With
    
        i4 = i4 + 1
    
    Next i3
    
    Erase 항삽입
    Erase 항제목

    '관 정리
    
    .Range("a6").Select
    
    관 = Left(ActiveCell.Value, 2)
    
    관삽입(1) = 6
    관제목(1) = Range("b6")
    i2 = 2
    
    For i = 1 To 150
    
        ActiveCell.Offset(1, 0).Select
    
        If Left(ActiveCell.Value, 2) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
        
        If Left(ActiveCell.Value, 2) = Empty Then
          ActiveCell.Offset(1, 0).Select
        End If
         
        If Left(ActiveCell.Value, 2) = Empty And 관 = Empty Then
             Exit For
        End If
    
    
        If 관 <> Left(ActiveCell.Value, 2) Then
    
            If 관 <> Left(ActiveCell.Value, 2) Then
    
                관삽입(i2) = ActiveCell.Row
                관제목(i2) = Range(열이름_관 & 관삽입(i2))
                i2 = i2 + 1
            End If
    
        End If
    
        관 = Left(ActiveCell.Value, 2)
    
    Next i
    관삽입(i2) = 관삽입(i2) - 2
    
    i4 = -2
    
    
    For i3 = 2 To i2 - 2
        .Range("A" & 관삽입(i3) + i4).Select
        Selection.EntireRow.Insert
    
        현재줄 = ActiveCell.Row
        .Range(열이름_관 & 현재줄).Select
        ActiveCell.FormulaR1C1 = 관제목(i3)
        Selection.Font.Bold = True
    
        .Range(열이름_관 & 현재줄 & ":" & 열이름_목 & 현재줄).Select
    
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
    
    
        .Range(열이름_예산액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_기간지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적수입 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_누적지출 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_예산잔액 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=subtotal(9,r[+" & 관삽입(i3 + 1) - 현재줄 + i4 & "]C:R[+1]C)"
        Selection.Font.Bold = True
    
        .Range(열이름_비율 & 현재줄).Select
        ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
        Selection.Font.Bold = True
        With Selection
            .HorizontalAlignment = xlRight
        End With
        
        .Range(열이름_관 & 현재줄 & ":" & 열이름_비율 & 현재줄).Select
        With Selection.Interior
            .ColorIndex = 37
        End With
     
        i4 = i4 + 1
    
    Next i3
    Erase 관삽입
    Erase 관제목
   
    .Range("b5:" & 열이름_비율 & "5").Select
    With Selection
        .Font.Bold = True
        .Interior.ColorIndex = 33
        .Font.ThemeColor = xlThemeColorLight1
    End With


' 총계 정리

    .Range(열이름_예산액 & "6").End(xlDown).Select

    현재줄 = ActiveCell.Row + 1
    Dim 처음줄 As Integer
    처음줄 = 현재줄 - 5

    .Range(열이름_관 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "총    계"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":d" & 현재줄).Select

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

    .Range(열이름_기간수입 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(열이름_기간지출 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(열이름_누적수입 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    .Range(열이름_누적지출 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":" & 열이름_비율 & 현재줄).Select
    With Selection.Interior
        .ColorIndex = 33
    End With


'잔액 정리
    현재줄 = 현재줄 + 1

    .Range(열이름_관 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "실 잔 액"
    Selection.Font.Bold = True

    .Range(열이름_관 & 현재줄 & ":" & 열이름_목 & 현재줄).Select

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

    .Range(열이름_기간지출 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    .Range(열이름_누적지출 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    .Range(열이름_관 & 현재줄 & ":" & 열이름_비율 & 현재줄).Select
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

    .PageSetup.PrintArea = "$b$2:$k$" & 현재줄

End With

'메모리 지우기
Application.CutCopyMode = False

End Sub


Sub 결산서정리2()

    Dim 항 As String
    Dim i As Integer, i2 As Integer, i3 As Integer, i4 As Integer
    
    Dim 결산서 As Worksheet
    Dim ws As Worksheet
    Set 결산서 = Worksheets("결산서")
    
    결산서.Select
    결산서.Copy After:=Sheets(9)
    
    Range("b5:k5").Select
    With Selection.Interior
        .ColorIndex = 33
    End With
    
    Set ws = ActiveSheet
ws.Range("i1").Select
ActiveCell.FormulaR1C1 = " 이 쉬트는 삭제하셔도 됩니다."

Dim 항삽입(30)

ws.Range("a6").Select

항 = Left(ActiveCell.Value, 4)

i2 = 1

For i = 1 To 300
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    If Left(ActiveCell.Value, 4) = "" Then
      항삽입(i2) = ActiveCell.Row
      i2 = i2 + 1
         Exit For
    End If

    If 항 <> Left(ActiveCell.Value, 4) Then
      항삽입(i2) = ActiveCell.Row
      i2 = i2 + 1
    End If
   
    항 = Left(ActiveCell.Value, 4)

Next i

Dim 현재줄 As Integer

i4 = 0
항삽입(0) = 6

For i3 = 1 To i2 - 1
    Range("A" & 항삽입(i3) + i4).Select
    Selection.EntireRow.Insert

    현재줄 = ActiveCell.Row
    ws.Range(열이름_항 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "항 소 계"
    Selection.Font.Bold = True
    
    Range(열이름_항 & 현재줄 & ":d" & 현재줄).Select

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

    ws.Range("e" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,R[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("f" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,R[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("j" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 항삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True
    
    ws.Range("k" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
    Selection.Font.Bold = True
    
    Range(열이름_항 & 현재줄 & ":k" & 현재줄).Select
    With Selection.Interior
        .ColorIndex = 35
    End With
    
    If 항삽입(i3 - 1) + i4 + 1 <> 현재줄 Then
        Range(열이름_항 & 항삽입(i3 - 1) + i4 + 1 & ":c" & 현재줄 - 1).Select
        Selection.Font.ColorIndex = 2
        
        Range(열이름_항 & 항삽입(i3 - 1) + i4 & ":c" & 현재줄 - 1).Select

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


'관 정리

Dim 관삽입(30)

ws.Range("a6").Select
Dim 관 As String
관 = Left(ActiveCell.Value, 2)

i2 = 1

For i = 1 To 150
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    
    If Left(ActiveCell.Value, 2) = Empty And 관 = Empty Then
         Exit For
    End If

    If 관 <> Left(ActiveCell.Value, 2) Then
      ActiveCell.Offset(1, 0).Range("A1").Select
      
        If 관 <> Left(ActiveCell.Value, 2) Then
      
            관삽입(i2) = ActiveCell.Row
            i2 = i2 + 1
        End If
      
    End If
    
    관 = Left(ActiveCell.Value, 2)

Next i

i4 = 0
관삽입(0) = 5

For i3 = 1 To i2 - 1
    Range("A" & 관삽입(i3) + i4).Select
    Selection.EntireRow.Insert

    현재줄 = ActiveCell.Row
    ws.Range(열이름_관 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "관 합 계"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":d" & 현재줄).Select

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

    ws.Range("e" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("f" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("j" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 현재줄 - 관삽입(i3 - 1) - i4 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("k" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]-RC[-1])/RC[-6]"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":k" & 현재줄).Select
    With Selection.Interior
        .ColorIndex = 34
    End With
    

    If 관삽입(i3 - 1) + i4 + 1 <> 현재줄 Then
        Range(열이름_관 & 관삽입(i3 - 1) + i4 + 1 & ":b" & 현재줄 - 1).Select
        Selection.Font.ColorIndex = 2

        Range(열이름_관 & 관삽입(i3 - 1) + i4 & ":b" & 현재줄 - 1).Select

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

Dim 처음줄 As Integer
' 총계 정리
    Range("A" & 현재줄 + 1).Select

    처음줄 = 현재줄 - 관삽입(0) + 1

    현재줄 = ActiveCell.Row

    ws.Range(열이름_관 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "총    계"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":d" & 현재줄).Select

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

    ws.Range("f" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("g" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("h" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    ws.Range("i" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=subtotal(9,r[-" & 처음줄 & "]C:R[-1]C)"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":k" & 현재줄).Select
    With Selection.Interior
        .ColorIndex = 33
    End With

'잔액 정리
    Range("A" & 현재줄 + 1).Select

    현재줄 = ActiveCell.Row

    ws.Range(열이름_관 & 현재줄).Select
    ActiveCell.FormulaR1C1 = "실 잔 액"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":d" & 현재줄).Select

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

    ws.Range("g" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    ws.Range("i" & 현재줄).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]-R[-1]C"
    Selection.Font.Bold = True

    Range(열이름_관 & 현재줄 & ":k" & 현재줄).Select
    With Selection.Interior
        .ColorIndex = 33
    End With
    
    ws.Shapes("Button 2").Select
    Selection.Cut
    ws.Shapes("Button 3").Select
    Selection.Cut
    ws.Shapes("Button 1").Select
    Selection.Cut

    ws.PageSetup.PrintArea = "$b$2:$k$" & 현재줄


'메모리 지우기
Application.CutCopyMode = False

End Sub
