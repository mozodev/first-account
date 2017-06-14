Attribute VB_Name = "Module1"
'module1 : 일계표, 입금원장, 출금원장
Option Explicit

Const 열이름_날짜 As String = "A"
Const 열이름_관항목 As String = "B"
Const 열이름_code As String = "C"
Const 열이름_관 As String = "D"
Const 열이름_항 As String = "E"
Const 열이름_목 As String = "F"
Const 열이름_세목 As String = "g"
Const 열이름_적요 As String = "h"
Const 열이름_수입 As String = "i"
Const 열이름_지출 As String = "j"
Const 열이름_은현 As String = "k"
Const 열이름_VAT As String = "l"
Const 열이름_대차 As String = "m"
Const 열이름_프로젝트 As String = "n"
Const 열이름_부서 As String = "o"
Const 열이름_현금잔액 As String = "p"
Const 열이름_통장잔액 As String = "q"
Const 열이름_총잔액 As String = "r"

Sub 일계표작성()
Attribute 일계표작성.VB_Description = " .이(가) 2006-10-26에 기록한 매크로"
Attribute 일계표작성.VB_ProcData.VB_Invoke_Func = "q\n14"

    Dim 회계원장 As Worksheet
    Dim 일계표 As Worksheet
    Set 회계원장 = Worksheets("회계원장")
    Set 일계표 = Worksheets("일계표")
    
    회계원장.Select
    Selection.End(xlToLeft).Select
    Dim 날짜 As String
    Dim 날짜2 As String
    Dim 오늘 As String
    Dim 전일 As String
    
    날짜 = ActiveCell.Value
    
    Dim i As Integer
    For i = 1 To 30
    
        ActiveCell.Offset(1, 0).Select
        날짜2 = ActiveCell.Value
        
        If 날짜 <> 날짜2 Then
            ActiveCell.Offset(-1, 0).Select
            Exit For
        End If
    
    Next i
    
    Dim 수입 As Long
    Dim 지출 As Long
    Dim 수입2 As Long
    Dim 지출2 As Long
    Dim 잔고 As Long
    Dim 잔고2 As Long
    Dim 잔고합 As Long
    Dim 전일잔고 As Long
    Dim 전일잔고2 As Long
    Dim 이번수입 As Long
    Dim 이번지출 As Long
    
    수입 = 0
    지출 = 0
    수입2 = 0
    지출2 = 0
    잔고 = 0
    잔고2 = 0
    
    Dim 현재줄 As Integer
    현재줄 = ActiveCell.Row
    
    회계원장.Range("a" & 현재줄).Select
    
    날짜 = ActiveCell.Value
    오늘 = 날짜
    
    잔고 = Range(열이름_통장잔액 & 현재줄).Value
    잔고2 = Range(열이름_현금잔액 & 현재줄).Value
    잔고합 = Range(열이름_총잔액 & 현재줄).Value

    For i = 1 To 300
    
        If 날짜 <> ActiveCell.Value Then
            현재줄 = ActiveCell.Row
            전일잔고 = Range(열이름_통장잔액 & 현재줄).Value
            전일잔고2 = Range(열이름_현금잔액 & 현재줄).Value
            전일 = Range("a" & 현재줄).Value
            Exit For
        End If
    
        현재줄 = ActiveCell.Row
        이번수입 = Range(열이름_수입 & 현재줄).Value
        이번지출 = Range(열이름_지출 & 현재줄).Value
    
        If Left(Range(열이름_code & 현재줄).Value, 2) = "00" Then
        
            수입 = 수입 + 이번수입
            지출 = 지출 + 이번지출
            
            수입2 = 수입2 + 이번수입
            지출2 = 지출2 + 이번지출
        
        Else
        
            If Range(열이름_은현 & 현재줄).Value = 0 Then
                수입 = 수입 + 이번수입
                지출 = 지출 + 이번지출
            End If
        
            If Range(열이름_은현 & 현재줄).Value = 1 Then
                수입2 = 수입2 + 이번수입
                지출2 = 지출2 + 이번지출
            End If
        
        End If

        ActiveCell.Offset(-1, 0).Select
    Next i
    
    일계표.Activate
    
    With 일계표
        '현금
        .Range("b3").Value = 오늘
        .Range("d6").Value = 전일잔고2
        .Range("e6").Value = 전일
        .Range("d7").Value = 수입2
        .Range("d8").Value = 지출2
        .Range("h9").Value = 잔고2
        '은행
        .Range("d11").Value = 전일잔고
        .Range("d12").Value = 수입
        .Range("d13").Value = 지출
        .Range("h14").Value = 잔고
        .Range("h16").Value = 잔고합
        '.Range("A1").Select
    End With
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
