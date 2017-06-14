Attribute VB_Name = "Module5"
'module5 : 원장날짜정렬, 원장인쇄, 시트잠금, 시트잠금해제
Option Explicit

Public Const PWD = "1234"
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
Const 최대행수 As Integer = 20000

Sub 원장날짜정렬()
Attribute 원장날짜정렬.VB_Description = ".이(가) 2006-11-14에 기록한 매크로"
Attribute 원장날짜정렬.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim ws_target As Worksheet
    Set ws_target = Worksheets("회계원장")
    ws_target.Unprotect PWD
    
    Dim 끝줄 As Integer
    끝줄 = ws_target.Range("A6").End(xlDown).Row

    ws_target.Range("A8:O" & 끝줄).Sort Key1:=ws_target.Range("A7"), Order1:=xlAscending, Key2:=ws_target.Range("B7") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
        
    ws_target.Protect PWD
    MsgBox "정렬되었습니다"

End Sub


Sub 원장인쇄(시작일 As String, 종료일 As String)

Dim ws As Worksheet
Set ws = Worksheets("회계원장")

ws.Range("a5").Select

Dim 시작줄 As Integer
Dim 끝줄 As Integer
Dim i As Integer

For i = 1 To 최대행수

    ActiveCell.Offset(1, 0).Range("A1").Select

    If 시작일 <= ActiveCell.Value Then
    
       시작줄 = ActiveCell.Row
       Exit For
       
    ElseIf ActiveCell.Value = Empty Then
       Exit For
    End If

Next i

If 시작줄 = Empty Then
    MsgBox "인쇄할 자료가 없습니다."
    Exit Sub
End If

ws.Range("a6").Select
Dim i2 As Integer

For i2 = i To 최대행수

    ActiveCell.Offset(1, 0).Range("A1").Select

    If (종료일 < ActiveCell.Value) Or (ActiveCell.Value = "") Then
      끝줄 = ActiveCell.Row - 1
      Exit For
    End If

Next i2

    ws.PageSetup.PrintArea = "$a$" & 시작줄 & ":$" & 열이름_총잔액 & "$" & 끝줄
    ws.PageSetup.Orientation = xlLandscape
    ActiveWindow.SelectedSheets.PrintPreview

'메모리 지우기
Application.CutCopyMode = False
Set ws = Nothing

End Sub

Sub 시트잠금(ByVal 대상시트 As String)
    If Worksheets("설정").Range("a2").Offset(, 1).Value = True Then
        Worksheets(대상시트).Protect PWD
    End If
End Sub

Sub 시트잠금해제(ByVal 대상시트 As String)
    Worksheets(대상시트).Unprotect PWD
End Sub
