Attribute VB_Name = "Module6"
'module6 : 예산서코드입력, 고정비입력
Option Explicit

Public code_changed As Boolean

Sub 계정과목_시트정렬()

    If Not code_changed Then
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = Worksheets("예산서")
    Dim 시작행 As Integer
    Dim 종료행 As Integer
    Dim r시작위치 As Range
    Set r시작위치 = ws.Range("관필드").Offset(3, 0)
    시작행 = r시작위치.Row
    종료행 = r시작위치.End(xlDown).Row
    ws.Range("A" & 시작행, "G" & 종료행).Select
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B4", "B" & 종료행), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("C4", "C" & 종료행), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("D4", "D" & 종료행), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A" & 시작행, "G" & 종료행)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.CutCopyMode = False
End Sub

Sub 예산서코드입력()
Attribute 예산서코드입력.VB_Description = ".이(가) 2006-11-15에 기록한 매크로"
Attribute 예산서코드입력.VB_ProcData.VB_Invoke_Func = "u\n14"

    Dim 관(100), 항(100), 목(100), 세목(100) As String
    Dim c1, c2, c3, c4 As Integer
    Dim c1_1, c2_2, c3_3, c4_4 As String

    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = Worksheets("예산서")
    ws.Activate
    
With ws
  
    c1 = 0
    c2 = 1
    c3 = 1
    c4 = 1
    
    관(4) = .Range("b4").Value
    항(4) = .Range("c4").Value
    목(4) = .Range("d4").Value
    세목(4) = .Range("d4").Value

    For i = 4 To 100
    
        관(i) = .Range("b" & i).Value
        항(i) = .Range("c" & i).Value
        목(i) = .Range("d" & i).Value
        세목(i) = .Range("e" & i).Value
        
        If 관(i) = "" Then
            Exit For
        End If
        
        If 관(i - 1) <> 관(i) Then
            c1 = c1 + 1
            c2 = 1
            c3 = 1
            c4 = 0
        Else
            If 항(i - 1) <> 항(i) Then
                c2 = c2 + 1
                c3 = 1
                c4 = 0
            Else
                If 목(i - 1) <> 목(i) Then
                    c3 = c3 + 1
                    c4 = 0
                End If
                
            End If
        
        End If
        
        c4 = c4 + 1
 
        If c1 < 10 Then
            c1_1 = "'0" & c1
        Else
            c1_1 = "'" & c1
        End If
        
        If c2 < 10 Then
            c2_2 = "0" & c2
        Else
            c2_2 = c2
        End If
        
        If c3 < 10 Then
            c3_3 = "0" & c3
        Else
            c3_3 = c3
        End If
        
        If c4 < 10 Then
            c4_4 = "0" & c4
        Else
            c4_4 = c4
        End If
         
        .Cells(i, 1) = c1_1 & c2_2 & c3_3 & c4_4
    
    Next i
    
    Dim ii As Integer
    
    For ii = 2 To i - 1
        .Range("o" & ii) = "=IF(RC[-14]>"""",RC[-14]&""/""&RC[-13]&""/""&RC[-12]&""/""&RC[-11]&""/""&RC[-10],"""")"
    Next ii

End With
Erase 관
Erase 항
Erase 목
Erase 세목
    ActiveWorkbook.Names.Add name:="항목선택", RefersToR1C1:="=예산서!R2C15:R" & i - 1 & "C15"

    ActiveWorkbook.Names.Add name:="관항목", RefersToR1C1:="=예산서!R2C1:R" & i - 1 & "C5"
    
End Sub
