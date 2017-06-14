Option Explicit
Const PWD = "1234"

Private Sub Worksheet_Activate()
    With Worksheets("회계원장")
        .Visible = True
        .Unprotect PWD
        If (Worksheets("설정").Range("시트잠금설정").Offset(, 1).Value = True) Then
            .Range("일자필드레이블").CurrentRegion.Locked = True
            .Protect PWD
        End If
        
        If Worksheets("설정").Range("결재3설정").Offset(0, 1).Value <> "" Then
            .Pictures("picture 15").Formula = "설정!$e$2:$h$3"
        Else
            .Pictures("picture 15").Formula = "설정!$e$2:$g$3"
        End If
        
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
    If (targetSheet.name = "첫페이지") Then
        Worksheets("회계원장").Visible = False
    End If
End Sub
