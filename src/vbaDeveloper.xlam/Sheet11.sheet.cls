Private Sub Worksheet_Activate()
    With Worksheets("품의서대장")
        .Visible = True
        .Unprotect
        If (Worksheets("설정").Range("a2").Offset(, 1).Value = True) Then
            .Range("품의날짜레이블").CurrentRegion.Locked = True
            .Protect
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
    If (targetSheet.name = "첫페이지") Then
        Worksheets("품의서대장").Visible = False
    End If
End Sub
    