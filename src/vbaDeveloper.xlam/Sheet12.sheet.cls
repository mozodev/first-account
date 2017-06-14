Private Sub Worksheet_Activate()
    With Worksheets("지출결의대장")
        .Visible = True
        .Unprotect
        'If (Worksheets("설정").Range("a2").Offset(, 1).Value = True) Then
        '    .Range("결의날짜레이블").CurrentRegion.Locked = True
        '    .Protect
        'End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    'MsgBox targetsheet.Name
    
    If (targetSheet.name = "첫페이지") Then
        Worksheets("지출결의대장").Visible = False
    End If
    If Err.Number <> 0 Then
        MsgBox (Err.Description)
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub