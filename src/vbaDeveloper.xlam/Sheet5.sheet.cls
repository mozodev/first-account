Private Sub Worksheet_Activate()
    Worksheets("도움말").Visible = True
End Sub

Private Sub Worksheet_Deactivate()
    If (ActiveSheet.name = "첫페이지") Then
        Worksheets("도움말").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub