Private Sub Worksheet_Activate()
    Worksheets("����").Visible = True
End Sub

Private Sub Worksheet_Deactivate()
    If (ActiveSheet.name = "ù������") Then
        Worksheets("����").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub