Private Sub Worksheet_Activate()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name <> "ù������" Then
            ws.Visible = False
        End If
    Next ws
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub