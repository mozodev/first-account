Private Sub Worksheet_Activate()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name <> "ù������" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    Application.StatusBar = ""
End Sub
