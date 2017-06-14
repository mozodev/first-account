Private Sub Worksheet_Activate()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name <> "첫페이지" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    Application.StatusBar = ""
End Sub
