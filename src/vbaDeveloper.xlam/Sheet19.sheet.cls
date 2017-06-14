
Private Sub Worksheet_Activate()
    With Worksheets("결산서")
        .Visible = True
        If Worksheets("설정").Range("결재3설정").Offset(0, 1).Value <> "" Then
            .Pictures("picture 6").Formula = "설정!$e$2:$h$3"
        Else
            .Pictures("picture 6").Formula = "설정!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim sh As Worksheet
    Set sh = ActiveSheet
    If (sh.name = "첫페이지") Then
        Worksheets("결산서").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Private Sub 결산서인쇄_Click()
    ActiveWindow.SelectedSheets.PrintPreview
End Sub