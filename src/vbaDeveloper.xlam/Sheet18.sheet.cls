Private Sub CommandButton1_Click()
    홈
End Sub

Private Sub CommandButton2_Click()
'    ActiveSheet.PageSetup.PrintArea = "$a$1:$e$37"
    ActiveWindow.SelectedSheets.PrintPreview
End Sub

Private Sub Worksheet_Activate()
    With Worksheets("항목분계장")
        .Visible = True
        If Worksheets("설정").Range("결재3설정").Offset(0, 1).Value <> "" Then
            .Pictures("picture 3").Formula = "설정!$e$2:$h$3"
        Else
            .Pictures("picture 3").Formula = "설정!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    If (ActiveSheet.name = "첫페이지") Then
        Worksheets("항목분계장").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub