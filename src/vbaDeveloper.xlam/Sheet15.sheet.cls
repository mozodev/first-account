
Private Sub CommandButton2_Click()
    홈
End Sub

Private Sub Worksheet_Activate()
    With Worksheets("품의서")
        .Visible = True
        If Worksheets("설정").Range("결재3설정").Offset(0, 1).Value <> "" Then
            .Pictures("picture 5").Formula = "설정!$e$2:$h$3"
        Else
            .Pictures("picture 5").Formula = "설정!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Worksheets("품의서").Visible = True
End Sub

Private Sub 품의서인쇄_Click()
    
    ActiveSheet.PageSetup.PrintArea = "$b$3:$l$24"
    ActiveWindow.SelectedSheets.PrintPreview

End Sub
