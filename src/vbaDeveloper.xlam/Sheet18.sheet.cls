Private Sub CommandButton1_Click()
    Ȩ
End Sub

Private Sub CommandButton2_Click()
'    ActiveSheet.PageSetup.PrintArea = "$a$1:$e$37"
    ActiveWindow.SelectedSheets.PrintPreview
End Sub

Private Sub Worksheet_Activate()
    With Worksheets("�׸�а���")
        .Visible = True
        If Worksheets("����").Range("����3����").Offset(0, 1).Value <> "" Then
            .Pictures("picture 3").Formula = "����!$e$2:$h$3"
        Else
            .Pictures("picture 3").Formula = "����!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    If (ActiveSheet.name = "ù������") Then
        Worksheets("�׸�а���").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub