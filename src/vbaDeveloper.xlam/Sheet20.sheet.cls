
Private Sub Worksheet_Activate()
    With Worksheets("��꼭")
        .Visible = True
        If Worksheets("����").Range("����3����").Offset(0, 1).Value <> "" Then
            .Pictures("picture 6").Formula = "����!$e$2:$h$3"
        Else
            .Pictures("picture 6").Formula = "����!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim sh As Worksheet
    Set sh = ActiveSheet
    If (sh.name = "ù������") Then
        Worksheets("��꼭").Visible = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Private Sub ��꼭�μ�_Click()
    ActiveWindow.SelectedSheets.PrintPreview
End Sub