Private Sub CommandButton2_Click()
    Ȩ
End Sub

Private Sub Worksheet_Activate()
    With Worksheets("������Ǽ�")
        .Visible = True
        If Worksheets("����").Range("����3����").Offset(0, 1).Value <> "" Then
            .Pictures("picture 5").Formula = "����!$e$2:$h$3"
        Else
            .Pictures("picture 5").Formula = "����!$e$2:$g$3"
        End If
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Worksheets("������Ǽ�").Visible = True
End Sub

Private Sub ������Ǽ��μ�_Click()
    
    ActiveSheet.PageSetup.PrintArea = "$b$3:$l$24"
    ActiveWindow.SelectedSheets.PrintPreview

End Sub

Private Sub ���Ǽ��μ�_Click()
    ActiveSheet.PageSetup.PrintArea = "$b$3:$l$24"
    ActiveWindow.SelectedSheets.PrintPreview
End Sub