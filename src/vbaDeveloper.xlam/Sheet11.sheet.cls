Private Sub Worksheet_Activate()
    With Worksheets("ǰ�Ǽ�����")
        .Visible = True
        .Unprotect
        If (Worksheets("����").Range("a2").Offset(, 1).Value = True) Then
            .Range("ǰ�ǳ�¥���̺�").CurrentRegion.Locked = True
            .Protect
        End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
    If (targetSheet.name = "ù������") Then
        Worksheets("ǰ�Ǽ�����").Visible = False
    End If
End Sub
    