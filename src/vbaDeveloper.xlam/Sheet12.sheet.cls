Private Sub Worksheet_Activate()
    With Worksheets("������Ǵ���")
        .Visible = True
        .Unprotect
        'If (Worksheets("����").Range("a2").Offset(, 1).Value = True) Then
        '    .Range("���ǳ�¥���̺�").CurrentRegion.Locked = True
        '    .Protect
        'End If
    End With
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    'MsgBox targetsheet.Name
    
    If (targetSheet.name = "ù������") Then
        Worksheets("������Ǵ���").Visible = False
    End If
    If Err.Number <> 0 Then
        MsgBox (Err.Description)
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub