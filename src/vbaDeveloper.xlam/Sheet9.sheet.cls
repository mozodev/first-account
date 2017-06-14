Private Sub Worksheet_Activate()
    Worksheets("예산서").Visible = True
End Sub

Private Sub Worksheet_Deactivate()
Dim targetSheet As Worksheet
Set targetSheet = ActiveSheet
'On Error GoTo err
    If (targetSheet.name = "첫페이지") Then
        With Worksheets("예산서")
            .columns.Hidden = False  '숨겨진 모든 필드를 되돌린다 -> 계정과목/예산볼때 일부 필드 번갈아 숨김
            .Visible = False
        End With
    End If
'err:
'    errwarn = MsgBox(Error(), vbOKOnly)
'    Resume Next
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub