Private Sub Worksheet_Activate()
    Worksheets("���꼭").Visible = True
End Sub

Private Sub Worksheet_Deactivate()
Dim targetSheet As Worksheet
Set targetSheet = ActiveSheet
'On Error GoTo err
    If (targetSheet.name = "ù������") Then
        With Worksheets("���꼭")
            .columns.Hidden = False  '������ ��� �ʵ带 �ǵ����� -> ��������/���꺼�� �Ϻ� �ʵ� ������ ����
            .Visible = False
        End With
    End If
'err:
'    errwarn = MsgBox(Error(), vbOKOnly)
'    Resume Next
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub