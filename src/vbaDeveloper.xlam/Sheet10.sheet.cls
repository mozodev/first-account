Option Explicit
Const PWD = "1234"

Private Sub Worksheet_Activate()
    With Worksheets("ȸ�����")
        .Visible = True
        .Unprotect PWD
        If (Worksheets("����").Range("��Ʈ��ݼ���").Offset(, 1).Value = True) Then
            .Range("�����ʵ巹�̺�").CurrentRegion.Locked = True
            .Protect PWD
        End If
        
        If Worksheets("����").Range("����3����").Offset(0, 1).Value <> "" Then
            .Pictures("picture 15").Formula = "����!$e$2:$h$3"
        Else
            .Pictures("picture 15").Formula = "����!$e$2:$g$3"
        End If
        
    End With
End Sub

Private Sub Worksheet_Deactivate()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
    If (targetSheet.name = "ù������") Then
        Worksheets("ȸ�����").Visible = False
    End If
End Sub