'Dim started As Integer

Private Sub Worksheet_Activate()
    'If Not started > 0 Then
    '    started = 0
    'End If
    
    'If started = 0 Then
    '    UserForm_����.Hide
    '
    '    ' #2 ������ �����϶� �ȳ� �޽���
    '    Range("A3:P1000").ClearContents
    '    MsgBox "������ �����͸� �� ��Ʈ (����° ��)�� �������ּ���"
    '    Range("A3").Select
    'End If
    
    ' #3 ��Ʈ�� �����Ͱ� �����, ����, ���׸�, ����, ����/����, ��/�� ���� �� ù �ٿ� ������� �޽���
    ' #4 ù �࿡ ��� �� �ִ��� Ȯ��
    ' #5 �� �� ��ȸ�ϸ� ���� & ȸ����忡 ���̱�
    ' #6 �߰��� �κ� ��ȸ�ϸ� ���׸� ���� : ������ ������ ǥ��, ������ ������ ��ϵ� �ڵ����� Ȯ��
    ' #6-1. �ű� ���׸��� ��� ������ ���׸� ����
    ' #7 ȸ����� ����

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' #3 ��Ʈ�� �����Ͱ� �����, ����, ���׸�, ����, ����/����, ��/�� ���� �� ù �ٿ� ������� �޽���
    'With Worksheets("��������")
    '    If .Range("A3").Value <> "" And started = 0 Then
    '        MsgBox "�ι�° �࿡ ���� �߰����ּ���(����, ��, ��, ��, ����, ����, ����, ��/��)"
    '        .Range("A2").Select
    '        started = 1
    '    End If
    'End With
End Sub