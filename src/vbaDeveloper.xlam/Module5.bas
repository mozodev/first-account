Attribute VB_Name = "Module5"
'module5 : ���峯¥����, �����μ�, ��Ʈ���, ��Ʈ�������
Option Explicit

Public Const PWD = "1234"
Const ���̸�_��¥ As String = "A"
Const ���̸�_���׸� As String = "B"
Const ���̸�_code As String = "C"
Const ���̸�_�� As String = "D"
Const ���̸�_�� As String = "E"
Const ���̸�_�� As String = "F"
Const ���̸�_���� As String = "g"
Const ���̸�_���� As String = "h"
Const ���̸�_���� As String = "i"
Const ���̸�_���� As String = "j"
Const ���̸�_���� As String = "k"
Const ���̸�_VAT As String = "l"
Const ���̸�_���� As String = "m"
Const ���̸�_������Ʈ As String = "n"
Const ���̸�_�μ� As String = "o"
Const ���̸�_�����ܾ� As String = "p"
Const ���̸�_�����ܾ� As String = "q"
Const ���̸�_���ܾ� As String = "r"
Const �ִ���� As Integer = 20000

Sub ���峯¥����()
Attribute ���峯¥����.VB_Description = ".��(��) 2006-11-14�� ����� ��ũ��"
Attribute ���峯¥����.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim ws_target As Worksheet
    Set ws_target = Worksheets("ȸ�����")
    ws_target.Unprotect PWD
    
    Dim ���� As Integer
    ���� = ws_target.Range("A6").End(xlDown).Row

    ws_target.Range("A8:O" & ����).Sort Key1:=ws_target.Range("A7"), Order1:=xlAscending, Key2:=ws_target.Range("B7") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
        
    ws_target.Protect PWD
    MsgBox "���ĵǾ����ϴ�"

End Sub


Sub �����μ�(������ As String, ������ As String)

Dim ws As Worksheet
Set ws = Worksheets("ȸ�����")

ws.Range("a5").Select

Dim ������ As Integer
Dim ���� As Integer
Dim i As Integer

For i = 1 To �ִ����

    ActiveCell.Offset(1, 0).Range("A1").Select

    If ������ <= ActiveCell.Value Then
    
       ������ = ActiveCell.Row
       Exit For
       
    ElseIf ActiveCell.Value = Empty Then
       Exit For
    End If

Next i

If ������ = Empty Then
    MsgBox "�μ��� �ڷᰡ �����ϴ�."
    Exit Sub
End If

ws.Range("a6").Select
Dim i2 As Integer

For i2 = i To �ִ����

    ActiveCell.Offset(1, 0).Range("A1").Select

    If (������ < ActiveCell.Value) Or (ActiveCell.Value = "") Then
      ���� = ActiveCell.Row - 1
      Exit For
    End If

Next i2

    ws.PageSetup.PrintArea = "$a$" & ������ & ":$" & ���̸�_���ܾ� & "$" & ����
    ws.PageSetup.Orientation = xlLandscape
    ActiveWindow.SelectedSheets.PrintPreview

'�޸� �����
Application.CutCopyMode = False
Set ws = Nothing

End Sub

Sub ��Ʈ���(ByVal ����Ʈ As String)
    If Worksheets("����").Range("a2").Offset(, 1).Value = True Then
        Worksheets(����Ʈ).Protect PWD
    End If
End Sub

Sub ��Ʈ�������(ByVal ����Ʈ As String)
    Worksheets(����Ʈ).Unprotect PWD
End Sub
