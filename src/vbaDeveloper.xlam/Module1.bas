Attribute VB_Name = "Module1"
'module1 : �ϰ�ǥ, �Աݿ���, ��ݿ���
Option Explicit

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

Sub �ϰ�ǥ�ۼ�()
Attribute �ϰ�ǥ�ۼ�.VB_Description = " .��(��) 2006-10-26�� ����� ��ũ��"
Attribute �ϰ�ǥ�ۼ�.VB_ProcData.VB_Invoke_Func = "q\n14"

    Dim ȸ����� As Worksheet
    Dim �ϰ�ǥ As Worksheet
    Set ȸ����� = Worksheets("ȸ�����")
    Set �ϰ�ǥ = Worksheets("�ϰ�ǥ")
    
    ȸ�����.Select
    Selection.End(xlToLeft).Select
    Dim ��¥ As String
    Dim ��¥2 As String
    Dim ���� As String
    Dim ���� As String
    
    ��¥ = ActiveCell.Value
    
    Dim i As Integer
    For i = 1 To 30
    
        ActiveCell.Offset(1, 0).Select
        ��¥2 = ActiveCell.Value
        
        If ��¥ <> ��¥2 Then
            ActiveCell.Offset(-1, 0).Select
            Exit For
        End If
    
    Next i
    
    Dim ���� As Long
    Dim ���� As Long
    Dim ����2 As Long
    Dim ����2 As Long
    Dim �ܰ� As Long
    Dim �ܰ�2 As Long
    Dim �ܰ��� As Long
    Dim �����ܰ� As Long
    Dim �����ܰ�2 As Long
    Dim �̹����� As Long
    Dim �̹����� As Long
    
    ���� = 0
    ���� = 0
    ����2 = 0
    ����2 = 0
    �ܰ� = 0
    �ܰ�2 = 0
    
    Dim ������ As Integer
    ������ = ActiveCell.Row
    
    ȸ�����.Range("a" & ������).Select
    
    ��¥ = ActiveCell.Value
    ���� = ��¥
    
    �ܰ� = Range(���̸�_�����ܾ� & ������).Value
    �ܰ�2 = Range(���̸�_�����ܾ� & ������).Value
    �ܰ��� = Range(���̸�_���ܾ� & ������).Value

    For i = 1 To 300
    
        If ��¥ <> ActiveCell.Value Then
            ������ = ActiveCell.Row
            �����ܰ� = Range(���̸�_�����ܾ� & ������).Value
            �����ܰ�2 = Range(���̸�_�����ܾ� & ������).Value
            ���� = Range("a" & ������).Value
            Exit For
        End If
    
        ������ = ActiveCell.Row
        �̹����� = Range(���̸�_���� & ������).Value
        �̹����� = Range(���̸�_���� & ������).Value
    
        If Left(Range(���̸�_code & ������).Value, 2) = "00" Then
        
            ���� = ���� + �̹�����
            ���� = ���� + �̹�����
            
            ����2 = ����2 + �̹�����
            ����2 = ����2 + �̹�����
        
        Else
        
            If Range(���̸�_���� & ������).Value = 0 Then
                ���� = ���� + �̹�����
                ���� = ���� + �̹�����
            End If
        
            If Range(���̸�_���� & ������).Value = 1 Then
                ����2 = ����2 + �̹�����
                ����2 = ����2 + �̹�����
            End If
        
        End If

        ActiveCell.Offset(-1, 0).Select
    Next i
    
    �ϰ�ǥ.Activate
    
    With �ϰ�ǥ
        '����
        .Range("b3").Value = ����
        .Range("d6").Value = �����ܰ�2
        .Range("e6").Value = ����
        .Range("d7").Value = ����2
        .Range("d8").Value = ����2
        .Range("h9").Value = �ܰ�2
        '����
        .Range("d11").Value = �����ܰ�
        .Range("d12").Value = ����
        .Range("d13").Value = ����
        .Range("h14").Value = �ܰ�
        .Range("h16").Value = �ܰ���
        '.Range("A1").Select
    End With
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
