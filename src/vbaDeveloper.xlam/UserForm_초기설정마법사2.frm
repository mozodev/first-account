VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_초기설정마법사2 
   Caption         =   "초기설정마법사_2단계"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865.001
   OleObjectBlob   =   "UserForm_초기설정마법사2.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_초기설정마법사2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    '설정값 저장
    Dim 기관명, 회계시작일, 담당자직함, 결재1직함, 결재2직함 As String
    Dim 시트잠금 As String
    
    기관명 = TextBox_기관명.Value
    회계시작일 = TextBox_회계시작일.Value
    담당자직함 = TextBox_담당자직함.Value
    결재1직함 = TextBox_결재1직함.Value
    결재2직함 = TextBox_결재2직함.Value

    시트잠금 = True
    
    With Worksheets("설정")
        If 담당자직함 <> "" Then
            .Range("E2").Value = 담당자직함
        End If
        If 결재1직함 <> "" Then
            .Range("F2").Value = 결재1직함
        End If
        If 결재2직함 <> "" Then
            .Range("G2").Value = 결재2직함
        End If
    End With
    
    With Worksheets("설정")
        .Range("시트잠금설정").Offset(0, 1).Value = 시트잠금
        If 기관명 <> "" And 기관명 <> .Range("기관명설정").Offset(0, 1).Value Then
            .Range("기관명설정").Offset(0, 1).Value = 기관명
        End If
        
        If 담당자직함 <> "" Then
            .Range("담당자직함설정").Offset(0, 1).Value = 담당자직함
        End If
        
        If 회계시작일 <> "" And 회계시작일 <> .Range("회계시작일설정").Offset(0, 1).Value Then
            .Range("회계시작일설정").Offset(0, 1).Value = 회계시작일
        End If
        
        If 결재1직함 <> "" Then
            .Range("결재1설정").Offset(0, 1).Value = 결재1직함
        End If
        
        If 결재2직함 <> "" Then
            .Range("결재2설정").Offset(0, 1).Value = 결재2직함
        End If

    End With
    
    With Worksheets("회계원장")
        .Unprotect PWD
        If 기관명 <> "" And 기관명 <> .Range("기관명").Value Then
            .Range("기관명").Value = 기관명
        End If

        If Worksheets("설정").Range("시트잠금설정").Offset(, 1).Value = True Then
            .Protect PWD
        End If
    End With
    
    Unload Me
    UserForm_초기설정마법사3.Show
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    UserForm_초기설정마법사1.Show
End Sub

