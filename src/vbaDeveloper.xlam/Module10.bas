Attribute VB_Name = "Module10"
'module10: 메뉴/인터페이스
Option Explicit
Public Parent As String

Sub 홈()
    Worksheets("첫페이지").Activate
End Sub

Sub 메뉴_도움말()
    Worksheets("도움말").Activate
End Sub

Sub 메뉴_상황도움말()
    '새로 만든 도움말
    UserForm_도움말.Show
End Sub

Sub 메뉴_설정()
    UserForm_front_config.Show
End Sub

Sub 메뉴_환경설정()
    UserForm_설정.Show
End Sub

Sub 메뉴_가져오기()
    UserForm_데이터관리.MultiPage1.Value = 0
    UserForm_데이터관리.Show
End Sub

Sub 메뉴_내보내기()
    UserForm_데이터관리.MultiPage1.Value = 1
    UserForm_데이터관리.Show
End Sub

Sub 메뉴_입출내역입력()
    If ActiveSheet.name = "회계원장" Then
        Parent = "회계원장"
    Else
        Worksheets("회계원장").Activate
    End If
    
    UserForm_입출금입력.Show
    Parent = ""
End Sub

Sub 메뉴_자주쓰는입출금()
    If ActiveSheet.name = "회계원장" Then
        Parent = "회계원장"
    Else
        Worksheets("회계원장").Activate
    End If
    UserForm_자주쓰는입출금.Show
End Sub

Sub 서브메뉴_입출내역수정()
    Dim 줄 As Integer
    If ActiveSheet.name = "회계원장" Then
        Parent = "회계원장"
        줄 = ActiveCell.Row
        If 줄 < 6 Then
            줄 = 6
        End If
        
        UserForm_입출금내역.load_입출금레코드 (줄)
        UserForm_입출금내역.Show
    End If
End Sub

Sub 서브메뉴_지출결의서생성()
    Dim 줄 As Integer, 새줄 As Integer
    Dim dataCount As Integer
    Dim ws As Worksheet, ws_ledger As Worksheet
    Set ws = Worksheets("지출결의대장")
    Set ws_ledger = Worksheets("회계원장")
    Dim error_code As Integer, answer As Integer
    
    If ActiveSheet.name = "회계원장" Then
        Parent = "회계원장"
        error_code = 0
        
        줄 = ActiveCell.Row
        If 줄 < 6 Then
            error_code = 1
        Else
            With ws_ledger.Range("A5")
                dataCount = .End(xlDown).Row - 5
            End With
            If 줄 > dataCount + 5 Then
                error_code = 1
            End If
        End If
        
        If error_code > 0 Then
            MsgBox "지출결의서를 생성할 내용이 들어 있는 줄을 선택해주세요"
            Exit Sub
        End If
        
        With ws_ledger.Range("A" & 줄)
            answer = MsgBox("지출결의대장에 입력하고 지출결의서 양식을 생성합니다 : " & .Value & " / " & .Offset(, 7).Value & " ( " & .Offset(, 9).Value & " ) ", vbYesNo + vbQuestion, "지출결의서 생성 확인")
            If answer <> vbYes Then
                Exit Sub
            End If
        End With
            
        ws.Activate
        
        '지출결의대장에 copy
        Call 회계원장_지출결의입력(줄)
        
        '지출결의대장의 줄번호 가져오기
        If ws.Range("결의날짜레이블").Offset(1, 0).Value Then
            새줄 = ws.Range("결의날짜레이블").End(xlDown).Row

            '지출결의서 생성
            ws.Range("A" & 새줄).Select

            Call 지출결의서작성(False)
        End If
        
    End If
End Sub

Sub 메뉴_프로젝트및부서설정()
    UserForm_프로젝트및부서.Show
End Sub

Sub 메뉴_계정과목입력()
    With Worksheets("예산서")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = True
        .columns(7).Hidden = False
        .Activate
    End With
    
    UserForm_계정과목.Show
End Sub

Sub 메뉴_계정과목보기()
    Worksheets("예산서").Activate
End Sub

Sub 메뉴_예산설정()
    With Worksheets("예산서")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = False
        .columns(7).Hidden = True
        .Activate
    End With
    UserForm_예산.Show
End Sub

Sub 메뉴_예산서보기()
    With Worksheets("예산서")
        .columns(1).Hidden = True
        .columns(5).Hidden = False
        .columns(6).Hidden = False
        .columns(7).Hidden = True
        .Activate
    End With
End Sub

Sub 메뉴_고정비입력()
    Worksheets("2014고정비").Activate
    'UserForm3.Show '미구현이므로 폼 비활성화
End Sub

Sub 메뉴_지출결의서작성()
    If ActiveSheet.name = "지출결의대장" Then
        Parent = "지출결의대장"
    Else
        Worksheets("지출결의대장").Activate
    End If
    On Error GoTo error
    UserForm_지출결의.Show
    Parent = ""
error:
    If Err.Number <> 0 Then

        MsgBox "오류번호 : " & Err.Number & vbCr & _
        "오류내용 : " & Err.Description, vbCritical, "오류"

    End If
End Sub

Sub 메뉴_회계원장조회()
    Worksheets("회계원장").Activate
End Sub

Sub 서브메뉴_회계원장인쇄()
    Worksheets("회계원장").Activate
    UserForm_원장인쇄설정.Show
End Sub

Sub 메뉴_결산서조회()
    If ActiveSheet.name = "회계원장" Then
        Parent = "회계원장"
    Else
        Worksheets("회계원장").Activate
    End If

    UserForm5.Show
    Parent = ""
End Sub

Sub 메뉴_지출결의대장보기()
    Worksheets("지출결의대장").Activate
End Sub

Sub 메뉴_지출결의서생성()
    
    If ActiveSheet.name = "지출결의대장" Then
        Parent = "지출결의대장"
    Else
        Parent = "지출결의대장_from_홈"
    End If
    Worksheets("지출결의대장").Activate

    UserForm_품의서결의서생성.MultiPage1.Value = 0
    UserForm_품의서결의서생성.Show
    Parent = ""
End Sub

Sub 메뉴_초기설정마법사()
    UserForm_초기설정마법사1.Show
End Sub

Sub 메뉴_원장데이터가져오기()
    Dim ws As Worksheet
    Set ws = Worksheets("가져오기")
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub
