VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_설정 
   Caption         =   "환경설정"
   ClientHeight    =   6100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   OleObjectBlob   =   "UserForm_설정.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_설정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const 최대행수 As Integer = 10000
Const 최대행수2 As Integer = 5000
'Const PWD = "1234"

'사용자가 직접 시트를 만지기보다 준비된 폼 들을 이용하도록 막음
Sub 시트보호(시트이름 As String)
    With Worksheets(시트이름)
        .Protect 'PWD '시트보호
    End With
   
    MsgBox "선택영역의 잠금 작업이 완료되었습니다.", vbExclamation, ""
End Sub

Private Sub CheckBox_시트잠금_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "엑셀 사용에 익숙하거나 기능을 변경하고픈 경우에만 해제하세요"
End Sub

'간편 업데이트 적용 기능 구현 중.
'VBA 코드만 업데이트 되고 시트에는 변동이 없는 경우 엑셀 파일과 같은 위치에 업데이트 파일을 두고 이 버튼 클릭하게
Private Sub CommandButton_update_Click()
    Dim Filt$, title$, fileName$, Message As VbMsgBoxResult
    Filt = "VB Files (*.bas; *.frm; *.cls)(*.bas; *.frm; *.cls)," & _
        "*.bas;*.frm;*.cls"
    Dim vbp As Object 'VBIDE.VBProject
    Dim vbc As Object 'VBIDE.VBComponent
    
    Set vbp = ActiveWorkbook.VBProject
    
    fileName = Application.GetOpenFilename(FileFilter:=Filt, _
        FilterIndex:=5, title:=title)
        
    If fileName <> vbNullString Then
        Dim n As Integer
        n = Len(fileName)
        Dim m As String
        m = Left(fileName, n - 4)
        Dim m_array() As String
        m_array() = Split(m, "\")
        m = m_array(UBound(m_array))
            
        On Error Resume Next
        Set vbc = vbp.VBComponents(m)
        If Err = 0 Then
            vbp.VBComponents.Remove vbc
        Else
            Err.Clear
        End If
        vbp.VBComponents.Import fileName
        MsgBox "업데이트 했습니다"
        On Error GoTo 0
    End If

End Sub

Function module_name(fileName As String)
    Dim n As Integer
    n = Len(fileName)
    Dim m As String
    m = Left(fileName, n - 4)
    Dim m_array() As String
    m_array() = Split(m, "\")
    module_name = m_array(UBound(m_array))
End Function

Private Sub CommandButton_updateAll_Click()

    Dim vbp As Object 'VBIDE.VBProject
    Dim vbc As Object 'VBIDE.VBComponent
    
    Set vbp = ActiveWorkbook.VBProject
    
    Dim sThisFilePath          As String
    Dim sFile                  As String
    
    sThisFilePath = ThisWorkbook.Path
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    
    sFile = Dir(sThisFilePath & "*.bas")
    Dim m As String
    
    Do While sFile <> vbNullString
        'MsgBox "The next file is " & sFile
        On Error Resume Next
        m = module_name(sFile)
        Set vbc = vbp.VBComponents(m)
        If Err = 0 Then
            vbp.VBComponents.Remove vbc
        Else
            Err.Clear
        End If
        vbp.VBComponents.Import sFile
        
        On Error GoTo 0
        sFile = Dir
    Loop
    MsgBox "업데이트 했습니다"
End Sub

Private Sub CommandButton_계정과목가져오기_Click()
    ' #1 빈 시트 염
    Dim ws As Worksheet
    Set ws = Worksheets("가져오기2")
    ws.Visible = xlSheetVisible
    ws.Activate
    Unload Me

    ' #2 데이터 복사하란 안내 메시지
   
    ' #3 시트에 데이터가 생기면, 일자, 관항목, 적요, 수입/지출, 은/현 구분 라벨 첫 줄에 넣으라는 메시지
    ' #4 첫 행에 모든 라벨 있는지 확인
    ' #5 각 열 순회하며 복사 & 회계원장에 붙이기
    ' #6 추가된 부분 순회하며 관항목 검증 : 없으면 붉은색 표시, 있으면 기존에 등록된 코드인지 확인
    ' #6-1. 신규 관항목이 들어 있으면 관항목 생성
    ' #7 회계원장 정렬
End Sub

Private Sub CommandButton_계정과목초기화_Click()
    Call 회계설정초기화("계정과목")
End Sub

Private Sub CommandButton_부서초기화_Click()
    Call 회계설정초기화("부서")
End Sub

Private Sub CommandButton_저장_Click()
    Dim orgName As String
    Dim accountStartDate As String
    Dim startDate As String
    Dim endDate As String
    Dim 담당자직함 As String
    Dim 결재1직함 As String
    Dim 결재2직함 As String
    Dim 결재3직함 As String
    Dim 시트잠금 As String
    
    orgName = TextBox_기관명.Value
    accountStartDate = TextBox_회계시작일.Value
    담당자직함 = TextBox_담당자직함.Value
    결재1직함 = TextBox_결재1직함.Value
    결재2직함 = TextBox_결재2직함.Value
    결재3직함 = TextBox_결재3직함.Value

    시트잠금 = CheckBox_시트잠금.Value
    
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
        If 결재3직함 <> "" Then
            .Range("H2").Value = 결재3직함
            '.Pictures("Picture 2").Formula = "$L$26:$O$27"
        Else
            '.Pictures("Picture 2").Formula = "$L$26:$N$27"
        End If
    End With
    
    With Worksheets("설정")
        .Range("시트잠금설정").Offset(0, 1).Value = 시트잠금
        If orgName <> "" And orgName <> .Range("기관명설정").Offset(0, 1).Value Then
            .Range("기관명설정").Offset(0, 1).Value = orgName
        End If
        
        If 담당자직함 <> "" Then
            .Range("담당자직함설정").Offset(0, 1).Value = 담당자직함
        End If
        
        If accountStartDate <> "" And accountStartDate <> .Range("회계시작일설정").Offset(0, 1).Value Then
            .Range("회계시작일설정").Offset(0, 1).Value = accountStartDate
        End If
        
        .Range("작업시작일설정").Offset(0, 1).Value = startDate
        .Range("작업종료일설정").Offset(0, 1).Value = endDate
        
        If 결재1직함 <> "" Then
            .Range("결재1설정").Offset(0, 1).Value = 결재1직함
        End If
        
        If 결재2직함 <> "" Then
            .Range("결재2설정").Offset(0, 1).Value = 결재2직함
        End If
        
        'If 결재3직함 <> "" Then
            .Range("결재3설정").Offset(0, 1).Value = 결재3직함
        'End If

    End With
    
    With Worksheets("회계원장")
        .Unprotect PWD
        If orgName <> "" And orgName <> .Range("기관명").Value Then
            .Range("기관명").Value = orgName
        End If

        If Worksheets("설정").Range("시트잠금설정").Offset(, 1).Value = True Then
            .Protect PWD
        End If
    End With
    
    MsgBox "설정되었습니다"
    Unload Me
    홈
End Sub

Private Sub CommandButton_지출결의초기화_Click()
    Call 회계자료초기화("지출결의")
End Sub

Private Sub CommandButton_취소_Click()
    Unload Me
    홈
End Sub

Private Sub CommandButton_프로젝트초기화_Click()
    Call 회계설정초기화("프로젝트")
End Sub

Private Sub CommandButton_회계설정전체초기화_Click()
    Call 회계설정초기화("전체")
End Sub

Private Sub CommandButton_회계원장가져오기_Click()
    ' #1 빈 시트 염
    Dim ws As Worksheet
    Set ws = Worksheets("가져오기")
    ws.Visible = xlSheetVisible
    ws.Activate
    Unload Me

    ' #2 데이터 복사하란 안내 메시지
   
    ' #3 시트에 데이터가 생기면, 일자, 관항목, 적요, 수입/지출, 은/현 구분 라벨 첫 줄에 넣으라는 메시지
    ' #4 첫 행에 모든 라벨 있는지 확인
    ' #5 각 열 순회하며 복사 & 회계원장에 붙이기
    ' #6 추가된 부분 순회하며 관항목 검증 : 없으면 붉은색 표시, 있으면 기존에 등록된 코드인지 확인
    ' #6-1. 신규 관항목이 들어 있으면 관항목 생성
    ' #7 회계원장 정렬
End Sub

Private Sub CommandButton_회계원장초기화_Click()
    Call 회계자료초기화("회계원장")
End Sub

Sub 회계자료초기화(종류 As String)
    Dim 코드 As String
    
    If 종류 = "회계원장" Or 종류 = "전체" Then
        Call 시트잠금해제("회계원장")
        Dim 이월금 As String
        이월금 = InputBox("이월금이 있으면 금액을 적어주세요. 없으면 0 혹은 그냥 확인을 눌러주세요")
        With Worksheets("회계원장")
            .Range("A6:O7").ClearContents
            .Range("A8:O" & 최대행수).ClearContents
            .Range("a6").Value = Worksheets("설정").Range("회계시작일설정").Offset(0, 1).Value
            .Range("a7").Value = .Range("a6").Value
            .Range("c6").NumberFormat = "@"
            .Range("c7").NumberFormat = "@"
            .Range("c6").Value = get_code("수입", "이월금", "이월금", "전기이월")
            .Range("b6").Value = .Range("c6").Value & "/수입/이월금/이월금/전기이월"
            .Range("c7").Value = "00010101"
            .Range("b7").Value = "00010101/예산외수입/"
            .Range("d6").Value = "수입"
            .Range("e6").Value = "이월금"
            .Range("f6").Value = "이월금"
            .Range("g6").Value = "전기이월"
            .Range("d7").Value = "예산외수입"
            .Range("H6").Value = "전년이월"
            .Range("i6").Value = 이월금
            .Range("i7").Value = 이월금
            .Range("H7").Value = "통장입금"
            .Range("K6:O7").ClearContents
            .Range("K6").Value = 1
            .Range("K7").Value = 0
        End With
        Call 시트잠금("회계원장")
        MsgBox "회계원장이 초기화되었습니다"
    End If
    
    If 종류 = "지출결의" Or 종류 = "전체" Then
        With Worksheets("지출결의대장")
            .Range("A4:I" & 최대행수2).ClearContents
        End With
        MsgBox "지출결의대장이 초기화되었습니다"
    End If

End Sub

Sub 회계설정초기화(종류 As String)
    If 종류 = "프로젝트" Or 종류 = "전체" Then
        With Worksheets("설정").Range("프로젝트설정레이블")
            Range(.Offset(1, 0), .End(xlDown)).ClearContents
        End With
        MsgBox "프로젝트 설정이 초기화되었습니다"
    End If
    
    If 종류 = "부서" Or 종류 = "전체" Then
        With Worksheets("설정").Range("부서설정레이블")
            Range(.Offset(1, 0), .End(xlDown)).ClearContents
        End With
        MsgBox "부서 설정이 초기화되었습니다"
    End If
    
    If 종류 = "계정과목" Or 종류 = "전체" Then
        With Worksheets("예산서")
            .Range("A4:G1000").ClearContents
            With .Range("b4")
                .Value = "수입"
                .Offset(, 1).Value = "이월금"
                .Offset(, 2).Value = "이월금"
                .Offset(, 3).Value = "전기이월"
            End With
            
            With .Range("b5")
                .Value = "지출"
                .Offset(, 1).Value = "예비비"
                .Offset(, 2).Value = "예비비"
                .Offset(, 3).Value = "예비비"
            End With
            
        End With
        Call 예산서코드입력
        MsgBox "계정과목 설정이 초기화되었습니다"
    End If
        
    If 종류 = "예산" Or 종류 = "전체" Then
        With Worksheets("예산서")
            .Range("F4", .Range("F4").End(xlDown)).ClearContents
        End With
        MsgBox "예산 설정이 초기화되었습니다"
    End If
    
End Sub

Sub 기본설정초기화()
    With Worksheets("설정")
        .Range("E2").Value = "" '결재1
        .Range("F2").Value = "" '결재2
        .Range("G2").Value = "" '결재3
        .Range("H2").Value = "" '결재4
        
        .Range("시트잠금설정").Offset(0, 1).Value = ""
        .Range("기관명설정").Offset(0, 1).Value = ""
        .Range("담당자직함설정").Offset(0, 1).Value = ""
        .Range("회계시작일설정").Offset(0, 1).Value = ""
    End With
End Sub

Private Sub CommandButton_회계자료전체초기화_Click()
    '전체 초기화를 한 경우에는 회계시작일 날짜도 올해 첫해로 수정
    Call init_firstday
    
    Call 회계자료초기화("전체")
End Sub


Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label_메시지.caption = "단체/조직명을 입력해주세요. 각종 대장 및 양식 출력시 표시됩니다"
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If MultiPage1.Value = 0 Then
        Label_메시지.caption = "필수적으로 설정해야 할 내용들입니다"
    ElseIf MultiPage1.Value = 1 Then
        Label_메시지.caption = "부가적으로 설정하는 내용들입니다"
    Else
        Label_메시지.caption = "입력한 회계자료를 모두 비우고 처음으로 되돌립니다"
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim accountStartDate As String
    'Dim mytoday As String
    'Dim days As Integer
    
    With Worksheets("설정")
        TextBox_기관명.Value = .Range("기관명설정").Offset(0, 1).Value
        If Not .Range("회계시작일설정").Offset(0, 1).Value = "" Then
            'mytoday = Date
            accountStartDate = .Range("회계시작일설정").Offset(0, 1).Value
            'days = DateDiff("d", mytoday, 회계시작일)
            'If days > 365 Then
            '    회계시작일 = Year(Now) & "-01-01"
            'End If
            TextBox_회계시작일.Value = accountStartDate
        Else
            'accountStartDate = Year(Now) & "-01-01"
            Call init_firstday
        End If
        
        TextBox_담당자직함.Value = .Range("담당자직함설정").Offset(0, 1).Value
        TextBox_결재1직함.Value = .Range("결재1설정").Offset(0, 1).Value
        TextBox_결재2직함.Value = .Range("결재2설정").Offset(0, 1).Value
        TextBox_결재3직함.Value = .Range("결재3설정").Offset(0, 1).Value

        CheckBox_시트잠금.Value = Worksheets("설정").Range("시트잠금설정").Offset(0, 1).Value
    End With
    MultiPage1.Value = 0  '첫페이지(기본설정)가 항상 먼저 뜨도록
End Sub
