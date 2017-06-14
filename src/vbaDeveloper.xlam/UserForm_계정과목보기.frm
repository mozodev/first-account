VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_계정과목보기 
   Caption         =   "계정과목보기"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595.001
   OleObjectBlob   =   "UserForm_계정과목보기.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_계정과목보기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_close_Click()
    Unload Me
End Sub

Sub 계정과목로드(분류 As String, 데이터소스 As String)
    Dim 전체 As Range, 찾은행 As Range, 레코드 As Range
    Dim x As Integer, y As Integer

    Dim 기준열 As Range
    Dim ws_source As Worksheet
    Set ws_source = Worksheets(데이터소스)

    Set 기준열 = ws_source.Range("샘플분류열라벨")
    Dim 관열 As Range, 항열 As Range, 목열 As Range, 세목열 As Range
    
    Set 관열 = ws_source.Range("샘플관열라벨")
    Set 항열 = ws_source.Range("샘플항열라벨")
    Set 목열 = ws_source.Range("샘플목열라벨")
    Set 세목열 = ws_source.Range("샘플세목열라벨")
    
    Dim vlist() As Variant
    
    x = 0

    ' Step 1 : 공통되는 샘플 항목을 먼저 가져옴
    Set 전체 = ws_source.Range("샘플관열라벨").CurrentRegion.columns(기준열.Column)
    Dim 행수 As Integer
       
    행수 = 전체.Rows.Count
    Dim 분류값 As String
    Dim i As Integer
     
    For i = 1 To 행수
        분류값 = ws_source.Range("A" & i).Offset(, 기준열.Column - 1).Value
        If 분류값 = "공통" Or 분류값 = 분류 Then
            ReDim Preserve vlist(5, x)
            Set 레코드 = ws_source.Range("A" & i).Resize(1, 기준열.Column) '열 숫자를 알아내서 바꾸자
            vlist(0, x) = 레코드.Cells(, 관열.Column)
            vlist(1, x) = 레코드.Cells(, 항열.Column)
            vlist(2, x) = 레코드.Cells(, 목열.Column)
            vlist(3, x) = 레코드.Cells(, 세목열.Column)
            vlist(4, x) = 레코드.Cells(, 기준열.Column)
    
            x = x + 1
            ListBox_계정과목보기.Column = vlist
        End If
    Next i
    
    If x = 0 Then
        MsgBox "검색결과가 존재하지 않습니다"
        ListBox_계정과목보기.Clear
    End If

End Sub

Sub 계정과목로드2(유형 As String, 데이터소스 As String)
    ' 기존 계정과목로드 함수를 백업한 것
    Dim 전체 As Range, 찾은행 As Range, 레코드 As Range, 기준열 As Range
    Dim x As Integer, y As Integer
    Dim 첫위치 As String, 필드명 As String
    
    Dim ws_source As Worksheet
    Set ws_source = Worksheets(데이터소스)
    
    Dim 라벨행 As Range
    Set 라벨행 = ws_source.Range("A2").CurrentRegion.Rows(2)
    Set 기준열 = ws_source.Range("분류열라벨")
    
    Dim 관열 As Range, 항열 As Range, 목열 As Range, 세목열 As Range
    Set 관열 = 라벨행.Find(What:="관", LookAt:=xlPart)
    Set 항열 = 라벨행.Find(What:="항", LookAt:=xlPart)
    Set 목열 = 라벨행.Find(What:="목", LookAt:=xlPart)
    Set 세목열 = 라벨행.Find(What:="세목", LookAt:=xlPart)
    
    Dim vlist() As Variant
    
    y = 0
    
    Set 전체 = ws_source.Range("샘플관열라벨").CurrentRegion.columns(기준열.Column)
    Set 찾은행 = 전체.Find(What:=1, LookAt:=xlPart)
    
    If Not 찾은행 Is Nothing Then
        첫위치 = 찾은행.Address
        
        Do
            ReDim Preserve vlist(5, x)
            Set 레코드 = 찾은행.End(xlToLeft).Resize(1, 라벨행.columns.Count)
            vlist(0, x) = 레코드.Cells(, 관열.Column)
            vlist(1, x) = 레코드.Cells(, 항열.Column)
            vlist(2, x) = 레코드.Cells(, 목열.Column)
            vlist(3, x) = 레코드.Cells(, 세목열.Column)
            vlist(4, x) = 레코드.Cells(, 기준열.Column)
    
            x = x + 1
    '        Y = 0
    '
            Set 찾은행 = 전체.FindNext(찾은행)
        Loop While Not 찾은행 Is Nothing And 찾은행.Address <> 첫위치
    
        ListBox_계정과목보기.Column = vlist
    Else
        MsgBox "검색결과가 존재하지 않습니다"
        ListBox_계정과목보기.Clear
    End If

End Sub

Private Sub UserForm_Initialize()
    With ListBox_계정과목보기
        .columnCount = 4
        .ColumnWidths = "1cm;2.7cm;3cm;3.5cm"
    End With
End Sub
