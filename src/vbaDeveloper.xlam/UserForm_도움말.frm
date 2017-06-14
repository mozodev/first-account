VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_도움말 
   Caption         =   "도움말"
   ClientHeight    =   6015
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   9765.001
   OleObjectBlob   =   "UserForm_도움말.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_도움말"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 도움말검색결과() As String
Dim 검색결과인덱스 As Integer
Dim 검색결과수 As Integer
Dim ws As Worksheet

Private Sub CommandButton_검색_Click()
    Set ws = Worksheets("기능도움말")
    Dim c As Range
    Dim 검색어 As String
    Dim i As Integer
    Dim 찾은코드 As String
    Dim 도움말수 As Integer
    도움말수 = ws.Range("기능코드레이블").End(xlDown).Row

    ReDim 도움말검색결과(2, 도움말수)
        
    i = 1
    검색어 = TextBox_도움말검색.Value
    검색결과인덱스 = 1
    
    With ws.Cells
        Set c = .Find(What:=검색어)
        If Not c Is Nothing Then
            Dim firstaddress As String
            firstaddress = c.Address
            Do
                찾은코드 = c.End(xlToLeft).Value
                If i = 1 Then
                    도움말검색결과(1, i) = 찾은코드
                    도움말검색결과(2, i) = c.Row
                    i = i + 1
                ElseIf i > 1 And 도움말검색결과(2, i - 1) <> c.Row Then
                    도움말검색결과(1, i) = 찾은코드
                    도움말검색결과(2, i) = c.Row
                    i = i + 1
                End If
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstaddress
        End If
    End With
    
    검색결과수 = i - 1

    Label_검색결과수.caption = "검색결과 : " & 검색결과수 & "건"
    If 검색결과수 > 0 Then
        
        '검색한 결과 중 첫번째 결과 표시
        Dim 행 As Integer
        Dim j As Integer
        행 = 도움말검색결과(2, 1)

        With ws.Range("A" & 행)
            대분류 = .Offset(0, 1).Value
            분류 = .Offset(0, 3).Value

            Select Case 대분류
                Case "일상회계"
                    MultiPage1.Value = 0
                    For j = 1 To ListBox_일상회계.ListCount

                        If ListBox_일상회계.List(j - 1, 1) = .Value Then
                            ListBox_일상회계.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "지출결의"
                    MultiPage1.Value = 1
                    For j = 1 To ListBox_지출결의.ListCount

                        If ListBox_지출결의.List(j - 1, 1) = .Value Then
                            ListBox_지출결의.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "지출품의"
                    MultiPage1.Value = 2
                    For j = 1 To ListBox_지출품의.ListCount

                        If ListBox_지출품의.List(j - 1, 1) = .Value Then
                            ListBox_지출품의.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "설정"
                    MultiPage1.Value = 3
                    For j = 1 To ListBox_설정.ListCount

                        If ListBox_설정.List(j - 1, 1) = .Value Then
                            ListBox_설정.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "예산"
                    MultiPage1.Value = 4
                    For j = 1 To ListBox_예산.ListCount

                        If ListBox_예산.List(j - 1, 1) = .Value Then
                            ListBox_예산.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "결산"
                    MultiPage1.Value = 5
                    For j = 1 To ListBox_결산.ListCount

                        If ListBox_결산.List(j - 1, 1) = .Value Then
                            ListBox_결산.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
                Case "자산채무"
                    MultiPage1.Value = 6
                    For j = 1 To ListBox_자산채무.ListCount

                        If ListBox_자산채무.List(j - 1, 1) = .Value Then
                            ListBox_자산채무.Selected(j - 1) = True
                            Exit For
                        End If
                    Next
            End Select
        End With

        CommandButton_다음찾기.Visible = True
    Else
        CommandButton_다음찾기.Visible = False
        TextBox_도움말검색.SetFocus
    End If
End Sub

Private Sub CommandButton_다음찾기_Click()
    검색결과인덱스 = 검색결과인덱스 + 1

    If 검색결과인덱스 <= 검색결과수 Then
        Call 검색결과표시(검색결과인덱스)
    Else
        MsgBox "더 이상의 검색결과는 없습니다"
        TextBox_도움말검색.SetFocus
    End If
End Sub

Sub 검색결과표시(인덱스 As Integer)
    Dim 행 As Integer
    Dim j As Integer
    행 = 도움말검색결과(2, 인덱스)
    Set ws = Worksheets("기능도움말")

    With ws.Range("A" & 행)
        대분류 = .Offset(0, 1).Value
        분류 = .Offset(0, 3).Value

        Select Case 대분류
            Case "일상회계"
                MultiPage1.Value = 0
                For j = 0 To ListBox_일상회계.ListCount - 1

                    If ListBox_일상회계.List(j, 1) = .Value Then
                        ListBox_일상회계.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "지출결의"
                MultiPage1.Value = 1
                For j = 0 To ListBox_지출결의.ListCount - 1

                    If ListBox_지출결의.List(j, 1) = .Value Then
                        ListBox_지출결의.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "지출품의"
                MultiPage1.Value = 2
                For j = 0 To ListBox_지출품의.ListCount - 1

                    If ListBox_지출품의.List(j, 1) = .Value Then
                        ListBox_지출품의.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "설정"
                MultiPage1.Value = 3
                For j = 0 To ListBox_설정.ListCount - 1

                    If ListBox_설정.List(j, 1) = .Value Then
                        ListBox_설정.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "예산"
                MultiPage1.Value = 4
                For j = 0 To ListBox_예산.ListCount - 1

                    If ListBox_예산.List(j, 1) = .Value Then
                        ListBox_예산.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "결산"
                MultiPage1.Value = 5
                For j = 0 To ListBox_결산.ListCount - 1

                    If ListBox_결산.List(j, 1) = .Value Then
                        ListBox_결산.Selected(j) = True
                        Exit For
                    End If
                Next
            Case "자산채무"
                MultiPage1.Value = 6
                For j = 0 To ListBox_자산채무.ListCount - 1

                    If ListBox_자산채무.List(j, 1) = .Value Then
                        ListBox_자산채무.Selected(j) = True
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub ListBox_결산_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_결산
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_결산도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_설정_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_설정
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_설정도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_예산_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_예산
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_예산도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_일상회계_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_일상회계
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_일상회계도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_자산채무_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_자산채무
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_자산채무도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_지출결의_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_지출결의
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_지출결의도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub ListBox_지출품의_Click()
    Dim i선택 As Integer
    Dim 도움말 As String
    
    With ListBox_지출품의
        i선택 = .ListIndex
        If i선택 > -1 Then
            도움말 = .List(i선택, 3)

            Label_지출품의도움말.caption = 도움말
        End If
    End With
End Sub

Private Sub MultiPage1_click(ByVal Index As Long)
    Select Case MultiPage1.SelectedItem.name
        Case "page_일상회계관리":  '일상회계관리
            Call listbox_초기화("일상회계")
        Case "page_지출결의":  '지출결의
            Call listbox_초기화("지출결의")
        Case "page_지출품의":  '지출품의
            Call listbox_초기화("지출품의")
        Case "page_설정":  '설정
            Call listbox_초기화("설정")
        Case "page_예산":  '예산
            Call listbox_초기화("예산")
        Case "page_결산":  '결산
            Call listbox_초기화("결산")
        Case Else '자산/채무관리
            Call listbox_초기화("자산채무")
    End Select
End Sub

Private Sub UserForm_Initialize()
    
    With ListBox_일상회계
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_지출결의
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_지출품의
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_설정
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_예산
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_결산
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    With ListBox_자산채무
        .columnCount = 3
        .ColumnWidths = "0cm;0cm;2cm"
    End With
    
    Call listbox_초기화("일상회계")
    Call listbox_초기화("지출결의")
    Call listbox_초기화("지출품의")
    Call listbox_초기화("설정")
    Call listbox_초기화("예산")
    Call listbox_초기화("결산")
    Call listbox_초기화("자산채무")
    
    MultiPage1.Value = 0  '첫페이지(일상회계관리)가 항상 먼저 뜨도록
End Sub

Sub listbox_초기화(도움말항목 As String)
    
    Dim 상황 As Range
    Dim ws As Worksheet
    Set ws = Worksheets("기능도움말")
    Dim vlist() As Variant
    Dim 도움말수 As Integer
    Dim x As Integer
    x = 0
    Const 헤더줄수 As Integer = 1
        
    Select Case 도움말항목
        Case "일상회계"
            ListBox_일상회계.Clear
        Case "지출결의"
            ListBox_지출결의.Clear
        Case "지출품의"
            ListBox_지출품의.Clear
        Case "설정"
            ListBox_설정.Clear
        Case "예산"
            ListBox_예산.Clear
        Case "결산"
            ListBox_결산.Clear
        Case "자산채무"
            ListBox_자산채무.Clear
    End Select
    
    With ws.Range("기능코드레이블")
        Set 상황 = .Offset(1)
        With 상황
            If .Value = "" Then
                도움말수 = 0
            Else
                If .Offset(1, 0).Value <> "" Then
                    도움말수 = .End(xlDown).Row - 헤더줄수
                End If
            End If
        End With
            
        If 도움말수 > 0 Then
            
            For i = 0 To 도움말수 - 1
                With 상황.Offset(i, 0)
            
                    If .Value <> "" Then
                        If .Offset(0, 1).Value = 도움말항목 Then
                            ReDim Preserve vlist(3, x)
                            vlist(1, x) = .Value
                            vlist(2, x) = .Offset(0, 3).Value
                            vlist(3, x) = .Offset(0, 4).Value
                            x = x + 1
                        End If
                    End If
                End With
            Next i
            
            If x > 0 Then
                Select Case 도움말항목
                    Case "일상회계"
                        ListBox_일상회계.Column = vlist
                    Case "지출결의"
                        ListBox_지출결의.Column = vlist
                    Case "지출품의"
                        ListBox_지출품의.Column = vlist
                    Case "설정"
                        ListBox_설정.Column = vlist
                    Case "예산"
                        ListBox_예산.Column = vlist
                    Case "결산"
                        ListBox_결산.Column = vlist
                    Case "자산채무"
                        ListBox_자산채무.Column = vlist
                End Select
                
            End If
        End If
    End With
End Sub
