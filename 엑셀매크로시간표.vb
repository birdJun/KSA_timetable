Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row = 1 And Target.Column = 2 Then
        If Target.Value = "" Then
            Me.Cells(1, 3).Value = "← 학번을 입력하세요. (OO-OOO 형식으로)"
            Me.Range("b4:ID40").ClearContents
            Exit Sub
        End If
        시간표그리기
    End If
End Sub

Private Sub 시간표그리기()
    Dim 학번 As String
    Dim 분반현황 As Worksheet
    Set 분반현황 = ThisWorkbook.Worksheets("분반현황")
    Dim 과목명(20), 분반(20) As String
    Dim 시간표 As Worksheet
    Set 시간표 = ThisWorkbook.Worksheets("2023-1학기 시간표")
    Dim n As Integer
    Dim a As Integer
    Dim 중복 As Integer
    중복 = 0
    a = 0
    Dim 같이듣는사람(240) As String
    Dim 같이듣는사람수(240) As Integer
    Dim temp1, temp2 As String
    
    학번 = Me.Cells(1, 2).Value
    자리 = 0
    
    Me.Range("b4:ID39").ClearContents
    
    For i = 7 To 414
        If Left(분반현황.Cells(i, 1), 6) = 학번 Then
            자리 = i
            Exit For
        End If
    Next i
    
    If 자리 = 0 Then
        Me.Cells(1, 3).Value = "학번이 올바르지 않습니다."
        Exit Sub
    End If
    Me.Cells(1, 3) = ""
    
    
    과목수 = 0
    For i = 4 To 400
        If 분반현황.Cells(자리, i).Value = 1 Then
            과목명길게 = 분반현황.Cells(2, i).Value
            뭐지 = Left(Right(과목명길게, 2), 1)
            If Left(Right(과목명길게, 2), 1) = "_" Then
                과목수 = 과목수 + 1
                분반(과목수) = Right(과목명길게, 1)
                과목명(과목수) = Left(과목명길게, Len(과목명길게) - 2)
                n = 0
                Me.Cells(과목수 + 19, 2).Value = ((과목명(과목수)) + " " + 분반(과목수) + "분반")
                For j = 7 To 414
                    If 분반현황.Cells(j, i).Value = 1 Then
                        Me.Cells(과목수 + 19, n + 3) = 분반현황.Cells(j, 1) + " " + 분반현황.Cells(j, 2)
                        n = n + 1
                        For k = 1 To 240
                            If 같이듣는사람(k) = (분반현황.Cells(j, 1) + " " + 분반현황.Cells(j, 2)) Then
                                같이듣는사람수(k) = 같이듣는사람수(k) + 1
                                중복 = 1
                            End If
                        Next k
                        If 중복 = 0 Then
                            a = a + 1
                            같이듣는사람(a) = (분반현황.Cells(j, 1) + " " + 분반현황.Cells(j, 2))
                            같이듣는사람수(a) = 1
                        End If
                        중복 = 0
                    End If
                Next j
            End If
        End If
    Next i
    
    'Range("A1:B20").Sort Key1:=Range("B1"), Order1:=xlAscending, Key2:=Range("A1"), Order2:=xlDescending'
    
    For 순서 = 1 To 240
        If 같이듣는사람수(순서) <> 0 Then
            Me.Cells(17, 순서 + 1).Value = 같이듣는사람(순서)
            Me.Cells(18, 순서 + 1).Value = 같이듣는사람수(순서)
        End If
    Next 순서
    
    
    
    For 요일 = 2 To 14 Step 3
        For 행 = 6 To 260
            For 과목번호 = 1 To 20
                If 과목명(과목번호) = Empty Then Exit For
                If 시간표.Cells(행, 요일) = 과목명(과목번호) And 시간표.Cells(행, 요일 + 1) = 분반(과목번호) Then
                    시간표행 = 시간표.Cells(행, 1) + 3
                    시간표열 = (요일 + 4) / 3
                    If Me.Cells(시간표행, 시간표열) <> "" Then
                        MsgBox ("에러")
                        Exit Sub
                    End If
                    기록할거 = Left(과목명(과목번호), Len(과목명(과목번호)) - 3) & Chr(10) & "분반:" & 분반(과목번호) & Chr(10) & 시간표.Cells(행, 요일 + 2) & " 선생님"
                    Me.Cells(시간표행, 시간표열).Value = 기록할거
                    
                End If
            Next 과목번호
        Next 행
    Next 요일
    
    Me.Cells(11, 6).Value = "클럽활동"
    Me.Cells(12, 6).Value = "클럽활동"
    
    If Left(학번, 2) = "21" Then
        Me.Cells(11, 4).Value = "졸업연구"
        Me.Cells(12, 4).Value = "졸업연구"
        Me.Cells(13, 4).Value = "졸업연구"
        Me.Cells(14, 4).Value = "졸업연구"
        Me.Cells(15, 4).Value = "졸업연구"
    End If
    
    If Left(학번, 2) = "20" Then
        Me.Cells(11, 4).Value = "졸업연구"
        Me.Cells(12, 4).Value = "졸업연구"
        Me.Cells(13, 4).Value = "졸업연구"
        Me.Cells(14, 4).Value = "졸업연구"
        Me.Cells(15, 4).Value = "졸업연구"
    End If
    
    If Left(학번, 2) = "19" Then
        Me.Cells(11, 4).Value = "졸업연구"
        Me.Cells(12, 4).Value = "졸업연구"
        Me.Cells(13, 4).Value = "졸업연구"
        Me.Cells(14, 4).Value = "졸업연구"
        Me.Cells(15, 4).Value = "졸업연구"
    End If
    
    If Left(학번, 2) = "23" Then
        Me.Cells(10, 5).Value = "창의설계활동"
        Me.Cells(11, 5).Value = "창의설계활동"
    End If
    
    
    For i = 2 To 240
        For j = i + 1 To 240
            If Int(Me.Cells(18, j).Value) > Int(Me.Cells(18, i).Value) Then
                temp = Me.Cells(18, i).Value
                temp2 = Me.Cells(17, i).Value
                Me.Cells(18, i).Value = Me.Cells(18, j).Value
                Me.Cells(18, j) = temp
                Me.Cells(17, i).Value = Me.Cells(17, j).Value
                Me.Cells(17, j) = temp2
            End If
        Next j
    Next i
                
End Sub