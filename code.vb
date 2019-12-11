
Sub TestNull()
    'If VBA.IsError(range("W4")(1, 1).Value = CVErr(xlErrNull) Then
        MsgBox "NULL detected"
    End If
End Sub
Sub UnhideAll()
    Range("15679:20060").EntireRow.Hidden = False
End Sub
Sub DeleteUnusedColumns()
    Range("NE:WD").EntireColumn.Delete
End Sub
Sub DeleteUnusedRows()
    Range("15680:20058").EntireRow.Delete
End Sub
Sub DeleteUnusedStudents()
    Dim validStudents As Variant
    validStudents = Sheets("Sheet3").Range("E2:E6350").Value
    Dim students As Variant
    students = Range("E2:E15679").Value

    Dim j As Long
    j = UBound(students, 1)
    For i = UBound(validStudents, 1) To LBound(validStudents, 1) Step -1
        Do Until students(j, 1) = validStudents(i, 1)
            Range(j + 1 & ":" & j + 1).EntireRow.Delete 'edit this line
            j = j - 1
            If j < LBound(students, 1) Then
                Exit Do 'End
            End If
        Loop
        j = j - 1
    Next
End Sub
Sub FindIsolated()
    Dim data As Variant
    data = Range("E3:E8592").Value
    For i = 1 To 8590 Step 2
        If Not data(i, 1) = data(i + 1, 1) Then
            MsgBox "i=" & i & " Row#=" & i + 2
            Range(i + 2 & ":" & i + 2).Select
            Exit Sub
        End If
    Next
    MsgBox "no isolated"
End Sub
Sub CheckTwoSetsAreSame()
    Dim a As Variant
    a = Range("E2:E6350").Value
    Dim b As Variant
    b = Range("E6355:E12703").Value
    For i = LBound(a, 1) To UBound(a, 1)
        If a(i, 1) <> b(i, 1) Then
            MsgBox "Wrong data"
        End If
    Next
    MsgBox "Done"
End Sub
Sub MergeReadMath()
    For i = 4 To 2511
        Range("WE" & i - 1 & ":BAY" & i - 1).Value = Range("W" & i & ":AEQ" & i).Value
        Range(i & ":" & i).EntireRow.Delete
    Next
End Sub
Sub SortWithQuestions()
    Dim orderStudent(6348) As Integer
    For i = 0 To 6348
        orderStudent(i) = i + 1
    Next
    
    Dim data() As Variant
    data = Range("W2:BIK6350").Value
    
    For j = UBound(data, 2) To 1 Step -1
        Call SortWithAQuestion(orderStudent, 0, 6348, data, CInt(j))
    Next
    
    'Integrate the sorted list, including student metadata
    Dim newData() As Variant
    data = Range("A2:BIL6350").Value
    newData = data
    For k = 1 To UBound(data, 1) 'row
        For ii = 1 To UBound(data, 2) 'column
            newData(k, ii) = data(orderStudent(k - 1), ii)
        Next
    Next
    
    Range("A2:BIL6350").Value = newData
    
End Sub
Sub SortWithAQuestion(ByRef orderStudent() As Integer, _
    head As Integer, tail As Integer, data() As Variant, _
    currentColumnNo As Integer)
    
    'Sort this column
    Dim nulled As Integer
    Dim valued As Integer
    nulled = head
    valued = head
    
    Do While nulled <= tail And valued <= tail
        While nulled < tail And Not VBA.IsError(data(orderStudent(nulled), currentColumnNo))
            nulled = nulled + 1
        Wend
        While valued < tail And VBA.IsError(data(orderStudent(valued), currentColumnNo))
            valued = valued + 1
        Wend
        'Pop the valued to top
        If (valued > nulled) And _
            Not VBA.IsError(data(orderStudent(valued), currentColumnNo)) And _
            VBA.IsError(data(orderStudent(nulled), currentColumnNo)) Then
            
            Dim temp As Integer
            temp = orderStudent(valued)
            For i = valued To nulled + 1 Step -1
                orderStudent(i) = orderStudent(i - 1)
            Next
            orderStudent(nulled) = temp
            nulled = nulled + 1
        Else
            valued = valued + 1
        End If
    Loop
    
End Sub
















