    ' Objective:
    ' Iterate through rows, for each unique PID found (made easy since PIDS are ordered)
    ' make sure the first "Project" type that corresponds with that PID is a
    ' "Project" type
    ' Note: Type=0 : Other type   (Dont Care)
    '       Type=1 : Project type
    
    Sub FilterSolutions_Click()
        Dim i, j As Integer
        i = 2
        
        Do Until IsEmpty(Cells(i, 1))
            cur_id = Cells(i, 1)
            cur_type = Cells(i, 2)
            ' 1. Unique Project ID begins with Project Type
            If cur_type = "Solution" Then
                If i > 2 Then
                    j = i - 1
                    prev_id = Cells(j, 1)
                    prev_type = Cells(j, 2)
                    If (prev_id = cur_id And prev_type <> "Solution") Then
                        ' 4. Continue iterating backwards to check whether
                        '    rows need to be removed
                        Do While (prev_id = cur_id And prev_type <> "Solution")
                            Rows(j).EntireRow.Delete
                            j = j - 1
                            prev_id = Cells(j, 1)
                            prev_type = Cells(j, 2)
                        Loop
                    End If
                End If
            End If
            i = i + 1
        Loop
    End Sub
        
