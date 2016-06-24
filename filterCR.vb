' Objective:
' Iterate through rows, for each unique CR_ID (ordered) found
' make sure the first "Solution" type that corresponds with that CR_ID is a
' "Solution" type

Sub FilterSolutions_Click()
    ' Filter BP 1
    ThisWorkbook.Sheets("test_sheet").Activate

    ' Declare/initiate our iterators
    Dim i, j As Integer
    i = 2
    
    ' Remove "non-solutions" on CR_ID criteria
    Do Until IsEmpty(Cells(i, 1))
        cur_id = Cells(i, 1)
        cur_type = Cells(i, 12)
        If cur_type = "Solution" Then
            If i > 2 Then
                j = i - 1
                prev_id = Cells(j, 1)
                prev_type = Cells(j, 12)
                If (prev_id = cur_id And prev_type <> "Solution") Then
                    Do While (prev_id = cur_id And prev_type <> "Solution")
                        Rows(j).EntireRow.Delete
                        j = j - 1
                        prev_id = Cells(j, 1)
                        prev_type = Cells(j, 12)
                    Loop
                End If
            End If
        End If
        i = i + 1
    Loop
    ' Remove Duplicates on CR_ID critera
    Cells.RemoveDuplicates Columns:=Array(1)
End Sub
