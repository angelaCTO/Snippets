Private Sub process_all_crs_Click()

    Dim ws As Worksheet
    Dim i, j As Integer
    
    ' Process selected BP CR sheets
    For Each ws In Worksheets
        Select Case ws.CodeName
          ' Note here, sheet "Name" may not be the same as its CodeName
          ' which is immutable
          '      "1st BP CRs", "2nd BP CRs", "3rd BP CRs", "4th BP CRs"
            Case "Sheet5", "Sheet8", "Sheet111", "Sheet14"
            With ws
                i = 2 ' Skip the header
                ' Iterate through the rows
                Do Until IsEmpty(.Cells(i, 1))
                    cur_id = .Cells(i, 1)
                    cur_type = .Cells(i, 12)
                    If cur_type = "Solution" And i > 2 Then
                            j = i - 1
                            prev_id = .Cells(j, 1)
                            prev_type = .Cells(j, 12)
                            ' Remove Non-Solutions on CR_ID criteria
                            If (prev_id = cur_id And prev_type <> "Solution") Then
                                Do While (prev_id = cur_id And prev_type <> "Solution")
                                    .Rows(j).EntireRow.Delete
                                    j = j - 1
                                    prev_id = .Cells(j, 1)
                                    prev_type = .Cells(j, 12)
                                Loop
                            End If
                    End If
                    i = i + 1
                Loop
                ' Remove Duplicates on CR_ID critera
                .Cells.RemoveDuplicates Columns:=Array(1)
            End With
        End Select
    Next ws
End Sub
