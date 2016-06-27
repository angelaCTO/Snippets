Private Sub process_all_crs_Click()

    Dim bp_cr As Worksheet
    Dim i As Integer, j As Integer
    
    ' Optimization Attempt - Disable all automatic events until all sheets
    ' have been processed. Counts will update automatically upon completion
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        
        ' Process selected BP CR sheets (skip the header) first removing
        ' non-solutions on CR_ID criteria then removing duplicates
        For Each bp_cr In Worksheets
            ' Note here, sheet "Name" may not be the same as its "CodeName"
            ' which is immutable (this is best)
            Select Case bp_cr.CodeName
                ' -> "1st BP CRs", "2nd BP CRs", "3rd BP CRs", "4th BP CRs"
                Case "Sheet5", "Sheet8", "Sheet111", "Sheet14"
                With bp_cr
                    i = 2
                    Do Until IsEmpty(.Cells(i, 1))
                        cur_id = .Cells(i, 1)
                        cur_type = .Cells(i, 12)
                        If cur_type = "Solution" And i > 2 Then
                                j = i - 1
                                prev_id = .Cells(j, 1)
                                prev_type = .Cells(j, 12)
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
                    .Cells.RemoveDuplicates Columns:=Array(1)
                End With
            End Select
        Next bp_cr
        
        ' Return Application To Original State
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
