' DOCS:
' -----------------------------------------------------------------------------
' Processes all BP CRs 1-4 sheets by first iterating through records
' Once a unique CR_ID has been found, ensure that it is of WR_Type
' Solution, else backtrack, remove all Non-Solutions that have
' matching CR_IDs, then remove duplicate records by CR_ID criteria
'
' NOTES:
'   Note(1), Optimization Attempt - Disable all automatic application events
'   until all sheets have been processed. Counts will update automatically
'   upon completion. Once script has completed, application settings will
'   to its original state
'
'   Note(2), sheet "Name" may not be the same as its "CodeName"
'   which is immutable (this is best) as future revisions will
'   not affect sheet coding. The following naming correspondence
'   "1st BP CRs" = Sheet5
'   "2nd BP CRs" = Sheet8
'   "3rd BP CRs" = Sheet111
'   "4th BP CRs" = Sheet14
'
'   Dev Note, Need to refractor and modularaize (when time permits)
'   Dirty hack runs two loops before remDups (h_i). Investigation required
'------------------------------------------------------------------------------
Private Sub process_all_crs_Click()

    Dim bp_cr As Worksheet
    Dim i As Integer, j As Integer
    Dim h_i As Integer ' TODO - Remove a/f Refact
    
    ' Refer To Note(1)
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        
        ' Refer To Note(2)
        For Each bp_cr In Worksheets
            Select Case bp_cr.CodeName
                Case "Sheet5", "Sheet8", "Sheet111", "Sheet14"
                With bp_cr
                    For h_i = 1 To 2 ' TODO - Remove a/f Refact
                        i = 2
                        Do Until IsEmpty(.Cells(i, 1))
                            cur_id = .Cells(i, 1)
                            cur_type = .Cells(i, 12)
                            If cur_type = "Solution" And i > 2 Then
                                 j = i - 1
                                 prev_id = .Cells(j, 1)
                                 prev_type = .Cells(j, 12)
                                    Do While (prev_id = cur_id And prev_type <> "Solution")
                                        .Rows(j).EntireRow.Delete
                                        j = j - 1
                                        prev_id = .Cells(j, 1)
                                        prev_type = .Cells(j, 12)
                                    Loop
                            End If
                            i = i + 1
                        Loop
                    Next h_i ' TODO - Remove a/f Refact
                    .Cells.RemoveDuplicates Columns:=Array(1)
                End With
            End Select
        Next bp_cr
        
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
