Private Sub Update_BP_Metrics_Click()

    With Application

        Dim curdate As String, fpath As String
        curdate = VBA.Format(DateSerial(Year(Date), Month(Date), Day(Date)), "mm-dd-yyyy")
    
        Dim bp_wb As Workbook
        MsgBox ("Please Select The BP Metrics Excel Data File To Update")
        With .FileDialog(msoFileDialogFilePicker)
            .AllowMultiSelect = False
            .Show
            If .SelectedItems.Count <> 0 Then
                fpath = .SelectedItems(1)
                'MsgBox (fexcel)
                Set bp_wb = Workbooks.Open(Filename:=fpath)
            End If
        End With
        
        ' Opens the "Totals" Sheet and append new roll data
        Dim find_row As Range
        Dim row_num As Range
        With bp_wb.Sheets("Sheet2")
            Set find_row = .Range("A:A").Find(What:="Total WRs on ", _
                                              After:=.Cells(18, 1), _
                                              LookIn:="xlValues")
            row_num = find_row.Row
            MsgBox (row_num)
        End With
        ' Finish Later
      
    End With
End Sub
