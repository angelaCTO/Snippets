Private Sub Process_Rolls_Click() 'TODO Refractor

    ' Optimization Attempt - Disable all automatic events until all sheets
    ' have been processed. Counts will update automatically upon completion
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False

        ' Create a new workbook
        Dim new_wb As Workbook
        Set new_wb = Workbooks.Add
    
        'Save to designated path, named by date
        Dim fname As String, fpath As String
        fname = VBA.Format(DateSerial(Year(Date), Month(Date), Day(Date)), "mm-dd-yyyy")
        MsgBox ("Please Select A Destination To Save New File To")
        
        With .FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            .Show
            If .SelectedItems.Count <> 0 Then
                fpath = .SelectedItems(1) & "\" & fname & ".xlsm"
                'MsgBox (fpath)
                new_wb.SaveAs Filename:=fpath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            End If
        End With
    
        ' Note here, need to close and reopen wb again to avoid Excel
        ' Runtime Error 1004 (User)
        new_wb.Close SaveChanges:=True
        Set new_wb = Workbooks.Open(Filename:=fpath)
    
        ' Create the new sheets
        Dim s_sht As Worksheet, c_sht As Worksheet, e_sht As Worksheet, cd_sht As Worksheet
        With new_wb
            Set e_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 2
                e_sht.Name = "ERolls"
            Set s_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 3
                s_sht.Name = "SRolls"
            Set c_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 4
                c_sht.Name = "CRolls"
            Set cd_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count)) 'codename = sheet 5
                cd_sht.Name = "CDRolls"
        End With
    
        ' Incrementors, +2 skip the headers
        Dim i As Integer
        i = 2
        Dim ei As Integer, ex As Integer
        ei = 2
        ex = 0
        Dim si As Integer, sx As Integer
        si = 2
        sx = 0
        Dim ci As Integer, cx As Integer
        ci = 2
        cx = 0
        Dim cdi As Integer, cdx As Integer
        cdi = 2
        cdx = 0
     
        ' Copy data to each corresponding roll sheet
        With ThisWorkbook.Sheets("Sheet1")
            ' Copy headers
            .Cells(1, 1).EntireRow.Copy
            e_sht.Paste Destination:=e_sht.Cells(1, 1).EntireRow
            .Cells(1, 1).EntireRow.Copy
            s_sht.Paste Destination:=s_sht.Cells(1, 1).EntireRow
            .Cells(1, 1).EntireRow.Copy
            c_sht.Paste Destination:=c_sht.Cells(1, 1).EntireRow
            .Cells(1, 1).EntireRow.Copy
            cd_sht.Paste Destination:=cd_sht.Cells(1, 1).EntireRow
        
            ' Copy records
            Do Until IsEmpty(.Cells(i, 1))
                If InStr(.Cells(i, 4), "E") <> 0 Then
                    If InStr(.Cells(i, 4), "xx") <> 0 Then
                        ex = ex + 1
                    End If
                    .Cells(i, 1).EntireRow.Copy
                    e_sht.Paste Destination:=e_sht.Cells(ei, 1).EntireRow
                    ei = ei + 1
                ElseIf InStr(.Cells(i, 4), "S") <> 0 Then
                    If InStr(.Cells(i, 4), "xx") <> 0 Then
                        sx = sx + 1
                    End If
                    .Cells(i, 1).EntireRow.Copy
                    s_sht.Paste Destination:=s_sht.Cells(si, 1).EntireRow
                    si = si + 1
                ElseIf InStr(.Cells(i, 4), "CD") <> 0 Then
                    If InStr(.Cells(i, 4), "xx") <> 0 Then
                        cdx = cdx + 1
                    End If
                    .Cells(i, 1).EntireRow.Copy
                    cd_sht.Paste Destination:=cd_sht.Cells(cdi, 1).EntireRow
                    cdi = cdi + 1
                Else
                    If InStr(.Cells(i, 4), "xx") <> 0 Then
                        cx = cx + 1
                    End If
                    .Cells(i, 1).EntireRow.Copy
                    c_sht.Paste Destination:=c_sht.Cells(ci, 1).EntireRow
                    ci = ci + 1
                End If
                i = i + 1
            Loop
        End With
    
        ' Write the counts into Sheet1
        With new_wb.Sheets("Sheet1")
            Dim e_counts As String
            e_counts = "Total = " & (ei - 2) & " Committed = " & ((ei - 2) - ex)
            .Range("A1") = "EROLL COUNTS:"
            .Range("A2") = e_counts
        
            Dim s_counts As String
            s_counts = "Total = " & (si - 2) & " Committed = " & ((si - 2) - sx)
            .Range("A4") = "SROLL COUNTS:"
            .Range("A5") = s_counts
        
            Dim c_counts As String
            c_counts = "Total = " & (ci - 2) & " Committed = " & ((ci - 2) - cx)
            .Range("A7") = "CROLL COUNTS:"
            .Range("A8") = c_counts
        
            Dim cd_counts As String
            cd_counts = "Total = " & (cdi - 2) & " Committed = " & ((cdi - 2) - cdx)
            .Range("A10") = "CD COUNTS:"
            .Range("A11") = cd_counts
        End With

        'new_wb.Close SaveChanges:=True
    
        ' Disable annoying pop-up
        .CutCopyMode = False
        
        ans = MsgBox("Would you like to close the current workbook?", 3, "Choose Options")
        If ans = 7 Then
            ThisWorkbook.Close SaveChanges:=True
        ElseIf ans = 6 Then
            ThisWorkbook.Close SaveChanges:=False
        End If
        
        ' Return Application To Original State
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
