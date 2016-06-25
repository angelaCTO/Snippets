Private Sub Process_Rolls_Click()
    ' Create a new workbook save to designated path, named by date
    Dim new_wb As Workbook
    Set new_wb = Workbooks.Add
    
    Dim fname As String, fpath As String
    fname = VBA.Format(DateSerial(Year(Date), Month(Date), Day(Date)), "mm-dd-yyyy")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count <> 0 Then
            fpath = .SelectedItems(1) & "\" & fname & ".xlsm"
            'MsgBox (fpath)
            new_wb.SaveAs Filename:=fpath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        End If
    End With
    
    ' Note here, need to close and reopen wb again to avoid Error 1004
    new_wb.Close SaveChanges:=True
    Set new_wb = Workbooks.Open(Filename:=fpath)
    
    ' Create the new sheets
    Dim s_sht As Worksheet, c_sht As Worksheet, e_sht As Worksheet, cd_sht As Worksheet
    With new_wb
        Set e_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 2
        Set s_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 3
        Set c_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))  'codename = sheet 4
        Set cd_sht = .Sheets.Add(After:=.Sheets(.Sheets.Count)) 'codename = sheet 5
         
        e_sht.Name = "ERolls"
        s_sht.Name = "SRolls"
        c_sht.Name = "CRolls"
        cd_sht.Name = "CDRolls"
    End With
    
    
    ' Row counters, skip the headers
    Dim i As Integer, ei As Integer, si As Integer, ci As Integer, cdi As Integer
    i = 2   ' current incrementer
    ei = 2  ' e roll incrementer
    si = 2  ' s roll incrementer
    ci = 2  ' c roll incrementer
    cdi = 2 ' cd roll incrementer
    
    With ThisWorkbook.Sheets("Sheet1") ' "start" sheet codenamed 1
        ' Copy the headers to each roll sheet
        .Cells(1, 1).EntireRow.Copy
        e_sht.Paste Destination:=e_sht.Cells(1, 1).EntireRow
        .Cells(1, 1).EntireRow.Copy
        s_sht.Paste Destination:=s_sht.Cells(1, 1).EntireRow
        .Cells(1, 1).EntireRow.Copy
        c_sht.Paste Destination:=c_sht.Cells(1, 1).EntireRow
        .Cells(1, 1).EntireRow.Copy
        cd_sht.Paste Destination:=cd_sht.Cells(1, 1).EntireRow
        
        ' Copy row data to corresponding sheet
        Do Until IsEmpty(.Cells(i, 1))
            If InStr(.Cells(i, 4), "E") <> 0 Then
                .Cells(i, 1).EntireRow.Copy
                e_sht.Paste Destination:=e_sht.Cells(ei, 1).EntireRow
                ei = ei + 1

            ElseIf InStr(.Cells(i, 4), "S") <> 0 Then
                .Cells(i, 1).EntireRow.Copy
                s_sht.Paste Destination:=s_sht.Cells(si, 1).EntireRow
                si = si + 1
            
            ElseIf InStr(.Cells(i, 4), "CD") <> 0 Then
                .Cells(i, 1).EntireRow.Copy
                cd_sht.Paste Destination:=cd_sht.Cells(cdi, 1).EntireRow
                cdi = cdi + 1
            
            Else
                .Cells(i, 1).EntireRow.Copy
                c_sht.Paste Destination:=c_sht.Cells(ci, 1).EntireRow
                ci = ci + 1
            End If
            i = i + 1
        Loop
    End With
    
    new_wb.Close SaveChanges:=True
    ThisWorkbook.Close SaveChanges:=True
    
End Sub
