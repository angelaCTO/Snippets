Sub ClearMasterSlide()

    ' Optimization
    With Application
        .Calculations = xLCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    
    With ActiveSheet
        .DisplayPageBreaks = False

    Dim i As Integer, j As Integer	
    Dim oPres As Presentation
    
    Set oPres = ActivePresentation

    On Error Resume Next
    
    With oPres
        For i = 1 To .Designs.Count
            For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
                    .Designs(i).SlideMaster.CustomLayouts(j).Delete
            Next
        Next i
     End With

    ' Return to normal state
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
    .DisplayPageBreaks = True
    
End Sub

