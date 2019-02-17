Sub SelectAlAllSheet()
    Application.ScreenUpdating = False

    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    For Each sheet In Worksheets
        sheet.Activate
        Call SelectA1
    Next
    currentSheet.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub SelectA1()
    Cells(1, 1).Activate
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub
