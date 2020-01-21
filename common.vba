Sub SelectA1AllSheet()
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

' シートが存在するか返す
Function SheetExists(sheetName As String)
    SheetExists = False
    Dim sheet As Worksheet
    For Each sheet In Sheets
        If sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Function GetLastRow(Optional col = 1)
    GetLastRow = Cells(Rows.Count, col).End(xlUp).Row
End Function

Sub 選択されているセル範囲内の図形を選択()
    Dim shp As shape
    Dim shapeRange As Range
    Dim selectRange As Range

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set selectRange = Selection

    For Each shp In ActiveSheet.Shapes
        Set shapeRange = Range(shp.TopLeftCell, shp.BottomRightCell)
        If Not (Intersect(shapeRange, selectRange) Is Nothing) Then
            shp.Select False
        End If
    Next
End Sub
