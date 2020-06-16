'Return number of last row which contain any value
Function Last_Row_Address_In_Sheet(workbookName As String, sheetName As String) As Long
    
    On Error GoTo Error_handler
        
        'Save file to don't count empties cells which was used but dont have value, but they are active
        ActiveWorkbook.Save
    
        Dim lastRow As Long: lastRow = Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1).SpecialCells(xlLastCell).Row
        
        If (lastRow = 1) Then
            
            If (Application.CountA(Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1).EntireRow) = 0) Then lastRow = 0
            
        End If
        
        Last_Row_Address_In_Sheet = lastRow
    
Error_handler:
Err.Clear

End Function
