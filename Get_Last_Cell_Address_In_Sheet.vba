'Return address of last cell in which is any value
Function Get_Last_Cell_Address_In_Sheet(workbookName As String, sheetName As String) As String
    
    On Error GoTo Error_handler
    
        'Save file to don't count empties cells which was used but don't have value, but they are active
        ActiveWorkbook.Save
    
        Dim lastCell As String: lastCell = Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1).SpecialCells(xlLastCell).Address
        
        If (lastCell = "$A$1") Then
            
            If (Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1) = "") Then lastCell = 0
            
        End If
        
        RANGE_Last_Cell_In_Sheet = lastCell
        
Error_handler:
Err.Clear

End Function
