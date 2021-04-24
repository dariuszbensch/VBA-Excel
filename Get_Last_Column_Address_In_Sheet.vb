'Return letter of last column in which is any value
Function Get_Last_Column_Address_In_Sheet(workbookName As String, sheetName As String) As String
    
    On Error GoTo Error_handler
    
        'Save file to don't count empties cells which was used but dont have value, but they are active
        ActiveWorkbook.Save
        
        Dim lastColumn As Long: lastColumn = Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1).SpecialCells(xlLastCell).column
        
        If (lastColumn = 1) Then
            If (Application.CountA(Workbooks(workbookName).Worksheets(sheetName).Cells(1, 1).EntireColumn) = 0) Then lastColumn = 0
        End If
        
        Get_Last_Column_Address_In_Sheet = Change_Number_To_Letter(lastColumn)
    
Error_handler:
Err.Clear

End Function




'Change number to letter - Number = letter from alphabet on position of number we sent
Private Function Change_Number_To_Letter(number As Long) As String
    
    On Error GoTo Error_handler
    Change_Number_To_Letter = Split(Cells(1, number).Address, "$")(1)
Error_handler:
Err.Clear

End Function
