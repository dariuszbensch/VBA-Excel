'Return letter of last column in which is any value in choosen row
'Workbook name need to have .extention
Function Get_Last_Column_Address_In_Row(workbookName As String, sheetName As String, rowNumber As Long) As String
    
    On Error GoTo Error_handler
    
        'Save file to don't count empties cells which was used but dont have value, but they are active
        ActiveWorkbook.Save
        
        Dim sheet As Worksheet: Set sheet = Workbooks(workbookName).Worksheets(sheetName)
        Dim lastColumn As Long: lastColumn = sheet.Cells(rowNumber, sheet.Columns.Count).End(xlToLeft).Column
        
        If (lastColumn = 1) Then
            If (sheet.Cells(rowNumber, 1) = "") Then lastColumn = 0
        End If
        
        Get_Last_Column_Address_In_Row = Change_Number_To_Letter(lastColumn)
  
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
