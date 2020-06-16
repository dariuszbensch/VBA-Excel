'Return number of last row ich which is any value in choosen column
Function Get_Last_Row_Address_In_Column(workbookName As String, sheetName As String, columnLetter As String) As Long
    
    On Error GoTo Error_handler
        
        'Save file to don't count empties cells which was used but dont have value, but they are active
        ActiveWorkbook.Save
    
        'Changing column letter to number
        Dim column As Long: column = Change_Letter_To_Number(columnLetter)
        
        Dim sheet As Worksheet: Set sheet = Workbooks(workbookName).Worksheets(sheetName)
        Dim lastRow As Long: lastRow = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
    
        If (lastRow = 1) Then
            If (sheet.Cells(1, column) = "") Then lastRow = 0
        End If

        Get_Last_Row_Address_In_Column = lastRow
    
Error_handler:
Err.Clear

End Function




'Change letter to number - Letter position in alphabet = number
Private Function Change_Letter_To_Number(letter As String) As Integer
    
    On Error GoTo Error_handler
    Change_Letter_To_Number = Range(letter & 1).column
Error_handler:
Err.Clear

End Function
