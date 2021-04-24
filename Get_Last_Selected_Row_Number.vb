'Return number of last row from selected range
Function Get_Last_Selected_Row_Number() As Long

    Dim selectionRange As Range: Set selectionRange = Selection
    Get_Last_Selected_Row_Number = selectionRange.Cells(selectionRange.Rows.Count, 1).row
    
End Function
