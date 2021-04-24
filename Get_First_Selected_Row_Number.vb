'Return number of first row from selected range
Function Get_First_Selected_Row_Number() As Long

    Dim selectionRange As Range: Set selectionRange = Selection
    Get_First_Selected_Row_Number = selectionRange.Cells(1, 1).row
    
End Function
