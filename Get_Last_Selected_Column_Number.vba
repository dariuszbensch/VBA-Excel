'Return number of last column from selected range
Function Get_Last_Selected_Column_Number() As Long

    Dim selectionRange As Range: Set selectionRange = Selection
    lastSelectedColumn = selectionRange.Cells(1, selectionRange.Columns.Count).Column
    
End Function
