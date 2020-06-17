'Return number of first column from selected range
Function Get_First_Selected_Column_Number() As Long

    Dim selectionRange As Range: Set selectionRange = Selection
    firstSelectedColumn = selectionRange.Cells(1, 1).Column
    
End Function
