'Moving every columns left
Sub Move_All_Cells_Left()

    If Application.CountA(ActiveCell.EntireColumn) = 0 Then Selection.EntireColumn.Delete
        
End Sub
