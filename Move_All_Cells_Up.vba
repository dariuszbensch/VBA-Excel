'Moving every rows up
Sub Move_All_Cells_Up()

    If Application.CountA(ActiveCell.EntireRow) = 0 Then Selection.EntireRow.Delete
    
End Sub
