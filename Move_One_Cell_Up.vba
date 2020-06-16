'Moving one cell in column up
Sub Move_One_Cell_Up()

    If Application.CountA(ActiveCell) = 0 Then Selection.Delete Shift:=xlUp
           
End Sub
