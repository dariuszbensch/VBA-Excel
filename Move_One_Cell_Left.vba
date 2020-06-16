'Moving one cell in row left
Sub Move_One_Cell_Left()

    If Application.CountA(ActiveCell) = 0 Then Selection.Delete Shift:=xlToLeft
            
End Sub
