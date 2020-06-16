Sub Move_One_Column_Left()

    If Application.CountA(ActiveCell) = 0 Then Selection.Delete Shift:=xlToLeft
            
End Sub
