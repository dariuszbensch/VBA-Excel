'Moving cell left
Sub Move_Cell_Left()

    If Application.CountA(ActiveCell) = 0 Then Selection.Delete Shift:=xlToLeft
            
End Sub
