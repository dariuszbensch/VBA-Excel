'Moving cell up
Sub Move_Cell_Up()

    If Application.CountA(ActiveCell) = 0 Then Selection.Delete Shift:=xlUp
           
End Sub
