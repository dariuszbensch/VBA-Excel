'Moving one cell in column up
Sub Move_One_Cell_Up()

    Dim check As Boolean: check = True

    'Checking if there is any data in current range
    For Each OneCell In Selection
        If OneCell <> "" Then
            check = False
            Exit For
        End If
    Next OneCell
    
    'If there is any data in selected row or range then it will not be deleted
    If check Then Selection.Delete Shift:=xlUp
    
End Sub
