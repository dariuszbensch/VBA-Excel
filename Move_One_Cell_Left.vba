'Moving one cell in row left
Sub Move_One_Cell_Left()

    Dim check As Boolean: check = True

    'Checking if there is any data in current range
    For Each OneCell In Selection
        If OneCell <> "" Then
            check = False
            Exit For
        End If
    Next OneCell

    'If there is any data in selected column or range then it will not be deleted
    If check Then Selection.Delete Shift:=xlToLeft
    
End Sub


