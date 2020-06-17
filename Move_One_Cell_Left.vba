'Moving one cell in row left
Sub Move_One_Cell_Left()

    Dim selectionRange As Range: Set selectionRange = Selection
    Dim firstCol As Long: firstCol = selectionRange.Cells(1, 1).Column
    Dim lastCol As Long: lastCol = selectionRange.Cells(1, selectionRange.Columns.Count).Column
    Dim check As Boolean: check = True
    
    'First checkin if there is any data in current column
    If Application.CountA(ActiveCell.EntireColumn) > 0 Then
    
        check = False
    
    Else
    
        'Then checking if there is any data in current range
        For i = firstCol To lastCol
            If Application.CountA(Cells(1, i).EntireColumn) > 0 Then
                check = False
                Exit For
            End If
        Next i
    
    End If

    'If there is any data in selected column or range then it will not be deleted
    If check Then Selection.Delete Shift:=xlToLeft
    
End Sub
