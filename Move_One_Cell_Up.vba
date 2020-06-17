'Moving one cell in column up
Sub Move_One_Cell_Up()

    Dim selectionRange As Range: Set selectionRange = Selection
    Dim firstRow As Long: firstCol = selectionRange.Cells(1, 1).row
    Dim lastRow As Long: lastCol = selectionRange.Cells(selectionRange.Rows.Count, 1).row
    Dim check As Boolean: check = True

    'First checkin if there is any data in current row
    If Application.CountA(ActiveCell.EntireRow) > 0 Then
        
        check = False
    
    Else
    
        'Then checking if there is any data in current range
        For i = firstCol To lastCol
            If Application.CountA(Cells(i, 1).EntireRow) > 0 Then
                check = False
                Exit For
            End If
        Next i
        
    End If
    
    
    'If there is any data in selected row or range then it will not be deleted
    If check Then Selection.Delete Shift:=xlUp
    
End Sub
