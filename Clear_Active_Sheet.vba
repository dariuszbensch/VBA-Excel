'Clear whole active sheet
Sub Clear_Active_Sheet()

    On Error GoTo Error_handler
    Dim answer As Integer: answer = MsgBox("Are you sure you want to clear the whole sheet?", vbYesNo + vbQuestion, "Empty Sheet")

    If (answer = vbYes) Then
    
        Dim selectionRange As Range: Set selectionRange = Range("A1", "XFD1048576")
        Dim s As Shape
    
        selectionRange.Clear
        
        For Each s In ActiveSheet.Shapes
            If Not Intersect(selectionRange, s.TopLeftCell) Is Nothing Then
                s.Delete
            End If
        Next s
    
    End If
    
    Cells.Select
    selectionRange.ColumnWidth = 8.43
    selectionRange.RowHeight = 15
    Selection.NumberFormat = "0"

Error_handler:
Err.Clear

Range("A1").Select

End Sub
