'Function clear selected data, formats, photos etc...
Sub Clear_Selected_Range()
    
    On Error GoTo Error_handler
    Dim selectionRange As Range: Set selectionRange = Selection
    Dim s As Shape

        Selection.Clear
        
        For Each s In ActiveSheet.Shapes
            If Not Intersect(selectionRange, s.TopLeftCell) Is Nothing Then
            s.Delete
            End If
        Next s
    
    Selection.NumberFormat = "0"
    
Error_handler:
Err.Clear

End Sub
