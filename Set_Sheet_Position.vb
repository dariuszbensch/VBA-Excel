Function Set_Sheet_Position(sheetName As String, newPosition As Long)

    On Error GoTo Error_handler
        
        Dim currentPosition As Integer: currentPosition = Sheets(sheetName).Index
        
        If newPosition <= 0 Then
            Worksheets(sheetName).Move Before:=Worksheets(1)
        ElseIf (newPosition > Sheets.Count) Then
            Worksheets(sheetName).Move After:=Worksheets(Sheets.Count)
        ElseIf (newPosition < currentPosition) Then
            Worksheets(sheetName).Move Before:=Worksheets(newPosition)
        ElseIf (newPosition > currentPosition) Then
            Worksheets(sheetName).Move After:=Worksheets(newPosition)
        End If
        
Error_handler:
Err.Clear

End Function
