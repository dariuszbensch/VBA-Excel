'VLOOKUP function
'resultTo - Address of Cell to which we want to send searched value
'lookupValueFrom - Address of Cell from which we getting value we will be looking for
'lookupRangeFrom - Table range in which will be looking data
'columnFromLookup - Number of column from which we will receive searched value
'match - False (exactly match),  True (estimated match)

Function Vlookup(resultToWorkbook As String, _
                      resultToSheet As String, _
                      resultToColumn As Long, _
                      resultToRow As Long, _
                      lookupValueFromWorkbook As String, _
                      lookupValueFromSheet As String, _
                      lookupValueFromColumn As Long, _
                      lookupValueFromRow As Long, _
                      lookupRangeFromWorkbook As String, _
                      lookupRangeFromSheet As String, _
                      lookupRangeFromTable As String, _
                      columnFromLookupTable As Long, _
                      match As Boolean)
                      

    On Error Resume Next

        Workbooks(resultToWorkbook).Worksheets(resultToSheet).Cells(resultToRow, resultToColumn).Value = Application.WorksheetFunction.VLookup( _
        Workbooks(lookupValueFromWorkbook).Worksheets(lookupValueFromSheet).Cells(lookupValueFromRow, lookupValueFromColumn).Value, _
        Workbooks(lookupRangeFromWorkbook).Worksheets(lookupRangeFromSheet).Range(lookupRangeFromTable), _
        columnFromLookupTable, _
        match)
        
    If Err.number <> 0 Then
        
        MsgBox "You used wrong values in VLOOKUP function!"
        Err.Clear
        
    End If
    

End Function
