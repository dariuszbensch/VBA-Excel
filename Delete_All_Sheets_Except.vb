'Deleting every sheets except those which has been sent in array

'Example declaration
'arr = Array("...","...")
'Call Select_Range(arr)
Sub Delete_All_Sheets_Except(arr As Variant)

    On Error GoTo Error_handler
    
        'Mute communicates
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim ws As Worksheet
        
        'Deleting every sheets except those which has been sent in array
        For Each ws In Application.ActiveWorkbook.Worksheets
            If Not Delete_Except_Checker(ws.name, arr) Then
                Sheets(ws.name).Select
                ActiveWindow.SelectedSheets.Delete
            End If
        Next
        
        'Unmute communicates
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    
Error_handler:
Err.Clear

End Sub



Private Function Delete_Except_Checker(name As Variant, arr As Variant) As Boolean
    On Error GoTo Error_handler
        'Set function default as false
        Delete_Except_Checker = False
        'Checking if table not contain elemen which has been sent, if not then return True
        For Each e In arr
            If e = name Then
                Delete_Except_Checker = True
                Exit For
            End If
        Next
Error_handler:
Err.Clear
End Function
