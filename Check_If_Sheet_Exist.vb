'Checking if sheet exist - return TRUE if sheet exist
Function Check_If_Sheet_Exist(sheetName As String) As Boolean

    On Error Resume Next
        Dim sheet As Worksheet: Set sheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If (Not sheet Is Nothing) Then
        Check_If_Sheet_Exist = True
    Else
       Check_If_Sheet_Exist = False
    End If
    
End Function
