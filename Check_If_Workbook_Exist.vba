'Checking if workbook exist
'Path should contain .extension
Function Check_If_Workbook_Exist(path As String) As Boolean

    Dim TestStr As String
        TestStr = ""
        
    If path <> "" Then
        On Error Resume Next
        TestStr = Dir(path)
        On Error GoTo 0
    End If
    
    If TestStr = "" Then
        Check_If_Workbook_Exist = False
    Else
        Check_If_Workbook_Exist = True
    End If
    
End Function
