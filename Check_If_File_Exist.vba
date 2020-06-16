'Checking if file exist
'Path should contain .extension
'Returning true if exist
Function Check_If_File_Exist(path As String) As Boolean

    Dim TestStr As String: TestStr = ""        
        
    If path <> "" Then
        On Error Resume Next
        TestStr = Dir(path)
        On Error GoTo 0
    End If
    
    If TestStr = "" Then
        Check_If_File_Exist = False
    Else
        Check_If_File_Exist = True
    End If
    
End Function
