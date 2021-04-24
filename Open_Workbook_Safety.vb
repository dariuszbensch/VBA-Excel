'Checking if any other file with macro is open
Function Open_Workbook_Safety()

    Dim counter As Long: counter = 0
    Dim thisWorkbookName As String: thisWorkbookName = ThisWorkbook.name
    Dim workbookName As String
    Dim filesList As String
        
    For i = 1 To Application.Workbooks.Count
    
        'Getting open files names
        workbookName = Application.Workbooks(i).name
    
        'Checking if opened file can cointain macro
        If Right(workbookName, 5) = ".xlsm" Then
            
            counter = counter + 1
            
            'Create list of open files containing macro
            If workbookName <> thisWorkbookName Then filesList = filesList + workbookName + vbCrLf
            
        End If
        
    Next
    
    
    If counter > 1 Then
    
        MsgBox "You already have opened files containing macros. " + _
               "To avoid bugs, please close all files containing macros, and then run this one. " + _
               vbCrLf + _
               vbCrLf + _
               "Please close these files:" + _
               vbCrLf + _
               vbCrLf + _
               filesList

    End If
    
End Function
