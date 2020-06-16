'Assign this macro to shape. After click on shape you will be moved to Sheet named same as text on the shape. If Sheet doesn't exist them will be created new one.
Sub Navigation_Shape()
    
    'Getting text from shape
    Dim sheetName As String: sheetName = ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text
    
    On Error GoTo Error_handler
        
        'If sheet doesn't exist then create new on the end and go there
        If Not Check_If_Sheet_Exist(sheetName) Then
            
            Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)).Name = sheetName
            Worksheets(sheetName).Activate
        
        'Else go to the sheet
        Else: Worksheets(sheetName).Activate
        
        End If
    
Error_handler:
Err.Clear
End Sub




'Checking if sheet exist
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
