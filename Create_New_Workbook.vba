'Create new workbook (.xlsx)
Function Create_New_Workbook(workbookName As String, Optional path As String = "")

On Error GoTo Error_handler

    'If path is not set then get path to current folder
    If path = "" Then path = ActiveWorkbook.path
     
    'Create new workbook (.xlsx) and saving in choosen location
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=(path & "\" & workbookName), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
        
Error_handler:
Err.Clear

End Function
