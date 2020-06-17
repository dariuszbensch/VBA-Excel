'Copy data betwen workbooks, sheets or ranges
Function Copy_Data(fromWorkbook As String, _
                   fromSheet As String, _
                   fromRange As String, _
                   toWorkbook As String, _
                   toSheet As String, _
                   toCell As String)


    'Switch off unnecessary comunicates
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    
    
    On Error GoTo Error_handler
        
        
        Dim APP As Excel.Application: Set APP = CreateObject("Excel.Application")
        Dim WBK As Workbook: Set WBK = APP.Workbooks.Open(Filename:=fromWorkbook, ReadOnly:=True, UpdateLinks:=False)
        Dim paste As Object
        
        'Copy data from source file
        WBK.Worksheets(fromSheet).Range(fromRange).Copy
        WBK.Close
        APP.Quit
    
    
        Set WBK = Nothing
        Set APP = Nothing
        
        
        'Paste data to destination file
        Set WBK = Workbooks.Open(toWorkbook)
        Set paste = WBK.Sheets(toSheet)
        paste.Activate
        Range(toCell).Select
        paste.paste
        
        
        'Optional can save and close destination file:
        'ActiveWorkbook.Save
        'Workbooks(ActiveWorkbook.Name).Close
        
        
Error_handler:
Err.Clear


    'Switch on comunicates
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True


Exit Function
End Function
