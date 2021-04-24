'Cut data betwen workbooks, sheets or ranges
Function Cut_Data(fromWorkbook As String, _
                  fromSheet As String, _
                  fromRange As String, _
                  toWorkbook As String, _
                  toSheet As String, _
                  toCell As String)


    'Switch off unnecessary comunicates
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    
    
    On Error GoTo Error_handler
        
        
        'Opening required workbooks
        Workbooks.Open (fromWorkbook)
        Workbooks.Open (toWorkbook)
        
        
        'Cut data from sourece to destination file. "Dir" extract file name from path
        Workbooks(Dir(fromWorkbook)).Worksheets(fromSheet).Range(fromRange).Cut _
        Workbooks(Dir(toWorkbook)).Worksheets(toSheet).Range(toCell)
        

        'Optional can save and close destination file:
        Workbooks(Dir(fromWorkbook)).Save
        Workbooks(Dir(toWorkbook)).Save
        Workbooks(Dir(fromWorkbook)).Close
        Workbooks(Dir(toWorkbook)).Close
        
        
Error_handler:
Err.Clear


    'Switch on comunicates
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True


Exit Function
End Function
