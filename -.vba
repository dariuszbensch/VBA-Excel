'Auto fit columns / rows if checkbox is active
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If ThisWorkbook.Worksheets("WbkName").CheckBox_ColumnsAdjusting.Value Then Cells.EntireColumn.AutoFit
    If ThisWorkbook.Worksheets("WbkName").CheckBox_RowsAdjusting.Value Then Cells.EntireRow.AutoFit
End Sub


'Auto fit columns / rows
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Sub

        
'Hide / Unhide
Worksheets("Sheet1").visible = False
Worksheets("Sheet1").visible = xlSheetHidden
Worksheets("Sheet1").visible = xlSheetVeryHidden
Worksheets("Sheet1").Visible = True
Worksheets("Sheet1").Visible = xlSheetVisible

        
'Communicates disable / enable
Application.DisplayAlerts = False       
Application.EnableEvents = False        

Application.DisplayAlerts = True
Application.EnableEvents = True

        
'Launch website on double click
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Launch_Website ("https://administrator.omega.voyager-wms.net/omni/search?term=" & ActiveCell.Value)
End Sub

        
'Assign keyboard shortcut to specify function during file opening
Sub Workbook_Open()
    'Assign SHIFT + key
     Application.OnKey "+{KEY}", "FunctonName"
    'Assign CTRL + key
     Application.OnKey "^{KEY}", "FunctonName"
End Sub
        
        
'Run specify function during file opening
Sub Workbook_Open()
    Call FunctionName("OptionalParameter")  
End Sub

        
'Helpful methods
        Windows("workbookName").Activate 'Go to choosen workbook. Workbook have to be open
        
        Workbooks("workbookName").path 'Return only path to choosen workbook. Workbook have to be open
        Workbooks("workbookName").Close 'Close workbook at given name
        Workbooks.Open fileName:="path" 'Open workbook under the given path
        
        Sheets("sheetName").Visible = xlSheetVeryHidden 'Hide sheet - working if sheet exist
        Sheets("sheetName").Visible = xlSheetVisible 'Unhide sheet - working if sheet exist
        Sheets("sheetName").Select 'Go to selected sheet
        Sheets("sheetName").Delete 'Delete choosen sheet. Method will be ask if we want to delete sheet
        Sheets.Add.Name = "sheetName" 'Create new sheet in active workbook with chosen name
        Sheets.Add  'Create new sheet in active workbook
        
        ActiveWorkbook.path 'Return path to the folder of current workbook is
        ActiveWorkbook.name 'Return active workbook name with .extension
        ActiveWorkbook.Save 'Saving current active workbook
        
        ActiveSheet.name 'Return active sheet name
        
        Selection.Address 'Return address of one selected cell or range
        FileDateTime("pathToFile") 'Return date and time of file modifications
                
        Application.AskToUpdateLinks = False 'Switch off comunicates
        Application.AskToUpdateLinks = True 'Switch on comunicates     
        Application.DisplayAlerts = False 'Switch off comunicates            
        Application.DisplayAlerts = True 'Switch on comunicates     
