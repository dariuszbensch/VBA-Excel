'Paste as text format
Function Paste_Plain_Text()

    On Error GoTo Error_handler
    
        If Application.CutCopyMode = xlCopy Or Application.CutCopyMode = xlCut Then
        
            ActiveSheet.Paste
            
        Else
    
            On Error Resume Next
    
                ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    
            If Err Then: Err.Clear
        
        End If
    
Error_handler:
Err.Clear

End Function
